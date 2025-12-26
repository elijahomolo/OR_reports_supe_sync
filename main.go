package main

import (
	"flag"
	"fmt"
	"log"
	"os"
	"strings"

	"github.com/xuri/excelize/v2"
)

type DriverInfo struct {
	FirstName   string
	LastName    string
	ColleagueID string
	ManagerName string
}

func normalize(s string) string {
	s = strings.ToLower(strings.TrimSpace(s))
	s = strings.ReplaceAll(s, "\u00a0", " ")
	s = strings.ReplaceAll(s, "\n", " ")
	s = strings.ReplaceAll(s, "\r", " ")
	s = strings.Join(strings.Fields(s), " ")
	return s
}

func cell(row []string, idx int) string {
	if idx < 0 || idx >= len(row) {
		return ""
	}
	return strings.TrimSpace(row[idx])
}

func splitDriverName(full string) (string, string) {
	full = strings.TrimSpace(full)
	if full == "" {
		return "", ""
	}

	if strings.Contains(full, ",") {
		parts := strings.Split(full, ",")
		if len(parts) == 2 {
			return strings.TrimSpace(parts[1]), strings.TrimSpace(parts[0])
		}
	}

	parts := strings.Fields(full)
	if len(parts) < 2 {
		return "", ""
	}

	return parts[0], parts[len(parts)-1]
}

func findHeaderRow(
	f *excelize.File,
	sheet string,
	required []string,
) (map[string]int, int, error) {

	rows, err := f.GetRows(sheet)
	if err != nil {
		return nil, 0, err
	}

	for rIdx, row := range rows {
		headers := map[string]int{}

		for cIdx, v := range row {
			n := normalize(v)
			if n != "" {
				headers[n] = cIdx
			}
		}

		matched := true
		for _, req := range required {
			if _, ok := headers[normalize(req)]; !ok {
				matched = false
				break
			}
		}

		if matched {
			return headers, rIdx, nil
		}
	}

	return nil, 0, fmt.Errorf("no valid header row found in %q", sheet)
}

func main() {
	allopsPath := flag.String("allops", "", "ALLOPS workforce file")
	v2Path := flag.String("v2", "", "V2 Overlooked & Returned file")
	flag.Parse()

	if *allopsPath == "" || *v2Path == "" {
		flag.Usage()
		os.Exit(1)
	}

	allops, err := excelize.OpenFile(*allopsPath)
	if err != nil {
		log.Fatal(err)
	}

	v2, err := excelize.OpenFile(*v2Path)
	if err != nil {
		log.Fatal(err)
	}

	// ---------------------------
	// ALLOPS HEADER DETECTION
	// ---------------------------
	allopsSheet := allops.GetSheetName(0)

	allopsHeaders, allopsHeaderRow, err := findHeaderRow(
		allops,
		allopsSheet,
		[]string{
			"preferred first name",
			"legal last name",
			"colleague id",
			"manager - name",
			"business area",
		},
	)
	if err != nil {
		log.Fatal(err)
	}

	// SAFETY ASSERTION
	if allopsHeaders["manager - name"] <= allopsHeaders["business area"] {
		log.Fatalf(
			"invalid ALLOPS header detection: manager-name column overlaps business-area column",
		)
	}

	// ---------------------------
	// V2 HEADER DETECTION
	// ---------------------------
	v2Sheet := "Data"
	v2Index, err := v2.GetSheetIndex(v2Sheet)
	if err != nil {
		log.Fatal(err)
	}
	if v2Index == -1 {
		log.Fatalf("sheet %q not found in V2 file", v2Sheet)
	}

	v2Headers, v2HeaderRow, err := findHeaderRow(
		v2,
		v2Sheet,
		[]string{
			"driver name",
			"supervisor name",
		},
	)
	if err != nil {
		log.Fatal(err)
	}

	// ---------------------------
	// BUILD ALLOPS LOOKUP
	// ---------------------------
	driverMap := make(map[string]DriverInfo)

	allopsRows, _ := allops.GetRows(allopsSheet)
	for i := allopsHeaderRow + 1; i < len(allopsRows); i++ {
		row := allopsRows[i]

		first := cell(row, allopsHeaders["preferred first name"])
		last := cell(row, allopsHeaders["legal last name"])
		id := cell(row, allopsHeaders["colleague id"])
		manager := cell(row, allopsHeaders["manager - name"])

		if first == "" || last == "" || id == "" {
			continue
		}

		key := normalize(first) + "|" + normalize(last)

		driverMap[key] = DriverInfo{
			FirstName:   first,
			LastName:    last,
			ColleagueID: id,
			ManagerName: manager,
		}
	}

	// ---------------------------
	// MASTER LIST SETUP
	// ---------------------------
	masterSheet := "master list"
	index, err := v2.GetSheetIndex(masterSheet)
	if err != nil {
		log.Fatal(err)
	}

	if idx := index; idx != -1 {
		v2.DeleteSheet(masterSheet)
	}
	v2.NewSheet(masterSheet)

	headers := []string{
		"Colleague ID",
		"First Name",
		"Last Name",
		"Supervisor Name",
	}

	for i, h := range headers {
		ref, _ := excelize.CoordinatesToCellName(i+1, 1)
		v2.SetCellValue(masterSheet, ref, h)
	}

	written := map[string]bool{}
	masterRow := 2

	// ---------------------------
	// PROCESS V2 DATA
	// ---------------------------
	v2Rows, _ := v2.GetRows(v2Sheet)
	for i := v2HeaderRow + 1; i < len(v2Rows); i++ {
		row := v2Rows[i]

		driver := cell(row, v2Headers["driver name"])
		first, last := splitDriverName(driver)
		if first == "" || last == "" {
			continue
		}

		info, ok := driverMap[normalize(first)+"|"+normalize(last)]
		if !ok {
			continue
		}

		// Deduplicated master list
		if !written[info.ColleagueID] {
			values := []string{
				info.ColleagueID,
				info.FirstName,
				info.LastName,
				info.ManagerName,
			}

			for c, v := range values {
				ref, _ := excelize.CoordinatesToCellName(c+1, masterRow)
				v2.SetCellValue(masterSheet, ref, v)
			}

			written[info.ColleagueID] = true
			masterRow++
		}

		// Supervisor correction
		current := cell(row, v2Headers["supervisor name"])
		if normalize(current) != normalize(info.ManagerName) {
			ref, _ := excelize.CoordinatesToCellName(
				v2Headers["supervisor name"]+1,
				i+1,
			)
			v2.SetCellValue(v2Sheet, ref, info.ManagerName)
		}
	}

	// ---------------------------
	// SAVE
	// ---------------------------
	out := "V2.Overlooked_and_Returned_VALIDATED.xlsx"
	if err := v2.SaveAs(out); err != nil {
		log.Fatal(err)
	}

	fmt.Println("✔ Data validated")
	fmt.Println("✔ Supervisor names corrected")
	fmt.Println("✔ Master list deduplicated")
	fmt.Println("✔ Output:", out)
}
