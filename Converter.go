package main

import (
	"fmt"
	"github.com/xuri/excelize/v2"
)

func main() {
	// Create a new Excel file and add a new sheet
	f := excelize.NewFile()
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()
	_, err := f.NewSheet("Sheet1")
	if err != nil {
		fmt.Println(err)
		return
	}

	// Set the headers
	_ = f.SetCellValue("Sheet1", "A1", "Header1")
	_ = f.SetCellValue("Sheet1", "B1", "Header2")
	_ = f.SetCellValue("Sheet1", "C1", "Header3")

	// Read data from a source Excel file
	sourceFile, err := excelize.OpenFile("sample.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
	defer func() {
		if err := sourceFile.Close(); err != nil {
			fmt.Println(err)
		}
	}()
	rows, err := sourceFile.GetRows("Sheet1")
	if err != nil {
		fmt.Println(err)
		return
	}
	for i, row := range rows {
		if i == 0 {
			// Skip the header row
			continue
		}
		// Read the first three columns
		for j := 1; j <= 3; j++ {
			cellValue := row[j-1]
			_ = f.SetCellValue("Sheet1", fmt.Sprintf("%s%d", toChar(j), i+1), cellValue)
		}
	}

	// Save the new Excel file
	if err := f.SaveAs("new.xlsx"); err != nil {
		fmt.Println(err)
		return
	}
}

func toChar(i int) string {
	result := ""
	for i > 0 {
		i--
		result = string('A'+i%26) + result
		i /= 26
	}
	return result
}
