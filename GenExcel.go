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

	// Set the headers and sample data
	f.SetCellValue("Sheet1", "A1", "Header1")
	f.SetCellValue("Sheet1", "B1", "Header2")
	f.SetCellValue("Sheet1", "C1", "Header3")
	f.SetCellValue("Sheet1", "A2", "Data1")
	f.SetCellValue("Sheet1", "B2", 10)
	f.SetCellValue("Sheet1", "C2", "Value1")
	f.SetCellValue("Sheet1", "A3", "Data2")
	f.SetCellValue("Sheet1", "B3", 20)
	f.SetCellValue("Sheet1", "C3", "Value2")

	// Save the sample Excel file
	if err := f.SaveAs("sample.xlsx"); err != nil {
		fmt.Println(err)
		return
	}
}
