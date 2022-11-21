package main

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

func main() {
	// create a new spreadsheet
	f := excelize.NewFile()
	var (
		// cell values
		data = [][]interface{}{
			{"Lớp 1", "Lớp 2", "Lớp 3", "Lớp 4", "Lớp 5", "Lớp 6", "Lớp 7", "Lớp 8", "Lớp 9", "Lớp 10", "Lớp 11", "Lớp 12"},
			{"Class1_math", "Potato", nil, "Drop Down 1", "Drop Down 2"},
			{"Apple", "Tomato"},
			{"Grapes", "Spinach"},
			{"Strawberry", "Onion"},
			{"Kiwi", "Cucumber"},
		}
		addr                    string
		err                     error
		cellsStyle, headerStyle int
	)
	// set each cell value
	for r, row := range data {
		if addr, err = excelize.JoinCellName("A", r+1); err != nil {
			fmt.Println(err)
			return
		}
		if err = f.SetSheetRow("Sheet2", addr, &row); err != nil {
			fmt.Println(err)
			return
		}
	}
	// set data validation
	dvRange1 := excelize.NewDataValidation(true)
	dvRange1.Sqref = "D3:D3"
	dvRange1.SetSqrefDropList("$A$1:$B$1")
	if err = f.AddDataValidation("Sheet1", dvRange1); err != nil {
		fmt.Println(err)
		return
	}
	dvRange2 := excelize.NewDataValidation(true)
	dvRange2.Sqref = "E3:E3"
	dvRange2.SetSqrefDropList("INDIRECT(D3)")
	if err = f.AddDataValidation("Sheet1", dvRange2); err != nil {
		fmt.Println(err)
		return
	}
	// set defined name
	if err = f.SetDefinedName(&excelize.DefinedName{
		Name:     "Fruits",
		RefersTo: "Sheet1!$A$2:$A$6",
		Scope:    "Sheet1",
	}); err != nil {
		fmt.Println(err)
		return
	}
	if err = f.SetDefinedName(&excelize.DefinedName{
		Name:     "Vegetables",
		RefersTo: "Sheet1!$B$2:$B$6",
		Scope:    "Sheet1",
	}); err != nil {
		fmt.Println(err)
		return
	}
	// set custom column width
	for col, width := range map[string]float64{
		"A": 12, "B": 12, "C": 6, "D": 12, "E": 12} {
		if err = f.SetColWidth("Sheet1", col, col, width); err != nil {
			fmt.Println(err)
			return
		}
	}
	// hide gridlines for the worksheet
	if err = f.SetSheetViewOptions("Sheet1", 0,
		excelize.ShowGridLines(false)); err != nil {
		fmt.Println(err)
		return
	}
	// define the border style
	border := []excelize.Border{
		{Type: "top", Style: 1, Color: "cccccc"},
		{Type: "left", Style: 1, Color: "cccccc"},
		{Type: "right", Style: 1, Color: "cccccc"},
		{Type: "bottom", Style: 1, Color: "cccccc"},
	}
	// define the style of cells
	if cellsStyle, err = f.NewStyle(&excelize.Style{
		Font:   &excelize.Font{Color: "333333"},
		Border: border}); err != nil {
		fmt.Println(err)
		return
	}
	// define the style of the header row
	if headerStyle, err = f.NewStyle(&excelize.Style{
		Font: &excelize.Font{Bold: true},
		Fill: excelize.Fill{
			Type: "pattern", Color: []string{"dae9f3"}, Pattern: 1},
		Border: border},
	); err != nil {
		fmt.Println(err)
		return
	}
	// set cell style
	if err = f.SetCellStyle("Sheet1", "A2", "B6", cellsStyle); err != nil {
		fmt.Println(err)
		return
	}
	if err = f.SetCellStyle("Sheet1", "D3", "E3", cellsStyle); err != nil {
		fmt.Println(err)
		return
	}
	// set cell style for the header row
	if err = f.SetCellStyle("Sheet1", "A1", "B1", headerStyle); err != nil {
		fmt.Println(err)
		return
	}
	if err = f.SetCellStyle("Sheet1", "D2", "E2", headerStyle); err != nil {
		fmt.Println(err)
		return
	}
	// save spreadsheet file
	if err := f.SaveAs("Book1.xlsx"); err != nil {
		fmt.Println(err)
	}
}
