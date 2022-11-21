package main

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

func main() {
	// create a new spreadsheet
	f := excelize.NewFile()
	var (
		// Class cell values
		classData = [][]interface{}{
			{"class1", "class2", "class3", "class4", "class5", "class6", "Lớp 7", "Lớp 8", "Lớp 9", "Lớp 10", "Lớp 11", "Lớp 12"},
			{"Class1_math", "Class2_math", "Class3_math", "Class4_math", "Class5_math", "Class6_math", "Class7_math", "Class8_math", "Class9_math", "Class10_math", "Class11_math", "Class12_math"},
			{"Class1_vietnamese", "Class2_vietnamese", "Class3_vietnamese", "Class4_vietnamese", "Class5_vietnamese", "Class6_chemistry", "Class7_chemistry", "Class8_chemistry", "Class9_chemistry", "Class10_chemistry", "Class11_chemistry", "Class12_chemistry"},
			{"", "", "", "", "", "", "Class7_bilogical", "Class8_bilogical", "Class9_bilogical", "Class10_bilogical", "Class11_bilogical", "Class12_bilogical"},
		}
		//
		subjectData = [][]interface{}{
			{"anh_van", "vat_ly", "ngu_van", "toan_hoc", "sinh_hoc", "hoa_hoc", "dia_ly", "lich_su", "IELTS", "toan_tu_duy", "tieng_viet"},
			{"", "", "", "class1_toan_hoc_chapter_1", "", "", "", "", "", "", "", "class1_tieng_viet_chapter_1"},
			{"", "", "", "class1_toan_hoc_chapter_2", "", "", "", "", "", "", "", "class1_tieng_viet_chapter_2"},
			{"", "", "", "class1_toan_hoc_chapter_3", "", "", "", "", "", "", "", "class1_tieng_viet_chapter_3"},
		}
		addr       string
		err        error
		cellsStyle int
	)
	//Class set each cell value
	for r, row := range classData {
		if addr, err = excelize.JoinCellName("A", r+1); err != nil {
			fmt.Println(err)
			return
		}

		f.NewSheet("Class_Subject")
		if err = f.SetSheetRow("Class_Subject", addr, &row); err != nil {
			fmt.Println(err)
			return
		}
	}
	//Subject set each cell value
	for r, row := range subjectData {
		if addr, err = excelize.JoinCellName("A", r+1); err != nil {
			fmt.Println(err)
			return
		}
		f.NewSheet("Class1_Subject_Chapter")
		if err = f.SetSheetRow("Class1_Subject_Chapter", addr, &row); err != nil {
			fmt.Println(err)
			return
		}
	}
	if err = f.SetDefinedName(&excelize.DefinedName{
		Name:     "class1",
		RefersTo: "Class_Subject!$A$2:$A$3",
		Scope:    "Class_Subject",
	}); err != nil {
		fmt.Println(err)
		return
	}
	if err = f.SetDefinedName(&excelize.DefinedName{
		Name:     "class2",
		RefersTo: "Class_Subject!$B$2:$B$3",
		Scope:    "Class_Subject",
	}); err != nil {
		fmt.Println(err)
		return
	}
	if err = f.SetDefinedName(&excelize.DefinedName{
		Name:     "class3",
		RefersTo: "Class_Subject!$C$2:$C$3",
		Scope:    "Class_Subject",
	}); err != nil {
		fmt.Println(err)
		return
	}
	// set data validation
	dvRange1 := excelize.NewDataValidation(true)
	dvRange1.Sqref = "D3:D3"
	dvRange1.SetSqrefDropList("Class_Subject!$A$1:$M$1")
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
	/*if err = f.SetDefinedName(&excelize.DefinedName{
		Name:     "Class",
		RefersTo: "Class1_Subject!$A$2:$A$6",
		Scope:    "Sheet1",
	}); err != nil {
		fmt.Println(err)
		return
	}*/
	//if err = f.SetDefinedName(&excelize.DefinedName{
	//	Name:     "Subject",
	//	RefersTo: "Sheet1!$B$2:$B$6",
	//	Scope:    "Sheet1",
	//}); err != nil {
	//	fmt.Println(err)
	//	return
	//}
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
	//if headerStyle, err = f.NewStyle(&excelize.Style{
	//	Font: &excelize.Font{Bold: true},
	//	Fill: excelize.Fill{
	//		Type: "pattern", Color: []string{"dae9f3"}, Pattern: 1},
	//	Border: border},
	//); err != nil {
	//	fmt.Println(err)
	//	return
	//}
	// set cell style
	//if err = f.SetCellStyle("Sheet1", "A2", "B6", cellsStyle); err != nil {
	//	fmt.Println(err)
	//	return
	//}
	if err = f.SetCellStyle("Sheet1", "D3", "E3", cellsStyle); err != nil {
		fmt.Println(err)
		return
	}
	// set cell style for the header row
	//if err = f.SetCellStyle("Sheet1", "A1", "B1", headerStyle); err != nil {
	//	fmt.Println(err)
	//	return
	//}
	//if err = f.SetCellStyle("Sheet1", "D2", "E2", headerStyle); err != nil {
	//	fmt.Println(err)
	//	return
	//}
	// save spreadsheet file
	if err := f.SaveAs("sample.xlsx"); err != nil {
		fmt.Println(err)
	}
}
