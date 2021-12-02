package main

import (
	"fmt"
	"net/http"
	"strconv"

	"github.com/matcheswork/goexcel"

	"github.com/xuri/excelize/v2"
)

func intToLetters(number int) (letters string) {
	number--
	firstLetterIdx := number / 26
	if firstLetterIdx >= 1 {
		letters += intToLetters(firstLetterIdx)
		letters += string(rune('A' + number%26))
	} else {
		letters += string(rune('A' + number))
	}

	return
}

func main() {
	mux := http.NewServeMux()

	mux.HandleFunc("/download", func(rw http.ResponseWriter, r *http.Request) {
		f := excelize.NewFile()
		sheet := "Sheet1"
		// // Create a new sheet.
		f.NewSheet(sheet)
		// // Set value of a cell.
		// f.SetCellValue("Sheet2", "A2", "Hello world.")
		// f.SetCellValue("Sheet1", "B2", 100)
		// // Set active sheet of the workbook.
		// f.SetActiveSheet(index)

		table := goexcel.Table{
			XAxisHeader: goexcel.Column{1, 2, 3, 4},
			YAxisHeader: []goexcel.Cell{1, 2, 3, 4},
			Data: []goexcel.Column{
				{1, 2, 3, 4},
				{1, 2, 3, 4},
				{1, 2, 3, 4},
				{1, 2, 3, 4},
			},
		}

		for i, v := range table.XAxisHeader {
			cellName := intToLetters(i+2) + strconv.Itoa(1)
			f.SetCellValue(sheet, cellName, v)
		}
		for i, v := range table.YAxisHeader {
			cellName := "A" + strconv.Itoa(i+2)
			f.SetCellValue(sheet, cellName, v)
		}

		for ix := range table.XAxisHeader {
			for iy := range table.YAxisHeader {
				cellName := intToLetters(ix+2) + strconv.Itoa(iy+2)

				f.SetCellValue(sheet, cellName, table.Data[ix][iy])
			}
		}

		sheet = "Sheet2"
		f.NewSheet(sheet)

		basicTable := goexcel.BasicTable{
			Data: [][]goexcel.Cell{
				{"Table", 1, 2, 3, 4},
				{1, 11, 21, 31, 41},
				{2, 21, 22, 32, 42},
				{3, 31, 23, 33, 43},
				{4, 41, 24, 34, 44},
			},
		}

		for iy, vy := range basicTable.Data {
			for ix, vxy := range vy {
				cellName := intToLetters(ix+1) + strconv.Itoa(iy+1)
				f.SetCellValue(sheet, cellName, vxy)
			}
		}

		sheet = "Sheet3"
		f.NewSheet(sheet)

		mapTable := goexcel.MapTable{
			XAxisHeader: []goexcel.Cell{"Column1", "Column2", "Column3", "Column4"},
			YAxisHeader: []goexcel.Cell{"Row1", "Row2", "Row3", "Row4"},
			Data: map[goexcel.Cell]map[goexcel.Cell]goexcel.Cell{
				"Row1": {
					"Column1": 11,
					"Column2": 12,
					"Column3": 13,
					"Column4": 14,
				},
				"Row2": {
					"Column1": 21,
					"Column3": 23,
					"Column4": 24,
				},
				"Row3": {
					"Column1": 31,
					"Column2": 32,
					"Column4": 34,
				},
				"Row4": {
					"Column1": 41,
					"Column2": 42,
					"Column3": 43,
					"Column4": 44,
				},
			},
		}

		for i, v := range mapTable.XAxisHeader {
			cellName := intToLetters(i+2) + strconv.Itoa(1)
			f.SetCellValue(sheet, cellName, v)
		}
		for i, v := range mapTable.YAxisHeader {
			cellName := "A" + strconv.Itoa(i+2)
			f.SetCellValue(sheet, cellName, v)
		}

		for iy, vy := range mapTable.YAxisHeader {
			for ix, vx := range mapTable.XAxisHeader {
				if mapTable.Data[vy][vx] == nil {
					continue
				}
				cellName := intToLetters(ix+2) + strconv.Itoa(iy+2)

				f.SetCellValue(sheet, cellName, mapTable.Data[vy][vx])
			}
		}

		rw.Header().Add("Content-type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
		_, err := f.WriteTo(rw)
		if err != nil {
			fmt.Fprintf(rw, "%s", err)
		}
	})

	http.ListenAndServe(":9874", mux)
}
