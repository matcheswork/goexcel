package main

import (
	"fmt"
	"net/http"

	"github.com/xuri/excelize/v2"
)

func main() {
	mux := http.NewServeMux()

	mux.HandleFunc("/download", func(rw http.ResponseWriter, r *http.Request) {
		f := excelize.NewFile()
		// Create a new sheet.
		index := f.NewSheet("Sheet2")
		// Set value of a cell.
		f.SetCellValue("Sheet2", "A2", "Hello world.")
		f.SetCellValue("Sheet1", "B2", 100)
		// Set active sheet of the workbook.
		f.SetActiveSheet(index)

		rw.Header().Add("Content-type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
		_, err := f.WriteTo(rw)
		if err != nil {
			fmt.Fprintf(rw, "%s", err)
		}
	})

	http.ListenAndServe(":9874", mux)
}
