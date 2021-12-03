package main

import (
	"fmt"
	"log"
	"net/http"

	"github.com/matcheswork/goexcel"

	"github.com/xuri/excelize/v2"
)

var mapTable = goexcel.MapTable{
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

func main() {
	mux := http.NewServeMux()

	mux.HandleFunc("/download", func(rw http.ResponseWriter, r *http.Request) {
		f := excelize.NewFile()
		exclr := goexcel.NewExcelerImpl(f)
		cnvrtr := goexcel.NewConverterImpl(exclr)

		writerTo, err := cnvrtr.Convert(mapTable)

		if err != nil {
			log.Fatalf("%+v", err)
		}

		rw.Header().Add("Content-type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
		_, err = writerTo.WriteTo(rw)
		if err != nil {
			fmt.Fprintf(rw, "%s", err)
		}
	})

	http.ListenAndServe(":9874", mux)
}
