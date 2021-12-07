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
	Data: map[int]map[int]goexcel.Cell{
		0: {
			0: 11,
			1: 12,
			2: 13,
			3: 14,
		},
		1: {
			0: 21,
			2: 23,
			3: 24,
		},
		2: {
			0: 31,
			1: 32,
			3: 34,
		},
		3: {
			0: 41,
			1: 42,
			2: 43,
			3: 44,
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
