package goexcel

import "io"

type Row []interface{}

type Column []Row

type Table struct {
	Data []Column
}

type Converter interface {
	Convert(t Table) io.WriterTo
}
