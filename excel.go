package goexcel

import "io"

type Cell interface{}

type MapTable struct {
	YAxisHeader []Cell
	XAxisHeader []Cell
	Data        map[Cell]map[Cell]Cell
}
type Converter interface {
	Convert(t MapTable) (io.WriterTo, error)
}
