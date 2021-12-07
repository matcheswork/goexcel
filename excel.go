package goexcel

import "io"

type Cell interface{}

type MapTable struct {
	YAxisHeader []Cell
	XAxisHeader []Cell
	Data        map[int]map[int]Cell
}
type Converter interface {
	Convert(t MapTable) (io.WriterTo, error)
}
