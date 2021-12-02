package goexcel

import "io"

type Cell interface{}

type Column []Cell

type Table struct {
	Data        []Column
	YAxisHeader Column
	XAxisHeader []Cell
}

// лучший для постройки таблицы/самый гибкий
type BasicTable struct {
	Data [][]Cell
}

// лучший для работы с точечными данными
// хранит только нужные данные (без пустых значений)
type MapTable struct {
	YAxisHeader []Cell
	XAxisHeader []Cell
	Data        map[Cell]map[Cell]Cell
}
type Converter interface {
	Convert(t Table) io.WriterTo
}
