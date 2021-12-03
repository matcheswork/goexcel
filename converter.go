package goexcel

import (
	"io"
	"strconv"
)

type ConverterImpl struct {
	exclr Exceler
}

func NewConverterImpl(exclr Exceler) *ConverterImpl {
	return &ConverterImpl{
		exclr,
	}
}

func (con *ConverterImpl) convertXHeader(sheet string, xAxisHeader []Cell) (io.WriterTo, error) {
	for i, v := range xAxisHeader {
		cellName := IntToLetters(i+2) + strconv.Itoa(1)
		con.exclr.SetCellValue(sheet, cellName, v)
	}

	return nil, nil
}

func (con *ConverterImpl) convertYHeader(sheet string, yAxisHeader []Cell) (io.WriterTo, error) {
	for i, v := range yAxisHeader {
		cellName := "A" + strconv.Itoa(i+2)
		con.exclr.SetCellValue(sheet, cellName, v)
	}

	return nil, nil
}

func (con *ConverterImpl) convertData(sheet string, t MapTable) (io.WriterTo, error) {
	for iy, vy := range t.YAxisHeader {
		for ix, vx := range t.XAxisHeader {
			if t.Data[vy][vx] == nil {
				continue
			}
			// adds 2 to indexes
			cellName := IntToLetters(ix+2) + strconv.Itoa(iy+2)

			con.exclr.SetCellValue(sheet, cellName, t.Data[vy][vx])
		}
	}
	return nil, nil
}

func (con *ConverterImpl) Convert(t MapTable) (io.WriterTo, error) {
	sheetName := "Sheet1"

	con.exclr.NewSheet(sheetName)

	con.convertXHeader(sheetName, t.XAxisHeader)
	con.convertYHeader(sheetName, t.YAxisHeader)
	con.convertData(sheetName, t)

	return con.exclr, nil
}
