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

func (con *ConverterImpl) calculateCellLetter(index int) string {
	headerOffset := 1 // 1 cell offset to the right to leave space for Y header

	return LetterIndexToExcelColumnName(uint64(index + headerOffset))
}

func (con *ConverterImpl) calculateCellNumber(index int) string {
	headerOffset := 1                      // 1 cell offset to the bottom to leave space for X header
	zeroBasedToOneBasedIndexingOffset := 1 // to compensate for excel 1 based indexing (A1, A2)

	return strconv.Itoa(index + headerOffset + zeroBasedToOneBasedIndexingOffset)
}

func (con *ConverterImpl) convertXHeader(sheet string, xAxisHeader []Cell) error {
	for i, v := range xAxisHeader {
		cellName := con.calculateCellLetter(i) + strconv.Itoa(1)
		err := con.exclr.SetCellValue(sheet, cellName, v)
		if err != nil {
			return err
		}
	}
	return nil
}

func (con *ConverterImpl) convertYHeader(sheet string, yAxisHeader []Cell) error {
	for i, v := range yAxisHeader {
		cellName := "A" + con.calculateCellNumber(i)
		err := con.exclr.SetCellValue(sheet, cellName, v)

		if err != nil {
			return err
		}
	}
	return nil
}

func (con *ConverterImpl) convertData(sheet string, t MapTable) error {
	for indexY := range t.YAxisHeader {
		for indexX := range t.XAxisHeader {
			if t.Data[indexY][indexX] == nil {
				continue
			}

			cellLetters := con.calculateCellLetter(indexX)
			cellNumber := con.calculateCellNumber(indexY)

			cellName := cellLetters + cellNumber

			err := con.exclr.SetCellValue(sheet, cellName, t.Data[indexY][indexX])

			if err != nil {
				return err
			}
		}
	}

	return nil
}

func (con *ConverterImpl) Convert(t MapTable) (io.WriterTo, error) {
	sheetName := "Sheet1"

	con.exclr.NewSheet(sheetName)

	err := con.convertXHeader(sheetName, t.XAxisHeader)
	if err != nil {
		return nil, err
	}

	err = con.convertYHeader(sheetName, t.YAxisHeader)
	if err != nil {
		return nil, err
	}

	err = con.convertData(sheetName, t)
	if err != nil {
		return nil, err
	}

	return con.exclr, nil
}
