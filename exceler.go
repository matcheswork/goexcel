package goexcel

import (
	"io"

	"github.com/xuri/excelize/v2"
)

type Exceler interface {
	NewSheet(name string) int
	SetCellValue(sheet string, axis string, value interface{}) error
	WriteTo(w io.Writer) (int64, error)
}

type ExcelerImpl struct {
	f *excelize.File
}

func NewExcelerImpl(f *excelize.File) *ExcelerImpl {
	return &ExcelerImpl{
		f,
	}
}

func (exclr *ExcelerImpl) NewSheet(name string) int {
	return exclr.f.NewSheet(name)
}

func (exclr *ExcelerImpl) SetCellValue(sheet string, axis string, value interface{}) error {
	return exclr.f.SetCellValue(sheet, axis, value)
}

func (exclr *ExcelerImpl) WriteTo(w io.Writer) (int64, error) {
	return exclr.f.WriteTo(w)
}
