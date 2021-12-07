package goexcel

import (
	"bytes"
	"io"
	"testing"

	"github.com/stretchr/testify/mock"
)

type mockExclr struct {
	mock.Mock
}

func (mx *mockExclr) NewSheet(name string) int {
	mx.Called(name)
	return 0
}

func (mx *mockExclr) SetCellValue(sheet string, axis string, value interface{}) error {
	mx.Called(sheet, axis, value)
	return nil
}

func (mx *mockExclr) WriteTo(w io.Writer) (int64, error) {
	mx.Called(w)
	return 0, nil
}

func TestConvert(t *testing.T) {
	var mapTable = MapTable{
		XAxisHeader: []Cell{"Column1", "Column2", "Column3", "Column4"},
		YAxisHeader: []Cell{"Row1", "Row2", "Row3", "Row4"},
		Data: map[int]map[int]Cell{
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

	testExclr := new(mockExclr)
	cnvrtr := NewConverterImpl(testExclr)

	testExclr.On("NewSheet", "Sheet1").Return(0)

	testExclr.On("SetCellValue", mock.Anything, mock.Anything, mock.Anything).Return(nil)

	testExclr.On("WriteTo", mock.Anything).Return(0, nil)

	writerTo, err := cnvrtr.Convert(mapTable)

	testExclr.AssertNumberOfCalls(t, "SetCellValue", 22)

	if err != nil {
		t.Errorf("Convert() gave error: %s", err)

	}

	var b bytes.Buffer
	_, err = writerTo.WriteTo(&b)

	if err != nil {
		t.Fatalf("WriteTo() gave error: %s", err)
	}
}

func TestConvertMissingRowData(t *testing.T) {
	var mapTable = MapTable{
		XAxisHeader: []Cell{"Column1", "Column2", "Column3", "Column4"},
		YAxisHeader: []Cell{"Row1", "Row2", "Row3", "Row4"},
		Data: map[int]map[int]Cell{
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
		},
	}

	testExclr := new(mockExclr)
	cnvrtr := NewConverterImpl(testExclr)

	testExclr.On("NewSheet", "Sheet1").Return(0)

	testExclr.On("SetCellValue", mock.Anything, mock.Anything, mock.Anything).Return(nil)

	testExclr.On("WriteTo", mock.Anything).Return(0, nil)

	writerTo, err := cnvrtr.Convert(mapTable)

	testExclr.AssertNumberOfCalls(t, "SetCellValue", 18)

	if err != nil {
		t.Errorf("Convert() gave error: %s", err)

	}

	var b bytes.Buffer
	_, err = writerTo.WriteTo(&b)

	if err != nil {
		t.Fatalf("WriteTo() gave error: %s", err)
	}
}

func TestConvertMissingColumnData(t *testing.T) {
	var mapTable = MapTable{
		XAxisHeader: []Cell{"Column1", "Column2", "Column3"},
		YAxisHeader: []Cell{"Row1", "Row2", "Row3", "Row4"},
		Data: map[int]map[int]Cell{
			0: {
				1: 12,
				2: 13,
				3: 14,
			},
			1: {
				2: 23,
				3: 24,
			},
			2: {
				1: 32,
				3: 34,
			},
			3: {
				1: 42,
				2: 43,
				3: 44,
			},
		},
	}

	testExclr := new(mockExclr)
	cnvrtr := NewConverterImpl(testExclr)

	testExclr.On("NewSheet", "Sheet1").Return(0)

	testExclr.On("SetCellValue", mock.Anything, mock.Anything, mock.Anything).Return(nil)

	testExclr.On("WriteTo", mock.Anything).Return(0, nil)

	writerTo, err := cnvrtr.Convert(mapTable)

	testExclr.AssertNumberOfCalls(t, "SetCellValue", 13)

	if err != nil {
		t.Errorf("Convert() gave error: %s", err)

	}

	var b bytes.Buffer
	_, err = writerTo.WriteTo(&b)

	if err != nil {
		t.Fatalf("WriteTo() gave error: %s", err)
	}
}
