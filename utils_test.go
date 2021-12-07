package goexcel

import (
	"testing"
)

func TestLetterIndexToColumnName(t *testing.T) {
	cells := []struct {
		index   uint64
		letters string
	}{
		{index: 0, letters: "A"},
		{index: 1, letters: "B"},
		{index: 25, letters: "Z"},
		{index: 26, letters: "AA"},
		{index: 27, letters: "AB"},
		{index: 51, letters: "AZ"},
		{index: 52, letters: "BA"},
		{index: 53, letters: "BB"},
		{index: 5633, letters: "HHR"},
	}
	for _, v := range cells {
		letters := LetterIndexToExcelColumnName(v.index)
		if v.letters != letters {
			t.Errorf("Index to value conversion was incorrect, got: %s, want: %s.", letters, v.letters)
		}
	}
}
