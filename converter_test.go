package goexcel

import (
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestConvert(t *testing.T) {
	converter := NewConverterImpl(nil)

	converter.Convert(MapTable{})
}

func TestConverterImpl(t *testing.T) {
	converter := NewConverterImpl(nil)

	assert.Implements(t, (*Converter)(nil), converter)
}
