package goexcel

func LetterIndexToExcelColumnName(index uint64) (letters string) {
	firstLetterIdx := index / 26
	if firstLetterIdx >= 1 {
		zeroBasedFirstLetterIdx := firstLetterIdx - 1

		letters += LetterIndexToExcelColumnName(zeroBasedFirstLetterIdx)
		letters += string(rune('A' + index%26))
	} else {
		letters += string(rune('A' + index))
	}

	return
}
