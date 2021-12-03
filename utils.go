package goexcel

func IntToLetters(number int) (letters string) {
	number--
	firstLetterIdx := number / 26
	if firstLetterIdx >= 1 {
		letters += IntToLetters(firstLetterIdx)
		letters += string(rune('A' + number%26))
	} else {
		letters += string(rune('A' + number))
	}

	return
}
