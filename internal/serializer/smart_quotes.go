package serializer

import "strings"

var smartQuoteReplacements = []struct {
	from string
	to   string
}{
	{"\u201C", "&#x201C;"}, // left double quote
	{"\u201D", "&#x201D;"}, // right double quote
	{"\u2018", "&#x2018;"}, // left single quote
	{"\u2019", "&#x2019;"}, // right single quote / apostrophe
	{"\u2013", "&#x2013;"}, // en dash
	{"\u2014", "&#x2014;"}, // em dash
}

func EncodeSmartQuotes(text string) string {
	result := text
	for _, rep := range smartQuoteReplacements {
		result = strings.ReplaceAll(result, rep.from, rep.to)
	}
	return result
}
