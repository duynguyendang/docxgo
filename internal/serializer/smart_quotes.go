package serializer

import "strings"

// smartQuoteReplacements maps smart Unicode quotes to their XML entity equivalents.
// This ensures that smart quotes survive XML round-tripping without encoding issues.
var smartQuoteReplacements = map[rune]string{
	'\u201C': "&#x201C;", // left double quotation mark
	'\u201D': "&#x201D;", // right double quotation mark
	'\u2018': "&#x2018;", // left single quotation mark
	'\u2019': "&#x2019;", // right single quotation mark (apostrophe)
	'\u2013': "&#x2013;", // en dash
	'\u2014': "&#x2014;", // em dash
}

// EncodeSmartQuotes replaces smart Unicode quotes with XML entities.
// This is applied to all text content during serialization to ensure
// the quotes survive any XML processing pipeline.
func EncodeSmartQuotes(text string) string {
	var sb strings.Builder
	sb.Grow(len(text) + len(text)/10) // pre-allocate with ~10% extra for entities

	for _, r := range text {
		if entity, ok := smartQuoteReplacements[r]; ok {
			sb.WriteString(entity)
		} else {
			sb.WriteRune(r)
		}
	}
	return sb.String()
}
