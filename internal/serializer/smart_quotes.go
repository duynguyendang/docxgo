package serializer

// EncodeSmartQuotes is kept for API compatibility.
// Modern XML/OOXML parsers handle UTF-8 directly, so we now output raw UTF-8.
func EncodeSmartQuotes(text string) string {
	return text
}
