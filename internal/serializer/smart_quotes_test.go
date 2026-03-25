package serializer

import (
	"testing"
)

func TestEncodeSmartQuotes(t *testing.T) {
	tests := []struct {
		name     string
		input    string
		expected string
	}{
		{
			name:     "left double quote",
			input:    "Hello \u201CWorld\u201D",
			expected: "Hello &#x201C;World&#x201D;",
		},
		{
			name:     "apostrophe",
			input:    "It\u2019s a test",
			expected: "It&#x2019;s a test",
		},
		{
			name:     "left single quote",
			input:    "\u2018Quote\u2019",
			expected: "&#x2018;Quote&#x2019;",
		},
		{
			name:     "en dash",
			input:    "2020\u20132021",
			expected: "2020&#x2013;2021",
		},
		{
			name:     "em dash",
			input:    "word\u2014word",
			expected: "word&#x2014;word",
		},
		{
			name:     "no smart quotes",
			input:    "plain text",
			expected: "plain text",
		},
		{
			name:     "empty string",
			input:    "",
			expected: "",
		},
		{
			name:     "mixed quotes",
			input:    "She said \u201CHello\u201D and he replied \u2018Hi\u2019",
			expected: "She said &#x201C;Hello&#x201D; and he replied &#x2018;Hi&#x2019;",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result := EncodeSmartQuotes(tt.input)
			if result != tt.expected {
				t.Errorf("EncodeSmartQuotes(%q) = %q, want %q", tt.input, result, tt.expected)
			}
		})
	}
}

func TestSerializeTextContent_SmartQuotes(t *testing.T) {
	s := NewRunSerializer()

	tests := []struct {
		name          string
		input         string
		expectedText  string
		expectedSpace string
	}{
		{
			name:          "smart quotes encoded",
			input:         "It\u2019s",
			expectedText:  "It&#x2019;s",
			expectedSpace: "",
		},
		{
			name:          "leading space preserved",
			input:         " word",
			expectedText:  " word",
			expectedSpace: "preserve",
		},
		{
			name:          "trailing space preserved",
			input:         "word ",
			expectedText:  "word ",
			expectedSpace: "preserve",
		},
		{
			name:          "leading tab preserved",
			input:         "\tword",
			expectedText:  "\tword",
			expectedSpace: "preserve",
		},
		{
			name:          "no space attribute when no whitespace",
			input:         "word",
			expectedText:  "word",
			expectedSpace: "",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result := s.serializeTextContent(tt.input)
			if result == nil {
				t.Fatal("serializeTextContent returned nil")
			}
			if result.Content != tt.expectedText {
				t.Errorf("Content = %q, want %q", result.Content, tt.expectedText)
			}
			if result.Space != tt.expectedSpace {
				t.Errorf("Space = %q, want %q", result.Space, tt.expectedSpace)
			}
		})
	}
}
