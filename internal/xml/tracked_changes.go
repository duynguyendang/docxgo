package xml

import "encoding/xml"

// TrackedInsertion represents w:ins element (tracked insertion).
type TrackedInsertion struct {
	XMLName xml.Name `xml:"w:ins"`
	ID      int      `xml:"w:id,attr"`
	Author  string   `xml:"w:author,attr"`
	Date    string   `xml:"w:date,attr"`
	Runs    []*Run   `xml:"w:r"`
}

// TrackedDeletion represents w:del element (tracked deletion).
type TrackedDeletion struct {
	XMLName xml.Name `xml:"w:del"`
	ID      int      `xml:"w:id,attr"`
	Author  string   `xml:"w:author,attr"`
	Date    string   `xml:"w:date,attr"`
	Runs    []*Run   `xml:"w:r"`
}

// DelText represents w:delText element (deleted text content).
type DelText struct {
	XMLName xml.Name `xml:"w:delText"`
	Space   string   `xml:"xml:space,attr,omitempty"`
	Content string   `xml:",chardata"`
}

// DelRun represents a w:r element inside w:del (uses w:delText instead of w:t).
type DelRun struct {
	XMLName    xml.Name       `xml:"w:r"`
	Properties *RunProperties `xml:"w:rPr,omitempty"`
	DelText    *DelText       `xml:"w:delText,omitempty"`
}
