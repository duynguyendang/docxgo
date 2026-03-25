package xml

import "encoding/xml"

// Comment represents w:comment element.
type Comment struct {
	XMLName  xml.Name `xml:"w:comment"`
	ID       int      `xml:"w:id,attr"`
	Author   string   `xml:"w:author,attr"`
	Date     string   `xml:"w:date,attr"`
	Initials string   `xml:"w:initials,attr,omitempty"`
	Body     []interface{}
}

// CommentRangeStart represents w:commentRangeStart element.
type CommentRangeStart struct {
	XMLName xml.Name `xml:"w:commentRangeStart"`
	ID      int      `xml:"w:id,attr"`
}

// CommentRangeEnd represents w:commentRangeEnd element.
type CommentRangeEnd struct {
	XMLName xml.Name `xml:"w:commentRangeEnd"`
	ID      int      `xml:"w:id,attr"`
}

// CommentReference represents w:commentReference element.
type CommentReference struct {
	XMLName xml.Name `xml:"w:commentReference"`
	ID      int      `xml:"w:id,attr"`
}

// CommentEx represents w15:commentEx element (extended comment properties).
type CommentEx struct {
	XMLName      xml.Name `xml:"w15:commentEx"`
	ParaID       string   `xml:"w15:paraId,attr"`
	ParaIDParent string   `xml:"w15:paraIdParent,attr,omitempty"`
	Done         string   `xml:"w15:done,attr,omitempty"`
}

// CommentID represents w16cid:commentId element.
type CommentID struct {
	XMLName   xml.Name `xml:"w16cid:commentId"`
	ParaID    string   `xml:"w16cid:paraId,attr"`
	DurableID string   `xml:"w16cid:durableId,attr"`
}

// CommentExtensible represents w16cex:commentExtensible element.
type CommentExtensible struct {
	XMLName   xml.Name `xml:"w16cex:commentExtensible"`
	DurableID string   `xml:"w16cex:durableId,attr"`
	DateUtc   string   `xml:"w16cex:dateUtc,attr"`
}
