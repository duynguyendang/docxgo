/*
MIT License

Copyright (c) 2025 Misael Monterroca

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
*/

package domain

// TrackedChange represents a tracked change (insertion or deletion) in a document.
type TrackedChange interface {
	// ID returns the unique identifier of this tracked change.
	ID() string
	// Author returns the author of this tracked change.
	Author() string
	// Date returns the ISO 8601 date string of when the change was made.
	Date() string
	// Type returns whether this is an insertion or deletion.
	Type() TrackedChangeType
	// Runs returns the runs contained within this tracked change.
	Runs() []Run
}

// TrackedChangeType distinguishes insertions from deletions.
type TrackedChangeType int

const (
	TrackedChangeInsertion TrackedChangeType = iota
	TrackedChangeDeletion
)

// Comment represents a comment attached to a range of text in the document.
type Comment interface {
	// ID returns the unique identifier of this comment.
	ID() string
	// Author returns the author of this comment.
	Author() string
	// Date returns the ISO 8601 date string of when the comment was created.
	Date() string
	// Initials returns the author's initials.
	Initials() string
	// Text returns the plain text content of this comment.
	Text() string
}
