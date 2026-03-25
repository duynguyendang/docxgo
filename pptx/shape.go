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

package pptx

// Shape represents a shape on a slide.
type Shape struct {
	id        int
	shapeType ShapeType
	x         int64
	y         int64
	w         int64
	h         int64
	fillColor *Color
	lineColor *Color
	lineWidth int
	texts     []TextRun
	alignment Alignment
	imageData []byte
	custGeom  *customGeom // Custom geometry for SVG paths
	prstGeom  string      // Preset geometry name override

	// Line/connector fields
	isLine     bool  // True if this is a line shape
	endX       int64 // Line end X (EMUs)
	endY       int64 // Line end Y (EMUs)
	arrowEnd   bool  // Arrowhead at end
	arrowStart bool  // Arrowhead at start
}

// TextRun represents a run of text with formatting.
type TextRun struct {
	Text       string
	FontSize   int // In points (e.g., 24 = 24pt)
	FontFamily string
	Color      *Color
	Bold       bool
	Italic     bool
	Underline  bool
}

// TextShape provides a fluent API for configuring text shapes.
type TextShape struct {
	shape *Shape
}

// SetPosition sets the position of the text shape in EMUs.
func (ts *TextShape) SetPosition(x, y int64) *TextShape {
	ts.shape.x = x
	ts.shape.y = y
	return ts
}

// SetSize sets the size of the text shape in EMUs.
func (ts *TextShape) SetSize(w, h int64) *TextShape {
	ts.shape.w = w
	ts.shape.h = h
	return ts
}

// SetFontSize sets the font size for all text runs in points.
func (ts *TextShape) SetFontSize(size int) *TextShape {
	for i := range ts.shape.texts {
		ts.shape.texts[i].FontSize = size
	}
	return ts
}

// SetBold sets bold for all text runs.
func (ts *TextShape) SetBold(bold bool) *TextShape {
	for i := range ts.shape.texts {
		ts.shape.texts[i].Bold = bold
	}
	return ts
}

// SetItalic sets italic for all text runs.
func (ts *TextShape) SetItalic(italic bool) *TextShape {
	for i := range ts.shape.texts {
		ts.shape.texts[i].Italic = italic
	}
	return ts
}

// SetUnderline sets underline for all text runs.
func (ts *TextShape) SetUnderline(underline bool) *TextShape {
	for i := range ts.shape.texts {
		ts.shape.texts[i].Underline = underline
	}
	return ts
}

// SetColor sets the text color for all text runs.
func (ts *TextShape) SetColor(color Color) *TextShape {
	for i := range ts.shape.texts {
		ts.shape.texts[i].Color = &color
	}
	return ts
}

// SetFontFamily sets the font family for all text runs.
func (ts *TextShape) SetFontFamily(font string) *TextShape {
	for i := range ts.shape.texts {
		ts.shape.texts[i].FontFamily = font
	}
	return ts
}

// SetAlignment sets the text alignment.
func (ts *TextShape) SetAlignment(align Alignment) *TextShape {
	ts.shape.alignment = align
	return ts
}

// SetFillColor sets the background fill color of the shape.
func (ts *TextShape) SetFillColor(color Color) *TextShape {
	ts.shape.fillColor = &color
	return ts
}

// SetLine sets the border line color and width.
func (ts *TextShape) SetLine(color Color, widthPt int) *TextShape {
	ts.shape.lineColor = &color
	ts.shape.lineWidth = widthPt * EmuPerPoint / 1270 // Convert pt to EMU line width
	return ts
}

// ShapeBuilder provides a fluent API for configuring auto shapes.
type ShapeBuilder struct {
	shape *Shape
}

// SetPosition sets the position of the shape in EMUs.
func (sb *ShapeBuilder) SetPosition(x, y int64) *ShapeBuilder {
	sb.shape.x = x
	sb.shape.y = y
	return sb
}

// SetSize sets the size of the shape in EMUs.
func (sb *ShapeBuilder) SetSize(w, h int64) *ShapeBuilder {
	sb.shape.w = w
	sb.shape.h = h
	return sb
}

// SetFillColor sets the fill color of the shape.
func (sb *ShapeBuilder) SetFillColor(color Color) *ShapeBuilder {
	sb.shape.fillColor = &color
	return sb
}

// SetLine sets the border line color and width.
func (sb *ShapeBuilder) SetLine(color Color, widthPt int) *ShapeBuilder {
	sb.shape.lineColor = &color
	sb.shape.lineWidth = widthPt * EmuPerPoint / 1270
	return sb
}

// SetNoFill makes the shape transparent (no fill).
func (sb *ShapeBuilder) SetNoFill() *ShapeBuilder {
	sb.shape.fillColor = nil
	return sb
}

// SetNoLine removes the shape border.
func (sb *ShapeBuilder) SetNoLine() *ShapeBuilder {
	sb.shape.lineColor = nil
	sb.shape.lineWidth = 0
	return sb
}

// ImageBuilder provides a fluent API for configuring image shapes.
type ImageBuilder struct {
	shape *Shape
}

// SetPosition sets the position of the image in EMUs.
func (ib *ImageBuilder) SetPosition(x, y int64) *ImageBuilder {
	ib.shape.x = x
	ib.shape.y = y
	return ib
}

// SetSize sets the size of the image in EMUs.
func (ib *ImageBuilder) SetSize(w, h int64) *ImageBuilder {
	ib.shape.w = w
	ib.shape.h = h
	return ib
}

// SetWidth sets the width while maintaining aspect ratio.
func (ib *ImageBuilder) SetWidth(w int64) *ImageBuilder {
	ib.shape.w = w
	return ib
}

// SetHeight sets the height while maintaining aspect ratio.
func (ib *ImageBuilder) SetHeight(h int64) *ImageBuilder {
	ib.shape.h = h
	return ib
}

// LineBuilder provides a fluent API for configuring line shapes.
type LineBuilder struct {
	shape *Shape
}

// SetColor sets the line color.
func (lb *LineBuilder) SetColor(color Color) *LineBuilder {
	lb.shape.lineColor = &color
	return lb
}

// SetWidth sets the line width in EMUs.
func (lb *LineBuilder) SetWidth(w int64) *LineBuilder {
	lb.shape.lineWidth = int(w)
	return lb
}

// SetWidthPt sets the line width in points.
func (lb *LineBuilder) SetWidthPt(pt float64) *LineBuilder {
	lb.shape.lineWidth = int(pt * 12700)
	return lb
}

// SetArrowEnd adds an arrowhead at the end of the line.
func (lb *LineBuilder) SetArrowEnd() *LineBuilder {
	lb.shape.arrowEnd = true
	return lb
}

// SetArrowStart adds an arrowhead at the start of the line.
func (lb *LineBuilder) SetArrowStart() *LineBuilder {
	lb.shape.arrowStart = true
	return lb
}
