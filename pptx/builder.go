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

// PresentationBuilder provides a fluent API for building presentations.
//
// Example:
//
//	builder := pptx.NewPresentationBuilder(
//	    pptx.WithTitle("My Presentation"),
//	    pptx.WithAuthor("John Doe"),
//	    pptx.WithLayout(pptx.Layout16x9),
//	)
//
//	builder.AddSlide().
//	    AddText("Title Slide").
//	    SetBold(true).
//	    SetFontSize(44).
//	    SetColor(pptx.Blue).
//	    End()
//
//	pres, _ := builder.Build()
//	pres.SaveAs("presentation.pptx")
type PresentationBuilder struct {
	pres   *Presentation
	errors []error
}

// PptxOption configures the presentation builder.
type PptxOption func(*Presentation)

// WithTitle sets the presentation title.
func WithTitle(title string) PptxOption {
	return func(p *Presentation) {
		p.metadata.Title = title
	}
}

// WithAuthor sets the presentation author.
func WithAuthor(author string) PptxOption {
	return func(p *Presentation) {
		p.metadata.Author = author
		p.metadata.Creator = author
	}
}

// WithLayout sets the slide layout.
func WithLayout(layout Layout) PptxOption {
	return func(p *Presentation) {
		p.layout = layout
		p.slideW, p.slideH = layout.Dimensions()
	}
}

// NewPresentationBuilder creates a new presentation builder with options.
func NewPresentationBuilder(opts ...PptxOption) *PresentationBuilder {
	pres := NewPresentation()
	for _, opt := range opts {
		opt(pres)
	}
	return &PresentationBuilder{
		pres:   pres,
		errors: make([]error, 0),
	}
}

// AddSlide adds a new slide and returns a SlideBuilder.
func (b *PresentationBuilder) AddSlide() *SlideBuilder {
	slide := b.pres.AddSlide()
	return &SlideBuilder{
		slide:  slide,
		parent: b,
	}
}

// Build returns the built presentation.
func (b *PresentationBuilder) Build() (*Presentation, error) {
	if len(b.errors) > 0 {
		return nil, b.errors[0]
	}
	return b.pres, nil
}

// SlideBuilder provides a fluent API for building slides.
type SlideBuilder struct {
	slide  *Slide
	parent *PresentationBuilder
}

// AddText adds a text element and returns a TextBuilder for chaining.
func (sb *SlideBuilder) AddText(text string) *TextBuilder {
	ts := sb.slide.AddText(text)
	return &TextBuilder{
		textShape: ts,
		parent:    sb,
	}
}

// AddShape adds an auto shape and returns a ShapeBuilder for chaining.
func (sb *SlideBuilder) AddShape(shapeType ShapeType) *SlideShapeBuilder {
	builder := sb.slide.AddShape(shapeType)
	return &SlideShapeBuilder{
		builder: builder,
		parent:  sb,
	}
}

// AddLine adds a line between two points and returns a LineBuilder for chaining.
func (sb *SlideBuilder) AddLine(x1, y1, x2, y2 int64) *SlideLineBuilder {
	builder := sb.slide.AddLine(x1, y1, x2, y2)
	return &SlideLineBuilder{
		builder: builder,
		parent:  sb,
	}
}

// SetBackgroundColor sets the slide background color.
func (sb *SlideBuilder) SetBackgroundColor(color Color) *SlideBuilder {
	sb.slide.SetBackgroundColor(color)
	return sb
}

// AddSVG parses SVG data and converts its elements to PPTX shapes.
func (sb *SlideBuilder) AddSVG(svgData []byte) error {
	return sb.slide.AddSVG(svgData)
}

// AddSVGFile reads an SVG file and converts its elements to PPTX shapes.
func (sb *SlideBuilder) AddSVGFile(path string) error {
	return sb.slide.AddSVGFile(path)
}

// End returns to the PresentationBuilder.
func (sb *SlideBuilder) End() *PresentationBuilder {
	return sb.parent
}

// TextBuilder provides a fluent API for configuring text on slides.
type TextBuilder struct {
	textShape *TextShape
	parent    *SlideBuilder
}

// SetFontSize sets the font size in points.
func (tb *TextBuilder) SetFontSize(size int) *TextBuilder {
	tb.textShape.SetFontSize(size)
	return tb
}

// SetBold sets bold formatting.
func (tb *TextBuilder) SetBold(bold bool) *TextBuilder {
	tb.textShape.SetBold(bold)
	return tb
}

// SetItalic sets italic formatting.
func (tb *TextBuilder) SetItalic(italic bool) *TextBuilder {
	tb.textShape.SetItalic(italic)
	return tb
}

// SetUnderline sets underline formatting.
func (tb *TextBuilder) SetUnderline(underline bool) *TextBuilder {
	tb.textShape.SetUnderline(underline)
	return tb
}

// SetColor sets the text color.
func (tb *TextBuilder) SetColor(color Color) *TextBuilder {
	tb.textShape.SetColor(color)
	return tb
}

// SetFontFamily sets the font family.
func (tb *TextBuilder) SetFontFamily(font string) *TextBuilder {
	tb.textShape.SetFontFamily(font)
	return tb
}

// SetAlignment sets the text alignment.
func (tb *TextBuilder) SetAlignment(align Alignment) *TextBuilder {
	tb.textShape.SetAlignment(align)
	return tb
}

// SetPosition sets the position in EMUs.
func (tb *TextBuilder) SetPosition(x, y int64) *TextBuilder {
	tb.textShape.SetPosition(x, y)
	return tb
}

// SetSize sets the size in EMUs.
func (tb *TextBuilder) SetSize(w, h int64) *TextBuilder {
	tb.textShape.SetSize(w, h)
	return tb
}

// SetFillColor sets the background fill color.
func (tb *TextBuilder) SetFillColor(color Color) *TextBuilder {
	tb.textShape.SetFillColor(color)
	return tb
}

// SetLine sets the border line.
func (tb *TextBuilder) SetLine(color Color, widthPt int) *TextBuilder {
	tb.textShape.SetLine(color, widthPt)
	return tb
}

// End returns to the SlideBuilder.
func (tb *TextBuilder) End() *SlideBuilder {
	return tb.parent
}

// SlideShapeBuilder wraps ShapeBuilder to maintain the builder chain.
type SlideShapeBuilder struct {
	builder *ShapeBuilder
	parent  *SlideBuilder
}

// SetPosition sets the shape position.
func (ssb *SlideShapeBuilder) SetPosition(x, y int64) *SlideShapeBuilder {
	ssb.builder.SetPosition(x, y)
	return ssb
}

// SetSize sets the shape size.
func (ssb *SlideShapeBuilder) SetSize(w, h int64) *SlideShapeBuilder {
	ssb.builder.SetSize(w, h)
	return ssb
}

// SetFillColor sets the fill color.
func (ssb *SlideShapeBuilder) SetFillColor(color Color) *SlideShapeBuilder {
	ssb.builder.SetFillColor(color)
	return ssb
}

// SetLine sets the border line.
func (ssb *SlideShapeBuilder) SetLine(color Color, widthPt int) *SlideShapeBuilder {
	ssb.builder.SetLine(color, widthPt)
	return ssb
}

// SetNoFill makes the shape transparent.
func (ssb *SlideShapeBuilder) SetNoFill() *SlideShapeBuilder {
	ssb.builder.SetNoFill()
	return ssb
}

// SetNoLine removes the border.
func (ssb *SlideShapeBuilder) SetNoLine() *SlideShapeBuilder {
	ssb.builder.SetNoLine()
	return ssb
}

// End returns to the SlideBuilder.
func (ssb *SlideShapeBuilder) End() *SlideBuilder {
	return ssb.parent
}

// SlideLineBuilder wraps LineBuilder to maintain the builder chain.
type SlideLineBuilder struct {
	builder *LineBuilder
	parent  *SlideBuilder
}

// SetColor sets the line color.
func (slb *SlideLineBuilder) SetColor(color Color) *SlideLineBuilder {
	slb.builder.SetColor(color)
	return slb
}

// SetWidthPt sets the line width in points.
func (slb *SlideLineBuilder) SetWidthPt(pt float64) *SlideLineBuilder {
	slb.builder.SetWidthPt(pt)
	return slb
}

// SetArrowEnd adds an arrowhead at the end.
func (slb *SlideLineBuilder) SetArrowEnd() *SlideLineBuilder {
	slb.builder.SetArrowEnd()
	return slb
}

// SetArrowStart adds an arrowhead at the start.
func (slb *SlideLineBuilder) SetArrowStart() *SlideLineBuilder {
	slb.builder.SetArrowStart()
	return slb
}

// End returns to the SlideBuilder.
func (slb *SlideLineBuilder) End() *SlideBuilder {
	return slb.parent
}
