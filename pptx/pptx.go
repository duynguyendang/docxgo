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

// Package pptx provides PowerPoint (.pptx) file creation for go-docx v3.
//
// This package provides a clean, builder-pattern API for creating PowerPoint
// presentations following the Office Open XML (OOXML) specification.
//
// # Quick Start
//
// Create a simple presentation:
//
//	pres := pptx.NewPresentation()
//	slide := pres.AddSlide()
//	slide.AddText("Hello, World!").SetBold(true).SetFontSize(32)
//	pres.SaveAs("hello.pptx")
//
// # Builder Pattern
//
// Use the fluent builder API for complex presentations:
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
package pptx

// Layout represents presentation slide dimensions.
type Layout int

// Common slide layouts.
const (
	// Layout16x9 is 10" x 5.625" (default widescreen).
	Layout16x9 Layout = iota
	// Layout16x10 is 10" x 6.25".
	Layout16x10
	// Layout4x3 is 10" x 7.5" (standard).
	Layout4x3
	// LayoutWIDE is 13.3" x 7.5".
	LayoutWIDE
	// LayoutA4 is A4 paper size.
	LayoutA4
)

// Dimensions returns the width and height in EMUs (English Metric Units).
// 1 inch = 914,400 EMUs.
func (l Layout) Dimensions() (width, height int64) {
	switch l {
	case Layout16x9:
		return 12192000, 6858000
	case Layout16x10:
		return 12192000, 7620000
	case Layout4x3:
		return 12192000, 9144000
	case LayoutWIDE:
		return 16192500, 9144000
	case LayoutA4:
		return 16838000, 23811000
	default:
		return 12192000, 6858000
	}
}

// Name returns the human-readable name of the layout.
func (l Layout) Name() string {
	switch l {
	case Layout16x9:
		return "16:9 Widescreen"
	case Layout16x10:
		return "16:10"
	case Layout4x3:
		return "4:3 Standard"
	case LayoutWIDE:
		return "Widescreen"
	case LayoutA4:
		return "A4"
	default:
		return "16:9 Widescreen"
	}
}

// Color represents an RGB color for PPTX elements.
type Color struct {
	R uint8
	G uint8
	B uint8
}

// Hex returns the 6-character hex string representation (without #).
func (c Color) Hex() string {
	return string([]byte{
		hexChar(c.R >> 4), hexChar(c.R & 0x0F),
		hexChar(c.G >> 4), hexChar(c.G & 0x0F),
		hexChar(c.B >> 4), hexChar(c.B & 0x0F),
	})
}

func hexChar(b byte) byte {
	if b < 10 {
		return '0' + b
	}
	return 'A' + b - 10
}

// Common color constants.
var (
	Black   = Color{0, 0, 0}
	White   = Color{255, 255, 255}
	Red     = Color{255, 0, 0}
	Green   = Color{0, 128, 0}
	Blue    = Color{0, 0, 255}
	Yellow  = Color{255, 255, 0}
	Cyan    = Color{0, 255, 255}
	Magenta = Color{255, 0, 255}
	Orange  = Color{255, 165, 0}
	Purple  = Color{128, 0, 128}
	Gray    = Color{128, 128, 128}
	Navy    = Color{0, 0, 128}
	Teal    = Color{0, 128, 128}
	Maroon  = Color{128, 0, 0}
)

// Alignment represents text alignment within a shape.
type Alignment int

// Alignment constants.
const (
	AlignmentLeft Alignment = iota
	AlignmentCenter
	AlignmentRight
	AlignmentJustify
)

// ShapeType represents the type of auto shape.
type ShapeType int

// Shape type constants.
const (
	ShapeRectangle ShapeType = iota
	ShapeEllipse
	ShapeTriangle
	ShapeLine
	ShapeRoundedRectangle
	ShapeArrow
	ShapeStar
	ShapeDiamond
)

// shapeNames maps ShapeType to OOXML preset geometry names.
var shapeNames = map[ShapeType]string{
	ShapeRectangle:        "rect",
	ShapeEllipse:          "ellipse",
	ShapeTriangle:         "triangle",
	ShapeLine:             "line",
	ShapeRoundedRectangle: "roundRect",
	ShapeArrow:            "rightArrow",
	ShapeStar:             "star5",
	ShapeDiamond:          "diamond",
}

// ShapeName returns the OOXML preset geometry name.
func (s ShapeType) ShapeName() string {
	if name, ok := shapeNames[s]; ok {
		return name
	}
	return "rect"
}

// Measurement helpers for EMU values.
const (
	// EMU per inch.
	EmuPerInch = 914400
	// EMU per centimeter.
	EmuPerCm = 360000
	// EMU per point.
	EmuPerPoint = 12700
)

// Inches converts inches to EMUs.
func Inches(inches float64) int64 {
	return int64(inches * EmuPerInch)
}

// Cm converts centimeters to EMUs.
func Cm(cm float64) int64 {
	return int64(cm * EmuPerCm)
}

// Points converts points to EMUs.
func Points(pts float64) int64 {
	return int64(pts * EmuPerPoint)
}
