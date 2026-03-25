# PPTX Support for docxgo

## Overview

The `pptx` package adds PowerPoint (.pptx) file creation to docxgo using the same clean architecture and builder patterns.

## Quick Start

```go
import "github.com/mmonterroca/docxgo/v2/pptx"

// Create presentation with builder pattern
builder := pptx.NewPresentationBuilder(
    pptx.WithTitle("My Presentation"),
    pptx.WithAuthor("docxgo"),
    pptx.WithLayout(pptx.Layout16x9),
)

// Add slides with fluent API
builder.AddSlide().
    AddText("Hello World").
    SetBold(true).
    SetFontSize(44).
    SetColor(pptx.Blue).
    SetPosition(pptx.Inches(1), pptx.Inches(2)).
    SetSize(pptx.Inches(8), pptx.Inches(1.5)).
    SetAlignment(pptx.AlignmentCenter).
    End()

// Build and save
pres, _ := builder.Build()
pres.SaveAs("output.pptx")
```

## Features

### Layouts
- `Layout16x9` - 10" x 5.625" (widescreen)
- `Layout16x10` - 10" x 6.25"
- `Layout4x3` - 10" x 7.5" (standard)
- `LayoutWIDE` - 13.3" x 7.5"
- `LayoutA4` - A4 paper size

### Shapes
- Rectangle, Rounded Rectangle, Ellipse
- Triangle, Line, Arrow, Star, Diamond

### Text Formatting
- Bold, Italic, Underline
- Font size and family
- Text color
- Alignment (left, center, right, justify)
- Fill color (background)

### Measurement Helpers
- `pptx.Inches(1.5)` - Convert inches to EMUs
- `pptx.Cm(2.5)` - Convert centimeters to EMUs
- `pptx.Points(12)` - Convert points to EMUs

### Colors
Predefined: Black, White, Red, Green, Blue, Yellow, Cyan, Magenta, Orange, Purple, Gray, Navy, Teal, Maroon

Custom: `pptx.Color{R: 30, G: 39, B: 97}`

## API Reference

### PresentationBuilder
- `NewPresentationBuilder(opts...)` - Create builder with options
- `AddSlide()` - Add a new slide, returns SlideBuilder
- `Build()` - Returns (*Presentation, error)

### SlideBuilder
- `AddText(text)` - Add text, returns TextBuilder
- `AddShape(type)` - Add shape, returns ShapeBuilder
- `SetBackgroundColor(color)` - Set slide background
- `End()` - Return to PresentationBuilder

### TextBuilder
- `SetBold(bool)` - Set bold
- `SetItalic(bool)` - Set italic
- `SetUnderline(bool)` - Set underline
- `SetFontSize(int)` - Set size in points
- `SetFontFamily(string)` - Set font
- `SetColor(Color)` - Set text color
- `SetFillColor(Color)` - Set background color
- `SetAlignment(Alignment)` - Set alignment
- `SetPosition(x, y)` - Set position in EMUs
- `SetSize(w, h)` - Set size in EMUs
- `End()` - Return to SlideBuilder

## Example

See `examples/16_pptx/main.go` for a complete example.
