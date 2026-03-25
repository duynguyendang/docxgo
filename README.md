# docxgo

Production-grade Microsoft Word `.docx` and PowerPoint `.pptx` file creation in Go.

[![Go Reference](https://pkg.go.dev/badge/github.com/duynguyendang/docxgo/v3.svg)](https://pkg.go.dev/github.com/duynguyendang/docxgo/v3)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/licenses/MIT)

## Overview

**docxgo** is a powerful, clean-architecture library for creating Microsoft Office documents in Go. Supports both Word (`.docx`) and PowerPoint (`.pptx`) files with a fluent builder API.

### Key Features

- ✅ **Word (.docx)** - Full document creation with tables, images, styles, themes
- ✅ **PowerPoint (.pptx)** - Slides with shapes, text, lines, arrows
- ✅ **SVG to PPTX** - Convert SVG elements to PowerPoint shapes
- ✅ **JSON to PPTX** - Generate architecture diagrams from JSON
- ✅ **Builder Pattern** - Fluent API for easy construction
- ✅ **Pure Go** - No external dependencies, no CGo
- ✅ **Open Source** - MIT License

---

## Installation

```bash
go get github.com/duynguyendang/docxgo/v3
```

### Requirements

- Go 1.23 or higher
- No external dependencies
- Works on Linux, macOS, Windows

---

## Quick Start

### Word Document

```go
package main

import (
    "log"
    docx "github.com/duynguyendang/docxgo/v3"
)

func main() {
    doc := docx.NewDocument()
    para, _ := doc.AddParagraph()
    run, _ := para.AddRun()
    run.SetText("Hello, World!")
    run.SetBold(true)
    doc.SaveAs("hello.docx")
}
```

### PowerPoint Presentation

```go
package main

import (
    "github.com/duynguyendang/docxgo/v3/pptx"
)

func main() {
    b := pptx.NewPresentationBuilder(
        pptx.WithTitle("My Presentation"),
        pptx.WithLayout(pptx.Layout16x9),
    )

    slide := b.AddSlide()
    slide.AddText("Hello PPTX").
        SetBold(true).
        SetFontSize(36).
        SetPosition(pptx.Inches(1), pptx.Inches(2)).
        SetSize(pptx.Inches(8), pptx.Inches(1)).
        End()

    slide.AddShape(pptx.ShapeRectangle).
        SetPosition(pptx.Inches(1), pptx.Inches(3)).
        SetSize(pptx.Inches(3), pptx.Inches(2)).
        SetFillColor(pptx.Blue).
        End()

    pres, _ := b.Build()
    pres.SaveAs("hello.pptx")
}
```

### SVG to PPTX

```go
package main

import (
    "os"
    "github.com/duynguyendang/docxgo/v3/pptx"
)

func main() {
    svgData, _ := os.ReadFile("diagram.svg")

    b := pptx.NewPresentationBuilder(pptx.WithLayout(pptx.Layout16x9))
    slide := b.AddSlide()
    slide.SetBackgroundColor(pptx.White)
    slide.AddSVG(svgData)

    pres, _ := b.Build()
    pres.SaveAs("diagram.pptx")
}
```

### JSON Architecture to PPTX

```go
package main

import (
    "encoding/json"
    "github.com/duynguyendang/docxgo/v3/pptx"
)

func main() {
    data := []byte(`{
        "canvas_size": {"width": 1000, "height": 562},
        "containers": [
            {"name": "CLIENT", "location": {"x": 10, "y": 100}, "size": {"width": 140, "height": 200}}
        ],
        "components": [
            {"name": "API", "location": {"x": 30, "y": 150}, "size": {"width": 100, "height": 50}}
        ],
        "flows": [
            {"from": "API", "to": "DB", "label": "Query", "type": "data"}
        ]
    }`)

    // Parse JSON and generate PPTX
    var arch Architecture
    json.Unmarshal(data, &arch)
    // ... generate PPTX from architecture
}
```

### Lines with Arrows

```go
slide.AddLine(
    pptx.Inches(1), pptx.Inches(2),  // start
    pptx.Inches(5), pptx.Inches(2),  // end
).
    SetColor(pptx.Color{R: 0, G: 120, B: 212}).
    SetWidthPt(1.5).
    SetArrowEnd().
    End()
```

---

## PPTX Features

| Feature | Status |
|---------|--------|
| Shapes (rect, ellipse, rounded rect) | ✅ |
| Text with formatting | ✅ |
| Lines with arrows | ✅ |
| SVG conversion | ✅ |
| Slide backgrounds | ✅ |
| Multiple layouts (16:9, 16:10, 4:3) | ✅ |
| Builder pattern API | ✅ |

### SVG Elements Supported

- `<rect>` - Rectangle shapes
- `<circle>` - Circles
- `<path>` - Path data (M, L, C, Z commands)
- `<polygon>` - Polygons
- `<line>` - Lines
- `<text>` - Text elements
- `<g>` - Groups with transforms

---

## Examples

See [`examples/`](examples/) directory:

### Word (DOCX)

| Example | Description |
|---------|-------------|
| 01_basic | Simple document with builder |
| 02_intermediate | Professional product catalog |
| 03_toc | Table of contents |
| 04_fields | TOC, page numbers, hyperlinks |
| 05_styles | 40+ built-in styles |
| 06_sections | Page layout control |
| 07_advanced | All features combined |
| 08_images | Image insertion |
| 09_advanced_tables | Cell merging, nested tables |
| 10_paragraph_spacing | Line and paragraph spacing |
| 11_multi_section | Multi-section layouts |
| 12_read_and_modify | Read and modify documents |
| 13_themes | Theme system (7 presets) |
| 14_mail_merge | Template engine |
| 15_external_template | MERGEFIELD support |

### PowerPoint (PPTX)

| Example | Description |
|---------|-------------|
| pptx | All PPTX features (shapes, text, SVG, JSON architecture, lines with arrows) |

---

## Architecture

```
github.com/duynguyendang/docxgo/v3/
├── domain/          # DOCX interfaces
├── internal/        # DOCX implementations
├── pkg/             # Utilities
├── pptx/            # PPTX package
│   ├── pptx.go      # Types, colors, layouts
│   ├── presentation.go # Presentation struct
│   ├── slide.go     # Slide and shapes
│   ├── shape.go     # Shape types
│   ├── builder.go   # Builder pattern
│   └── svg.go       # SVG parser and converter
├── themes/          # DOCX themes
└── examples/        # 18+ examples
```

---

## Error Handling

All operations return explicit errors:

```go
pres, err := builder.Build()
if err != nil {
    log.Fatal(err)
}

err := pres.SaveAs("output.pptx")
if err != nil {
    log.Fatal(err)
}
```

---

## Credits

This project is forked from [mmonterroca/docxgo](https://github.com/mmonterroca/docxgo) with added PPTX support and SVG conversion.

---

## License

MIT License - free for commercial and personal use.

```
Copyright (C) 2024-2026 Duy Nguyen
```

---

## Links

- **GitHub**: [github.com/duynguyendang/docxgo](https://github.com/duynguyendang/docxgo)
- **Go Reference**: [pkg.go.dev/github.com/duynguyendang/docxgo/v3](https://pkg.go.dev/github.com/duynguyendang/docxgo/v3)
- **Issues**: [GitHub Issues](https://github.com/duynguyendang/docxgo/issues)
