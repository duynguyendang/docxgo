# Credits & Project History

## Current Maintainer

**Duy Nguyen**  
GitHub: [@duynguyendang](https://github.com/duynguyendang)  
Role: PPTX Support, SVG Conversion, Current Maintainer

## Previous Maintainer

**Misael Monterroca**  
GitHub: [@mmonterroca](https://github.com/mmonterroca)  
Role: Original v2 Architect & Developer  
Original Repo: https://github.com/mmonterroca/docxgo

---

## Project Genealogy

```
2020-2022: gonfva/docxlib
           └─ Basic OOXML manipulation

2022-2023: fumiama/go-docx (fork)
           └─ Images, tables, shapes

2023-2025: mmonterroca/docxgo v1/v2
           └─ Clean architecture, production-grade DOCX
           └─ https://github.com/mmonterroca/docxgo

2025-2026: duynguyendang/docxgo (fork)
           └─ Added PPTX support
           └─ Added SVG to PPTX conversion
           └─ Added JSON architecture diagrams
```

---

## Major Contributions

### Duy Nguyen (2025-2026)

**PowerPoint (PPTX) Support**
- `pptx/` package with full slide creation
- Shapes: rectangle, ellipse, rounded rectangle, line
- Text with formatting (bold, italic, color, font size)
- Lines with arrowheads
- Builder pattern API matching DOCX style
- Multiple slide layouts (16:9, 16:10, 4:3)

**SVG to PPTX Conversion**
- SVG parser for rect, circle, path, polygon, text, groups
- Transform support (translate)
- Named color support
- ViewBox scaling
- Custom geometry generation

**JSON to PPTX**
- Architecture diagram generation from JSON
- Container and component layout
- Flow lines with arrows
- Color mapping

### Misael Monterroca (2023-2025)

**v2 Clean Architecture**
- Interface-based design with dependency injection
- Separation of concerns (domain, internal, pkg)
- Type-safe implementations
- Comprehensive error handling

**Core DOCX Features**
- Document, Paragraph, Run, Table interfaces
- Builder pattern with fluent API
- Headers, footers, page numbers
- Table of contents (TOC)
- Hyperlinks and fields
- Images (9 formats)
- Themes (7 presets)
- Mail merge / template engine

### fumiama (2022-2023)

- Images and picture handling
- Tables with complex structures
- Shapes and drawing objects
- Go module structure

### Gonzalo Fernández-Victorio (2020-2022)

- Original OOXML structure definitions
- Basic document parsing and writing
- ZIP-based .docx format support

---

## License

MIT License

```
Copyright (C) 2025-2026 Duy Nguyen (PPTX, SVG)
Copyright (C) 2024-2025 Misael Monterroca (v2 architecture)
Copyright (C) 2022-2024 fumiama (v1 enhancements)
Copyright (C) 2020-2022 Gonzalo Fernández-Victorio (original library)
```
