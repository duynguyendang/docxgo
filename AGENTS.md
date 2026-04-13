# Coding Agent Guide for docxgo

## Project Overview

**docxgo** is a pure Go library for creating Microsoft Office documents (`.docx` and `.pptx`). No external dependencies, no CGo.

- **Version**: 3.0.0
- **Go**: 1.23+
- **License**: MIT

---

## Architecture

### Package Structure

```
docxgo/
├── docx.go           # Public API: NewDocument(), OpenDocument(), NewField()
├── builder.go        # Fluent builder: DocumentBuilder, ParagraphBuilder, TableBuilder
├── options.go        # Options: WithTitle, WithAuthor, WithPageSize, etc.
├── domain/           # Interfaces (ports): Document, Paragraph, Run, Table, Section, Image
├── internal/
│   ├── core/         # Domain implementations (adapters)
│   ├── manager/      # ID, relationship, media, style managers
│   ├── serializer/  # Domain → XML conversion
│   ├── writer/       # XML → ZIP archive
│   ├── reader/       # Reading existing DOCX files
│   └── xml/          # OOXML XML structures
├── pkg/
│   ├── errors/       # Structured errors: DocxError, ValidationError
│   ├── color/        # Color utilities
│   └── constants/    # OOXML constants (twips, measurements)
└── pptx/             # PowerPoint creation (separate API)
```

### Flow

```
User Code → docx.go/builder.go (API) → domain/ (interfaces) → internal/core/ (implementations)
                                                                      ↓
                                                    internal/{serializer, writer, manager}/
```

### Key Patterns

| Pattern | Usage |
|---------|-------|
| **Ports/Adapters** | `domain/` = interfaces, `internal/core/` = implementations |
| **Builder** | `NewDocumentBuilder()` for fluent API |
| **Manager** | `IDGenerator`, `RelationshipManager`, `MediaManager` |
| **Error Accumulation** | Builders accumulate errors, surface at `Build()` |

---

## Key Files

| File | Purpose |
|------|---------|
| `docx.go` | `NewDocument()`, `OpenDocument()`, `NewField()` factory |
| `builder.go` | Fluent API: `DocumentBuilder.AddParagraph().Bold().End()` |
| `domain/document.go` | `Document` interface |
| `domain/paragraph.go` | `Paragraph` interface |
| `domain/run.go` | `Run` interface |
| `domain/table.go` | `Table`, `TableRow`, `TableCell` interfaces |
| `domain/section.go` | `Section`, `Header`, `Footer` interfaces |
| `internal/core/document.go` | Document implementation |
| `internal/serializer/` | Domain → XML → ZIP conversion |

---

## Common Tasks

### Create a Document

```go
doc := docx.NewDocument()
para, _ := doc.AddParagraph()
run, _ := para.AddRun()
run.SetText("Hello!")
run.SetBold(true)
doc.SaveAs("output.docx")
```

### Using Builder API

```go
builder := docx.NewDocumentBuilder()
builder.AddParagraph().
    Text("Title").
    Bold().
    FontSize(32).
    End()

doc, _ := builder.Build()
doc.SaveAs("output.docx")
```

### Create a Table

```go
table, _ := doc.AddTable(3, 2)  // 3 rows, 2 columns
table.SetStyle(docx.TableStyleGrid)

row0, _ := table.Row(0)
cell, _ := row0.Cell(0)
cellPara, _ := cell.AddParagraph()
cellRun, _ := cellPara.AddRun()
cellRun.SetText("Header")
```

### Add an Image

```go
para, _ := doc.AddParagraph()
img, _ := para.AddImage("photo.jpg")
size := domain.NewImageSizeInches(3.0, 2.0)
img.SetSize(size)
```

### Add Fields (TOC, Page Numbers)

```go
// Table of Contents
tocPara, _ := doc.AddParagraph()
toc, _ := tocPara.AddField(domain.FieldTypeTOC)

// Page number in footer
section, _ := doc.DefaultSection()
footer, _ := section.Footer(domain.FooterDefault)
footerPara, _ := footer.AddParagraph()
pageNum, _ := footerPara.AddField(domain.FieldTypePageNumber)
```

### Configure Page Layout

```go
section, _ := doc.DefaultSection()
section.SetPageSize(domain.PageSizeLetter)
section.SetOrientation(domain.OrientationLandscape)
section.SetMargins(domain.Margins{
    Top:    1440,  // 1 inch in twips (1440 twips = 1 inch)
    Right:  1440,
    Bottom: 1440,
    Left:   1440,
})
```

---

## Measurements

| Unit | Used For | Conversion |
|------|----------|------------|
| Twips | Margins, spacing, indents | 1440 twips = 1 inch |
| Half-points | Font sizes | 28 = 14pt |
| EMUs | Images | 914400 EMUs = 1 inch |
| DXA | Borders | 1/20 of a point |

```go
// 1 inch margins
margins := domain.Margins{Top: 1440, Right: 1440, Bottom: 1440, Left: 1440}

// 14pt font
run.SetSize(28)

// 3x2 inch image
size := domain.NewImageSizeInches(3.0, 2.0)
```

---

## Error Handling

Methods return errors for validation failures:

```go
para, err := doc.AddParagraph()
if err != nil {
    return fmt.Errorf("add paragraph: %w", err)
}

run, err := para.AddRun()
if err != nil {
    return fmt.Errorf("add run: %w", err)
}

if err := run.SetText("Hello"); err != nil {
    return fmt.Errorf("set text: %w", err)
}

if err := doc.Validate(); err != nil {
    return fmt.Errorf("validate: %w", err)
}
```

**Structured errors** in `pkg/errors`:
- `errors.NewValidationError(op, field, value, msg)`
- `errors.InvalidState(op, msg)`
- `errors.Unsupported(op, feature)`

---

## Thread Safety

**NOT thread-safe.** Each Document instance should be used by a single goroutine.

---

## Testing

| Package | Coverage |
|---------|----------|
| domain/ | 100% |
| pkg/errors/ | 100% |
| internal/core/ | ~54% |
| internal/serializer/ | ~57% |
| internal/writer/ | ~50% |

```bash
# Run all tests
go test ./...

# Run with coverage
go test -coverprofile=coverage.out ./...
go tool cover -html=coverage.out

# Run specific package
go test ./internal/core/... -v
```

---

## Build & Lint

```bash
# Build
go build ./...

# Format
go fmt ./...

# Vet
go vet ./...

# Lint (if golangci-lint configured)
golangci-lint run
```

---

## Adding New Features

### 1. Domain Interface (port)

Add interface in `domain/`:
```go
type MyFeature interface {
    DoSomething() error
}
```

### 2. Implementation (adapter)

Implement in `internal/core/`:
```go
type myFeature struct {
    // fields
}

func (m *myFeature) DoSomething() error {
    // implementation
    return nil
}
```

### 3. Wire up in document

In `internal/core/document.go`, create and return the implementation.

### 4. Add serializer

In `internal/serializer/`, convert the domain object to XML structures per OOXML spec.

### 5. Add tests

Create `*_test.go` alongside source with table-driven tests.

---

## Important Notes

- **License headers**: MIT. Do not introduce AGPL or other licenses.
- **No external deps**: Keep `go.mod` clean. Only standard library + internal packages.
- **OOXML compliance**: Generated files must be valid OOXML (Word/PowerPoint open them).
- **Error messages**: Use `strconv.Itoa()` or `fmt.Sprintf("%d", i)` for integers in strings, NOT `string(rune(i))`.

---

## Quick Reference

```go
// Create document
doc := docx.NewDocument()

// Add content
para, _ := doc.AddParagraph()
run, _ := para.AddRun()
run.SetText("text")
run.SetBold(true)
run.SetSize(28)  // 14pt
run.SetColor(docx.Blue)

// Table
table, _ := doc.AddTable(3, 3)
table.SetStyle(docx.TableStyleGrid)
row := table.Row(0)
cell := row.Cell(0)

// Save
doc.SaveAs("file.docx")
```
