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

import (
	"encoding/xml"
	"fmt"
	"math"
	"os"
)

// Slide represents a single slide in the presentation.
type Slide struct {
	id      int
	pres    *Presentation
	shapes  []*Shape
	bgColor *Color
	images  []*SlideImage // Images added to this slide
}

// SlideImage represents an image embedded in a slide.
type SlideImage struct {
	data      []byte
	extension string // "svg", "png", "jpg", etc.
	id        int    // Shape ID
	x         int64
	y         int64
	w         int64
	h         int64
}

// SetBackgroundColor sets the slide background color.
func (s *Slide) SetBackgroundColor(color Color) *Slide {
	s.bgColor = &color
	return s
}

// AddText adds a text shape to the slide.
func (s *Slide) AddText(text string) *TextShape {
	shape := &TextShape{
		shape: &Shape{
			id:        s.pres.nextShapeID(),
			shapeType: ShapeRectangle,
			x:         Inches(1),
			y:         Inches(1),
			w:         Inches(8),
			h:         Inches(1),
			texts:     []TextRun{{Text: text}},
		},
	}
	s.shapes = append(s.shapes, shape.shape)
	return shape
}

// AddShape adds an auto shape to the slide.
func (s *Slide) AddShape(shapeType ShapeType) *ShapeBuilder {
	shape := &Shape{
		id:        s.pres.nextShapeID(),
		shapeType: shapeType,
		x:         Inches(1),
		y:         Inches(1),
		w:         Inches(2),
		h:         Inches(2),
		fillColor: nil,
		lineColor: nil,
		lineWidth: 0,
	}
	s.shapes = append(s.shapes, shape)
	return &ShapeBuilder{shape: shape}
}

// AddLine adds a line shape between two points.
func (s *Slide) AddLine(x1, y1, x2, y2 int64) *LineBuilder {
	shape := &Shape{
		id:        s.pres.nextShapeID(),
		shapeType: ShapeLine,
		x:         x1,
		y:         y1,
		w:         x2 - x1,
		h:         y2 - y1,
		isLine:    true,
		endX:      x2,
		endY:      y2,
		lineColor: &Color{R: 80, G: 80, B: 80},
		lineWidth: 9525, // ~1pt
	}
	s.shapes = append(s.shapes, shape)
	return &LineBuilder{shape: shape}
}

// AddImage adds an image shape to the slide from raw bytes.
func (s *Slide) AddImage(data []byte) *ImageBuilder {
	shape := &Shape{
		id:        s.pres.nextShapeID(),
		shapeType: ShapeRectangle, // Placeholder
		x:         Inches(1),
		y:         Inches(1),
		w:         Inches(4),
		h:         Inches(3),
		imageData: data,
	}
	s.shapes = append(s.shapes, shape)
	return &ImageBuilder{shape: shape}
}

// AddSVG parses an SVG document and converts its elements to PPTX shapes.
func (s *Slide) AddSVG(svgData []byte) error {
	doc, err := ParseSVG(svgData)
	if err != nil {
		return err
	}

	// Calculate scale to fit slide
	slideW := float64(s.pres.slideW) / 914400.0 // Convert EMU to inches
	slideH := float64(s.pres.slideH) / 914400.0

	// SVG viewBox dimensions in SVG units
	svgW := doc.ViewBox[2]
	svgH := doc.ViewBox[3]

	if svgW == 0 || svgH == 0 {
		svgW = doc.Width
		svgH = doc.Height
	}

	// Calculate scale factor (use smaller to fit, with 0.9 margin)
	margin := 0.9
	scaleX := (slideW * margin * 914400) / svgW
	scaleY := (slideH * margin * 914400) / svgH
	scale := math.Min(scaleX, scaleY)

	// Offset to center
	offsetX := (s.pres.slideW - int64(svgW*scale)) / 2
	offsetY := (s.pres.slideH - int64(svgH*scale)) / 2

	// Convert all SVG elements to PPTX shapes with scaling
	for _, elem := range doc.Elements {
		pptxShapes := elem.ToPPTXShapesScaled(&s.pres.shapeID, scale, offsetX, offsetY, doc.ViewBox)
		for _, ps := range pptxShapes {
			s.shapes = append(s.shapes, ps.toShape())
		}
	}

	return nil
}

// AddSVGFile reads an SVG file and adds its shapes to the slide.
func (s *Slide) AddSVGFile(path string) error {
	data, err := readFile(path)
	if err != nil {
		return fmt.Errorf("pptx: failed to read SVG file: %w", err)
	}
	return s.AddSVG(data)
}

func readFile(path string) ([]byte, error) {
	return os.ReadFile(path)
}

// Shapes returns all shapes on the slide.
func (s *Slide) Shapes() []*Shape {
	return s.shapes
}

// slideXML represents the XML structure of a slide.
type slideXML struct {
	XMLName   xml.Name  `xml:"p:sld"`
	XmlnsR    string    `xml:"xmlns:r,attr"`
	XmlnsP    string    `xml:"xmlns:p,attr"`
	XmlnsA    string    `xml:"xmlns:a,attr"`
	CSld      cSld      `xml:"p:cSld"`
	ClrMapOvr clrMapOvr `xml:"p:clrMapOvr"`
}

type clrMapOvr struct {
	MasterClrMapping masterClrMapping `xml:"a:masterClrMapping"`
}

type masterClrMapping struct {
	// Empty - inherits from master
}

type cSld struct {
	Bg     *bg    `xml:"p:bg,omitempty"`
	SpTree spTree `xml:"p:spTree"`
}

type bg struct {
	BgPr *bgPr `xml:"p:bgPr,omitempty"`
}

type bgPr struct {
	SolidFill *solidFill `xml:"a:solidFill,omitempty"`
}

type solidFill struct {
	SrgbClr *srgbClr `xml:"a:srgbClr,omitempty"`
}

type srgbClr struct {
	Val string `xml:"val,attr"`
}

type spTree struct {
	NvGrpSpPr nvGrpSpPr `xml:"p:nvGrpSpPr"`
	GrpSpPr   grpSpPr   `xml:"p:grpSpPr"`
	Sp        []spXML   `xml:"p:sp"`
}

type nvGrpSpPr struct {
	CNvPr      cNvPr      `xml:"p:cNvPr"`
	CNvGrpSpPr cNvGrpSpPr `xml:"p:cNvGrpSpPr"`
	NvPr       nvPr       `xml:"p:nvPr"`
}

type cNvPr struct {
	ID   int    `xml:"id,attr"`
	Name string `xml:"name,attr"`
}

type cNvGrpSpPr struct{}
type nvPr struct{}

type grpSpPr struct {
	Xfrm *xfrm `xml:"a:xfrm,omitempty"`
}

type xfrm struct {
	Off off `xml:"a:off"`
	Ext ext `xml:"a:ext"`
}

type off struct {
	X int64 `xml:"x,attr"`
	Y int64 `xml:"y,attr"`
}

type ext struct {
	Cx int64 `xml:"cx,attr"`
	Cy int64 `xml:"cy,attr"`
}

type spXML struct {
	NvSpPr nvSpPr   `xml:"p:nvSpPr"`
	SpPr   spPr     `xml:"p:spPr"`
	Style  *spStyle `xml:"p:style,omitempty"`
	TxBody *txBody  `xml:"p:txBody,omitempty"`
}

type spStyle struct {
	LnRef     styleRef `xml:"a:lnRef"`
	FillRef   styleRef `xml:"a:fillRef"`
	EffectRef styleRef `xml:"a:effectRef"`
	FontRef   fontRef  `xml:"a:fontRef"`
}

type styleRef struct {
	Idx       int       `xml:"idx,attr"`
	SchemeClr schemeClr `xml:"a:schemeClr"`
}

type schemeClr struct {
	Val string `xml:"val,attr"`
}

type fontRef struct {
	Idx       string    `xml:"idx,attr"`
	SchemeClr schemeClr `xml:"a:schemeClr"`
}

type nvSpPr struct {
	CNvPr   cNvPr   `xml:"p:cNvPr"`
	CNvSpPr cNvSpPr `xml:"p:cNvSpPr"`
	NvPr    nvPr    `xml:"p:nvPr"`
}

type cNvSpPr struct {
	TxBox bool `xml:"txBox,attr,omitempty"`
}

type spPr struct {
	Xfrm      *shapeXfrm   `xml:"a:xfrm,omitempty"`
	PrstGeom  *prstGeomXML `xml:"a:prstGeom,omitempty"`
	CustGeom  *custGeomXML `xml:"a:custGeom,omitempty"`
	SolidFill *solidFill   `xml:"a:solidFill,omitempty"`
	NoFill    *noFill      `xml:"a:noFill,omitempty"`
	Ln        *ln          `xml:"a:ln,omitempty"`
}

type noFill struct {
	// Empty - indicates no fill
}

type shapeXfrm struct {
	Off off `xml:"a:off"`
	Ext ext `xml:"a:ext"`
}

type prstGeomXML struct {
	Prst  string `xml:"prst,attr"`
	AvLst avLst  `xml:"a:avLst"`
}

type avLst struct {
	// Empty - no adjust values
}

type custGeomXML struct {
	AvLst   avLst      `xml:"a:avLst"`
	PathLst pathLstXML `xml:"a:pathLst"`
}

type pathLstXML struct {
	Path []pathXML `xml:"a:path"`
}

type pathXML struct {
	W    int64  `xml:"w,attr"`
	H    int64  `xml:"h,attr"`
	Path string `xml:",chardata"`
}

type ln struct {
	W         int        `xml:"w,attr,omitempty"`
	SolidFill *solidFill `xml:"a:solidFill,omitempty"`
	HeadEnd   *arrowEnd  `xml:"a:headEnd,omitempty"`
	TailEnd   *arrowEnd  `xml:"a:tailEnd,omitempty"`
}

type arrowEnd struct {
	Type string `xml:"type,attr"` // "triangle", "stealth", etc.
	W    string `xml:"w,attr"`    // "medium", "large", etc.
	Ln   string `xml:"len,attr"`  // "medium", "large", etc.
}

type txBody struct {
	BodyPr bodyPr `xml:"a:bodyPr"`
	P      []pXML `xml:"a:p"`
}

type bodyPr struct {
	Wrap   string `xml:"wrap,attr,omitempty"`
	Anchor string `xml:"anchor,attr,omitempty"`
}

type pXML struct {
	PPr *pPr   `xml:"a:pPr,omitempty"`
	R   []rXML `xml:"a:r"`
}

type pPr struct {
	Algn string `xml:"algn,attr,omitempty"`
}

type rXML struct {
	RPr rPr    `xml:"a:rPr"`
	T   string `xml:"a:t"`
}

type rPr struct {
	Lang      string     `xml:"lang,attr,omitempty"`
	Sz        int        `xml:"sz,attr,omitempty"`
	B         int        `xml:"b,attr,omitempty"`
	I         int        `xml:"i,attr,omitempty"`
	U         string     `xml:"u,attr,omitempty"`
	SolidFill *solidFill `xml:"a:solidFill,omitempty"`
	Latin     *latin     `xml:"a:latin,omitempty"`
}

type latin struct {
	Typeface string `xml:"typeface,attr"`
}

func (s *Slide) toXML() slideXML {
	sx := slideXML{
		XmlnsR: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
		XmlnsP: "http://schemas.openxmlformats.org/presentationml/2006/main",
		XmlnsA: "http://schemas.openxmlformats.org/drawingml/2006/main",
		CSld: cSld{
			SpTree: spTree{
				NvGrpSpPr: nvGrpSpPr{
					CNvPr: cNvPr{ID: 1, Name: ""},
				},
			},
		},
	}

	// Background color
	if s.bgColor != nil {
		sx.CSld.Bg = &bg{
			BgPr: &bgPr{
				SolidFill: &solidFill{
					SrgbClr: &srgbClr{Val: s.bgColor.Hex()},
				},
			},
		}
	}

	// Add shapes
	for _, shape := range s.shapes {
		sx.CSld.SpTree.Sp = append(sx.CSld.SpTree.Sp, shape.toXML())
	}

	return sx
}

func (sh *Shape) toXML() spXML {
	sp := spXML{
		NvSpPr: nvSpPr{
			CNvPr: cNvPr{
				ID:   sh.id,
				Name: fmt.Sprintf("Shape %d", sh.id),
			},
			CNvSpPr: cNvSpPr{TxBox: len(sh.texts) > 0},
		},
		SpPr: spPr{
			Xfrm: &shapeXfrm{
				Off: off{X: sh.x, Y: sh.y},
				Ext: ext{Cx: sh.w, Cy: sh.h},
			},
		},
		Style: &spStyle{
			LnRef:     styleRef{Idx: 1, SchemeClr: schemeClr{Val: "accent1"}},
			FillRef:   styleRef{Idx: 3, SchemeClr: schemeClr{Val: "accent1"}},
			EffectRef: styleRef{Idx: 2, SchemeClr: schemeClr{Val: "accent1"}},
			FontRef:   fontRef{Idx: "minor", SchemeClr: schemeClr{Val: "lt1"}},
		},
	}

	// Geometry - either custom or preset
	if sh.custGeom != nil {
		// Custom geometry from SVG path
		sp.SpPr.CustGeom = sh.custGeom.toXML()
	} else {
		// Preset geometry
		prstName := sh.shapeType.ShapeName()
		if sh.prstGeom != "" {
			prstName = sh.prstGeom
		}
		sp.SpPr.PrstGeom = &prstGeomXML{
			Prst:  prstName,
			AvLst: avLst{},
		}
	}

	// Fill color
	if sh.fillColor != nil {
		sp.SpPr.SolidFill = &solidFill{
			SrgbClr: &srgbClr{Val: sh.fillColor.Hex()},
		}
	}

	// Line/border
	if sh.lineColor != nil || sh.isLine {
		w := sh.lineWidth
		if w == 0 {
			w = 9525 // ~1pt default
		}
		lineColor := sh.lineColor
		if lineColor == nil {
			lineColor = &Color{R: 80, G: 80, B: 80}
		}
		lineLn := &ln{
			W: w,
			SolidFill: &solidFill{
				SrgbClr: &srgbClr{Val: lineColor.Hex()},
			},
		}

		// Add arrowheads for line shapes
		if sh.arrowStart {
			lineLn.HeadEnd = &arrowEnd{
				Type: "triangle",
				W:    "medium",
				Ln:   "medium",
			}
		}
		if sh.arrowEnd {
			lineLn.TailEnd = &arrowEnd{
				Type: "triangle",
				W:    "medium",
				Ln:   "medium",
			}
		}

		sp.SpPr.Ln = lineLn

		// For line shapes, no fill
		if sh.isLine {
			sp.SpPr.NoFill = &noFill{}
		}
	}

	// Text body
	if len(sh.texts) > 0 {
		tx := &txBody{
			BodyPr: bodyPr{Wrap: "square"},
		}

		// Group consecutive texts with same formatting into runs
		p := pXML{}
		for _, t := range sh.texts {
			r := rXML{
				RPr: rPr{
					Lang: "en-US",
					Sz:   t.FontSize * 100, // Font size in hundredths of point
					B:    boolToInt(t.Bold),
					I:    boolToInt(t.Italic),
					U:    underlineStyle(t.Underline),
				},
				T: t.Text,
			}

			if t.Color != nil {
				r.RPr.SolidFill = &solidFill{
					SrgbClr: &srgbClr{Val: t.Color.Hex()},
				}
			}

			if t.FontFamily != "" {
				r.RPr.Latin = &latin{Typeface: t.FontFamily}
			}

			p.R = append(p.R, r)
		}

		// Alignment
		if sh.alignment != AlignmentLeft {
			alignStr := "l"
			switch sh.alignment {
			case AlignmentCenter:
				alignStr = "ctr"
			case AlignmentRight:
				alignStr = "r"
			case AlignmentJustify:
				alignStr = "just"
			}
			p.PPr = &pPr{Algn: alignStr}
		}

		tx.P = []pXML{p}
		sp.TxBody = tx
	}

	return sp
}

func boolToInt(b bool) int {
	if b {
		return 1
	}
	return 0
}

func underlineStyle(u bool) string {
	if u {
		return "sng"
	}
	return ""
}
