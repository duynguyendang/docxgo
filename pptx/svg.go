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
	"strconv"
	"strings"
)

// SVGDocument represents a parsed SVG document.
type SVGDocument struct {
	Width    float64
	Height   float64
	ViewBox  [4]float64
	Elements []SVGElement
}

// SVGElement represents an SVG element that can be converted to PPTX.
type SVGElement interface {
	ToPPTXShapes(id *int) []pptxShape
	ToPPTXShapesScaled(id *int, scale float64, offsetX, offsetY int64, viewBox [4]float64) []pptxShape
}

// SVGRect represents an SVG rectangle.
type SVGRect struct {
	X, Y, Width, Height float64
	Fill                string
	Stroke              string
	StrokeWidth         float64
	Transform           string
}

// SVGCircle represents an SVG circle.
type SVGCircle struct {
	CX, CY, R   float64
	Fill        string
	Stroke      string
	StrokeWidth float64
	Transform   string
}

// SVGPath represents an SVG path element.
type SVGPath struct {
	D           string
	Fill        string
	Stroke      string
	StrokeWidth float64
	Transform   string
}

// SVGGroup represents an SVG group element.
type SVGGroup struct {
	Transform string
	Elements  []SVGElement
}

// pptxShape is an internal representation of a shape ready for PPTX output.
type pptxShape struct {
	ID        int
	Name      string
	X, Y      int64 // EMUs
	W, H      int64 // EMUs
	Fill      *string
	Stroke    *string
	StrokeW   int64
	CustGeom  *customGeom // Custom geometry for paths
	PrstGeom  string      // Preset geometry name
	Text      string      // Text content
	FontSize  int         // Font size in hundredths of point
	Bold      bool        // Bold text
	TextColor *string     // Text color (hex)
}

// customGeom represents PPTX custom geometry with path data.
type customGeom struct {
	Paths []geomPath
}

// geomPath represents a path within custom geometry.
type geomPath struct {
	W, H int64  // Path size in EMUs
	Path string // GDML path commands
}

// svgRoot represents the root SVG element for XML parsing.
type svgRoot struct {
	XMLName  xml.Name        `xml:"svg"`
	Width    string          `xml:"width,attr"`
	Height   string          `xml:"height,attr"`
	ViewBox  string          `xml:"viewBox,attr"`
	Rects    []svgRectXML    `xml:"rect"`
	Paths    []svgPathXML    `xml:"path"`
	Circles  []svgCircleXML  `xml:"circle"`
	Ellipses []svgEllipseXML `xml:"ellipse"`
	Lines    []svgLineXML    `xml:"line"`
	Polys    []svgPolyXML    `xml:"polygon"`
	Texts    []svgTextXML    `xml:"text"`
	Groups   []svgGroupXML   `xml:"g"`
}

type svgRectXML struct {
	X         string `xml:"x,attr"`
	Y         string `xml:"y,attr"`
	Width     string `xml:"width,attr"`
	Height    string `xml:"height,attr"`
	Fill      string `xml:"fill,attr"`
	Stroke    string `xml:"stroke,attr"`
	StrokeW   string `xml:"stroke-width,attr"`
	Transform string `xml:"transform,attr"`
}

type svgPathXML struct {
	D         string `xml:"d,attr"`
	Fill      string `xml:"fill,attr"`
	Stroke    string `xml:"stroke,attr"`
	StrokeW   string `xml:"stroke-width,attr"`
	Transform string `xml:"transform,attr"`
}

type svgCircleXML struct {
	CX        string `xml:"cx,attr"`
	CY        string `xml:"cy,attr"`
	R         string `xml:"r,attr"`
	Fill      string `xml:"fill,attr"`
	Stroke    string `xml:"stroke,attr"`
	StrokeW   string `xml:"stroke-width,attr"`
	Transform string `xml:"transform,attr"`
}

type svgEllipseXML struct {
	CX        string `xml:"cx,attr"`
	CY        string `xml:"cy,attr"`
	RX        string `xml:"rx,attr"`
	RY        string `xml:"ry,attr"`
	Fill      string `xml:"fill,attr"`
	Stroke    string `xml:"stroke,attr"`
	StrokeW   string `xml:"stroke-width,attr"`
	Transform string `xml:"transform,attr"`
}

type svgLineXML struct {
	X1        string `xml:"x1,attr"`
	Y1        string `xml:"y1,attr"`
	X2        string `xml:"x2,attr"`
	Y2        string `xml:"y2,attr"`
	Stroke    string `xml:"stroke,attr"`
	StrokeW   string `xml:"stroke-width,attr"`
	Transform string `xml:"transform,attr"`
}

type svgPolyXML struct {
	Points    string `xml:"points,attr"`
	Fill      string `xml:"fill,attr"`
	Stroke    string `xml:"stroke,attr"`
	StrokeW   string `xml:"stroke-width,attr"`
	Transform string `xml:"transform,attr"`
}

type svgTextXML struct {
	X          string `xml:"x,attr"`
	Y          string `xml:"y,attr"`
	FontSize   string `xml:"font-size,attr"`
	FontWeight string `xml:"font-weight,attr"`
	Fill       string `xml:"fill,attr"`
	Anchor     string `xml:"text-anchor,attr"`
	Content    string `xml:",chardata"`
	Transform  string `xml:"transform,attr"`
}

type svgGroupXML struct {
	Transform string          `xml:"transform,attr"`
	Rects     []svgRectXML    `xml:"rect"`
	Paths     []svgPathXML    `xml:"path"`
	Circles   []svgCircleXML  `xml:"circle"`
	Ellipses  []svgEllipseXML `xml:"ellipse"`
	Lines     []svgLineXML    `xml:"line"`
	Polys     []svgPolyXML    `xml:"polygon"`
}

// ParseSVG parses SVG XML bytes and returns an SVGDocument.
func ParseSVG(data []byte) (*SVGDocument, error) {
	var root svgRoot
	if err := xml.Unmarshal(data, &root); err != nil {
		return nil, fmt.Errorf("pptx: failed to parse SVG: %w", err)
	}

	doc := &SVGDocument{}

	// Parse dimensions
	doc.Width = parseSVGFloat(root.Width)
	doc.Height = parseSVGFloat(root.Height)

	// Parse viewBox
	if root.ViewBox != "" {
		parts := strings.Fields(root.ViewBox)
		if len(parts) == 4 {
			doc.ViewBox = [4]float64{
				parseSVGFloat(parts[0]),
				parseSVGFloat(parts[1]),
				parseSVGFloat(parts[2]),
				parseSVGFloat(parts[3]),
			}
		}
	}

	// If no viewBox, use dimensions
	if doc.ViewBox[2] == 0 {
		doc.ViewBox = [4]float64{0, 0, doc.Width, doc.Height}
	}

	// Parse elements
	for _, r := range root.Rects {
		doc.Elements = append(doc.Elements, parseRect(r))
	}
	for _, p := range root.Paths {
		doc.Elements = append(doc.Elements, parsePath(p))
	}
	for _, c := range root.Circles {
		doc.Elements = append(doc.Elements, parseCircle(c))
	}
	for _, e := range root.Ellipses {
		doc.Elements = append(doc.Elements, parseEllipse(e))
	}
	for _, l := range root.Lines {
		doc.Elements = append(doc.Elements, parseLine(l))
	}
	for _, p := range root.Polys {
		doc.Elements = append(doc.Elements, parsePolygon(p))
	}
	for _, t := range root.Texts {
		doc.Elements = append(doc.Elements, parseText(t))
	}

	// Parse groups
	for _, g := range root.Groups {
		group := parseGroup(g)
		if len(group.Elements) > 0 {
			doc.Elements = append(doc.Elements, group)
		}
	}

	return doc, nil
}

func parseRect(r svgRectXML) SVGRect {
	return SVGRect{
		X:           parseSVGFloat(r.X),
		Y:           parseSVGFloat(r.Y),
		Width:       parseSVGFloat(r.Width),
		Height:      parseSVGFloat(r.Height),
		Fill:        r.Fill,
		Stroke:      r.Stroke,
		StrokeWidth: parseSVGFloat(r.StrokeW),
		Transform:   r.Transform,
	}
}

func parsePath(p svgPathXML) SVGPath {
	return SVGPath{
		D:           p.D,
		Fill:        p.Fill,
		Stroke:      p.Stroke,
		StrokeWidth: parseSVGFloat(p.StrokeW),
		Transform:   p.Transform,
	}
}

func parseCircle(c svgCircleXML) SVGCircle {
	return SVGCircle{
		CX:          parseSVGFloat(c.CX),
		CY:          parseSVGFloat(c.CY),
		R:           parseSVGFloat(c.R),
		Fill:        c.Fill,
		Stroke:      c.Stroke,
		StrokeWidth: parseSVGFloat(c.StrokeW),
		Transform:   c.Transform,
	}
}

func parseEllipse(e svgEllipseXML) SVGElement {
	// Convert ellipse to path
	rx := parseSVGFloat(e.RX)
	ry := parseSVGFloat(e.RY)
	cx := parseSVGFloat(e.CX)
	cy := parseSVGFloat(e.CY)

	// Approximate ellipse with bezier curves
	// M cx-rx,cy C cx-rx,cy-ry*0.552 cx-rx*0.552,cy-ry cx,cy-ry ...
	k := 0.5522847498
	d := fmt.Sprintf("M %g,%g C %g,%g %g,%g %g,%g C %g,%g %g,%g %g,%g C %g,%g %g,%g %g,%g C %g,%g %g,%g %g,%g Z",
		cx-rx, cy,
		cx-rx, cy-ry*k, cx-rx*k, cy-ry, cx, cy-ry,
		cx+rx*k, cy-ry, cx+rx, cy-ry*k, cx+rx, cy,
		cx+rx, cy+ry*k, cx+rx*k, cy+ry, cx, cy+ry,
		cx-rx*k, cy+ry, cx-rx, cy+ry*k, cx-rx, cy,
	)

	return SVGPath{
		D:           d,
		Fill:        e.Fill,
		Stroke:      e.Stroke,
		StrokeWidth: parseSVGFloat(e.StrokeW),
		Transform:   e.Transform,
	}
}

func parseLine(l svgLineXML) SVGPath {
	x1 := parseSVGFloat(l.X1)
	y1 := parseSVGFloat(l.Y1)
	x2 := parseSVGFloat(l.X2)
	y2 := parseSVGFloat(l.Y2)

	return SVGPath{
		D:           fmt.Sprintf("M %g,%g L %g,%g", x1, y1, x2, y2),
		Fill:        "none",
		Stroke:      l.Stroke,
		StrokeWidth: parseSVGFloat(l.StrokeW),
		Transform:   l.Transform,
	}
}

func parsePolygon(p svgPolyXML) SVGPath {
	points := strings.FieldsFunc(p.Points, func(r rune) bool {
		return r == ' ' || r == ','
	})

	var d strings.Builder
	for i := 0; i+1 < len(points); i += 2 {
		if i == 0 {
			d.WriteString("M ")
		} else {
			d.WriteString(" L ")
		}
		d.WriteString(points[i])
		d.WriteString(",")
		d.WriteString(points[i+1])
	}
	d.WriteString(" Z")

	return SVGPath{
		D:           d.String(),
		Fill:        p.Fill,
		Stroke:      p.Stroke,
		StrokeWidth: parseSVGFloat(p.StrokeW),
		Transform:   p.Transform,
	}
}

// SVGText represents an SVG text element.
type SVGText struct {
	X          float64
	Y          float64
	Content    string
	FontSize   float64
	FontWeight string
	Fill       string
	Anchor     string
	Transform  string
}

func parseText(t svgTextXML) SVGText {
	return SVGText{
		X:          parseSVGFloat(t.X),
		Y:          parseSVGFloat(t.Y),
		Content:    strings.TrimSpace(t.Content),
		FontSize:   parseSVGFloat(t.FontSize),
		FontWeight: t.FontWeight,
		Fill:       t.Fill,
		Anchor:     t.Anchor,
		Transform:  t.Transform,
	}
}

// ToPPTXShapes converts SVGText to PPTX shapes.
func (t SVGText) ToPPTXShapes(id *int) []pptxShape {
	return t.ToPPTXShapesScaled(id, 1, 0, 0, [4]float64{0, 0, 0, 0})
}

// ToPPTXShapesScaled converts SVGText to PPTX text shapes with scaling.
func (t SVGText) ToPPTXShapesScaled(id *int, scale float64, offsetX, offsetY int64, viewBox [4]float64) []pptxShape {
	if t.Content == "" {
		return nil
	}

	*id++

	// Calculate position - text y is baseline, adjust to top
	fontSize := t.FontSize
	if fontSize == 0 {
		fontSize = 12
	}
	x := int64((t.X-viewBox[0])*scale) + offsetX
	y := int64((t.Y-viewBox[1])*scale) + offsetY - int64(fontSize*scale)

	// Estimate text box size
	w := int64(fontSize * float64(len(t.Content)) * 0.65 * scale)
	h := int64(fontSize * 1.4 * scale)

	// Adjust for anchor
	if t.Anchor == "middle" {
		x -= w / 2
	}

	// Text color from fill attribute
	var textColor *string
	fillColor := parseColor(t.Fill)
	if fillColor != "" {
		textColor = &fillColor
	}

	return []pptxShape{{
		ID:   *id,
		Name: fmt.Sprintf("Text %d", *id),
		X:    x,
		Y:    y,
		W:    w,
		H:    h,
		// No fill - text boxes should be transparent
		Text:      t.Content,
		FontSize:  int(fontSize * 100), // hundredths of point
		Bold:      t.FontWeight == "bold",
		TextColor: textColor,
	}}
}

func parseGroup(g svgGroupXML) SVGGroup {
	group := SVGGroup{
		Transform: g.Transform,
	}

	// Parse translate from transform string
	tx, ty := parseTranslate(g.Transform)

	for _, r := range g.Rects {
		r.X = fmt.Sprintf("%g", parseSVGFloat(r.X)+tx)
		r.Y = fmt.Sprintf("%g", parseSVGFloat(r.Y)+ty)
		r.Transform = "" // Clear transform since we applied it
		group.Elements = append(group.Elements, parseRect(r))
	}
	for _, p := range g.Paths {
		// Apply translate to all path coordinates
		if tx != 0 || ty != 0 {
			p.D = applyTranslateToPath(p.D, tx, ty)
		}
		p.Transform = ""
		group.Elements = append(group.Elements, parsePath(p))
	}
	for _, c := range g.Circles {
		c.CX = fmt.Sprintf("%g", parseSVGFloat(c.CX)+tx)
		c.CY = fmt.Sprintf("%g", parseSVGFloat(c.CY)+ty)
		c.Transform = ""
		group.Elements = append(group.Elements, parseCircle(c))
	}
	for _, e := range g.Ellipses {
		e.CX = fmt.Sprintf("%g", parseSVGFloat(e.CX)+tx)
		e.CY = fmt.Sprintf("%g", parseSVGFloat(e.CY)+ty)
		e.Transform = ""
		group.Elements = append(group.Elements, parseEllipse(e))
	}
	for _, l := range g.Lines {
		l.X1 = fmt.Sprintf("%g", parseSVGFloat(l.X1)+tx)
		l.Y1 = fmt.Sprintf("%g", parseSVGFloat(l.Y1)+ty)
		l.X2 = fmt.Sprintf("%g", parseSVGFloat(l.X2)+tx)
		l.Y2 = fmt.Sprintf("%g", parseSVGFloat(l.Y2)+ty)
		l.Transform = ""
		group.Elements = append(group.Elements, parseLine(l))
	}
	for _, p := range g.Polys {
		// Apply translate to polygon points
		if tx != 0 || ty != 0 {
			points := strings.FieldsFunc(p.Points, func(r rune) bool {
				return r == ' ' || r == ','
			})
			var newPoints []string
			for i := 0; i+1 < len(points); i += 2 {
				x := parseSVGFloat(points[i]) + tx
				y := parseSVGFloat(points[i+1]) + ty
				newPoints = append(newPoints, fmt.Sprintf("%g,%g", x, y))
			}
			p.Points = strings.Join(newPoints, " ")
		}
		p.Transform = ""
		group.Elements = append(group.Elements, parsePolygon(p))
	}

	return group
}

// parseTranslate extracts translate(x, y) from a transform string
func parseTranslate(transform string) (float64, float64) {
	if transform == "" {
		return 0, 0
	}

	// Look for translate(x, y) or translate(x y)
	idx := strings.Index(transform, "translate(")
	if idx == -1 {
		return 0, 0
	}

	rest := transform[idx+len("translate("):]
	endIdx := strings.Index(rest, ")")
	if endIdx == -1 {
		return 0, 0
	}

	coords := strings.FieldsFunc(rest[:endIdx], func(r rune) bool {
		return r == ',' || r == ' '
	})

	if len(coords) >= 2 {
		return parseSVGFloat(coords[0]), parseSVGFloat(coords[1])
	} else if len(coords) == 1 {
		return parseSVGFloat(coords[0]), 0
	}

	return 0, 0
}

// applyTranslateToPath applies a translate transform to all coordinates in an SVG path
func applyTranslateToPath(d string, tx, ty float64) string {
	tokens := tokenizePath(d)
	var result strings.Builder

	for _, tok := range tokens {
		cmd := tok.Cmd
		args := tok.Args

		result.WriteByte(byte(cmd))
		result.WriteByte(' ')

		switch cmd {
		case 'M', 'm', 'L', 'l', 'T', 't':
			// x y pairs
			for i := 0; i+1 < len(args); i += 2 {
				x := args[i] + tx
				y := args[i+1] + ty
				result.WriteString(fmt.Sprintf("%g %g ", x, y))
			}
		case 'H', 'h':
			// x only
			for _, x := range args {
				result.WriteString(fmt.Sprintf("%g ", x+tx))
			}
		case 'V', 'v':
			// y only
			for _, y := range args {
				result.WriteString(fmt.Sprintf("%g ", y+ty))
			}
		case 'C', 'c':
			// x1 y1 x2 y2 x y (6 values)
			for i := 0; i+5 < len(args); i += 6 {
				result.WriteString(fmt.Sprintf("%g %g %g %g %g %g ",
					args[i]+tx, args[i+1]+ty,
					args[i+2]+tx, args[i+3]+ty,
					args[i+4]+tx, args[i+5]+ty))
			}
		case 'S', 's', 'Q', 'q':
			// x2 y2 x y (4 values)
			for i := 0; i+3 < len(args); i += 4 {
				result.WriteString(fmt.Sprintf("%g %g %g %g ",
					args[i]+tx, args[i+1]+ty,
					args[i+2]+tx, args[i+3]+ty))
			}
		case 'A', 'a':
			// rx ry x-rotation large-arc-flag sweep-flag x y (7 values)
			for i := 0; i+6 < len(args); i += 7 {
				result.WriteString(fmt.Sprintf("%g %g %g %g %g %g %g ",
					args[i], args[i+1], args[i+2],
					args[i+3], args[i+4],
					args[i+5]+tx, args[i+6]+ty))
			}
		case 'Z', 'z':
			// No arguments
		default:
			// Pass through unknown commands
			for _, arg := range args {
				result.WriteString(fmt.Sprintf("%g ", arg))
			}
		}
	}

	return result.String()
}

func combineTransforms(parent, child string) string {
	if parent == "" {
		return child
	}
	if child == "" {
		return parent
	}
	return parent + " " + child
}

func parseSVGFloat(s string) float64 {
	if s == "" {
		return 0
	}
	// Remove any unit suffix
	s = strings.TrimSpace(s)
	s = strings.TrimSuffix(s, "px")
	s = strings.TrimSuffix(s, "pt")
	s = strings.TrimSuffix(s, "em")
	s = strings.TrimSuffix(s, "%")

	val, _ := strconv.ParseFloat(s, 64)
	return val
}

func parseColor(s string) string {
	if s == "" || s == "none" || s == "transparent" {
		return ""
	}

	// Handle SVG named colors
	namedColors := map[string]string{
		"black":   "000000",
		"white":   "FFFFFF",
		"red":     "FF0000",
		"green":   "008000",
		"blue":    "0000FF",
		"yellow":  "FFFF00",
		"cyan":    "00FFFF",
		"magenta": "FF00FF",
		"orange":  "FFA500",
		"purple":  "800080",
		"gray":    "808080",
		"grey":    "808080",
		"navy":    "000080",
		"teal":    "008080",
		"maroon":  "800000",
		"lime":    "00FF00",
		"aqua":    "00FFFF",
		"silver":  "C0C0C0",
		"olive":   "808000",
		"fuchsia": "FF00FF",
	}

	if hex, ok := namedColors[strings.ToLower(s)]; ok {
		return hex
	}

	// Remove # prefix
	return strings.TrimPrefix(s, "#")
}

// SVG to PPTX conversion

// SVGToEMU converts SVG coordinates to EMU (English Metric Units).
// 1 SVG unit = 9525 EMUs (approximate for screen resolution)
func SVGToEMU(svg float64) int64 {
	return int64(math.Round(svg * 9525))
}

// ToPPTXShapes converts SVGRect to PPTX shapes.
func (r SVGRect) ToPPTXShapes(id *int) []pptxShape {
	*id++
	fill := parseColor(r.Fill)
	stroke := parseColor(r.Stroke)
	strokeW := SVGToEMU(r.StrokeWidth)

	shape := pptxShape{
		ID:       *id,
		Name:     fmt.Sprintf("Rect %d", *id),
		X:        SVGToEMU(r.X),
		Y:        SVGToEMU(r.Y),
		W:        SVGToEMU(r.Width),
		H:        SVGToEMU(r.Height),
		PrstGeom: "rect",
	}

	if fill != "" {
		shape.Fill = &fill
	}
	if stroke != "" {
		shape.Stroke = &stroke
		shape.StrokeW = strokeW
	}

	return []pptxShape{shape}
}

// ToPPTXShapes converts SVGCircle to PPTX shapes.
func (c SVGCircle) ToPPTXShapes(id *int) []pptxShape {
	*id++
	fill := parseColor(c.Fill)
	stroke := parseColor(c.Stroke)
	strokeW := SVGToEMU(c.StrokeWidth)

	shape := pptxShape{
		ID:       *id,
		Name:     fmt.Sprintf("Circle %d", *id),
		X:        SVGToEMU(c.CX - c.R),
		Y:        SVGToEMU(c.CY - c.R),
		W:        SVGToEMU(c.R * 2),
		H:        SVGToEMU(c.R * 2),
		PrstGeom: "ellipse",
	}

	if fill != "" {
		shape.Fill = &fill
	}
	if stroke != "" {
		shape.Stroke = &stroke
		shape.StrokeW = strokeW
	}

	return []pptxShape{shape}
}

// ToPPTXShapes converts SVGPath to PPTX shapes with custom geometry.
func (p SVGPath) ToPPTXShapes(id *int) []pptxShape {
	*id++
	fill := parseColor(p.Fill)
	stroke := parseColor(p.Stroke)
	strokeW := SVGToEMU(p.StrokeWidth)

	// Parse path to get bounding box and convert to GDML
	gdml, bounds := convertSVGPathToGDML(p.D)

	shape := pptxShape{
		ID:   *id,
		Name: fmt.Sprintf("Path %d", *id),
		X:    SVGToEMU(bounds.X),
		Y:    SVGToEMU(bounds.Y),
		W:    SVGToEMU(bounds.W),
		H:    SVGToEMU(bounds.H),
		CustGeom: &customGeom{
			Paths: []geomPath{
				{
					W:    SVGToEMU(bounds.W),
					H:    SVGToEMU(bounds.H),
					Path: gdml,
				},
			},
		},
	}

	if fill != "" {
		shape.Fill = &fill
	}
	if stroke != "" {
		shape.Stroke = &stroke
		shape.StrokeW = strokeW
	}

	return []pptxShape{shape}
}

// ToPPTXShapes converts SVGGroup to PPTX shapes.
func (g SVGGroup) ToPPTXShapes(id *int) []pptxShape {
	var shapes []pptxShape
	for _, elem := range g.Elements {
		shapes = append(shapes, elem.ToPPTXShapes(id)...)
	}
	return shapes
}

type bbox struct {
	X, Y, W, H float64
}

// convertSVGPathToGDML converts SVG path data to PPTX GDML format.
func convertSVGPathToGDML(svgPath string) (string, bbox) {
	// Tokenize the path
	tokens := tokenizePath(svgPath)

	var gdml strings.Builder
	minX, minY := math.MaxFloat64, math.MaxFloat64
	maxX, maxY := -math.MaxFloat64, -math.MaxFloat64

	for _, tok := range tokens {
		cmd := tok.Cmd
		args := tok.Args

		switch cmd {
		case 'M', 'm':
			if len(args) >= 2 {
				x, y := args[0], args[1]
				if cmd == 'm' {
					gdml.WriteString(fmt.Sprintf("M %d %d ", SVGToEMU(x), SVGToEMU(y)))
				} else {
					gdml.WriteString(fmt.Sprintf("M %d %d ", SVGToEMU(x), SVGToEMU(y)))
				}
				updateBounds(&minX, &minY, &maxX, &maxY, x, y)
			}
		case 'L', 'l':
			for i := 0; i+1 < len(args); i += 2 {
				x, y := args[i], args[i+1]
				gdml.WriteString(fmt.Sprintf("L %d %d ", SVGToEMU(x), SVGToEMU(y)))
				updateBounds(&minX, &minY, &maxX, &maxY, x, y)
			}
		case 'H', 'h':
			for _, x := range args {
				gdml.WriteString(fmt.Sprintf("L %d 0 ", SVGToEMU(x)))
				updateBounds(&minX, &minY, &maxX, &maxY, x, 0)
			}
		case 'V', 'v':
			for _, y := range args {
				gdml.WriteString(fmt.Sprintf("L 0 %d ", SVGToEMU(y)))
				updateBounds(&minX, &minY, &maxX, &maxY, 0, y)
			}
		case 'C', 'c':
			for i := 0; i+5 < len(args); i += 6 {
				x1, y1 := args[i], args[i+1]
				x2, y2 := args[i+2], args[i+3]
				x, y := args[i+4], args[i+5]
				gdml.WriteString(fmt.Sprintf("C %d %d %d %d %d %d ",
					SVGToEMU(x1), SVGToEMU(y1),
					SVGToEMU(x2), SVGToEMU(y2),
					SVGToEMU(x), SVGToEMU(y)))
				updateBounds(&minX, &minY, &maxX, &maxY, x1, y1)
				updateBounds(&minX, &minY, &maxX, &maxY, x2, y2)
				updateBounds(&minX, &minY, &maxX, &maxY, x, y)
			}
		case 'S', 's':
			for i := 0; i+3 < len(args); i += 4 {
				x2, y2 := args[i], args[i+1]
				x, y := args[i+2], args[i+3]
				gdml.WriteString(fmt.Sprintf("C %d %d %d %d %d %d ",
					SVGToEMU(x2), SVGToEMU(y2),
					SVGToEMU(x2), SVGToEMU(y2),
					SVGToEMU(x), SVGToEMU(y)))
				updateBounds(&minX, &minY, &maxX, &maxY, x2, y2)
				updateBounds(&minX, &minY, &maxX, &maxY, x, y)
			}
		case 'Q', 'q':
			for i := 0; i+3 < len(args); i += 4 {
				x1, y1 := args[i], args[i+1]
				x, y := args[i+2], args[i+3]
				// Convert quadratic to cubic bezier
				gdml.WriteString(fmt.Sprintf("C %d %d %d %d %d %d ",
					SVGToEMU(x1), SVGToEMU(y1),
					SVGToEMU(x1), SVGToEMU(y1),
					SVGToEMU(x), SVGToEMU(y)))
				updateBounds(&minX, &minY, &maxX, &maxY, x1, y1)
				updateBounds(&minX, &minY, &maxX, &maxY, x, y)
			}
		case 'Z', 'z':
			gdml.WriteString("Z ")
		}
	}

	if minX == math.MaxFloat64 {
		minX, minY, maxX, maxY = 0, 0, 0, 0
	}

	return gdml.String(), bbox{
		X: minX,
		Y: minY,
		W: maxX - minX,
		H: maxY - minY,
	}
}

type pathToken struct {
	Cmd  rune
	Args []float64
}

func tokenizePath(d string) []pathToken {
	var tokens []pathToken
	var current *pathToken
	var numStr strings.Builder

	flush := func() {
		if numStr.Len() > 0 {
			if current != nil {
				val, _ := strconv.ParseFloat(numStr.String(), 64)
				current.Args = append(current.Args, val)
			}
			numStr.Reset()
		}
	}

	for i := 0; i < len(d); i++ {
		ch := d[i]
		switch {
		case ch >= 'A' && ch <= 'Z' || ch >= 'a' && ch <= 'z':
			flush()
			if current != nil && len(current.Args) > 0 {
				tokens = append(tokens, *current)
			}
			current = &pathToken{Cmd: rune(ch)}
		case ch == '-' || ch == '.' || (ch >= '0' && ch <= '9'):
			if ch == '-' && numStr.Len() > 0 {
				flush()
			}
			numStr.WriteByte(ch)
		case ch == ',' || ch == ' ' || ch == '\n' || ch == '\t' || ch == '\r':
			flush()
		}
	}

	flush()
	if current != nil && len(current.Args) > 0 {
		tokens = append(tokens, *current)
	}

	return tokens
}

func updateBounds(minX, minY, maxX, maxY *float64, x, y float64) {
	if x < *minX {
		*minX = x
	}
	if y < *minY {
		*minY = y
	}
	if x > *maxX {
		*maxX = x
	}
	if y > *maxY {
		*maxY = y
	}
}

// toShape converts pptxShape to the internal Shape type.
func (ps pptxShape) toShape() *Shape {
	var fillColor, lineColor *Color

	if ps.Fill != nil && *ps.Fill != "" {
		c := hexToColor(*ps.Fill)
		fillColor = &c
	}
	if ps.Stroke != nil && *ps.Stroke != "" {
		c := hexToColor(*ps.Stroke)
		lineColor = &c
	}

	shape := &Shape{
		id:        ps.ID,
		shapeType: ShapeRectangle,
		x:         ps.X,
		y:         ps.Y,
		w:         ps.W,
		h:         ps.H,
		fillColor: fillColor,
		lineColor: lineColor,
		lineWidth: int(ps.StrokeW),
		custGeom:  ps.CustGeom,
		prstGeom:  ps.PrstGeom,
	}

	// Add text if present
	if ps.Text != "" {
		fontSize := ps.FontSize / 100 // Convert from hundredths to points
		if fontSize == 0 {
			fontSize = 12
		}
		var textColor *Color
		if ps.TextColor != nil && *ps.TextColor != "" {
			c := hexToColor(*ps.TextColor)
			textColor = &c
		}
		shape.texts = []TextRun{{
			Text:     ps.Text,
			FontSize: fontSize,
			Bold:     ps.Bold,
			Color:    textColor,
		}}
	}

	return shape
}

func hexToColor(hex string) Color {
	hex = strings.TrimPrefix(hex, "#")
	if len(hex) == 3 {
		hex = string([]byte{hex[0], hex[0], hex[1], hex[1], hex[2], hex[2]})
	}
	if len(hex) != 6 {
		return Black
	}

	r, _ := strconv.ParseUint(hex[0:2], 16, 8)
	g, _ := strconv.ParseUint(hex[2:4], 16, 8)
	b, _ := strconv.ParseUint(hex[4:6], 16, 8)

	return Color{R: uint8(r), G: uint8(g), B: uint8(b)}
}

// toXML converts customGeom to XML representation.
func (cg *customGeom) toXML() *custGeomXML {
	xml := &custGeomXML{
		AvLst:   avLst{},
		PathLst: pathLstXML{},
	}

	for _, p := range cg.Paths {
		xml.PathLst.Path = append(xml.PathLst.Path, pathXML{
			W:    p.W,
			H:    p.H,
			Path: p.Path,
		})
	}

	return xml
}

// ToPPTXShapesScaled converts SVG elements to PPTX shapes with scaling.
func (r SVGRect) ToPPTXShapesScaled(id *int, scale float64, offsetX, offsetY int64, viewBox [4]float64) []pptxShape {
	*id++
	fill := parseColor(r.Fill)
	stroke := parseColor(r.Stroke)

	// Apply viewBox offset and scale
	x := int64((r.X-viewBox[0])*scale) + offsetX
	y := int64((r.Y-viewBox[1])*scale) + offsetY
	w := int64(r.Width * scale)
	h := int64(r.Height * scale)

	shape := pptxShape{
		ID:       *id,
		Name:     fmt.Sprintf("Rect %d", *id),
		X:        x,
		Y:        y,
		W:        w,
		H:        h,
		PrstGeom: "rect",
	}

	if fill != "" {
		shape.Fill = &fill
	} else if r.Fill == "" {
		// Default fill for SVG rects
		defaultFill := "000000"
		shape.Fill = &defaultFill
	}
	if stroke != "" {
		shape.Stroke = &stroke
		shape.StrokeW = int64(r.StrokeWidth * scale)
	}

	return []pptxShape{shape}
}

// ToPPTXShapesScaled converts SVGCircle to PPTX shapes with scaling.
func (c SVGCircle) ToPPTXShapesScaled(id *int, scale float64, offsetX, offsetY int64, viewBox [4]float64) []pptxShape {
	*id++
	fill := parseColor(c.Fill)
	stroke := parseColor(c.Stroke)

	x := int64((c.CX-c.R-viewBox[0])*scale) + offsetX
	y := int64((c.CY-c.R-viewBox[1])*scale) + offsetY
	w := int64(c.R * 2 * scale)
	h := int64(c.R * 2 * scale)

	shape := pptxShape{
		ID:       *id,
		Name:     fmt.Sprintf("Circle %d", *id),
		X:        x,
		Y:        y,
		W:        w,
		H:        h,
		PrstGeom: "ellipse",
	}

	if fill != "" {
		shape.Fill = &fill
	}
	if stroke != "" {
		shape.Stroke = &stroke
		shape.StrokeW = int64(c.StrokeWidth * scale)
	}

	return []pptxShape{shape}
}

// ToPPTXShapesScaled converts SVGPath to PPTX shapes with scaling.
func (p SVGPath) ToPPTXShapesScaled(id *int, scale float64, offsetX, offsetY int64, viewBox [4]float64) []pptxShape {
	*id++
	fill := parseColor(p.Fill)
	stroke := parseColor(p.Stroke)

	// Parse path to get bounding box and convert to GDML
	gdml, bounds := convertSVGPathToGDMLScaled(p.D, scale, viewBox)

	// Calculate position with offset
	x := int64(bounds.X) + offsetX
	y := int64(bounds.Y) + offsetY
	w := int64(bounds.W)
	h := int64(bounds.H)

	// Skip if no size
	if w <= 0 || h <= 0 {
		*id-- // Revert ID increment
		return nil
	}

	shape := pptxShape{
		ID:   *id,
		Name: fmt.Sprintf("Path %d", *id),
		X:    x,
		Y:    y,
		W:    w,
		H:    h,
		CustGeom: &customGeom{
			Paths: []geomPath{
				{
					W:    w,
					H:    h,
					Path: gdml,
				},
			},
		},
	}

	// Handle fill vs stroke
	if p.Fill == "none" && stroke != "" {
		// Stroke-only path - use stroke color as fill for visibility
		// (PPTX doesn't render thin strokes well on custom geometry)
		shape.Fill = &stroke
	} else if fill != "" {
		shape.Fill = &fill
		if stroke != "" {
			shape.Stroke = &stroke
			shape.StrokeW = int64(p.StrokeWidth * scale)
		}
	} else if p.Fill == "" || p.Fill == "black" {
		// Default fill for SVG paths
		defaultFill := "000000"
		shape.Fill = &defaultFill
	}

	return []pptxShape{shape}
}

// ToPPTXShapesScaled converts SVGGroup to PPTX shapes with scaling.
func (g SVGGroup) ToPPTXShapesScaled(id *int, scale float64, offsetX, offsetY int64, viewBox [4]float64) []pptxShape {
	var shapes []pptxShape
	for _, elem := range g.Elements {
		shapes = append(shapes, elem.ToPPTXShapesScaled(id, scale, offsetX, offsetY, viewBox)...)
	}
	return shapes
}

// convertSVGPathToGDMLScaled converts SVG path data to GDML with scaling.
// Returns multiple GDML paths if the SVG path contains multiple sub-paths.
func convertSVGPathToGDMLScaled(svgPath string, scale float64, viewBox [4]float64) (string, bbox) {
	tokens := tokenizePath(svgPath)

	// Split into sub-paths (each M command starts a new sub-path)
	type coord struct{ x, y float64 }
	var subPaths []struct {
		tokens []pathToken
	}

	var current *struct {
		tokens []pathToken
	}

	for _, tok := range tokens {
		if tok.Cmd == 'M' || tok.Cmd == 'm' {
			if current != nil && len(current.tokens) > 0 {
				subPaths = append(subPaths, *current)
			}
			current = &struct{ tokens []pathToken }{tokens: []pathToken{tok}}
		} else if current != nil {
			current.tokens = append(current.tokens, tok)
		}
	}
	if current != nil && len(current.tokens) > 0 {
		subPaths = append(subPaths, *current)
	}

	// If only one sub-path, process normally
	if len(subPaths) <= 1 {
		return convertSinglePathToGDML(tokens, scale, viewBox)
	}

	// Process all sub-paths and combine
	minX, minY := math.MaxFloat64, math.MaxFloat64
	maxX, maxY := -math.MaxFloat64, -math.MaxFloat64
	vbX, vbY := viewBox[0], viewBox[1]

	type scaledCoord struct{ x, y float64 }
	var allCoords []struct {
		cmd  rune
		vals []scaledCoord
	}

	for _, sp := range subPaths {
		for _, tok := range sp.tokens {
			cmd := tok.Cmd
			args := tok.Args

			switch cmd {
			case 'M', 'm':
				if len(args) >= 2 {
					x, y := args[0], args[1]
					sx := (x - vbX) * scale
					sy := (y - vbY) * scale
					allCoords = append(allCoords, struct {
						cmd  rune
						vals []scaledCoord
					}{cmd: 'M', vals: []scaledCoord{{sx, sy}}})
					updateBounds(&minX, &minY, &maxX, &maxY, sx, sy)
				}
			case 'L', 'l':
				var vals []scaledCoord
				for i := 0; i+1 < len(args); i += 2 {
					x, y := args[i], args[i+1]
					sx := (x - vbX) * scale
					sy := (y - vbY) * scale
					vals = append(vals, scaledCoord{sx, sy})
					updateBounds(&minX, &minY, &maxX, &maxY, sx, sy)
				}
				allCoords = append(allCoords, struct {
					cmd  rune
					vals []scaledCoord
				}{cmd: 'L', vals: vals})
			case 'C', 'c':
				var vals []scaledCoord
				for i := 0; i+5 < len(args); i += 6 {
					x1, y1 := args[i], args[i+1]
					x2, y2 := args[i+2], args[i+3]
					x, y := args[i+4], args[i+5]
					sx1 := (x1 - vbX) * scale
					sy1 := (y1 - vbY) * scale
					sx2 := (x2 - vbX) * scale
					sy2 := (y2 - vbY) * scale
					sx := (x - vbX) * scale
					sy := (y - vbY) * scale
					vals = append(vals, scaledCoord{sx1, sy1}, scaledCoord{sx2, sy2}, scaledCoord{sx, sy})
					updateBounds(&minX, &minY, &maxX, &maxY, sx1, sy1)
					updateBounds(&minX, &minY, &maxX, &maxY, sx2, sy2)
					updateBounds(&minX, &minY, &maxX, &maxY, sx, sy)
				}
				allCoords = append(allCoords, struct {
					cmd  rune
					vals []scaledCoord
				}{cmd: 'C', vals: vals})
			case 'Z', 'z':
				allCoords = append(allCoords, struct {
					cmd  rune
					vals []scaledCoord
				}{cmd: 'Z'})
			}
		}
	}

	if minX == math.MaxFloat64 {
		return "", bbox{0, 0, 0, 0}
	}

	// Normalize coordinates
	var gdml strings.Builder
	inPath := false

	// Use max coordinate for both w and h to ensure all points fit
	allMax := math.Max(maxX, maxY)
	padding := allMax * 0.15
	if padding < 10000 {
		padding = 10000
	}
	pathW := int64(allMax + padding)
	pathH := int64(allMax + padding)

	for _, pc := range allCoords {
		switch pc.cmd {
		case 'M':
			// Close previous sub-path
			if inPath {
				gdml.WriteString("Z ")
			}
			if len(pc.vals) > 0 {
				gdml.WriteString(fmt.Sprintf("M %d %d ", int64(pc.vals[0].x+padding), int64(pc.vals[0].y+padding)))
				inPath = true
			}
		case 'L':
			for _, v := range pc.vals {
				gdml.WriteString(fmt.Sprintf("L %d %d ", int64(v.x+padding), int64(v.y+padding)))
			}
		case 'C':
			for i := 0; i+2 < len(pc.vals); i += 3 {
				gdml.WriteString(fmt.Sprintf("C %d %d %d %d %d %d ",
					int64(pc.vals[i].x+padding), int64(pc.vals[i].y+padding),
					int64(pc.vals[i+1].x+padding), int64(pc.vals[i+1].y+padding),
					int64(pc.vals[i+2].x+padding), int64(pc.vals[i+2].y+padding)))
			}
		case 'Z':
			gdml.WriteString("Z ")
			inPath = false
		}
	}
	// Close final sub-path if not already closed
	if inPath {
		gdml.WriteString("Z ")
	}

	return gdml.String(), bbox{
		X: minX - padding,
		Y: minY - padding,
		W: float64(pathW),
		H: float64(pathH),
	}
}

// convertSinglePathToGDML converts a single path (no multiple M commands)
func convertSinglePathToGDML(tokens []pathToken, scale float64, viewBox [4]float64) (string, bbox) {
	type coord struct{ x, y float64 }

	minX, minY := math.MaxFloat64, math.MaxFloat64
	maxX, maxY := -math.MaxFloat64, -math.MaxFloat64
	vbX, vbY := viewBox[0], viewBox[1]

	type scaledCoord struct{ x, y float64 }
	var coords []struct {
		cmd  rune
		vals []scaledCoord
	}

	for _, tok := range tokens {
		cmd := tok.Cmd
		args := tok.Args

		switch cmd {
		case 'M', 'm':
			if len(args) >= 2 {
				x, y := args[0], args[1]
				sx := (x - vbX) * scale
				sy := (y - vbY) * scale
				coords = append(coords, struct {
					cmd  rune
					vals []scaledCoord
				}{cmd: 'M', vals: []scaledCoord{{sx, sy}}})
				updateBounds(&minX, &minY, &maxX, &maxY, sx, sy)
			}
		case 'L', 'l':
			var vals []scaledCoord
			for i := 0; i+1 < len(args); i += 2 {
				x, y := args[i], args[i+1]
				sx := (x - vbX) * scale
				sy := (y - vbY) * scale
				vals = append(vals, scaledCoord{sx, sy})
				updateBounds(&minX, &minY, &maxX, &maxY, sx, sy)
			}
			coords = append(coords, struct {
				cmd  rune
				vals []scaledCoord
			}{cmd: 'L', vals: vals})
		case 'C', 'c':
			var vals []scaledCoord
			for i := 0; i+5 < len(args); i += 6 {
				x1, y1 := args[i], args[i+1]
				x2, y2 := args[i+2], args[i+3]
				x, y := args[i+4], args[i+5]
				sx1 := (x1 - vbX) * scale
				sy1 := (y1 - vbY) * scale
				sx2 := (x2 - vbX) * scale
				sy2 := (y2 - vbY) * scale
				sx := (x - vbX) * scale
				sy := (y - vbY) * scale
				vals = append(vals, scaledCoord{sx1, sy1}, scaledCoord{sx2, sy2}, scaledCoord{sx, sy})
				updateBounds(&minX, &minY, &maxX, &maxY, sx1, sy1)
				updateBounds(&minX, &minY, &maxX, &maxY, sx2, sy2)
				updateBounds(&minX, &minY, &maxX, &maxY, sx, sy)
			}
			coords = append(coords, struct {
				cmd  rune
				vals []scaledCoord
			}{cmd: 'C', vals: vals})
		case 'Z', 'z':
			coords = append(coords, struct {
				cmd  rune
				vals []scaledCoord
			}{cmd: 'Z'})
		}
	}

	if minX == math.MaxFloat64 {
		return "", bbox{0, 0, 0, 0}
	}

	// Normalize
	var gdml strings.Builder
	inPath := false

	// Use max coordinate for both w and h to ensure all points fit
	allMax := math.Max(maxX, maxY)
	padding := allMax * 0.15
	if padding < 10000 {
		padding = 10000
	}

	for _, pc := range coords {
		switch pc.cmd {
		case 'M':
			// Close previous sub-path
			if inPath {
				gdml.WriteString("Z ")
			}
			if len(pc.vals) > 0 {
				gdml.WriteString(fmt.Sprintf("M %d %d ", int64(pc.vals[0].x+padding), int64(pc.vals[0].y+padding)))
				inPath = true
			}
		case 'L':
			for _, v := range pc.vals {
				gdml.WriteString(fmt.Sprintf("L %d %d ", int64(v.x+padding), int64(v.y+padding)))
			}
		case 'C':
			for i := 0; i+2 < len(pc.vals); i += 3 {
				gdml.WriteString(fmt.Sprintf("C %d %d %d %d %d %d ",
					int64(pc.vals[i].x+padding), int64(pc.vals[i].y+padding),
					int64(pc.vals[i+1].x+padding), int64(pc.vals[i+1].y+padding),
					int64(pc.vals[i+2].x+padding), int64(pc.vals[i+2].y+padding)))
			}
		case 'Z':
			gdml.WriteString("Z ")
			inPath = false
		}
	}
	// Close final sub-path if not already closed
	if inPath {
		gdml.WriteString("Z ")
	}

	return gdml.String(), bbox{
		X: minX - padding,
		Y: minY - padding,
		W: allMax + padding*2,
		H: allMax + padding*2,
	}
}
