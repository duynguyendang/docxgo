package main

import (
	"encoding/json"
	"fmt"

	"github.com/duynguyendang/docxgo/v3/pptx"
)

func main() {
	builder := pptx.NewPresentationBuilder(
		pptx.WithTitle("docxgo PPTX Demo"),
		pptx.WithLayout(pptx.Layout16x9),
	)

	// Slide 1-2: Title
	addTitleSlides(builder)

	// Slide 3: Features
	addFeaturesSlide(builder)

	// Slide 4: Shapes
	addShapesSlide(builder)

	// Slide 5: Text formatting
	addTextSlide(builder)

	// Slide 6: SVG inline
	addSVGSlide(builder)

	// Slide 7: SVG from file
	addSVGFileSlide(builder)

	// Slide 8: Architecture
	addArchSlide(builder)

	p, _ := builder.Build()
	p.SaveAs("pptx_demo.pptx")
	fmt.Println("Created: pptx_demo.pptx (8 slides)")
}

func addTitleSlides(b *pptx.PresentationBuilder) {
	b.AddSlide().
		SetBackgroundColor(pptx.Color{R: 30, G: 39, B: 97}).
		AddText("docxgo PPTX").
		SetBold(true).SetFontSize(48).SetColor(pptx.White).
		SetPosition(pptx.Inches(1), pptx.Inches(2)).
		SetSize(pptx.Inches(8), pptx.Inches(1.5)).
		SetAlignment(pptx.AlignmentCenter).End()

	b.AddSlide().
		SetBackgroundColor(pptx.Color{R: 30, G: 39, B: 97}).
		AddText("Pure Go PowerPoint Library").
		SetFontSize(24).SetColor(pptx.Color{R: 200, G: 200, B: 200}).
		SetPosition(pptx.Inches(1), pptx.Inches(3)).
		SetSize(pptx.Inches(8), pptx.Inches(1)).
		SetAlignment(pptx.AlignmentCenter).End()
}

func addFeaturesSlide(b *pptx.PresentationBuilder) {
	s := b.AddSlide()
	s.AddText("Features").SetBold(true).SetFontSize(36).
		SetColor(pptx.Color{R: 30, G: 39, B: 97}).
		SetPosition(pptx.Inches(1), pptx.Inches(0.5)).
		SetSize(pptx.Inches(8), pptx.Inches(0.8)).End()

	for i, f := range []string{
		"Pure Go - No external dependencies",
		"Builder Pattern - Fluent API",
		"SVG to PPTX conversion",
		"Lines with arrows",
		"JSON architecture diagrams",
	} {
		s.AddText(f).SetFontSize(18).
			SetColor(pptx.Color{R: 60, G: 60, B: 60}).
			SetPosition(pptx.Inches(1), pptx.Inches(1.5+float64(i)*0.5)).
			SetSize(pptx.Inches(8), pptx.Inches(0.5)).End()
	}
}

func addShapesSlide(b *pptx.PresentationBuilder) {
	s := b.AddSlide()
	s.AddText("Shape Gallery").SetBold(true).SetFontSize(36).
		SetColor(pptx.Color{R: 30, G: 39, B: 97}).
		SetPosition(pptx.Inches(3), pptx.Inches(0.3)).
		SetSize(pptx.Inches(6), pptx.Inches(0.8)).
		SetAlignment(pptx.AlignmentCenter).End()

	for i, c := range []pptx.Color{
		{R: 249, G: 97, B: 103}, {R: 6, G: 90, B: 130},
		{R: 44, G: 95, B: 45}, {R: 184, G: 80, B: 66},
		{R: 132, G: 181, B: 159},
	} {
		s.AddShape(pptx.ShapeRoundedRectangle).
			SetPosition(pptx.Inches(0.5+float64(i)*1.8), pptx.Inches(1.5)).
			SetSize(pptx.Inches(1.5), pptx.Inches(1.5)).
			SetFillColor(c).End()
	}
	for i, c := range []pptx.Color{
		{R: 249, G: 231, B: 149}, {R: 28, G: 114, B: 147},
		{R: 151, G: 188, B: 98}, {R: 231, G: 232, B: 209},
		{R: 105, G: 162, B: 151},
	} {
		s.AddShape(pptx.ShapeEllipse).
			SetPosition(pptx.Inches(0.5+float64(i)*1.8), pptx.Inches(3.5)).
			SetSize(pptx.Inches(1.5), pptx.Inches(1.5)).
			SetFillColor(c).End()
	}
}

func addTextSlide(b *pptx.PresentationBuilder) {
	s := b.AddSlide()
	s.AddText("Text Formatting").SetBold(true).SetFontSize(40).
		SetColor(pptx.Color{R: 30, G: 39, B: 97}).
		SetPosition(pptx.Inches(1), pptx.Inches(0.5)).
		SetSize(pptx.Inches(8), pptx.Inches(1)).
		SetAlignment(pptx.AlignmentCenter).End()

	s.AddText("Bold").SetBold(true).SetFontSize(24).
		SetPosition(pptx.Inches(1), pptx.Inches(2)).
		SetSize(pptx.Inches(2), pptx.Inches(0.5)).End()
	s.AddText("Italic").SetItalic(true).SetFontSize(24).
		SetPosition(pptx.Inches(4), pptx.Inches(2)).
		SetSize(pptx.Inches(2), pptx.Inches(0.5)).End()
	s.AddText("Underline").SetUnderline(true).SetFontSize(24).
		SetPosition(pptx.Inches(7), pptx.Inches(2)).
		SetSize(pptx.Inches(2), pptx.Inches(0.5)).End()

	s.AddText("With Background").SetFontSize(20).SetColor(pptx.White).
		SetFillColor(pptx.Color{R: 30, G: 39, B: 97}).
		SetPosition(pptx.Inches(3), pptx.Inches(3.5)).
		SetSize(pptx.Inches(4), pptx.Inches(0.7)).
		SetAlignment(pptx.AlignmentCenter).End()
}

func addSVGSlide(b *pptx.PresentationBuilder) {
	s := b.AddSlide()
	s.SetBackgroundColor(pptx.White)
	s.AddText("SVG Inline").SetBold(true).SetFontSize(24).
		SetColor(pptx.Color{R: 30, G: 39, B: 97}).
		SetPosition(pptx.Inches(0.5), pptx.Inches(0.2)).
		SetSize(pptx.Inches(4), pptx.Inches(0.5)).End()

	s.AddSVG([]byte(`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 400 300">
		<rect x="20" y="50" width="100" height="80" fill="#e8f5e9" stroke="#4caf50"/>
		<text x="70" y="45" text-anchor="middle" font-size="11" font-weight="bold">CLIENT</text>
		<rect x="150" y="50" width="120" height="80" fill="#e3f2fd" stroke="#2196f3"/>
		<text x="210" y="45" text-anchor="middle" font-size="11" font-weight="bold">SERVER</text>
		<rect x="300" y="50" width="80" height="80" fill="#f3e5f5" stroke="#7b1fa2"/>
		<text x="340" y="45" text-anchor="middle" font-size="11" font-weight="bold">DB</text>
		<line x1="120" y1="90" x2="150" y2="90" stroke="#666" stroke-width="2"/>
		<polygon points="150,90 143,85 143,95" fill="#666"/>
		<line x1="270" y1="90" x2="300" y2="90" stroke="#666" stroke-width="2"/>
		<polygon points="300,90 293,85 293,95" fill="#666"/>
	</svg>`))
}

func addSVGFileSlide(b *pptx.PresentationBuilder) {
	s := b.AddSlide()
	s.SetBackgroundColor(pptx.White)
	s.AddText("SVG Complex").SetBold(true).SetFontSize(24).
		SetColor(pptx.Color{R: 30, G: 39, B: 97}).
		SetPosition(pptx.Inches(0.5), pptx.Inches(0.2)).
		SetSize(pptx.Inches(4), pptx.Inches(0.5)).End()

	s.AddSVG([]byte(`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 400 300">
		<rect x="0" y="0" width="400" height="300" fill="#f8f9fa"/>
		<rect x="20" y="100" width="120" height="150" fill="#e8f5e9" stroke="#4caf50" stroke-width="2"/>
		<text x="80" y="90" text-anchor="middle" font-size="11" font-weight="bold">CLIENT</text>
		<rect x="30" y="120" width="100" height="35" rx="5" fill="white" stroke="#81c784"/>
		<rect x="30" y="165" width="100" height="35" rx="5" fill="white" stroke="#81c784"/>
		<rect x="160" y="20" width="220" height="260" fill="#e3f2fd" stroke="#2196f3" stroke-width="2"/>
		<text x="270" y="45" text-anchor="middle" font-size="12" font-weight="bold">AZURE CLOUD</text>
		<rect x="180" y="140" width="90" height="50" fill="white" stroke="#1976d2"/>
		<text x="225" y="170" text-anchor="middle" font-size="9">Front Door</text>
		<rect x="190" y="60" width="170" height="200" fill="#bbdefb" stroke="#1565c0"/>
		<text x="275" y="80" text-anchor="middle" font-size="10" font-weight="bold">AI GATEWAY</text>
		<rect x="200" y="90" width="150" height="40" fill="#c8e6c9" stroke="#388e3c"/>
		<text x="275" y="115" text-anchor="middle" font-size="9">Security</text>
		<rect x="200" y="195" width="150" height="25" fill="#fff9c4" stroke="#f9a825"/>
		<text x="275" y="212" text-anchor="middle" font-size="9">Governance</text>
		<line x1="140" y1="175" x2="180" y2="165" stroke="#666" stroke-width="2"/>
		<polygon points="180,165 172,162 172,168" fill="#666"/>
	</svg>`))
}

func addArchSlide(b *pptx.PresentationBuilder) {
	s := b.AddSlide()
	s.SetBackgroundColor(pptx.White)

	s.AddText("Architecture from JSON").SetBold(true).SetFontSize(24).
		SetColor(pptx.Color{R: 30, G: 39, B: 97}).
		SetPosition(pptx.Inches(0.5), pptx.Inches(0.1)).
		SetSize(pptx.Inches(5), pptx.Inches(0.4)).End()

	data := []byte(`{
		"canvas_size": {"width": 1000, "height": 562},
		"containers": [
			{"id": "c1", "name": "CLIENT", "location": {"x": 10, "y": 100}, "size": {"width": 140, "height": 200}, "color": "Green-Light"},
			{"id": "c2", "name": "AZURE", "location": {"x": 200, "y": 20}, "size": {"width": 500, "height": 500}, "color": "Blue-Outline"}
		],
		"components": [
			{"id": "p1", "name": "Dev Teams", "location": {"x": 30, "y": 150}, "size": {"width": 100, "height": 50}},
			{"id": "p2", "name": "Front Door", "location": {"x": 220, "y": 200}, "size": {"width": 100, "height": 60}},
			{"id": "p3", "name": "Gateway", "location": {"x": 350, "y": 150}, "size": {"width": 120, "height": 50}}
		],
		"flows": [
			{"from": "p1", "to": "p2", "label": "API", "type": "request"},
			{"from": "p2", "to": "p3", "label": "Route", "type": "logic"}
		]
	}`)

	var arch struct {
		CanvasSize struct{ Width, Height int } `json:"canvas_size"`
		Containers []struct {
			ID, Name, Color string
			Location        struct{ X, Y float64 }
			Size            struct{ Width, Height int }
		} `json:"containers"`
		Components []struct {
			ID       string
			Name     string
			Location struct{ X, Y float64 }
			Size     struct{ Width, Height int }
		} `json:"components"`
		Flows []struct {
			From, To, Label, Type string
		} `json:"flows"`
	}
	json.Unmarshal(data, &arch)

	sx, sy := 12.0/float64(arch.CanvasSize.Width), 7.0/float64(arch.CanvasSize.Height)
	ox, oy := 0.5, 0.3

	colorMap := map[string]pptx.Color{
		"Green-Light":  {R: 235, G: 251, B: 238},
		"Blue-Outline": {R: 230, G: 245, B: 255},
		"Blue-Fill":    {R: 200, G: 225, B: 255},
	}

	for _, c := range arch.Containers {
		x, y := ox+c.Location.X*sx, oy+c.Location.Y*sy
		w, h := float64(c.Size.Width)*sx, float64(c.Size.Height)*sy
		s.AddShape(pptx.ShapeRoundedRectangle).
			SetPosition(pptx.Inches(x), pptx.Inches(y)).
			SetSize(pptx.Inches(w), pptx.Inches(h)).
			SetFillColor(colorMap[c.Color]).
			SetLine(pptx.Color{R: 0, G: 120, B: 212}, 1).End()
		s.AddText(c.Name).SetFontSize(9).SetBold(true).
			SetColor(pptx.Color{R: 30, G: 39, B: 97}).
			SetPosition(pptx.Inches(x+0.1), pptx.Inches(y+0.05)).
			SetSize(pptx.Inches(w-0.2), pptx.Inches(0.25)).End()
	}

	for _, c := range arch.Components {
		x, y := ox+c.Location.X*sx, oy+c.Location.Y*sy
		w, h := float64(c.Size.Width)*sx, float64(c.Size.Height)*sy
		s.AddShape(pptx.ShapeRectangle).
			SetPosition(pptx.Inches(x), pptx.Inches(y)).
			SetSize(pptx.Inches(w), pptx.Inches(h)).
			SetFillColor(pptx.White).
			SetLine(pptx.Color{R: 100, G: 100, B: 100}, 1).End()
		s.AddText(c.Name).SetFontSize(8).
			SetColor(pptx.Color{R: 30, G: 30, B: 30}).
			SetPosition(pptx.Inches(x+0.05), pptx.Inches(y+0.05)).
			SetSize(pptx.Inches(w-0.1), pptx.Inches(h-0.1)).
			SetAlignment(pptx.AlignmentCenter).End()
	}

	flowColors := map[string]pptx.Color{
		"request": {R: 0, G: 120, B: 212},
		"logic":   {R: 0, G: 150, B: 0},
	}

	for _, f := range arch.Flows {
		var fx, fy, tx, ty float64
		for _, c := range arch.Components {
			if c.ID == f.From {
				fx = ox + (c.Location.X+float64(c.Size.Width))*sx
				fy = oy + c.Location.Y*sy + float64(c.Size.Height)*sy/2
			}
			if c.ID == f.To {
				tx = ox + c.Location.X*sx
				ty = oy + c.Location.Y*sy + float64(c.Size.Height)*sy/2
			}
		}
		clr := flowColors[f.Type]
		s.AddLine(pptx.Inches(fx), pptx.Inches(fy), pptx.Inches(tx), pptx.Inches(ty)).
			SetColor(clr).SetWidthPt(1.5).SetArrowEnd().End()
	}
}
