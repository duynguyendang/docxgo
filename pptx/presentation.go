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
	"archive/zip"
	"encoding/xml"
	"fmt"
	"os"
	"time"
)

// Presentation represents a PowerPoint presentation.
type Presentation struct {
	slides   []*Slide
	metadata *Metadata
	layout   Layout
	slideW   int64
	slideH   int64
	slideID  int
	shapeID  int
}

// Metadata contains presentation properties.
type Metadata struct {
	Title    string
	Author   string
	Subject  string
	Creator  string
	Keywords []string
	Created  time.Time
}

// NewPresentation creates a new empty presentation.
func NewPresentation() *Presentation {
	w, h := Layout16x9.Dimensions()
	p := &Presentation{
		metadata: &Metadata{
			Creator: "docxgo",
			Created: time.Now(),
		},
		layout:  Layout16x9,
		slideW:  w,
		slideH:  h,
		slideID: 256,
		shapeID: 1,
	}
	return p
}

// SetLayout sets the presentation layout (aspect ratio).
func (p *Presentation) SetLayout(layout Layout) *Presentation {
	p.layout = layout
	p.slideW, p.slideH = layout.Dimensions()
	return p
}

// SetTitle sets the presentation title.
func (p *Presentation) SetTitle(title string) *Presentation {
	p.metadata.Title = title
	return p
}

// SetAuthor sets the presentation author.
func (p *Presentation) SetAuthor(author string) *Presentation {
	p.metadata.Author = author
	p.metadata.Creator = author
	return p
}

// Metadata returns the presentation metadata.
func (p *Presentation) Metadata() *Metadata {
	return p.metadata
}

// AddSlide adds a new slide to the presentation.
func (p *Presentation) AddSlide() *Slide {
	p.slideID++
	slide := &Slide{
		id:      p.slideID,
		pres:    p,
		shapes:  make([]*Shape, 0),
		bgColor: nil,
	}
	p.slides = append(p.slides, slide)
	return slide
}

// Slides returns all slides in the presentation.
func (p *Presentation) Slides() []*Slide {
	return p.slides
}

// SlideCount returns the number of slides.
func (p *Presentation) SlideCount() int {
	return len(p.slides)
}

// SaveAs saves the presentation to a .pptx file.
func (p *Presentation) SaveAs(path string) error {
	f, err := os.Create(path)
	if err != nil {
		return fmt.Errorf("pptx: failed to create file: %w", err)
	}
	defer f.Close()

	return p.WriteTo(f)
}

// WriteTo writes the presentation to an io.Writer.
func (p *Presentation) WriteTo(w *os.File) error {
	zw := zip.NewWriter(w)
	defer zw.Close()

	// Write [Content_Types].xml
	if err := p.writeContentTypes(zw); err != nil {
		return err
	}

	// Write _rels/.rels
	if err := p.writeRootRels(zw); err != nil {
		return err
	}

	// Write docProps/app.xml
	if err := p.writeAppProps(zw); err != nil {
		return err
	}

	// Write docProps/core.xml
	if err := p.writeCoreProps(zw); err != nil {
		return err
	}

	// Write ppt/presentation.xml
	if err := p.writePresentation(zw); err != nil {
		return err
	}

	// Write ppt/_rels/presentation.xml.rels
	if err := p.writePresentationRels(zw); err != nil {
		return err
	}

	// Write ppt/theme/theme1.xml
	if err := p.writeTheme(zw); err != nil {
		return err
	}

	// Write ppt/slideMasters/slideMaster1.xml
	if err := p.writeSlideMaster(zw); err != nil {
		return err
	}

	// Write ppt/slideMasters/_rels/slideMaster1.xml.rels
	if err := p.writeSlideMasterRels(zw); err != nil {
		return err
	}

	// Write ppt/slideLayouts/slideLayout1.xml
	if err := p.writeSlideLayout(zw); err != nil {
		return err
	}

	// Write ppt/slideLayouts/_rels/slideLayout1.xml.rels
	if err := p.writeSlideLayoutRels(zw); err != nil {
		return err
	}

	// Write slides
	for i, slide := range p.slides {
		if err := p.writeSlide(zw, slide, i+1); err != nil {
			return err
		}
	}

	return nil
}

func (p *Presentation) nextShapeID() int {
	p.shapeID++
	return p.shapeID
}

// XML structures for PPTX

type contentTypes struct {
	XMLName   xml.Name      `xml:"Types"`
	Xmlns     string        `xml:"xmlns,attr"`
	Defaults  []contentType `xml:"Default"`
	Overrides []contentType `xml:"Override"`
}

type contentType struct {
	Extension   string `xml:"Extension,attr,omitempty"`
	ContentType string `xml:"ContentType,attr,omitempty"`
	PartName    string `xml:"PartName,attr,omitempty"`
}

type rels struct {
	XMLName      xml.Name       `xml:"Relationships"`
	Xmlns        string         `xml:"xmlns,attr"`
	Relationship []relationship `xml:"Relationship"`
}

type relationship struct {
	ID     string `xml:"Id,attr"`
	Type   string `xml:"Type,attr"`
	Target string `xml:"Target,attr"`
}

type coreProperties struct {
	XMLName      xml.Name `xml:"cp:coreProperties"`
	XmlnsCp      string   `xml:"xmlns:cp,attr"`
	XmlnsDc      string   `xml:"xmlns:dc,attr"`
	XmlnsDcterms string   `xml:"xmlns:dcterms,attr"`
	XmlnsXsi     string   `xml:"xmlns:xsi,attr"`
	Creator      string   `xml:"dc:creator"`
	Created      string   `xml:"dcterms:created"`
}

type appProperties struct {
	XMLName     xml.Name `xml:"Properties"`
	Xmlns       string   `xml:"xmlns,attr"`
	XmlnsVt     string   `xml:"xmlns:vt,attr"`
	Application string   `xml:"Application"`
	Slides      int      `xml:"Slides"`
}

type presentationXML struct {
	XMLName  xml.Name `xml:"p:presentation"`
	XmlnsR   string   `xml:"xmlns:r,attr"`
	XmlnsP   string   `xml:"xmlns:p,attr"`
	XmlnsA   string   `xml:"xmlns:a,attr"`
	XmlnsMC  string   `xml:"xmlns:mc,attr"`
	XmlnsNS1 string   `xml:"xmlns:ns1,attr,omitempty"`
	SldSz    sldSz    `xml:"p:sldSz"`
	SldIdLst sldIdLst `xml:"p:sldIdLst"`
}

type sldSz struct {
	Cx int64 `xml:"cx,attr"`
	Cy int64 `xml:"cy,attr"`
}

type sldIdLst struct {
	SldId []sldId `xml:"p:sldId"`
}

type sldId struct {
	ID  int    `xml:"id,attr"`
	RID string `xml:"r:id,attr"`
}

type themeXML struct {
	XMLName xml.Name `xml:"a:theme"`
	XmlnsA  string   `xml:"xmlns:a,attr"`
	Name    string   `xml:"name,attr"`
}

type slideMasterXML struct {
	XMLName xml.Name `xml:"p:sldMaster"`
	XmlnsR  string   `xml:"xmlns:r,attr"`
	XmlnsP  string   `xml:"xmlns:p,attr"`
	XmlnsA  string   `xml:"xmlns:a,attr"`
}

type slideLayoutXML struct {
	XMLName xml.Name `xml:"p:sldLayout"`
	XmlnsR  string   `xml:"xmlns:r,attr"`
	XmlnsP  string   `xml:"xmlns:p,attr"`
	XmlnsA  string   `xml:"xmlns:a,attr"`
}

func (p *Presentation) writeContentTypes(zw *zip.Writer) error {
	ct := contentTypes{
		Xmlns: "http://schemas.openxmlformats.org/package/2006/content-types",
		Defaults: []contentType{
			{Extension: "rels", ContentType: "application/vnd.openxmlformats-package.relationships+xml"},
			{Extension: "xml", ContentType: "application/xml"},
			{Extension: "png", ContentType: "image/png"},
			{Extension: "jpeg", ContentType: "image/jpeg"},
			{Extension: "jpg", ContentType: "image/jpeg"},
			{Extension: "gif", ContentType: "image/gif"},
			{Extension: "emf", ContentType: "image/x-emf"},
		},
		Overrides: []contentType{
			{PartName: "/ppt/presentation.xml", ContentType: "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"},
			{PartName: "/ppt/theme/theme1.xml", ContentType: "application/vnd.openxmlformats-officedocument.theme+xml"},
			{PartName: "/ppt/slideMasters/slideMaster1.xml", ContentType: "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"},
			{PartName: "/ppt/slideLayouts/slideLayout1.xml", ContentType: "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"},
			{PartName: "/docProps/core.xml", ContentType: "application/vnd.openxmlformats-package.core-properties+xml"},
			{PartName: "/docProps/app.xml", ContentType: "application/vnd.openxmlformats-officedocument.extended-properties+xml"},
		},
	}

	for i := range p.slides {
		ct.Overrides = append(ct.Overrides, contentType{
			PartName:    fmt.Sprintf("/ppt/slides/slide%d.xml", i+1),
			ContentType: "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
		})
	}

	return writeXMLToZip(zw, "[Content_Types].xml", ct)
}

func (p *Presentation) writeRootRels(zw *zip.Writer) error {
	r := rels{
		Xmlns: "http://schemas.openxmlformats.org/package/2006/relationships",
		Relationship: []relationship{
			{ID: "rId1", Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", Target: "ppt/presentation.xml"},
			{ID: "rId2", Type: "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties", Target: "docProps/core.xml"},
			{ID: "rId3", Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties", Target: "docProps/app.xml"},
		},
	}
	return writeXMLToZip(zw, "_rels/.rels", r)
}

func (p *Presentation) writeAppProps(zw *zip.Writer) error {
	ap := appProperties{
		Xmlns:       "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
		XmlnsVt:     "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
		Application: "docxgo",
		Slides:      len(p.slides),
	}
	return writeXMLToZip(zw, "docProps/app.xml", ap)
}

func (p *Presentation) writeCoreProps(zw *zip.Writer) error {
	cp := coreProperties{
		XmlnsCp:      "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
		XmlnsDc:      "http://purl.org/dc/elements/1.1/",
		XmlnsDcterms: "http://purl.org/dc/terms/",
		XmlnsXsi:     "http://www.w3.org/2001/XMLSchema-instance",
		Creator:      p.metadata.Creator,
		Created:      p.metadata.Created.Format(time.RFC3339),
	}
	return writeXMLToZip(zw, "docProps/core.xml", cp)
}

func (p *Presentation) writePresentation(zw *zip.Writer) error {
	pres := presentationXML{
		XmlnsR:  "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
		XmlnsP:  "http://schemas.openxmlformats.org/presentationml/2006/main",
		XmlnsA:  "http://schemas.openxmlformats.org/drawingml/2006/main",
		XmlnsMC: "http://schemas.openxmlformats.org/markup-compatibility/2006",
		SldSz:   sldSz{Cx: p.slideW, Cy: p.slideH},
		SldIdLst: sldIdLst{
			SldId: make([]sldId, len(p.slides)),
		},
	}

	for i := range p.slides {
		pres.SldIdLst.SldId[i] = sldId{
			ID:  p.slides[i].id,
			RID: fmt.Sprintf("rId%d", i+3), // rId1 is theme, rId2 is slideMaster
		}
	}

	return writeXMLToZip(zw, "ppt/presentation.xml", pres)
}

func (p *Presentation) writePresentationRels(zw *zip.Writer) error {
	r := rels{
		Xmlns: "http://schemas.openxmlformats.org/package/2006/relationships",
		Relationship: []relationship{
			{ID: "rId1", Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme", Target: "theme/theme1.xml"},
			{ID: "rId2", Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster", Target: "slideMasters/slideMaster1.xml"},
		},
	}

	for i := range p.slides {
		r.Relationship = append(r.Relationship, relationship{
			ID:     fmt.Sprintf("rId%d", i+3),
			Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
			Target: fmt.Sprintf("slides/slide%d.xml", i+1),
		})
	}

	return writeXMLToZip(zw, "ppt/_rels/presentation.xml.rels", r)
}

func (p *Presentation) writeTheme(zw *zip.Writer) error {
	// Minimal theme with default Office theme
	theme := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText"/></a:dk1>
      <a:lt1><a:sysClr val="window"/></a:lt1>
      <a:dk2><a:srgbClr val="1F497D"/></a:dk2>
      <a:lt2><a:srgbClr val="EEECE1"/></a:lt2>
      <a:accent1><a:srgbClr val="4F81BD"/></a:accent1>
      <a:accent2><a:srgbClr val="F79646"/></a:accent2>
      <a:accent3><a:srgbClr val="9BBB59"/></a:accent3>
      <a:accent4><a:srgbClr val="8064A2"/></a:accent4>
      <a:accent5><a:srgbClr val="4BACC6"/></a:accent5>
      <a:accent6><a:srgbClr val="F36523"/></a:accent6>
      <a:hlink><a:srgbClr val="0000FF"/></a:hlink>
      <a:folHlink><a:srgbClr val="800080"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
      <a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office">
      <a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="2700000"/></a:gradFill></a:fillStyleLst>
      <a:lnStyleLst><a:ln w="9525" cap="flat"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln></a:lnStyleLst>
      <a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>
      <a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
</a:theme>`

	f, err := zw.Create("ppt/theme/theme1.xml")
	if err != nil {
		return fmt.Errorf("pptx: failed to create theme: %w", err)
	}
	_, err = f.Write([]byte(theme))
	return err
}

func (p *Presentation) writeSlideMaster(zw *zip.Writer) error {
	master := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr/>
    </p:spTree>
  </p:cSld>
  <p:sldLayoutIdLst>
    <p:sldLayoutId id="2147483649" r:id="rId1"/>
  </p:sldLayoutIdLst>
  <p:txStyles>
    <p:titleStyle/>
    <p:bodyStyle/>
    <p:otherStyle/>
  </p:txStyles>
</p:sldMaster>`

	f, err := zw.Create("ppt/slideMasters/slideMaster1.xml")
	if err != nil {
		return fmt.Errorf("pptx: failed to create slide master: %w", err)
	}
	_, err = f.Write([]byte(master))
	return err
}

func (p *Presentation) writeSlideMasterRels(zw *zip.Writer) error {
	r := rels{
		Xmlns: "http://schemas.openxmlformats.org/package/2006/relationships",
		Relationship: []relationship{
			{ID: "rId1", Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout", Target: "../slideLayouts/slideLayout1.xml"},
			{ID: "rId2", Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme", Target: "../theme/theme1.xml"},
		},
	}
	return writeXMLToZip(zw, "ppt/slideMasters/_rels/slideMaster1.xml.rels", r)
}

func (p *Presentation) writeSlideLayout(zw *zip.Writer) error {
	layout := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             type="blank" preserve="1">
  <p:cSld name="Blank">
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr/>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr>
    <a:masterClrMapping/>
  </p:clrMapOvr>
</p:sldLayout>`

	f, err := zw.Create("ppt/slideLayouts/slideLayout1.xml")
	if err != nil {
		return fmt.Errorf("pptx: failed to create slide layout: %w", err)
	}
	_, err = f.Write([]byte(layout))
	return err
}

func (p *Presentation) writeSlideLayoutRels(zw *zip.Writer) error {
	r := rels{
		Xmlns: "http://schemas.openxmlformats.org/package/2006/relationships",
		Relationship: []relationship{
			{ID: "rId1", Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster", Target: "../slideMasters/slideMaster1.xml"},
		},
	}
	return writeXMLToZip(zw, "ppt/slideLayouts/_rels/slideLayout1.xml.rels", r)
}

func (p *Presentation) writeSlide(zw *zip.Writer, slide *Slide, num int) error {
	slideXML := slide.toXML()
	return writeXMLToZip(zw, fmt.Sprintf("ppt/slides/slide%d.xml", num), slideXML)
}

func writeXMLToZip(zw *zip.Writer, name string, v any) error {
	f, err := zw.Create(name)
	if err != nil {
		return fmt.Errorf("pptx: failed to create %s: %w", name, err)
	}

	data, err := xml.MarshalIndent(v, "", "  ")
	if err != nil {
		return fmt.Errorf("pptx: failed to marshal %s: %w", name, err)
	}

	// Write XML header
	if _, err := f.Write([]byte(xml.Header)); err != nil {
		return err
	}
	_, err = f.Write(data)
	return err
}
