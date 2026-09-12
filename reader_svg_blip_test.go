package gopresentation

import (
	"bytes"
	"encoding/xml"
	"image"
	"image/color"
	"image/png"
	"testing"
)

// Reading an image reference out of a <p:pic>.
//
// Microsoft's SVG extension puts the reference for an SVG part on
// <asvg:svgBlip r:embed="...">, a child of <a:blip>, rather than on <a:blip>
// itself. PowerPoint writes both the raster fallback and the SVG, but a
// generator that emits SVG artwork only leaves the parent blip bare. The reader
// used to look solely at <a:blip r:embed>, so such a picture came back with no
// data at all and was drawn as an empty frame — the whole picture, silently
// missing.

// svgBlipNS is the namespace of the Microsoft SVG extension that defines
// asvg:svgBlip.
const svgBlipNS = "http://schemas.microsoft.com/office/drawing/2016/SVG/main"

// svgBlipURI is the extension URI that carries svgBlip.
const svgBlipURI = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}"

// svgIcon is a minimal but non-trivial SVG document: a filled circle under a
// filled polygon, so anything that renders it has to honour both fill and
// geometry.
const svgIcon = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 16 16">` +
	`<circle cx="8" cy="8" r="7" fill="#336699"/>` +
	`<polygon points="8,3 13,12 3,12" fill="#ffffff"/>` +
	`</svg>`

// slideWithPicture returns a slide holding one <p:pic> whose blipFill body is
// inserted verbatim, so a test can choose exactly which reference form the
// picture carries.
func slideWithPicture(blipFillBody string) string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
		`<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"` +
		` xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"` +
		` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"` +
		` xmlns:asvg="` + svgBlipNS + `">` +
		`<p:cSld><p:spTree>` +
		`<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
		`<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>` +
		`<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>` +
		`<p:pic>` +
		`<p:nvPicPr><p:cNvPr id="2" name="icon.svg"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
		blipFillBody +
		`<p:spPr><a:xfrm><a:off x="1000000" y="1000000"/><a:ext cx="2000000" cy="2000000"/></a:xfrm>` +
		`<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>` +
		`</p:pic>` +
		`</p:spTree></p:cSld></p:sld>`
}

// svgBlipFillOnly is the form under test: no r:embed on <a:blip>, the one and
// only reference on the nested svgBlip.
func svgBlipFillOnly() string {
	return `<p:blipFill><a:blip><a:extLst>` +
		`<a:ext uri="` + svgBlipURI + `"><asvg:svgBlip r:embed="rId2"/></a:ext>` +
		`</a:extLst></a:blip><a:stretch><a:fillRect/></a:stretch></p:blipFill>`
}

// rasterBlipFill is the ordinary form, with the reference on <a:blip>. rId1 is
// the PNG part, so an ordinary picture is expected back as image/png.
func rasterBlipFill() string {
	return `<p:blipFill><a:blip r:embed="rId1"/>` +
		`<a:stretch><a:fillRect/></a:stretch></p:blipFill>`
}

// bothBlipFill is the form PowerPoint writes: a raster fallback on <a:blip>
// plus the SVG in the extension.
func bothBlipFill() string {
	return `<p:blipFill><a:blip r:embed="rId1"><a:extLst>` +
		`<a:ext uri="` + svgBlipURI + `"><asvg:svgBlip r:embed="rId2"/></a:ext>` +
		`</a:extLst></a:blip><a:stretch><a:fillRect/></a:stretch></p:blipFill>`
}

// picturePackage assembles a package holding one slide with the given blipFill
// body, plus the two media parts its relationships point at.
//
// Only presentation.xml, its rels and the slide are structurally required; the
// reader treats everything else as optional, which keeps the fixture small
// enough to read as a specification.
func picturePackage(t *testing.T, blipFillBody string) []byte {
	t.Helper()
	return buildZip(t, map[string][]byte{
		"ppt/presentation.xml": []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
			`<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"` +
			` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
			`<p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst></p:presentation>`),
		"ppt/_rels/presentation.xml.rels": []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
			`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
			`<Relationship Id="rId1"` +
			` Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"` +
			` Target="slides/slide1.xml"/></Relationships>`),
		"ppt/slides/slide1.xml": []byte(slideWithPicture(blipFillBody)),
		// Targets are relative to the slide, as PowerPoint writes them.
		"ppt/slides/_rels/slide1.xml.rels": []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
			`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
			`<Relationship Id="rId1"` +
			` Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"` +
			` Target="media/image1.png"/>` +
			`<Relationship Id="rId2"` +
			` Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"` +
			` Target="media/image2.svg"/></Relationships>`),
		"ppt/slides/media/image1.png": smallPNG(t),
		"ppt/slides/media/image2.svg": []byte(svgIcon),
	})
}

// readPicture returns the only picture of a one-slide package.
func readPicture(t *testing.T, data []byte) *DrawingShape {
	t.Helper()
	reader, err := NewReader(ReaderPowerPoint2007)
	if err != nil {
		t.Fatalf("new reader: %v", err)
	}
	pres, err := reader.ReadFromReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatalf("read package: %v", err)
	}
	if pres.GetSlideCount() != 1 {
		t.Fatalf("slide count = %d, want 1", pres.GetSlideCount())
	}
	slide, err := pres.GetSlide(0)
	if err != nil {
		t.Fatalf("get slide: %v", err)
	}
	for _, s := range slide.GetShapes() {
		if d, ok := s.(*DrawingShape); ok {
			return d
		}
	}
	t.Fatal("no picture found on the slide")
	return nil
}

// TestReaderFollowsSVGBlipReference is the regression test: the reference is on
// the extension element, and the picture must come back with its bytes.
func TestReaderFollowsSVGBlipReference(t *testing.T) {
	d := readPicture(t, picturePackage(t, svgBlipFillOnly()))

	if len(d.GetImageData()) == 0 {
		t.Fatal("picture whose only reference is asvg:svgBlip read back with no image data; " +
			"the extension carries the reference and the reader must follow it")
	}
	if got := string(d.GetImageData()); got != svgIcon {
		t.Errorf("image data is not the SVG part:\n got %q\nwant %q", got, svgIcon)
	}
	if got := d.GetMimeType(); got != "image/svg+xml" {
		t.Errorf("mime type = %q, want %q", got, "image/svg+xml")
	}
}

// TestReaderPrefersRasterOverSVGInSameBlip pins which reference wins when both
// are present, which is the form PowerPoint writes. Either would render, but
// the choice has to be deliberate rather than an accident of element order: the
// reference on <a:blip> is reached first and is kept.
func TestReaderPrefersRasterOverSVGInSameBlip(t *testing.T) {
	d := readPicture(t, picturePackage(t, bothBlipFill()))

	if len(d.GetImageData()) == 0 {
		t.Fatal("picture with both references read back with no image data")
	}
	if got := d.GetMimeType(); got != "image/png" {
		t.Errorf("mime type = %q, want %q: the reference on <a:blip> is the raster fallback "+
			"and is resolved before the extension is reached", got, "image/png")
	}
}

// TestReaderStillFollowsPlainBlipReference guards the ordinary path against the
// refactor that added the extension handling: a reference on <a:blip> alone
// must keep working.
func TestReaderStillFollowsPlainBlipReference(t *testing.T) {
	d := readPicture(t, picturePackage(t, rasterBlipFill()))

	if len(d.GetImageData()) == 0 {
		t.Fatal("picture with an r:embed on <a:blip> read back with no image data")
	}
	if got := d.GetMimeType(); got != "image/png" {
		t.Errorf("mime type = %q, want %q", got, "image/png")
	}
}

// TestReaderReportsUnresolvableBlipReference covers the failure side: a
// reference naming no relationship leaves the picture with no data, which the
// renderer draws as a labelled placeholder rather than as a blank rectangle.
func TestReaderReportsUnresolvableBlipReference(t *testing.T) {
	missingRel := `<p:blipFill><a:blip><a:extLst>` +
		`<a:ext uri="` + svgBlipURI + `"><asvg:svgBlip r:embed="rId99"/></a:ext>` +
		`</a:extLst></a:blip><a:stretch><a:fillRect/></a:stretch></p:blipFill>`
	if d := readPicture(t, picturePackage(t, missingRel)); len(d.GetImageData()) != 0 {
		t.Errorf("picture referencing an unknown relationship read back %d bytes of data",
			len(d.GetImageData()))
	}
}

// TestEmbedRelIDFindsTheReferenceOnEitherPrefix checks the attribute lookup both
// reference forms go through. It compares local names, because the relationships
// prefix is bound by the document rather than fixed by the schema.
func TestEmbedRelIDFindsTheReferenceOnEitherPrefix(t *testing.T) {
	const src = `<a:blip xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"` +
		` xmlns:rr="http://schemas.openxmlformats.org/officeDocument/2006/relationships"` +
		` cstate="print" rr:embed="rId7"/>`

	dec := xml.NewDecoder(bytes.NewReader([]byte(src)))
	for {
		tok, err := dec.Token()
		if err != nil {
			t.Fatalf("parse %q: %v", src, err)
		}
		start, ok := tok.(xml.StartElement)
		if !ok {
			continue
		}
		if got := embedRelID(start); got != "rId7" {
			t.Errorf("embedRelID = %q, want %q", got, "rId7")
		}
		if got := embedRelID(xml.StartElement{Name: xml.Name{Local: "blip"}}); got != "" {
			t.Errorf("embedRelID with no attributes = %q, want %q", got, "")
		}
		return
	}
}

// smallPNG returns a 1x1 opaque PNG, used as the raster fallback in fixtures.
func smallPNG(t *testing.T) []byte {
	t.Helper()
	img := image.NewRGBA(image.Rect(0, 0, 1, 1))
	img.Set(0, 0, color.RGBA{R: 0x33, G: 0x66, B: 0x99, A: 0xff})
	var buf bytes.Buffer
	if err := png.Encode(&buf, img); err != nil {
		t.Fatalf("encode png: %v", err)
	}
	return buf.Bytes()
}
