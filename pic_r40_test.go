package gopresentation

// The r40 picture semantics tests.
//
// The comparison deck's slide 20 shows a framed code screenshot: a p:pic
// whose blip carries a duotone recolour (black → D9C3A5 tint50 satMod180),
// whose spPr carries a 7pt white mat (a:ln), an outer shadow and a
// snip2DiagRect frame. The pre-r40 reader kept only the pixels — no duotone,
// no border, no shadow, no frame clipping — so the card rendered as a raw
// white-background screenshot where PowerPoint shows a cream, matte-framed
// exhibit. Everything was pinned against PowerPoint COM exports
// (out_deck/r40 build_r40_e1.py, e1/e2 goldens):
//
//   - duotone maps every pixel onto the straight line between its two
//     colours at the pixel's Rec.709 luma, in sRGB gamma space:
//     out = A + (B-A)·t with t = (0.2126R + 0.7152G + 0.0722B)/255. The
//     sixteen gray bands of the calibration ramp land band-exact on
//     round(A+(B-A)·v/255) and the pure primaries map to 54/182/18 —
//     the three weights times 255.
//   - the duotone colours carry their own transform chain, folded at parse
//     time (D9C3A5 tint50 satMod180 → (245,228,208); our chain rounds to
//     (245,229,208), the ±1 the shared tint→satMod pipeline costs).

import (
	"bytes"
	"image"
	"image/color"
	"image/png"
	"strings"
	"testing"
)

// TestDuotoneMapsAtRec709Luma pins the pixel mapping itself: the two-colour
// line, the Rec.709 weights, gamma-space arithmetic, alpha pass-through.
func TestDuotoneMapsAtRec709Luma(t *testing.T) {
	img := image.NewRGBA(image.Rect(0, 0, 6, 1))
	cols := []color.RGBA{
		{R: 255, G: 255, B: 255, A: 255}, // white
		{R: 0, G: 0, B: 0, A: 255},       // black
		{R: 255, G: 0, B: 0, A: 255},     // red
		{R: 0, G: 255, B: 0, A: 255},     // lime
		{R: 0, G: 0, B: 255, A: 255},     // blue
		{R: 255, G: 255, B: 255, A: 128}, // translucent white
	}
	for i, c := range cols {
		img.SetRGBA(i, 0, c)
	}
	applyDuotone(img, NewColor("000000"), NewColor("FF0000"))

	want := [][3]uint8{
		{255, 0, 0}, // white: luma 1 → the light colour itself
		{0, 0, 0},   // black: luma 0 → the dark colour itself
		{54, 0, 0},  // red: 0.2126·255
		{182, 0, 0}, // lime: 0.7152·255
		{18, 0, 0},  // blue: 0.0722·255
		{128, 0, 0}, // translucent white: colour maps like white, alpha kept
	}
	for i, w := range want {
		got := img.RGBAAt(i, 0)
		if got.R != w[0] || got.G != w[1] || got.B != w[2] {
			t.Errorf("duotone[%d]: want rgb%v got rgb(%d,%d,%d)", i, w, got.R, got.G, got.B)
		}
	}
	if a := img.RGBAAt(5, 0).A; a != 128 {
		t.Errorf("duotone must not touch alpha: got %d, want 128", a)
	}
}

// duotoneCardSlide mirrors the comparison deck's slide 20 card: duotone with
// a transform chain on the light colour, a crop, the 7pt white mat, a shadow
// and the snipped-corner frame.
const duotoneCardSlide = `
<p:pic>
  <p:nvPicPr><p:cNvPr id="20" name="Card"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
  <p:blipFill>
    <a:blip r:embed="rId101"><a:duotone><a:prstClr val="black"/><a:srgbClr val="D9C3A5"><a:tint val="50000"/><a:satMod val="180000"/></a:srgbClr></a:duotone></a:blip>
    <a:srcRect l="11667" t="30745" r="52475" b="19431"/>
    <a:stretch><a:fillRect/></a:stretch>
  </p:blipFill>
  <p:spPr>
    <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="2880320"/></a:xfrm>
    <a:prstGeom prst="snip2DiagRect"><a:avLst/></a:prstGeom>
    <a:ln w="88900" cap="sq"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln>
    <a:effectLst><a:outerShdw blurRad="88900" algn="tl" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="45000"/></a:srgbClr></a:outerShdw></a:effectLst>
  </p:spPr>
</p:pic>`

// readPicFixture reads a slide whose single picture references a tiny red
// PNG through rId101.
func readPicFixture(t *testing.T, picXML string) *Presentation {
	t.Helper()
	parts := zipParts(t, writeToBytes(t, New()))
	var buf bytes.Buffer
	src := image.NewRGBA(image.Rect(0, 0, 4, 4))
	for y := 0; y < 4; y++ {
		for x := 0; x < 4; x++ {
			src.SetRGBA(x, y, color.RGBA{R: 255, A: 255})
		}
	}
	if err := png.Encode(&buf, src); err != nil {
		t.Fatalf("encode fixture image: %v", err)
	}
	parts["ppt/media/image1.png"] = buf.Bytes()
	rels := string(parts["ppt/slides/_rels/slide1.xml.rels"])
	rels = strings.Replace(rels, "</Relationships>",
		`<Relationship Id="rId101" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/></Relationships>`, 1)
	parts["ppt/slides/_rels/slide1.xml.rels"] = []byte(rels)
	parts["ppt/slides/slide1.xml"] = []byte(bodyPrSlide(picXML))
	pkg := buildZip(t, parts)
	pres, err := (&PPTXReader{}).ReadFromReader(bytes.NewReader(pkg), int64(len(pkg)))
	if err != nil {
		t.Fatalf("read the package: %v", err)
	}
	return pres
}

func findDrawing(shapes []Shape) *DrawingShape {
	for _, sh := range shapes {
		if d, ok := sh.(*DrawingShape); ok {
			return d
		}
	}
	return nil
}

// TestReaderParsesPicFrameAndDuotone: the whole frame survives the read —
// duotone endpoints with the transform chain folded in, geometry name, matte
// border and shadow.
func TestReaderParsesPicFrameAndDuotone(t *testing.T) {
	pres := readPicFixture(t, duotoneCardSlide)
	d := findDrawing(pres.GetAllSlides()[0].GetShapes())
	if d == nil {
		t.Fatal("picture dropped")
	}
	if !d.hasDuotone {
		t.Fatal("duotone lost")
	}
	if d.duotoneA.ARGB != "FF000000" {
		t.Errorf("duotoneA = %s, want FF000000 (prstClr black)", d.duotoneA.ARGB)
	}
	// PPT reads (245,228,208); our tint→satMod chain rounds to (245,229,208).
	if d.duotoneB.ARGB != "FFF5E5D0" {
		t.Errorf("duotoneB = %s, want FFF5E5D0 (D9C3A5 tint50 satMod180, ±1 of PPT F5E4D0)", d.duotoneB.ARGB)
	}
	if d.presetGeom != "snip2DiagRect" {
		t.Errorf("presetGeom = %q, want snip2DiagRect", d.presetGeom)
	}
	if d.border == nil {
		t.Fatal("frame border lost")
	}
	if d.border.Width != 7 {
		t.Errorf("border width = %d, want 7 (88900 EMU)", d.border.Width)
	}
	if d.border.Color.ARGB != "FFFFFFFF" {
		t.Errorf("border colour = %s, want FFFFFFFF", d.border.Color.ARGB)
	}
	if d.shadow == nil || !d.shadow.Visible {
		t.Fatal("frame shadow lost")
	}
	if d.shadow.BlurRadius != 7 {
		t.Errorf("shadow blur = %d, want 7", d.shadow.BlurRadius)
	}
	if d.shadow.Alpha != 45 {
		t.Errorf("shadow alpha = %d, want 45", d.shadow.Alpha)
	}
}

// TestWriterRoundTripsPicFrame: reading then writing keeps the duotone, the
// frame geometry and the matte. The duotone comes back as resolved colours —
// the transforms were folded at parse time — which renders identically.
func TestWriterRoundTripsPicFrame(t *testing.T) {
	pres := readPicFixture(t, duotoneCardSlide)
	parts := zipParts(t, writeToBytes(t, pres))
	xml := string(parts["ppt/slides/slide1.xml"])
	for _, want := range []string{
		`<a:duotone><a:srgbClr val="000000"/><a:srgbClr val="F5E5D0"/></a:duotone>`,
		`<a:prstGeom prst="snip2DiagRect">`,
		`<a:ln w="88900"`,
		`<a:outerShdw blurRad="88900"`,
	} {
		if !strings.Contains(xml, want) {
			t.Errorf("written slide XML lacks %q", want)
		}
	}
}

// TestPicFrameRendersClippedWithMatteAndShadow: the render sentinel. The
// snipped corners must show the background (the image used to be pasted as a
// full rectangle), the interior must show the image, the matte must show on
// the frame edge and the shadow must darken below the box.
func TestPicFrameRendersClippedWithMatteAndShadow(t *testing.T) {
	fc := NewFontCache()
	opts := goldenOptions(fc)
	var buf bytes.Buffer
	src := image.NewRGBA(image.Rect(0, 0, 4, 4))
	for y := 0; y < 4; y++ {
		for x := 0; x < 4; x++ {
			src.SetRGBA(x, y, color.RGBA{R: 255, A: 255})
		}
	}
	if err := png.Encode(&buf, src); err != nil {
		t.Fatalf("encode fixture: %v", err)
	}

	p := New()
	d := NewDrawingShape()
	d.SetImageData(buf.Bytes(), "image/png")
	d.BaseShape.SetOffsetX(914400).SetOffsetY(914400).SetWidth(2540000).SetHeight(2540000) // 200pt box at 72pt
	d.presetGeom = "snip2DiagRect"
	d.border = &Border{Style: BorderSolid, Width: 7}
	d.border.Color = NewColor("000000")
	d.shadow = NewShadow()
	d.shadow.Visible = true
	d.shadow.Color = NewColor("000000")
	d.shadow.Alpha = 100
	d.shadow.BlurRadius = 2
	p.GetActiveSlide().AddShape(d)

	img, err := p.SlideToImage(0, opts)
	if err != nil {
		t.Fatalf("render: %v", err)
	}
	rgba, ok := img.(*image.RGBA)
	if !ok {
		t.Fatalf("render is %T, want *image.RGBA", img)
	}
	box := emuRect(t, p, 914400, 914400, 2540000, 2540000, opts.Width)
	probe := func(x, y int) color.RGBA { return rgba.RGBAAt(x, y) }

	// Top-right snipped corner: outside the frame polygon → background.
	c := probe(box.Max.X-2, box.Min.Y+2)
	if c.R != 255 || c.G != 255 || c.B != 255 {
		t.Errorf("snipped corner pixel (%d,%d) = rgb(%d,%d,%d), want untouched background",
			box.Max.X-2, box.Min.Y+2, c.R, c.G, c.B)
	}
	// Interior well inside the frame → the image (red) under the matte.
	c = probe(box.Min.X+box.Dx()/2, box.Min.Y+box.Dy()/2)
	if c.R < 200 || c.G > 100 {
		t.Errorf("centre pixel = rgb(%d,%d,%d), want the red image dominant", c.R, c.G, c.B)
	}
	// The matte: a black 7pt line along the left edge, inside the box.
	foundDark := false
	for dy := -8; dy <= 8; dy++ {
		for ddx := 0; ddx < 12; ddx++ {
			pc := probe(box.Min.X+ddx, box.Min.Y+box.Dy()/2+dy)
			if pc.R < 100 && pc.G < 100 && pc.B < 100 {
				foundDark = true
			}
		}
	}
	if !foundDark {
		t.Error("no dark matte pixels along the left edge — the frame line was not drawn")
	}
	// Shadow: black 100% alpha, small blur — below the bottom edge must be
	// darkened, and it is drawn only when the shadow exists at all.
	c = probe(box.Min.X+box.Dx()/2, box.Max.Y+4)
	if c.R > 220 {
		t.Errorf("pixel below the box = rgb(%d,%d,%d), want visible shadow darkening", c.R, c.G, c.B)
	}
}
