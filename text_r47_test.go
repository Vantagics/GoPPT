package gopresentation

// The r47 connector 3D-frame tests.
//
// Pinned against COM exports: a connector whose spPr carries
// <a:scene3d><a:sp3d><a:bevelT prst="coolSlant"> renders its stroke as a
// bevelled band — upper side lit (~1.15x the line colour), lower side in
// shade (~0.56x), a 1px highlight just past the lower edge. Five slides of
// the 00022693 bench (20/23/26/27 flow diagrams) carry the effect on every
// connector; the arrowheads stay the plain line colour, no bevel bands.

import (
	"bytes"
	"strings"
	"testing"
)

const bevelCxnSlide = `<p:cxnSp>
  <p:nvCxnSpPr><p:cNvPr id="2" name="Arrow"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
  <p:spPr>
    <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="0"/></a:xfrm>
    <a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>
    <a:ln w="76200"><a:headEnd type="none" w="sm" len="sm"/><a:tailEnd type="triangle" w="sm" len="sm"/></a:ln>
    <a:scene3d><a:camera prst="orthographicFront"/><a:lightRig rig="threePt" dir="t"/></a:scene3d>
    <a:sp3d><a:bevelT w="165100" prst="coolSlant"/></a:sp3d>
  </p:spPr>
  <p:style><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef></p:style>
</p:cxnSp>`

// TestReaderCarriesConnectorBevel: the <a:bevelT prst> inside the
// connector's <a:sp3d> must land on the line shape — dropping it flatlined
// the flow diagrams (slide23 measured 6.94 vs 5.60 once rendered).
func TestReaderCarriesConnectorBevel(t *testing.T) {
	parts := zipParts(t, writeToBytes(t, New()))
	parts["ppt/slides/slide1.xml"] = []byte(bodyPrSlide(bevelCxnSlide))
	pkg := buildZip(t, parts)
	pres, err := (&PPTXReader{}).ReadFromReader(bytes.NewReader(pkg), int64(len(pkg)))
	if err != nil {
		t.Fatalf("read the package: %v", err)
	}
	var ln *LineShape
	for _, sh := range pres.GetAllSlides()[0].GetShapes() {
		if l, ok := sh.(*LineShape); ok {
			ln = l
		}
	}
	if ln == nil {
		t.Fatal("connector dropped")
	}
	if got := ln.GetBevelTop(); got != "coolSlant" {
		t.Errorf("bevelTop = %q, want %q", got, "coolSlant")
	}
}

// TestWriterRoundTripsConnectorBevel: a bevelled connector writes scene3d
// then sp3d after the <a:ln> (CT_ShapeProperties order) and reads back with
// the preset intact; a plain connector writes neither element.
func TestWriterRoundTripsConnectorBevel(t *testing.T) {
	p := New()
	sl := p.GetActiveSlide()
	ln := sl.CreateLineShape()
	ln.SetOffsetX(914400).SetOffsetY(914400).SetSize(3657600, 0)
	ln.SetLineWidth(6)
	ln.SetLineColor(NewColor("FF1F497D"))
	ln.SetBevelTop("coolSlant")

	data := writeToBytes(t, p)
	xml := string(zipParts(t, data)["ppt/slides/slide1.xml"])
	if !strings.Contains(xml, `<a:scene3d>`) || !strings.Contains(xml, `prst="coolSlant"`) {
		t.Errorf("the bevel did not reach the XML:\n%s", xml)
	}
	lnEnd := strings.Index(xml, "</a:ln>")
	s3d := strings.Index(xml, "<a:scene3d>")
	sp3 := strings.Index(xml, "<a:sp3d>")
	if lnEnd < 0 || s3d < lnEnd || sp3 < s3d {
		t.Errorf("scene3d/sp3d must follow </a:ln> in order: lnEnd=%d scene3d=%d sp3d=%d", lnEnd, s3d, sp3)
	}

	pres, err := (&PPTXReader{}).ReadFromReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatalf("re-read: %v", err)
	}
	var back *LineShape
	for _, sh := range pres.GetAllSlides()[0].GetShapes() {
		if l, ok := sh.(*LineShape); ok {
			back = l
		}
	}
	if back == nil || back.GetBevelTop() != "coolSlant" {
		t.Errorf("bevel lost on round trip: %+v", back)
	}

	// A connector without the effect writes no 3D elements.
	p2 := New()
	l2 := p2.GetActiveSlide().CreateLineShape()
	l2.SetOffsetX(914400).SetOffsetY(914400).SetSize(3657600, 0)
	xml2 := string(zipParts(t, writeToBytes(t, p2))["ppt/slides/slide1.xml"])
	if strings.Contains(xml2, "<a:sp3d>") {
		t.Errorf("plain connector must not carry sp3d:\n%s", xml2)
	}
}

// TestBevelLineBands: a horizontal 6pt bevelled connector splits across its
// normal — lit above, base colour at the axis, shade below, highlight just
// past the lower edge.
func TestBevelLineBands(t *testing.T) {
	p := New()
	sl := p.GetActiveSlide()
	ln := sl.CreateLineShape()
	ln.BaseShape.SetOffsetX(571500).SetOffsetY(2857500)
	ln.BaseShape.SetWidth(4572000).SetHeight(0)
	ln.SetLineWidth(6)
	ln.SetLineColor(NewColor("FF4F81BD"))
	ln.SetBevelTop("coolSlant")

	opts := DefaultRenderOptions()
	opts.Width = 1600
	img, err := p.SlideToImage(0, opts)
	if err != nil {
		t.Fatalf("render: %v", err)
	}

	// Line runs y=500 (2857500 EMU * 1600/9144000 = 500), x 100..900.
	baseLum := 132 // 4F81BD = (79,129,189) → (79+129+189)/3
	px := func(x, y int) int {
		r, g, b, _ := img.At(x, y).RGBA()
		return (int(r>>8) + int(g>>8) + int(b>>8)) / 3
	}

	// Upper side (y-3): lit band, brighter than the base colour.
	if got := px(400, 497); got <= baseLum {
		t.Errorf("lit band y-3 luminance = %d, want > %d", got, baseLum)
	}
	// Lower side (y+3): shade, well below the base colour.
	if got := px(400, 503); got >= baseLum-20 {
		t.Errorf("shade band y+3 luminance = %d, want < %d", got, baseLum-20)
	}
	// Highlight just past the lower edge (y+5..y+10): near-white.
	best := 0
	for y := 505; y <= 510; y++ {
		if l := px(400, y); l > best {
			best = l
		}
	}
	if best < 200 {
		t.Errorf("bottom highlight peak luminance = %d, want >= 200", best)
	}
}
