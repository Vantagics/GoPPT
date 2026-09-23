package gopresentation

import (
	"bytes"
	"testing"
)

// TestMatchPlaceholderDefByIndex pins the OOXML inheritance key: a slide
// placeholder that declares only <p:ph idx="1"/> (PowerPoint drops the type
// on content placeholders) must find the master's type="body" idx="1"
// definition by index. Requiring type AND idx to match left slide17's
// full-width body rendering as a one-word-per-line column.
func TestMatchPlaceholderDefByIndex(t *testing.T) {
	defs := []layoutPlaceholder{
		{phType: "title", phIdx: 0},
		{phType: "body", phIdx: 1},
		{phType: "sldNum", phIdx: 12},
	}
	ph := NewPlaceholderShape("") // idx-only: no type in the slide markup
	ph.phIdx = 1
	got := matchPlaceholderDef(defs, ph)
	if got == nil || got.phType != "body" || got.phIdx != 1 {
		t.Fatalf("idx-only placeholder matched %+v, want body idx=1", got)
	}
	// An index with no carrier still falls through to the type walk.
	ph2 := NewPlaceholderShape(PlaceholderType("body"))
	ph2.phIdx = 7
	got2 := matchPlaceholderDef(defs, ph2)
	if got2 == nil || got2.phType != "body" {
		t.Fatalf("unmatched idx fell through to type walk, got %+v", got2)
	}
}

// TestReaderSlide17BodyInheritsMasterGeometry is the end-to-end guard: a
// slide body placeholder that carries no type and no xfrm must inherit the
// master body's full-width geometry rather than leaving a default-size box
// that re-wraps every paragraph (deck 00022693 slide17 rendered its body as
// a one-word-per-line column).
func TestReaderSlide17BodyInheritsMasterGeometry(t *testing.T) {
	parts := zipParts(t, writeToBytes(t, New()))
	// Master: define the body placeholder with the full-width geometry the
	// real deck's master carries (scaled down).
	parts["ppt/slideMasters/slideMaster1.xml"] = []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="457200" y="0"/><a:ext cx="8229600" cy="1143000"/></a:xfrm></p:spPr>
      <p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>
    </p:sp>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="3" name="Body"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="457200" y="1325562"/><a:ext cx="8229600" cy="4525963"/></a:xfrm></p:spPr>
      <p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>
    </p:sp>
  </p:spTree></p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst/>
</p:sldMaster>`)
	// Layout: declares the body placeholder but no geometry — the ladder
	// must fall through to the master.
	parts["ppt/slideLayouts/slideLayout1.xml"] = []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="2" name="Body"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>
      <p:spPr/>
      <p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>
    </p:sp>
  </p:spTree></p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>`)
	// Slide: the PowerPoint style idx-only placeholder — no type, no xfrm.
	parts["ppt/slides/slide1.xml"] = []byte(bodyPrSlide(`<p:sp>
  <p:nvSpPr><p:cNvPr id="2" name="Content"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph idx="1"/></p:nvPr></p:nvSpPr>
  <p:spPr/>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Full width body</a:t></a:r></a:p></p:txBody>
</p:sp>`))
	pkg := buildZip(t, parts)
	pres, err := (&PPTXReader{}).ReadFromReader(bytes.NewReader(pkg), int64(len(pkg)))
	if err != nil {
		t.Fatalf("read the package: %v", err)
	}
	var body *PlaceholderShape
	for _, sh := range pres.GetAllSlides()[0].GetShapes() {
		if ph, ok := sh.(*PlaceholderShape); ok && ph.phIdx == 1 && ph.phType == "" {
			body = ph
		}
	}
	if body == nil {
		t.Fatal("idx-only body placeholder dropped")
	}
	// The inherited width must be the master's full-width box.
	if body.width != 8229600 || body.height != 4525963 {
		t.Errorf("body placeholder size = %dx%d, want 8229600x4525963 inherited from master", body.width, body.height)
	}
}

// TestScatterDefaultStrokeMatchesLineChart pins the scatter chart's default
// series stroke end to end: a scatter series with no <a:ln w> must render
// with the line chart's 2.25pt default, not a hairline (chart2 of deck
// 00022693 declares no width yet the COM gold shows ~5px lines). The stroke
// thickness is asserted by ink mass — a hairline covers far fewer pixels
// than the 2.25pt default at the same length.
func TestScatterDefaultStrokeMatchesLineChart(t *testing.T) {
	fc := NewFontCache()
	if pickInstalledFont(fc) == "" {
		t.Skip("no fonts installed on this machine")
	}
	p := New()
	slide := p.GetActiveSlide()
	chart := slide.CreateChartShape()
	chart.BaseShape.SetOffsetX(500000).SetOffsetY(500000)
	chart.BaseShape.SetWidth(8500000).SetHeight(5000000)
	sc := NewScatterChart()
	ser := NewChartSeriesOrdered("",
		[]string{"1", "2", "3", "4", "5", "6", "7", "8"}, []float64{1, 2, 1.5, 3, 2.5, 4, 3.5, 5})
	sc.AddSeries(ser)
	chart.GetPlotArea().SetType(sc)

	opts := DefaultRenderOptions()
	opts.Width = 1600
	opts.FontCache = fc
	img, err := p.SlideToImage(0, opts)
	if err != nil {
		t.Fatalf("render: %v", err)
	}
	// Count the series-stroke pixels: the default palette's first accent
	// (4F81BD-ish blue) over the plot area.
	ink := 0
	for y := 0; y < img.Bounds().Dy(); y++ {
		for x := 0; x < img.Bounds().Dx(); x++ {
			r, g, b, _ := img.At(x, y).RGBA()
			r8, g8, b8 := int(r>>8), int(g>>8), int(b>>8)
			if b8 > 120 && b8-r8 > 40 && b8-g8 > 20 && r8 < 150 {
				ink++
			}
		}
	}
	// An 8-point polyline across a ~970px-wide plot is ~1200px long; a
	// hairline covers ~1200 px of ink, the 2.25pt default ~6000.
	if ink < 3000 {
		t.Errorf("scatter line ink = %d px, want >=3000 (2.25pt stroke, not hairline)", ink)
	}
}
