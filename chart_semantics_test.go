package gopresentation

import (
	"bytes"
	"testing"
)

// The chart semantics tests.
//
// Deck 00022823 slide23 (a horizontal bar chart) exposed four gaps at once:
// the first category drew at the top instead of the bottom, the value axis
// ignored its number format, <c:dPt> point overrides clobbered the series
// fill, and <c:manualLayout> was never applied. Each test here pins one of
// those behaviours with PowerPoint-measured expectations.

// dPtChartXML mirrors deck 00022823 chart4.xml: a blue series whose first
// data point is overridden green, a "#,##0" value axis and a manualLayout.
const dPtChartXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart>
    <c:autoTitleDeleted val="1"/>
    <c:plotArea>
      <c:layout><c:manualLayout><c:layoutTarget val="inner"/><c:xMode val="edge"/><c:yMode val="edge"/><c:x val="7.7777777777777779E-2"/><c:y val="5.2757793764988022E-2"/><c:w val="0.84310606566128388"/><c:h val="0.83932853717026379"/></c:manualLayout></c:layout>
      <c:barChart>
        <c:barDir val="bar"/>
        <c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:tx><c:strRef><c:f>Sheet3!$C$31</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>max gates</c:v></c:pt></c:strCache></c:strRef></c:tx>
          <c:spPr><a:solidFill><a:srgbClr val="0070C0"/></a:solidFill></c:spPr>
          <c:dPt>
            <c:idx val="0"/>
            <c:bubble3D val="0"/>
            <c:spPr><a:solidFill><a:srgbClr val="00B050"/></a:solidFill></c:spPr>
          </c:dPt>
          <c:cat><c:strRef><c:f>Sheet3!$D$30:$G$30</c:f><c:strCache><c:ptCount val="4"/><c:pt idx="0"><c:v>Our Results</c:v></c:pt><c:pt idx="1"><c:v>TASTY</c:v></c:pt></c:strCache></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet3!$D$31:$G$31</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="2"/><c:pt idx="0"><c:v>1073741824</c:v></c:pt><c:pt idx="1"><c:v>4194304</c:v></c:pt></c:numCache></c:numRef></c:val>
        </c:ser>
        <c:axId val="1"/>
        <c:axId val="2"/>
      </c:barChart>
      <c:catAx>
        <c:axId val="1"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="l"/>
        <c:numFmt formatCode="General" sourceLinked="1"/>
        <c:crossAx val="2"/>
      </c:catAx>
      <c:valAx>
        <c:axId val="2"/>
        <c:scaling><c:orientation val="minMax"/><c:max val="1200000000"/><c:min val="0"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="b"/>
        <c:numFmt formatCode="#,##0" sourceLinked="0"/>
        <c:crossAx val="1"/>
        <c:majorUnit val="1000000000"/>
      </c:valAx>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`

// TestChartReadsDataPointOverrides: the <c:dPt> colour must land on the
// point, not clobber the series fill — the old reader routed every series-
// level spPr colour to serFill, so the override (which comes second) turned
// the whole series green.
func TestChartReadsDataPointOverrides(t *testing.T) {
	chart := parseChartXML([]byte(dPtChartXML), nil)
	if chart == nil {
		t.Fatal("parseChartXML returned nil")
	}
	bc, ok := chart.plotArea.GetType().(*BarChart)
	if !ok {
		t.Fatalf("plot type = %T, want *BarChart", chart.plotArea.GetType())
	}
	if len(bc.Series) != 1 {
		t.Fatalf("series count = %d, want 1", len(bc.Series))
	}
	s := bc.Series[0]
	if want := NewColor("0070C0"); s.FillColor.ARGB != want.ARGB {
		t.Errorf("series fill = %s, want %s (dPt must not clobber it)", s.FillColor.ARGB, want.ARGB)
	}
	if len(s.PointColors) != 1 {
		t.Fatalf("point overrides = %v, want exactly idx 0", s.PointColors)
	}
	if want := NewColor("00B050"); s.PointColors[0].ARGB != want.ARGB {
		t.Errorf("point 0 colour = %s, want %s", s.PointColors[0].ARGB, want.ARGB)
	}
}

// TestChartReadsAxisNumberFormatAndManualLayout pins the axis numFmt and the
// manualLayout fractions (exponent notation included, as PowerPoint writes).
func TestChartReadsAxisNumberFormatAndManualLayout(t *testing.T) {
	chart := parseChartXML([]byte(dPtChartXML), nil)
	if chart == nil {
		t.Fatal("parseChartXML returned nil")
	}
	axY := chart.plotArea.GetAxisY()
	if axY == nil {
		t.Fatal("value axis missing")
	}
	if axY.NumberFormat != "#,##0" {
		t.Errorf("value axis numFmt = %q, want %q", axY.NumberFormat, "#,##0")
	}
	if axX := chart.plotArea.GetAxisX(); axX != nil && axX.NumberFormat != "" {
		t.Errorf("category axis numFmt = %q, want empty (General)", axX.NumberFormat)
	}
	L := chart.plotArea.layout
	if L == nil {
		t.Fatal("manualLayout not read")
	}
	// The exact fractions from deck 00022823 chart4.xml.
	if L.x < 0.07777 || L.x > 0.07778 {
		t.Errorf("layout x = %v, want 0.077777...", L.x)
	}
	if L.y < 0.052757 || L.y > 0.052758 {
		t.Errorf("layout y = %v, want 5.275779e-2", L.y)
	}
	if L.w < 0.84310 || L.w > 0.84311 {
		t.Errorf("layout w = %v, want 0.843106...", L.w)
	}
	if L.h < 0.839328 || L.h > 0.839329 {
		t.Errorf("layout h = %v, want 0.839328...", L.h)
	}
}

// TestChartFormatNumberWith covers the "#,##0" family the renderer resolves.
func TestChartFormatNumberWith(t *testing.T) {
	cases := []struct {
		v      float64
		format string
		want   string
	}{
		{0, "#,##0", "0"},
		{1000000000, "#,##0", "1,000,000,000"},
		{1073741824, "#,##0", "1,073,741,824"},
		{-23456, "#,##0", "-23,456"},
		{1234.56, "#,##0.0", "1,234.6"},
		{7, "#,##0.00", "7.00"},
		{0.25, "0%", "25%"},
		{0.256, "0.0%", "25.6%"},
		{42, "General", "42"},
		{42, "", "42"},
		// Unmodelled codes fall back to the plain rendering.
		{1234.5, `#\ ???/???`, "1234.5"},
	}
	for _, c := range cases {
		if got := chartFormatNumberWith(c.v, c.format); got != c.want {
			t.Errorf("chartFormatNumberWith(%v, %q) = %q, want %q", c.v, c.format, got, c.want)
		}
	}
}

// TestHorizontalBarFirstCategoryAtBottom renders a two-category horizontal
// bar chart whose first point is overridden green: OOXML draws category 0 at
// the bottom, so the green bar must land in the lower half and the blue
// series fill in the upper half.
func TestHorizontalBarFirstCategoryAtBottom(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()
	chart := slide.CreateChartShape()
	chart.BaseShape.SetOffsetX(457200).SetOffsetY(457200)
	chart.BaseShape.SetWidth(8229600).SetHeight(4572000)

	bar := NewBarChart()
	bar.BarDirection = BarDirectionHorizontal
	ser := NewChartSeriesOrdered("gates", []string{"Big", "Small"}, []float64{1000, 100})
	ser.SetFillColor(NewColor("0070C0"))
	ser.PointColors = map[int]Color{0: NewColor("00B050")}
	bar.AddSeries(ser)
	chart.GetPlotArea().SetType(bar)

	fc := NewFontCache()
	requireAnyFont(t, fc)
	opts := DefaultRenderOptions()
	opts.Width = 720
	opts.FontCache = fc
	img, err := p.SlideToImage(0, opts)
	if err != nil {
		t.Fatalf("render: %v", err)
	}
	b := img.Bounds()
	// Frame: offset 457200 EMU = 36px, height 4572000 = 360px. Rows split at
	// the frame's vertical centre, not the image's.
	midY := 36 + 360/2
	count := func(pred func(r, g, bl uint32) bool) (top, bottom int) {
		for y := b.Min.Y; y < b.Max.Y; y++ {
			for x := b.Min.X; x < b.Max.X; x++ {
				r, g, bl, _ := img.At(x, y).RGBA()
				if pred(r>>8, g>>8, bl>>8) {
					if y < midY {
						top++
					} else {
						bottom++
					}
				}
			}
		}
		return
	}
	greenTop, greenBottom := count(func(r, g, bl uint32) bool {
		return g > 120 && g > r+40 && g > bl+40
	})
	blueTop, blueBottom := count(func(r, g, bl uint32) bool {
		return bl > 120 && bl > r+40 && bl > g+40
	})
	if greenTop != 0 || greenBottom == 0 {
		t.Errorf("green (category 0) pixels top=%d bottom=%d, want all in the bottom half", greenTop, greenBottom)
	}
	if blueTop == 0 || blueBottom > blueTop/10 {
		t.Errorf("blue (category 1) pixels top=%d bottom=%d, want them in the top half", blueTop, blueBottom)
	}
}

// TestWriterRoundTripsDataPointColors: the writer must emit <c:dPt> and the
// axis <c:numFmt> so a read-modify-write keeps the highlights and format.
func TestWriterRoundTripsDataPointColors(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()
	chart := slide.CreateChartShape()
	bar := NewBarChart()
	bar.BarDirection = BarDirectionHorizontal
	ser := NewChartSeriesOrdered("gates", []string{"A", "B"}, []float64{10, 20})
	ser.SetFillColor(NewColor("0070C0"))
	ser.PointColors = map[int]Color{0: NewColor("00B050")}
	bar.AddSeries(ser)
	chart.GetPlotArea().SetType(bar)
	chart.GetPlotArea().GetAxisY().NumberFormat = "#,##0"

	data := writeToBytes(t, p)
	pres, err := (&PPTXReader{}).ReadFromReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatalf("read back: %v", err)
	}
	var got *ChartShape
	for _, sh := range pres.GetAllSlides()[0].shapes {
		if cs, ok := sh.(*ChartShape); ok {
			got = cs
			break
		}
	}
	if got == nil {
		t.Fatal("chart shape dropped on round-trip")
	}
	bc, ok := got.plotArea.GetType().(*BarChart)
	if !ok || len(bc.Series) != 1 {
		t.Fatalf("round-tripped chart = %T with %d series", got.plotArea.GetType(), len(bc.Series))
	}
	rs := bc.Series[0]
	if want := NewColor("0070C0"); rs.FillColor.ARGB != want.ARGB {
		t.Errorf("series fill = %s, want %s", rs.FillColor.ARGB, want.ARGB)
	}
	if len(rs.PointColors) != 1 || rs.PointColors[0].ARGB != NewColor("00B050").ARGB {
		t.Errorf("point colours = %v, want idx0 = 00B050", rs.PointColors)
	}
	if got.plotArea.GetAxisY().NumberFormat != "#,##0" {
		t.Errorf("axis numFmt = %q, want #,##0", got.plotArea.GetAxisY().NumberFormat)
	}
}

// TestChartManualLayoutPinsPlotRect: with a manualLayout the plot rect is the
// layout fractions of the frame, verified here via the value-axis scale the
// bars map onto — 100 of max 120 must reach 5/6 of the data rect width. The
// layout is frame-relative, so with no title/legend the bars must start near
// x + 0.0778*w rather than after a label gutter sized from the frame edge.
func TestChartManualLayoutPinsPlotRect(t *testing.T) {
	const chartXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart><c:plotArea>
    <c:layout><c:manualLayout><c:layoutTarget val="inner"/><c:xMode val="edge"/><c:yMode val="edge"/><c:x val="0.5"/><c:y val="0.1"/><c:w val="0.4"/><c:h val="0.8"/></c:manualLayout></c:layout>
    <c:barChart>
      <c:barDir val="bar"/>
      <c:grouping val="clustered"/>
      <c:ser>
        <c:idx val="0"/><c:order val="0"/>
        <c:tx><c:strRef><c:f>S!$A$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>s</c:v></c:pt></c:strCache></c:strRef></c:tx>
        <c:spPr><a:solidFill><a:srgbClr val="0070C0"/></a:solidFill></c:spPr>
        <c:cat><c:strRef><c:f>S!$A$2</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>k</c:v></c:pt></c:strCache></c:strRef></c:cat>
        <c:val><c:numRef><c:f>S!$A$3</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="1"/><c:pt idx="0"><c:v>100</c:v></c:pt></c:numCache></c:numRef></c:val>
      </c:ser>
      <c:axId val="1"/><c:axId val="2"/>
    </c:barChart>
    <c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="2"/></c:catAx>
    <c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/><c:max val="120"/><c:min val="0"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="1"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`

	parts := zipParts(t, writeToBytes(t, New()))
	parts["ppt/slides/slide1.xml"] = []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
    <p:graphicFrame>
      <p:nvGraphicFramePr><p:cNvPr id="2" name="Chart 1"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
      <p:xfrm><a:off x="0" y="0"/><a:ext cx="9525000" cy="5715000"/></p:xfrm>
      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId2"/></a:graphicData></a:graphic>
    </p:graphicFrame>
  </p:spTree></p:cSld>
</p:sld>`)
	parts["ppt/slides/_rels/slide1.xml.rels"] = []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
		`<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>` +
		`<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>` +
		`</Relationships>`)
	parts["ppt/charts/chart1.xml"] = []byte(chartXML)
	pkg := buildZip(t, parts)
	pres, err := (&PPTXReader{}).ReadFromReader(bytes.NewReader(pkg), int64(len(pkg)))
	if err != nil {
		t.Fatalf("read: %v", err)
	}
	slide := pres.GetAllSlides()[0]
	var cs *ChartShape
	for _, sh := range slide.shapes {
		if c, ok := sh.(*ChartShape); ok {
			cs = c
			break
		}
	}
	if cs == nil {
		t.Fatal("chart shape missing")
	}
	if cs.BaseShape.GetWidth() != 9525000 {
		t.Fatalf("frame width %d EMU, want 9525000", cs.BaseShape.GetWidth())
	}

	fc := NewFontCache()
	requireAnyFont(t, fc)
	opts := DefaultRenderOptions()
	opts.Width = 720
	opts.FontCache = fc
	img, err := pres.SlideToImage(0, opts)
	if err != nil {
		t.Fatalf("render: %v", err)
	}
	b := img.Bounds()
	imgW := b.Dx()
	firstBlue := -1
	lastBlue := -1
	for y := b.Min.Y; y < b.Max.Y; y++ {
		for x := b.Min.X; x < b.Max.X; x++ {
			r, g, bl, _ := img.At(x, y).RGBA()
			_ = g
			if bl > 120 && int(bl>>8) > int(r>>8)+40 {
				if firstBlue < 0 {
					firstBlue = x
				}
				lastBlue = x
			}
		}
	}
	if firstBlue < 0 {
		t.Fatal("no blue bar rendered")
	}
	// The frame spans 9525000 of the 9144000-EMU slide width; scale the
	// rendered width accordingly (the px-per-EMU factor is the renderer's
	// business, the fractions are ours).
	frameW := float64(imgW) * 9525000.0 / 9144000.0
	wantStart := 0.5 * frameW // layout x=0.5
	if float64(firstBlue) < wantStart-8 || float64(firstBlue) > wantStart+8 {
		t.Errorf("bar starts at x=%d, want ~%.0f (layout x=0.5 of the %.0fpx frame)", firstBlue, wantStart, frameW)
	}
	// Value 100 of max 120 lands at layout x + 5/6 of layout w.
	wantEnd := wantStart + (100.0/120.0)*0.4*frameW
	if float64(lastBlue) < wantEnd-8 || float64(lastBlue) > wantEnd+8 {
		t.Errorf("bar ends at x=%d, want ~%.0f (100 of max 120 in a %.0fpx rect)", lastBlue, wantEnd, 0.4*frameW)
	}
}

// numFmtChartXML is the layout-pinned bar chart with a configurable value
// axis format; used to prove the tick labels really change with numFmt.
func numFmtChartXML(numFmt string) string {
	fmtEl := ""
	if numFmt != "" {
		fmtEl = `<c:numFmt formatCode="` + numFmt + `" sourceLinked="0"/>`
	}
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart><c:plotArea>
    <c:layout><c:manualLayout><c:layoutTarget val="inner"/><c:xMode val="edge"/><c:yMode val="edge"/><c:x val="0.1"/><c:y val="0.1"/><c:w val="0.8"/><c:h val="0.8"/></c:manualLayout></c:layout>
    <c:barChart>
      <c:barDir val="bar"/>
      <c:grouping val="clustered"/>
      <c:ser>
        <c:idx val="0"/><c:order val="0"/>
        <c:tx><c:strRef><c:f>S!$A$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>s</c:v></c:pt></c:strCache></c:strRef></c:tx>
        <c:spPr><a:solidFill><a:srgbClr val="0070C0"/></a:solidFill></c:spPr>
        <c:cat><c:strRef><c:f>S!$A$2</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>k</c:v></c:pt></c:strCache></c:strRef></c:cat>
        <c:val><c:numRef><c:f>S!$A$3</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="1"/><c:pt idx="0"><c:v>1000000000</c:v></c:pt></c:numCache></c:numRef></c:val>
      </c:ser>
      <c:axId val="1"/><c:axId val="2"/>
    </c:barChart>
    <c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="2"/></c:catAx>
    <c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/><c:max val="1200000000"/><c:min val="0"/></c:scaling><c:delete val="0"/><c:axPos val="b"/>` + fmtEl + `<c:crossAx val="1"/><c:majorUnit val="1000000000"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`
}

// renderNumFmtChart renders the chart above and counts dark text pixels in
// the value-tick-label strip below the pinned plot.
func renderNumFmtChart(t *testing.T, numFmt string) int {
	t.Helper()
	parts := zipParts(t, writeToBytes(t, New()))
	parts["ppt/slides/slide1.xml"] = []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
    <p:graphicFrame>
      <p:nvGraphicFramePr><p:cNvPr id="2" name="Chart 1"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
      <p:xfrm><a:off x="0" y="0"/><a:ext cx="9525000" cy="5715000"/></p:xfrm>
      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId2"/></a:graphicData></a:graphic>
    </p:graphicFrame>
  </p:spTree></p:cSld>
</p:sld>`)
	parts["ppt/slides/_rels/slide1.xml.rels"] = []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
		`<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>` +
		`<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>` +
		`</Relationships>`)
	parts["ppt/charts/chart1.xml"] = []byte(numFmtChartXML(numFmt))
	pkg := buildZip(t, parts)
	pres, err := (&PPTXReader{}).ReadFromReader(bytes.NewReader(pkg), int64(len(pkg)))
	if err != nil {
		t.Fatalf("read: %v", err)
	}
	fc := NewFontCache()
	requireAnyFont(t, fc)
	opts := DefaultRenderOptions()
	opts.Width = 720
	opts.FontCache = fc
	img, err := pres.SlideToImage(0, opts)
	if err != nil {
		t.Fatalf("render: %v", err)
	}
	b := img.Bounds()
	// Plot bottom = 0.9 of the frame height (layout y+h); the tick labels sit
	// just below it. Frame height = 5715000/6858000 of the image height.
	top := int(0.9 * 5715000.0 / 6858000.0 * float64(b.Dy()))
	minX, maxX := -1, -1
	for y := top + 2; y < top+40 && y < b.Max.Y; y++ {
		for x := b.Min.X; x < b.Max.X; x++ {
			r, g, bl, _ := img.At(x, y).RGBA()
			if r < 100<<8 && g < 100<<8 && bl < 100<<8 {
				if minX < 0 {
					minX = x
				}
				maxX = x
			}
		}
	}
	if minX < 0 {
		t.Fatal("no tick label text rendered below the plot")
	}
	return maxX - minX
}

// TestChartAxisNumberFormatChangesTickLabels: with "#,##0" the tick
// "1,000,000,000" carries two commas the General rendering lacks — the
// label strip must get visibly darker. This is the renderer-side sentinel
// for the axisNumFormat path the reader feeds.
func TestChartAxisNumberFormatChangesTickLabels(t *testing.T) {
	general := renderNumFmtChart(t, "")
	formatted := renderNumFmtChart(t, "#,##0")
	if formatted < general+3 {
		t.Errorf("tick label width general=%d formatted=%d; the #,##0 format did not reach the renderer", general, formatted)
	}
}
