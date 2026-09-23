package gopresentation

// The r44 horizontal-bar-chart semantics tests.
//
// Pinned against COM exports of deck 00022823 slides 22/23 (the "solver
// components" and "Results: Scalability" horizontal bar charts):
//
//   - The value axis of a horizontal bar chart runs along the plot bottom
//     and PowerPoint draws it like any other axis line; the renderer used to
//     draw only the category axis line.
//   - The value-axis tick labels below that axis sit a full em down
//     (glyph top = axis + 1em − 2px: 84px baseline drop at 24pt on slide23,
//     70px at 20pt on slide22), not at +2+ascent.
//   - The plot's right edge pulls in so the tick label centred at it clears
//     the chart frame (slide23: pinned 1667 → measured 1639).
//   - A chartSpace-level <c:txPr> is the default text size for every element
//     that declares none (chart3 states 20pt there and renders 20pt value
//     labels off a size-less value axis).
//   - An axis <c:spPr><a:ln> overrides the 134-grey default stroke, and
//     <a:noFill> removes the line entirely (chart5's left value axis).

import (
	"image"
	"image/color"
	"testing"
)

const axisLineBlackChart = `<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
<c:chart><c:plotArea><c:layout/>
<c:barChart><c:barDir val="bar"/><c:grouping val="clustered"/>
<c:ser><c:idx val="0"/><c:order val="0"/>
<c:cat><c:strRef><c:f>Sheet1!$A$1:$A$2</c:f><c:strCache><c:ptCount val="2"/>
<c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt>
</c:strCache></c:strRef></c:cat>
<c:val><c:numRef><c:f>Sheet1!$B$1:$B$2</c:f><c:numCache><c:ptCount val="2"/>
<c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt>
</c:numCache></c:numRef></c:val>
</c:ser>
<c:axId val="111"/><c:axId val="222"/>
</c:barChart>
<c:catAx><c:axId val="111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:majorTickMark val="none"/><c:tickLblPos val="low"/><c:crossAx val="222"/></c:catAx>
<c:valAx><c:axId val="222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:majorTickMark val="out"/><c:tickLblPos val="nextTo"/><c:spPr><a:ln w="19050"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></c:spPr><c:crossAx val="111"/></c:valAx>
</c:plotArea></c:chart></c:chartSpace>`

const axisLineNoFillChart = `<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
<c:chart><c:plotArea><c:layout/>
<c:scatterChart><c:scatterStyle val="lineMarker"/>
<c:ser><c:idx val="0"/><c:order val="0"/>
<c:xVal><c:numRef><c:f>Sheet1!$A$1:$A$3</c:f><c:numCache><c:ptCount val="3"/>
<c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt>
</c:numCache></c:numRef></c:xVal>
<c:yVal><c:numRef><c:f>Sheet1!$B$1:$B$3</c:f><c:numCache><c:ptCount val="3"/>
<c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt>
</c:numCache></c:numRef></c:yVal>
</c:ser>
<c:axId val="111"/><c:axId val="222"/>
</c:scatterChart>
<c:valAx><c:axId val="111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:majorTickMark val="none"/><c:crossAx val="222"/></c:valAx>
<c:valAx><c:axId val="222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:majorTickMark val="none"/><c:spPr><a:ln><a:noFill/></a:ln></c:spPr><c:crossAx val="111"/></c:valAx>
</c:plotArea></c:chart></c:chartSpace>`

const chartSpaceDefaultFontChart = `<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
<c:chart><c:plotArea><c:layout/>
<c:barChart><c:barDir val="bar"/><c:grouping val="clustered"/>
<c:ser><c:idx val="0"/><c:order val="0"/>
<c:dLbls><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1200"/></a:pPr><a:endParaRPr/></a:p></c:txPr></c:dLbls>
<c:cat><c:strRef><c:f>Sheet1!$A$1:$A$2</c:f><c:strCache><c:ptCount val="2"/>
<c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt>
</c:strCache></c:strRef></c:cat>
<c:val><c:numRef><c:f>Sheet1!$B$1:$B$2</c:f><c:numCache><c:ptCount val="2"/>
<c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt>
</c:numCache></c:numRef></c:val>
</c:ser>
<c:axId val="111"/><c:axId val="222"/>
</c:barChart>
<c:catAx><c:axId val="111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:crossAx val="222"/><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="2400"/></a:pPr><a:endParaRPr/></a:p></c:txPr></c:catAx>
<c:valAx><c:axId val="222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="111"/><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:endParaRPr/></a:p></c:txPr></c:valAx>
</c:plotArea></c:chart>
<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="2000"/></a:pPr><a:endParaRPr/></a:p></c:txPr>
</c:chartSpace>`

// TestReaderAxisLineStroke reads the axis line's own <c:spPr><a:ln>: an
// explicit colour/width replaces the 134-grey default (chart3's value axis
// is black 0.75pt) and <a:noFill> removes the line (chart5's left axis).
func TestReaderAxisLineStroke(t *testing.T) {
	chart := parseChartXML([]byte(axisLineBlackChart), nil)
	if chart == nil {
		t.Fatal("parseChartXML returned nil")
	}
	val := chart.plotArea.GetAxisY()
	if val == nil {
		t.Fatal("value axis dropped")
	}
	if val.OutlineNoFill {
		t.Error("black axis line read as noFill")
	}
	if got := val.OutlineColor.ARGB; got != "FFFF0000" {
		t.Errorf("axis line colour = %q, want FFFF0000", got)
	}
	if got := val.OutlineWidth; got != 2 {
		t.Errorf("axis line width = %d pt, want 2 (19050 EMU)", got)
	}

	chart2 := parseChartXML([]byte(axisLineNoFillChart), nil)
	if chart2 == nil {
		t.Fatal("parseChartXML returned nil for the noFill fixture")
	}
	left := chart2.plotArea.GetAxisY()
	if left == nil {
		t.Fatal("left axis dropped")
	}
	if !left.OutlineNoFill {
		t.Error("noFill axis line not recorded")
	}
}

// TestRendererAxisStrokeHonoursDeclaration: the renderer draws the declared
// stroke, skips a noFill axis, and keeps the 134-grey default otherwise.
func TestRendererAxisStrokeHonoursDeclaration(t *testing.T) {
	r := &renderer{scaleX: 1600.0 / 9144000.0}
	def := color.RGBA{R: 134, G: 134, B: 134, A: 255}

	c, w, ok := r.chartAxisStroke(nil, def, 2)
	if !ok || c != def || w != 2 {
		t.Errorf("nil axis stroke = %v/%d/%v, want default", c, w, ok)
	}

	black := NewColor("000000")
	a := &ChartAxis{OutlineColor: black, OutlineWidth: 1}
	c, w, ok = r.chartAxisStroke(a, def, 2)
	if !ok || c != (color.RGBA{R: 0, G: 0, B: 0, A: 255}) || w < 2 {
		t.Errorf("declared stroke = %v/%d/%v, want black width >=2", c, w, ok)
	}

	nf := &ChartAxis{OutlineNoFill: true}
	if _, _, ok = r.chartAxisStroke(nf, def, 2); ok {
		t.Error("noFill axis still drew a line")
	}
}

// TestChartSpaceDefaultFontSizeCascades: the chartSpace-level <c:txPr> size
// applies to every chart font that declares none — deck 2 chart3 states 20pt
// there and its size-less value axis renders 20pt tick labels, not the 10pt
// model fallback. An element's own size (the 24pt category axis, the 12pt
// data labels) still wins.
func TestChartSpaceDefaultFontSizeCascades(t *testing.T) {
	chart := parseChartXML([]byte(chartSpaceDefaultFontChart), nil)
	if chart == nil {
		t.Fatal("parseChartXML returned nil")
	}
	if ax := chart.plotArea.GetAxisY(); ax.Font.Size != 20 {
		t.Errorf("size-less value axis font = %d pt, want 20 from chartSpace default", ax.Font.Size)
	}
	if ax := chart.plotArea.GetAxisX(); ax.Font.Size != 24 {
		t.Errorf("category axis font = %d pt, want its own 24", ax.Font.Size)
	}
	if chart.legend.Font.Size != 20 {
		t.Errorf("legend font = %d pt, want 20 from chartSpace default", chart.legend.Font.Size)
	}
	ser := chart.plotArea.GetType().(*BarChart).Series[0]
	if ser.Font == nil || ser.Font.Size != 12 {
		t.Errorf("series label font = %v, want its own 12", ser.Font)
	}
}

// renderHorizontalBar renders an API-built horizontal bar chart at 1600px and
// returns the image plus the row of the plot-bottom value-axis line (the
// long 134-grey horizontal run) and the plot line's right end.
func renderHorizontalBar(t *testing.T, mut func(c *ChartShape)) (image.Image, int, int) {
	t.Helper()
	fc := NewFontCache()
	if pickInstalledFont(fc) == "" {
		t.Skip("no fonts installed on this machine")
	}
	p := New()
	slide := p.GetActiveSlide()
	chart := slide.CreateChartShape()
	chart.BaseShape.SetOffsetX(500000).SetOffsetY(500000)
	chart.BaseShape.SetWidth(8600000).SetHeight(5000000)
	bc := NewBarChart()
	bc.BarDirection = BarDirectionHorizontal
	ser := NewChartSeriesOrdered("s", []string{"A", "B", "C", "D"}, []float64{1, 2, 3, 4})
	bc.AddSeries(ser)
	chart.GetPlotArea().SetType(bc)
	if mut != nil {
		mut(chart)
	}

	opts := DefaultRenderOptions()
	opts.Width = 1600
	opts.FontCache = fc
	img, err := p.SlideToImage(0, opts)
	if err != nil {
		t.Fatalf("render: %v", err)
	}
	return img, 0, 0
}

// findBarAxisLine locates the plot-bottom value-axis line: the row with the
// longest run of 134-grey pixels, plus that run's right end.
func findBarAxisLine(t *testing.T, img image.Image) (row, rightEnd int) {
	t.Helper()
	w := img.Bounds().Dx()
	h := img.Bounds().Dy()
	bestRow, bestLen, bestRight := -1, 0, 0
	for y := h / 3; y < h; y++ {
		run, right := 0, 0
		for x := 0; x < w; x++ {
			r, g, b, _ := img.At(x, y).RGBA()
			r8, g8, b8 := int(r>>8), int(g>>8), int(b>>8)
			if g8 >= 120 && g8 <= 150 && r8 >= 120 && r8 <= 150 && b8 >= 120 && b8 <= 150 {
				run++
				right = x
			}
		}
		if run > bestLen {
			bestRow, bestLen, bestRight = y, run, right
		}
	}
	if bestLen < 300 {
		t.Fatalf("no plot-bottom axis line found (best run %d px at y=%d)", bestLen, bestRow)
	}
	return bestRow, bestRight
}

// TestHorizontalBarValueAxisLineAndLabelDrop pins two COM-measured behaviours
// of deck 00022823 slide22/23: the value axis line along the plot bottom is
// drawn (the old code drew only the category axis line), and the tick labels
// below it start a full em down (glyph top = axis + 1em − 2px), not hugging
// the axis at +2+ascent.
func TestHorizontalBarValueAxisLineAndLabelDrop(t *testing.T) {
	img, _, _ := renderHorizontalBar(t, func(c *ChartShape) {
		// 24pt tick labels: the COM drop measures 84px at this size.
		c.GetPlotArea().GetAxisY().Font.Size = 24
	})
	axisRow, _ := findBarAxisLine(t, img)

	// First text-ink row below the axis line.
	h := img.Bounds().Dy()
	firstText := -1
	for y := axisRow + 3; y < h && y < axisRow+200; y++ {
		cnt := 0
		for x := 0; x < img.Bounds().Dx(); x++ {
			r, g, b, _ := img.At(x, y).RGBA()
			if int(r>>8) < 110 && int(g>>8) < 110 && int(b>>8) < 110 {
				cnt++
			}
		}
		if cnt >= 3 {
			firstText = y
			break
		}
	}
	if firstText < 0 {
		t.Fatal("no tick label ink below the axis line")
	}
	drop := firstText - axisRow
	// 24pt em = 53px: the COM glyph top lands at +51. The old +2+ascent put
	// it at +8; a reverted or retuned constant drifts out of this window.
	if drop < 30 || drop > 75 {
		t.Errorf("tick label glyph top %dpx below the axis, want ~51 (1em−2) — got drop %d", 51, drop)
	}
}

// TestHorizontalBarPlotRightClearsFrame pins the plot-right clamp: the tick
// label centred at the plot's right edge must clear the chart frame, so the
// plot cannot keep a manual-layout right edge that runs to the frame border
// (deck 00022823 slide23: pinned 1667 → PowerPoint 1639). The axis line's
// right end is the plot's right edge.
func TestHorizontalBarPlotRightClearsFrame(t *testing.T) {
	const maxValue = 1000000000.0
	img, _, _ := renderHorizontalBar(t, func(c *ChartShape) {
		c.GetPlotArea().layout = &chartManualLayout{x: 0.08, y: 0.30, w: 0.92, h: 0.60}
		v := 1200000000.0
		c.GetPlotArea().GetAxisY().MaxBounds = &v
		c.GetPlotArea().GetAxisY().NumberFormat = "#,##0"
		ser := NewChartSeriesOrdered("s", []string{"Only"}, []float64{maxValue})
		bc := NewBarChart()
		bc.BarDirection = BarDirectionHorizontal
		bc.AddSeries(ser)
		c.GetPlotArea().SetType(bc)
	})
	_, rightEnd := findBarAxisLine(t, img)

	// Frame right: offset 500000 + width 8600000 EMU at 1600px/9144000.
	spanEMU := float64(500000 + 8600000)
	frameRight := int(spanEMU / 9144000.0 * 1600.0)
	// The widest tick label is ~90px at the 10pt default, so the COM clamp
	// lands ~(half label + 18px) left of the frame edge. The unclamped
	// manual layout runs to frameRight+ (0.92 extends past the frame here).
	if rightEnd > frameRight-30 {
		t.Errorf("plot right edge %d not clamped inside the frame right %d (want <= frameRight-30)", rightEnd, frameRight)
	}
	if rightEnd < frameRight/2 {
		t.Errorf("plot right edge %d collapsed, clamp over-shrank", rightEnd)
	}
}
