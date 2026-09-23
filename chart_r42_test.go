package gopresentation

// The r42 chart-axis and series-stroke semantics tests.
//
// Four defaults were pinned against COM exports of deck 00022823 (slides
// 5/6, the "Cost per Genome" date-axis line chart):
//
//   - A chart axis that omits <c:majorTickMark> renders OUTSIDE tick marks
//     (chart1/chart2 show them; chart3/4/5 declare val="none" and show
//     none). The reader used to inherit NewChartAxis's API-level "none".
//   - A line-chart series without an explicit stroke width renders at
//     PowerPoint's 28575 EMU (2.25pt) default, not a hairline.
//   - <c:marker><c:size val="8"/> is 8 POINTS wide: the radius must scale
//     through 12700 EMU/pt before scaleX. The old expression fed the raw
//     point count to a px-per-EMU scale, collapsing the diamond to a 2px dot.
//   - A rotated date label's ink column centers ~5px right of its tick on
//     the COM golds (the ascent side of the glyph box).

import (
	"image"
	"image/color"
	"testing"
)

const tickDefaultChart = `<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
<c:chart><c:plotArea><c:layout/>
<c:lineChart><c:grouping val="standard"/>
<c:ser><c:idx val="0"/><c:order val="0"/>
<c:marker><c:symbol val="diamond"/><c:size val="8"/></c:marker>
<c:cat><c:strRef><c:f>Sheet1!$A$1:$A$3</c:f><c:strCache><c:ptCount val="3"/>
<c:pt idx="0"><c:v>Sep 2001</c:v></c:pt><c:pt idx="1"><c:v>Feb 2002</c:v></c:pt><c:pt idx="2"><c:v>Jul 2002</c:v></c:pt>
</c:strCache></c:strRef></c:cat>
<c:val><c:numRef><c:f>Sheet1!$B$1:$B$3</c:f><c:numCache><c:ptCount val="3"/>
<c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt>
</c:numCache></c:numRef></c:val>
</c:ser>
<c:marker val="1"/>
<c:axId val="111"/><c:axId val="222"/>
</c:lineChart>
<c:catAx><c:axId val="111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="222"/></c:catAx>
<c:valAx><c:axId val="222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:crossAx val="111"/></c:valAx>
</c:plotArea></c:chart></c:chartSpace>`

const tickNoneChart = `<?xml version="1.0" encoding="UTF-8"?>
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
<c:catAx><c:axId val="111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:majorTickMark val="none"/><c:crossAx val="222"/></c:catAx>
<c:valAx><c:axId val="222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:majorTickMark val="none"/><c:crossAx val="111"/></c:valAx>
</c:plotArea></c:chart></c:chartSpace>`

// TestReaderAxisTickMarkDefaultsOut: an axis without <c:majorTickMark> gets
// PowerPoint's "out" default; an explicit val="none" still wins (deck 2's
// bar/scatter charts declare none and the COM exports show no ticks).
func TestReaderAxisTickMarkDefaultsOut(t *testing.T) {
	chart := parseChartXML([]byte(tickDefaultChart), nil)
	if chart == nil {
		t.Fatal("parseChartXML returned nil")
	}
	cat := chart.plotArea.GetAxisX()
	val := chart.plotArea.GetAxisY()
	if cat == nil || val == nil {
		t.Fatalf("axes dropped: cat=%v val=%v", cat != nil, val != nil)
	}
	if cat.MajorTickMark != TickMarkOutside {
		t.Errorf("category axis majorTickMark = %q, want %q (element absent)", cat.MajorTickMark, TickMarkOutside)
	}
	if val.MajorTickMark != TickMarkOutside {
		t.Errorf("value axis majorTickMark = %q, want %q (element absent)", val.MajorTickMark, TickMarkOutside)
	}

	chart2 := parseChartXML([]byte(tickNoneChart), nil)
	if chart2 == nil {
		t.Fatal("parseChartXML returned nil for the none fixture")
	}
	if ax := chart2.plotArea.GetAxisX(); ax.MajorTickMark != TickMarkNone {
		t.Errorf("explicit none axis majorTickMark = %q, want none", ax.MajorTickMark)
	}
	if ax := chart2.plotArea.GetAxisY(); ax.MajorTickMark != TickMarkNone {
		t.Errorf("explicit none axis majorTickMark = %q, want none", ax.MajorTickMark)
	}
}

// TestAPITimeAxisKeepsNoneDefault: axes created through the API keep the
// tick-mark-free default — only file-parsed axes gain PowerPoint's "out".
func TestAPITimeAxisKeepsNoneDefault(t *testing.T) {
	ax := NewChartAxis()
	if ax.MajorTickMark != TickMarkNone {
		t.Errorf("API axis majorTickMark = %q, want none", ax.MajorTickMark)
	}
}

// TestLineSeriesDefaultStrokeIsTwoTwentyFive: a line series without an
// explicit width draws at 28575 EMU (2.25pt) — deck 2 slide05's unstyled
// series is ~5px wide at 1600px/10in, not the 1-2px hairline this drew.
func TestLineSeriesDefaultStrokeIsTwoTwentyFive(t *testing.T) {
	r := &renderer{scaleX: 1600.0 / 9144000.0}
	ser := &ChartSeries{}
	if got := r.chartSeriesLineWidth(ser, 2); got != 5 {
		t.Errorf("default series width = %d, want 5 (28575 EMU at 160dpi)", got)
	}
	// An explicit width still wins.
	ser.Outline = &SeriesOutline{Width: 1}
	if got := r.chartSeriesLineWidth(ser, 2); got != 2 {
		t.Errorf("explicit 1pt width = %d, want 2", got)
	}
}

// TestMarkerSizeIsPoints: <c:size val="8"/> is an 8pt diameter — the drawn
// diamond must span ~16px at 160dpi, not the 4px dot the unscaled radius
// produced. Rendered ink is measured because the radius lives inside.
func TestMarkerSizeIsPoints(t *testing.T) {
	r := &renderer{scaleX: 1600.0 / 9144000.0, scaleY: 1200.0 / 6858000.0}
	r.img = image.NewRGBA(image.Rect(0, 0, 64, 64))
	ser := &ChartSeries{Marker: &SeriesMarker{Symbol: MarkerDiamond, Size: 8}}
	blue := color.RGBA{R: 79, G: 129, B: 189, A: 255}
	r.drawChartMarker(ser, 32, 32, blue)

	span := func(pred func(color.RGBA) bool) int {
		minY, maxY := -1, -1
		for y := 0; y < 64; y++ {
			for x := 0; x < 64; x++ {
				cr, cg, cb, _ := r.img.At(x, y).RGBA()
				if pred(color.RGBA{R: uint8(cr >> 8), G: uint8(cg >> 8), B: uint8(cb >> 8), A: 255}) {
					if minY < 0 {
						minY = y
					}
					maxY = y
				}
			}
		}
		if minY < 0 {
			return 0
		}
		return maxY - minY + 1
	}
	h := span(func(c color.RGBA) bool { return c.B > 120 && c.B-c.R > 40 })
	if h < 12 {
		t.Errorf("8pt diamond ink height = %d, want >=12 (16px diameter minus AA); "+
			"an unscaled point count collapses it to ~4", h)
	}
}

// TestRotatedDateLabelOffsetPinned pins the rotated date label's ink-column
// shift: the renderer's offset must equal 2.25pt (5px at 160dpi) — the COM
// golds center the label column that far right of the tick.
func TestRotatedDateLabelOffsetPinned(t *testing.T) {
	r := &renderer{scaleX: 1600.0 / 9144000.0}
	if got := r.rotatedTickLabelOffset(); got != 5 {
		t.Errorf("rotated label ink offset = %d, want 5 (2.25pt at 160dpi)", got)
	}
	// A 4:3 deck rendered at 640px halves it.
	r2 := &renderer{scaleX: 640.0 / 9144000.0}
	if got := r2.rotatedTickLabelOffset(); got != 2 {
		t.Errorf("rotated label ink offset at 640px = %d, want 2", got)
	}
}

// TestChartValueLabelGapPinned pins the value-label clearance shared by the
// plot reservation and the label draw: one label line height + 4px. Break the
// helper and the labels detach from the axis (slide05's left band ghosts).
func TestChartValueLabelGapPinned(t *testing.T) {
	r := &renderer{scaleX: 1600.0 / 9144000.0}
	face := r.chartFace(&Font{Name: "Calibri", Size: 16}, "")
	if face == nil {
		t.Fatal("no face for the axis font")
	}
	want := r.chartLineHeight(face) + 4
	if got := r.chartValueLabelGap(face); got != want {
		t.Errorf("value label gap = %d, want %d (line height + 4)", got, want)
	}
}
