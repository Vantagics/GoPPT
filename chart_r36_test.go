package gopresentation

import (
	"math"
	"strings"
	"testing"
	"time"
)

// The r36 chart semantics tests.
//
// Deck 00022823 slides 05/06/29 (line and scatter charts) exposed five gaps
// at once: logarithmic value axes ticked linearly, a second value axis was
// always routed to axisY even for scatter charts, <c:dateAx> was not read at
// all, <a:prstDash> was dropped, and blank points past the numCache count
// were plotted as fake zeros hugging the axis floor. Each test pins one of
// those behaviours with PowerPoint-measured expectations.

const r36ChartHeader = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart>
    <c:autoTitleDeleted val="1"/>
    <c:plotArea>`

const r36ChartFooter = `    </c:plotArea>
  </c:chart>
</c:chartSpace>`

// logAxisValAx mirrors deck 00022823 chart5's value axis: base 2, min pinned
// to 128, data topping out at 8192.
const logAxisValAx = `
      <c:valAx>
        <c:axId val="2"/>
        <c:scaling><c:orientation val="minMax"/><c:logBase val="2"/><c:min val="128"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="l"/>
        <c:numFmt formatCode="General" sourceLinked="1"/>
        <c:crossAx val="1"/>
      </c:valAx>`

// TestChartReadsLogBase pins that <c:logBase> lands on the model axis.
func TestChartReadsLogBase(t *testing.T) {
	xml := r36ChartHeader + `
      <c:lineChart>
        <c:grouping val="standard"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:cat><c:strRef><c:f>Sheet1!$A$1:$A$3</c:f><c:strCache><c:ptCount val="3"/><c:pt idx="0"><c:v>a</c:v></c:pt><c:pt idx="1"><c:v>b</c:v></c:pt><c:pt idx="2"><c:v>c</c:v></c:pt></c:strCache></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$B$1:$B$3</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="3"/><c:pt idx="0"><c:v>128</c:v></c:pt><c:pt idx="1"><c:v>1024</c:v></c:pt><c:pt idx="2"><c:v>8192</c:v></c:pt></c:numCache></c:numRef></c:val>
        </c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:lineChart>
      <c:catAx>
        <c:axId val="1"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/><c:axPos val="b"/>
        <c:crossAx val="2"/>
      </c:catAx>` + logAxisValAx + r36ChartFooter

	chart := parseChartXML([]byte(xml), nil)
	if chart == nil {
		t.Fatal("parseChartXML returned nil")
	}
	axY := chart.plotArea.GetAxisY()
	if axY == nil {
		t.Fatal("value axis missing")
	}
	if axY.LogBase != 2 {
		t.Errorf("LogBase = %v, want 2", axY.LogBase)
	}
	if axY.MinBounds == nil || *axY.MinBounds != 128 {
		t.Errorf("MinBounds = %v, want 128", axY.MinBounds)
	}
}

// TestChartLogTicksAtPowersOfBase: chartComputeScale snaps the bounds
// outward to whole powers of the base and ticks() lands on each power.
// COM slide29 (base 2, min 128) labels 128, 256, ..., 8192; a base-10 axis
// with data max 129.9 tops out at 1000, not 200.
func TestChartLogTicksAtPowersOfBase(t *testing.T) {
	ax := &ChartAxis{LogBase: 2, Visible: true}
	cs := chartComputeScale(128, 8192, ax, false)
	want := []float64{128, 256, 512, 1024, 2048, 4096, 8192}
	got := cs.ticks()
	if len(got) != len(want) {
		t.Fatalf("ticks = %v, want %v", got, want)
	}
	for i := range want {
		if got[i] != want[i] {
			t.Errorf("tick[%d] = %v, want %v", i, got[i], want[i])
		}
	}

	ax10 := &ChartAxis{LogBase: 10, Visible: true}
	cs10 := chartComputeScale(10, 129.9, ax10, false)
	if cs10.max != 1000 {
		t.Errorf("base-10 max = %v, want 1000 (base^ceil(log10 129.9))", cs10.max)
	}
}

// TestChartValueAxisRoutingByPlotType pins the dual-value-axis rule: only
// scatter/bubble charts route an axPos="b"/"t" value axis to axisX. The
// first cut routed every "b" value axis to axisX, which threw deck 00022823
// slide23's horizontal bar chart into regression (its value axis also sits
// on "b" but must stay axisY).
func TestChartValueAxisRoutingByPlotType(t *testing.T) {
	barXML := r36ChartHeader + `
      <c:barChart>
        <c:barDir val="bar"/>
        <c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:cat><c:strRef><c:f>Sheet1!$A$1:$A$2</c:f><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>x</c:v></c:pt><c:pt idx="1"><c:v>y</c:v></c:pt></c:strCache></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$B$1:$B$2</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="2"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numCache></c:numRef></c:val>
        </c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:barChart>
      <c:catAx>
        <c:axId val="1"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/><c:axPos val="l"/>
        <c:crossAx val="2"/>
      </c:catAx>
      <c:valAx>
        <c:axId val="2"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/><c:axPos val="b"/>
        <c:crossAx val="1"/>
      </c:valAx>` + r36ChartFooter

	chart := parseChartXML([]byte(barXML), nil)
	if chart == nil {
		t.Fatal("bar chart: parseChartXML returned nil")
	}
	if chart.plotArea.GetAxisY() == nil {
		t.Error("bar chart: axPos=b value axis must stay on axisY")
	}

	scatterXML := r36ChartHeader + `
      <c:scatterChart>
        <c:scatterStyle val="lineMarker"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:xVal><c:numRef><c:f>Sheet1!$A$1:$A$2</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="2"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numCache></c:numRef></c:xVal>
          <c:yVal><c:numRef><c:f>Sheet1!$B$1:$B$2</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="2"/><c:pt idx="0"><c:v>3</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numCache></c:numRef></c:yVal>
        </c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:scatterChart>
      <c:valAx>
        <c:axId val="1"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/><c:axPos val="b"/>
        <c:crossAx val="2"/>
      </c:valAx>
      <c:valAx>
        <c:axId val="2"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/><c:axPos val="l"/>
        <c:crossAx val="1"/>
      </c:valAx>` + r36ChartFooter

	schart := parseChartXML([]byte(scatterXML), nil)
	if schart == nil {
		t.Fatal("scatter chart: parseChartXML returned nil")
	}
	// NewPlotArea pre-seeds both axis slots, so the routing must be checked
	// by identity: the bottom-edge value axis lands in axisX and the left
	// one in axisY. With the routing disabled both valAx commits pile into
	// axisY (last wins) and axisX stays the untouched default.
	if axX := schart.plotArea.GetAxisX(); axX == nil || axX.Position != "b" {
		got := "nil"
		if axX != nil {
			got = axX.Position
		}
		t.Errorf("scatter chart: axisX.Position = %q, want \"b\" (bottom value axis must route to axisX)", got)
	}
	if axY := schart.plotArea.GetAxisY(); axY == nil || axY.Position != "l" {
		got := "nil"
		if axY != nil {
			got = axY.Position
		}
		t.Errorf("scatter chart: axisY.Position = %q, want \"l\"", got)
	}
}

// TestChartDateSerialIs1900Epoch: serial 37164 is 2001-09-30 over the
// 1900 workbook epoch — deck 00022823 chart1 declares <c:date1904/> yet COM
// renders its serial 37164 as "Sep 2001", which only the 1900 system gives.
func TestChartDateSerialIs1900Epoch(t *testing.T) {
	tm := chartSerialToTime(37164)
	if tm.Year() != 2001 || tm.Month() != time.September || tm.Day() != 30 {
		t.Errorf("chartSerialToTime(37164) = %v, want 2001-09-30", tm)
	}
	if got := chartFormatDate(37164, "mmm yyyy"); got != "Sep 2001" {
		t.Errorf("chartFormatDate(37164, mmm yyyy) = %q, want %q", got, "Sep 2001")
	}
}

// TestChartMonthTicksClampDay pins the day-clamp on calendar-month ticks:
// Sep 30 + 5 months must land on Feb 28 (PowerPoint), not Mar 2 (Go's
// AddDate normalisation). COM slide05's tick positions require it.
func TestChartMonthTicksClampDay(t *testing.T) {
	cs := chartScale{
		min:        chartTimeToSerial(time.Date(2001, 9, 30, 0, 0, 0, 0, time.UTC)),
		max:        chartTimeToSerial(time.Date(2003, 9, 30, 0, 0, 0, 0, time.UTC)),
		stepMonths: 5,
	}
	ticks := cs.ticks()
	if len(ticks) < 3 {
		t.Fatalf("ticks = %v, want at least 3", ticks)
	}
	wantFeb := chartTimeToSerial(time.Date(2002, 2, 28, 0, 0, 0, 0, time.UTC))
	if ticks[1] != wantFeb {
		t.Errorf("tick[1] = %v (%s), want Feb 28 2002 — day overflow must clamp to month end",
			ticks[1], chartSerialToTime(ticks[1]))
	}
	// Sep 30 + 10 months preserves the day (Jul 30, Excel EDATE semantics);
	// the clamp only fires when the target month is shorter than the day.
	wantJul := chartTimeToSerial(time.Date(2002, 7, 30, 0, 0, 0, 0, time.UTC))
	if ticks[2] != wantJul {
		t.Errorf("tick[2] = %v (%s), want Jul 30 2002",
			ticks[2], chartSerialToTime(ticks[2]))
	}
}

// TestChartSeriesDashRoundTrip: the reader must lift <a:prstDash> out of the
// series <a:ln> into LineDash, and the writer must emit it back. The first
// reader cut keyed the guard on pendTarget, which the colour element's
// EndElement had already cleared, so the dash was silently dropped.
func TestChartSeriesDashRoundTrip(t *testing.T) {
	xml := r36ChartHeader + `
      <c:lineChart>
        <c:grouping val="standard"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:spPr><a:ln w="19050"><a:solidFill><a:srgbClr val="808080"/></a:solidFill><a:prstDash val="sysDash"/></a:ln></c:spPr>
          <c:cat><c:strRef><c:f>Sheet1!$A$1:$A$2</c:f><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>a</c:v></c:pt><c:pt idx="1"><c:v>b</c:v></c:pt></c:strCache></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$B$1:$B$2</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="2"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numCache></c:numRef></c:val>
        </c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:lineChart>
      <c:catAx>
        <c:axId val="1"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/><c:axPos val="b"/>
        <c:crossAx val="2"/>
      </c:catAx>
      <c:valAx>
        <c:axId val="2"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/><c:axPos val="l"/>
        <c:crossAx val="1"/>
      </c:valAx>` + r36ChartFooter

	chart := parseChartXML([]byte(xml), nil)
	if chart == nil {
		t.Fatal("parseChartXML returned nil")
	}
	lc, ok := chart.plotArea.GetType().(*LineChart)
	if !ok {
		t.Fatalf("plot type = %T, want *LineChart", chart.plotArea.GetType())
	}
	if len(lc.Series) != 1 {
		t.Fatalf("series count = %d, want 1", len(lc.Series))
	}
	if got := lc.Series[0].LineDash; got != "sysDash" {
		t.Errorf("LineDash = %q, want %q", got, "sysDash")
	}

	// Writer side: the dash preset must come back out inside <a:ln>, after
	// a well-formed solidFill (the first cut dropped </a:solidFill> — the
	// fidelity suite caught it, this assertion keeps it honest here too).
	out := lineElementXML(NewColor("808080"), 2, "sysDash")
	wantOut := `<a:ln w="25400"><a:solidFill><a:srgbClr val="808080"/></a:solidFill><a:prstDash val="sysDash"/></a:ln>`
	if out != wantOut {
		t.Errorf("lineElementXML =\n%s\nwant\n%s", out, wantOut)
	}
	if out := lineElementXML(NewColor("808080"), 2, ""); strings.Contains(out, "prstDash") {
		t.Errorf("solid line must not carry prstDash: %s", out)
	}
}

// TestChartMissingPointsAreNaN: when numCache's ptCount exceeds the number
// of <c:pt> elements, the blanks must read as NaN so the renderer breaks the
// line there — chart1's "Cost per Genome" ends mid-axis in COM rather than
// drawing a fake zero line along the floor.
func TestChartMissingPointsAreNaN(t *testing.T) {
	xml := r36ChartHeader + `
      <c:lineChart>
        <c:grouping val="standard"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:cat><c:strRef><c:f>Sheet1!$A$1:$A$4</c:f><c:strCache><c:ptCount val="4"/><c:pt idx="0"><c:v>a</c:v></c:pt><c:pt idx="1"><c:v>b</c:v></c:pt><c:pt idx="2"><c:v>c</c:v></c:pt><c:pt idx="3"><c:v>d</c:v></c:pt></c:strCache></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$B$1:$B$4</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="4"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt></c:numCache></c:numRef></c:val>
        </c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:lineChart>
      <c:catAx>
        <c:axId val="1"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/><c:axPos val="b"/>
        <c:crossAx val="2"/>
      </c:catAx>
      <c:valAx>
        <c:axId val="2"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/><c:axPos val="l"/>
        <c:crossAx val="1"/>
      </c:valAx>` + r36ChartFooter

	chart := parseChartXML([]byte(xml), nil)
	if chart == nil {
		t.Fatal("parseChartXML returned nil")
	}
	lc := chart.plotArea.GetType().(*LineChart)
	s := lc.Series[0]
	if got := s.Values["d"]; !math.IsNaN(got) {
		t.Errorf("Values[d] = %v, want NaN (missing numCache point)", got)
	}
	if got := s.Values["c"]; got != 3 {
		t.Errorf("Values[c] = %v, want 3", got)
	}
	// The series range must ignore the NaN, not treat it as zero/minimum.
	lo, hi := chartSeriesRange([]*ChartSeries{s})
	if lo != 1 || hi != 3 {
		t.Errorf("chartSeriesRange = (%v, %v), want (1, 3)", lo, hi)
	}
}

// TestChartIsDateFormat: the date-vs-number label branch keys on mmm/mmmm/yy
// tokens (the codes PowerPoint actually puts on date axes).
func TestChartIsDateFormat(t *testing.T) {
	cases := map[string]bool{
		"mmm yyyy":   true,
		"mmmm-yy":    true,
		`mmm\ yyyy`:  true,
		"#,##0":      false,
		"General":    false,
		"0.0":        false,
		`"$"#,##0.0`: false,
	}
	for format, want := range cases {
		if got := chartIsDateFormat(format); got != want {
			t.Errorf("chartIsDateFormat(%q) = %v, want %v", format, got, want)
		}
	}
}
