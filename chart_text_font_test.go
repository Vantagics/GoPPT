package gopresentation

import (
	"image"
	"strings"
	"testing"
)

// Chart text fonts, end to end.
//
// A chart's strings carry the same font model as slide text — a Latin face and
// an East Asian face — but all three layers of the chart path dropped it: the
// writer emitted only the size and bold flag, the reader never looked at
// <a:latin>/<a:ea>, and the rasteriser resolved chart text through the Latin
// path with no glyph-coverage check. The symptom was Chinese chart labels drawn
// as .notdef boxes while every layer reported success: the declared name
// resolves, so FontDiagnostics has nothing to complain about, and the chart
// goldens use ASCII strings, so they compared tofu against tofu quite happily.
// Each layer therefore gets an assertion here.
//
// Fixing the run fixed the preview and left the *file* wrong. A chart states a
// title's font on the run inside <c:title> but every other label's font —
// axis tick labels, legend entries, data labels — in a <c:txPr> belonging to
// the element, and the writer emitted no <c:txPr> at all. The preview was
// right because the rasteriser resolves the model's font itself; PowerPoint,
// which reads the part, fell back to the theme font for every one of them. So
// the tests below cover both the run and the <c:txPr>, and pin the slot each
// one has to occupy.

// fixtureChartWithFonts builds a bar chart that declares a font on every text
// element a chart can have: the chart title, both axis titles, both axes' tick
// labels, the legend and the data labels. One write therefore exercises every
// place the chart writer emits a font, and one round trip every place the
// reader reads one back.
func fixtureChartWithFonts() *Presentation {
	p := New()
	chart := p.GetActiveSlide().CreateChartShape()
	chart.BaseShape.SetOffsetX(500000).SetOffsetY(500000)
	chart.BaseShape.SetWidth(7000000).SetHeight(4500000)

	title := chart.GetTitle()
	title.SetText("收入趋势").SetVisible(true)
	title.Font.SetName("Microsoft YaHei")
	title.Font.SetSize(18).SetBold(true)
	title.Font.NameEA = "SimSun"

	bar := NewBarChart()
	series := NewChartSeriesOrdered("S", []string{"Q1", "Q2"}, []float64{10, 20})
	series.ShowValue = true
	series.LabelPosition = LabelOutsideEnd
	series.Font.SetName("Microsoft YaHei")
	series.Font.NameEA = "SimSun"
	bar.AddSeries(series)
	chart.GetPlotArea().SetType(bar)

	legend := chart.GetLegend()
	legend.Visible = true
	legend.Font.SetName("Microsoft YaHei")
	legend.Font.NameEA = "SimSun"

	for _, ax := range []*ChartAxis{chart.GetPlotArea().GetAxisX(), chart.GetPlotArea().GetAxisY()} {
		ax.SetTitle("季度")
		ax.Font.SetName("Microsoft YaHei")
		ax.Font.NameEA = "SimSun"
	}
	// Bounds and a major unit on the value axis pin the order of the elements
	// the axis slot test below checks: CT_Scaling puts max before min, and
	// CT_ValAx puts the units after the crossing pair.
	axY := chart.GetPlotArea().GetAxisY()
	axY.SetMinBounds(0)
	axY.SetMaxBounds(40)
	axY.SetMajorUnit(5)
	return p
}

// chartPartOf returns the chart part of a one-chart presentation as text.
func chartPartOf(t *testing.T, p *Presentation) string {
	t.Helper()
	part, ok := zipParts(t, writeToBytes(t, p))["ppt/charts/chart1.xml"]
	if !ok {
		t.Fatal("chart part was not written")
	}
	return string(part)
}

// blockIn returns the slice of text between an opening and a closing tag. The
// parts are small and hand-formatted, which makes a text slice enough to say
// "this element carries that".
func blockIn(t *testing.T, text, open, close string) string {
	t.Helper()
	i := strings.Index(text, open)
	if i < 0 {
		t.Fatalf("%s is not in the part; got:\n%s", open, text)
	}
	j := strings.Index(text[i:], close)
	if j < 0 {
		t.Fatalf("%s is never closed by %s; got:\n%s", open, close, text)
	}
	return text[i : i+j+len(close)]
}

// assertOrdered fails when the given tags do not appear in the block in that
// order.
func assertOrdered(t *testing.T, what, block string, tags ...string) {
	t.Helper()
	at := -1
	for _, tag := range tags {
		i := strings.Index(block, tag)
		if i < 0 {
			t.Errorf("%s: %s is missing; got:\n%s", what, tag, block)
			return
		}
		if i < at {
			t.Errorf("%s: %s is out of schema order (want %s); got:\n%s",
				what, tag, strings.Join(tags, " -> "), block)
			return
		}
		at = i
	}
}

// TestChartTextFontReachesEveryTextElement covers the writer half. The chart
// writer knew each element's font — the model carries it — and serialised a
// size and a bold flag on the title runs only, so the rest was gone the next
// time the part was opened.
//
// Note what replaced the old assertion here: it counted <a:ea> occurrences and
// expected three, which is exactly what the writer emitted, so it was derived
// from the code under test — and the two label kinds that carried no font at
// all could therefore not fail it. The elements are named individually now.
func TestChartTextFontReachesEveryTextElement(t *testing.T) {
	text := chartPartOf(t, fixtureChartWithFonts())

	title := blockIn(t, text, "<c:title>", "</c:title>")
	if !strings.Contains(title,
		`<a:rPr lang="en-US" sz="1800" b="1" i="0"><a:latin typeface="Microsoft YaHei"/><a:ea typeface="SimSun"/></a:rPr>`) {
		t.Errorf("the chart title run does not carry its size, bold flag and both typefaces together; got:\n%s", title)
	}

	// The const strings are what PowerPoint reads a label's font from.
	const labelFont = `<a:defRPr lang="en-US" sz="1000" b="0" i="0"><a:latin typeface="Microsoft YaHei"/><a:ea typeface="SimSun"/></a:defRPr>`
	for _, el := range []struct{ what, open, close string }{
		{"the category axis", "<c:catAx>", "</c:catAx>"},
		{"the value axis", "<c:valAx>", "</c:valAx>"},
		{"the legend", "<c:legend>", "</c:legend>"},
		{"the data labels", "<c:dLbls>", "</c:dLbls>"},
	} {
		block := blockIn(t, text, el.open, el.close)
		if !strings.Contains(block, labelFont) {
			t.Errorf("%s carries no <c:txPr> naming the font PowerPoint would draw it with; got:\n%s",
				el.what, block)
		}
	}
}

// TestChartLabelFontSitsInTheSchemaSlot pins where <c:txPr> has to appear.
//
// CT_CatAx, CT_ValAx and CT_DLbls all validate the order of their children, and
// our reader matches on the local name without looking at order at all — so a
// misplaced element round-trips here and is only ever rejected by PowerPoint.
// That asymmetry is how the chart writer's order drifted in the first place.
// The sequences below are the ones in dml-chart.xsd.
func TestChartLabelFontSitsInTheSchemaSlot(t *testing.T) {
	text := chartPartOf(t, fixtureChartWithFonts())

	assertOrdered(t, "category axis",
		blockIn(t, text, "<c:catAx>", "</c:catAx>"),
		`<c:tickLblPos `, `<c:txPr>`, `<c:crossAx `, `<c:crosses `)

	assertOrdered(t, "value axis",
		blockIn(t, text, "<c:valAx>", "</c:valAx>"),
		// CT_ValAx splits the units out after the crossing pair.
		`<c:tickLblPos `, `<c:txPr>`, `<c:crossAx `, `<c:crosses `, `<c:majorUnit `)

	assertOrdered(t, "value axis scaling",
		blockIn(t, blockIn(t, text, "<c:valAx>", "</c:valAx>"), "<c:scaling>", "</c:scaling>"),
		`<c:orientation `, `<c:max `, `<c:min `)

	assertOrdered(t, "data labels",
		blockIn(t, text, "<c:dLbls>", "</c:dLbls>"),
		`<c:txPr>`, `<c:dLblPos `, `<c:showVal `)

	assertOrdered(t, "legend",
		blockIn(t, text, "<c:legend>", "</c:legend>"),
		`<c:overlay `, `<c:txPr>`)
}

// chartTextFontXML is a minimal bar chart part whose title runs declare a Latin
// and an East Asian typeface, as PowerPoint writes them.
const chartTextFontXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart>
    <c:title>
      <c:tx>
        <c:rich>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:r>
              <a:rPr lang="en-US" sz="1400" b="0"><a:latin typeface="Microsoft YaHei"/><a:ea typeface="SimSun"/></a:rPr>
              <a:t>收入趋势</a:t>
            </a:r>
          </a:p>
        </c:rich>
      </c:tx>
      <c:overlay val="0"/>
    </c:title>
    <c:plotArea>
      <c:layout/>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:varyColors val="0"/>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>S</c:v></c:pt></c:strCache></c:strRef></c:tx>
          <c:cat><c:strRef><c:f>Sheet1!$A$2</c:f><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt></c:strCache></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$B$2</c:f><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt></c:numCache></c:numRef></c:val>
        </c:ser>
        <c:axId val="1"/>
        <c:axId val="2"/>
      </c:barChart>
      <c:catAx>
        <c:axId val="1"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="b"/>
        <c:crossAx val="2"/>
        <c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"><a:latin typeface="Microsoft YaHei"/><a:ea typeface="SimSun"/></a:rPr><a:t>季度</a:t></a:r></a:p></c:rich></c:tx></c:title>
      </c:catAx>
      <c:valAx>
        <c:axId val="2"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="l"/>
        <c:crossAx val="1"/>
      </c:valAx>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`

// TestChartTextFontIsReadFromTypeface covers the reader half. Both typefaces
// live on the run's <a:rPr>, and reading the text without them left every chart
// string on the model default.
func TestChartTextFontIsReadFromTypeface(t *testing.T) {
	chart := parseChartXML([]byte(chartTextFontXML), nil)
	if chart == nil {
		t.Fatal("parseChartXML returned nil for a valid chart part")
	}

	title := chart.GetTitle()
	if title.Font == nil {
		t.Fatal("chart title has no font after parsing")
	}
	if got, want := title.Font.Name, "Microsoft YaHei"; got != want {
		t.Errorf("chart title font = %q, want %q", got, want)
	}
	if got, want := title.Font.NameEA, "SimSun"; got != want {
		t.Errorf("chart title East Asian font = %q, want %q", got, want)
	}

	axX := chart.GetPlotArea().GetAxisX()
	if axX == nil {
		t.Fatal("category axis was not read")
	}
	if got, want := axX.Title, "季度"; got != want {
		t.Errorf("category axis title = %q, want %q", got, want)
	}
	if got, want := axX.Font.Name, "Microsoft YaHei"; got != want {
		t.Errorf("category axis font = %q, want %q", got, want)
	}
	if got, want := axX.Font.NameEA, "SimSun"; got != want {
		t.Errorf("category axis East Asian font = %q, want %q", got, want)
	}
}

// TestChartTextFontSurvivesRoundTrip is the integration case: a chart written
// with a font must read back with it. Either half alone leaves this failing, and
// nothing else in the suite notices, because a chart that has lost its font
// still renders — just not with the right face.
func TestChartTextFontSurvivesRoundTrip(t *testing.T) {
	rt := firstChartShape(t, roundTrip(t, fixtureChartWithFonts()))

	title := rt.GetTitle()
	if title.Font == nil {
		t.Fatal("chart title has no font after the round trip")
	}
	if got, want := title.Text, "收入趋势"; got != want {
		t.Errorf("chart title text = %q, want %q", got, want)
	}
	if got, want := title.Font.Name, "Microsoft YaHei"; got != want {
		t.Errorf("chart title font = %q, want %q", got, want)
	}
	if got, want := title.Font.NameEA, "SimSun"; got != want {
		t.Errorf("chart title East Asian font = %q, want %q", got, want)
	}

	axX := rt.GetPlotArea().GetAxisX()
	if axX == nil {
		t.Fatal("category axis was lost by the round trip")
	}
	if got, want := axX.Title, "季度"; got != want {
		t.Errorf("category axis title = %q, want %q", got, want)
	}
	if got, want := axX.Font.Name, "Microsoft YaHei"; got != want {
		t.Errorf("category axis font = %q, want %q", got, want)
	}
	if got, want := axX.Font.NameEA, "SimSun"; got != want {
		t.Errorf("category axis East Asian font = %q, want %q", got, want)
	}
}

// chartLabelFontXML is a bar chart written the way PowerPoint writes one: no
// font on any title run, and a <c:txPr> on each label-bearing element instead.
// The legend's entry names the theme's minor font, which is a reference rather
// than a face.
const chartLabelFontXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart>
    <c:plotArea>
      <c:layout/>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:varyColors val="0"/>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>S</c:v></c:pt></c:strCache></c:strRef></c:tx>
          <c:dLbls>
            <c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900" b="1"><a:latin typeface="Microsoft YaHei"/><a:ea typeface="SimSun"/></a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>
            <c:showVal val="1"/>
          </c:dLbls>
          <c:cat><c:strRef><c:f>Sheet1!$A$2</c:f><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt></c:strCache></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$B$2</c:f><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt></c:numCache></c:numRef></c:val>
        </c:ser>
        <c:axId val="1"/>
        <c:axId val="2"/>
      </c:barChart>
      <c:catAx>
        <c:axId val="1"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="b"/>
        <c:tickLblPos val="nextTo"/>
        <c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1200" b="0"><a:latin typeface="Microsoft YaHei"/><a:ea typeface="SimSun"/></a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>
        <c:crossAx val="2"/>
        <c:crosses val="autoZero"/>
      </c:catAx>
      <c:valAx>
        <c:axId val="2"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="l"/>
        <c:tickLblPos val="nextTo"/>
        <c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1200" b="1"><a:latin typeface="Microsoft YaHei"/><a:ea typeface="SimSun"/></a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>
        <c:crossAx val="1"/>
        <c:crosses val="autoZero"/>
      </c:valAx>
    </c:plotArea>
    <c:legend>
      <c:legendPos val="b"/>
      <c:overlay val="0"/>
      <c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1000" b="0"><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/></a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>
    </c:legend>
    <c:plotVisOnly val="1"/>
    <c:dispBlanksAs val="zero"/>
  </c:chart>
</c:chartSpace>`

// TestChartLabelFontIsReadFromTxPr covers the reader half of the label fonts.
// PowerPoint takes an axis', a legend's and a data label's font from <c:txPr>,
// and the reader used to read the typefaces of title runs only — so every label
// came back on the model default even though the part named a face.
func TestChartLabelFontIsReadFromTxPr(t *testing.T) {
	chart := parseChartXML([]byte(chartLabelFontXML), nil)
	if chart == nil {
		t.Fatal("parseChartXML returned nil for a valid chart part")
	}

	axX := chart.GetPlotArea().GetAxisX()
	if axX == nil {
		t.Fatal("category axis was not read")
	}
	if got, want := axX.Font.Name, "Microsoft YaHei"; got != want {
		t.Errorf("category axis font = %q, want %q (from its <c:txPr>)", got, want)
	}
	if got, want := axX.Font.NameEA, "SimSun"; got != want {
		t.Errorf("category axis East Asian font = %q, want %q", got, want)
	}
	if got, want := axX.Font.Size, 12; got != want {
		t.Errorf("category axis font size = %d, want %d (sz=\"1200\")", got, want)
	}

	axY := chart.GetPlotArea().GetAxisY()
	if axY == nil {
		t.Fatal("value axis was not read")
	}
	if got, want := axY.Font.Name, "Microsoft YaHei"; got != want {
		t.Errorf("value axis font = %q, want %q", got, want)
	}
	if !axY.Font.Bold {
		t.Error("value axis font is not bold, but its <c:txPr> states b=\"1\"")
	}

	bar, ok := chart.GetPlotArea().GetType().(*BarChart)
	if !ok {
		t.Fatalf("chart type = %T, want *BarChart", chart.GetPlotArea().GetType())
	}
	if len(bar.Series) != 1 {
		t.Fatalf("read %d series, want 1", len(bar.Series))
	}
	if got, want := bar.Series[0].Font.NameEA, "SimSun"; got != want {
		t.Errorf("data-label East Asian font = %q, want %q (from the series' <c:txPr>)", got, want)
	}
	if got, want := bar.Series[0].Font.Size, 9; got != want {
		t.Errorf("data-label font size = %d, want %d (sz=\"900\")", got, want)
	}
	if !bar.Series[0].Font.Bold {
		t.Error("data-label font is not bold, but its <c:txPr> states b=\"1\"")
	}
}

// TestThemeFontReferenceIsNotAFontName: PowerPoint writes the theme's minor
// font into a <c:txPr> as "+mn-lt". That is a reference, not a name, so the
// model has to keep its own default rather than gain a face called "+mn-lt" —
// which no font cache can resolve, and which would then be reported as a
// substitution on every render.
func TestThemeFontReferenceIsNotAFontName(t *testing.T) {
	chart := parseChartXML([]byte(chartLabelFontXML), nil)
	if chart == nil {
		t.Fatal("parseChartXML returned nil for a valid chart part")
	}
	if got, want := chart.GetLegend().Font.Name, NewFont().Name; got != want {
		t.Errorf("legend font = %q, want the model default %q: \"+mn-lt\" is a theme reference, not a font name",
			got, want)
	}
}

// TestChartLabelFontSurvivesRoundTrip is the integration case for the label
// fonts: a chart whose axes, legend and data labels all name a font must read
// back with them. Either half alone leaves this failing, and nothing else in
// the suite notices, because a chart that has lost the font still renders —
// just not with the right face.
func TestChartLabelFontSurvivesRoundTrip(t *testing.T) {
	rt := firstChartShape(t, roundTrip(t, fixtureChartWithFonts()))

	for _, ax := range []struct {
		what string
		axis *ChartAxis
	}{
		{"category axis", rt.GetPlotArea().GetAxisX()},
		{"value axis", rt.GetPlotArea().GetAxisY()},
	} {
		if ax.axis == nil {
			t.Fatalf("%s was lost by the round trip", ax.what)
		}
		if ax.axis.Font == nil {
			t.Fatalf("%s has no font after the round trip", ax.what)
		}
		if got, want := ax.axis.Font.NameEA, "SimSun"; got != want {
			t.Errorf("%s East Asian font = %q, want %q", ax.what, got, want)
		}
	}

	legend := rt.GetLegend()
	if legend == nil || legend.Font == nil {
		t.Fatal("the legend font was lost by the round trip")
	}
	if got, want := legend.Font.NameEA, "SimSun"; got != want {
		t.Errorf("legend East Asian font = %q, want %q", got, want)
	}

	bar, ok := rt.GetPlotArea().GetType().(*BarChart)
	if !ok {
		t.Fatalf("chart type = %T, want *BarChart", rt.GetPlotArea().GetType())
	}
	if len(bar.Series) == 0 {
		t.Fatal("the series was lost by the round trip")
	}
	if bar.Series[0].Font == nil {
		t.Fatal("the series has no font after the round trip")
	}
	if got, want := bar.Series[0].Font.NameEA, "SimSun"; got != want {
		t.Errorf("data-label East Asian font = %q, want %q", got, want)
	}
}

// TestAxisTickMarksAndMinorGridlinesAreWritten covers three settings that were
// carried by the model and parsed by the reader, and written by nobody: the two
// tick-mark styles and a category axis' minor gridlines. All three reverted to
// the model default the next time the part was saved.
func TestAxisTickMarksAndMinorGridlinesAreWritten(t *testing.T) {
	p := fixtureChartWithFonts()
	axX := firstChartShape(t, p).GetPlotArea().GetAxisX()
	axX.SetMajorTickMark(TickMarkOutside)
	axX.SetMinorTickMark(TickMarkInside)
	axX.SetMinorGridlines(&Gridlines{Width: 1, Color: ColorBlack})

	cat := blockIn(t, chartPartOf(t, p), "<c:catAx>", "</c:catAx>")
	for _, want := range []string{
		`<c:majorTickMark val="out"/>`,
		`<c:minorTickMark val="in"/>`,
		"<c:minorGridlines>",
	} {
		if !strings.Contains(cat, want) {
			t.Errorf("the category axis does not carry %s; got:\n%s", want, cat)
		}
	}

	got := firstChartShape(t, roundTrip(t, p)).GetPlotArea().GetAxisX()
	if got.MajorTickMark != TickMarkOutside {
		t.Errorf("major tick mark after the round trip = %q, want %q", got.MajorTickMark, TickMarkOutside)
	}
	if got.MinorTickMark != TickMarkInside {
		t.Errorf("minor tick mark after the round trip = %q, want %q", got.MinorTickMark, TickMarkInside)
	}
	if got.MinorGridlines == nil {
		t.Error("the category axis lost its minor gridlines in the round trip")
	}
}

// TestChartTitleCJKRendersWithoutTofu covers the rasteriser half.
//
// Two charts are built that are identical in every respect except the three
// characters of their title. Titles that are drawn with a face lacking those
// glyphs come out as three identical .notdef boxes each, so the two slides are
// pixel-identical; titles drawn with a face that has the glyphs differ. The
// test therefore needs no reference image and cannot be satisfied by a
// uniformly wrong render.
func TestChartTitleCJKRendersWithoutTofu(t *testing.T) {
	const textA, textB = "本季度", "一二三"
	sample := textA + textB

	fc := NewFontCache()
	latin := latinOnlyFont(fc, sample)
	if latin == "" {
		t.Skip("no installed Latin-only font; the tofu comparison needs one")
	}
	if coverageFont(fc, sample) == "" {
		t.Skip("no installed font covers the sample; a tofu comparison would be meaningless")
	}

	build := func(title string) *Presentation {
		p := New()
		chart := p.GetActiveSlide().CreateChartShape()
		chart.BaseShape.SetOffsetX(300000).SetOffsetY(300000)
		chart.BaseShape.SetWidth(8500000).SetHeight(5000000)
		chart.GetTitle().SetText(title).SetVisible(true)
		// A face that is installed, resolves, and cannot draw a single one of
		// the characters: exactly the case that used to render as tofu while
		// reporting a perfect font match.
		chart.GetTitle().Font.SetName(latin)
		chart.GetLegend().Visible = false

		bar := NewBarChart()
		bar.AddSeries(NewChartSeriesOrdered("S", []string{"Q1", "Q2"}, []float64{10, 20}))
		chart.GetPlotArea().SetType(bar)
		return p
	}

	render := func(p *Presentation) image.Image {
		t.Helper()
		opts := DefaultRenderOptions()
		opts.Width = 1600
		opts.FontCache = fc
		img, err := p.SlideToImage(0, opts)
		if err != nil {
			t.Fatalf("render: %v", err)
		}
		return img
	}

	presA, presB := build(textA), build(textB)
	imgA, imgB := render(presA), render(presB)

	// The title strip. Everything below it is identical between the two
	// slides, so including some of it costs nothing and keeps the band from
	// depending on font metrics.
	strip := emuRect(t, presA, 300000, 300000, 8500000, 5000000, 1600)
	strip.Max.Y = strip.Min.Y + 200

	inStrip := countDiffInRect(imgA, imgB, strip)
	if inStrip == 0 {
		t.Errorf("two different three-character titles (%q vs %q, both declared as %q) rendered identically "+
			"inside the title strip: the chart title is drawn as missing-glyph boxes", textA, textB, latin)
	}

	if whole := countDiffInRect(imgA, imgB, imgA.Bounds()); whole != inStrip {
		t.Errorf("%d of %d differing pixels lie outside the title strip; the comparison is measuring "+
			"something other than the title", whole-inStrip, whole)
	}
}

// countDiffInRect counts pixels inside rect whose channels differ by more than
// the golden tolerance.
func countDiffInRect(a, b image.Image, rect image.Rectangle) int {
	n := 0
	for y := rect.Min.Y; y < rect.Max.Y; y++ {
		for x := rect.Min.X; x < rect.Max.X; x++ {
			ar, ag, ab, aa := a.At(x, y).RGBA()
			br, bg, bb, ba := b.At(x, y).RGBA()
			d := channelDelta(ar>>8, br>>8)
			if v := channelDelta(ag>>8, bg>>8); v > d {
				d = v
			}
			if v := channelDelta(ab>>8, bb>>8); v > d {
				d = v
			}
			if v := channelDelta(aa>>8, ba>>8); v > d {
				d = v
			}
			if d > goldenChannelTolerance {
				n++
			}
		}
	}
	return n
}

// latinOnlyFont returns an installed font that cannot draw any character of
// sample, or "" when the machine has none.
//
// The precondition asks the font cache directly and never touches the chart
// font-selection code, so reverting that code makes the test FAIL rather than
// SKIP and hide the regression.
func latinOnlyFont(fc *FontCache, sample string) string {
	candidates := []string{"Arial", "Calibri", "Segoe UI", "Tahoma", "Verdana", "DejaVu Sans", "Helvetica"}
	for _, name := range candidates {
		if !fc.HasFont(name, false, false) {
			continue
		}
		drawable := false
		for _, r := range sample {
			if fc.CoversRune(name, false, false, r) {
				drawable = true
				break
			}
		}
		if !drawable {
			return name
		}
	}
	return ""
}

// coverageFont returns the name of an installed font in the built-in East Asian
// fallback chain that can draw every character of sample, or "" when none can.
// Like latinOnlyFont it works strictly off the font cache.
func coverageFont(fc *FontCache, sample string) string {
	for _, name := range defaultCJKFallbackChain {
		if !fc.HasFont(name, false, false) {
			continue
		}
		all := true
		for _, r := range sample {
			if !fc.CoversRune(name, false, false, r) {
				all = false
				break
			}
		}
		if all {
			return name
		}
	}
	return ""
}
