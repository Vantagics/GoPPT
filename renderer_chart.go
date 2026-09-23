package gopresentation

import (
	"fmt"
	"image"
	"image/color"
	"math"
	"strconv"
	"strings"
	"time"

	"golang.org/x/image/font"
	"golang.org/x/image/math/fixed"
)

// ---------------------------------------------------------------------------
// Chart rasterization
//
// Charts are drawn from the ChartShape model. The layout follows PowerPoint's
// defaults: a title strip at the top, a legend along the chosen edge, a value
// axis with gridlines and tick labels, a category axis with labels, and the
// series plotted inside the remaining data area. Chart area fill/outline come
// from the model (nil means no fill and no border, as in PowerPoint).
// ---------------------------------------------------------------------------

// chartScale is the resolved value-axis scale.
type chartScale struct {
	min  float64
	max  float64
	step float64
	// logBase > 1 turns every mapping into logarithmic space
	// (<c:logBase>: tick marks sit at whole powers of the base).
	logBase float64
	// stepMonths > 0 marks a date-axis scale ticking every N calendar
	// months from the axis minimum (COM slide05: 3318-day span ticks every
	// 5 months — Sep 2001, Feb 2002, Jul 2002, ... — 22 labels).
	stepMonths int
}

func (cs chartScale) span() float64 {
	if cs.max <= cs.min {
		return 1
	}
	return cs.max - cs.min
}

// ratio maps a data value to a 0..1 position along the value axis.
func (cs chartScale) ratio(v float64) float64 {
	if cs.logBase > 1 && v > 0 && cs.min > 0 && cs.max > cs.min {
		return math.Log(v/cs.min) / math.Log(cs.max/cs.min)
	}
	return (v - cs.min) / cs.span()
}

// ratioClamped is ratio() clamped to the 0..1 range.
func (cs chartScale) ratioClamped(v float64) float64 {
	r := cs.ratio(v)
	if r < 0 {
		return 0
	}
	if r > 1 {
		return 1
	}
	return r
}

// ticks returns the major tick values between min and max.
func (cs chartScale) ticks() []float64 {
	if cs.stepMonths > 0 {
		start := chartSerialToTime(cs.min)
		var out []float64
		for k := 0; ; k++ {
			t := start.AddDate(0, k*cs.stepMonths, 0)
			// Go normalises day overflow into the next month (Sep 30 +
			// 5 months → Mar 2); PowerPoint clamps to the month's end
			// (Feb 28), which is what the COM tick positions show.
			if t.Day() < start.Day() {
				t = t.AddDate(0, 0, -t.Day())
			}
			s := chartTimeToSerial(t)
			if s > cs.max+0.5 {
				break
			}
			out = append(out, s)
			if len(out) > 64 {
				break
			}
		}
		if len(out) == 0 {
			out = append(out, cs.min)
		}
		return out
	}
	if cs.logBase > 1 && cs.min > 0 {
		// Logarithmic axes tick at whole powers of the base, starting at
		// the first power at or above the axis minimum (COM slide29: base
		// 2 with min 128 labels 128, 256, ..., 8192).
		var out []float64
		exp := math.Ceil(math.Log(cs.min) / math.Log(cs.logBase))
		for {
			v := math.Pow(cs.logBase, exp)
			if v > cs.max*1.0001 || v < cs.min*0.9999 {
				break
			}
			out = append(out, v)
			exp++
			if len(out) > 64 {
				break
			}
		}
		if len(out) == 0 {
			out = append(out, cs.min)
		}
		return out
	}
	if cs.step <= 0 {
		return []float64{cs.min, cs.max}
	}
	n := int(math.Round((cs.max - cs.min) / cs.step))
	if n < 0 {
		n = 0
	}
	if n > 200 {
		n = 200
	}
	out := make([]float64, 0, n+1)
	for i := 0; i <= n; i++ {
		out = append(out, cs.min+float64(i)*cs.step)
	}
	return out
}

// chartPlotArea is the data rectangle plus the scale it maps to.
type chartPlotArea struct {
	x, y, w, h     int
	ox, oy, ow, oh int // the outer area the data rectangle was fitted into
	scale          chartScale
	cats           []string
	categoriesOnY  bool // true for horizontal bar charts
	// flipCats mirrors a horizontal bar chart's rows: OOXML draws the first
	// category at the bottom (deck 00022823 slide23: "Fairplay" idx3 top,
	// "Our Results" idx0 bottom), the opposite of the label order in the file.
	flipCats bool
	// catLabelGap is the space between the category labels' right edge and
	// the plot: the legacy 5px inset, or one label line height when a
	// manualLayout pins the plot (measured 50-54px on 24pt labels).
	catLabelGap int
	// scaleX, when set, is the numeric scale for a scatter chart's X axis;
	// without it scatter points space themselves on the category index.
	scaleX *chartScale
}

func (p chartPlotArea) valueY(v float64) int {
	return p.y + p.h - int(p.scale.ratioClamped(v)*float64(p.h))
}

func (p chartPlotArea) valueX(v float64) int {
	return p.x + int(p.scale.ratioClamped(v)*float64(p.w))
}

// --- scale helpers ---

// chartSeriesRange returns the min/max across every series value.
func chartSeriesRange(series []*ChartSeries) (float64, float64) {
	minV, maxV := math.MaxFloat64, -math.MaxFloat64
	found := false
	for _, s := range series {
		for _, cat := range s.Categories {
			v := s.Values[cat]
			if math.IsNaN(v) {
				continue
			}
			if !found {
				minV, maxV, found = v, v, true
				continue
			}
			if v < minV {
				minV = v
			}
			if v > maxV {
				maxV = v
			}
		}
	}
	if !found {
		return 0, 1
	}
	return minV, maxV
}

// chartStackTotals returns the per-category totals used by stacked charts.
func chartStackTotals(series []*ChartSeries, cats []string) (float64, float64) {
	minV, maxV := math.MaxFloat64, -math.MaxFloat64
	if len(cats) == 0 {
		return 0, 1
	}
	for _, cat := range cats {
		total := 0.0
		for _, s := range series {
			total += s.Values[cat]
		}
		if total < minV {
			minV = total
		}
		if total > maxV {
			maxV = total
		}
	}
	return minV, maxV
}

// chartNiceStep picks a "nice" tick step (1/2/5 × 10^n) for a value range.
func chartNiceStep(rng float64, targetTicks int) float64 {
	if rng <= 0 || targetTicks <= 0 {
		return 1
	}
	raw := rng / float64(targetTicks)
	mag := math.Pow(10, math.Floor(math.Log10(raw)))
	norm := raw / mag
	var mult float64
	switch {
	case norm <= 1:
		mult = 1
	case norm <= 2:
		mult = 2
	case norm <= 5:
		mult = 5
	default:
		mult = 10
	}
	return mult * mag
}

// chartComputeScale resolves the value-axis scale, honouring explicit bounds
// and major units when the author set them.
func chartComputeScale(minV, maxV float64, ax *ChartAxis, startAtZero bool) chartScale {
	if ax != nil && ax.MinBounds != nil {
		minV = *ax.MinBounds
	}
	if ax != nil && ax.MaxBounds != nil {
		maxV = *ax.MaxBounds
	}
	if ax != nil && ax.LogBase > 1 && minV > 0 {
		// Logarithmic scale: the bounds snap outward to whole powers of
		// the base and ticks land on each power (majorUnit is ignored,
		// as in PowerPoint).
		lb := ax.LogBase
		if ax == nil || ax.MinBounds == nil {
			minV = math.Pow(lb, math.Floor(math.Log(minV)/math.Log(lb)+1e-9))
		}
		if ax == nil || ax.MaxBounds == nil {
			maxV = math.Pow(lb, math.Ceil(math.Log(maxV)/math.Log(lb)-1e-9))
		}
		if maxV <= minV {
			maxV = minV * lb
		}
		return chartScale{min: minV, max: maxV, logBase: lb}
	}
	// Column/line/area/scatter charts start at zero unless overridden.
	if startAtZero && (ax == nil || ax.MinBounds == nil) && minV > 0 {
		minV = 0
	}
	if maxV <= minV {
		maxV = minV + 1
	}
	step := 0.0
	if ax != nil && ax.MajorUnit != nil && *ax.MajorUnit > 0 {
		step = *ax.MajorUnit
	} else {
		step = chartNiceStep(maxV-minV, 5)
	}
	if step <= 0 {
		step = 1
	}
	// Snap the bounds outward to whole steps so labels come out even, unless
	// the author pinned them.
	if ax == nil || ax.MaxBounds == nil {
		maxV = math.Ceil(maxV/step) * step
	}
	if ax == nil || ax.MinBounds == nil {
		minV = math.Floor(minV/step) * step
	}
	return chartScale{min: minV, max: maxV, step: step}
}

// chartFormatNumber renders a value for an axis or data label.
func chartFormatNumber(v float64) string {
	if math.Abs(v) < 1e-9 {
		return "0"
	}
	if v == math.Trunc(v) && math.Abs(v) < 1e15 {
		return strconv.FormatInt(int64(v), 10)
	}
	return strconv.FormatFloat(v, 'g', 8, 64)
}

// chartFormatNumberWith renders a value with an Excel-style format code.
// Only the codes PowerPoint actually puts on value axes are honoured —
// "#,##0" (thousands groups + optional decimals), trailing "%" and a quoted
// literal "$" prefix (the accounting code on deck 00022823 chart1's value
// axis renders "$10,000") — anything else falls back to the plain rendering
// rather than guessing wrong.
func chartFormatNumberWith(v float64, format string) string {
	format = strings.TrimSpace(format)
	if format == "" || format == "General" {
		return chartFormatNumber(v)
	}
	pct := strings.Contains(format, "%")
	grouped := strings.Contains(format, "#,##0")
	dollar := strings.Contains(format, `"$"`)
	if !grouped && !pct {
		return chartFormatNumber(v)
	}
	av := math.Abs(v)
	if pct {
		av *= 100
	}
	dec := 0
	if i := strings.IndexByte(format, '.'); i >= 0 {
		for _, c := range format[i+1:] {
			if c == '0' || c == '#' {
				dec++
				continue
			}
			break
		}
	}
	s := strconv.FormatFloat(av, 'f', dec, 64)
	if grouped {
		intPart, frac := s, ""
		if i := strings.IndexByte(s, '.'); i >= 0 {
			intPart, frac = s[:i], s[i:]
		}
		n := len(intPart)
		b := make([]byte, 0, n+n/3)
		for i := 0; i < n; i++ {
			if i > 0 && (n-i)%3 == 0 {
				b = append(b, ',')
			}
			b = append(b, intPart[i])
		}
		s = string(b) + frac
	}
	if pct {
		s += "%"
	}
	if dollar {
		s = "$" + s
	}
	if v < 0 || (v == 0 && math.Signbit(v)) {
		s = "-" + s
	}
	return s
}

// axisNumFormat returns the tick-label format of an axis ("" when absent).
func axisNumFormat(ax *ChartAxis) string {
	if ax == nil {
		return ""
	}
	return ax.NumberFormat
}

// --- font / text helpers ---

// chartFace resolves a font for chart text. A nil or unset font falls back to
// the renderer default (10pt).
//
// sample is the text the face is about to draw, and it decides the face.
// Chart text used to resolve through the Latin path only, so a chart part that
// left the font unstated — or named a Latin face — drew every East Asian
// character as a .notdef box. The failure is silent by construction: the face's
// name resolves, so FontDiagnostics reports a perfect match, and only the
// pixels show the tofu. When the sample does contain East Asian characters the
// face is therefore chosen by glyph coverage first, exactly as the slide text
// path does, and any substitution is reported.
func (r *renderer) chartFace(f *Font, sample string) font.Face {
	if f == nil {
		f = NewFont()
	}
	if f.Size <= 0 {
		clone := *f
		clone.Size = 10
		f = &clone
	}
	if containsCJK(sample) {
		face, used := r.getCJKFace(f, sample)
		r.noteCJKFont(f, used)
		if face != nil {
			return face
		}
	}
	return r.getFace(f)
}

// chartFaceForCats resolves the face for a set of labels drawn as one group.
func (r *renderer) chartFaceForCats(f *Font, cats []string) font.Face {
	return r.chartFace(f, strings.Join(cats, ""))
}

// chartFontColor returns the colour to use for a font, defaulting to the
// standard chart text grey.
func chartFontColor(f *Font) color.RGBA {
	if f != nil && f.Color.ARGB != "" && f.Color.GetAlpha() != 0 {
		return argbToRGBA(f.Color)
	}
	return color.RGBA{R: 64, G: 64, B: 64, A: 255}
}

// chartLineHeight is the pixel height of a single line of text.
func (r *renderer) chartLineHeight(face font.Face) int {
	m := face.Metrics()
	return (m.Ascent + m.Descent).Ceil()
}

// drawChartText draws text with its baseline at (x, baselineY).
func (r *renderer) drawChartText(text string, face font.Face, c color.RGBA, x, baselineY int) {
	if text == "" || face == nil {
		return
	}
	d := &font.Drawer{
		Dst:  r.img,
		Src:  image.NewUniform(drawSrc(c)),
		Face: face,
		Dot:  fixed.P(x, baselineY),
	}
	d.DrawString(text)
}

// chartTextWidth measures a string in pixels.
func chartTextWidth(face font.Face, s string) int {
	if face == nil || s == "" {
		return 0
	}
	return font.MeasureString(face, s).Ceil()
}

// chartTruncate shortens s with an ellipsis so it fits within maxW pixels.
func chartTruncate(face font.Face, s string, maxW int) string {
	if maxW <= 0 || face == nil || s == "" {
		return ""
	}
	if chartTextWidth(face, s) <= maxW {
		return s
	}
	runes := []rune(s)
	for len(runes) > 1 {
		runes = runes[:len(runes)-1]
		candidate := string(runes) + "…"
		if chartTextWidth(face, candidate) <= maxW {
			return candidate
		}
	}
	return ""
}

// chartMaxLabelWidth returns the widest label, capped at limit.
func chartMaxLabelWidth(face font.Face, labels []string, limit int) int {
	w := 0
	for _, l := range labels {
		if tw := chartTextWidth(face, l); tw > w {
			w = tw
		}
	}
	if limit > 0 && w > limit {
		w = limit
	}
	return w
}

// --- chart entry points ---

func (r *renderer) renderChart(s *ChartShape) {
	x := r.emuToPixelX(s.offsetX)
	y := r.emuToPixelY(s.offsetY)
	w := r.emuToPixelX(s.width)
	h := r.emuToPixelY(s.height)
	if w <= 0 || h <= 0 {
		return
	}
	frame := image.Rect(x, y, x+w, y+h)

	// Chart area fill / outline. Both default to "none" in PowerPoint, which
	// means the slide background shows through.
	if s.fill != nil && s.fill.Type != FillNone {
		r.renderFill(s.fill, frame)
	}
	if s.border != nil && s.border.Style != BorderNone {
		pw := maxInt(int(float64(maxInt(s.border.Width, 1))*12700.0*r.scaleX), 1)
		r.drawRectBorder(frame, argbToRGBA(s.border.Color), pw, s.border.Style)
	}

	pad := maxInt(int(4.0*r.scaleX), 2)

	// Title strip at the top.
	titleH := 0
	if s.title != nil && s.title.Visible && s.title.Text != "" {
		face := r.chartFace(s.title.Font, s.title.Text)
		titleH = r.chartLineHeight(face) + 4
		r.drawStringCentered(s.title.Text, face, chartFontColor(s.title.Font),
			image.Rect(x+pad, y+pad, x+w-pad, y+pad+titleH))
		titleH += pad
	}

	ct := s.plotArea.GetType()
	if ct == nil {
		return
	}

	// Legend reserves space along one edge.
	legendPos := LegendBottom
	legendH := 0
	legendNames, _ := r.chartLegendEntries(s)
	if s.legend != nil && s.legend.Visible && len(legendNames) > 0 {
		legendPos = s.legend.Position
		legendH = r.chartLegendHeight(s, x, w)
	}

	innerX := x + pad
	innerY := y + pad + titleH
	innerW := w - 2*pad
	innerH := h - 2*pad - titleH

	switch legendPos {
	case LegendLeft, LegendRight:
		legendW := r.chartLegendWidth(s, innerW)
		if legendW > innerW/2 {
			legendW = innerW / 2
		}
		if legendPos == LegendLeft {
			innerX += legendW
		}
		innerW -= legendW
	case LegendTop, LegendTopRight:
		innerY += legendH
		innerH -= legendH
	default: // bottom
		innerH -= legendH
	}

	// <c:manualLayout> pins the inner plot rect as fractions of the chart
	// frame. Measured against COM exports the placement is exact (deck
	// 00022823 slide22: layout right edge 1449.5 vs measured 1449), so it is
	// applied after the title/legend strips, on frame-relative coordinates.
	if L := s.plotArea.layout; L != nil && L.w > 0 && L.h > 0 {
		nx, ny := x+int(L.x*float64(w)), y+int(L.y*float64(h))
		nx2, ny2 := x+int((L.x+L.w)*float64(w)), y+int((L.y+L.h)*float64(h))
		innerX, innerY, innerW, innerH = nx, ny, nx2-nx, ny2-ny
	}

	if innerW < 8 || innerH < 8 {
		return
	}

	switch c := ct.(type) {
	case *BarChart:
		r.renderBarChart(c, s, innerX, innerY, innerW, innerH)
	case *Bar3DChart:
		r.renderBarChart(&c.BarChart, s, innerX, innerY, innerW, innerH)
	case *LineChart:
		r.renderLineChart(c, s, innerX, innerY, innerW, innerH)
	case *PieChart:
		r.renderPieChart(c.Series, s, innerX, innerY, innerW, innerH)
	case *Pie3DChart:
		r.renderPieChart(c.Series, s, innerX, innerY, innerW, innerH)
	case *DoughnutChart:
		r.renderDoughnutChart(c, s, innerX, innerY, innerW, innerH)
	case *AreaChart:
		r.renderAreaChart(c, s, innerX, innerY, innerW, innerH)
	case *ScatterChart:
		r.renderScatterChart(c, s, innerX, innerY, innerW, innerH)
	case *RadarChart:
		r.renderRadarChart(c, s, innerX, innerY, innerW, innerH)
	}

	if s.legend != nil && s.legend.Visible && len(legendNames) > 0 {
		r.renderChartLegend(s, x, y, w, h, titleH, legendPos)
	}
}

// --- plot area & axes ---

// chartPlotAreaFor reserves room for the axis labels and returns the data rect.
func (r *renderer) chartPlotAreaFor(s *ChartShape, px, py, pw, ph int, sc chartScale, cats []string, categoriesOnY bool) chartPlotArea {
	axV := s.plotArea.GetAxisY()
	axX := s.plotArea.GetAxisX()
	// The value labels are formatted numbers, so they never carry East Asian
	// text; the category labels are the document's own strings and may.
	valueFace := r.chartFace(axisFont(axV), "")
	catFace := r.chartFaceForCats(axisFont(axX), cats)

	// With a manualLayout the rect handed in IS the plot rect: the category
	// labels live outside it (right-aligned, one line height away), and only
	// when they cannot fit does the plot shove right — PowerPoint measured
	// behaviour on deck 00022823 slides 22/23.
	if s.plotArea.layout != nil {
		base := chartPlotArea{
			ox: px, oy: py, ow: pw, oh: ph,
			scale: sc, cats: cats, categoriesOnY: categoriesOnY,
			catLabelGap: r.chartLineHeight(catFace),
		}
		base.x, base.y, base.w, base.h = px, py, pw, ph
		if categoriesOnY {
			gap := base.catLabelGap
			lw := chartMaxLabelWidth(catFace, cats, 0)
			frameX := r.emuToPixelX(s.offsetX)
			if px-gap-lw < frameX {
				// Labels don't fit between the frame edge and the pinned
				// plot — slide23 does exactly this (labels end 275, pinned
				// plot starts 165, PowerPoint moves the plot to 329).
				shift := frameX + lw + gap - px
				base.x = px + shift
				base.w = pw - shift
			}
		}
		if base.w < 4 || base.h < 4 {
			base.x, base.y, base.w, base.h = px, py, pw, ph
		}
		return base
	}

	left, top, right, bottom := 3, 2, 4, 2
	if axV == nil || axV.Visible {
		if categoriesOnY {
			bottom += r.chartLineHeight(valueFace) + 2
		} else {
			lw := 0
			for _, t := range sc.ticks() {
				if tw := chartTextWidth(valueFace, chartFormatNumberWith(t, axisNumFormat(axV))); tw > lw {
					lw = tw
				}
			}
			// Reserve the widest label plus the value-label gap (one line
			// height + 4px, see chartValueLabelGap).
			left = lw + r.chartValueLabelGap(valueFace)
		}
	}
	if axX == nil || axX.Visible {
		if categoriesOnY {
			left += chartMaxLabelWidth(catFace, cats, pw/2) + 5
		} else if axX != nil && chartIsDateFormat(axisNumFormat(axX)) {
			// Rotated date labels need a column as tall as the longest
			// formatted label, plus the label line height. The extra
			// top/left/right insets are PowerPoint's measured auto layout
			// for this chart family (COM slide05: plot top 58, left 287,
			// right 1537 on a frame at 27,27,1533x1093); the value-label
			// width and its line-height clearance are already reserved
			// above, so this term is what lands the axis at x=287.
			left += r.chartLineHeight(catFace) - 16
			top += r.chartLineHeight(valueFace) - 9
			right += 20
			maxLen := 0
			for _, cat := range cats {
				label := cat
				if v, err := strconv.ParseFloat(strings.TrimSpace(cat), 64); err == nil {
					label = chartFormatDate(v, axisNumFormat(axX))
				}
				if l := chartTextWidth(catFace, label); l > maxLen {
					maxLen = l
				}
			}
			bottom += maxLen + 2 + r.chartLineHeight(catFace)
		} else {
			bottom += r.chartLineHeight(catFace) + 2
		}
	}
	// Axis titles get their own strip: the value-axis title sits beyond the
	// value labels, the category-axis title beyond the category labels.
	if axV != nil && axV.Visible && strings.TrimSpace(axV.Title) != "" {
		if categoriesOnY {
			bottom += r.chartLineHeight(valueFace) + 3
		} else {
			left += r.chartLineHeight(valueFace) + 3
		}
	}
	if axX != nil && axX.Visible && strings.TrimSpace(axX.Title) != "" {
		if categoriesOnY {
			left += r.chartLineHeight(catFace) + 3
		} else {
			bottom += r.chartLineHeight(catFace) + 3
		}
	}

	dataX, dataY := px+left, py+top
	dataW, dataH := pw-left-right, ph-top-bottom
	base := chartPlotArea{
		ox: px, oy: py, ow: pw, oh: ph,
		scale: sc, cats: cats, categoriesOnY: categoriesOnY,
		catLabelGap: 5,
	}
	if dataW < 4 || dataH < 4 {
		// Not enough room for labels — draw the series in the full area.
		base.x, base.y, base.w, base.h = px, py, pw, ph
		return base
	}
	base.x, base.y, base.w, base.h = dataX, dataY, dataW, dataH
	return base
}

// axisFont returns an axis font, or nil when the axis is absent.
func axisFont(a *ChartAxis) *Font {
	if a == nil {
		return nil
	}
	return a.Font
}

// drawChartAxes renders the value-axis gridlines plus the value and category
// tick labels.
func (r *renderer) drawChartAxes(s *ChartShape, p chartPlotArea) {
	axV := s.plotArea.GetAxisY()
	axX := s.plotArea.GetAxisX()

	// Major gridlines. Drawn only when the chart declares them: charts 3/4
	// of deck 00022823 carry no <c:majorGridlines> and the COM exports show
	// none, while every chart that does declare them keeps its grid — unless
	// the stroke is <a:noFill/> (chart5's Y axis), which stays invisible.
	//
	// PowerPoint's default gridline stroke (no <a:ln> spPr) measures
	// 134,134,134 at ~2px/160dpi ≈ 9525 EMU (0.75pt) on both test decks —
	// the light 1px #D9D9D9 this used to draw appears in neither.
	if axV != nil && axV.MajorGridlines != nil && !axV.MajorGridlines.NoFill {
		gridColor := color.RGBA{R: 134, G: 134, B: 134, A: 255}
		// drawLineAA's Wu sub-lines land ~1px narrower than the requested
		// width, so +1 reproduces PowerPoint's crisp 2px stroke.
		gridW := maxInt(int(9525.0*r.scaleX+0.5)+1, 2)
		if axV.MajorGridlines.ColorSet {
			gridColor = argbToRGBA(axV.MajorGridlines.Color)
		}
		if axV.MajorGridlines.Width > 0 {
			gridW = axV.MajorGridlines.Width
		}
		for _, t := range p.scale.ticks() {
			if p.categoriesOnY {
				gx := p.valueX(t)
				r.drawLineAA(gx, p.y, gx, p.y+p.h, gridColor, gridW)
			} else {
				gy := p.valueY(t)
				r.drawLineAA(p.x, gy, p.x+p.w, gy, gridColor, gridW)
			}
		}
	}

	// Axis lines. PowerPoint's default axis stroke matches the gridlines:
	// 134,134,134 at 0.75pt (measured on deck 00022823 slide05's value axis
	// and slide22/23's category axes).
	axisLine := color.RGBA{R: 134, G: 134, B: 134, A: 255}
	axisW := maxInt(int(9525.0*r.scaleX+0.5)+1, 2)
	// Major tick marks: "out" is PowerPoint's default for chart axes — deck
	// 00022823 chart1/chart2 omit <c:majorTickMark> and the COM export shows
	// 11px ticks, while chart3/4/5 declare val="none" and show none. Length
	// measures ~5pt on the COM golds.
	tickLen := maxInt(int(63500.0*r.scaleX+0.5), 2)
	if p.categoriesOnY {
		r.drawLineAA(p.x, p.y, p.x, p.y+p.h, axisLine, axisW)
	} else {
		r.drawLineAA(p.x, p.y+p.h, p.x+p.w, p.y+p.h, axisLine, axisW)
		// A numeric X axis (scatter, or a date category axis) gets a
		// vertical value-axis line too — the COM exports show it.
		if p.scaleX != nil {
			r.drawLineAA(p.x, p.y, p.x, p.y+p.h, axisLine, axisW)
		}
	}

	// Major tick marks. Value-axis ticks poke outward (left for a vertical
	// axis); the tick at the scale maximum is not drawn (COM slide05: the
	// 1e8 gridline at y=58 carries no tick while the decades below do).
	if axV != nil && axV.MajorTickMark != "" && axV.MajorTickMark != TickMarkNone && axV.Visible {
		tickColor := axisLine
		for _, t := range p.scale.ticks() {
			if !p.categoriesOnY {
				if maxV := p.scale.max; maxV > p.scale.min && t >= maxV {
					continue
				}
			}
			if p.categoriesOnY {
				gx := p.valueX(t)
				r.drawLineAA(gx, p.y+p.h, gx, p.y+p.h+tickLen, tickColor, axisW)
			} else {
				gy := p.valueY(t)
				r.drawLineAA(p.x-tickLen, gy, p.x, gy, tickColor, axisW)
			}
		}
	}
	// Category-axis ticks along the bottom (numeric/date scales tick at
	// scale positions): short vertical strokes below the axis line.
	if !p.categoriesOnY && p.scaleX != nil && axX != nil && axX.MajorTickMark != "" && axX.MajorTickMark != TickMarkNone && axX.Visible {
		for _, t := range p.scaleX.ticks() {
			tx := p.x + int(p.scaleX.ratioClamped(t)*float64(p.w))
			r.drawLineAA(tx, p.y+p.h, tx, p.y+p.h+tickLen, axisLine, axisW)
		}
	}

	// Value labels.
	if axV == nil || axV.Visible {
		face := r.chartFace(axisFont(axV), "")
		fc := chartFontColor(axisFont(axV))
		ascent := face.Metrics().Ascent.Ceil()
		descent := face.Metrics().Descent.Ceil()
		for _, t := range p.scale.ticks() {
			label := chartFormatNumberWith(t, axisNumFormat(axV))
			if p.categoriesOnY {
				tw := chartTextWidth(face, label)
				r.drawChartText(label, face, fc, p.valueX(t)-tw/2, p.y+p.h+2+ascent)
			} else {
				tw := chartTextWidth(face, label)
				gy := p.valueY(t)
				r.drawChartText(label, face, fc, p.x-r.chartValueLabelGap(face)-tw, gy+(ascent+descent)/2-descent)
			}
		}
	}

	// X-axis labels: numeric scales (scatter / date axes) tick at scale
	// positions; everything else labels the categories evenly.
	if p.scaleX != nil {
		if axX == nil || axX.Visible {
			face := r.chartFace(axisFont(axX), "")
			fc := chartFontColor(axisFont(axX))
			m := face.Metrics()
			ascent := m.Ascent.Ceil()
			dateFmt := chartIsDateFormat(axisNumFormat(axX))
			for _, t := range p.scaleX.ticks() {
				var label string
				if dateFmt {
					label = chartFormatDate(t, axisNumFormat(axX))
				} else {
					label = chartFormatNumberWith(t, axisNumFormat(axX))
				}
				tx := p.x + int(p.scaleX.ratioClamped(t)*float64(p.w))
				if dateFmt {
					// PowerPoint rotates date labels that would collide
					// horizontally; they read bottom-to-top under the axis.
					tx += r.rotatedTickLabelOffset()
					r.drawChartRotatedLabel(label, face, fc, tx, p.y+p.h+3)
				} else {
					tw := chartTextWidth(face, label)
					r.drawChartText(label, face, fc, tx-tw/2, p.y+p.h+2+ascent)
				}
			}
		}
	} else if (axX == nil || axX.Visible) && len(p.cats) > 0 {
		face := r.chartFaceForCats(axisFont(axX), p.cats)
		fc := chartFontColor(axisFont(axX))
		ascent := face.Metrics().Ascent.Ceil()
		descent := face.Metrics().Descent.Ceil()
		gap := p.catLabelGap
		if gap <= 0 {
			gap = 5
		}
		if p.categoriesOnY {
			rowH := float64(p.h) / float64(len(p.cats))
			for i, cat := range p.cats {
				row := i
				if p.flipCats {
					row = len(p.cats) - 1 - i
				}
				cy := p.y + int((float64(row)+0.5)*rowH)
				tw := chartTextWidth(face, cat)
				r.drawChartText(cat, face, fc, p.x-gap-tw, cy+(ascent+descent)/2-descent)
			}
		} else {
			catW := float64(p.w) / float64(len(p.cats))
			for i, cat := range p.cats {
				label := cat
				tw := chartTextWidth(face, label)
				if tw > int(catW)-2 {
					label = chartTruncate(face, label, int(catW)-2)
					tw = chartTextWidth(face, label)
				}
				cx := p.x + int((float64(i)+0.5)*catW)
				r.drawChartText(label, face, fc, cx-tw/2, p.y+p.h+2+ascent)
			}
		}
	}

	r.drawChartAxisTitles(s, p)
}

// drawChartAxisTitles renders the value- and category-axis titles. The title of
// the axis that runs vertically is rotated 270°, matching PowerPoint.
//
// With a pinned manualLayout the plot rect no longer touches the frame, so
// both title strips are measured from the frame edge: the value title spans
// frame-left..plot-left, the category title sits centred in the strip
// between the tick labels and the frame bottom (COM slide29: title centre
// 1074 in the 1025..1133 strip).
func (r *renderer) drawChartAxisTitles(s *ChartShape, p chartPlotArea) {
	axV := s.plotArea.GetAxisY()
	axX := s.plotArea.GetAxisX()
	frameL := r.emuToPixelX(s.offsetX)
	frameB := r.emuToPixelY(s.offsetY) + r.emuToPixelY(s.height)

	if axV != nil && axV.Visible && strings.TrimSpace(axV.Title) != "" {
		face := r.chartFace(axV.Font, axV.Title)
		fc := chartFontColor(axV.Font)
		if p.categoriesOnY {
			// Value axis runs along the bottom.
			h := r.chartLineHeight(face)
			y := p.oy + p.oh - h
			if y < p.y+p.h {
				y = p.y + p.h
			}
			r.drawChartTextCentered(axV.Title, face, fc, p.x, y, p.w, h)
		} else if w := p.x - frameL; w > 0 {
			r.drawChartVerticalCentered(axV.Title, face, fc, frameL, p.y, w, p.h)
		}
	}
	if axX != nil && axX.Visible && strings.TrimSpace(axX.Title) != "" {
		face := r.chartFace(axX.Font, axX.Title)
		fc := chartFontColor(axX.Font)
		if p.categoriesOnY {
			if w := p.x - frameL; w > 0 {
				r.drawChartVerticalCentered(axX.Title, face, fc, frameL, p.y, w, p.h)
			}
		} else {
			h := r.chartLineHeight(face)
			y := p.oy + p.oh - h
			if y < p.y+p.h {
				y = p.y + p.h
			}
			if frameB > p.y+p.h {
				// Centre the title in the strip between the tick labels
				// and the frame bottom.
				stripTop := y + h + 4
				if cy := stripTop + (frameB-stripTop-h)/2; cy > y {
					y = cy
				}
			}
			r.drawChartTextCentered(axX.Title, face, fc, p.x, y, p.w, h)
		}
	}
}

// drawChartTextCentered draws text centred inside a box.
func (r *renderer) drawChartTextCentered(text string, face font.Face, c color.RGBA, x, y, w, h int) {
	if text == "" || face == nil || w <= 0 || h <= 0 {
		return
	}
	tw := chartTextWidth(face, text)
	m := face.Metrics()
	lineH := (m.Ascent + m.Descent).Ceil()
	baseline := y + (h-lineH)/2 + m.Ascent.Ceil()
	r.drawChartText(text, face, c, x+(w-tw)/2, baseline)
}

// drawChartVerticalCentered draws text rotated 270° (bottom-to-top) centred in
// a box, which is how PowerPoint renders a vertical axis title.
func (r *renderer) drawChartVerticalCentered(text string, face font.Face, c color.RGBA, x, y, w, h int) {
	if text == "" || face == nil || w <= 0 || h <= 0 {
		return
	}
	// Render into a buffer with swapped dimensions, then rotate back.
	tmp := image.NewRGBA(image.Rect(0, 0, h, w))
	tr := r.subRenderer(tmp)
	tr.drawChartTextCentered(text, face, c, 0, 0, h, w)
	rotateAndComposite(r.img, tmp, x, y, w, h, 270)
}

// --- bar chart ---

func (r *renderer) renderBarChart(c *BarChart, s *ChartShape, px, py, pw, ph int) {
	if len(c.Series) == 0 || pw <= 0 || ph <= 0 {
		return
	}
	cats := c.Series[0].Categories
	if len(cats) == 0 {
		return
	}
	palette := chartColors()

	stacked := c.BarGrouping == BarGroupingStacked || c.BarGrouping == BarGroupingPercentStacked
	percent := c.BarGrouping == BarGroupingPercentStacked
	horizontal := c.BarDirection == BarDirectionHorizontal

	minV, maxV := chartSeriesRange(c.Series)
	if stacked {
		minV, maxV = chartStackTotals(c.Series, cats)
	}
	if percent {
		minV, maxV = 0, 100
	}
	axCat := s.plotArea.GetAxisX()
	sc := chartComputeScale(minV, maxV, s.plotArea.GetAxisY(), !percent)
	plot := r.chartPlotAreaFor(s, px, py, pw, ph, sc, cats, horizontal)
	// Horizontal bar charts draw the first category at the bottom (OOXML
	// semantics; maxMin orientation restores file order).
	if horizontal {
		plot.flipCats = axCat == nil || !axCat.ReversedOrder
	}
	r.drawChartAxes(s, plot)

	gapPct := c.GapWidthPercent
	if gapPct < 0 {
		gapPct = 150
	}
	// PowerPoint's gap width is a percentage of the bar/cluster width.
	groupGapFactor := 100.0 / (100.0 + float64(gapPct))

	nCats := len(cats)
	nSeries := len(c.Series)

	if horizontal {
		rowH := float64(plot.h) / float64(nCats)
		clusterH := rowH * groupGapFactor
		if clusterH < 1 {
			clusterH = 1
		}
		for ci, cat := range cats {
			row := ci
			if plot.flipCats {
				row = nCats - 1 - ci
			}
			rowY := float64(plot.y) + float64(row)*rowH + (rowH-clusterH)/2
			if stacked {
				acc := 0.0
				for si, ser := range c.Series {
					v := ser.Values[cat]
					if percent {
						v = chartPercentShare(c.Series, cat, v)
					}
					x0 := plot.valueX(acc)
					x1 := plot.valueX(acc + v)
					r.fillHBar(int(rowY), int(clusterH), x0, x1,
						seriesPointColor(ser, si, ci, palette))
					acc += v
				}
			} else {
				barH := clusterH / float64(nSeries)
				step := barH
				ov := c.OverlapPercent
				if ov != 0 {
					den := float64(nSeries) - float64(nSeries-1)*float64(ov)/100.0
					if den > 0.1 {
						barH = clusterH / den
						step = barH * (1 - float64(ov)/100.0)
					}
				}
				for si, ser := range c.Series {
					v := ser.Values[cat]
					y0 := int(rowY + float64(si)*step)
					bh := int(barH)
					x0 := plot.valueX(0)
					x1 := plot.valueX(v)
					r.fillHBar(y0, bh, x0, x1,
						seriesPointColor(ser, si, ci, palette))
				}
			}
		}
	} else {
		catW := float64(plot.w) / float64(nCats)
		clusterW := catW * groupGapFactor
		if clusterW < 1 {
			clusterW = 1
		}
		for ci, cat := range cats {
			groupX := float64(plot.x) + float64(ci)*catW + (catW-clusterW)/2
			if stacked {
				acc := 0.0
				for si, ser := range c.Series {
					v := ser.Values[cat]
					if percent {
						v = chartPercentShare(c.Series, cat, v)
					}
					y0 := plot.valueY(acc)
					y1 := plot.valueY(acc + v)
					r.fillVBar(int(groupX), int(clusterW), y0, y1,
						seriesPointColor(ser, si, ci, palette))
					acc += v
				}
			} else {
				barW := clusterW / float64(nSeries)
				step := barW
				ov := c.OverlapPercent
				if ov != 0 {
					den := float64(nSeries) - float64(nSeries-1)*float64(ov)/100.0
					if den > 0.1 {
						barW = clusterW / den
						step = barW * (1 - float64(ov)/100.0)
					}
				}
				for si, ser := range c.Series {
					v := ser.Values[cat]
					bx := int(groupX + float64(si)*step)
					bw := int(barW)
					y0 := plot.valueY(0)
					y1 := plot.valueY(v)
					r.fillVBar(bx, bw, y0, y1,
						seriesPointColor(ser, si, ci, palette))
				}
			}
		}
	}
}

// chartPercentShare returns a value's percentage share within its category,
// used for percentStacked bar charts.
func chartPercentShare(series []*ChartSeries, cat string, v float64) float64 {
	total := 0.0
	for _, s := range series {
		total += s.Values[cat]
	}
	if total == 0 {
		return 0
	}
	return v / total * 100
}

// fillVBar fills a vertical bar occupying x..x+w between pixel rows y0 and y1.
func (r *renderer) fillVBar(x, w, y0, y1 int, c color.RGBA) {
	if w < 1 {
		w = 1
	}
	top, bot := y0, y1
	if top > bot {
		top, bot = bot, top
	}
	if bot-top < 1 {
		bot = top + 1
	}
	r.fillRectBlend(image.Rect(x, top, x+w, bot), c)
}

// fillHBar fills a horizontal bar occupying y..y+h between pixel columns x0 and x1.
func (r *renderer) fillHBar(y, h, x0, x1 int, c color.RGBA) {
	if h < 1 {
		h = 1
	}
	left, right := x0, x1
	if left > right {
		left, right = right, left
	}
	if right-left < 1 {
		right = left + 1
	}
	r.fillRectBlend(image.Rect(left, y, right, y+h), c)
}

// --- line chart ---

// chartSeriesLineWidth resolves the stroke width in pixels for a series, or
// fallback when the series states none.
//
// SeriesOutline.Width is in points — the same unit Border.Width uses, because
// the reader divides the EMU it reads by 12700 — so it has to be scaled by
// 12700 EMU/point before scaleX turns EMU into pixels. The expression this
// replaced multiplied and divided by 12700 in the same breath, which cancelled
// out and left the point count multiplied by scaleX: about 1e-4, so it always
// rounded down to the 1px floor and every series line came out hairline thin no
// matter what the document asked for.
func (r *renderer) chartSeriesLineWidth(ser *ChartSeries, fallback int) int {
	if ser != nil && ser.Outline != nil && ser.Outline.Width > 0 {
		return maxInt(int(float64(ser.Outline.Width)*12700.0*r.scaleX), 1)
	}
	if fallback <= 0 {
		return fallback
	}
	// A chart series without an explicit stroke width renders at PowerPoint's
	// default 28575 EMU (2.25pt) — COM exports of deck 00022823 slide05 show
	// the unstyled "Cost per Genome" line ~5px wide at 160dpi, not the 1-2px
	// hairline this used to draw. The fallback parameter only signals whether
	// the caller wants a default at all.
	w := int(28575.0*r.scaleX + 0.5)
	return maxInt(w, 1)
}

func (r *renderer) renderLineChart(c *LineChart, s *ChartShape, px, py, pw, ph int) {
	if len(c.Series) == 0 || pw <= 0 || ph <= 0 {
		return
	}
	cats := c.Series[0].Categories
	if len(cats) == 0 {
		return
	}
	palette := chartColors()

	minV, maxV := chartSeriesRange(c.Series)
	sc := chartComputeScale(minV, maxV, s.plotArea.GetAxisY(), true)
	plot := r.chartPlotAreaFor(s, px, py, pw, ph, sc, cats, false)
	// A date category axis maps categories at their serial positions on a
	// time scale, not at even category spacing (deck 00022823 chart1: the
	// gaps between points vary between ~5 and ~7 months).
	if axX := s.plotArea.GetAxisX(); axX != nil && chartIsDateFormat(axisNumFormat(axX)) {
		if min0, err0 := strconv.ParseFloat(strings.TrimSpace(cats[0]), 64); err0 == nil {
			if max0, err1 := strconv.ParseFloat(strings.TrimSpace(cats[len(cats)-1]), 64); err1 == nil && max0 > min0 {
				// COM slide05: a 3318-day (109-month) span over a ~1250px
				// plot ticks every 5 calendar months — roughly plot
				// width / 57 per tick, rounded to whole months.
				spanMonths := (max0 - min0) / 30.4375
				months := int(math.Round(spanMonths / float64(maxInt(2, plot.w/57))))
				if months < 1 {
					months = 1
				}
				scX := chartScale{min: min0, max: max0, stepMonths: months}
				plot.scaleX = &scX
			}
		}
	}
	r.drawChartAxes(s, plot)

	nPts := len(cats)
	for si, ser := range c.Series {
		sc2 := getSeriesColor(ser, si, palette)
		lw := r.chartSeriesLineWidth(ser, 2)
		// Blank points (NaN) break the line — PowerPoint's dispBlanksAs
		// "gap" treatment.
		var runXs, runYs []int
		flush := func() {
			if len(runXs) > 1 {
				if c.IsSmooth && len(runXs) >= 3 {
					r.drawSmoothCurve(runXs, runYs, sc2, lw)
				} else {
					r.drawChartLineDashed(runXs, runYs, sc2, lw, ser.LineDash)
				}
			}
			for i := range runXs {
				r.drawChartMarker(ser, runXs[i], runYs[i], sc2)
			}
			runXs, runYs = runXs[:0], runYs[:0]
		}
		for i, cat := range cats {
			v := ser.Values[cat]
			if math.IsNaN(v) {
				flush()
				continue
			}
			runXs = append(runXs, chartCatX(plot, i, nPts, cat))
			runYs = append(runYs, plot.valueY(v))
		}
		flush()
	}
}

// chartCatX returns the pixel X for a category: numeric date scales map the
// serial, everything else spaces categories evenly.
func chartCatX(p chartPlotArea, i, n int, cat string) int {
	if p.scaleX != nil {
		if v, err := strconv.ParseFloat(strings.TrimSpace(cat), 64); err == nil {
			return p.x + int(p.scaleX.ratioClamped(v)*float64(p.w))
		}
	}
	return chartPointX(p, i, n)
}

// chartPointX returns the pixel X for the i-th of n category points.
func chartPointX(p chartPlotArea, i, n int) int {
	if n <= 1 {
		return p.x + p.w/2
	}
	return p.x + i*p.w/(n-1)
}

// drawSmoothCurve draws a Catmull-Rom style smoothed polyline through points.
func (r *renderer) drawSmoothCurve(xs, ys []int, c color.RGBA, width int) {
	if len(xs) < 2 {
		return
	}
	const steps = 8
	for i := 0; i < len(xs)-1; i++ {
		p0 := i - 1
		if p0 < 0 {
			p0 = 0
		}
		p3 := i + 2
		if p3 >= len(xs) {
			p3 = len(xs) - 1
		}
		prevX, prevY := xs[i], ys[i]
		for t := 1; t <= steps; t++ {
			ft := float64(t) / float64(steps)
			x := catmullRom(float64(xs[p0]), float64(xs[i]), float64(xs[i+1]), float64(xs[p3]), ft)
			y := catmullRom(float64(ys[p0]), float64(ys[i]), float64(ys[i+1]), float64(ys[p3]), ft)
			ix, iy := int(math.Round(x)), int(math.Round(y))
			r.drawLineAA(prevX, prevY, ix, iy, c, width)
			prevX, prevY = ix, iy
		}
	}
}

// catmullRom evaluates a Catmull-Rom spline segment at t in [0,1].
func catmullRom(p0, p1, p2, p3, t float64) float64 {
	t2 := t * t
	t3 := t2 * t
	return 0.5 * ((2 * p1) +
		(-p0+p2)*t +
		(2*p0-5*p1+4*p2-p3)*t2 +
		(-p0+3*p1-3*p2+p3)*t3)
}

// drawChartMarker draws a series marker using its configured symbol.
func (r *renderer) drawChartMarker(ser *ChartSeries, x, y int, c color.RGBA) {
	symbol := MarkerCircle
	size := 5
	if ser.Marker != nil {
		if ser.Marker.Symbol != "" {
			symbol = ser.Marker.Symbol
		}
		if ser.Marker.Size > 0 {
			size = ser.Marker.Size
		}
	}
	// Marker size is in points (<c:size val="8"/> = 8pt diameter): scale it
	// through 12700 EMU/pt like every other point-denominated length. The
	// old expression fed the raw point count to scaleX (px per EMU), which
	// collapsed 8pt to a 2px dot — deck 00022823 slide05's diamond markers
	// vanished entirely at 160dpi.
	rad := maxInt(int(float64(size)*12700.0*r.scaleX/2), 2)
	switch symbol {
	case MarkerNone:
		return
	case MarkerSquare:
		r.fillRectBlend(image.Rect(x-rad, y-rad, x+rad, y+rad), c)
	case MarkerDiamond:
		pts := []fpoint{
			{float64(x), float64(y - rad)},
			{float64(x + rad), float64(y)},
			{float64(x), float64(y + rad)},
			{float64(x - rad), float64(y)},
		}
		r.fillPolygon(pts, c)
	case MarkerTriangle:
		pts := []fpoint{
			{float64(x), float64(y - rad)},
			{float64(x + rad), float64(y + rad)},
			{float64(x - rad), float64(y + rad)},
		}
		r.fillPolygon(pts, c)
	case MarkerPlus, MarkerX:
		r.drawLineAA(x-rad, y, x+rad, y, c, 1)
		r.drawLineAA(x, y-rad, x, y+rad, c, 1)
		if symbol == MarkerX {
			r.drawLineAA(x-rad, y-rad, x+rad, y+rad, c, 1)
			r.drawLineAA(x-rad, y+rad, x+rad, y-rad, c, 1)
		}
	default: // circle / dot / dash / star / anything else
		r.fillEllipseAA(x-rad, y-rad, rad*2+1, rad*2+1, c)
	}
}

// rotatedTickLabelOffset shifts a rotated date label's ink column right of
// its tick: on the COM golds (deck 00022823 slide05/06) the label centers
// ~5px (2.25pt at 160dpi) right of the tick — the ascent side of the glyph
// box eats into the left of the column.
func (r *renderer) rotatedTickLabelOffset() int {
	return int(28575.0*r.scaleX + 0.5)
}

// chartValueLabelGap is the clearance between the value-axis labels' right
// edge and the axis line: one label line height plus 4px (COM slide05:
// "$10,000,000" ink ends 42px left of the axis at 16pt). Both the plot-area
// reservation and the label draw use it, or the labels detach from the axis.
func (r *renderer) chartValueLabelGap(face font.Face) int {
	return r.chartLineHeight(face) + 4
}

// --- pie / doughnut ---

func (r *renderer) renderPieChart(series []*ChartSeries, s *ChartShape, px, py, pw, ph int) {
	if len(series) == 0 || len(series[0].Categories) == 0 {
		return
	}
	palette := chartColors()
	ser := series[0]

	vals := make([]float64, 0, len(ser.Categories))
	total := 0.0
	for _, cat := range ser.Categories {
		v := ser.Values[cat]
		if v > 0 {
			vals = append(vals, v)
			total += v
		} else {
			vals = append(vals, 0)
		}
	}
	if total == 0 {
		return
	}

	cx := px + pw/2
	cy := py + ph/2
	radius := minInt(pw, ph)/2 - 2
	if radius < 5 {
		return
	}

	startAngle := -math.Pi / 2
	for i, cat := range ser.Categories {
		v := vals[i]
		if v <= 0 {
			continue
		}
		sweep := 2 * math.Pi * v / total
		endAngle := startAngle + sweep
		sc := palette[i%len(palette)]
		r.fillPieSlice(cx, cy, radius, startAngle, endAngle, sc)
		r.drawPieLabel(ser, cat, v/total*100, cx, cy, radius, startAngle, endAngle)
		startAngle = endAngle
	}
}

// drawPieLabel draws a data label for a pie/doughnut slice when requested.
func (r *renderer) drawPieLabel(ser *ChartSeries, cat string, percent float64, cx, cy, radius int, startAngle, endAngle float64) {
	if ser == nil || (!ser.ShowValue && !ser.ShowPercentage && !ser.ShowCategoryName) {
		return
	}
	parts := make([]string, 0, 2)
	if ser.ShowCategoryName {
		parts = append(parts, cat)
	}
	if ser.ShowValue {
		parts = append(parts, chartFormatNumber(ser.Values[cat]))
	}
	if ser.ShowPercentage {
		parts = append(parts, strconv.FormatFloat(percent, 'f', 1, 64)+"%")
	}
	if len(parts) == 0 {
		return
	}
	text := parts[0]
	for _, p := range parts[1:] {
		sep := ser.Separator
		if sep == "" {
			sep = ", "
		}
		text += sep + p
	}
	face := r.chartFace(ser.Font, text)
	fc := chartFontColor(ser.Font)
	mid := (startAngle + endAngle) / 2
	lx := cx + int(float64(radius)*0.65*math.Cos(mid))
	ly := cy + int(float64(radius)*0.65*math.Sin(mid))
	ascent := face.Metrics().Ascent.Ceil()
	descent := face.Metrics().Descent.Ceil()
	tw := chartTextWidth(face, text)
	r.drawChartText(text, face, fc, lx-tw/2, ly+(ascent+descent)/2-descent)
}

func (r *renderer) renderDoughnutChart(c *DoughnutChart, s *ChartShape, px, py, pw, ph int) {
	if len(c.Series) == 0 || len(c.Series[0].Categories) == 0 {
		return
	}
	palette := chartColors()
	ser := c.Series[0]

	vals := make([]float64, 0, len(ser.Categories))
	total := 0.0
	for _, cat := range ser.Categories {
		v := ser.Values[cat]
		if v > 0 {
			vals = append(vals, v)
			total += v
		} else {
			vals = append(vals, 0)
		}
	}
	if total == 0 {
		return
	}

	cx := px + pw/2
	cy := py + ph/2
	outerR := minInt(pw, ph)/2 - 2
	innerR := outerR * c.HoleSize / 100
	if outerR < 5 {
		return
	}

	startAngle := -math.Pi / 2
	for i, cat := range ser.Categories {
		v := vals[i]
		if v <= 0 {
			continue
		}
		sweep := 2 * math.Pi * v / total
		endAngle := startAngle + sweep
		sc := palette[i%len(palette)]
		r.fillDoughnutSlice(cx, cy, innerR, outerR, startAngle, endAngle, sc)
		r.drawPieLabel(ser, cat, v/total*100, cx, cy, outerR, startAngle, endAngle)
		startAngle = endAngle
	}
}

// fillPieSlice fills a pie slice using a scanline approach.
func (r *renderer) fillPieSlice(cx, cy, radius int, startAngle, endAngle float64, c color.RGBA) {
	r2 := radius * radius
	for dy := -radius; dy <= radius; dy++ {
		dy2 := dy * dy
		if dy2 > r2 {
			continue
		}
		maxDx := int(math.Sqrt(float64(r2 - dy2)))
		for dx := -maxDx; dx <= maxDx; dx++ {
			angle := math.Atan2(float64(dy), float64(dx))
			if angleInSweep(angle, startAngle, endAngle) {
				r.blendPixel(cx+dx, cy+dy, c)
			}
		}
	}
}

// fillDoughnutSlice fills a doughnut slice (an annulus segment).
func (r *renderer) fillDoughnutSlice(cx, cy, innerR, outerR int, startAngle, endAngle float64, c color.RGBA) {
	or2 := outerR * outerR
	ir2 := innerR * innerR
	for dy := -outerR; dy <= outerR; dy++ {
		dy2 := dy * dy
		if dy2 > or2 {
			continue
		}
		maxDx := int(math.Sqrt(float64(or2 - dy2)))
		for dx := -maxDx; dx <= maxDx; dx++ {
			d2 := dx*dx + dy2
			if d2 < ir2 {
				continue
			}
			angle := math.Atan2(float64(dy), float64(dx))
			if angleInSweep(angle, startAngle, endAngle) {
				r.blendPixel(cx+dx, cy+dy, c)
			}
		}
	}
}

// angleInSweep reports whether angle lies within the sweep from start to end.
func angleInSweep(angle, start, end float64) bool {
	norm := func(a float64) float64 {
		for a < 0 {
			a += 2 * math.Pi
		}
		for a >= 2*math.Pi {
			a -= 2 * math.Pi
		}
		return a
	}
	a := norm(angle)
	s := norm(start)
	e := norm(end)
	if s <= e {
		return a >= s && a <= e
	}
	return a >= s || a <= e
}

// --- area chart ---

func (r *renderer) renderAreaChart(c *AreaChart, s *ChartShape, px, py, pw, ph int) {
	if len(c.Series) == 0 || pw <= 0 || ph <= 0 {
		return
	}
	cats := c.Series[0].Categories
	if len(cats) == 0 {
		return
	}
	palette := chartColors()

	minV, maxV := chartSeriesRange(c.Series)
	sc := chartComputeScale(minV, maxV, s.plotArea.GetAxisY(), true)
	plot := r.chartPlotAreaFor(s, px, py, pw, ph, sc, cats, false)
	r.drawChartAxes(s, plot)

	nPts := len(cats)
	for si, ser := range c.Series {
		sc2 := getSeriesColor(ser, si, palette)
		fillC := color.RGBA{R: sc2.R, G: sc2.G, B: sc2.B, A: 128}
		pts := make([]fpoint, 0, nPts+2)
		for i, cat := range cats {
			pts = append(pts, fpoint{float64(chartPointX(plot, i, nPts)), float64(plot.valueY(ser.Values[cat]))})
		}
		baseline := float64(plot.valueY(sc.min))
		pts = append(pts, fpoint{pts[len(pts)-1].x, baseline})
		pts = append(pts, fpoint{pts[0].x, baseline})
		r.fillPolygon(pts, fillC)

		for i := 1; i < nPts; i++ {
			r.drawLineAA(int(pts[i-1].x), int(pts[i-1].y), int(pts[i].x), int(pts[i].y), sc2, 2)
		}
	}
}

// --- scatter chart ---

func (r *renderer) renderScatterChart(c *ScatterChart, s *ChartShape, px, py, pw, ph int) {
	if len(c.Series) == 0 || pw <= 0 || ph <= 0 {
		return
	}
	palette := chartColors()

	minV, maxV := chartSeriesRange(c.Series)
	sc := chartComputeScale(minV, maxV, s.plotArea.GetAxisY(), false)
	plot := r.chartPlotAreaFor(s, px, py, pw, ph, sc, c.Series[0].Categories, false)
	// The X axis is numeric for scatter charts. Only pin a real scale when
	// the author constrained it (log base / explicit bounds); otherwise the
	// legacy even-spacing mapping keeps earlier renders stable.
	if axX := s.plotArea.GetAxisX(); axX != nil && (axX.LogBase > 1 || axX.MinBounds != nil || axX.MaxBounds != nil) {
		minX, maxX := math.MaxFloat64, -math.MaxFloat64
		for _, cat := range c.Series[0].Categories {
			if f, err := strconv.ParseFloat(cat, 64); err == nil {
				if f < minX {
					minX = f
				}
				if f > maxX {
					maxX = f
				}
			}
		}
		if minX <= maxX {
			scX := chartComputeScale(minX, maxX, axX, false)
			plot.scaleX = &scX
		}
	}
	r.drawChartAxes(s, plot)

	for si, ser := range c.Series {
		sc2 := getSeriesColor(ser, si, palette)
		// Scatter charts keep the legacy hairline default: the old deck's
		// scatter series render thin in the COM golds too, and 2.25pt has
		// only been verified for the line-chart family (deck 00022823
		// slide05). Revisit with a COM variant experiment before changing.
		lw := r.chartSeriesLineWidth(ser, 0)
		if lw <= 0 {
			if ser != nil && ser.Outline != nil && ser.Outline.Width > 0 {
				lw = maxInt(int(float64(ser.Outline.Width)*12700.0*r.scaleX), 1)
			} else {
				lw = 1
			}
		}
		n := len(ser.Categories)
		var runXs, runYs []int
		flush := func() {
			if len(runXs) > 1 {
				r.drawChartLineDashed(runXs, runYs, sc2, lw, ser.LineDash)
			}
			for i := range runXs {
				r.drawChartMarker(ser, runXs[i], runYs[i], sc2)
			}
			runXs, runYs = runXs[:0], runYs[:0]
		}
		for i := 0; i < n; i++ {
			cat := ser.Categories[i]
			v := ser.Values[cat]
			if math.IsNaN(v) {
				flush()
				continue
			}
			runXs = append(runXs, r.scatterX(plot, i, n, cat))
			runYs = append(runYs, plot.valueY(v))
		}
		flush()
	}
}

// scatterX maps a scatter point to a pixel X. When the axis carries a real
// scale (log base or explicit bounds) it wins; otherwise numeric categories
// plot on a linear axis and anything else falls back to even spacing.
func (r *renderer) scatterX(plot chartPlotArea, i, n int, cat string) int {
	if plot.scaleX != nil {
		if v, err := strconv.ParseFloat(cat, 64); err == nil {
			return plot.x + int(plot.scaleX.ratioClamped(v)*float64(plot.w))
		}
	}
	if v, err := strconv.ParseFloat(cat, 64); err == nil {
		minX, maxX := math.MaxFloat64, -math.MaxFloat64
		for _, name := range plot.cats {
			if f, err := strconv.ParseFloat(name, 64); err == nil {
				if f < minX {
					minX = f
				}
				if f > maxX {
					maxX = f
				}
			}
		}
		if minX <= maxX {
			if maxX-minX < 1e-12 {
				return plot.x + plot.w/2
			}
			ratio := (v - minX) / (maxX - minX)
			return plot.x + int(ratio*float64(plot.w))
		}
	}
	return chartPointX(plot, i, n)
}

// --- radar chart ---

func (r *renderer) renderRadarChart(c *RadarChart, s *ChartShape, px, py, pw, ph int) {
	if len(c.Series) == 0 || pw <= 0 || ph <= 0 {
		return
	}
	palette := chartColors()

	maxVal := 0.0
	for _, ser := range c.Series {
		for _, v := range ser.Values {
			if v > maxVal {
				maxVal = v
			}
		}
	}
	if maxVal <= 0 {
		maxVal = 1
	}

	cats := c.Series[0].Categories
	nCats := len(cats)
	if nCats == 0 {
		return
	}

	cx := px + pw/2
	cy := py + ph/2

	// Category labels are drawn just outside the outermost ring, so the radius
	// has to leave room for them. Sizing the radius from pw/ph alone lets the
	// labels spill out of the plot rect and collide with the chart title above
	// and the legend below.
	axX := s.plotArea.GetAxisX()
	catFace := r.chartFaceForCats(axisFont(axX), cats)
	catColor := chartFontColor(axisFont(axX))
	const labelGap = 8
	labelInsetX, labelInsetY := 0, 0
	if axX == nil || axX.Visible {
		maxTW := 0
		for _, cat := range cats {
			if tw := chartTextWidth(catFace, cat); tw > maxTW {
				maxTW = tw
			}
		}
		labelInsetX = labelGap + maxTW/2 + 2
		labelInsetY = labelGap + r.chartLineHeight(catFace)
	}

	radius := minInt(pw/2-labelInsetX, ph/2-labelInsetY)
	if radius < 5 {
		return
	}

	// Grid rings.
	gridColor := color.RGBA{R: 217, G: 217, B: 217, A: 255}
	rings := 5
	for ring := 1; ring <= rings; ring++ {
		rr := float64(radius) * float64(ring) / float64(rings)
		pts := make([]fpoint, nCats)
		for i := 0; i < nCats; i++ {
			angle := 2*math.Pi*float64(i)/float64(nCats) - math.Pi/2
			pts[i] = fpoint{float64(cx) + rr*math.Cos(angle), float64(cy) + rr*math.Sin(angle)}
		}
		for i := 0; i < nCats; i++ {
			j := (i + 1) % nCats
			r.drawLineAA(int(pts[i].x), int(pts[i].y), int(pts[j].x), int(pts[j].y), gridColor, 1)
		}
	}
	// Spokes.
	for i := 0; i < nCats; i++ {
		angle := 2*math.Pi*float64(i)/float64(nCats) - math.Pi/2
		ex := cx + int(float64(radius)*math.Cos(angle))
		ey := cy + int(float64(radius)*math.Sin(angle))
		r.drawLineAA(cx, cy, ex, ey, gridColor, 1)
	}

	// Category labels around the perimeter.
	if axX == nil || axX.Visible {
		ascent := catFace.Metrics().Ascent.Ceil()
		descent := catFace.Metrics().Descent.Ceil()
		for i, cat := range cats {
			angle := 2*math.Pi*float64(i)/float64(nCats) - math.Pi/2
			lx := cx + int(float64(radius+labelGap)*math.Cos(angle))
			ly := cy + int(float64(radius+labelGap)*math.Sin(angle))
			tw := chartTextWidth(catFace, cat)
			r.drawChartText(cat, catFace, catColor, lx-tw/2, ly+(ascent+descent)/2-descent)
		}
	}

	// Series polygons.
	for si, ser := range c.Series {
		sc2 := getSeriesColor(ser, si, palette)
		pts := make([]fpoint, nCats)
		for i, cat := range cats {
			v := ser.Values[cat]
			angle := 2*math.Pi*float64(i)/float64(nCats) - math.Pi/2
			dist := float64(radius) * v / maxVal
			pts[i] = fpoint{
				x: float64(cx) + dist*math.Cos(angle),
				y: float64(cy) + dist*math.Sin(angle),
			}
		}
		fillC := color.RGBA{R: sc2.R, G: sc2.G, B: sc2.B, A: 64}
		r.fillPolygon(pts, fillC)
		for i := 0; i < nCats; i++ {
			j := (i + 1) % nCats
			r.drawLineAA(int(pts[i].x), int(pts[i].y), int(pts[j].x), int(pts[j].y), sc2, 2)
		}
	}
}

// --- legend ---

// chartLegendEntries returns the legend labels and swatch colours for a chart.
func (r *renderer) chartLegendEntries(s *ChartShape) ([]string, []color.RGBA) {
	ct := s.plotArea.GetType()
	if ct == nil {
		return nil, nil
	}
	palette := chartColors()
	var names []string
	var colors []color.RGBA

	bySeries := func(series []*ChartSeries) {
		for i, ser := range series {
			names = append(names, ser.Title)
			colors = append(colors, getSeriesColor(ser, i, palette))
		}
	}
	byCategory := func(series []*ChartSeries) {
		if len(series) == 0 {
			return
		}
		for i, cat := range series[0].Categories {
			names = append(names, cat)
			colors = append(colors, palette[i%len(palette)])
		}
	}

	switch c := ct.(type) {
	case *BarChart:
		bySeries(c.Series)
	case *Bar3DChart:
		bySeries(c.Series)
	case *LineChart:
		bySeries(c.Series)
	case *AreaChart:
		bySeries(c.Series)
	case *ScatterChart:
		bySeries(c.Series)
	case *RadarChart:
		bySeries(c.Series)
	case *PieChart:
		byCategory(c.Series)
	case *Pie3DChart:
		byCategory(c.Series)
	case *DoughnutChart:
		byCategory(c.Series)
	}
	return names, colors
}

// chartLegendWidth is the horizontal space needed for a left/right legend.
func (r *renderer) chartLegendWidth(s *ChartShape, maxW int) int {
	names, _ := r.chartLegendEntries(s)
	face := r.chartFace(s.legend.Font, strings.Join(names, ""))
	w := 0
	for _, n := range names {
		if tw := chartTextWidth(face, n); tw > w {
			w = tw
		}
	}
	return minInt(w+18, maxW)
}

// chartLegendHeight is the vertical space needed for a top/bottom legend,
// accounting for wrapping into multiple rows.
func (r *renderer) chartLegendHeight(s *ChartShape, x, w int) int {
	names, _ := r.chartLegendEntries(s)
	if len(names) == 0 {
		return 0
	}
	face := r.chartFace(s.legend.Font, strings.Join(names, ""))
	lineH := r.chartLineHeight(face) + 4
	avail := w - 8
	rowW := 0
	rows := 1
	for _, n := range names {
		entryW := 14 + chartTextWidth(face, n) + 12
		if rowW > 0 && rowW+entryW > avail {
			rows++
			rowW = entryW
		} else {
			rowW += entryW
		}
	}
	return rows*lineH + 4
}

func (r *renderer) renderChartLegend(s *ChartShape, x, y, w, h, titleH int, pos LegendPosition) {
	names, colors := r.chartLegendEntries(s)
	if len(names) == 0 {
		return
	}
	face := r.chartFace(s.legend.Font, strings.Join(names, ""))
	fc := chartFontColor(s.legend.Font)
	lineH := r.chartLineHeight(face) + 4
	box := 9

	drawEntry := func(nx, ny int, name string, c color.RGBA) int {
		r.fillRectFast(image.Rect(nx, ny+(lineH-box)/2, nx+box, ny+(lineH-box)/2+box), c)
		baseline := ny + lineH/2 + face.Metrics().Ascent.Ceil()/2
		r.drawChartText(name, face, fc, nx+box+5, baseline)
		return 14 + chartTextWidth(face, name) + 12
	}

	switch pos {
	case LegendLeft, LegendRight:
		legendW := r.chartLegendWidth(s, w/2)
		lx := x + 4
		if pos == LegendRight {
			lx = x + w - legendW - 4
		}
		innerX := lx + 6
		ly := y + titleH + 4
		if innerX+legendW > x+w {
			innerX = x + 4
		}
		for i, name := range names {
			drawEntry(innerX, ly+i*lineH, name, colors[i])
		}
	default: // bottom / top / topRight
		legendH := r.chartLegendHeight(s, x, w)
		ly := y + h - legendH
		if pos == LegendTop || pos == LegendTopRight {
			ly = y + titleH
		}
		if ly < y {
			ly = y
		}
		avail := w - 8
		// Pre-compute rows so a row can be centred.
		type entry struct {
			name string
			c    color.RGBA
			w    int
		}
		var rows [][]entry
		var cur []entry
		curW := 0
		for i, name := range names {
			ew := 14 + chartTextWidth(face, name) + 12
			if curW > 0 && curW+ew > avail {
				rows = append(rows, cur)
				cur = nil
				curW = 0
			}
			cur = append(cur, entry{name, colors[i], ew})
			curW += ew
		}
		if len(cur) > 0 {
			rows = append(rows, cur)
		}
		for ri, row := range rows {
			total := 0
			for _, e := range row {
				total += e.w
			}
			nx := x + 4
			if pos == LegendTopRight {
				nx = x + w - 4 - total
			} else {
				nx = x + (w-total)/2
			}
			ny := ly + ri*lineH
			for _, e := range row {
				nx += drawEntry(nx, ny, e.name, e.c)
			}
		}
	}
}

// --- date axis & dash helpers ---

// chartIsDateFormat reports whether a number format code renders dates
// (month and/or year tokens like "mmm yyyy"); such axes plot date serials.
func chartIsDateFormat(format string) bool {
	f := strings.ToLower(format)
	return strings.Contains(f, "mmm") || strings.Contains(f, "yy")
}

// chartFormatDate renders an Excel date serial with a mmm/mmmm/yy/yyyy
// format code. Serials in chart caches are 1900-epoch even when the part
// declares <c:date1904/> — deck 00022823 chart1 carries the flag yet its
// serial 37164 renders as "Sep 2001" in COM, which is the 1900 system.
func chartFormatDate(serial float64, format string) string {
	day := time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC)
	t := day.AddDate(0, 0, int(math.Round(serial)))
	f := strings.ReplaceAll(format, "\\", "")
	var b strings.Builder
	runes := []rune(f)
	for i := 0; i < len(runes); i++ {
		switch c := runes[i]; c {
		case 'm', 'M':
			n := 1
			for i+n < len(runes) && (runes[i+n] == 'm' || runes[i+n] == 'M') {
				n++
			}
			i += n - 1
			switch n {
			case 4:
				b.WriteString(t.Format("January"))
			case 3:
				b.WriteString(t.Format("Jan"))
			case 2:
				b.WriteString(fmt.Sprintf("%02d", int(t.Month())))
			default:
				b.WriteString(strconv.Itoa(int(t.Month())))
			}
		case 'y', 'Y':
			n := 1
			for i+n < len(runes) && (runes[i+n] == 'y' || runes[i+n] == 'Y') {
				n++
			}
			i += n - 1
			y := t.Year()
			if n >= 3 {
				b.WriteString(strconv.Itoa(y))
			} else {
				b.WriteString(fmt.Sprintf("%02d", y%100))
			}
		default:
			b.WriteRune(c)
		}
	}
	return b.String()
}

// chartSerialToTime / chartTimeToSerial convert between Excel date serials
// and time.Time over the 1899-12-30 workbook epoch (see chartFormatDate).
func chartSerialToTime(s float64) time.Time {
	return time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC).AddDate(0, 0, int(math.Round(s)))
}

func chartTimeToSerial(t time.Time) float64 {
	return t.Sub(time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC)).Hours() / 24
}

// chartDashPattern maps an <a:prstDash> preset to on/off run lengths in
// multiples of the stroke width. nil means solid.
func chartDashPattern(name string) []float64 {
	switch name {
	case "dash":
		return []float64{4, 3}
	case "sysDash":
		return []float64{4, 1}
	case "dot":
		return []float64{1, 2}
	case "sysDot":
		return []float64{1, 1}
	case "lgDash":
		return []float64{8, 4}
	case "lgDashDot":
		return []float64{8, 2, 1, 2}
	case "lgDashDotDot":
		return []float64{8, 2, 1, 2, 1, 2}
	case "dashDot":
		return []float64{6, 2, 1, 2}
	}
	return nil
}

// drawChartLineDashed strokes a polyline, breaking the stroke into the
// preset's on/off runs (scaled by the line width) along its length.
func (r *renderer) drawChartLineDashed(xs, ys []int, c color.RGBA, width int, dash string) {
	pat := chartDashPattern(dash)
	if pat == nil {
		for i := 1; i < len(xs); i++ {
			r.drawLineAA(xs[i-1], ys[i-1], xs[i], ys[i], c, width)
		}
		return
	}
	w := float64(maxInt(width, 1))
	runs := make([]float64, len(pat))
	for i, p := range pat {
		runs[i] = math.Max(p*w, 1.5)
	}
	// Walk the polyline, consuming pattern runs across segment boundaries.
	on := true
	consumed := runs[0] // remaining length of the current run
	rx0, ry0 := float64(xs[0]), float64(ys[0])
	segStartX, segStartY := rx0, ry0
	pi := 0
	for i := 1; i < len(xs); i++ {
		sx, sy := float64(xs[i]), float64(ys[i])
		segLen := math.Hypot(sx-rx0, sy-ry0)
		if segLen <= 0 {
			continue
		}
		pos := 0.0
		for pos < segLen {
			run := consumed
			if pos+run > segLen {
				run = segLen - pos
			}
			ex := rx0 + (sx-rx0)*(pos+run)/segLen
			ey := ry0 + (sy-ry0)*(pos+run)/segLen
			if on {
				r.drawLineAA(int(segStartX), int(segStartY), int(ex), int(ey), c, width)
			}
			segStartX, segStartY = ex, ey
			pos += run
			consumed -= run
			if consumed <= 1e-9 {
				pi = (pi + 1) % len(runs)
				consumed = runs[pi]
				on = pi%2 == 0
			}
		}
		rx0, ry0 = sx, sy
	}
}

// drawChartRotatedLabel draws text rotated 270 degrees (reading
// bottom-to-top) with the column centred on cx below the axis - how
// PowerPoint fits long date labels along a category axis.
func (r *renderer) drawChartRotatedLabel(text string, face font.Face, c color.RGBA, cx, top int) {
	if text == "" || face == nil {
		return
	}
	tw := chartTextWidth(face, text)
	m := face.Metrics()
	lineH := (m.Ascent + m.Descent).Ceil()
	h := tw + lineH
	// The buffer is the destination region transposed: text is drawn
	// horizontally across its full length, then rotated into the narrow
	// column.
	tmp := image.NewRGBA(image.Rect(0, 0, h, lineH))
	tr := r.subRenderer(tmp)
	tr.drawChartTextCentered(text, face, c, 0, 0, h, lineH)
	rotateAndComposite(r.img, tmp, cx-lineH/2, top, lineH, h, 270)
}
