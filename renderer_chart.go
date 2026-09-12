package gopresentation

import (
	"image"
	"image/color"
	"math"
	"strconv"
	"strings"

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
}

func (cs chartScale) span() float64 {
	if cs.max <= cs.min {
		return 1
	}
	return cs.max - cs.min
}

// ratio maps a data value to a 0..1 position along the value axis.
func (cs chartScale) ratio(v float64) float64 { return (v - cs.min) / cs.span() }

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

// --- font / text helpers ---

// chartFace resolves a font for chart text. A nil or unset font falls back to
// the renderer default (10pt).
func (r *renderer) chartFace(f *Font) font.Face {
	if f == nil {
		f = NewFont()
	}
	if f.Size <= 0 {
		clone := *f
		clone.Size = 10
		f = &clone
	}
	return r.getFace(f)
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
		face := r.chartFace(s.title.Font)
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
	valueFace := r.chartFace(axisFont(axV))
	catFace := r.chartFace(axisFont(axX))

	left, top, right, bottom := 3, 2, 4, 2
	if axV == nil || axV.Visible {
		if categoriesOnY {
			bottom += r.chartLineHeight(valueFace) + 2
		} else {
			lw := 0
			for _, t := range sc.ticks() {
				if tw := chartTextWidth(valueFace, chartFormatNumber(t)); tw > lw {
					lw = tw
				}
			}
			left = lw + 5
		}
	}
	if axX == nil || axX.Visible {
		if categoriesOnY {
			left += chartMaxLabelWidth(catFace, cats, pw/2) + 5
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

	// Major gridlines. PowerPoint shows them by default on the value axis.
	gridColor := color.RGBA{R: 217, G: 217, B: 217, A: 255}
	gridW := 1
	if axV != nil && axV.MajorGridlines != nil {
		gridColor = argbToRGBA(axV.MajorGridlines.Color)
		gridW = maxInt(axV.MajorGridlines.Width, 1)
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

	// Category axis line.
	axisLine := color.RGBA{R: 191, G: 191, B: 191, A: 255}
	if p.categoriesOnY {
		r.drawLineAA(p.x, p.y, p.x, p.y+p.h, axisLine, 1)
	} else {
		r.drawLineAA(p.x, p.y+p.h, p.x+p.w, p.y+p.h, axisLine, 1)
	}

	// Value labels.
	if axV == nil || axV.Visible {
		face := r.chartFace(axisFont(axV))
		fc := chartFontColor(axisFont(axV))
		ascent := face.Metrics().Ascent.Ceil()
		descent := face.Metrics().Descent.Ceil()
		for _, t := range p.scale.ticks() {
			label := chartFormatNumber(t)
			if p.categoriesOnY {
				tw := chartTextWidth(face, label)
				r.drawChartText(label, face, fc, p.valueX(t)-tw/2, p.y+p.h+2+ascent)
			} else {
				tw := chartTextWidth(face, label)
				gy := p.valueY(t)
				r.drawChartText(label, face, fc, p.x-5-tw, gy+(ascent+descent)/2-descent)
			}
		}
	}

	// Category labels.
	if (axX == nil || axX.Visible) && len(p.cats) > 0 {
		face := r.chartFace(axisFont(axX))
		fc := chartFontColor(axisFont(axX))
		ascent := face.Metrics().Ascent.Ceil()
		descent := face.Metrics().Descent.Ceil()
		if p.categoriesOnY {
			rowH := float64(p.h) / float64(len(p.cats))
			for i, cat := range p.cats {
				cy := p.y + int((float64(i)+0.5)*rowH)
				tw := chartTextWidth(face, cat)
				r.drawChartText(cat, face, fc, p.x-5-tw, cy+(ascent+descent)/2-descent)
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
func (r *renderer) drawChartAxisTitles(s *ChartShape, p chartPlotArea) {
	axV := s.plotArea.GetAxisY()
	axX := s.plotArea.GetAxisX()

	if axV != nil && axV.Visible && strings.TrimSpace(axV.Title) != "" {
		face := r.chartFace(axV.Font)
		fc := chartFontColor(axV.Font)
		if p.categoriesOnY {
			// Value axis runs along the bottom.
			h := r.chartLineHeight(face)
			y := p.oy + p.oh - h
			if y < p.y+p.h {
				y = p.y + p.h
			}
			r.drawChartTextCentered(axV.Title, face, fc, p.x, y, p.w, h)
		} else if w := p.x - p.ox; w > 0 {
			r.drawChartVerticalCentered(axV.Title, face, fc, p.ox, p.y, w, p.h)
		}
	}
	if axX != nil && axX.Visible && strings.TrimSpace(axX.Title) != "" {
		face := r.chartFace(axX.Font)
		fc := chartFontColor(axX.Font)
		if p.categoriesOnY {
			if w := p.x - p.ox; w > 0 {
				r.drawChartVerticalCentered(axX.Title, face, fc, p.ox, p.y, w, p.h)
			}
		} else {
			h := r.chartLineHeight(face)
			y := p.oy + p.oh - h
			if y < p.y+p.h {
				y = p.y + p.h
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
	sc := chartComputeScale(minV, maxV, s.plotArea.GetAxisY(), !percent)
	plot := r.chartPlotAreaFor(s, px, py, pw, ph, sc, cats, horizontal)
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
			rowY := float64(plot.y) + float64(ci)*rowH + (rowH-clusterH)/2
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
						getSeriesColor(ser, si, palette))
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
						getSeriesColor(ser, si, palette))
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
						getSeriesColor(ser, si, palette))
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
						getSeriesColor(ser, si, palette))
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
	return fallback
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
	r.drawChartAxes(s, plot)

	nPts := len(cats)
	for si, ser := range c.Series {
		sc2 := getSeriesColor(ser, si, palette)
		lw := r.chartSeriesLineWidth(ser, 2)
		xs := make([]int, nPts)
		ys := make([]int, nPts)
		for i, cat := range cats {
			xs[i] = chartPointX(plot, i, nPts)
			ys[i] = plot.valueY(ser.Values[cat])
		}
		if c.IsSmooth && nPts >= 3 {
			r.drawSmoothCurve(xs, ys, sc2, lw)
		} else {
			for i := 1; i < nPts; i++ {
				r.drawLineAA(xs[i-1], ys[i-1], xs[i], ys[i], sc2, lw)
			}
		}
		// Markers
		for i := range xs {
			r.drawChartMarker(ser, xs[i], ys[i], sc2)
		}
	}
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
	rad := maxInt(int(float64(size)*r.scaleX/2), 2)
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
	face := r.chartFace(ser.Font)
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
	r.drawChartAxes(s, plot)

	for si, ser := range c.Series {
		sc2 := getSeriesColor(ser, si, palette)
		lw := r.chartSeriesLineWidth(ser, 1)
		n := len(ser.Categories)
		xs := make([]int, n)
		ys := make([]int, n)
		for i, cat := range ser.Categories {
			xs[i] = r.scatterX(plot, i, n, cat)
			ys[i] = plot.valueY(ser.Values[cat])
		}
		if c.IsSmooth && n >= 3 {
			r.drawSmoothCurve(xs, ys, sc2, lw)
		} else if n > 1 {
			for i := 1; i < n; i++ {
				r.drawLineAA(xs[i-1], ys[i-1], xs[i], ys[i], sc2, lw)
			}
		}
		for i := range xs {
			r.drawChartMarker(ser, xs[i], ys[i], sc2)
		}
	}
}

// scatterX maps a scatter point to a pixel X. Numeric category values are
// plotted on a linear axis; anything else falls back to even spacing.
func (r *renderer) scatterX(plot chartPlotArea, i, n int, cat string) int {
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
	catFace := r.chartFace(axisFont(axX))
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
	face := r.chartFace(s.legend.Font)
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
	face := r.chartFace(s.legend.Font)
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
	face := r.chartFace(s.legend.Font)
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
