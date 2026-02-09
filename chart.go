package gopresentation

import "sort"

// ChartShape represents a chart embedded in a slide.
type ChartShape struct {
	BaseShape
	title       *ChartTitle
	plotArea    *PlotArea
	legend      *ChartLegend
	view3D      *View3D
	displayBlankAs string
}

// Chart display blank constants.
const (
	ChartBlankAsGap  = "gap"
	ChartBlankAsZero = "zero"
	ChartBlankAsSpan = "span"
)

func (c *ChartShape) GetType() ShapeType { return ShapeTypeChart }

// NewChartShape creates a new chart shape.
func NewChartShape() *ChartShape {
	return &ChartShape{
		title:          NewChartTitle(),
		plotArea:       NewPlotArea(),
		legend:         NewChartLegend(),
		view3D:         NewView3D(),
		displayBlankAs: ChartBlankAsZero,
	}
}

// GetTitle returns the chart title.
func (c *ChartShape) GetTitle() *ChartTitle { return c.title }

// GetPlotArea returns the plot area.
func (c *ChartShape) GetPlotArea() *PlotArea { return c.plotArea }

// GetLegend returns the chart legend.
func (c *ChartShape) GetLegend() *ChartLegend { return c.legend }

// GetView3D returns the 3D view settings.
func (c *ChartShape) GetView3D() *View3D { return c.view3D }

// SetDisplayBlankAs sets how blank values are displayed.
func (c *ChartShape) SetDisplayBlankAs(mode string) { c.displayBlankAs = mode }

// GetDisplayBlankAs returns how blank values are displayed.
func (c *ChartShape) GetDisplayBlankAs() string { return c.displayBlankAs }

// ChartTitle represents a chart title.
type ChartTitle struct {
	Text    string
	Visible bool
	Font    *Font
}

// NewChartTitle creates a new chart title.
func NewChartTitle() *ChartTitle {
	return &ChartTitle{
		Visible: true,
		Font:    NewFont(),
	}
}

// SetText sets the title text.
func (ct *ChartTitle) SetText(text string) *ChartTitle {
	ct.Text = text
	return ct
}

// SetVisible sets the title visibility.
func (ct *ChartTitle) SetVisible(v bool) *ChartTitle {
	ct.Visible = v
	return ct
}

// PlotArea represents the chart plot area.
type PlotArea struct {
	chartType ChartType
	axisX     *ChartAxis
	axisY     *ChartAxis
}

// NewPlotArea creates a new plot area.
func NewPlotArea() *PlotArea {
	return &PlotArea{
		axisX: NewChartAxis(),
		axisY: NewChartAxis(),
	}
}

// SetType sets the chart type.
func (pa *PlotArea) SetType(ct ChartType) { pa.chartType = ct }

// GetType returns the chart type.
func (pa *PlotArea) GetType() ChartType { return pa.chartType }

// GetAxisX returns the X axis.
func (pa *PlotArea) GetAxisX() *ChartAxis { return pa.axisX }

// GetAxisY returns the Y axis.
func (pa *PlotArea) GetAxisY() *ChartAxis { return pa.axisY }

// ChartAxis represents a chart axis.
type ChartAxis struct {
	Title         string
	TitleRotation int
	Visible       bool
	MinBounds     *float64
	MaxBounds     *float64
	MinorUnit     *float64
	MajorUnit     *float64
	CrossesAt     string
	ReversedOrder bool
	Font          *Font
	MajorGridlines *Gridlines
	MinorGridlines *Gridlines
	MajorTickMark  string
	MinorTickMark  string
	TickLabelPos   string
	OutlineWidth   int
	OutlineColor   Color
}

// Axis crossing constants.
const (
	AxisCrossesAuto = "autoZero"
	AxisCrossesMin  = "min"
	AxisCrossesMax  = "max"
)

// Tick mark constants.
const (
	TickMarkNone    = "none"
	TickMarkInside  = "in"
	TickMarkOutside = "out"
	TickMarkCross   = "cross"
)

// Tick label position constants.
const (
	TickLabelPosNextTo = "nextTo"
	TickLabelPosHigh   = "high"
	TickLabelPosLow    = "low"
)

// NewChartAxis creates a new chart axis.
func NewChartAxis() *ChartAxis {
	return &ChartAxis{
		Visible:       true,
		CrossesAt:     AxisCrossesAuto,
		Font:          NewFont(),
		MajorTickMark: TickMarkNone,
		MinorTickMark: TickMarkNone,
		TickLabelPos:  TickLabelPosNextTo,
	}
}

// SetTitle sets the axis title.
func (a *ChartAxis) SetTitle(title string) *ChartAxis {
	a.Title = title
	return a
}

// SetTitleRotation sets the axis title rotation in degrees.
func (a *ChartAxis) SetTitleRotation(deg int) *ChartAxis {
	a.TitleRotation = deg
	return a
}

// SetVisible sets axis visibility.
func (a *ChartAxis) SetVisible(v bool) *ChartAxis {
	a.Visible = v
	return a
}

// SetMinBounds sets the minimum bounds.
func (a *ChartAxis) SetMinBounds(v float64) *ChartAxis {
	a.MinBounds = &v
	return a
}

// ClearMinBounds clears the minimum bounds.
func (a *ChartAxis) ClearMinBounds() *ChartAxis {
	a.MinBounds = nil
	return a
}

// SetMaxBounds sets the maximum bounds.
func (a *ChartAxis) SetMaxBounds(v float64) *ChartAxis {
	a.MaxBounds = &v
	return a
}

// ClearMaxBounds clears the maximum bounds.
func (a *ChartAxis) ClearMaxBounds() *ChartAxis {
	a.MaxBounds = nil
	return a
}

// SetMinorUnit sets the minor unit.
func (a *ChartAxis) SetMinorUnit(v float64) *ChartAxis {
	a.MinorUnit = &v
	return a
}

// SetMajorUnit sets the major unit.
func (a *ChartAxis) SetMajorUnit(v float64) *ChartAxis {
	a.MajorUnit = &v
	return a
}

// SetCrossesAt sets where the axis crosses.
func (a *ChartAxis) SetCrossesAt(v string) *ChartAxis {
	a.CrossesAt = v
	return a
}

// SetReversedOrder sets whether the axis is reversed.
func (a *ChartAxis) SetReversedOrder(v bool) *ChartAxis {
	a.ReversedOrder = v
	return a
}

// SetMajorGridlines sets the major gridlines.
func (a *ChartAxis) SetMajorGridlines(g *Gridlines) *ChartAxis {
	a.MajorGridlines = g
	return a
}

// SetMinorGridlines sets the minor gridlines.
func (a *ChartAxis) SetMinorGridlines(g *Gridlines) *ChartAxis {
	a.MinorGridlines = g
	return a
}

// SetMajorTickMark sets the major tick mark style.
func (a *ChartAxis) SetMajorTickMark(v string) *ChartAxis {
	a.MajorTickMark = v
	return a
}

// SetMinorTickMark sets the minor tick mark style.
func (a *ChartAxis) SetMinorTickMark(v string) *ChartAxis {
	a.MinorTickMark = v
	return a
}

// SetTickLabelPosition sets the tick label position.
func (a *ChartAxis) SetTickLabelPosition(v string) *ChartAxis {
	a.TickLabelPos = v
	return a
}

// Gridlines represents chart gridlines.
type Gridlines struct {
	Width int
	Color Color
}

// NewGridlines creates new gridlines.
func NewGridlines() *Gridlines {
	return &Gridlines{
		Width: 1,
		Color: ColorBlack,
	}
}

// ChartLegend represents a chart legend.
type ChartLegend struct {
	Visible  bool
	Position LegendPosition
	Font     *Font
}

// LegendPosition represents the legend position.
type LegendPosition string

const (
	LegendBottom LegendPosition = "b"
	LegendTop    LegendPosition = "t"
	LegendLeft   LegendPosition = "l"
	LegendRight  LegendPosition = "r"
	LegendTopRight LegendPosition = "tr"
)

// NewChartLegend creates a new chart legend.
func NewChartLegend() *ChartLegend {
	return &ChartLegend{
		Visible:  true,
		Position: LegendBottom,
		Font:     NewFont(),
	}
}

// View3D represents 3D view settings.
type View3D struct {
	RotX          int
	RotY          int
	DepthPercent  int
	HeightPercent *int
	RightAngleAxes bool
}

// NewView3D creates new 3D view settings.
func NewView3D() *View3D {
	hp := 100
	return &View3D{
		RotX:           15,
		RotY:           20,
		DepthPercent:   100,
		HeightPercent:  &hp,
		RightAngleAxes: true,
	}
}

// SetHeightPercent sets the height percent. Pass nil to enable autoscale.
func (v *View3D) SetHeightPercent(hp *int) { v.HeightPercent = hp }

// --- Chart Types ---

// ChartType is the interface for chart types.
type ChartType interface {
	GetChartTypeName() string
}

// ChartSeries represents a data series in a chart.
type ChartSeries struct {
	Title             string
	Values            map[string]float64 // category -> value
	Categories        []string           // ordered category names
	FillColor         Color
	ShowCategoryName  bool
	ShowLegendKey     bool
	ShowPercentage    bool
	ShowSeriesName    bool
	ShowValue         bool
	Separator         string
	LabelPosition     string
	Font              *Font
	Outline           *SeriesOutline
	Marker            *SeriesMarker
}

// Series label position constants.
const (
	LabelInsideEnd  = "inEnd"
	LabelOutsideEnd = "outEnd"
	LabelCenter     = "ctr"
	LabelInsideBase = "inBase"
	LabelBestFit    = "bestFit"
)

// NewChartSeries creates a new chart series.
// Note: map iteration order is non-deterministic in Go, so category order may vary.
// Use NewChartSeriesOrdered for deterministic ordering.
func NewChartSeries(title string, data map[string]float64) *ChartSeries {
	cats := make([]string, 0, len(data))
	for k := range data {
		cats = append(cats, k)
	}
	// Sort for deterministic output
	sort.Strings(cats)
	return &ChartSeries{
		Title:      title,
		Values:     data,
		Categories: cats,
		Font:       NewFont(),
		Separator:  ",",
	}
}

// NewChartSeriesOrdered creates a series with ordered categories.
// If len(values) < len(categories), missing values default to 0.
// Extra values beyond len(categories) are ignored.
func NewChartSeriesOrdered(title string, categories []string, values []float64) *ChartSeries {
	data := make(map[string]float64, len(categories))
	for i, cat := range categories {
		if i < len(values) {
			data[cat] = values[i]
		} else {
			data[cat] = 0
		}
	}
	return &ChartSeries{
		Title:      title,
		Values:     data,
		Categories: categories,
		Font:       NewFont(),
		Separator:  ",",
	}
}

// SetFillColor sets the series fill color.
func (s *ChartSeries) SetFillColor(c Color) *ChartSeries {
	s.FillColor = c
	return s
}

// SetLabelPosition sets the data label position.
func (s *ChartSeries) SetLabelPosition(pos string) *ChartSeries {
	s.LabelPosition = pos
	return s
}

// SeriesOutline represents a series outline.
type SeriesOutline struct {
	Width int
	Color Color
}

// SeriesMarker represents a series marker.
type SeriesMarker struct {
	Symbol string
	Size   int
}

// Marker symbol constants.
const (
	MarkerCircle   = "circle"
	MarkerDash     = "dash"
	MarkerDiamond  = "diamond"
	MarkerDot      = "dot"
	MarkerPlus     = "plus"
	MarkerSquare   = "square"
	MarkerStar     = "star"
	MarkerTriangle = "triangle"
	MarkerX        = "x"
	MarkerNone     = "none"
)

// --- Concrete Chart Types ---

// BarChart represents a bar/column chart.
type BarChart struct {
	Series          []*ChartSeries
	BarGrouping     string
	BarDirection    string
	GapWidthPercent int
	OverlapPercent  int
}

// Bar grouping constants.
const (
	BarGroupingClustered      = "clustered"
	BarGroupingStacked        = "stacked"
	BarGroupingPercentStacked = "percentStacked"
)

// Bar direction constants.
const (
	BarDirectionVertical   = "col"
	BarDirectionHorizontal = "bar"
)

func (b *BarChart) GetChartTypeName() string { return "bar" }

// NewBarChart creates a new bar chart.
func NewBarChart() *BarChart {
	return &BarChart{
		Series:          make([]*ChartSeries, 0),
		BarGrouping:     BarGroupingClustered,
		BarDirection:    BarDirectionVertical,
		GapWidthPercent: 150,
		OverlapPercent:  0,
	}
}

// AddSeries adds a data series.
func (b *BarChart) AddSeries(s *ChartSeries) *BarChart {
	b.Series = append(b.Series, s)
	return b
}

// SetBarGrouping sets the bar grouping type.
func (b *BarChart) SetBarGrouping(g string) *BarChart {
	b.BarGrouping = g
	if g == BarGroupingStacked || g == BarGroupingPercentStacked {
		b.OverlapPercent = 100
	} else {
		b.OverlapPercent = 0
	}
	return b
}

// SetGapWidthPercent sets the gap width percentage (0-500).
func (b *BarChart) SetGapWidthPercent(v int) *BarChart {
	if v < 0 {
		v = 0
	}
	if v > 500 {
		v = 500
	}
	b.GapWidthPercent = v
	return b
}

// SetOverlapPercent sets the overlap percentage (-100 to 100).
func (b *BarChart) SetOverlapPercent(v int) *BarChart {
	if v < -100 {
		v = -100
	}
	if v > 100 {
		v = 100
	}
	b.OverlapPercent = v
	return b
}

// Bar3DChart represents a 3D bar chart.
type Bar3DChart struct {
	BarChart
}

func (b *Bar3DChart) GetChartTypeName() string { return "bar3D" }

// NewBar3DChart creates a new 3D bar chart.
func NewBar3DChart() *Bar3DChart {
	return &Bar3DChart{BarChart: *NewBarChart()}
}

// LineChart represents a line chart.
type LineChart struct {
	Series   []*ChartSeries
	IsSmooth bool
}

func (l *LineChart) GetChartTypeName() string { return "line" }

// NewLineChart creates a new line chart.
func NewLineChart() *LineChart {
	return &LineChart{
		Series: make([]*ChartSeries, 0),
	}
}

// AddSeries adds a data series.
func (l *LineChart) AddSeries(s *ChartSeries) *LineChart {
	l.Series = append(l.Series, s)
	return l
}

// SetSmooth sets whether the line is smooth.
func (l *LineChart) SetSmooth(v bool) *LineChart {
	l.IsSmooth = v
	return l
}

// AreaChart represents an area chart.
type AreaChart struct {
	Series []*ChartSeries
}

func (a *AreaChart) GetChartTypeName() string { return "area" }

// NewAreaChart creates a new area chart.
func NewAreaChart() *AreaChart {
	return &AreaChart{Series: make([]*ChartSeries, 0)}
}

// AddSeries adds a data series.
func (a *AreaChart) AddSeries(s *ChartSeries) *AreaChart {
	a.Series = append(a.Series, s)
	return a
}

// PieChart represents a pie chart.
type PieChart struct {
	Series []*ChartSeries
}

func (p *PieChart) GetChartTypeName() string { return "pie" }

// NewPieChart creates a new pie chart.
func NewPieChart() *PieChart {
	return &PieChart{Series: make([]*ChartSeries, 0)}
}

// AddSeries adds a data series.
func (p *PieChart) AddSeries(s *ChartSeries) *PieChart {
	p.Series = append(p.Series, s)
	return p
}

// Pie3DChart represents a 3D pie chart.
type Pie3DChart struct {
	PieChart
}

func (p *Pie3DChart) GetChartTypeName() string { return "pie3D" }

// NewPie3DChart creates a new 3D pie chart.
func NewPie3DChart() *Pie3DChart {
	return &Pie3DChart{PieChart: *NewPieChart()}
}

// DoughnutChart represents a doughnut chart.
type DoughnutChart struct {
	Series    []*ChartSeries
	HoleSize  int // percentage 10-90
}

func (d *DoughnutChart) GetChartTypeName() string { return "doughnut" }

// NewDoughnutChart creates a new doughnut chart.
func NewDoughnutChart() *DoughnutChart {
	return &DoughnutChart{
		Series:   make([]*ChartSeries, 0),
		HoleSize: 50,
	}
}

// AddSeries adds a data series.
func (d *DoughnutChart) AddSeries(s *ChartSeries) *DoughnutChart {
	d.Series = append(d.Series, s)
	return d
}

// ScatterChart represents a scatter chart.
type ScatterChart struct {
	Series   []*ChartSeries
	IsSmooth bool
}

func (s *ScatterChart) GetChartTypeName() string { return "scatter" }

// NewScatterChart creates a new scatter chart.
func NewScatterChart() *ScatterChart {
	return &ScatterChart{Series: make([]*ChartSeries, 0)}
}

// AddSeries adds a data series.
func (s *ScatterChart) AddSeries(series *ChartSeries) *ScatterChart {
	s.Series = append(s.Series, series)
	return s
}

// SetSmooth sets whether the line is smooth.
func (s *ScatterChart) SetSmooth(v bool) *ScatterChart {
	s.IsSmooth = v
	return s
}

// RadarChart represents a radar chart.
type RadarChart struct {
	Series []*ChartSeries
}

func (r *RadarChart) GetChartTypeName() string { return "radar" }

// NewRadarChart creates a new radar chart.
func NewRadarChart() *RadarChart {
	return &RadarChart{Series: make([]*ChartSeries, 0)}
}

// AddSeries adds a data series.
func (r *RadarChart) AddSeries(s *ChartSeries) *RadarChart {
	r.Series = append(r.Series, s)
	return r
}
