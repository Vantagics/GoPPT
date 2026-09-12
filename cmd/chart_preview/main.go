// Command chart_preview rasterizes one slide per native chart type to PNG.
//
// It builds charts from scratch and renders them, so it exercises the
// write -> rasterize path of the built-in renderer without needing an input
// .pptx. That makes it a quick way to eyeball chart output (bar grouping, gap
// width, axes and gridlines, data labels, legend placement, ...) after a
// renderer change.
//
// Usage:
//
//	go run ./cmd/chart_preview
//	go run ./cmd/chart_preview -draft -width 480   # cheap pass for a quick look
//
// Images are written to cmd/chart_preview/out/.
//
// The run builds one FontCache and reuses it for every chart, because
// constructing a cache scans the system font directories. That is the pattern to
// copy when rendering a whole deck.
package main

import (
	"flag"
	"fmt"
	"os"
	"path/filepath"

	ppt "github.com/VantageDataChat/GoPPT"
)

const outDir = "cmd/chart_preview/out"

var (
	categories = []string{"Q1", "Q2", "Q3", "Q4"}
	palette    = []string{"FF4472C4", "FFED7D31", "FFA5A5A5", "FFFFC000"}
)

// preview describes a single rendered chart.
type preview struct {
	file  string // output base name, without extension
	title string // chart title
	build func() ppt.ChartType
}

// series builds a series with an explicit fill colour, so previews stay stable
// no matter which colour scheme the default theme uses.
func series(title string, values []float64, paletteIndex int) *ppt.ChartSeries {
	s := ppt.NewChartSeriesOrdered(title, categories, values)
	return s.SetFillColor(ppt.NewColor(palette[paletteIndex%len(palette)]))
}

func previews() []preview {
	return []preview{
		{
			file:  "bar_clustered",
			title: "Clustered Column",
			build: func() ppt.ChartType {
				c := ppt.NewBarChart()
				c.SetBarGrouping(ppt.BarGroupingClustered)
				c.AddSeries(series("Product A", []float64{42, 58, 71, 63}, 0))
				c.AddSeries(series("Product B", []float64{31, 44, 52, 68}, 1))
				return c
			},
		},
		{
			file:  "bar_stacked",
			title: "Stacked Column",
			build: func() ppt.ChartType {
				c := ppt.NewBarChart()
				c.SetBarGrouping(ppt.BarGroupingStacked)
				c.SetGapWidthPercent(80)
				c.AddSeries(series("Product A", []float64{42, 58, 71, 63}, 0))
				c.AddSeries(series("Product B", []float64{31, 44, 52, 68}, 1))
				return c
			},
		},
		{
			file:  "bar_percent_stacked",
			title: "100% Stacked Column",
			build: func() ppt.ChartType {
				c := ppt.NewBarChart()
				c.SetBarGrouping(ppt.BarGroupingPercentStacked)
				c.SetGapWidthPercent(80)
				c.AddSeries(series("Product A", []float64{42, 58, 71, 63}, 0))
				c.AddSeries(series("Product B", []float64{31, 44, 52, 68}, 1))
				return c
			},
		},
		{
			file:  "bar_horizontal",
			title: "Horizontal Bar",
			build: func() ppt.ChartType {
				c := ppt.NewBarChart()
				c.BarDirection = ppt.BarDirectionHorizontal
				c.SetGapWidthPercent(60)
				c.AddSeries(series("Product A", []float64{42, 58, 71, 63}, 0))
				c.AddSeries(series("Product B", []float64{31, 44, 52, 68}, 1))
				return c
			},
		},
		{
			file:  "bar_3d",
			title: "3-D Column",
			build: func() ppt.ChartType {
				c := ppt.NewBar3DChart()
				c.AddSeries(series("Product A", []float64{42, 58, 71, 63}, 0))
				c.AddSeries(series("Product B", []float64{31, 44, 52, 68}, 1))
				return c
			},
		},
		{
			file:  "line_smooth",
			title: "Smoothed Line",
			build: func() ppt.ChartType {
				c := ppt.NewLineChart()
				c.SetSmooth(true)
				a := series("Revenue", []float64{12, 30, 24, 45}, 0)
				a.Marker = &ppt.SeriesMarker{Symbol: ppt.MarkerCircle, Size: 7}
				b := series("Target", []float64{20, 26, 33, 38}, 1)
				b.Marker = &ppt.SeriesMarker{Symbol: ppt.MarkerSquare, Size: 7}
				c.AddSeries(a)
				c.AddSeries(b)
				return c
			},
		},
		{
			file:  "area",
			title: "Area",
			build: func() ppt.ChartType {
				c := ppt.NewAreaChart()
				c.AddSeries(series("North", []float64{20, 32, 28, 41}, 0))
				c.AddSeries(series("South", []float64{14, 19, 25, 22}, 1))
				return c
			},
		},
		{
			file:  "pie",
			title: "Pie",
			build: func() ppt.ChartType {
				c := ppt.NewPieChart()
				s := series("Share", []float64{35, 25, 25, 15}, 0)
				s.ShowPercentage = true
				c.AddSeries(s)
				return c
			},
		},
		{
			file:  "pie_3d",
			title: "3-D Pie",
			build: func() ppt.ChartType {
				c := ppt.NewPie3DChart()
				s := series("Share", []float64{35, 25, 25, 15}, 0)
				s.ShowPercentage = true
				c.AddSeries(s)
				return c
			},
		},
		{
			file:  "doughnut",
			title: "Doughnut",
			build: func() ppt.ChartType {
				c := ppt.NewDoughnutChart()
				c.HoleSize = 60
				s := series("Share", []float64{40, 30, 20, 10}, 0)
				s.ShowPercentage = true
				c.AddSeries(s)
				return c
			},
		},
		{
			file:  "scatter",
			title: "Scatter",
			build: func() ppt.ChartType {
				c := ppt.NewScatterChart()
				s := series("Samples", []float64{5, 12, 9, 20}, 0)
				s.Marker = &ppt.SeriesMarker{Symbol: ppt.MarkerCircle, Size: 8}
				c.AddSeries(s)
				return c
			},
		},
		{
			file:  "radar",
			title: "Radar",
			build: func() ppt.ChartType {
				c := ppt.NewRadarChart()
				c.AddSeries(series("Team A", []float64{70, 85, 60, 90}, 0))
				c.AddSeries(series("Team B", []float64{55, 65, 80, 70}, 1))
				return c
			},
		},
	}
}

func main() {
	draft := flag.Bool("draft", false, "render with Draft mode: skip anti-aliasing, shadows and image smoothing")
	width := flag.Int("width", 1280, "output width in pixels")
	flag.Parse()

	if err := os.MkdirAll(outDir, 0o755); err != nil {
		fmt.Fprintf(os.Stderr, "create output dir: %v\n", err)
		os.Exit(1)
	}

	// One FontCache for the whole run. Constructing a cache scans the system
	// font directories, so leaving FontCache unset would repeat that scan for
	// every one of the previews below.
	fc := ppt.NewFontCache()
	opts := ppt.DefaultRenderOptions()
	opts.Width = *width
	opts.Draft = *draft
	opts.FontCache = fc

	cases := previews()
	for _, tc := range cases {
		// A fresh presentation per chart keeps each image focused on one chart.
		pres := ppt.New()
		slide, err := pres.GetSlide(0)
		if err != nil {
			fmt.Fprintf(os.Stderr, "slide for %s: %v\n", tc.file, err)
			os.Exit(1)
		}

		chart := slide.CreateChartShape()
		chart.BaseShape.SetOffsetX(600000)
		chart.BaseShape.SetOffsetY(400000)
		chart.BaseShape.SetWidth(8000000)
		chart.BaseShape.SetHeight(4800000)

		chart.GetTitle().SetText(tc.title)
		chart.GetTitle().SetVisible(true)
		chart.GetTitle().Font.SetBold(true).SetSize(16)

		chart.GetLegend().Visible = true
		chart.GetLegend().Position = ppt.LegendBottom

		chart.GetPlotArea().SetType(tc.build())

		out := filepath.Join(outDir, tc.file+".png")
		if err := pres.SaveSlideAsImage(0, out, opts); err != nil {
			fmt.Fprintf(os.Stderr, "render %s: %v\n", tc.file, err)
			os.Exit(1)
		}
		fmt.Printf("OK: %s\n", out)
	}

	mode := "full"
	if *draft {
		mode = "draft"
	}
	fmt.Printf("Rendered %d chart previews to %s (%s quality, %dpx)\n",
		len(cases), outDir, mode, opts.Width)
}
