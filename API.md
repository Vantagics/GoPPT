# GoPPT API Reference / API 参考文档

[English](#english) | [中文](#中文)

---

<a id="english"></a>

## English

Package: `github.com/Vantagics/GoPPT`

All dimensions use EMU (English Metric Units): 1 inch = 914400 EMU, 1 cm = 360000 EMU, 1 pt = 12700 EMU.

---

### Presentation

The root object representing a PowerPoint file.

```go
// Create a new presentation (includes one blank slide)
p := ppt.New()

// Document properties
p.GetDocumentProperties().Title = "Title"
p.GetDocumentProperties().Creator = "Author"
p.GetDocumentProperties().Description = "Description"
p.GetDocumentProperties().Subject = "Subject"
p.GetDocumentProperties().Keywords = "go, pptx"
p.GetDocumentProperties().Category = "Report"
p.GetDocumentProperties().Company = "ACME"

// Custom properties
p.GetDocumentProperties().SetCustomProperty("version", "1.0", ppt.PropertyTypeString)
p.GetDocumentProperties().GetCustomPropertyValue("version") // "1.0"
// Custom properties are written to docProps/custom.xml, which is declared in
// [Content_Types].xml and related from _rels/.rels. The part is emitted only
// when at least one custom property is set.

// Presentation properties
p.GetPresentationProperties().SetZoom(1.5)
p.GetPresentationProperties().SetLastView(ppt.ViewSlide)
p.GetPresentationProperties().SetSlideshowType(ppt.SlideshowTypePresent)
// Zoom, last view and slideshow type are written (to ppt/viewProps.xml and
// ppt/presProps.xml). Comment visibility and MarkAsFinal are in-memory only:
// the extension that carries them is not written, so they do not survive a save.
p.GetPresentationProperties().MarkAsFinal()

// Layout
p.GetLayout().SetLayout(ppt.LayoutScreen16x9)
p.GetLayout().SetCustomLayout(9144000, 6858000) // custom EMU dimensions
```

| Layout Constant | Description |
|---|---|
| `LayoutScreen4x3` | 10" × 7.5" (default) |
| `LayoutScreen16x9` | 13.33" × 7.5" |
| `LayoutScreen16x10` | 12" × 7.5" |
| `LayoutA4` | A4 landscape |
| `LayoutLetter` | US Letter |
| `LayoutCustom` | Custom dimensions |

---

### Slides

```go
slide := p.CreateSlide()           // create and append
slide := p.GetActiveSlide()        // get current active slide
p.SetActiveSlideIndex(1)           // switch active slide
slide, _ := p.GetSlide(0)         // get by index
slides := p.GetAllSlides()         // get all
count := p.GetSlideCount()         // count
p.RemoveSlideByIndex(0)            // remove

slide.SetName("Intro")
slide.SetNotes("Speaker notes here")
slide.SetVisible(true)
slide.SetBackground(ppt.NewFill().SetSolid(ppt.ColorWhite))
```

---

### Shapes

All shapes share a common `BaseShape` with position, size, fill, border, shadow, and hyperlink.

```go
// Common BaseShape methods (available on all shapes)
shape.BaseShape.SetOffsetX(914400)   // 1 inch from left
shape.BaseShape.SetOffsetY(914400)   // 1 inch from top
shape.BaseShape.SetWidth(5000000)
shape.BaseShape.SetHeight(3000000)
shape.BaseShape.SetName("My Shape")
shape.BaseShape.SetRotation(45)      // degrees
shape.BaseShape.SetFill(fill)
shape.BaseShape.SetBorder(border)
shape.BaseShape.SetShadow(shadow)
shape.BaseShape.SetHyperlink(ppt.NewHyperlink("https://example.com")) // not serialized yet — see Hyperlink
```

#### RichTextShape

```go
rt := slide.CreateRichTextShape()
rt.SetOffsetX(100).SetOffsetY(100).SetWidth(8000000).SetHeight(1000000)
rt.SetWordWrap(true)
rt.SetAutoFit(ppt.AutoFitNormal)
rt.SetColumns(2)

// Paragraphs and text runs
para := rt.GetActiveParagraph()
tr := para.CreateTextRun("Hello World")
tr.GetFont().SetBold(true).SetSize(24).SetColor(ppt.ColorRed).SetName("Arial")
tr.GetFont().SetItalic(true).SetUnderline(ppt.UnderlineSingle).SetStrikethrough(true)

// Line break
para.CreateBreak()
para.CreateTextRun("Second line")

// New paragraph
para2 := rt.CreateParagraph()
para2.GetAlignment().SetHorizontal(ppt.HorizontalCenter)
para2.SetLineSpacing(200)
para2.SetSpaceBefore(100)
para2.SetSpaceAfter(50)
```

#### DrawingShape (Images)

```go
// From byte data
img := slide.CreateDrawingShape()
img.SetImageData(imageBytes, "image/png")
img.SetWidth(2000000).SetHeight(1500000).SetOffsetX(100).SetOffsetY(100)

// From file path
img2 := ppt.NewDrawingShape()
img2.SetPath("/path/to/image.jpg")
img2.SetWidth(2000000).SetHeight(1500000)
slide.AddShape(img2)
```

Supported formats: PNG, JPEG, GIF, BMP, SVG.

SVG is read and rasterised for rendering over a documented subset of the format
(`svg`, `g`, `polygon`, `polyline`, `line`, `rect`, `circle`, `ellipse`, `path`,
with `fill`/`stroke`/`stroke-width`/`fill-rule`, opacity, transforms and
`viewBox`). A construct outside the subset — `text`, `image`, gradients,
patterns, clipping, masks, filters, `<use>`, markers — is refused rather than
partially drawn, and the picture is shown as a labelled placeholder.
`<defs>` is never painted. An SVG can also be written, and its bytes and
`image/svg+xml` content type are stored correctly, but the Microsoft
`asvg:svgBlip` extension element is not emitted, so such a picture may not
display in PowerPoint; pass a raster fallback as the picture's data if the file
has to open there. See *SVG pictures* in the README.

#### TableShape

```go
table := slide.CreateTableShape(3, 4) // 3 rows, 4 columns
table.SetWidth(8000000).SetHeight(2000000)
table.BaseShape.SetOffsetX(500000).SetOffsetY(2000000)

cell := table.GetCell(0, 0) // row 0, col 0
cell.SetText("Header")
cell.SetFill(ppt.NewFill().SetSolid(ppt.ColorBlue))
cell.SetColSpan(2)
cell.SetRowSpan(1)

// Cell borders are per side, and the width is in points.
cell.GetBorders().Top.SetSolidFill(ppt.ColorRed).SetWidth(2)
cell.GetBorders().Bottom.Style = ppt.BorderDash // BorderSolid / BorderDash / BorderDot

// A merge is written as gridSpan/rowSpan on the spanning cell plus the empty
// continuation cells PowerPoint expects for the positions it covers, so a row
// always holds one <a:tc> per column. Text on a cell another cell spans is not
// written — that position belongs to the span.
//
// A table read from a file keeps its <a:gridCol> widths and <a:tr> heights when
// it is saved again; a table built through the API is split evenly.
```

#### AutoShape

```go
shape := slide.CreateAutoShape()
shape.SetAutoShapeType(ppt.AutoShapeRoundedRect)
shape.BaseShape.SetOffsetX(100).SetOffsetY(100).SetWidth(2000000).SetHeight(1000000)
shape.SetText("Inside the shape")
shape.BaseShape.SetFill(ppt.NewFill().SetSolid(ppt.ColorYellow))
```

| AutoShape Type | Constant |
|---|---|
| Rectangle | `AutoShapeRectangle` |
| Rounded Rectangle | `AutoShapeRoundedRect` |
| Ellipse | `AutoShapeEllipse` |
| Triangle | `AutoShapeTriangle` |
| Diamond | `AutoShapeDiamond` |
| Pentagon | `AutoShapePentagon` |
| Hexagon | `AutoShapeHexagon` |
| Star (4/5 point) | `AutoShapeStar4`, `AutoShapeStar5` |
| Arrows | `AutoShapeArrowRight/Left/Up/Down` |
| Heart | `AutoShapeHeart` |
| Lightning Bolt | `AutoShapeLightningBolt` |

#### LineShape

```go
line := slide.CreateLineShape()
line.BaseShape.SetOffsetX(0).SetOffsetY(0).SetWidth(5000000).SetHeight(0)
line.SetLineWidth(2).SetLineColor(ppt.ColorRed).SetLineStyle(ppt.BorderSolid)
```

#### GroupShape

```go
group := slide.CreateGroupShape()
group.BaseShape.SetOffsetX(0).SetOffsetY(0).SetWidth(5000000).SetHeight(3000000)

child := ppt.NewRichTextShape()
child.SetOffsetX(0).SetOffsetY(0).SetWidth(2000000).SetHeight(500000)
child.CreateTextRun("Inside group")
group.AddShape(child)

group.GetShapes()      // []Shape
group.GetShapeCount()  // int
group.RemoveShape(0)   // remove by index
```

#### PlaceholderShape

```go
ph := slide.CreatePlaceholderShape(ppt.PlaceholderTitle)
ph.BaseShape.SetOffsetX(500000).SetOffsetY(300000).SetWidth(8000000).SetHeight(1000000)
ph.CreateTextRun("Slide Title")
ph.SetPlaceholderIndex(0)
```

| Placeholder Type | Constant |
|---|---|
| Title | `PlaceholderTitle` |
| Body | `PlaceholderBody` |
| Center Title | `PlaceholderCtrTitle` |
| Subtitle | `PlaceholderSubTitle` |
| Date | `PlaceholderDate` |
| Footer | `PlaceholderFooter` |
| Slide Number | `PlaceholderSlideNum` |

#### UnsupportedShape

Created by the reader, not by callers, for OOXML constructs the library cannot draw (SmartArt, OLE objects, unreadable chart parts). It keeps the original frame geometry and renders as a labelled placeholder box; see [Unsupported shapes](#unsupported-shapes).

```go
u := ppt.NewUnsupportedShape("SmartArt diagram") // mostly useful in tests
u.BaseShape.SetOffsetX(2000000).SetOffsetY(1500000).SetWidth(4000000).SetHeight(2500000)
slide.AddShape(u)
```

`GetType()` returns `ppt.ShapeTypeUnsupported`, which `ShapeType.String()` renders as `"Unsupported"`.

---

### Charts

```go
chart := slide.CreateChartShape()
chart.BaseShape.SetOffsetX(500000).SetOffsetY(500000)
chart.BaseShape.SetWidth(7000000).SetHeight(4500000)

// Title
chart.GetTitle().SetText("My Chart").SetVisible(true)
chart.GetTitle().Font.SetBold(true).SetSize(14)

// Legend
chart.GetLegend().Visible = true
chart.GetLegend().Position = ppt.LegendBottom // b, t, l, r, tr

// Display blank values
chart.SetDisplayBlankAs(ppt.ChartBlankAsZero) // "gap", "zero", "span"

// 3D view (for 3D charts)
chart.GetView3D().RotX = 15
chart.GetView3D().RotY = 20
```

#### Chart Types

```go
// Bar / Column
bar := ppt.NewBarChart()
bar.SetBarGrouping(ppt.BarGroupingClustered) // clustered, stacked, percentStacked
bar.SetGapWidthPercent(150)  // 0-500
bar.SetOverlapPercent(0)     // -100 to 100
bar.AddSeries(ppt.NewChartSeriesOrdered("Sales", categories, values))

// 3D Bar
bar3d := ppt.NewBar3DChart()

// Line
line := ppt.NewLineChart()
line.SetSmooth(true)

// Area
area := ppt.NewAreaChart()

// Pie
pie := ppt.NewPieChart()

// 3D Pie
pie3d := ppt.NewPie3DChart()

// Doughnut
doughnut := ppt.NewDoughnutChart()
doughnut.HoleSize = 75 // 10-90

// Scatter
scatter := ppt.NewScatterChart()
scatter.SetSmooth(true)

// Radar
radar := ppt.NewRadarChart()
```

#### Chart Series

```go
s := ppt.NewChartSeriesOrdered("Series Name",
    []string{"Cat1", "Cat2", "Cat3"},
    []float64{10, 20, 30},
)
s.SetFillColor(ppt.ColorRed)
s.Outline = &ppt.SeriesOutline{Width: 3, Color: ppt.ColorRed} // Width is in points
s.SetLabelPosition(ppt.LabelOutsideEnd)
s.ShowValue = true
s.ShowCategoryName = true
s.ShowPercentage = true
s.ShowSeriesName = true
s.Separator = ", "
s.Marker = &ppt.SeriesMarker{Symbol: ppt.MarkerCircle, Size: 5}
```

Where a series' colour is emitted depends on the chart type. A **stroked** series (line, scatter)
puts its colour in `<c:spPr><a:ln>`; a **filled** series (bar, pie, area, radar) puts it in
`<c:spPr><a:solidFill>`. The writer picks the right one from the chart type, so a line series'
colour is not written somewhere PowerPoint ignores. `SeriesOutline.Width` is in points, like
`Border.Width`, and is converted to EMU on write (`pt × 12700`) and back on read.

Like a shape, a chart's own plot area supports fill and outline. A chart's area style is stored on
the shape itself (`GetFill` / `GetBorder`, the same fill and border the other shapes use) and is
written as the chart part's `<c:chartSpace><c:spPr>`, not onto the `p:graphicFrame`:

```go
chart.GetFill().SetSolid(ppt.NewColor("EEEEEE"))
chart.GetBorder().SetSolidFill(ppt.NewColor("445566")).SetWidth(2) // Width is in points
```

#### Chart Axes

```go
axX := chart.GetPlotArea().GetAxisX()
axX.SetTitle("Category").SetVisible(true)
axX.SetReversedOrder(true)
axX.SetMajorGridlines(ppt.NewGridlines())

axY := chart.GetPlotArea().GetAxisY()
axY.SetTitle("Value").SetMinBounds(0).SetMaxBounds(100)
axY.SetMajorUnit(20).SetMinorUnit(5)
axY.SetMinorGridlines(&ppt.Gridlines{Width: 1, Color: ppt.ColorBlack})
```

#### Reading Charts

Charts embedded as native chart parts are read back as `*ppt.ChartShape`, so an opened
presentation round-trips through the writer without losing them.

```go
reader := &ppt.PPTXReader{}
pres, _ := reader.Read("input.pptx")

slide, _ := pres.GetSlide(0)
for _, sh := range slide.Shapes() {
    if chart, ok := sh.(*ppt.ChartShape); ok {
        // The concrete chart type is preserved (bar, bar3D, line, area,
        // pie, pie3D, doughnut, scatter, radar), along with series,
        // categories, values, grouping, gap/overlap and hole size.
        fmt.Println(chart.GetTitle().Text)
    }
}
```

The reader follows `p:graphicFrame` → `a:graphicData` (uri ending in `/chart`) →
`c:chart r:id`, resolves the relationship in `ppt/slides/_rels/slideN.xml.rels`, and parses
`ppt/charts/chartN.xml`. Series colours written as `schemeClr` are resolved against the
presentation theme, and `<a:alpha>` values are folded into the ARGB colour. Non-contiguous
`<c:pt idx="…">` indices are preserved rather than compacted, so sparse caches keep their
original category alignment. Titles and legends are only reported when the source chart
actually declares them — nothing is invented.

Chart styling survives a full read → modify → write round trip. A chart area's fill and outline
are read from `<c:chartSpace><c:spPr>` into `GetFill` / `GetBorder`, and each series' outline
(`SeriesOutline` — colour and width in points) is read from the series `<c:spPr><a:ln>`. The
writer emits both again, so opening a deck that had a shaded chart area and saving it no longer
drops the shading. For a series, the writer places the colour where the chart type expects it:
on `<a:ln>` for a stroked series (line, scatter) and in `<a:solidFill>` for a filled one (bar,
pie, area, radar).

---

### Styles

#### Color

```go
ppt.ColorBlack   // FF000000
ppt.ColorWhite   // FFFFFFFF
ppt.ColorRed     // FFFF0000
ppt.ColorGreen   // FF00FF00
ppt.ColorBlue    // FF0000FF
ppt.ColorYellow  // FFFFFF00

custom := ppt.NewColor("FF8800")     // RGB (auto-adds FF alpha)
custom2 := ppt.NewColor("80FF8800")  // ARGB with transparency
```

#### Font

```go
font := ppt.NewFont()
font.SetName("Arial").SetSize(12)
font.SetBold(true).SetItalic(true)
font.SetColor(ppt.ColorRed)
font.SetUnderline(ppt.UnderlineSingle) // none, sng, dbl, heavy, dash, wavy
font.SetStrikethrough(true)
```

#### Fill

```go
solid := ppt.NewFill().SetSolid(ppt.ColorBlue)
gradient := ppt.NewFill().SetGradientLinear(ppt.ColorRed, ppt.ColorBlue, 90)
```

#### Border

```go
border := &ppt.Border{
    Style: ppt.BorderSolid, // none, solid, dash, dot
    Width: 2,
    Color: ppt.ColorBlack,
}
```

#### Shadow

```go
shadow := ppt.NewShadow()
shadow.SetVisible(true).SetDirection(45).SetDistance(5)
shadow.BlurRadius = 3
shadow.Color = ppt.Color{ARGB: "80000000"}
shadow.Alpha = 50
```

#### Alignment

```go
align := ppt.NewAlignment()
align.SetHorizontal(ppt.HorizontalCenter) // l, ctr, r, just, dist
align.SetVertical(ppt.VerticalMiddle)      // t, ctr, b
align.Level = 2 // indentation level
```

#### Hyperlink

A hyperlink lives on a text run. Both kinds are written to the file and read
back, so a presentation that is opened and saved again keeps them:

```go
ppt.NewHyperlink("https://example.com")       // external
ppt.NewInternalHyperlink(2)                     // link to slide 2

run.SetHyperlink(link)          // put it on a run
run.GetHyperlink().URL          // https://example.com
run.GetHyperlink().SlideNumber  // 2
```

An internal link names a slide by number. A number no slide in the presentation
backs is not written at all, rather than written as a relationship pointing at a
slide part that does not exist.

Not serialized: a hyperlink set on the shape itself with
`shape.BaseShape.SetHyperlink`. It is stored on the shape but never reaches the
file — no `<a:hlinkClick>` is written into the shape's `<p:cNvPr>`, and the
reader does not look for one. Put the link on a run instead.

---

### Bullets

```go
// Character bullet
bullet := ppt.NewBullet().SetCharBullet("•", "Arial")
bullet.SetColor(ppt.ColorRed).SetSize(120)

// Numeric bullet
bullet2 := ppt.NewBullet().SetNumericBullet(ppt.NumFormatArabicPeriod, 1)

para.SetBullet(bullet)
```

| Numeric Format | Constant |
|---|---|
| 1. 2. 3. | `NumFormatArabicPeriod` |
| 1) 2) 3) | `NumFormatArabicParen` |
| I. II. III. | `NumFormatRomanUcPeriod` |
| i. ii. iii. | `NumFormatRomanLcPeriod` |
| A. B. C. | `NumFormatAlphaUcPeriod` |
| a. b. c. | `NumFormatAlphaLcPeriod` |

---

### Comments

```go
author := ppt.NewCommentAuthor("John Doe", "JD")
comment := ppt.NewComment()
comment.SetAuthor(author).SetText("Review this").SetPosition(100, 200)
comment.SetDate(time.Now())
slide.AddComment(comment)
```

Comments round-trip. A comment is stored across two parts — the text body in
`ppt/comments/commentN.xml` and the author table in `ppt/commentAuthors.xml` — and both are
written and read, so an author's name and initials and the comment's timestamp survive a save.
The text body is emitted in the shape `p:text` requires (`a:bodyPr` plus paragraphs of `a:t`
runs), which is what PowerPoint reads. Reading also accepts the bare `<p:text>text</p:text>`
form written by older versions of this library.

---

### Writer / Reader

```go
// Write to file
w, _ := ppt.NewWriter(p, ppt.WriterPowerPoint2007)
w.(*ppt.PPTXWriter).Save("output.pptx")

// Write to io.Writer
var buf bytes.Buffer
w.WriteTo(&buf)

// Read from file
reader := &ppt.PPTXReader{}
pres, err := reader.Read("input.pptx")

// Read from io.ReaderAt
pres, err := reader.ReadFromReader(readerAt, size)

// Convenience wrappers
pres, err := ppt.Open("input.pptx")
pres, err := ppt.ReadFrom(readerAt, size)
```

Charts stored as native chart parts (`ppt/charts/chartN.xml`) are read back into `*ppt.ChartShape`, so a chart slide survives a write → read → render round trip. See [Charts](#charts) for the details, and [Panic safety](#panic-safety) for how malformed packages are reported.

Both read entry points return an error rather than panicking on malformed packages, including packages that are not PPTX files at all.

---

### Rendering

GoPPT includes a built-in renderer that exports slides as images with rendering quality close to Microsoft PowerPoint.

```go
reader := &ppt.PPTXReader{}
pres, _ := reader.Read("input.pptx")

opts := ppt.DefaultRenderOptions()
opts.Width = 1920 // output width in pixels; height follows the slide aspect ratio

// One slide, in memory
img, err := pres.SlideToImage(0, opts)

// Every slide, in memory
imgs, err := pres.SlidesToImages(opts)

// One slide to a file (PNG or JPEG per opts.Format)
err = pres.SaveSlideAsImage(0, "slide_1.png", opts)

// Every slide to files; the pattern must contain %d, numbered from 1
err = pres.SaveSlidesAsImages("slide_%d.png", opts)
```

All four take a slide index (or iterate every slide) plus an optional `*RenderOptions`; passing `nil` uses the defaults. `SlideToImage` returns a standard `image.Image`.

None of them modifies the `*RenderOptions` you pass: defaults are filled in on a copy, so one options value can be shared by concurrent renders. `SlidesToImages` and `SaveSlidesAsImages` build a single `FontCache` for the whole call unless you supply one, so the font directories are scanned once per call rather than once per slide. A slide size the library cannot turn into pixels — a nil layout, or a non-positive width or height — is reported as an ordinary error rather than as a recovered panic.

#### RenderOptions

| Field | Meaning |
| --- | --- |
| `Width` | Output width in pixels (default 960). Height follows the slide aspect ratio. |
| `Format` | `ppt.ImageFormatPNG` (default) or `ppt.ImageFormatJPEG`. |
| `JPEGQuality` | JPEG quality, 1-100 (default 90). |
| `BackgroundColor` | Overrides the slide background; `nil` uses the slide's own background or white. |
| `DPI` | DPI used for font sizing (default 96). |
| `FontDirs` | Extra directories to search for fonts; system directories are always searched. |
| `FontCache` | Share one `*ppt.FontCache` across renders. |
| `FontFallback` | Font names tried in order when the document's font is unavailable, ahead of the built-in chain. |
| `FontDiagnostics` | Collects the fonts substituted or not found during the render. |
| `OnFontFallback` | Callback fired for each font request that could not be satisfied exactly. |
| `OverlayOpacityScale` | Scales the opacity of semi-transparent fills (1.0 = unchanged). |
| `Draft` | Trades fidelity for speed in batch previews: skips anti-aliasing on lines and ellipses, skips shadows, and scales images with nearest-neighbour instead of bilinear. Text stays anti-aliased. The output keeps the same size and content. |

#### Fonts and missing-font diagnostics

Rendering with no access to the document's fonts cannot tell you *why* text looks wrong. Pass a `FontDiagnostics` to get the fonts that were substituted or not found at all, and optionally `OnFontFallback` to observe each event as it happens:

```go
diag := ppt.NewFontDiagnostics()
opts := ppt.DefaultRenderOptions()
opts.Width = 1600
opts.FontDiagnostics = diag
// Pin CJK text to a font you know is installed, or ship alongside the app.
opts.FontFallback = []string{"Noto Sans CJK SC"}

if _, err := pres.SlideToImage(0, opts); err != nil {
    return err
}

if diag.UsedBitmapFallback() {
    // Some font could not be found, so the built-in bitmap font was used
    // instead; CJK glyphs render as blank boxes (tofu).
    log.Printf("unresolved fonts: %v", diag.MissingNames())
}
log.Print(diag.Summary()) // one-line human-readable description
```

`FontCache.HasFont(name, bold, italic)` and `FontCache.FindFontName(name, bold, italic)` let you check availability up front, before rendering anything.

Availability is not coverage, and the renderer treats them differently. `FontCache.CoversRune(name, bold, italic, r)` asks the stronger question — does this font have a real glyph for `r`, read from its glyph index table, where index 0 means `.notdef`. For East Asian text the renderer picks a face by that test rather than by name, because a document can declare a Latin font as its East Asian font (`<a:ea typeface="Arial"/>`) and every name-based check will report a perfect match while the text draws as boxes. Enclosed symbols used inside CJK text — `①`, `㈠`, `㎡` — are classified as East Asian for this purpose, and each candidate face is checked against the characters the run actually contains. Answers are memoised per (font, style, rune), and a font that cannot cover *every* character of a run still wins when it covers the most, so one undrawable symbol does not push the whole run back to a Latin face.

Sharing a `FontCache` matters: constructing one scans the font directories. Build it once and assign `opts.FontCache` when rendering many slides or many files — `SlidesToImages` and `SaveSlidesAsImages` each reuse a single cache across the slides of one call. A cache is safe for concurrent use, so a worker pool rendering different decks in parallel can share one. Unresolved font requests are remembered too, so a deck naming an uninstalled font does not re-run the lookup for every text run. No render entry point writes back into the options you pass it.

`FontCache.LoadFont(name, path)` and `FontCache.LoadFontData(name, data)` register a font under a name of your choosing, and that registration outranks anything the directory scan later finds under the same name. Register a bundled font at startup and the render resolves to it even on a machine that has a different build of the same font installed, which is what makes output reproducible across machines.

#### Batch previews and draft mode

For contact sheets and bulk checks, set `Draft` and a modest `Width`. Draft skips anti-aliasing on lines and ellipses, skips shadows, and scales images with nearest-neighbour instead of bilinear. It does not change the output dimensions or the content, only the finish — text in particular stays anti-aliased, since glyph rasterisation cannot be turned off without replacing the text renderer.

```go
fc := ppt.NewFontCache()      // one scan, shared by every render
opts := ppt.DefaultRenderOptions()
opts.FontCache = fc
opts.Width = 480
opts.Draft = true

err := pres.SaveSlidesAsImages(filepath.Join(out, "slide_%d.png"), opts)
```

Indicative timings from a 960 px render (AMD Ryzen 7 8745HS), full versus draft — the ratios matter more than the absolute numbers:

| Slide | Full | Draft |
| --- | --- | --- |
| Text + chart | 692 µs | 406 µs |
| Shape with shadow | 521 µs | 161 µs |
| Full-bleed image | 11.2 ms | 5.0 ms |

Draft is a fidelity setting, not a different renderer: render the slides you actually inspect with `Draft` left at its default.

#### Unsupported shapes

Some constructs in a real deck cannot be rasterized: SmartArt diagrams, OLE objects, and charts whose part is missing or unparseable. The reader does not drop them. It keeps an `UnsupportedShape` carrying the original frame geometry, and the renderer paints a labelled amber box in its place, so a preview never shows a blank region that is indistinguishable from correctly rendered empty content.

```go
for _, s := range pres.UnsupportedShapes() {
    // Label() is e.g. "Unsupported: SmartArt diagram".
    log.Printf("not rendered: %s at (%d,%d) %dx%d",
        s.Label(), s.GetOffsetX(), s.GetOffsetY(), s.GetWidth(), s.GetHeight())
}
```

| Method | Meaning |
| --- | --- |
| `pres.UnsupportedShapes()` | Every unsupported shape in the deck, in slide order, descending into groups. Empty when everything is renderable. |
| `UnsupportedShape.GetReason()` | Short description, e.g. `"SmartArt diagram"`, or `"chart (its part could not be read)"`. |
| `UnsupportedShape.GetContentType()` | The `a:graphicData` uri the shape came from, for precise classification. |
| `UnsupportedShape.Label()` | Text drawn inside the placeholder box. |
| `UnsupportedShape.GetType()` | Always `ppt.ShapeTypeUnsupported`. |

The original XML is not retained, so writing the presentation back out drops the shape; the rest of the slide round-trips normally. The caption is shrunk to fit the box, and a box too small for a legible caption keeps only its dashed frame — the label is library-generated text, so it is never allowed to spill outside the placeholder. For the full list of which OOXML constructs are read, rendered, placeholdered, or flattened, see the shape support matrix in [README.md](README.md#shape-support-matrix).

#### Panic safety

`Open`, `Read`, `ReadFromReader`, `GetSlide`, `SlideToImage`, `SlidesToImages`, `SaveSlideAsImage` and `SaveSlidesAsImages` recover panics raised while parsing or rasterizing malformed input and return them as a `*ppt.PanicError` instead of crashing the process, so callers do not need their own `recover` wrappers. Use `ppt.ErrIsPanic(err)` to detect this case, or `errors.As(err, &ppt.PanicError{})` to read the recovered value and stack.

The renderer uses a dual font-face architecture: HintingNone faces for text layout (matching PowerPoint's DirectWrite metrics) and HintingFull faces for crisp glyph rendering. CJK text receives special handling with kinsoku line-breaking rules and tuned line-height calculations.

Chart shapes are rasterized natively instead of being flattened into placeholder images. Clustered, stacked and percent-stacked bars (both vertical and horizontal) honour `GapWidthPercent` and `OverlapPercent`; line charts support smoothed curves with per-series markers; area, pie, 3D pie, doughnut, scatter and radar each have dedicated draw paths. 3-D chart types (`bar3D`, `pie3D`) are drawn as their 2-D equivalents — the renderer performs no 3-D projection. Value axes derive "nice" tick steps on a 1 / 2 / 5 × 10ⁿ progression (or use explicit bounds and major units), draw major gridlines and tick labels, and axis titles are rotated along the axis. Data labels, a legend that wraps and reserves space according to `LegendPosition`, and series colours resolved from `srgbClr` / `schemeClr` / `sysClr` / `prstClr` are all supported.

<a id="中文"></a>

## 中文

包路径：`github.com/Vantagics/GoPPT`

所有尺寸使用 EMU（英制公制单位）：1 英寸 = 914400 EMU，1 厘米 = 360000 EMU，1 磅 = 12700 EMU。

---

### 演示文稿 (Presentation)

根对象，代表一个 PowerPoint 文件。

```go
// 创建新演示文稿（自动包含一张空白幻灯片）
p := ppt.New()

// 文档属性
p.GetDocumentProperties().Title = "标题"
p.GetDocumentProperties().Creator = "作者"
p.GetDocumentProperties().Description = "描述"

// 自定义属性
p.GetDocumentProperties().SetCustomProperty("版本", "1.0", ppt.PropertyTypeString)
// 自定义属性写入 docProps/custom.xml，并在 [Content_Types].xml 中声明、
// 由 _rels/.rels 建立关系；只有设置了至少一个自定义属性时才会写出该部件。

// 演示文稿属性
p.GetPresentationProperties().SetZoom(1.5)
p.GetPresentationProperties().SetLastView(ppt.ViewSlide)
p.GetPresentationProperties().SetSlideshowType(ppt.SlideshowTypePresent)
// 缩放、上次视图、放映方式会写入文件（ppt/viewProps.xml 与 ppt/presProps.xml）。
// SetCommentVisible 与 MarkAsFinal 仅存在于内存：承载它们的扩展未写出，
// 保存后不会保留。

// 布局
p.GetLayout().SetLayout(ppt.LayoutScreen16x9)
p.GetLayout().SetCustomLayout(9144000, 6858000) // 自定义 EMU 尺寸
```

| 布局常量 | 说明 |
|---|---|
| `LayoutScreen4x3` | 10" × 7.5"（默认） |
| `LayoutScreen16x9` | 13.33" × 7.5" |
| `LayoutScreen16x10` | 12" × 7.5" |
| `LayoutA4` | A4 横向 |
| `LayoutLetter` | US Letter |
| `LayoutCustom` | 自定义尺寸 |

---

### 幻灯片 (Slide)

```go
slide := p.CreateSlide()           // 创建并添加
slide := p.GetActiveSlide()        // 获取当前活动幻灯片
p.SetActiveSlideIndex(1)           // 切换活动幻灯片
slide, _ := p.GetSlide(0)         // 按索引获取
slides := p.GetAllSlides()         // 获取全部
count := p.GetSlideCount()         // 计数
p.RemoveSlideByIndex(0)            // 删除

slide.SetName("简介")
slide.SetNotes("演讲者备注")
slide.SetVisible(true)
slide.SetBackground(ppt.NewFill().SetSolid(ppt.ColorWhite))
```

---

### 形状 (Shapes)

所有形状共享 `BaseShape`，包含位置、大小、填充、边框、阴影和超链接。

```go
// BaseShape 通用方法
shape.BaseShape.SetOffsetX(914400)   // 距左 1 英寸
shape.BaseShape.SetOffsetY(914400)   // 距顶 1 英寸
shape.BaseShape.SetWidth(5000000)
shape.BaseShape.SetHeight(3000000)
shape.BaseShape.SetName("我的形状")
shape.BaseShape.SetRotation(45)      // 度
shape.BaseShape.SetFill(fill)
shape.BaseShape.SetBorder(border)
shape.BaseShape.SetShadow(shadow)
shape.BaseShape.SetHyperlink(ppt.NewHyperlink("https://example.com")) // 尚未序列化 — 见「超链接」
```

#### 富文本形状 (RichTextShape)

```go
rt := slide.CreateRichTextShape()
rt.SetOffsetX(100).SetOffsetY(100).SetWidth(8000000).SetHeight(1000000)
rt.SetWordWrap(true)
rt.SetAutoFit(ppt.AutoFitNormal)
rt.SetColumns(2)

// 段落和文本运行
para := rt.GetActiveParagraph()
tr := para.CreateTextRun("你好世界")
tr.GetFont().SetBold(true).SetSize(24).SetColor(ppt.ColorRed).SetName("微软雅黑")

// 换行
para.CreateBreak()
para.CreateTextRun("第二行")

// 新段落
para2 := rt.CreateParagraph()
para2.GetAlignment().SetHorizontal(ppt.HorizontalCenter)
para2.SetLineSpacing(200)
```

#### 图片形状 (DrawingShape)

```go
// 从字节数据
img := slide.CreateDrawingShape()
img.SetImageData(imageBytes, "image/png")
img.SetWidth(2000000).SetHeight(1500000)

// 从文件路径
img2 := ppt.NewDrawingShape()
img2.SetPath("/path/to/image.jpg")
slide.AddShape(img2)
```

支持格式：PNG、JPEG、GIF、BMP、SVG。

SVG 可读取，并按一套明确的格式子集栅格化用于渲染（`svg`、`g`、`polygon`、
`polyline`、`line`、`rect`、`circle`、`ellipse`、`path`，以及
`fill`/`stroke`/`stroke-width`/`fill-rule`、不透明度、变换、`viewBox`）。
子集之外的构造——`text`、`image`、渐变、图案、裁剪、遮罩、滤镜、`<use>`、
marker——会被拒绝而不是只画一部分，该图片显示为带标注的占位框。
`<defs>` 永远不会被绘制。SVG 也可以写出，其字节与 `image/svg+xml`
Content-Type 都会正确存储，但不会写出微软的 `asvg:svgBlip` 扩展元素，
因此这类图片在 PowerPoint 中可能不显示；若文件必须在那里打开，
请把位图回退作为图片数据传入。详见 README 的「SVG 图片」。

#### 表格形状 (TableShape)

```go
table := slide.CreateTableShape(3, 4) // 3 行 4 列
table.SetWidth(8000000).SetHeight(2000000)

cell := table.GetCell(0, 0)
cell.SetText("表头")
cell.SetFill(ppt.NewFill().SetSolid(ppt.ColorBlue))
cell.SetColSpan(2)

// 单元格边框按边设置，宽度单位是磅。
cell.GetBorders().Top.SetSolidFill(ppt.ColorRed).SetWidth(2)
cell.GetBorders().Bottom.Style = ppt.BorderDash // BorderSolid / BorderDash / BorderDot

// 合并会写成跨格单元上的 gridSpan/rowSpan，以及 PowerPoint 期望的、覆盖位置上的
// 空续格，因此每行始终有与列数相同的 <a:tc>。被其他单元跨过的单元格上的文字不会
// 写出——那个位置属于合并区域。
//
// 从文件读入的表格在再次保存时会保留其 <a:gridCol> 列宽与 <a:tr> 行高；
// 通过 API 新建的表格则按列数均分。
```

#### 自动形状 (AutoShape)

```go
shape := slide.CreateAutoShape()
shape.SetAutoShapeType(ppt.AutoShapeEllipse)
shape.BaseShape.SetOffsetX(100).SetOffsetY(100).SetWidth(2000000).SetHeight(1000000)
shape.SetText("形状内文字")
shape.BaseShape.SetFill(ppt.NewFill().SetSolid(ppt.ColorYellow))
```

| 形状类型 | 常量 |
|---|---|
| 矩形 | `AutoShapeRectangle` |
| 圆角矩形 | `AutoShapeRoundedRect` |
| 椭圆 | `AutoShapeEllipse` |
| 三角形 | `AutoShapeTriangle` |
| 菱形 | `AutoShapeDiamond` |
| 五边形 | `AutoShapePentagon` |
| 六边形 | `AutoShapeHexagon` |
| 星形 | `AutoShapeStar4`, `AutoShapeStar5` |
| 箭头 | `AutoShapeArrowRight/Left/Up/Down` |
| 心形 | `AutoShapeHeart` |
| 闪电 | `AutoShapeLightningBolt` |

#### 线条形状 (LineShape)

```go
line := slide.CreateLineShape()
line.BaseShape.SetOffsetX(0).SetOffsetY(0).SetWidth(5000000).SetHeight(0)
line.SetLineWidth(2).SetLineColor(ppt.ColorRed)
```

#### 组合形状 (GroupShape)

```go
group := slide.CreateGroupShape()
group.BaseShape.SetOffsetX(0).SetOffsetY(0).SetWidth(5000000).SetHeight(3000000)

child := ppt.NewRichTextShape()
child.SetOffsetX(0).SetOffsetY(0).SetWidth(2000000).SetHeight(500000)
child.CreateTextRun("组内文字")
group.AddShape(child)
```

#### 占位符形状 (PlaceholderShape)

```go
ph := slide.CreatePlaceholderShape(ppt.PlaceholderTitle)
ph.BaseShape.SetOffsetX(500000).SetOffsetY(300000).SetWidth(8000000).SetHeight(1000000)
ph.CreateTextRun("幻灯片标题")
```

| 占位符类型 | 常量 |
|---|---|
| 标题 | `PlaceholderTitle` |
| 正文 | `PlaceholderBody` |
| 居中标题 | `PlaceholderCtrTitle` |
| 副标题 | `PlaceholderSubTitle` |
| 日期 | `PlaceholderDate` |
| 页脚 | `PlaceholderFooter` |
| 页码 | `PlaceholderSlideNum` |

#### 不支持的形状 (UnsupportedShape)

由 reader 创建（而非调用方），用于本库无法绘制的 OOXML 结构（SmartArt、OLE 对象、无法读取的图表部件）。它保留原始框体几何，并渲染为带标注的占位框；详见 [不支持的形状](#不支持的形状)。

```go
u := ppt.NewUnsupportedShape("SmartArt diagram") // 主要用于测试
u.BaseShape.SetOffsetX(2000000).SetOffsetY(1500000).SetWidth(4000000).SetHeight(2500000)
slide.AddShape(u)
```

`GetType()` 返回 `ppt.ShapeTypeUnsupported`，经 `ShapeType.String()` 输出为 `"Unsupported"`。

---

### 图表 (Charts)

```go
chart := slide.CreateChartShape()
chart.BaseShape.SetOffsetX(500000).SetOffsetY(500000)
chart.BaseShape.SetWidth(7000000).SetHeight(4500000)

// 标题
chart.GetTitle().SetText("我的图表").SetVisible(true)

// 图例
chart.GetLegend().Visible = true
chart.GetLegend().Position = ppt.LegendBottom
```

#### 图表类型

```go
// 柱状图
bar := ppt.NewBarChart()
bar.SetBarGrouping(ppt.BarGroupingClustered) // clustered, stacked, percentStacked
bar.AddSeries(ppt.NewChartSeriesOrdered("销售额", categories, values))

// 3D 柱状图
bar3d := ppt.NewBar3DChart()

// 折线图
line := ppt.NewLineChart()
line.SetSmooth(true)

// 面积图
area := ppt.NewAreaChart()

// 饼图 / 3D 饼图
pie := ppt.NewPieChart()
pie3d := ppt.NewPie3DChart()

// 环形图
doughnut := ppt.NewDoughnutChart()
doughnut.HoleSize = 75

// 散点图
scatter := ppt.NewScatterChart()

// 雷达图
radar := ppt.NewRadarChart()
```

#### 数据系列

```go
s := ppt.NewChartSeriesOrdered("系列名称",
    []string{"类别1", "类别2", "类别3"},
    []float64{10, 20, 30},
)
s.SetFillColor(ppt.ColorRed)
s.Outline = &ppt.SeriesOutline{Width: 3, Color: ppt.ColorRed} // Width 单位为磅
s.ShowValue = true
s.ShowPercentage = true
s.Marker = &ppt.SeriesMarker{Symbol: ppt.MarkerCircle, Size: 5}
```

系列颜色写到哪个元素上取决于图表类型。**描边型**系列（折线、散点）把颜色写在
`<c:spPr><a:ln>`；**填充型**系列（柱状、饼图、面积、雷达）把颜色写在 `<c:spPr><a:solidFill>`。
写入器会依据图表类型选择正确的元素，因此折线系列的颜色不会写到 PowerPoint 忽略的位置。
`SeriesOutline.Width` 的单位是磅，与 `Border.Width` 一致：写入时换算为 EMU（`磅 × 12700`），
读取时再换算回来。

与普通形状一样，图表自身的绘图区域也支持填充与描边。图表的区域样式存在形状本身
（`GetFill` / `GetBorder`，与其它形状用的是同一套填充与边框），写出时对应图表部件的
`<c:chartSpace><c:spPr>`，而不是 `p:graphicFrame`：

```go
chart.GetFill().SetSolid(ppt.NewColor("EEEEEE"))
chart.GetBorder().SetSolidFill(ppt.NewColor("445566")).SetWidth(2) // Width 单位为磅
```

#### 坐标轴

```go
axX := chart.GetPlotArea().GetAxisX()
axX.SetTitle("类别").SetVisible(true)
axX.SetMajorGridlines(ppt.NewGridlines())

axY := chart.GetPlotArea().GetAxisY()
axY.SetTitle("数值").SetMinBounds(0).SetMaxBounds(100)
axY.SetMajorUnit(20).SetMinorUnit(5)
```

#### 读取图表

以原生图表部件（`ppt/charts/chartN.xml`）保存的图表会被读取为 `*ppt.ChartShape`，
因此打开演示文稿后再写出不会丢失图表。

```go
reader := &ppt.PPTXReader{}
pres, _ := reader.Read("输入.pptx")

slide, _ := pres.GetSlide(0)
for _, sh := range slide.Shapes() {
    if chart, ok := sh.(*ppt.ChartShape); ok {
        // 图表类型（柱状/3D 柱状/折线/面积/饼图/3D 饼图/环形/散点/雷达）、
        // 数据系列、类别、数值、分组方式、间隙宽度与重叠比例、孔径均被保留。
        fmt.Println(chart.GetTitle().Text)
    }
}
```

读取流程为：`p:graphicFrame` → `a:graphicData`（uri 以 `/chart` 结尾）→ `c:chart r:id`，
再通过 `ppt/slides/_rels/slideN.xml.rels` 解析关系，最终解析 `ppt/charts/chartN.xml`。
以 `schemeClr` 表示的数据系列颜色会结合演示文稿主题解析，`<a:alpha>` 透明度会折算进 ARGB 颜色。
非连续的 `<c:pt idx="…">` 索引会被保留而不会压缩，因此稀疏数据缓存仍与原类别一一对应。
标题与图例仅在源图表确有声明时才生效，不会凭空生成。

图表样式可完整走通「读取 → 修改 → 写出」回环。图表区域的填充与描边会从
`<c:chartSpace><c:spPr>` 读入 `GetFill` / `GetBorder`，每个系列的描边（`SeriesOutline`——
颜色与宽度，单位为磅）则从系列 `<c:spPr><a:ln>` 读入。写入器会把两者重新写出，因此打开
一个带图表区域底色的演示文稿再保存，不会再丢失该底色。对于系列，写入器会把颜色放到图表
类型期望的位置：描边型系列（折线、散点）放在 `<a:ln>`，填充型系列（柱状、饼图、面积、雷达）
放在 `<a:solidFill>`。

---

### 样式 (Styles)

#### 颜色

```go
ppt.ColorBlack   // FF000000
ppt.ColorWhite   // FFFFFFFF
ppt.ColorRed     // FFFF0000
ppt.ColorGreen   // FF00FF00
ppt.ColorBlue    // FF0000FF
ppt.ColorYellow  // FFFFFF00

custom := ppt.NewColor("FF8800")     // RGB（自动添加 FF 透明度）
custom2 := ppt.NewColor("80FF8800")  // ARGB 含透明度
```

#### 字体

```go
font := ppt.NewFont()
font.SetName("微软雅黑").SetSize(12)
font.SetBold(true).SetItalic(true)
font.SetColor(ppt.ColorRed)
font.SetUnderline(ppt.UnderlineSingle) // none, sng, dbl, heavy, dash, wavy
font.SetStrikethrough(true)
```

#### 填充

```go
solid := ppt.NewFill().SetSolid(ppt.ColorBlue)
gradient := ppt.NewFill().SetGradientLinear(ppt.ColorRed, ppt.ColorBlue, 90)
```

#### 边框

```go
border := &ppt.Border{
    Style: ppt.BorderSolid, // none, solid, dash, dot
    Width: 2,
    Color: ppt.ColorBlack,
}
```

#### 阴影

```go
shadow := ppt.NewShadow()
shadow.SetVisible(true).SetDirection(45).SetDistance(5)
```

#### 对齐

```go
align := ppt.NewAlignment()
align.SetHorizontal(ppt.HorizontalCenter) // l, ctr, r, just, dist
align.SetVertical(ppt.VerticalMiddle)      // t, ctr, b
```

#### 超链接

超链接挂在文本 run 上。两种链接都会写入文件并读回，因此「打开再保存」不会丢：

```go
ppt.NewHyperlink("https://example.com")  // 外部链接
ppt.NewInternalHyperlink(2)               // 链接到第 2 张幻灯片

run.SetHyperlink(link)          // 挂到 run 上
run.GetHyperlink().URL          // https://example.com
run.GetHyperlink().SlideNumber  // 2
```

内部链接用幻灯片序号指定目标。指向不存在的幻灯片时**不写出**，而不是写一条指向不存在部件的
关系。

尚未序列化：用 `shape.BaseShape.SetHyperlink` 设在形状本身上的超链接。它只存在模型里，不会
进文件——形状的 `<p:cNvPr>` 里不会写 `<a:hlinkClick>`，读取端也不找它。请把链接挂在 run 上。

---

### 项目符号 (Bullets)

```go
// 字符符号
bullet := ppt.NewBullet().SetCharBullet("•", "Arial")
bullet.SetColor(ppt.ColorRed).SetSize(120)

// 数字编号
bullet2 := ppt.NewBullet().SetNumericBullet(ppt.NumFormatArabicPeriod, 1)

para.SetBullet(bullet)
```

| 编号格式 | 常量 |
|---|---|
| 1. 2. 3. | `NumFormatArabicPeriod` |
| 1) 2) 3) | `NumFormatArabicParen` |
| I. II. III. | `NumFormatRomanUcPeriod` |
| i. ii. iii. | `NumFormatRomanLcPeriod` |
| A. B. C. | `NumFormatAlphaUcPeriod` |
| a. b. c. | `NumFormatAlphaLcPeriod` |

---

### 批注 (Comments)

```go
author := ppt.NewCommentAuthor("张三", "ZS")
comment := ppt.NewComment()
comment.SetAuthor(author).SetText("请审阅").SetPosition(100, 200)
slide.AddComment(comment)
```

批注支持完整往返。一条批注分两部分存储——正文在 `ppt/comments/commentN.xml`，作者表在
`ppt/commentAuthors.xml`——两者都会写出并读取，因此作者的姓名、缩写以及批注时间戳在保存后
依然保留。正文按 `p:text` 所要求的形态写出（`a:bodyPr` 加若干段 `a:t` 文本运行），这也是
PowerPoint 会读取的形态；读取时仍兼容本库旧版本写出的裸 `<p:text>文本</p:text>` 形式。

---

### 读写 (Writer / Reader)

```go
// 写入文件
w, _ := ppt.NewWriter(p, ppt.WriterPowerPoint2007)
w.(*ppt.PPTXWriter).Save("输出.pptx")

// 写入 io.Writer
var buf bytes.Buffer
w.WriteTo(&buf)

// 从文件读取
reader := &ppt.PPTXReader{}
pres, err := reader.Read("输入.pptx")

// 从 io.ReaderAt 读取
pres, err := reader.ReadFromReader(readerAt, size)

// 便捷包装
pres, err := ppt.Open("输入.pptx")
pres, err := ppt.ReadFrom(readerAt, size)
```

以原生图表部件（`ppt/charts/chartN.xml`）保存的图表会被读回为 `*ppt.ChartShape`，因此图表页可以完整走通「写入 → 读取 → 渲染」回环。详见[图表 (Charts)](#图表-charts)，畸形包的报错方式见[防崩溃保证](#防崩溃保证)。

两个读取入口在遇到畸形包（包括根本不是 PPTX 的文件）时返回 error 而非 panic。

---

### 渲染 (Rendering)

GoPPT 内置渲染器，可将幻灯片导出为图片，渲染效果接近 Microsoft PowerPoint。

```go
reader := &ppt.PPTXReader{}
pres, _ := reader.Read("输入.pptx")

opts := ppt.DefaultRenderOptions()
opts.Width = 1920 // 输出宽度（像素），高度按幻灯片宽高比自动推算

// 单张幻灯片 → 内存图像
img, err := pres.SlideToImage(0, opts)

// 全部幻灯片 → 内存图像
imgs, err := pres.SlidesToImages(opts)

// 单张幻灯片 → 文件（格式由 opts.Format 决定：PNG 或 JPEG）
err = pres.SaveSlideAsImage(0, "slide_1.png", opts)

// 全部幻灯片 → 文件；pattern 必须含 %d，从 1 开始编号
err = pres.SaveSlidesAsImages("slide_%d.png", opts)
```

四个入口都接收幻灯片索引（或遍历全部幻灯片）与可选的 `*RenderOptions`；传入 `nil` 表示使用默认值。`SlideToImage` 返回标准的 `image.Image`。

四者都不会修改你传入的 `*RenderOptions`：默认值是在副本上补齐的，因此同一份 options 可以安全地交给并发的多次渲染。除非你自行提供 `FontCache`，`SlidesToImages` 与 `SaveSlidesAsImages` 会在整次调用中只构造一份缓存，因此字体目录是每次调用扫描一次，而不是每张幻灯片扫描一次。本库无法转成像素的幻灯片尺寸——布局为 nil，或宽高非正——会作为普通错误返回，而不是一个被 recover 的 panic。

#### RenderOptions

| 字段 | 含义 |
| --- | --- |
| `Width` | 输出宽度（像素），默认 960；高度按幻灯片宽高比推算。 |
| `Format` | `ppt.ImageFormatPNG`（默认）或 `ppt.ImageFormatJPEG`。 |
| `JPEGQuality` | JPEG 质量，1-100，默认 90。 |
| `BackgroundColor` | 覆盖幻灯片背景；为 `nil` 时使用幻灯片自身背景或白色。 |
| `DPI` | 字号换算所用 DPI，默认 96。 |
| `FontDirs` | 额外的字体搜索目录；系统字体目录始终会被搜索。 |
| `FontCache` | 在多次渲染之间共享同一个 `*ppt.FontCache`。 |
| `FontFallback` | 文档字体缺失时按顺序尝试的字体名，优先于内置回退链。 |
| `FontDiagnostics` | 收集本次渲染中被替换或未找到的字体。 |
| `OnFontFallback` | 每当某个字体请求无法精确满足时触发的回调。 |
| `OverlayOpacityScale` | 缩放半透明填充的不透明度（1.0 表示不变）。 |
| `Draft` | 在批量预览中以画质换速度：跳过线条与椭圆的抗锯齿、跳过阴影、图片缩放改用最近邻而非双线性。文字仍保留抗锯齿。输出尺寸与内容不变。 |

#### 字体与缺字诊断

当渲染进程拿不到文档所需字体时，单看图片无法判断是文档生成有误还是本机缺字体。传入 `FontDiagnostics` 即可获得被替换或完全找不到的字体清单，也可用 `OnFontFallback` 实时观察每一次回退：

```go
diag := ppt.NewFontDiagnostics()
opts := ppt.DefaultRenderOptions()
opts.Width = 1600
opts.FontDiagnostics = diag
// 把中文固定到已知已安装（或随程序分发）的字体上
opts.FontFallback = []string{"Noto Sans CJK SC"}

if _, err := pres.SlideToImage(0, opts); err != nil {
    return err
}

if diag.UsedBitmapFallback() {
    // 有字体未能找到，已退化为内置位图字体，中文会渲染成豆腐块
    log.Printf("未解析的字体: %v", diag.MissingNames())
}
log.Print(diag.Summary()) // 一行式可读摘要
```

`FontCache.HasFont(name, bold, italic)` 与 `FontCache.FindFontName(name, bold, italic)` 可在渲染前直接查询字体是否可用。

「字体可用」不等于「字形存在」，渲染器对这两者的处理不同。`FontCache.CoversRune(name, bold, italic, r)` 问的是更强的问题——该字体是否有 `r` 的真实字形，答案直接读自字体的字形索引表，索引 0 即 `.notdef`。对于中日韩文字，渲染器按这一判定而非字体名来选 face，因为文档完全可能把一款拉丁字体声明成东亚字体（`<a:ea typeface="Arial"/>`），此时所有按名字的检查都会报告完美匹配，而文字实际画成一个个方框。CJK 文本中常用的带圈符号——`①`、`㈠`、`㎡`——为此也归入东亚类，每个候选 face 都会拿该 run 实际包含的字去校验。结果按（字体, 样式, 字符）缓存；并且当没有任何字体能覆盖 run 中**全部**字符时，覆盖最多的那个字体仍会被选中，因此一个画不出的符号不会把整个 run 推回拉丁字体。

共享 `FontCache` 很重要：构造缓存会扫描字体目录。批量预览时请只构造一次并赋给 `opts.FontCache`——`SlidesToImages` 与 `SaveSlidesAsImages` 在单次调用内都已对多张幻灯片复用同一个缓存。缓存可安全并发使用，因此并行渲染多份文档的 worker pool 可以共用一个。查不到的字体请求同样会被记住，所以引用了未安装字体的文档不会对每个文本 run 重跑一次查找。任何渲染入口都不会回写你传入的 options。

`FontCache.LoadFont(name, path)` 与 `FontCache.LoadFontData(name, data)` 可用你指定的名字注册字体，且该注册的优先级高于目录扫描之后在同名下找到的字体。在启动时注册随程序自带的字体，即使宿主机装有同名字体的另一个版本，渲染也会命中你注册的那一份——这正是跨机器输出可复现的前提。

#### 批量预览与 draft 模式

生成缩略图总览或做批量检查时，请设置 `Draft` 并使用较小的 `Width`。Draft 会跳过线条与椭圆的抗锯齿、跳过阴影，并把图片缩放从双线性改为最近邻。它不改变输出尺寸与内容，只改变精细度——文字仍保留抗锯齿，因为除非替换文本渲染器，字形栅格化无法关闭。

```go
fc := ppt.NewFontCache()      // 只扫描一次，供所有渲染复用
opts := ppt.DefaultRenderOptions()
opts.FontCache = fc
opts.Width = 480
opts.Draft = true

err := pres.SaveSlidesAsImages(filepath.Join(out, "slide_%d.png"), opts)
```

以下为 960 px 渲染的参考耗时（AMD Ryzen 7 8745HS），重点看比值而非绝对值：

| 幻灯片 | 完整 | Draft |
| --- | --- | --- |
| 文字 + 图表 | 692 µs | 406 µs |
| 带阴影形状 | 521 µs | 161 µs |
| 满版图片 | 11.2 ms | 5.0 ms |

Draft 是画质设置，而不是另一个渲染器：需要逐张核对的幻灯片请保持 `Draft` 默认关闭。

#### 不支持的形状

真实文件中总有一些结构无法栅格化：SmartArt 图示、OLE 对象，以及图表部件缺失或无法解析的图表。reader 不会丢弃它们，而是保留一个携带原始框体几何的 `UnsupportedShape`，renderer 在原位置绘制带标注的琥珀色方框，因此预览中不会出现与「正常渲染的空内容」无法区分的空白区域。

```go
for _, s := range pres.UnsupportedShapes() {
    // Label() 形如 "Unsupported: SmartArt diagram"。
    log.Printf("未渲染：%s 于 (%d,%d) %dx%d",
        s.Label(), s.GetOffsetX(), s.GetOffsetY(), s.GetWidth(), s.GetHeight())
}
```

| 方法 | 说明 |
| --- | --- |
| `pres.UnsupportedShapes()` | 返回全文档中不受支持的形状，按幻灯片顺序并递归进入组合。全部可渲染时为空。 |
| `UnsupportedShape.GetReason()` | 简短描述，如 `"SmartArt diagram"` 或 `"chart (its part could not be read)"`。 |
| `UnsupportedShape.GetContentType()` | 来源的 `a:graphicData` uri，便于精确分类。 |
| `UnsupportedShape.Label()` | 绘制在占位框内的文本。 |
| `UnsupportedShape.GetType()` | 恒为 `ppt.ShapeTypeUnsupported`。 |

原始 XML 不会被保留，因此再次写出该文件时会丢失该形状；幻灯片上的其余内容仍可正常往返。标注文字会自动缩小以适配框体；框体小到放不下可读标注时只保留虚线框——标注是本库自己生成的文字，绝不允许溢出到占位框之外。完整的「读取／渲染／占位／退化」对照表见 [README.md](README.md#形状支持矩阵)。

#### 防崩溃保证

`Open`、`Read`、`ReadFromReader`、`GetSlide`、`SlideToImage`、`SlidesToImages`、`SaveSlideAsImage`、`SaveSlidesAsImages` 在解析或栅格化畸形输入时若发生 panic，会将其 recover 并以 `*ppt.PanicError` 返回，而不会让调用方进程崩溃，因此调用侧无需再自行包 `recover`。可用 `ppt.ErrIsPanic(err)` 判断该情形，或用 `errors.As(err, &ppt.PanicError{})` 取出 panic 值与堆栈。

渲染器采用双字体度量架构：HintingNone 字体用于文本排版（匹配 PowerPoint DirectWrite 的度量），HintingFull 字体用于清晰的字形渲染。CJK 文本有专门的处理，包括禁則処理换行规则和优化的行高计算。

图表形状采用原生栅格化，而不会退化为占位图片。分组柱状图、堆积柱状图与百分比堆积柱状图（含横向柱状图）均支持 `GapWidthPercent` 与 `OverlapPercent`；折线图支持平滑曲线与逐系列标记；面积图、饼图、3D 饼图、环形图、散点图与雷达图各有独立绘制路径。3D 图表类型（`bar3D`、`pie3D`）按对应的 2D 图形绘制——渲染器不做 3D 投影。数值轴按 1 / 2 / 5 × 10ⁿ 的“取整”步长推算刻度（也可使用显式上下界与主单位），绘制主网格线与刻度标签，坐标轴标题沿轴线旋转排布。同时支持数据标签、按 `LegendPosition` 自动换行并预留空间的图例，以及由 `srgbClr` / `schemeClr` / `sysClr` / `prstClr` 解析出的数据系列颜色。
