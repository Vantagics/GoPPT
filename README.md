# GoPPT

[English](#english) | [中文](#中文)

---

<a id="english"></a>

## English

A pure Go library for creating, reading, and writing PowerPoint (.pptx) files. Zero external dependencies. Inspired by [PHPOffice/PHPPresentation](https://github.com/PHPOffice/PHPPresentation).

### Rendering

GoPPT includes a built-in slide renderer that produces PNG images closely matching Microsoft PowerPoint's native rendering. The renderer features:

- Dual font-face measurement system (HintingNone for layout, HintingFull for rendering) to match PowerPoint's DirectWrite text metrics
- CJK-aware line wrapping with kinsoku (禁則処理) punctuation handling
- Accurate text box layout with auto-fit, auto-shrink, and overflow control
- Shape rendering: fills, borders, shadows, custom geometry paths, arrowheads
- Chart rendering: bar (clustered / stacked / percent-stacked, vertical or horizontal), line (optionally smoothed), area, pie, doughnut, scatter and radar charts, complete with value/category axes, gridlines, tick labels, axis titles, data labels and configurable legend placement. 3-D chart types (`bar3D`, `pie3D`) are flattened onto their 2-D equivalents — there is no 3-D projection
- Image compositing with rotation, flip, and group transforms

Run `go run ./cmd/chart_preview` to regenerate a sample render of every chart type into `cmd/chart_preview/out/`.

**Chart styling round-trips.** A chart's area fill and outline (`chart.fill`, `chart.border`) and each series' outline (`SeriesOutline` — colour and width, in points) are read from the part, preserved on save, and drawn by the renderer. Where a series' colour lives depends on the chart type, and the writer follows the same rule the reader does: a **stroked** series (line, scatter) carries its colour in `<c:spPr><a:ln>`, which is where PowerPoint reads it, while a **filled** series (bar, pie, area, radar) carries it in `<c:spPr><a:solidFill>`. Writing a line series' colour as a fill would look correct through this library — the reader falls back to the fill when no outline colour is present — while PowerPoint silently dropped it, so the colour is now emitted on the element PowerPoint actually honours.

#### Shape support matrix

| OOXML construct | Read | Render | Notes |
| --- | --- | --- | --- |
| `p:sp` text box / autoshape | yes | yes | includes placeholder shapes |
| `p:pic` picture | yes | yes | PNG, JPEG, GIF, BMP, TIFF, SVG |
| `p:pic` whose reference sits on the Microsoft SVG extension | yes | yes | the (a:blip) may carry no r:embed at all; the only reference is then the nested `asvg:svgBlip r:embed`. A raster reference on `a:blip` always wins over the SVG one, so a fallback bitmap is preferred when both are present |
| `p:pic` containing SVG | yes | yes, for a documented subset | see *SVG pictures* below |
| `p:pic` with `a:srcRect` | yes | yes | crop percentages are applied |
| picture whose data cannot be decoded | yes | labelled placeholder box | the label names the format, or the SVG construct that was refused; EMF/WMF fall back to an extracted bitmap first |
| `p:cxnSp` connector / line | yes | yes | |
| `p:grpSp` group | yes | yes | recursive, with the child-to-group coordinate transform |
| `p:graphicFrame` containing `a:tbl` | yes | yes | |
| `p:graphicFrame` containing a chart part | yes | yes | every chart type listed above |
| chart nested inside a group | yes | yes | |
| 3-D chart types (`bar3D`, `pie3D`) | yes | flattened to 2-D | no 3-D projection |
| `p:graphicFrame` containing SmartArt (`dgm:relIds`) | placeholder | placeholder box | not rasterized |
| `p:graphicFrame` containing an OLE object | placeholder | placeholder box | not rasterized |
| chart whose part is missing or unparseable | placeholder | placeholder box | |

A construct marked *placeholder* is not dropped. The reader keeps an `UnsupportedShape` with the original frame geometry, and the renderer draws a labelled amber box where the graphic was, so a preview never shows an unexplained blank region. Query them with `pres.UnsupportedShapes()`:

```go
if bad := pres.UnsupportedShapes(); len(bad) > 0 {
    for _, s := range bad {
        log.Printf("not rendered: %s at (%d,%d) %dx%d",
            s.Label(), s.GetOffsetX(), s.GetOffsetY(), s.GetWidth(), s.GetHeight())
    }
}
```

Note that the original XML of an unsupported shape is not retained, so writing the deck back out drops it — the alternative would be failing the whole read because of one construct. Everything else on the slide round-trips normally. The caption is shrunk to fit the box, and a box too small for a legible caption keeps only its dashed frame: the label is library-generated text, so it is never allowed to spill over neighbouring content.

#### SVG pictures

PowerPoint stores SVG graphics alongside a raster fallback, and the same `a:blip` that carries a bitmap can carry an SVG instead. Both forms are read, and both are rasterised by the renderer.

There is no SVG engine among the dependencies, so a rasteriser for a **documented subset** is built in. It handles `svg`, `g`, `polygon`, `polyline`, `line`, `rect`, `circle`, `ellipse` and `path`, with `fill`, `stroke`, `stroke-width`, `fill-rule` and opacity, CSS colour syntax, transforms, and `viewBox` with `preserveAspectRatio` (default `meet`, plus `none`). That covers the icon and connector graphics that design-tool chains emit — Font Awesome paths, chevrons, rule lines — which is what the feature is for.

Everything outside the subset is **refused, not mis-drawn**: `text`, `image`, gradients, patterns, clipping, masks, filters, `<use>` and markers all stop the rasterisation with an error naming the construct, and the renderer draws a labelled placeholder instead of whatever partial picture the remaining elements would have made. Two consequences are worth stating plainly:

- **`<defs>` is not drawn.** Its contents are definitions that are only painted where something references them, so a well-formed document may hold a gradient and never use it. Definitions are skipped wherever they appear, and an unused one is not an error.
- **A picture is either rasterised in full or replaced by a placeholder.** There is no partial rendering, because a half-drawn vector graphic reads as a real result while being wrong.

An SVG *can* be written: the bytes go into the media part with the `image/svg+xml` content type. The Microsoft extension element is not emitted, though, so a picture written this way renders correctly in this library — which detects SVG by content, not by declaration — but PowerPoint may not display it. Writing SVG for PowerPoint's benefit is not implemented; supply the raster fallback as the picture's data if the deck has to open correctly there.

Two limits apply when reading a package from an untrusted source. Each part is capped at 50 MB decompressed, and group nesting is capped at 64 levels, because depth is chosen by the file while the renderer and the writer both recurse over the resulting tree — a stack overflow is a fatal error that `recover` cannot turn into a `*PanicError`. Past the cap, the extra levels are parsed and discarded rather than attached to the slide. Real decks nest a handful of levels deep, so the cap only trips on malformed input.

#### Batch previews

Rendering a whole deck calls for cheaper settings than judging a single slide, and two things dominate the cost.

**Share one `FontCache`.** Constructing a cache scans the system font directories, so building one per render repeats that scan every time. Build it once:

```go
fc := ppt.NewFontCache()                      // scans font directories once
opts := ppt.DefaultRenderOptions()
opts.FontCache = fc                           // reused by every render below
opts.Width = 480                              // a preview does not need 960+
opts.Draft = true                             // skip anti-aliasing, shadows, image smoothing

for _, path := range deckPaths {
    pres, err := (&ppt.PPTXReader{}).Read(path)
    if err != nil {
        continue
    }
    if err := pres.SaveSlidesAsImages(filepath.Join(out, base+"_%d.png"), opts); err != nil {
        log.Print(err)
    }
}
```

`SlidesToImages` and `SaveSlidesAsImages` already reuse one cache across the slides of a single call; the example above shares it across files too. Neither modifies the `RenderOptions` you pass in, so one options value can be shared by concurrent renders.

A font registered with `FontCache.LoadFont` or `LoadFontData` takes precedence over one the directory scan finds under the same name, so bundling a font is enough to make rendering reproducible on a machine that has a different build of it installed. This is worth knowing when you ship a CJK font with a preview service: register it at startup and every render resolves to it, not to whatever the host happens to have.

**Turn on `Draft`.** It skips anti-aliasing on lines and ellipses, skips shadows entirely, and scales images with nearest-neighbour instead of bilinear. Text stays anti-aliased, because glyph rasterisation cannot be switched off without replacing the text renderer — a preview still has to be readable. Indicative timings from a 960 px render (AMD Ryzen 7 8745HS), so treat the ratios rather than the absolute numbers as the point:

| Slide | Full | Draft | Draft at 480 px |
| --- | --- | --- | --- |
| Text + chart | 692 µs | 406 µs | 250 µs |
| Shape with shadow | 521 µs | 161 µs | — |
| Full-bleed image | 11.2 ms | 5.0 ms | — |

Draft mode is a fidelity trade, not a different renderer: the output has the same dimensions and the same content. Use it for contact sheets and batch checks, and render the slides you actually inspect at full quality.

When rendering unattended, pass a `FontDiagnostics` to `RenderOptions` to get the list of fonts that were missing or substituted, and set `FontFallback` to pin CJK text to a font you know is installed. See [API.md](API.md) for details.

### Features

- Create and save `.pptx` files (OOXML / PowerPoint 2007+)
- Read existing `.pptx` files with full round-trip support, including charts embedded as native chart parts (`ppt/charts/chartN.xml`)
- Rich text with fonts, colors, bold, italic, underline, strikethrough
- Images (PNG, JPEG, GIF, BMP, SVG) from bytes or file path
- Tables with cell formatting and fills
- Auto shapes (rectangle, ellipse, triangle, arrows, stars, etc.)
- Line shapes with style and color
- Charts: Bar, Bar3D, Line, Area, Pie, Pie3D, Doughnut, Scatter, Radar
- Configurable font fallback chain, plus diagnostics that report every font that was substituted or not found during a render
- East Asian text is matched to a font by **glyph coverage, not by name**. A document whose generator copies one font-family list into both `<a:latin>` and `<a:ea>` declares a Latin face as its East Asian font; the name resolves, so a name-based lookup reports a perfect match while every Chinese character is drawn as that font's `.notdef` box. Each candidate is therefore checked against the characters the run actually contains, and a Latin face is skipped in favour of the fallback chain. Enclosed symbols used inside CJK text (`①`, `㈠`, `㎡`) are classified with it, so they reach the same check instead of going straight to a Latin face
- Public read/render entry points recover panics caused by malformed input and return them as `*PanicError` instead of crashing the process
- Group shapes and Placeholder shapes
- Unsupported OOXML constructs (SmartArt, OLE objects, unreadable chart parts) are kept as visible placeholders and are enumerable via `UnsupportedShapes()`, instead of being silently dropped
- Bullets (character and numeric)
- Comments with authors — name, initials, timestamp and position survive a write → read round trip
- Speaker notes
- Slide backgrounds (solid and gradient)
- Animations (basic grouping)
- Document properties and custom properties
- Multiple slide layouts (4:3, 16:9, 16:10, A4, Letter, custom)
- `go test ./...` covers chart round-trips, font resolution, malformed-input handling and rendering regressions

### Installation

```bash
go get github.com/Vantagics/GoPPT
```

### Quick Start

```go
package main

import (
    "log"
    ppt "github.com/Vantagics/GoPPT"
)

func main() {
    // Create a new presentation
    p := ppt.New()

    // Set document properties
    p.GetDocumentProperties().Title = "My Presentation"
    p.GetDocumentProperties().Creator = "GoPPT"

    // First slide (created automatically)
    slide := p.GetActiveSlide()

    // Add a title
    title := slide.CreateRichTextShape()
    title.SetOffsetX(500000).SetOffsetY(300000)
    title.SetWidth(8000000).SetHeight(1000000)
    tr := title.CreateTextRun("Hello, GoPPT!")
    tr.GetFont().SetSize(28).SetBold(true).SetColor(ppt.ColorBlue)

    // Add a subtitle
    subtitle := slide.CreateRichTextShape()
    subtitle.SetOffsetX(500000).SetOffsetY(1500000)
    subtitle.SetWidth(8000000).SetHeight(600000)
    subtitle.CreateTextRun("Pure Go PowerPoint library")

    // Second slide with a chart
    slide2 := p.CreateSlide()
    chart := slide2.CreateChartShape()
    chart.BaseShape.SetOffsetX(500000).SetOffsetY(500000)
    chart.BaseShape.SetWidth(7000000).SetHeight(4500000)
    chart.GetTitle().SetText("Sales Report")

    bar := ppt.NewBarChart()
    bar.AddSeries(ppt.NewChartSeriesOrdered("Revenue",
        []string{"Q1", "Q2", "Q3", "Q4"},
        []float64{120, 180, 150, 210},
    ))
    chart.GetPlotArea().SetType(bar)

    // Save
    w, _ := ppt.NewWriter(p, ppt.WriterPowerPoint2007)
    if err := w.(*ppt.PPTXWriter).Save("presentation.pptx"); err != nil {
        log.Fatal(err)
    }
}
```

### Reading a Presentation

```go
reader := &ppt.PPTXReader{}
pres, err := reader.Read("presentation.pptx")
if err != nil {
    log.Fatal(err)
}

for i, slide := range pres.GetAllSlides() {
    fmt.Printf("Slide %d: %d shapes\n", i+1, len(slide.GetShapes()))
}
```

### Writing to io.Writer

```go
var buf bytes.Buffer
w, _ := ppt.NewWriter(p, ppt.WriterPowerPoint2007)
w.WriteTo(&buf)
// buf.Bytes() contains the .pptx data
```

### More Examples

See [API Documentation](API.md) for the full reference, or check `example_test.go` for a working example.

---

<a id="中文"></a>

## 中文

纯 Go 语言实现的 PowerPoint (.pptx) 文件创建、读取和写入库。零外部依赖。灵感来自 [PHPOffice/PHPPresentation](https://github.com/PHPOffice/PHPPresentation)。

### 渲染能力

GoPPT 内置幻灯片渲染器，可将幻灯片导出为 PNG 图片，渲染效果接近 Microsoft PowerPoint 原生渲染。渲染器特性：

- 双字体度量系统（HintingNone 用于排版，HintingFull 用于渲染），匹配 PowerPoint DirectWrite 的文本度量
- CJK 感知的自动换行，支持禁則処理（行首行尾标点规则）
- 精确的文本框排版：自动适应、自动缩放、溢出控制
- 形状渲染：填充、边框、阴影、自定义几何路径、箭头
- 图表渲染：柱状图（簇状 / 堆积 / 百分比堆积，纵向或横向）、折线图（可平滑）、面积图、饼图、环形图、散点图、雷达图，包含数值轴 / 分类轴、网格线、刻度标签、坐标轴标题、数据标签，以及可配置的图例位置。3D 图表类型（`bar3D`、`pie3D`）会退化为对应的 2D 图形——渲染器不做 3D 投影
- 图片合成：旋转、翻转、组合变换

运行 `go run ./cmd/chart_preview` 可重新生成全部图表类型的示例渲染图，输出到 `cmd/chart_preview/out/`。

#### 形状支持矩阵

| OOXML 结构 | 读取 | 渲染 | 说明 |
| --- | --- | --- | --- |
| `p:sp` 文本框 / 自动形状 | 支持 | 支持 | 含占位符形状 |
| `p:pic` 图片 | 支持 | 支持 | PNG、JPEG、GIF、BMP、TIFF、SVG |
| `p:pic` 的引用写在微软 SVG 扩展上 | 支持 | 支持 | `a:blip` 上可能完全没有 `r:embed`，唯一的引用是嵌套的 `asvg:svgBlip r:embed`。若 `a:blip` 上有位图引用，则始终优先于 SVG，因此两者都存在时使用回退位图 |
| `p:pic` 含 SVG | 支持 | 支持（有明确子集） | 见下文「SVG 图片」 |
| `p:pic` 带 `a:srcRect` 裁剪 | 支持 | 支持 | 会应用裁剪百分比 |
| 图片数据无法解码 | 支持 | 带标注的占位框 | 标注会写明格式，或被拒绝的 SVG 构造；EMF/WMF 会先尝试提取内嵌位图 |
| `p:cxnSp` 连接线 / 线条 | 支持 | 支持 | |
| `p:grpSp` 组合 | 支持 | 支持 | 递归渲染，并应用子坐标→组合坐标变换 |
| `p:graphicFrame` 含 `a:tbl` 表格 | 支持 | 支持 | |
| `p:graphicFrame` 含图表部件 | 支持 | 支持 | 上文列出的全部图表类型 |
| 组合内嵌的图表 | 支持 | 支持 | |
| 3D 图表类型（`bar3D`、`pie3D`） | 支持 | 退化为 2D | 不做 3D 投影 |
| `p:graphicFrame` 含 SmartArt（`dgm:relIds`） | 占位 | 占位框 | 不做栅格化 |
| `p:graphicFrame` 含 OLE 对象 | 占位 | 占位框 | 不做栅格化 |
| 图表部件缺失或无法解析 | 占位 | 占位框 | |

标记为「占位」的结构不会被丢弃：reader 会保留一个带原始框体几何的 `UnsupportedShape`，renderer 在原位置绘制一个带标注的琥珀色方框，因此预览图中不会出现无法解释的空白区域。可通过 `pres.UnsupportedShapes()` 查询：

```go
if bad := pres.UnsupportedShapes(); len(bad) > 0 {
    for _, s := range bad {
        log.Printf("未渲染：%s 于 (%d,%d) %dx%d",
            s.Label(), s.GetOffsetX(), s.GetOffsetY(), s.GetWidth(), s.GetHeight())
    }
}
```

需要注意：不支持形状的原始 XML 不会被保留，因此再次写出该文件时会丢失该形状——否则一个无关的构造会导致整份文件读取失败。幻灯片上的其余内容仍可正常往返。标注文字会自动缩小以适配框体；框体小到放不下可读的标注时只保留虚线框——标注是本库自己生成的文字，绝不允许溢出到相邻内容上。

#### SVG 图片

PowerPoint 存放 SVG 图形时通常会带一份位图回退；承载位图的同一处 `a:blip` 也可以直接承载 SVG。两种形式都能读取，也都由 renderer 栅格化。

依赖项里没有 SVG 引擎，因此内置了一个针对**明确子集**的栅格化器，支持 `svg`、`g`、`polygon`、`polyline`、`line`、`rect`、`circle`、`ellipse`、`path`，以及 `fill`、`stroke`、`stroke-width`、`fill-rule`、不透明度、CSS 颜色写法、变换，和 `viewBox` 配合 `preserveAspectRatio`（默认 `meet`，也支持 `none`）。这已覆盖设计工具链常导出的图标与连接线图形——Font Awesome 路径、箭头形、分隔线——这正是该功能的目标场景。

子集之外的构造一律**拒绝，而不是画错**：`text`、`image`、渐变、图案、裁剪、遮罩、滤镜、`<use>`、marker 都会让栅格化以错误终止，错误信息会指出具体构造，随后 renderer 绘制带标注的占位框，而不是画出「剩下那些元素凑出来的半张图」。有两点需要说明白：

- **`<defs>` 不会被绘制。** 它的内容是定义，只在被引用处才绘制，因此一份合法文档完全可能定义了一个渐变却从未使用。定义在其出现的任何位置都会被跳过，未使用的定义也不算错误。
- **一张图要么完整栅格化，要么整体换成占位框。** 不存在部分渲染，因为半画出来的矢量图看起来像个真实结果，实际是错的。

SVG **可以**写出：字节会写入 media 部件，Content-Type 为 `image/svg+xml`。但不会写出微软的扩展元素，因此这样写出的图片在本库中渲染正常——本库按内容而非按声明识别 SVG——PowerPoint 却可能不显示它。本库未实现「为 PowerPoint 写出 SVG」；如果文稿必须在那里正确打开，请把位图回退作为图片数据传入。

从不可信来源读取文件时有两处上限：单个部件解压后不超过 50 MB；组合嵌套不超过 64 层。后者是因为嵌套深度完全由文件决定，而 renderer 与 writer 都会对生成的树做递归——栈溢出是致命错误，`recover` 无法把它转成 `*PanicError`。超出上限的层级会被正常解析但不挂到幻灯片上。真实文稿的嵌套只有几层，因此该上限只会被畸形输入触发。

#### 批量预览

渲染整份文档所需的设置比逐张检查幻灯片要省得多，而开销主要来自两处。

**共享同一个 `FontCache`。** 构造缓存会扫描系统字体目录，每次渲染都新建一份等于反复扫描。请只构造一次：

```go
fc := ppt.NewFontCache()                      // 只扫描一次字体目录
opts := ppt.DefaultRenderOptions()
opts.FontCache = fc                           // 后续每次渲染都复用
opts.Width = 480                              // 预览用不到 960+
opts.Draft = true                             // 跳过抗锯齿、阴影与图片平滑

for _, path := range deckPaths {
    pres, err := (&ppt.PPTXReader{}).Read(path)
    if err != nil {
        continue
    }
    if err := pres.SaveSlidesAsImages(filepath.Join(out, base+"_%d.png"), opts); err != nil {
        log.Print(err)
    }
}
```

`SlidesToImages` 与 `SaveSlidesAsImages` 在单次调用内本就会对多张幻灯片复用同一个缓存；上面的写法把缓存也跨文件复用了。两者都不会修改你传入的 `RenderOptions`，因此同一份 options 可以安全地交给并发的多次渲染。

用 `FontCache.LoadFont` 或 `LoadFontData` 注册的字体优先级高于目录扫描在同名下找到的字体，因此只要随程序自带字体，就能在装有不同版本该字体的机器上得到一致的渲染结果。对随预览服务一起分发中文字体的场景尤其有意义：启动时注册，之后每次渲染都命中它，而不是命中宿主机上碰巧存在的那个。

**开启 `Draft`。** 它会跳过线条与椭圆的抗锯齿、完全跳过阴影，并把图片缩放从双线性改为最近邻。文字仍保留抗锯齿——除非替换文本渲染器，否则字形栅格化无法关闭，而预览必须保持可读。以下为 960 px 渲染的参考耗时（AMD Ryzen 7 8745HS），重点看比值而非绝对值：

| 幻灯片 | 完整 | Draft | Draft + 480 px |
| --- | --- | --- | --- |
| 文字 + 图表 | 692 µs | 406 µs | 250 µs |
| 带阴影形状 | 521 µs | 161 µs | — |
| 满版图片 | 11.2 ms | 5.0 ms | — |

Draft 是画质取舍，而不是另一个渲染器：输出尺寸与内容都一致。适合缩略图总览与批量检查；需要逐张核对的幻灯片请用完整画质渲染。

批量无人值守渲染时，可为 `RenderOptions` 传入 `FontDiagnostics` 以获取缺失/被替换的字体清单，并用 `FontFallback` 把中文固定到已知已安装的字体上。详见 [API.md](API.md)。

### 功能特性

- 创建和保存 `.pptx` 文件（OOXML / PowerPoint 2007+）
- 读取现有 `.pptx` 文件，支持完整的读写往返，包括以原生图表部件（`ppt/charts/chartN.xml`）形式嵌入的图表
- 富文本：字体、颜色、粗体、斜体、下划线、删除线
- 图片（PNG、JPEG、GIF、BMP、SVG），支持字节数据或文件路径
- 表格，支持单元格格式和填充
- 自动形状（矩形、椭圆、三角形、箭头、星形等）
- 线条形状，支持样式和颜色
- 图表：柱状图、3D柱状图、折线图、面积图、饼图、3D饼图、环形图、散点图、雷达图
- 可配置的字体回退链，并能报告每次渲染中被替换或未找到的字体
- 中日韩文字按**字形覆盖**而非字体名匹配。有些生成器会把同一份 font-family 列表同时写进 `<a:latin>` 和 `<a:ea>`，于是把一款拉丁字体声明成了「东亚字体」；字体名能解析，因此按名字查找会报告完美匹配，而每个汉字实际都画成了该字体的 `.notdef` 方框。所以每个候选字体都会拿该 run 实际包含的字去校验，拉丁字体会被跳过并改用回退链。CJK 文本中常用的带圈符号（`①`、`㈠`、`㎡`）也一并纳入该判定，从而进入同一套覆盖检查，而不是直接落到拉丁字体上
- 公共读取／渲染入口在遇到畸形输入时会 recover panic 并以 `*PanicError` 返回，不会让进程崩溃
- 组合形状和占位符形状
- 不支持的 OOXML 结构（SmartArt、OLE 对象、无法读取的图表部件）会保留为可见占位框，并可通过 `UnsupportedShapes()` 枚举，而不是被静默丢弃
- 项目符号（字符和数字编号）
- 批注（含作者信息）——作者姓名、缩写、时间戳与位置均可完整走通「写入 → 读取」往返
- 演讲者备注
- 幻灯片背景（纯色和渐变）
- 动画（基础分组）
- 文档属性和自定义属性
- 多种幻灯片布局（4:3、16:9、16:10、A4、Letter、自定义）
- `go test ./...` 覆盖图表读写往返、字体解析、畸形输入处理与渲染回归

### 安装

```bash
go get github.com/Vantagics/GoPPT
```

### 快速开始

```go
package main

import (
    "log"
    ppt "github.com/Vantagics/GoPPT"
)

func main() {
    // 创建新演示文稿
    p := ppt.New()

    // 设置文档属性
    p.GetDocumentProperties().Title = "我的演示文稿"
    p.GetDocumentProperties().Creator = "GoPPT"

    // 第一张幻灯片（自动创建）
    slide := p.GetActiveSlide()

    // 添加标题
    title := slide.CreateRichTextShape()
    title.SetOffsetX(500000).SetOffsetY(300000)
    title.SetWidth(8000000).SetHeight(1000000)
    tr := title.CreateTextRun("你好，GoPPT！")
    tr.GetFont().SetSize(28).SetBold(true).SetColor(ppt.ColorBlue)

    // 添加副标题
    subtitle := slide.CreateRichTextShape()
    subtitle.SetOffsetX(500000).SetOffsetY(1500000)
    subtitle.SetWidth(8000000).SetHeight(600000)
    subtitle.CreateTextRun("纯 Go 语言 PowerPoint 库")

    // 第二张幻灯片：图表
    slide2 := p.CreateSlide()
    chart := slide2.CreateChartShape()
    chart.BaseShape.SetOffsetX(500000).SetOffsetY(500000)
    chart.BaseShape.SetWidth(7000000).SetHeight(4500000)
    chart.GetTitle().SetText("销售报告")

    bar := ppt.NewBarChart()
    bar.AddSeries(ppt.NewChartSeriesOrdered("营收",
        []string{"第一季度", "第二季度", "第三季度", "第四季度"},
        []float64{120, 180, 150, 210},
    ))
    chart.GetPlotArea().SetType(bar)

    // 保存
    w, _ := ppt.NewWriter(p, ppt.WriterPowerPoint2007)
    if err := w.(*ppt.PPTXWriter).Save("演示文稿.pptx"); err != nil {
        log.Fatal(err)
    }
}
```

### 读取演示文稿

```go
reader := &ppt.PPTXReader{}
pres, err := reader.Read("演示文稿.pptx")
if err != nil {
    log.Fatal(err)
}

for i, slide := range pres.GetAllSlides() {
    fmt.Printf("幻灯片 %d：%d 个形状\n", i+1, len(slide.GetShapes()))
}
```

### 写入到 io.Writer

```go
var buf bytes.Buffer
w, _ := ppt.NewWriter(p, ppt.WriterPowerPoint2007)
w.WriteTo(&buf)
// buf.Bytes() 包含 .pptx 数据
```

### 更多示例

请参阅 [API 文档](API.md) 获取完整参考，或查看 `example_test.go` 获取可运行的示例。

---

## License

MIT
