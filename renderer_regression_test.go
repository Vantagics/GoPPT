package gopresentation

import (
	"image"
	"image/color"
	"strings"
	"testing"
	"unicode/utf8"

	"golang.org/x/image/font"
	"golang.org/x/image/font/basicfont"
)

// Structural rendering assertions.
//
// Unlike the golden images in golden_test.go, these checks do not depend on
// which fonts are installed, so they must hold on every machine. They cover the
// complex paths the README calls out: kinsoku line breaking, rotation, shadows
// and tables.

// --- kinsoku (禁則処理) ------------------------------------------------------

// TestKinsokuCharacterClasses pins the character sets that drive kinsoku, so a
// change to them is caught without needing to render anything.
func TestKinsokuCharacterClasses(t *testing.T) {
	// Closing punctuation must not start a line.
	for _, r := range []rune{'。', '，', '、', '；', '：', '！', '？', '…',
		'）', '】', '》', '」', '』', '〕', '｝', '］', ')', ']', '}', '>', '.', ',', ';', ':', '!', '?'} {
		if !isCJKClosingPunct(r) {
			t.Errorf("isCJKClosingPunct(%q) = false, want true", r)
		}
	}
	// Opening punctuation must not end a line.
	for _, r := range []rune{'（', '【', '《', '「', '『', '〈', '〔', '｛', '［',
		'(', '[', '{', '<', '\u201C', '\u2018'} {
		if !isCJKOpeningPunct(r) {
			t.Errorf("isCJKOpeningPunct(%q) = false, want true", r)
		}
	}
	// Ordinary characters belong to neither class.
	for _, r := range []rune{'a', 'Z', '0', '中', '文', '한'} {
		if isCJKClosingPunct(r) {
			t.Errorf("isCJKClosingPunct(%q) = true, want false", r)
		}
		if isCJKOpeningPunct(r) {
			t.Errorf("isCJKOpeningPunct(%q) = true, want false", r)
		}
	}
	// A run counts as punctuation only when every glyph is punctuation.
	if !isClosingPunctRun("。）") {
		t.Error(`isClosingPunctRun("。）") = false, want true`)
	}
	if isClosingPunctRun("。a") {
		t.Error(`isClosingPunctRun("。a") = true, want false`)
	}
	if isClosingPunctRun("") {
		t.Error(`isClosingPunctRun("") = true, want false`)
	}
}

// wrapRunsForTest wraps text runs with a deterministic built-in face, so the
// result does not depend on installed fonts.
func wrapRunsForTest(runs []textRun, maxWidth int) []string {
	r := &renderer{}
	lines := r.wrapRunLine(runs, maxWidth)
	out := make([]string, 0, len(lines))
	for _, ln := range lines {
		var sb strings.Builder
		for _, run := range ln.runs {
			sb.WriteString(run.text)
		}
		out = append(out, sb.String())
	}
	return out
}

func asciiRun(text string) textRun {
	face := basicfont.Face7x13
	return textRun{
		text:        text,
		font:        NewFont(),
		face:        face,
		measureFace: face,
		width:       faceWidth(face, text),
	}
}

func faceWidth(face font.Face, text string) int {
	return measureStringWithKern(face, text).Ceil()
}

// TestKinsokuClosingPunctStaysOnPreviousLine drives the wrapping code with a
// full line followed by a lone closing bracket: the bracket must be kept on the
// current line rather than starting a new one.
func TestKinsokuClosingPunctStaysOnPreviousLine(t *testing.T) {
	// basicfont.Face7x13 advances 7px per ASCII glyph, so 20 characters are
	// 140px wide. A 141px limit leaves room for nothing more.
	body := asciiRun(strings.Repeat("a", 20))
	if got := faceWidth(body.face, body.text); got != 140 {
		t.Skipf("basicfont metrics changed (width=%d, want 140); adjust the fixture", got)
	}
	runs := []textRun{body, asciiRun(")")}

	lines := wrapRunsForTest(runs, 141)
	if len(lines) != 1 {
		t.Fatalf("lines = %q, want the bracket kept on the first line", lines)
	}
	if !strings.HasSuffix(lines[0], ")") {
		t.Errorf("line = %q, want it to end with the closing bracket", lines[0])
	}
}

// TestWrapBreaksAtWidth is a sanity check that wrapping still happens at all, so
// the kinsoku test above cannot pass vacuously. It uses space-separated words
// because a single long Latin word is deliberately not broken mid-word.
func TestWrapBreaksAtWidth(t *testing.T) {
	const words = "alpha bravo charlie delta echo foxtrot golf hotel india"
	runs := []textRun{asciiRun(words)}
	// Six 7px glyphs plus a space fit per line at 49px.
	lines := wrapRunsForTest(runs, 49)
	if len(lines) < 3 {
		t.Fatalf("lines = %q, want at least 3 lines at 49px", lines)
	}
	for i, ln := range lines {
		if w := faceWidth(basicfont.Face7x13, ln); w > 49 {
			t.Errorf("line %d is %dpx wide, exceeding the 49px limit: %q", i, w, ln)
		}
	}
	// Wrapping must not duplicate or drop text.
	if got := strings.Join(strings.Fields(strings.Join(lines, " ")), " "); got != words {
		t.Errorf("wrapped text = %q, want %q", got, words)
	}
}

// TestCJKWrapNoLineStartsByPunctuation checks the real thing: a Chinese
// paragraph in a narrow box must never begin a line with punctuation that is
// prohibited at line start.
func TestCJKWrapNoLineStartsByPunctuation(t *testing.T) {
	fc := NewFontCache()
	requireCJKFont(t, fc)

	const text = "本季度营收同比增长百分之十八，主要来自华东与华南地区的新增客户（含两家制造业龙头）。" +
		"成本端，原材料价格回落，毛利率提升至百分之四十二。"

	p := goldenFixtureCJKWrap()
	pres := p
	slide := pres.GetAllSlides()[0]

	// Reuse the renderer's own run construction so the test exercises the real
	// path (CJK splitting included) rather than a hand-built approximation. The
	// scale must match a real render: point sizes are converted with it, and a
	// wrong scale produces an absurd face size that fails to build. It is taken
	// from the presentation's own slide size, never assumed.
	const renderWidthPx = 640
	layout := pres.GetLayout()
	if layout == nil || layout.CX <= 0 {
		t.Fatal("fixture presentation has no slide size")
	}
	r := &renderer{
		fontCache:    fc,
		scaleX:       float64(renderWidthPx) / float64(layout.CX),
		dpi:          96,
		fontFallback: mergeFallbackChain(nil, defaultFontFallbackChain),
		cjkFallback:  mergeFallbackChain(nil, defaultCJKFallbackChain),
	}
	shape, ok := slide.GetShapes()[0].(*RichTextShape)
	if !ok {
		t.Fatalf("fixture shape type = %T, want *RichTextShape", slide.GetShapes()[0])
	}
	runs := r.buildParaTextRuns(shape.GetActiveParagraph().GetElements())
	if len(runs) == 0 {
		t.Fatal("no text runs built from the fixture")
	}

	// Derive a wrap width from a real font so the line breaks land mid-text.
	probe := runs[0].mface()
	if probe == nil {
		t.Fatal("no measure face; cannot lay out text")
	}
	charWidth := faceWidth(probe, "中")
	if charWidth <= 0 {
		t.Skip("measure face cannot measure CJK characters")
	}
	maxWidth := charWidth * 16

	lines := r.wrapRunLine(runs, maxWidth)
	if len(lines) < 2 {
		t.Fatalf("expected the paragraph to wrap into several lines at %dpx, got %d", maxWidth, len(lines))
	}

	for i, ln := range lines {
		var sb strings.Builder
		for _, run := range ln.runs {
			sb.WriteString(run.text)
		}
		lineText := sb.String()
		if lineText == "" {
			continue
		}
		first, _ := utf8.DecodeRuneInString(lineText)
		if isCJKClosingPunct(first) {
			t.Errorf("line %d starts with line-start-prohibited punctuation %q: %q", i, first, lineText)
		}
	}

	// The text must survive wrapping intact.
	var got strings.Builder
	for _, ln := range lines {
		for _, run := range ln.runs {
			got.WriteString(run.text)
		}
	}
	if got.String() != text {
		t.Errorf("wrapped text differs from the source:\n got %q\nwant %q", got.String(), text)
	}
}

// --- shape transforms -------------------------------------------------------

// emuRect converts an EMU rectangle into pixels for a slide rendered at
// widthPx. Tests must derive geometry this way rather than assuming a slide
// size: New() creates a 4:3 presentation, not 16:9.
func emuRect(t *testing.T, pres *Presentation, x, y, w, h int64, widthPx int) image.Rectangle {
	t.Helper()
	layout := pres.GetLayout()
	if layout == nil || layout.CX <= 0 || layout.CY <= 0 {
		t.Fatalf("presentation has no usable slide size: %+v", layout)
	}
	scaleX := float64(widthPx) / float64(layout.CX)
	imgH := float64(widthPx) * float64(layout.CY) / float64(layout.CX)
	scaleY := imgH / float64(layout.CY)
	return image.Rect(
		int(float64(x)*scaleX), int(float64(y)*scaleY),
		int(float64(x+w)*scaleX), int(float64(y+h)*scaleY),
	)
}

// inkBounds returns the bounding box of non-white pixels, or the empty
// rectangle when the image is blank.
func inkBounds(img image.Image) image.Rectangle {
	b := img.Bounds()
	minX, minY := b.Max.X, b.Max.Y
	maxX, maxY := b.Min.X, b.Min.Y
	found := false
	for y := b.Min.Y; y < b.Max.Y; y++ {
		for x := b.Min.X; x < b.Max.X; x++ {
			r, g, bb, a := img.At(x, y).RGBA()
			if a>>8 < 8 {
				continue
			}
			if r>>8 < 250 || g>>8 < 250 || bb>>8 < 250 {
				found = true
				if x < minX {
					minX = x
				}
				if y < minY {
					minY = y
				}
				if x > maxX {
					maxX = x
				}
				if y > maxY {
					maxY = y
				}
			}
		}
	}
	if !found {
		return image.Rectangle{}
	}
	return image.Rect(minX, minY, maxX+1, maxY+1)
}

// TestRotationChangesInkExtent verifies rotation is actually applied: a wide,
// short bar rotated by 90 degrees must end up taller than it is wide, with the
// extents swapped. The shape is centred and small enough that the rotated
// result stays on the slide, so this measures the rotation rather than the clip.
func TestRotationChangesInkExtent(t *testing.T) {
	fc := NewFontCache()
	opts := goldenOptions(fc)

	newShape := func(rotation int) *Presentation {
		p := New()
		sh := p.GetActiveSlide().CreateAutoShape()
		sh.SetAutoShapeType(AutoShapeRectangle)
		sh.BaseShape.SetOffsetX(4096000).SetOffsetY(2979000)
		sh.BaseShape.SetWidth(4000000).SetHeight(900000)
		sh.BaseShape.SetRotation(rotation)
		sh.SetSolidFill(NewColor("4472C4"))
		return p
	}

	img, err := newShape(0).SlideToImage(0, opts)
	if err != nil {
		t.Fatalf("render unrotated: %v", err)
	}
	base := inkBounds(img)
	if base.Empty() {
		t.Fatal("unrotated shape rendered nothing")
	}
	if base.Dx() <= base.Dy() {
		t.Fatalf("unrotated shape ink is %v, want wider than tall", base)
	}
	if base.Min.X <= 0 || base.Min.Y <= 0 || base.Max.X >= img.Bounds().Max.X || base.Max.Y >= img.Bounds().Max.Y {
		t.Fatalf("unrotated shape touches the image edge (%v in %v); the fixture cannot measure rotation", base, img.Bounds())
	}

	rimg, err := newShape(90).SlideToImage(0, opts)
	if err != nil {
		t.Fatalf("render rotated: %v", err)
	}
	rbox := inkBounds(rimg)
	if rbox.Empty() {
		t.Fatal("rotated shape rendered nothing")
	}
	if rbox.Dy() <= rbox.Dx() {
		t.Errorf("rotated shape ink is %v, want taller than wide after a 90 degree rotation", rbox)
	}
	if rbox.Min.Y <= 0 || rbox.Max.Y >= rimg.Bounds().Max.Y {
		t.Fatalf("rotated shape is clipped (%v in %v); the fixture cannot measure rotation", rbox, rimg.Bounds())
	}
	// A 90 degree rotation swaps the extents, so the one should approximate the
	// other. Allow a few pixels for anti-aliasing at the edges.
	if absInt(rbox.Dx()-base.Dy()) > 6 || absInt(rbox.Dy()-base.Dx()) > 6 {
		t.Errorf("rotation did not swap extents: unrotated %v (w=%d,h=%d), rotated %v (w=%d,h=%d)",
			base, base.Dx(), base.Dy(), rbox, rbox.Dx(), rbox.Dy())
	}
}

func absInt(v int) int {
	if v < 0 {
		return -v
	}
	return v
}

// TestShadowDrawsOutsideShape verifies a shadow paints outside the shape's own
// rectangle. It compares against the same shape with no shadow, so it does not
// depend on the shadow's direction convention.
func TestShadowDrawsOutsideShape(t *testing.T) {
	fc := NewFontCache()
	opts := goldenOptions(fc)

	without := New()
	plain := without.GetActiveSlide().CreateAutoShape()
	plain.SetAutoShapeType(AutoShapeRectangle)
	plain.BaseShape.SetOffsetX(2000000).SetOffsetY(1500000)
	plain.BaseShape.SetWidth(4000000).SetHeight(2500000)
	plain.SetSolidFill(NewColor("ED7D31"))
	plainImg, err := without.SlideToImage(0, opts)
	if err != nil {
		t.Fatalf("render without shadow: %v", err)
	}

	// Pixel rectangle of the shape, derived from the presentation's slide size.
	shapeRect := emuRect(t, without, 2000000, 1500000, 4000000, 2500000, opts.Width)
	// A border ring just outside the shape, where only a shadow can draw.
	ring := image.Rect(shapeRect.Min.X-2, shapeRect.Min.Y-2, shapeRect.Max.X+24, shapeRect.Max.Y+24)

	with := New()
	shadowed := with.GetActiveSlide().CreateAutoShape()
	shadowed.SetAutoShapeType(AutoShapeRectangle)
	shadowed.BaseShape.SetOffsetX(2000000).SetOffsetY(1500000)
	shadowed.BaseShape.SetWidth(4000000).SetHeight(2500000)
	shadowed.SetSolidFill(NewColor("ED7D31"))
	sd := NewShadow()
	sd.Visible = true
	sd.Direction = 45
	sd.Distance = 14
	sd.BlurRadius = 10
	sd.Color = NewColor("000000")
	sd.Alpha = 70
	shadowed.BaseShape.SetShadow(sd)
	shadowImg, err := with.SlideToImage(0, opts)
	if err != nil {
		t.Fatalf("render with shadow: %v", err)
	}

	plainRing := countInkIn(plainImg, ring)
	shadowRing := countInkIn(shadowImg, ring)
	if shadowRing <= plainRing {
		t.Errorf("shadow added no ink outside the shape: %d -> %d inked pixels in %v",
			plainRing, shadowRing, ring)
	}
}

// countInkIn counts non-white pixels inside rect.
func countInkIn(img image.Image, rect image.Rectangle) int {
	rect = rect.Intersect(img.Bounds())
	n := 0
	for y := rect.Min.Y; y < rect.Max.Y; y++ {
		for x := rect.Min.X; x < rect.Max.X; x++ {
			r, g, b, a := img.At(x, y).RGBA()
			if a>>8 < 8 {
				continue
			}
			if r>>8 < 250 || g>>8 < 250 || b>>8 < 250 {
				n++
			}
		}
	}
	return n
}

// --- tables -----------------------------------------------------------------

// TestTableRendersEveryCell verifies the table path draws content for every
// cell; a table that silently renders as an empty frame is a common regression.
func TestTableRendersEveryCell(t *testing.T) {
	fc := NewFontCache()
	requireAnyFont(t, fc)

	pres := goldenFixtureTable()
	opts := goldenOptions(fc)
	img, err := pres.SlideToImage(0, opts)
	if err != nil {
		t.Fatalf("render: %v", err)
	}

	tbl, ok := pres.GetAllSlides()[0].GetShapes()[0].(*TableShape)
	if !ok {
		t.Fatalf("fixture shape type = %T, want *TableShape", pres.GetAllSlides()[0].GetShapes()[0])
	}

	// Sampling geometry comes from the presentation's own slide size: New()
	// builds a 4:3 slide, so assuming 16:9 would place the window off by the
	// ratio and could mask an empty row.
	box := emuRect(t, pres, 400000, 400000, 8000000, 4000000, opts.Width)
	x0, y0 := box.Min.X, box.Min.Y
	w, h := box.Dx(), box.Dy()
	cellW := w / tbl.GetNumCols()
	cellH := h / tbl.GetNumRows()

	for r := 0; r < tbl.GetNumRows(); r++ {
		for c := 0; c < tbl.GetNumCols(); c++ {
			cell := image.Rect(
				x0+c*cellW, y0+r*cellH,
				x0+(c+1)*cellW, y0+(r+1)*cellH,
			)
			if countInkIn(img, cell) == 0 {
				t.Errorf("table cell (%d,%d) rendered nothing in %v", r, c, cell)
			}
		}
	}
}

// --- text box basics --------------------------------------------------------

// TestTranslucentTextLightensItsBackground guards the colour convention at the
// image/draw boundary.
//
// Colour values inside the renderer carry straight alpha, because blendPixel
// multiplies by that alpha itself. color.RGBA, however, is defined by Go as
// alpha-premultiplied, and image/draw and font.Drawer rely on that: a
// straight-alpha white at 12% handed to font.Drawer has channels far above its
// alpha, the fixed-point blend overflows the byte it writes into, and the glyph
// comes out *darker* than the surface it was drawn on. That is exactly what
// happened to the oversized translucent section numerals on the real deck this
// was found with — they rendered as dark outlines on a dark background instead
// of as faint light watermarks.
//
// The assertion is directional rather than numeric: whichever colour a
// translucent fill produces, it must move the pixel toward that colour and not
// away from it.
func TestTranslucentTextLightensItsBackground(t *testing.T) {
	fc := NewFontCache()
	requireAnyFont(t, fc)

	// A dark slide with white text at 12% opacity, the shape of the numeral on
	// the deck's section dividers.
	background := color.RGBA{R: 6, G: 32, B: 50, A: 255}
	pres := New()
	shape := pres.GetActiveSlide().CreateRichTextShape()
	shape.BaseShape.SetOffsetX(800000).SetOffsetY(800000)
	shape.BaseShape.SetWidth(7000000).SetHeight(2000000)
	run := shape.CreateTextRun("Ag")
	run.GetFont().SetSize(72).SetColor(NewColor("1EFFFFFF"))

	opts := goldenOptions(fc)
	opts.BackgroundColor = &background
	img, err := pres.SlideToImage(0, opts)
	if err != nil {
		t.Fatalf("render: %v", err)
	}
	rgba := toRGBA(t, img)

	box := emuRect(t, pres, 800000, 800000, 7000000, 2000000, opts.Width)
	box = box.Intersect(rgba.Bounds())
	bgLuma := luma(background)

	inked, darker := 0, 0
	for y := box.Min.Y; y < box.Max.Y; y++ {
		for x := box.Min.X; x < box.Max.X; x++ {
			c := rgba.RGBAAt(x, y)
			if c == background {
				continue
			}
			inked++
			if luma(c) < bgLuma {
				darker++
			}
		}
	}
	if inked == 0 {
		t.Fatal("the translucent text drew no pixels at all")
	}
	if darker > 0 {
		t.Errorf("%d of %d glyph pixels are darker than the background they were drawn over "+
			"(background luma %d); white at 12%% must produce a LIGHTER pixel, so the "+
			"straight-alpha colour is reaching image/draw as if it were premultiplied",
			darker, inked, bgLuma)
	}
}

// luma returns the perceived brightness of an opaque colour.
func luma(c color.RGBA) int {
	return (299*int(c.R) + 587*int(c.G) + 114*int(c.B)) / 1000
}

// TestBorderWidthRoundTrip guards a units bug: Border.Width is stored in points
// (the reader divides the EMU it reads by 12700), so the writer has to emit
// points*12700 EMU. Emitting the raw value made a border collapse to about
// 1/12700 of its intended width on every write/read cycle — a 1pt border came
// back as 1/12700pt and rendered as nothing.
func TestBorderWidthRoundTrip(t *testing.T) {
	p := New()
	sh := p.GetActiveSlide().CreateAutoShape()
	sh.SetAutoShapeType(AutoShapeRectangle)
	sh.BaseShape.SetOffsetX(500000).SetOffsetY(500000)
	sh.BaseShape.SetWidth(4000000).SetHeight(2000000)
	sh.SetSolidFill(NewColor("FFFFFF"))

	border := NewBorder()
	border.SetSolidFill(NewColor("FF0000"))
	border.SetWidth(2)
	sh.BaseShape.SetBorder(border)

	path := writeChartPPTX(t, p)
	pres, err := Open(path)
	if err != nil {
		t.Fatalf("Open: %v", err)
	}
	as, ok := pres.GetAllSlides()[0].GetShapes()[0].(*AutoShape)
	if !ok {
		t.Fatalf("shape type = %T, want *AutoShape", pres.GetAllSlides()[0].GetShapes()[0])
	}
	got := as.GetBorder()
	if got == nil || got.Style != BorderSolid {
		t.Fatalf("border not restored after round trip: %+v", got)
	}
	if got.Width != 2 {
		t.Errorf("border width = %d, want 2 (points); the writer must emit W*12700 EMU", got.Width)
	}
}

// TestBorderRendersAtRequestedWeight checks the fix end to end: a 1pt border
// must be thin, not a filled blob. The units bug made it thousands of pixels
// thick, which turned the shape into a solid rectangle of the border colour.
func TestBorderRendersAtRequestedWeight(t *testing.T) {
	fc := NewFontCache()
	opts := goldenOptions(fc)

	p := New()
	sh := p.GetActiveSlide().CreateAutoShape()
	sh.SetAutoShapeType(AutoShapeRectangle)
	sh.BaseShape.SetOffsetX(2000000).SetOffsetY(1500000)
	sh.BaseShape.SetWidth(4000000).SetHeight(2500000)
	sh.SetSolidFill(NewColor("FFFFFF"))
	border := NewBorder()
	border.SetSolidFill(NewColor("FF0000"))
	border.SetWidth(1)
	sh.BaseShape.SetBorder(border)

	img, err := p.SlideToImage(0, opts)
	if err != nil {
		t.Fatalf("render: %v", err)
	}

	// Pixel rectangle of the shape, derived from the presentation's own slide
	// size (New() is 4:3, not 16:9).
	shapeRect := emuRect(t, p, 2000000, 1500000, 4000000, 2500000, opts.Width)

	// The interior must stay white: if the thick-border bug regresses, the
	// border colour floods the whole rectangle.
	interior := shapeRect.Inset(8)
	if n := countRedIn(img, interior); n > 0 {
		t.Errorf("%d red pixels inside the shape interior (%v); a 1pt border should not fill it", n, interior)
	}
	// But the border itself must be drawn.
	edge := image.Rect(shapeRect.Min.X-2, shapeRect.Min.Y-2, shapeRect.Max.X+2, shapeRect.Min.Y+3)
	if n := countRedIn(img, edge); n == 0 {
		t.Errorf("no border pixels found along the top edge (%v)", edge)
	}
}

// countRedIn counts pixels that are predominantly red inside rect.
func countRedIn(img image.Image, rect image.Rectangle) int {
	rect = rect.Intersect(img.Bounds())
	n := 0
	for y := rect.Min.Y; y < rect.Max.Y; y++ {
		for x := rect.Min.X; x < rect.Max.X; x++ {
			r, g, b, _ := img.At(x, y).RGBA()
			r8, g8, b8 := r>>8, g>>8, b>>8
			if r8 > 150 && g8 < 120 && b8 < 120 {
				n++
			}
		}
	}
	return n
}

// TestEmptySlideHasNoInk guards the simplest possible regression: rendering a
// slide with nothing on it must produce a blank image, not stray marks.
func TestEmptySlideHasNoInk(t *testing.T) {
	fc := NewFontCache()
	opts := goldenOptions(fc)
	img, err := New().SlideToImage(0, opts)
	if err != nil {
		t.Fatalf("render: %v", err)
	}
	if ink := inkBounds(img); !ink.Empty() {
		t.Errorf("empty slide produced ink at %v", ink)
	}
}

// TestSlideIndexOutOfRange checks the render entry points reject bad indices
// with an error rather than panicking.
func TestSlideIndexOutOfRange(t *testing.T) {
	p := New()
	opts := DefaultRenderOptions()
	opts.Width = 160

	for _, idx := range []int{-1, 1, 99} {
		if _, err := p.SlideToImage(idx, opts); err == nil {
			t.Errorf("SlideToImage(%d) = nil error, want error", idx)
		}
	}
	if err := p.SaveSlideAsImage(5, t.TempDir()+"/x.png", opts); err == nil {
		t.Error("SaveSlideAsImage(5) = nil error, want error")
	}
}

// TestSlidesToImagesDoesNotMutateOptions pins that rendering leaves the caller's
// options untouched. The derived FontCache used to be written back into the
// caller's struct, which is both surprising and a data race when one set of
// options covers two presentations rendered concurrently.
func TestSlidesToImagesDoesNotMutateOptions(t *testing.T) {
	p := New()
	shape := p.GetActiveSlide().CreateRichTextShape()
	shape.BaseShape.SetOffsetX(400000).SetOffsetY(400000)
	shape.BaseShape.SetWidth(4000000).SetHeight(1000000)
	shape.CreateTextRun("hello").GetFont().SetSize(18)

	opts := DefaultRenderOptions()
	opts.Width = 320

	if _, err := p.SlidesToImages(opts); err != nil {
		t.Fatalf("SlidesToImages: %v", err)
	}
	if opts.FontCache != nil {
		t.Error("SlidesToImages wrote its derived FontCache into the caller's options")
	}
	if opts.Width != 320 {
		t.Errorf("SlidesToImages changed opts.Width to %d, want 320", opts.Width)
	}
}
