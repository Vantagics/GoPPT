package gopresentation

import (
	"image"
	"testing"
)

// Bullet rendering — the half that had never been drawn in a test.
//
// bullet.go, reader_slide.go and writer_slide.go all carry a per-paragraph
// Bullet with StartAt and Size, and bullets_test.go covers the file half of
// that round trip. The drawing half had no coverage at all (buildBulletRun sat
// at 0%), and it read StartAt verbatim for every paragraph: because PowerPoint
// stores one <a:buAutoNum> per paragraph and counts the list implicitly, a
// three-item numbered list rendered "1. 1. 1." while the file said
// startAt="1" on the first item and left the rest to PowerPoint.
//
// Bullet.Size (<a:buSzPct>, a percentage of the text size) was read and written
// and then never applied, so a bullet set to 200% drew at 100% in the preview.
// An empty Bullet.Style was a third, smaller disagreement: the writer
// substituted "•" and the renderer drew a blank.

// numberedListPresentation builds one shape holding one paragraph per item,
// each carrying a numbered bullet of the given format from the given start
// number. The shape is given the full slide width with wrapping off so that
// every item occupies exactly one text line — otherwise a narrow default box
// wraps the bullet away from its text and the lines cannot be compared.
func numberedListPresentation(items []string, format string, startAt int) *Presentation {
	p := New()
	layout := p.GetLayout()
	shape := p.GetActiveSlide().CreateRichTextShape()
	shape.SetOffsetX(0).SetOffsetY(0)
	if layout != nil {
		shape.SetWidth(layout.CX)
	}
	shape.SetHeight(3000000)
	shape.SetWordWrap(false)
	for i, item := range items {
		if i > 0 {
			shape.CreateParagraph()
		}
		shape.CreateTextRun(item)
	}
	for _, para := range shape.GetParagraphs() {
		para.SetBullet(NewBullet().SetNumericBullet(format, startAt))
	}
	return p
}

// bulletPrefix measures the prefix the renderer draws for a bullet. The prefix
// is the string that reaches the line, so asserting on it is asserting on the
// visible numbering; the measuring passes and the drawing pass all go through
// buildBulletRun, so one call covers all three.
func bulletPrefix(t *testing.T, fc *FontCache, b *Bullet, ordinal int) textRun {
	t.Helper()
	para := NewParagraph()
	tr := para.CreateTextRun("item")
	f := NewFont()
	f.Size = 12
	tr.SetFont(f)

	// A renderer carrying the same fallback configuration the real one gets;
	// building a bare literal would silently drop the fallback chains.
	// fontSizePixels multiplies the point size by 12700 (EMU per point) and then
	// by scaleX, so scaleX = 1/12700 puts one point on one pixel and the measured
	// prefix is directly comparable between sizes.
	const ptToPixel = 1.0 / 12700.0
	base := &renderer{
		scaleX:       ptToPixel,
		scaleY:       ptToPixel,
		fontCache:    fc,
		fontFallback: defaultFontFallbackChain,
		cjkFallback:  defaultCJKFallbackChain,
	}
	r := base.subRenderer(image.NewRGBA(image.Rect(0, 0, 4, 4)))
	return r.buildBulletRun(b, para, ordinal)
}

// TestNumberedListIncrementsRenderedNumbers is the headline: the three items of
// a numbered list must not draw the same number. All three paragraphs hold the
// same text, so the only thing that can differ between the three rendered lines
// is the bullet prefix.
func TestNumberedListIncrementsRenderedNumbers(t *testing.T) {
	fc := NewFontCache()
	requireAnyFont(t, fc)

	p := numberedListPresentation([]string{"item", "item", "item"}, NumFormatArabicPeriod, 1)
	img, err := p.SlideToImage(0, goldenOptions(fc))
	if err != nil {
		t.Fatalf("render: %v", err)
	}

	bands := inkRowBands(img)
	if len(bands) != 3 {
		t.Fatalf("found %d text lines in a three-item list; the fixture cannot compare the bullets", len(bands))
	}
	// "Not uniformly equal" rather than pairwise inequality: a font that maps
	// every digit onto the same .notdef box draws three lines that differ from
	// nothing, and one reversed glyph pair would slip past a pairwise check.
	if !bandsDiffer(img, bands[0], bands[1]) || !bandsDiffer(img, bands[1], bands[2]) {
		t.Errorf("all three items of a numbered list rendered the same line; a numbered list has to count 1, 2, 3")
	}
}

// TestNumberedListStartingMidSequenceStillCounts checks the start number is not
// only read but used: a list beginning at 3 has to render 3, 4, 5 — three
// distinct lines — rather than the same number three times.
func TestNumberedListStartingMidSequenceStillCounts(t *testing.T) {
	fc := NewFontCache()
	requireAnyFont(t, fc)

	p := numberedListPresentation([]string{"item", "item", "item"}, NumFormatArabicPeriod, 3)
	img, err := p.SlideToImage(0, goldenOptions(fc))
	if err != nil {
		t.Fatalf("render: %v", err)
	}
	bands := inkRowBands(img)
	if len(bands) != 3 {
		t.Fatalf("found %d text lines in a three-item list; the fixture cannot compare the bullets", len(bands))
	}
	if !bandsDiffer(img, bands[0], bands[1]) || !bandsDiffer(img, bands[1], bands[2]) {
		t.Errorf("a numbered list starting at 3 rendered the same line three times; the start number did not reach the renderer")
	}
}

// TestBulletOrdinalsTrackTheList pins the counting rule that
// buildBulletRun cannot express on its own: consecutive numbered paragraphs
// continue one list, and anything else starts a new one.
func TestBulletOrdinalsTrackTheList(t *testing.T) {
	num := func(format string, startAt int) *Paragraph {
		para := NewParagraph()
		para.CreateTextRun("item")
		para.SetBullet(NewBullet().SetNumericBullet(format, startAt))
		return para
	}
	plain := func() *Paragraph {
		para := NewParagraph()
		para.CreateTextRun("text")
		return para
	}
	ch := func() *Paragraph {
		para := NewParagraph()
		para.CreateTextRun("item")
		para.SetBullet(NewBullet().SetCharBullet("•"))
		return para
	}

	cases := []struct {
		name string
		in   []*Paragraph
		want []int
	}{
		{"a three-item list counts up", []*Paragraph{num(NumFormatArabicPeriod, 1), num(NumFormatArabicPeriod, 1), num(NumFormatArabicPeriod, 1)}, []int{1, 2, 3}},
		{"the start number belongs to the list", []*Paragraph{num(NumFormatArabicPeriod, 3), num(NumFormatArabicPeriod, 1), num(NumFormatArabicPeriod, 1)}, []int{3, 4, 5}},
		{"a zero start number becomes one", []*Paragraph{num(NumFormatArabicPeriod, 0), num(NumFormatArabicPeriod, 0)}, []int{1, 2}},
		{"a roman list starts its own count", []*Paragraph{num(NumFormatRomanUcPeriod, 1), num(NumFormatRomanUcPeriod, 1)}, []int{1, 2}},
		{"a different format starts a new list", []*Paragraph{num(NumFormatArabicPeriod, 1), num(NumFormatRomanUcPeriod, 1)}, []int{1, 1}},
		{"plain text breaks the list", []*Paragraph{num(NumFormatArabicPeriod, 1), plain(), num(NumFormatArabicPeriod, 1)}, []int{1, 0, 1}},
		{"a character bullet breaks the list", []*Paragraph{num(NumFormatArabicPeriod, 1), ch(), num(NumFormatArabicPeriod, 1)}, []int{1, 0, 1}},
		{"no bullet at all", []*Paragraph{plain()}, []int{0}},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			got := bulletOrdinals(tc.in)
			if len(got) != len(tc.want) {
				t.Fatalf("got %d ordinals, want %d", len(got), len(tc.want))
			}
			for i := range tc.want {
				if got[i] != tc.want[i] {
					t.Errorf("ordinal %d = %d, want %d (all: %v)", i, got[i], tc.want[i], got)
					break
				}
			}
		})
	}
}

// TestNumberedBulletPrefixCarriesTheOrdinal shows the number reaching the drawn
// string, which is what the pixel test above can only see indirectly.
func TestNumberedBulletPrefixCarriesTheOrdinal(t *testing.T) {
	fc := NewFontCache()
	requireAnyFont(t, fc)

	b := NewBullet().SetNumericBullet(NumFormatArabicPeriod, 1)
	for ordinal, want := range map[int]string{1: "1. ", 2: "2. ", 10: "10. "} {
		if got := bulletPrefix(t, fc, b, ordinal).text; got != want {
			t.Errorf("prefix for ordinal %d = %q, want %q", ordinal, got, want)
		}
	}
}

// TestRomanBulletPrefixFormatsTheOrdinal covers the other formatters reaching
// the render path; formatBulletNumber was at 0% coverage along with the rest.
func TestRomanBulletPrefixFormatsTheOrdinal(t *testing.T) {
	fc := NewFontCache()
	requireAnyFont(t, fc)

	cases := []struct {
		format  string
		ordinal int
		want    string
	}{
		{NumFormatRomanUcPeriod, 4, "IV. "},
		{NumFormatRomanLcPeriod, 4, "iv. "},
		{NumFormatAlphaUcPeriod, 2, "B. "},
		{NumFormatAlphaLcParen, 3, "c) "},
		{NumFormatArabicParen, 5, "5) "},
	}
	for _, tc := range cases {
		b := NewBullet().SetNumericBullet(tc.format, 1)
		if got := bulletPrefix(t, fc, b, tc.ordinal).text; got != tc.want {
			t.Errorf("prefix for %s at %d = %q, want %q", tc.format, tc.ordinal, got, tc.want)
		}
	}
}

// TestBulletSizeScalesTheDrawnBullet covers <a:buSzPct>: the bullet is a
// percentage of the text size, so a 300% bullet must draw larger than a 100%
// one. The width of the prefix is what the layout uses, so it is the honest
// thing to measure.
func TestBulletSizeScalesTheDrawnBullet(t *testing.T) {
	fc := NewFontCache()
	requireAnyFont(t, fc)

	size := func(pct int) textRun {
		b := NewBullet().SetCharBullet("•")
		b.Size = pct
		return bulletPrefix(t, fc, b, 0)
	}
	base := size(100)
	big := size(300)
	if big.width <= base.width {
		t.Errorf("a 300%% bullet measured %d px, no wider than the 100%% bullet's %d px; Bullet.Size never reached the renderer",
			big.width, base.width)
	}
	if big.font.Size <= base.font.Size {
		t.Errorf("a 300%% bullet was drawn at %dpt, not larger than the 100%% bullet's %dpt", big.font.Size, base.font.Size)
	}
}

// TestEmptyCharBulletDrawsTheDefault matches the writer's substitution: an
// empty Style reaches the file as <a:buChar char="•"/>, so the preview has to
// draw a bullet rather than the bare space it drew before.
func TestEmptyCharBulletDrawsTheDefault(t *testing.T) {
	fc := NewFontCache()
	requireAnyFont(t, fc)

	b := NewBullet().SetCharBullet("")
	got := bulletPrefix(t, fc, b, 0).text
	if got != defaultBulletChar+" " {
		t.Errorf("prefix for an empty character bullet = %q, want %q", got, defaultBulletChar+" ")
	}
}

// inkRowBands splits the inked rows of an image into contiguous runs, one per
// text line. A numbered list has to be segmented into one band per item before
// the bullets can be compared with each other.
func inkRowBands(img image.Image) []image.Rectangle {
	b := img.Bounds()
	var bands []image.Rectangle
	start, end := -1, -1
	for y := b.Min.Y; y < b.Max.Y; y++ {
		inked := false
		for x := b.Min.X; x < b.Max.X; x++ {
			r, g, bl, a := img.At(x, y).RGBA()
			if a>>8 >= 8 && (r>>8 < 250 || g>>8 < 250 || bl>>8 < 250) {
				inked = true
				break
			}
		}
		switch {
		case inked && start < 0:
			start = y
			end = y + 1
		case inked:
			end = y + 1
		case start >= 0:
			bands = append(bands, image.Rect(b.Min.X, start, b.Max.X, end))
			start, end = -1, -1
		}
	}
	if start >= 0 {
		bands = append(bands, image.Rect(b.Min.X, start, b.Max.X, end))
	}
	return bands
}

// bandsDiffer reports whether two bands differ in height or in any pixel.
func bandsDiffer(img image.Image, a, b image.Rectangle) bool {
	if a.Dy() != b.Dy() {
		return true
	}
	for y := 0; y < a.Dy(); y++ {
		for x := a.Min.X; x < a.Max.X; x++ {
			r1, g1, b1, a1 := img.At(x, a.Min.Y+y).RGBA()
			r2, g2, b2, a2 := img.At(x, b.Min.Y+y).RGBA()
			if r1 != r2 || g1 != g2 || b1 != b2 || a1 != a2 {
				return true
			}
		}
	}
	return false
}
