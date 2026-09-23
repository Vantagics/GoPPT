package gopresentation

// A placeholder carries the same <p:style> fallback a regular shape does.
// slide18 of 00022823 draws its "We can do any computation..." banner with a
// content placeholder whose <p:spPr> names no fill at all — everything comes
// from fillRef (the accent gradient), lnRef (the outline) and fontRef (the
// white text). The reader used to drop all of it on the placeholder commit
// path, so the banner rendered as an invisible white-on-white text box.

import (
	"bytes"
	"testing"
)

const placeholderStyleSlide = `
<p:sp>
  <p:nvSpPr><p:cNvPr id="2" name="Banner"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="500000" y="500000"/><a:ext cx="3000000" cy="1000000"/></a:xfrm>
  </p:spPr>
  <p:style>
    <a:lnRef idx="1"><a:schemeClr val="accent2"/></a:lnRef>
    <a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>
    <a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>
  </p:style>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>ANY</a:t></a:r></a:p></p:txBody>
</p:sp>`

func readPlaceholderStyleFixture(t *testing.T) *Presentation {
	t.Helper()
	parts := zipParts(t, writeToBytes(t, New()))
	parts["ppt/slides/slide1.xml"] = []byte(bodyPrSlide(placeholderStyleSlide))
	pkg := buildZip(t, parts)
	pres, err := (&PPTXReader{}).ReadFromReader(bytes.NewReader(pkg), int64(len(pkg)))
	if err != nil {
		t.Fatalf("read the package: %v", err)
	}
	return pres
}

func findPlaceholder(shapes []Shape) *PlaceholderShape {
	for _, sh := range shapes {
		if ph, ok := sh.(*PlaceholderShape); ok {
			return ph
		}
	}
	return nil
}

// TestPlaceholderReadsStyleFallbacks: fillRef becomes the fill, lnRef the
// border, fontRef the default run colour — none of them may be dropped just
// because the shape is a placeholder.
func TestPlaceholderReadsStyleFallbacks(t *testing.T) {
	pres := readPlaceholderStyleFixture(t)
	ph := findPlaceholder(pres.GetAllSlides()[0].GetShapes())
	if ph == nil {
		t.Fatal("placeholder shape dropped")
	}
	if ph.fill == nil || ph.fill.Type != FillSolid {
		t.Fatalf("placeholder has no solid fill from fillRef: %+v", ph.fill)
	}
	if want := (Color{ARGB: "FF4472C4"}); ph.fill.Color != want {
		t.Errorf("fillRef colour = %s, want %s (accent1)", ph.fill.Color.ARGB, want.ARGB)
	}
	if ph.border == nil || ph.border.Style != BorderSolid {
		t.Fatalf("placeholder has no solid border from lnRef: %+v", ph.border)
	}
	if want := (Color{ARGB: "FFED7D31"}); ph.border.Color != want {
		t.Errorf("lnRef colour = %s, want %s (accent2)", ph.border.Color.ARGB, want.ARGB)
	}
	paras := ph.GetParagraphs()
	if len(paras) != 1 {
		t.Fatalf("paragraphs = %d, want 1", len(paras))
	}
	runs := paras[0].GetElements()
	if len(runs) == 0 {
		t.Fatal("no runs in the placeholder paragraph")
	}
	if tr, ok := runs[0].(*TextRun); ok {
		if want := (Color{ARGB: "FFFFFFFF"}); tr.font.Color != want {
			t.Errorf("fontRef run colour = %+v, want FFFFFFFF (lt1, the white banner text)", tr.font.Color)
		}
	} else {
		t.Fatalf("first paragraph element = %T, want *TextRun", runs[0])
	}
}

// TestWriterRoundTripsPlaceholderFill: a styled placeholder must survive a
// read-modify-write with its fill and border — the writer used to emit a bare
// <p:spPr> for placeholders.
func TestWriterRoundTripsPlaceholderFill(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()
	ph := NewPlaceholderShape(PlaceholderBody)
	ph.SetPlaceholderIndex(1)
	ph.SetPosition(500000, 500000)
	ph.SetSize(3000000, 1000000)
	ph.SetFill(NewFill().SetSolid(NewColor("4472C4")))
	ph.SetText("styled banner")
	b := NewBorder()
	b.SetSolidFill(NewColor("ED7D31"))
	b.Width = 1
	ph.SetBorder(b)
	slide.AddShape(ph)

	data := writeToBytes(t, p)
	pres, err := (&PPTXReader{}).ReadFromReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatalf("read back: %v", err)
	}
	got := findPlaceholder(pres.GetAllSlides()[0].GetShapes())
	if got == nil {
		t.Fatal("placeholder dropped on round-trip")
	}
	if got.fill == nil || got.fill.Type != FillSolid || got.fill.Color.ARGB != "FF4472C4" {
		t.Errorf("round-tripped fill = %+v, want solid FF4472C4", got.fill)
	}
	if got.border == nil || got.border.Style != BorderSolid || got.border.Color.ARGB != "FFED7D31" {
		t.Errorf("round-tripped border = %+v, want solid FFED7D31", got.border)
	}
}

// TestPlaceholderStyleFillRenders: the themed fill must reach pixels — an
// accent-blue banner with white text inside it, not a white-on-white text box.
func TestPlaceholderStyleFillRenders(t *testing.T) {
	pres := readPlaceholderStyleFixture(t)
	fc := NewFontCache()
	requireAnyFont(t, fc)
	opts := DefaultRenderOptions()
	opts.Width = 720
	opts.FontCache = fc
	img, err := pres.SlideToImage(0, opts)
	if err != nil {
		t.Fatalf("render: %v", err)
	}
	// Box: 500000..3500000 EMU x, 500000..1500000 EMU y → 40..280 px, 40..120 px.
	blue, white := 0, 0
	for y := 45; y < 115; y++ {
		for x := 45; x < 275; x++ {
			r, g, b, _ := img.At(x, y).RGBA()
			r8, g8, b8 := r>>8, g>>8, b>>8
			if b8 > 150 && b8 > r8+40 && g8 > r8 && g8 < b8 {
				blue++
			}
			if r8 > 235 && g8 > 235 && b8 > 235 {
				white++
			}
		}
	}
	if blue < 500 {
		t.Errorf("accent fill pixels = %d, want the banner box painted blue", blue)
	}
	if white < 20 {
		t.Errorf("white text pixels inside the box = %d, want the fontRef-white text visible on the fill", white)
	}
}
