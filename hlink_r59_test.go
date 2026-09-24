package gopresentation

import (
	"image"
	"strings"
	"testing"
)

// Round 59: a linked run is always underlined. PowerPoint takes the underline
// from the link, not from the run's u attribute — deck 00022823's slide 7
// declares u="none" on the citation title's hlinkClick runs and the COM
// export underlines them anyway (the UI greys the underline control out for
// linked text). The renderer forces UnderlineSingle at draw time on a copy
// of the font, so the model — and with it the writer's u attribute — stays
// exactly as the file declared it.

func hlinkUnderlineShape(attrs, extra string) string {
	return `<p:sp>
  <p:nvSpPr><p:cNvPr id="2" name="Linked"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="500000" y="500000"/><a:ext cx="6000000" cy="800000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
  <p:txBody>
    <a:bodyPr/><a:lstStyle/>
    <a:p><a:r><a:rPr lang="en-US" sz="1400" ` + attrs + `><a:solidFill><a:schemeClr val="tx1"/></a:solidFill>` + extra + `</a:rPr><a:t>linked text runs here</a:t></a:r></a:p>
  </p:txBody>
</p:sp>`
}

// hlinkUnderlineRender reads a one-shape slide whose run carries rpr (u attr
// and hlinkClick) and renders it at 1600px, returning the image.
func hlinkUnderlineRender(t *testing.T, attrs, extra string) image.Image {
	t.Helper()
	if pickInstalledFont(NewFontCache()) == "" {
		t.Skip("no fonts installed on this machine")
	}
	theme := themeTestTheme(themeOfficeFillStyles, themeOfficeLnStyles)
	rels := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="http://example.com/" TargetMode="External"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`
	pres := themeTestRead(t, theme, hlinkUnderlineShape(attrs, extra), rels)
	opts := DefaultRenderOptions()
	opts.Width = 1600
	opts.FontCache = NewFontCache()
	img, err := pres.SlideToImage(0, opts)
	if err != nil {
		t.Fatalf("render: %v", err)
	}
	return img
}

// blueBands returns the contiguous blue-ink row bands (the run renders in the
// theme's hlink colour, round 30) within the shape's box.
func blueBands(t *testing.T, img image.Image) [][2]int {
	t.Helper()
	lum := func(x, y int) (int, int, int) {
		r, _, b, _ := img.At(x, y).RGBA()
		return int(r >> 8), 0, int(b >> 8)
	}
	var bands [][2]int
	inb := false
	for y := 40; y < 220; y++ {
		ink := false
		for x := 60; x < 1300; x++ {
			r, _, b := lum(x, y)
			if b > r+40 && b > 120 && r < 150 {
				ink = true
				break
			}
		}
		if ink && !inb {
			bands = append(bands, [2]int{y, y})
			inb = true
		} else if ink {
			bands[len(bands)-1][1] = y
		} else {
			inb = false
		}
	}
	return bands
}

// The linked run with u="none" draws a glyph band and, below it, the forced
// underline band. Without the link the same declaration draws glyphs only.
func TestHyperlinkForcesUnderlineDespiteNone(t *testing.T) {
	linked := blueBands(t, hlinkUnderlineRender(t, `u="none" i="1"`, `<a:hlinkClick r:id="rId1"/>`))
	if len(linked) < 2 {
		t.Fatalf("linked run bands=%v, want glyphs + forced underline", linked)
	}
	glyphs := linked[len(linked)-2]
	ul := linked[len(linked)-1]
	if gap := ul[0] - glyphs[1]; gap < 1 || gap > 8 {
		t.Errorf("underline gap below glyphs = %dpx, want the 0.1em band (3-4px at 14pt); bands=%v", gap, linked)
	}
	if got := ul[1] - ul[0] + 1; got < 1 || got > 4 {
		t.Errorf("underline thickness = %dpx, want 2-3px at 14pt; bands=%v", got, linked)
	}

	plain := blueBands(t, hlinkUnderlineRender(t, `u="none" i="1"`, ``))
	// No hlink: the run paints in its declared tx1 (black) — no blue ink at
	// all, which is the negative control in itself.
	if len(plain) != 0 {
		t.Errorf("unlinked run has blue bands %v, want none (its colour is tx1, not the hlink)", plain)
	}
}

// The forced underline must live in the renderer only: the model font keeps
// UnderlineNone and the writer re-emits the declared u="none" — a render-
// then-save cycle must not rewrite the file's formatting.
func TestHyperlinkUnderlineForcingLeavesModelAndWriterUntouched(t *testing.T) {
	theme := themeTestTheme(themeOfficeFillStyles, themeOfficeLnStyles)
	shape := hlinkUnderlineShape(`u="none" i="1"`, `<a:hlinkClick r:id="rId1"/>`)
	rels := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="http://example.com/" TargetMode="External"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`
	pres := themeTestRead(t, theme, shape, rels)
	rt := pres.GetAllSlides()[0].GetShapes()[0].(*RichTextShape)
	tr := rt.GetParagraphs()[0].GetElements()[0].(*TextRun)
	if tr.font == nil || tr.font.Underline != UnderlineNone {
		t.Fatalf("model underline = %v, want none as declared", tr.font)
	}
	// Render — the forcing happens on a copy inside the draw path.
	opts := DefaultRenderOptions()
	opts.Width = 1600
	opts.FontCache = NewFontCache()
	if _, err := pres.SlideToImage(0, opts); err != nil {
		t.Fatalf("render: %v", err)
	}
	if tr.font.Underline != UnderlineNone {
		t.Fatalf("after render the model underline = %q, want untouched none", tr.font.Underline)
	}
	parts := zipParts(t, writeToBytes(t, pres))
	slide1 := string(parts["ppt/slides/slide1.xml"])
	// The writer normalises u="none" away (it is the default), so the bite
	// is the opposite: rendering must not have upgraded the run to an
	// explicit underline in the saved file.
	if strings.Contains(slide1, `u="sng"`) {
		t.Errorf("rendering upgraded the linked run's underline in the saved file: %s", slide1)
	}
}
