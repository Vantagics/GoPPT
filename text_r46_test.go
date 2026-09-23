package gopresentation

import (
	"bytes"
	"strings"
	"testing"
)

// A runless paragraph's line box rides its <a:endParaRPr> (round 26), but
// PowerPoint-written empty paragraphs usually declare no sz there — slide36's
// blank line between two groups advances a full 1.2 × the level default in
// the COM export while the 14px fallback left it 71px short. The size comes
// from the same ladder the sibling runs resolve through.
func TestEmptyParagraphInheritsLevelDefaultSize(t *testing.T) {
	slide := strings.Replace(masterLevelSlideShape,
		"<a:p><a:pPr lvl=\"1\"/><a:r><a:rPr lang=\"en-US\"/><a:t>LEVEL1</a:t></a:r></a:p>",
		"<a:p><a:pPr lvl=\"1\"/><a:r><a:rPr lang=\"en-US\"/><a:t>LEVEL1</a:t></a:r></a:p>"+
			"<a:p><a:pPr marL=\"0\" indent=\"0\"><a:buNone/></a:pPr><a:endParaRPr lang=\"en-US\"/></a:p>",
		1)
	parts := zipParts(t, writeToBytes(t, New()))
	parts["ppt/slides/slide1.xml"] = []byte(bodyPrSlide(slide))
	parts["ppt/slideMasters/slideMaster1.xml"] = []byte(masterLevelMaster)
	pkg := buildZip(t, parts)
	pres, err := (&PPTXReader{}).ReadFromReader(bytes.NewReader(pkg), int64(len(pkg)))
	if err != nil {
		t.Fatalf("read the package: %v", err)
	}
	ph, ok := firstShapeOfType[*PlaceholderShape](pres)
	if !ok {
		t.Fatal("no body placeholder found")
	}
	paras := ph.GetParagraphs()
	if len(paras) != 3 {
		t.Fatalf("got %d paragraphs, want 3", len(paras))
	}
	// A paragraph with runs keeps endParaRPrSize 0 — the writer must not
	// grow its XML with an invented size.
	if got := paras[0].endParaRPrSize; got != 0 {
		t.Errorf("paragraph with runs endParaRPrSize = %d, want 0", got)
	}
	// The empty paragraph rides the level default (lvl1pPr 44pt = 4400).
	if got := paras[2].endParaRPrSize; got != 4400 {
		t.Errorf("empty paragraph endParaRPrSize = %d, want 4400 (level default)", got)
	}
}

// The u46 COM variant deck pinned the underline geometry: offset below the
// baseline = floor(0.1em) and thickness = round(em/16 + 0.1) — 18pt draws a
// 3px band, 32pt a 5px band. The old constant "+2, 1px" drew a hairline that
// all but vanished at report scale.
func TestUnderlineBandMatchesPptGeometry(t *testing.T) {
	fc := NewFontCache()
	if pickInstalledFont(fc) == "" {
		t.Skip("no fonts installed on this machine")
	}
	p := New()
	slide := p.GetActiveSlide()
	for i, sz := range []int{18, 32} {
		tb := slide.CreateRichTextShape()
		tb.BaseShape.SetOffsetX(914400).SetOffsetY(int64(300000 + i*2600000))
		tb.BaseShape.SetSize(7315200, 1200000)
		run := tb.GetParagraphs()[0].CreateTextRun("Dataset check")
		run.GetFont().SetSize(sz).SetUnderline(UnderlineSingle)
	}
	opts := DefaultRenderOptions()
	opts.Width = 1600
	opts.FontCache = fc
	img, err := p.SlideToImage(0, opts)
	if err != nil {
		t.Fatalf("render: %v", err)
	}
	lum := func(x, y int) int {
		r, g, b, _ := img.At(x, y).RGBA()
		return (int(r>>8) + int(g>>8) + int(b>>8)) / 3
	}
	// Shapes sit at offsetY 300000 + i*2600000 EMU (52px + i*455px at this
	// scale), each holding one unwrapped line.
	for i, sz := range []int{18, 32} {
		wantThick := map[int]int{18: 3, 32: 5}[sz]
		y0 := 52 + i*455
		// Row profile across the text's x-extent: the glyph band(s), a gap,
		// then the underline band.
		ink := func(y int) bool {
			for x := 150; x < 700; x++ {
				if lum(x, y) < 128 {
					return true
				}
			}
			return false
		}
		var bands [][2]int
		inb := false
		for y := y0; y < y0+220; y++ {
			if ink(y) && !inb {
				bands = append(bands, [2]int{y, y})
				inb = true
			} else if ink(y) {
				bands[len(bands)-1][1] = y
			} else {
				inb = false
			}
		}
		if len(bands) < 2 {
			t.Fatalf("%dpt: bands=%v, want glyphs + underline", sz, bands)
		}
		ul := bands[len(bands)-1]
		glyphs := bands[len(bands)-2]
		if got := ul[1] - ul[0] + 1; got != wantThick {
			t.Errorf("%dpt underline thickness = %dpx, want %d (COM variant law); bands=%v", sz, got, wantThick, bands)
		}
		if gap := ul[0] - glyphs[1]; gap < 2 || gap > 9 {
			t.Errorf("%dpt underline gap below glyphs = %dpx, want the floor(0.1em) band (3-7px at these sizes); bands=%v", sz, gap, bands)
		}
	}
}
