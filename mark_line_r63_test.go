// Round 63: the paragraph-mark line after a trailing <a:br>, and the
// blank-line fallback that looks BACKWARD first.
//
// The slide09 COM variant family of 00022693 pinned the laws:
//
//   - "text br" draws the text line PLUS a mark line of its own — deleting
//     the br deletes that line (vG/vH: -86), and growing the endParaRPr sz
//     moves everything below by exactly the line's growth (vC: 16pt -> 36pt
//     = +53.3px = one 1.2 x 20pt line at 2.222 px/pt).
//   - The mark line's size is the endParaRPr sz when declared, otherwise
//     the previous run's size.
//   - A blank line terminated by an UNSIZED br inherits the run BEFORE it
//     (the text it follows): an unsized br after 32pt text advances a full
//     1.2 x 32pt line, and pinning that same br to sz=1600 shrinks it to
//     the 16pt line (vB: -43). Only a paragraph-LEADING br — nothing before
//     it — falls through to the next run (slide19, unchanged).
//
// Before this round wrapRunLine dropped the mark line entirely and sized an
// unsized blank by a forward scan that found nothing on a trailing br,
// falling back to the 14px stub.

package gopresentation

import (
	"bytes"
	"testing"

	"golang.org/x/image/font/basicfont"
)

// r63Renderer is a bare renderer at the comparison deck's scale: 1600px
// over a 9144000 EMU (10in) wide slide, so 1pt = 2.2222px and 1.2 x 16pt =
// 42.67 -> 43px, 1.2 x 32pt = 85.33 -> 85px.
func r63Renderer() *renderer {
	return &renderer{scaleX: 1600.0 / 9144000.0}
}

func r63TextRun(text string, sizePt int) textRun {
	f := &Font{Name: "Calibri", Size: sizePt}
	return textRun{
		text:        text,
		font:        f,
		face:        basicfont.Face7x13,
		measureFace: basicfont.Face7x13,
		width:       7 * len(text),
	}
}

func r63BreakRun(sizeHundredths int) textRun {
	tr := textRun{text: "\n", face: basicfont.Face7x13, measureFace: basicfont.Face7x13}
	if sizeHundredths > 0 {
		f := NewFont()
		f.Size = sizeHundredths / 100
		tr.font = f
	}
	return tr
}

// TestTrailingBreakLeavesMarkLine pins Law 1/2 at the wrap level: a run
// sequence ending in a break gains a final mark line, sized by the
// endParaRPr when the paragraph declares one and by the previous run
// otherwise. Without a trailing break no line appears — the paragraph mark
// shares the last text line.
func TestTrailingBreakLeavesMarkLine(t *testing.T) {
	r := r63Renderer()

	// Explicit endParaRPr sz=1600: the mark line is the 16pt line.
	lines := r.wrapRunLine([]textRun{r63TextRun("Top", 32), r63BreakRun(0)}, 9999, 1600)
	if len(lines) != 2 {
		t.Fatalf("text+br with endParaRPr 1600: got %d lines, want 2 (text + mark)", len(lines))
	}
	if got := lines[1].lineHeight; got != 43 {
		t.Errorf("mark line height = %d, want 43 (1.2 x 16pt at 1600px/10in)", got)
	}

	// No endParaRPr size: the mark inherits the previous run (32pt -> 85).
	lines = r.wrapRunLine([]textRun{r63TextRun("Top", 32), r63BreakRun(0)}, 9999, 0)
	if len(lines) != 2 {
		t.Fatalf("text+br unsized mark: got %d lines, want 2", len(lines))
	}
	if got := lines[1].lineHeight; got != 85 {
		t.Errorf("unsized mark line height = %d, want 85 (1.2 x 32pt previous run)", got)
	}

	// No trailing break: the mark shares the text line — no extra line.
	lines = r.wrapRunLine([]textRun{r63TextRun("Top", 32)}, 9999, 1600)
	if len(lines) != 1 {
		t.Errorf("text without br: got %d lines, want 1 (no mark line)", len(lines))
	}
}

// TestBlankBreakFallsBackToPreviousRun pins Law 3: the blank line a br
// terminates inherits the run BEFORE it when the br declares no size. The
// forward-only fallback this round replaced left a trailing br's blank at
// the 14px stub (slide09's 129px hole).
func TestBlankBreakFallsBackToPreviousRun(t *testing.T) {
	r := r63Renderer()

	// text(32pt) br(unsized) br(unsized), endParaRPr 1600:
	// [text 85][blank 85 (previous run)][mark 43 (endParaRPr)].
	lines := r.wrapRunLine([]textRun{r63TextRun("Top", 32), r63BreakRun(0), r63BreakRun(0)}, 9999, 1600)
	if len(lines) != 3 {
		t.Fatalf("text+br+br: got %d lines, want 3 (text + blank + mark)", len(lines))
	}
	if got := lines[1].lineHeight; got != 85 {
		t.Errorf("blank after unsized br = %d, want 85 (1.2 x 32pt previous run); the forward-only fallback leaves the 14px stub here", got)
	}
	if got := lines[2].lineHeight; got != 43 {
		t.Errorf("mark line = %d, want 43 (1.2 x 16pt endParaRPr)", got)
	}

	// An EXPLICIT sz on the br still wins over the previous run (slide09 vB).
	lines = r.wrapRunLine([]textRun{r63TextRun("Top", 32), r63BreakRun(1600)}, 9999, 0)
	if len(lines) != 2 {
		t.Fatalf("text+sized br: got %d lines, want 2", len(lines))
	}
	if got := lines[1].lineHeight; got != 43 {
		t.Errorf("blank after br sz=1600 = %d, want 43 (the br's own size wins)", got)
	}

	// A paragraph-LEADING br has nothing before it: the next run speaks
	// (slide19's law, unchanged by the backward fallback).
	lines = r.wrapRunLine([]textRun{r63BreakRun(0), r63TextRun("After", 18)}, 9999, 0)
	if len(lines) != 2 {
		t.Fatalf("leading br: got %d lines, want 2", len(lines))
	}
	if got := lines[0].lineHeight; got != 48 {
		t.Errorf("leading-br blank = %d, want 48 (1.2 x 18pt next run)", got)
	}
}

// TestTrailingBreakMarkLineEndToEnd is the render-level guard through the
// reader: a paragraph ending in <a:br> pushes the paragraph below it down by
// the mark line, and the mark line's size comes from the endParaRPr.
func TestTrailingBreakMarkLineEndToEnd(t *testing.T) {
	fc := NewFontCache()
	if pickInstalledFont(fc) == "" {
		t.Skip("no fonts installed on this machine")
	}
	bodyA := `<a:p><a:r><a:rPr lang="en-US" sz="3200"/><a:t>Top</a:t></a:r><a:br><a:rPr lang="en-US"/></a:br><a:endParaRPr lang="en-US" sz="1600"/></a:p>` +
		`<a:p><a:r><a:rPr lang="en-US" sz="3200"/><a:t>Bottom</a:t></a:r></a:p>`
	bodyB := `<a:p><a:r><a:rPr lang="en-US" sz="3200"/><a:t>Top</a:t></a:r></a:p>` +
		`<a:p><a:r><a:rPr lang="en-US" sz="3200"/><a:t>Bottom</a:t></a:r></a:p>`
	sp := func(x int, body string) string {
		return `<p:sp><p:nvSpPr><p:cNvPr id="2" name="TB"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr>` +
			`<a:xfrm><a:off x="` + itoa(x) + `" y="500000"/><a:ext cx="3000000" cy="3000000"/></a:xfrm>` +
			`<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr wrap="none"/>` +
			`<a:lstStyle/>` + body + `</p:txBody></p:sp>`
	}
	parts := zipParts(t, writeToBytes(t, New()))
	parts["ppt/slides/slide1.xml"] = []byte(bodyPrSlide(sp(500000, bodyA) + sp(5200000, bodyB)))
	pkg := buildZip(t, parts)
	pres, err := (&PPTXReader{}).ReadFromReader(bytes.NewReader(pkg), int64(len(pkg)))
	if err != nil {
		t.Fatalf("read the package: %v", err)
	}

	render := func() [][]int {
		opts := DefaultRenderOptions()
		opts.Width = 1600
		opts.FontCache = fc
		img, err := pres.SlideToImage(0, opts)
		if err != nil {
			t.Fatalf("render: %v", err)
		}
		bands := [][]int{}
		for _, xwin := range [][2]int{{150, 700}, {960, 1450}} {
			var rows []int
			for y := 0; y < img.Bounds().Dy(); y++ {
				cnt := 0
				for x := xwin[0]; x < xwin[1] && x < img.Bounds().Dx(); x++ {
					r, g, b, _ := img.At(x, y).RGBA()
					if int(r>>8) < 110 && int(g>>8) < 110 && int(b>>8) < 110 {
						cnt++
					}
				}
				if cnt >= 2 {
					rows = append(rows, y)
				}
			}
			// First row of the second band = first row after a >=8px gap.
			bands = append(bands, rows)
		}
		return bands
	}
	rowsA := render()[0]
	rowsB := render()[1]
	gapBetween := func(rows []int) int {
		prev := rows[0]
		for _, y := range rows[1:] {
			if y-prev >= 8 {
				return y - rows[0]
			}
			prev = y
		}
		t.Fatalf("ink rows never split into two bands: %v", rows)
		return 0
	}
	got := gapBetween(rowsA) - gapBetween(rowsB)
	// The mark line is 1.2 x 16pt = 42.67 -> 43px at this scale; allow
	// rasterisation slack.
	if got < 36 || got > 50 {
		t.Errorf("trailing-br box pushes Bottom down by %dpx vs the no-br box, want ~43px (the 16pt mark line)", got)
	}
}

func itoa(v int) string {
	if v == 0 {
		return "0"
	}
	var buf [16]byte
	i := len(buf)
	for v > 0 {
		i--
		buf[i] = byte('0' + v%10)
		v /= 10
	}
	return string(buf[i:])
}
