package gopresentation

import (
	"image"
	"image/draw"
	"strings"
	"testing"

	"golang.org/x/image/font"
	"golang.org/x/image/font/basicfont"
	"golang.org/x/image/math/fixed"
)

// widerFace wraps a face and reports every glyph 1px wider, the way a
// HintingFull face rounds each advance up to whole pixels. The wrap code must
// not consult it: the wrap width rides the measure face alone.
type widerFace struct {
	font.Face
}

func (f widerFace) GlyphAdvance(r rune) (fixed.Int26_6, bool) {
	a, ok := f.Face.GlyphAdvance(r)
	return a + 64, ok
}

func (f widerFace) Glyph(dot fixed.Point26_6, r rune) (dr image.Rectangle, mask image.Image, maskp image.Point, advance fixed.Int26_6, ok bool) {
	dr, mask, maskp, advance, ok = f.Face.Glyph(dot, r)
	return dr, mask, maskp, advance + 64, ok
}

var _ font.Face = widerFace{}

// wrapLineTexts flattens wrapped lines to their text.
func wrapLineTexts(lines []textLine) []string {
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

// wrapRunsForIndent wraps with the hanging-indent model the bulleted SmartArt
// paragraphs take, using a deterministic built-in face.
func wrapRunsForIndent(runs []textRun, firstW, contW int) []string {
	r := &renderer{}
	return wrapLineTexts(r.wrapRunLineWithIndent(runs, firstW, contW, 0))
}

// TestWrapFitsLineExactlyAtMaxWidth pins that a line whose advance equals the
// wrap limit stays on one line. PowerPoint compares the advance against the
// available width directly; the old 1px reservation wrapped customGeo slide47's
// "model curricula" — exactly 240px on a 240px continuation line — into a
// phantom third line the COM export does not have.
func TestWrapFitsLineExactlyAtMaxWidth(t *testing.T) {
	// basicfont.Face7x13 advances 7px per ASCII glyph, so "aaa bbb" is 49px.
	const text = "aaa bbb"
	if got := faceWidth(basicfont.Face7x13, text); got != 49 {
		t.Skipf("basicfont metrics changed (width=%d, want 49); adjust the fixture", got)
	}
	runs := []textRun{asciiRun(text)}

	lines := wrapRunsForTest(runs, 49)
	if len(lines) != 1 {
		t.Fatalf("lines = %q, want a 49px line to fit a 49px wrap limit", lines)
	}
}

// TestWrapWithIndentFitsLineExactlyAtMaxWidth pins the same no-reservation rule
// on the hanging-indent path the bulleted body columns take.
func TestWrapWithIndentFitsLineExactlyAtMaxWidth(t *testing.T) {
	const text = "aaa bbb"
	if got := faceWidth(basicfont.Face7x13, text); got != 49 {
		t.Skipf("basicfont metrics changed (width=%d, want 49); adjust the fixture", got)
	}
	runs := []textRun{asciiRun(text)}

	lines := wrapRunsForIndent(runs, 49, 49)
	if len(lines) != 1 {
		t.Fatalf("lines = %q, want a 49px line to fit a 49px continuation width", lines)
	}
}

// TestWrapWidthRidesMeasureFace pins that the wrap decision measures with the
// measure face only. The render face's per-glyph hinted advances read wider
// (each glyph rounds up a pixel here); taking the larger of the two faces —
// the pre-precision heuristic — over-wrapped boundary lines PowerPoint keeps
// on one, which is exactly how slide47's "model curricula" gained a phantom
// third line against the COM export.
func TestWrapWidthRidesMeasureFace(t *testing.T) {
	const text = "aaa bbb"
	if got := faceWidth(basicfont.Face7x13, text); got != 49 {
		t.Skipf("basicfont metrics changed (width=%d, want 49); adjust the fixture", got)
	}
	// Measure face: 49px total. Render face: 56px (each glyph +1px).
	run := textRun{
		text:        text,
		font:        NewFont(),
		face:        widerFace{basicfont.Face7x13},
		measureFace: basicfont.Face7x13,
		width:       faceWidth(basicfont.Face7x13, text),
	}
	runs := []textRun{run}

	// The measure advance fits 49px exactly; the hinted render advance (56px)
	// does not. The line must fit.
	lines := wrapRunsForIndent(runs, 49, 49)
	if len(lines) != 1 {
		t.Fatalf("lines = %q, want the wrap to ride the 49px measure face, not the 56px render face", lines)
	}
}

var _ = draw.Over
