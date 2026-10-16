package gopresentation

import (
	"image"
	"testing"
)

// The r69 round pinned, with an 8-slide COM-exported ladder (out_deck/
// r69probe.pptx: same box as customGeo slide47's title — off 685800,425450,
// ext 7772400,987425, anchor=ctr — varying the normAutofit levers), that a
// CENTRED autofit block's first baseline rides lnSpcReduction exactly like
// every other box quantity (pitch, block height, top-anchor ascent):
//
//	single line "Tasks for Districts" 44pt, COM baselines:
//	  F no levers        197
//	  H fontScale 77.5%  188   (fontScale: correct before this round)
//	  G lnRed 20%        186   ≈ H  — reduction does NOT push the line down
//	  D both             179
//	two lines, fontScale+lnRed: line1 144, line2 217 (pitch 73 = 0.96 em)
//
// The pre-fix renderer reduced the block height and the pitches but drew the
// first baseline from the UNREDUCED ascent, dropping every centred autofit
// block by reduction/2 (slide47's title +18px, probe G +20). These tests pin
// the deltas against the no-lever twin, ±6 for the residual font-face noise
// the probe already shows on lnRed=0 rows (F −3, B −2).
func inkRows(t *testing.T, img image.Image, rect image.Rectangle) [][2]int {
	t.Helper()
	var rows [][2]int
	inLine, start := false, 0
	for y := rect.Min.Y; y < rect.Max.Y; y++ {
		cnt := 0
		for x := rect.Min.X; x < rect.Max.X; x += 2 {
			r, g, b, a := img.At(x, y).RGBA()
			if a != 0 && (r+g+b)/3 < 30000 {
				cnt++
			}
		}
		on := cnt > 3
		if on && !inLine {
			inLine, start = true, y
		} else if !on && inLine {
			inLine = false
			rows = append(rows, [2]int{start, y - 1})
		}
	}
	if inLine {
		rows = append(rows, [2]int{start, rect.Max.Y - 1})
	}
	if len(rows) == 0 {
		t.Fatal("no ink rendered in the box")
	}
	return rows
}

// renderCenteredLadder builds the ladder textbox (title-box geometry,
// anchor=ctr, runs inherit the theme font at 44pt — same as the spliced probe
// XML) with the given levers and renders it at 1600px.
func renderCenteredLadder(t *testing.T, fontScale, lnSpcReduction int, lines ...string) image.Image {
	t.Helper()
	pres := New()
	shape := pres.GetActiveSlide().CreateRichTextShape()
	shape.SetOffsetX(685800).SetOffsetY(425450)
	shape.SetWidth(7772400).SetHeight(987425)
	shape.SetTextAnchor(TextAnchorMiddle)
	shape.autoFit = AutoFitNormal
	shape.fontScale = fontScale
	shape.lnSpcReduction = lnSpcReduction
	for i, text := range lines {
		var para *Paragraph
		if i == 0 {
			para = shape.GetParagraphs()[0]
		} else {
			para = shape.CreateParagraph()
		}
		run := para.CreateTextRun(text)
		run.GetFont().SetSize(44)
		run.GetFont().Color = NewColor("000000")
	}
	opts := DefaultRenderOptions()
	opts.Width = 1600
	img, err := pres.SlideToImage(0, opts)
	if err != nil {
		t.Fatalf("render: %v", err)
	}
	return img
}

// TestCenteredAutofitLnRedDoesNotSinkTheBlock is the biting test: with
// lnSpcReduction the centred block must move UP toward the fontScale-only
// position (COM G≈H), not down. Pre-fix the single-line baseline sat at 195
// (both levers) and 206 (reduction only) against COM's 179/186.
func TestCenteredAutofitLnRedDoesNotSinkTheBlock(t *testing.T) {
	baselineOf := func(img image.Image) int {
		rect := image.Rect(120, 40, 1480, 330)
		rows := inkRows(t, img, rect)
		if len(rows) != 1 {
			t.Fatalf("expected a single ink row, got %v", rows)
		}
		return rows[0][1] // "Tasks for Districts" has no descenders
	}
	base := baselineOf(renderCenteredLadder(t, 0, 0, "Tasks for Districts"))
	lnRed := baselineOf(renderCenteredLadder(t, 0, 20000, "Tasks for Districts"))
	both := baselineOf(renderCenteredLadder(t, 77500, 20000, "Tasks for Districts"))
	scale := baselineOf(renderCenteredLadder(t, 77500, 0, "Tasks for Districts"))

	// COM deltas: G−F = −11, D−F = −18, H−F = −9. The face-metric noise the
	// probe shows (±3) applies to every row, so ±6 windows on the deltas.
	if got := lnRed - base; got < -17 || got > -5 {
		t.Errorf("lnRed-only baseline delta = %+d (base %d, lnRed %d), want in [-17,-5] per COM (-11); the unreduced first-ascent sink is back", got, base, lnRed)
	}
	if got := both - base; got < -24 || got > -12 {
		t.Errorf("both-levers baseline delta = %+d (base %d, both %d), want in [-24,-12] per COM (-18)", got, base, both)
	}
	if got := scale - base; got < -15 || got > -3 {
		t.Errorf("fontScale-only baseline delta = %+d (base %d, scale %d), want in [-15,-3] per COM (-9)", got, base, scale)
	}
}

// The two-line block keeps the reduced pitch AND lands its first line at the
// fontScale-only position: COM line1 ink 88..143 / line2 162..216 (pitch 73).
// Pre-fix the whole block sat +18 low (108.. instead of 88..).
func TestCenteredAutofitTwoLineBlockTopRidesReduction(t *testing.T) {
	img := renderCenteredLadder(t, 77500, 20000,
		"Preparation for New Standards", "Tasks for Districts")
	rows := inkRows(t, img, image.Rect(120, 40, 1480, 330))
	if len(rows) != 2 {
		t.Fatalf("expected two ink rows, got %v", rows)
	}
	if rows[0][0] > 100 {
		t.Errorf("line1 ink top = %d, want ≤100 (COM 88, post-fix ≈94); the block sank by the unreduced first ascent", rows[0][0])
	}
	pitch := rows[1][0] - rows[0][0]
	if pitch < 68 || pitch > 78 {
		t.Errorf("line pitch = %d, want ≈73 (0.96 em at 34.1pt effective)", pitch)
	}
	if rows[1][1] > 224 {
		t.Errorf("line2 ink bottom = %d, want ≤224 (COM 216, post-fix ≈218)", rows[1][1])
	}
}
