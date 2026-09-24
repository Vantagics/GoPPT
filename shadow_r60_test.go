package gopresentation

import (
	"image"
	"testing"
)

// Round 60, two laws pinned by old-deck slides 30/34:
//
//  1. A no-fill roundRect casts the shadow of its PEN TRACE (the stroke
//     ring), never of the area it encloses — slide30's 6pt accent6
//     annotation rings sit over a table that stays visible through them,
//     while the COM export still shows their outerShdw around the stroke.
//     The renderer used to blur the whole rounded-rect silhouette, painting
//     the enclosed cells dark.
//
//  2. The bullet is followed by a TAB whose landing is marL whenever the
//     pen after the bullet GLYPH alone has not passed it — the trailing
//     space of the bullet text must not count. Slide34's en-dash: glyph pen
//     10px short of marL, glyph+space 3px past, and the COM export puts the
//     text exactly on marL.

func ringShadowShape() string {
	return `<p:sp>
  <p:nvSpPr><p:cNvPr id="2" name="Ring"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="500000" y="500000"/><a:ext cx="3000000" cy="2000000"/></a:xfrm>
    <a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>
    <a:noFill/>
    <a:ln w="76200"><a:solidFill><a:schemeClr val="accent6"/></a:solidFill></a:ln>
    <a:effectLst><a:outerShdw blurRad="50800" dist="38100" dir="2700000" algn="tl" rotWithShape="0"><a:prstClr val="black"><a:alpha val="40000"/></a:prstClr></a:outerShdw></a:effectLst>
  </p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody>
</p:sp>`
}

func renderOneShape(t *testing.T, shape string) image.Image {
	t.Helper()
	if pickInstalledFont(NewFontCache()) == "" {
		t.Skip("no fonts installed on this machine")
	}
	theme := themeTestTheme(themeOfficeFillStyles, themeOfficeLnStyles)
	rels := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`
	pres := themeTestRead(t, theme, shape, rels)
	opts := DefaultRenderOptions()
	opts.Width = 1600
	opts.FontCache = NewFontCache()
	img, err := pres.SlideToImage(0, opts)
	if err != nil {
		t.Fatalf("render: %v", err)
	}
	return img
}

// The enclosed area of a no-fill shadowed roundRect must stay background:
// the silhouette PowerPoint blurs is the stroke ring, so the interior —
// hundreds of pixels from any ink — sees no shadow at all.
func TestNoFillRoundRectShadowSparesInterior(t *testing.T) {
	img := renderOneShape(t, ringShadowShape())
	lum := func(x, y int) int {
		r, g, b, _ := img.At(x, y).RGBA()
		return int((r>>8 + g>>8 + b>>8) / 3)
	}
	// box: x=87.5 y=87.5 w=525 h=350 px at 1600 — centre (350, 262).
	minLum := 255
	for y := 250; y < 275; y++ {
		for x := 338; x < 363; x++ {
			if l := lum(x, y); l < minLum {
				minLum = l
			}
		}
	}
	if minLum < 245 {
		t.Errorf("ring interior darkened to %d, want ~255 (no-fill shapes cast no area shadow)", minLum)
	}
	// The ring's own shadow must still be there: just below the bottom
	// stroke, offset along dir=45°, the blur darkens the white.
	darkest := 255
	for y := 444; y < 459; y++ {
		for x := 338; x < 363; x++ {
			if l := lum(x, y); l < darkest {
				darkest = l
			}
		}
	}
	if darkest > 220 {
		t.Errorf("no ring shadow below the stroke (darkest %d), want blurred ink from the outerShdw", darkest)
	}
}

func bulletTabShape() string {
	return `<p:sp>
  <p:nvSpPr><p:cNvPr id="2" name="L"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="500000" y="500000"/><a:ext cx="6000000" cy="1200000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/>
    <a:p><a:pPr marL="688975" indent="-288925"><a:buFont typeface="Calibri"/><a:buChar char="&#8211;"/></a:pPr><a:r><a:rPr lang="en-US" sz="3200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:rPr><a:t>Integration</a:t></a:r></a:p>
  </p:txBody>
</p:sp>`
}

// The text after a hanging-indent bullet lands on marL: advance from the
// dash ink to the first text glyph equals the hang (288925 EMU = 50.6px at
// 1600) plus the two glyphs' left side bearings — NOT the dash's natural
// advance plus a space, which overshoots marL by most of a space width.
func TestBulletTabLandsTextOnMarL(t *testing.T) {
	img := renderOneShape(t, bulletTabShape())
	lum := func(x, y int) int {
		r, g, b, _ := img.At(x, y).RGBA()
		return int((r>>8 + g>>8 + b>>8) / 3)
	}
	// find the text line band inside the box (y 87..350)
	bandTop, bandBot := -1, -1
	for y := 100; y < 340; y++ {
		ink := 0
		for x := 100; x < 1200; x++ {
			if lum(x, y) < 110 {
				ink++
			}
		}
		if ink > 8 {
			if bandTop < 0 {
				bandTop = y
			}
			bandBot = y
		}
	}
	if bandTop < 0 {
		t.Fatal("no text band found")
	}
	// glyph runs across the band
	type run struct{ x0, x1 int }
	var runs []run
	inr := false
	for x := 100; x < 900; x++ {
		ink := false
		for y := bandTop; y <= bandBot; y++ {
			if lum(x, y) < 110 {
				ink = true
				break
			}
		}
		if ink && !inr {
			runs = append(runs, run{x, x})
			inr = true
		} else if ink {
			runs[len(runs)-1].x1 = x
		} else {
			inr = false
		}
	}
	if len(runs) < 2 {
		t.Fatalf("want dash + text glyph runs, got %v", runs)
	}
	dashStart := runs[0].x0
	// the text starts at the first run after a gap of >= 8px following the dash
	textStart := -1
	for i := 1; i < len(runs); i++ {
		if runs[i].x0-runs[i-1].x1 >= 8 {
			textStart = runs[i].x0
			break
		}
	}
	if textStart < 0 {
		t.Fatalf("no gap after the dash found; runs=%v", runs[:minInt(len(runs), 6)])
	}
	hangPx := float64(288925) * 1600 / 9144000 // 50.56px
	adv := float64(textStart - dashStart)
	if adv < hangPx-4 || adv > hangPx+10 {
		t.Errorf("dash→text advance %.1fpx, want the hang %.1fpx ± bearings (tab lands on marL)", adv, hangPx)
	}
}
