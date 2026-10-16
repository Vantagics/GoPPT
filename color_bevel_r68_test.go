package gopresentation

// r68 fixes — pinned against the r68 COM probe deck (out_deck/r68_probe.pptx,
// exports out_deck/r68_com) and PowerPoint's own colour-transform maths.
//
// Two real defects were closed:
//
//  1. The OOXML colour modifiers <a:hueOff>, <a:hueMod>, <a:satOff>,
//     <a:alphaOff>, <a:alphaMod> were silently ignored. slide47's SmartArt
//     columns each carry a hueOff (and satOff/lumOff) that rotates the shared
//     accent5 fill blue→green→peach; without it every column rendered the same
//     blue. The probe (rect E, schemeClr accent5 + tint + alpha) matched
//     exactly, so only the hue/sat/alpha rotations were missing.
//  2. The bevel flat-face gain (bevelFaceGain = 1.19, calibrated for the
//     matte/metal default) was applied to EVERY material. A "dkEdge" face is lit
//     flat (COM centre == the un-bevelled fill), and a "translucentPowder" face
//     only mildly (~1.10). The probe pinned both: dkEdge centre 212,228,246 vs
//     the un-bevelled 212,229,236, translucentPowder 222,236,250 vs 229,247,255
//     (the old 1.19 reading).

import (
	"image/color"
	"testing"
)

// --- Unit tests: the modifier helpers hit exact HSL/alpha values ------------

func TestApplyHueOffExact(t *testing.T) {
	// Pure red, hue 0. +120 -> green, -120 -> blue.
	c := NewColor("FFFF0000")
	applyHueOff(&c, 120)
	if c.ARGB != "FF00FF00" {
		t.Errorf("hueOff +120 on red = %s, want FF00FF00", c.ARGB)
	}
	c = NewColor("FFFF0000")
	applyHueOff(&c, -120)
	if c.ARGB != "FF0000FF" {
		t.Errorf("hueOff -120 on red = %s, want FF0000FF", c.ARGB)
	}
}

func TestApplyHueModExact(t *testing.T) {
	// Green, hue 120. x2 -> 240 = blue; x0.5 -> 60 = yellow.
	c := NewColor("FF00FF00")
	applyHueMod(&c, 2.0)
	if c.ARGB != "FF0000FF" {
		t.Errorf("hueMod 2.0 on green = %s, want FF0000FF", c.ARGB)
	}
	c = NewColor("FF00FF00")
	applyHueMod(&c, 0.5)
	if c.ARGB != "FFFFFF00" {
		t.Errorf("hueMod 0.5 on green = %s, want FFFF00", c.ARGB)
	}
}

func TestApplySatOffExact(t *testing.T) {
	// Red, saturation 1. Drop to 0 -> grey 808080.
	c := NewColor("FFFF0000")
	applySatOff(&c, -1.0)
	if c.ARGB != "FF808080" {
		t.Errorf("satOff -1.0 on red = %s, want FF808080", c.ARGB)
	}
}

func TestApplyAlphaOffExact(t *testing.T) {
	c := NewColor("FFFF0000")
	applyAlphaOff(&c, -0.5)
	if c.ARGB != "80FF0000" {
		t.Errorf("alphaOff -0.5 on opaque red = %s, want 80FF0000", c.ARGB)
	}
}

func TestApplyAlphaModExact(t *testing.T) {
	c := NewColor("FFFF0000")
	applyAlphaMod(&c, 0.5)
	if c.ARGB != "80FF0000" {
		t.Errorf("alphaMod 0.5 on opaque red = %s, want 80FF0000", c.ARGB)
	}
}

// --- Render tests: the modifiers and material gate actually fire ------------

// bevelRect builds a 3in solid-fill square; `extra` is appended inside <p:spPr>
// (use it for scene3d/sp3d). The fill is accent5-tinted, matching the probe.
func bevelRect(fillInner, extra string) string {
	return `<p:sp>
  <p:nvSpPr><p:cNvPr id="2" name="R"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="200000" y="200000"/><a:ext cx="3000000" cy="3000000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="4BACC6">` + fillInner + `</a:srgbClr></a:solidFill>` + extra + `
  </p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>
</p:sp>`
}

func renderCenter(t *testing.T, shapes string) color.RGBA {
	t.Helper()
	pres := readGradFixture(t, shapes)
	fc := NewFontCache()
	requireAnyFont(t, fc)
	opts := DefaultRenderOptions()
	opts.Width = 960
	opts.FontCache = fc
	img, err := pres.SlideToImage(0, opts)
	if err != nil {
		t.Fatalf("render: %v", err)
	}
	r, g, b, a := img.At(178, 178).RGBA()
	return color.RGBA{uint8(r >> 8), uint8(g >> 8), uint8(b >> 8), uint8(a >> 8)}
}

func TestHueOffRotatesFillInPipeline(t *testing.T) {
	// No hue rotation: accent5 tinted — blue-dominant (B is the largest channel).
	base := renderCenter(t, bevelRect(`<a:tint val="40000"/><a:alpha val="90000"/>`, ""))
	if !(base.B >= base.R && base.B >= base.G) {
		t.Fatalf("base fill not blue-dominant: %+v", base)
	}
	// Tint+alpha regression anchor (matches the probe rect A / COM 213,230,236).
	if absDiff(int(base.R), 212) > 6 || absDiff(int(base.G), 229) > 6 || absDiff(int(base.B), 236) > 6 {
		t.Errorf("base tint+alpha pixel = %+v, want ~{212,229,236}", base)
	}
	// With a hueOff of -10740482/60000 ≈ -179°, the fill rotates to peach:
	// green and red overtake blue. Before the fix hueOff was ignored and this
	// stayed blue-dominant, so the assertion below bit.
	rotated := renderCenter(t, bevelRect(`<a:tint val="40000"/><a:alpha val="90000"/><a:hueOff val="-10740482"/>`, ""))
	if rotated.B >= rotated.R && rotated.B >= rotated.G {
		t.Errorf("hueOff fill still blue-dominant: %+v (hueOff was ignored)", rotated)
	}
	if absDiff(int(rotated.R), int(base.R)) < 8 && absDiff(int(rotated.B), int(base.B)) < 8 {
		t.Errorf("hueOff produced no change vs base: base %+v rotated %+v", base, rotated)
	}
}

func TestBevelMaterialGate(t *testing.T) {
	flat := renderCenter(t, bevelRect(`<a:tint val="40000"/><a:alpha val="90000"/>`, ""))
	scene := `<a:scene3d><a:camera prst="orthographicFront"/><a:lightRig rig="chilly" dir="t"/></a:scene3d>`

	// dkEdge must light the face flat — centre indistinguishable from the
	// un-bevelled fill. The old code brightened it by 1.19, so this bit.
	dk := renderCenter(t, bevelRect(`<a:tint val="40000"/><a:alpha val="90000"/>`,
		scene+`<a:sp3d extrusionH="1700" prstMaterial="dkEdge"><a:bevelT w="25400" h="6350" prst="softRound"/></a:sp3d>`))
	if absDiff(int(dk.R), int(flat.R)) > 6 || absDiff(int(dk.G), int(flat.G)) > 6 || absDiff(int(dk.B), int(flat.B)) > 6 {
		t.Errorf("dkEdge bevel centre %+v should equal flat %+v (no face gain)", dk, flat)
	}

	// translucentPowder lifts the face mildly: brighter than flat, but well
	// short of the old 1.19 reading (which gave ~229,247,255). The probe pinned
	// COM at 222,236,250.
	pw := renderCenter(t, bevelRect(`<a:tint val="40000"/><a:alpha val="90000"/>`,
		scene+`<a:sp3d prstMaterial="translucentPowder"><a:bevelT w="127000" h="25400" prst="softRound"/></a:sp3d>`))
	if pw.R <= flat.R {
		t.Errorf("translucentPowder face %+v not brighter than flat %+v", pw, flat)
	}
	if pw.R >= 228 {
		t.Errorf("translucentPowder face R=%d too bright (old 1.19 bug gave ~229)", pw.R)
	}
	if absDiff(int(pw.R), 222) > 8 || absDiff(int(pw.G), 236) > 8 || absDiff(int(pw.B), 250) > 8 {
		t.Errorf("translucentPowder face %+v, want ~{222,236,250} (COM probe)", pw)
	}
}

func absDiff(a, b int) int {
	if a > b {
		return a - b
	}
	return b - a
}

// TestSchemeClrHueOffInPipeline mirrors TestHueOffRotatesFillInPipeline but
// drives a schemeClr (val="accent5") instead of srgbClr — that is exactly how
// slide47's SmartArt columns are coloured, and is the path that was actually
// broken before r68.
func TestSchemeClrHueOffInPipeline(t *testing.T) {
	base := renderCenter(t, bevelRect(`<a:schemeClr val="accent5"><a:tint val="40000"/><a:alpha val="90000"/></a:schemeClr>`, ""))
	if !(base.B >= base.R && base.B >= base.G) {
		t.Fatalf("schemeClr base not blue-dominant: %+v", base)
	}
	rotated := renderCenter(t, bevelRect(`<a:schemeClr val="accent5"><a:tint val="40000"/><a:alpha val="90000"/><a:hueOff val="-10740482"/></a:schemeClr>`, ""))
	if rotated.B >= rotated.R && rotated.B >= rotated.G {
		t.Errorf("schemeClr hueOff fill still blue-dominant: %+v (hueOff was ignored on the schemeClr path)", rotated)
	}
	if absDiff(int(rotated.R), int(base.R)) < 8 && absDiff(int(rotated.B), int(base.B)) < 8 {
		t.Errorf("schemeClr hueOff produced no change vs base: base %+v rotated %+v", base, rotated)
	}
}
