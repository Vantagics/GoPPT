package gopresentation

import (
	"math"
	"testing"
)

// r64 COM probe facts pinned here (probe decks out_deck/r64_lin.pptx,
// r64_neutral.pptx, r64_stops.pptx; exports out_deck/r64/ppt{,n,s}):
//
//   - A TWO-stop linear gradient is rendered by PowerPoint with a symmetric
//     sigmoid position easing t^1.6/(t^1.6+(1-t)^1.6) and interpolation in a
//     gamma-2.2 linearised space (per-box rmse 0.7-5.4 sRGB units; plain sRGB
//     linear interpolation misses by 36-44).
//   - A THREE-or-more-stop linear gradient is rendered in plain sRGB space
//     with NO easing (rmse 0.4-1.8) — the old Go behaviour is already right.

func TestGradEasedRampMidpoints(t *testing.T) {
	ramp := gradEasedRamp([4]uint8{255, 255, 255, 255}, [4]uint8{0, 0, 0, 255})
	const steps = 4096
	// Endpoints and centre must be exact regardless of the easing.
	for i, want := range map[int]uint8{0: 255, steps / 2: 186, steps: 0} {
		// t=0.5: e=0.5, gamma-2.2 midpoint 0.5^(1/2.2)=0.72974 -> 186.
		if got := ramp[i][0]; got != want {
			t.Errorf("ramp[%d].R = %d, want %d", i, got, want)
		}
	}
	// Quarter/tenth points must match the COM-measured white->black profile
	// (observed 251 @ f=0.1, 237 @ f=0.25 interpolated, 51 @ f=0.9 mirror).
	for i, want := range map[int]int{
		steps / 10:       252, // f=0.10
		steps / 4:        237, // f=0.25
		steps - steps/4:  107, // f=0.75 (mirror of 0.25 in g22 space)
		steps - steps/10: 51,  // f=0.90
	} {
		if got := int(ramp[i][0]); got != want {
			t.Errorf("ramp[%d].R = %d, want %d", i, got, want)
		}
	}
	// A plain sRGB ramp would give 191 at the quarter point; the eased
	// gamma-2.2 ramp holds the light end much longer (237).
	if q := ramp[steps/4][0]; q < 220 {
		t.Errorf("ramp[quarter].R = %d looks like plain sRGB interpolation", q)
	}
}

func TestTwoStopLinearGradientRendersEased(t *testing.T) {
	shapes := `<p:sp>
  <p:nvSpPr><p:cNvPr id="2" name="G"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="200000" y="200000"/><a:ext cx="3000000" cy="3000000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:gradFill rotWithShape="1">
      <a:gsLst>
        <a:gs pos="0"><a:srgbClr val="FFFFFF"/></a:gs>
        <a:gs pos="100000"><a:srgbClr val="000000"/></a:gs>
      </a:gsLst>
      <a:lin ang="0" scaled="1"/>
    </a:gradFill>
  </p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>
</p:sp>`
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
	// Default slide is 9144000 EMU (10in) wide; at 960px the 200000..
	// 3200000 EMU box spans 21.0..336.0 px. Sample the horizontal centre
	// row and compare against the quantised ramp.
	y := 178
	sample := func(x int) (int, int, int) {
		r, g, b, _ := img.At(x, y).RGBA()
		return int(r >> 8), int(g >> 8), int(b >> 8)
	}
	ramp := gradEasedRamp([4]uint8{255, 255, 255, 255}, [4]uint8{0, 0, 0, 255})
	boxMin, boxMax := 21.0, 336.0
	for _, f := range []float64{0.1, 0.25, 0.4, 0.5, 0.6, 0.75, 0.9} {
		x := int(boxMin + f*(boxMax-boxMin))
		r, g, b := sample(x)
		want := ramp[int(f*4096+0.5)]
		if math.Abs(float64(r-int(want[0]))) > 2 || math.Abs(float64(g-int(want[1]))) > 2 || math.Abs(float64(b-int(want[2]))) > 2 {
			t.Errorf("f=%.2f pixel (%d,%d,%d), want ramp %v", f, r, g, b, want)
		}
	}
	// Spot-check the eased midpoint against the COM-measured value 185-186.
	r, _, _ := sample(int((boxMin + boxMax) / 2))
	if r < 183 || r > 189 {
		t.Errorf("midpoint pixel R = %d, want ~186 (PowerPoint COM measured 185)", r)
	}
}

func TestThreeStopLinearGradientStaysPlainSRGB(t *testing.T) {
	shapes := `<p:sp>
  <p:nvSpPr><p:cNvPr id="2" name="G"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="200000" y="200000"/><a:ext cx="3000000" cy="3000000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:gradFill rotWithShape="1">
      <a:gsLst>
        <a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs>
        <a:gs pos="50000"><a:srgbClr val="00FF00"/></a:gs>
        <a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs>
      </a:gsLst>
      <a:lin ang="0" scaled="1"/>
    </a:gradFill>
  </p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>
</p:sp>`
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
	// f=0.25 is the exact midpoint of the red->green segment: plain sRGB
	// interpolation gives (128,128,0), and PowerPoint agrees (COM probe
	// rmse 0.48 for this box). An eased gamma-2.2 reading would give
	// roughly (188,188,0). Box spans 21.0..336.0 px at 960px (see above).
	boxMin, boxMax := 21.0, 336.0
	y := 178
	x := int(boxMin + 0.25*(boxMax-boxMin))
	r, g, b, _ := img.At(x, y).RGBA()
	r8, g8, b8 := int(r>>8), int(g>>8), int(b>>8)
	if math.Abs(float64(r8-128)) > 3 || math.Abs(float64(g8-128)) > 3 || b8 != 0 {
		t.Errorf("three-stop f=0.25 pixel (%d,%d,%d), want ~(128,128,0) plain sRGB", r8, g8, b8)
	}
}
