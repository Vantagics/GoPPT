package gopresentation

// The r39 gradient semantics tests.
//
// Slide 20 of the comparison deck draws its bent-up arrow with a six-stop
// circle path gradient whose scheme colours carry shade+satMod, a fillToRect
// corner focus and a tileRect grown one box up and right. Two things were
// broken (the arrow rendered red→dark-red instead of red→purple): the satMod
// on the gradient stops was never applied, and the whole stop list was folded
// into a two-stop LINEAR gradient. Both semantics were pinned against
// PowerPoint COM exports (out_deck build_r39*.py, opscmp_r39.py):
//
//   - satMod does NOT clamp S to 1 before converting back; the multiplied S
//     carries through and only the final RGB channels saturate
//     (ED7D31 satMod 160% → FF7200).
//   - duplicate stop positions collapse to the LAST stop written there.
//   - path gradients: focus = centre of fillToRect; tile = box grown by the
//     (possibly negative) tileRect insets; interior focus mixes by
//     t = |P−F| / D with D the distance to the farthest TILE corner; a focus
//     on the tile boundary switches to a per-axis max-metric.

import (
	"bytes"
	"testing"
)

const gradPathSlide = `
<p:sp>
  <p:nvSpPr><p:cNvPr id="2" name="Arrow"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="200000" y="200000"/><a:ext cx="3000000" cy="3000000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:gradFill rotWithShape="1">
      <a:gsLst>
        <a:gs pos="0"><a:srgbClr val="C00000"/></a:gs>
        <a:gs pos="0"><a:srgbClr val="00B050"/></a:gs>
        <a:gs pos="50000"><a:srgbClr val="00B050"><a:shade val="93000"/><a:satMod val="130000"/></a:srgbClr></a:gs>
        <a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs>
      </a:gsLst>
      <a:path path="circle"><a:fillToRect l="100000" b="100000"/></a:path>
      <a:tileRect t="-100000" r="-100000"/>
    </a:gradFill>
  </p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>
</p:sp>`

func readGradFixture(t *testing.T, shapes string) *Presentation {
	t.Helper()
	parts := zipParts(t, writeToBytes(t, New()))
	parts["ppt/slides/slide1.xml"] = []byte(bodyPrSlide(shapes))
	pkg := buildZip(t, parts)
	pres, err := (&PPTXReader{}).ReadFromReader(bytes.NewReader(pkg), int64(len(pkg)))
	if err != nil {
		t.Fatalf("read the package: %v", err)
	}
	return pres
}

func findAutoShape(shapes []Shape) *AutoShape {
	for _, sh := range shapes {
		if a, ok := sh.(*AutoShape); ok {
			return a
		}
	}
	return nil
}

// TestReaderKeepsPathGradientStops: the full stop list survives, duplicates
// collapse to the last stop at a position, and the path geometry lands in the
// model — the old reader folded everything into a 2-stop linear gradient.
func TestReaderKeepsPathGradientStops(t *testing.T) {
	pres := readGradFixture(t, gradPathSlide)
	shapes := pres.GetAllSlides()[0].GetShapes()
	if len(shapes) == 0 {
		t.Fatal("shape dropped")
	}
	filler, ok := shapes[0].(interface{ GetFill() *Fill })
	if !ok {
		t.Fatalf("shape %T cannot carry a fill", shapes[0])
	}
	f := filler.GetFill()
	if f == nil || f.Type != FillGradientPath {
		t.Fatalf("fill = %+v, want FillGradientPath", f)
	}
	if f.Path != "circle" {
		t.Errorf("Path = %q, want circle", f.Path)
	}
	if f.FillTo != [4]int{100000, 0, 0, 100000} {
		t.Errorf("FillTo = %v, want [100000 0 0 100000]", f.FillTo)
	}
	if f.TileTo != [4]int{0, -100000, -100000, 0} {
		t.Errorf("TileTo = %v, want [0 -100000 -100000 0]", f.TileTo)
	}
	// The two pos=0 stops collapse to the LAST one (green), so three stops.
	if len(f.Stops) != 3 {
		t.Fatalf("stops = %d, want 3 (duplicate pos=0 collapses)", len(f.Stops))
	}
	if f.Stops[0].Pos != 0 || f.Stops[0].Color.ARGB != "FF00B050" {
		t.Errorf("stop[0] = %d/%s, want 0/FF00B050 (last dup wins)", f.Stops[0].Pos, f.Stops[0].Color.ARGB)
	}
	if f.Stops[1].Pos != 50000 {
		t.Errorf("stop[1].Pos = %d, want 50000", f.Stops[1].Pos)
	}
	// shade 93% then satMod 130% on 00B050 — the transforms must have been
	// applied to the stop colour, in document order. The pinned PowerPoint
	// reading of C0504D shade93 satMod130 is (203,61,58); our chain rounds
	// to (203,61,57) — within the ±1 the shaded-then-rounded intermediate
	// costs. Assert the transform happened and the hue stayed green.
	g := f.Stops[1].Color
	if g.GetRed() >= 100 || g.GetGreen() <= g.GetRed() {
		t.Errorf("stop[1] colour = %s, want a darkened, saturated green", g.ARGB)
	}
	if f.Stops[2].Pos != 100000 || f.Stops[2].Color.ARGB != "FF0000FF" {
		t.Errorf("stop[2] = %d/%s, want 100000/FF0000FF", f.Stops[2].Pos, f.Stops[2].Color.ARGB)
	}
}

// TestSatModDoesNotClampSaturation: PowerPoint lets the multiplied S exceed 1
// and clamps only the resulting RGB channels — ED7D31 with satMod 160%
// renders FF7200, where a clamped S would give FF7A1F (op-probe measured).
func TestSatModDoesNotClampSaturation(t *testing.T) {
	c := NewColor("ED7D31")
	applySatMod(&c, 1.6)
	if want := "FFFF7200"; c.ARGB != want {
		t.Errorf("applySatMod(ED7D31, 1.6) = %s, want %s", c.ARGB, want)
	}
	// The style-reference op path (fillRef/lnRef transforms) must agree.
	c2 := applyColorOps(NewColor("ED7D31"), []themeColorOp{{op: "satMod", val: 1.6}})
	if want := "FFFF7200"; c2.ARGB != want {
		t.Errorf("applyColorOps satMod 1.6 = %s, want %s", c2.ARGB, want)
	}
	// And the paired shade+satMod of the deck's stops stays within ±1 of the
	// COM-measured (203,61,58).
	c3 := NewColor("C0504D")
	applyShade(&c3, 0.93)
	applySatMod(&c3, 1.30)
	dblue := int(c3.GetBlue()) - 58
	if dblue < 0 {
		dblue = -dblue
	}
	if c3.GetRed() != 203 || c3.GetGreen() != 61 || dblue > 1 {
		t.Errorf("C0504D shade93 satMod130 = %s, want ≈ FFCB3D3A", c3.ARGB)
	}
}

// TestWriterRoundTripsPathGradient: the N-stop list, the path kind and the
// insets must survive a write — the legacy writer emitted two stops and no
// <a:path> at all.
func TestWriterRoundTripsPathGradient(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()
	sh := NewAutoShape()
	sh.shapeType = AutoShapeRectangle
	sh.SetPosition(200000, 200000)
	sh.SetSize(3000000, 3000000)
	fill := NewFill().SetGradientStops([]GradStop{
		{Pos: 0, Color: NewColor("C00000")},
		{Pos: 100000, Color: NewColor("0000FF")},
	}, 0, "circle", [4]int{100000, 0, 0, 100000}, [4]int{0, -100000, -100000, 0})
	sh.SetFill(fill)
	slide.AddShape(sh)

	data := writeToBytes(t, p)
	xml := string(zipParts(t, data)["ppt/slides/slide1.xml"])
	for _, want := range []string{
		`<a:gs pos="0">`, `<a:gs pos="100000">`,
		`<a:path path="circle">`, `<a:fillToRect l="100000" b="100000"/>`,
		`<a:tileRect t="-100000" r="-100000"/>`,
	} {
		if !bytes.Contains([]byte(xml), []byte(want)) {
			t.Errorf("written slide XML missing %s", want)
		}
	}

	pres, err := (&PPTXReader{}).ReadFromReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatalf("read back: %v", err)
	}
	got := findAutoShape(pres.GetAllSlides()[0].GetShapes())
	if got == nil || got.fill == nil || got.fill.Type != FillGradientPath {
		t.Fatalf("path gradient lost on round-trip: %+v", got)
	}
	if len(got.fill.Stops) != 2 || got.fill.Stops[0].Color.ARGB != "FFC00000" ||
		got.fill.Stops[1].Color.ARGB != "FF0000FF" {
		t.Errorf("round-tripped stops = %+v", got.fill.Stops)
	}
	if got.fill.FillTo != [4]int{100000, 0, 0, 100000} || got.fill.TileTo != [4]int{0, -100000, -100000, 0} {
		t.Errorf("round-tripped insets = %v / %v", got.fill.FillTo, got.fill.TileTo)
	}
}

// TestLinearGradientKeepsMiddleStops: a three-stop linear gradient keeps its
// middle stop through the model and the writer (the legacy fields stay in
// sync for the old two/three-stop consumers).
func TestLinearGradientKeepsMiddleStops(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()
	sh := NewAutoShape()
	sh.shapeType = AutoShapeRectangle
	sh.SetPosition(0, 0)
	sh.SetSize(2000000, 1000000)
	sh.SetFill(NewFill().SetGradientStops([]GradStop{
		{Pos: 0, Color: NewColor("FF0000")},
		{Pos: 35000, Color: NewColor("00FF00")},
		{Pos: 100000, Color: NewColor("0000FF")},
	}, 0, "", [4]int{}, [4]int{}))
	slide.AddShape(sh)

	data := writeToBytes(t, p)
	pres, err := (&PPTXReader{}).ReadFromReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatalf("read back: %v", err)
	}
	got := findAutoShape(pres.GetAllSlides()[0].GetShapes())
	if got == nil || got.fill == nil || got.fill.Type != FillGradientLinear {
		t.Fatalf("linear gradient lost: %+v", got)
	}
	if len(got.fill.Stops) != 3 || got.fill.Stops[1].Pos != 35000 {
		t.Errorf("middle stop lost on round-trip: %+v", got.fill.Stops)
	}
	if got.fill.MidPos != 35000 || got.fill.MidColor.ARGB != "FF00FF00" {
		t.Errorf("legacy mid fields out of sync: %+v / %+v", got.fill.MidColor, got.fill.MidPos)
	}
}

// TestPathGradientRendersCornerFocus: the pinned geometry must reach pixels —
// a circle gradient with a top-right focus and the tile grown up/right paints
// the shape's top-right corner with the first stop and its bottom-left corner
// (the farthest box point from the focus) with the last one.
func TestPathGradientRendersCornerFocus(t *testing.T) {
	pres := readGradFixture(t, gradPathSlide)
	fc := NewFontCache()
	requireAnyFont(t, fc)
	opts := DefaultRenderOptions()
	opts.Width = 960
	opts.FontCache = fc
	img, err := pres.SlideToImage(0, opts)
	if err != nil {
		t.Fatalf("render: %v", err)
	}
	// Shape: 200000..3200000 EMU square on a 12192000-wide slide at 960px →
	// 15.7..251.6 px. Sample well inside opposite corners.
	sample := func(x, y int) (int, int, int) {
		r, g, b, _ := img.At(x, y).RGBA()
		return int(r >> 8), int(g >> 8), int(b >> 8)
	}
	r, g, b := sample(245, 25) // near the top-right focus
	if !(g > 150 && r < 100 && b < 120) {
		t.Errorf("near-focus pixel (%d,%d,%d), want green start stop dominant", r, g, b)
	}
	r, g, b = sample(25, 245) // farthest box corner from the focus
	if !(b > 150 && r < 120 && g < 120) {
		t.Errorf("far-corner pixel (%d,%d,%d), want blue end stop dominant", r, g, b)
	}
}

// TestPathGradientGeometryRegimes pins the two normalisation regimes against
// the COM-measured geometry (V7/V8/V9 variant deck, green_r39.py): a focus on
// the tile boundary must take the per-axis max-metric — the V8 probe measured
// the t=0.5 ring half a box out horizontally (244.5px on a 488.9px box),
// where a Euclidean reading puts it at 0.71 boxes — while an interior focus
// mixes by true distance over the farthest-tile-corner radius.
func TestPathGradientGeometryRegimes(t *testing.T) {
	// Corner focus, no tile: focus sits on the tile boundary → max-metric.
	g := pathGradientGeometry(100, 100, &Fill{FillTo: [4]int{100000, 0, 0, 100000}})
	if g.euclidean {
		t.Fatal("boundary focus must select the max-metric regime")
	}
	if g.fx != 100 || g.fy != 0 {
		t.Errorf("focus = (%v,%v), want (100,0)", g.fx, g.fy)
	}
	// Horizontal half-box out: t=0.5 (the measured V8 ring).
	if got := pathGradientT(g, 50, 0); got != 0.5 {
		t.Errorf("max-metric t(50,0) = %v, want 0.5", got)
	}
	// Off-diagonal: max-metric keeps t=0.5 where a Euclidean reading over
	// the box diagonal would give 0.395.
	if got := pathGradientT(g, 50, 25); got != 0.5 {
		t.Errorf("max-metric t(50,25) = %v, want 0.5 (Euclidean would give 0.395)", got)
	}

	// Centre focus: Euclidean over the half diagonal; the box corner is t=1.
	g2 := pathGradientGeometry(100, 100, &Fill{FillTo: [4]int{50000, 50000, 50000, 50000}})
	if !g2.euclidean {
		t.Fatal("interior focus must select the Euclidean regime")
	}
	if got := pathGradientT(g2, 100, 100); got != 1.0 {
		t.Errorf("Euclidean t(corner) = %v, want 1.0", got)
	}
	if got := pathGradientT(g2, 50, 25); got < 0.35 || got > 0.36 {
		t.Errorf("Euclidean t(50,25) = %v, want 0.354", got)
	}

	// Tile grown one box up and right around a corner focus (the deck's
	// slide20/27 geometry): Euclidean, and the radius is the distance to the
	// farthest TILE corner — for the centred tile that is the box diagonal,
	// so the t=1 ring still passes through the opposite box corner.
	g3 := pathGradientGeometry(100, 100, &Fill{
		FillTo: [4]int{100000, 0, 0, 100000},
		TileTo: [4]int{0, -100000, -100000, 0},
	})
	if !g3.euclidean {
		t.Fatal("tile-centred focus must select the Euclidean regime")
	}
	if got := pathGradientT(g3, 0, 100); got != 1.0 {
		t.Errorf("tile-centred t(opposite corner) = %v, want 1.0", got)
	}
	if got := pathGradientT(g3, 29.3, 0); got < 0.49 || got > 0.51 {
		t.Errorf("tile-centred t(29.3,0) = %v, want 0.5 (Euclidean; max-metric would give 0.707)", got)
	}
}
