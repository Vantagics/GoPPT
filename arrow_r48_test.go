package gopresentation

import (
	"math"
	"testing"
)

// TestArrowPresetGeometryLaw pins the OOXML straight-block-arrow polygon:
// halfShaft = cross*adj1/200000, head = ss*adj2/100000 (ss = min(w,h)),
// defaults adj1 = adj2 = 50000. The old hard-coded shaft 0.4 / head 0.35*w
// drew deck 00022823 slide20's themed arrows with a head one third too long
// and a shaft too thin.
func TestArrowPresetGeometryLaw(t *testing.T) {
	// Defaults, rightArrow 300x200: ss=200, shaft y 50..150, head x 200..300.
	pts := arrowPresetPoints(AutoShapeArrowRight, 0, 0, 300, 200, nil)
	want := []fpoint{
		{0, 50}, {200, 50}, {200, 0}, {300, 100},
		{200, 200}, {200, 150}, {0, 150},
	}
	if len(pts) != len(want) {
		t.Fatalf("rightArrow points = %d, want %d", len(pts), len(want))
	}
	for i, p := range want {
		if math.Abs(pts[i].x-p.x) > 0.01 || math.Abs(pts[i].y-p.y) > 0.01 {
			t.Errorf("rightArrow pt%d = (%v,%v), want (%v,%v)", i, pts[i].x, pts[i].y, p.x, p.y)
		}
	}
	// adj drives both: adj1=30000 (shaft 30% of height), adj2=80000 with
	// ss=200 caps head at 160 of the 300 width.
	pts = arrowPresetPoints(AutoShapeArrowRight, 0, 0, 300, 200, map[string]int{"adj1": 30000, "adj2": 80000})
	if math.Abs(pts[0].y-70) > 0.01 || math.Abs(pts[6].y-130) > 0.01 {
		t.Errorf("adj1=30000 shaft = %v..%v, want 70..130", pts[0].y, pts[6].y)
	}
	if math.Abs(pts[1].x-140) > 0.01 {
		t.Errorf("adj2=80000 head base x = %v, want 140", pts[1].x)
	}
	// downArrow transposes: shaft is a fraction of the WIDTH, head a
	// fraction of ss along the height.
	pts = arrowPresetPoints(AutoShapeArrowDown, 0, 0, 200, 300, nil)
	if math.Abs(pts[0].x-50) > 0.01 || math.Abs(pts[1].x-150) > 0.01 {
		t.Errorf("downArrow shaft x = %v..%v, want 50..150", pts[0].x, pts[1].x)
	}
	if math.Abs(pts[2].y-200) > 0.01 || math.Abs(pts[4].y-300) > 0.01 {
		t.Errorf("downArrow head = base y %v / tip y %v, want 200/300", pts[2].y, pts[4].y)
	}
	// adj2 larger than the box allows pins at the cap: 200 wide, 100 tall,
	// adj2=500000 caps at 100000*h/ss = 100000 (head = h).
	pts = arrowPresetPoints(AutoShapeArrowDown, 0, 0, 200, 100, map[string]int{"adj2": 500000})
	if math.Abs(pts[2].y-0) > 0.01 {
		t.Errorf("pinned head base y = %v, want 0 (head spans the whole height)", pts[2].y)
	}
}

// TestArrowRendersOutlineNotBox is the render sentinel: a filled arrow with
// a border shows the fill inside the shaft, no ink on the bounding box's
// top edge away from the outline (the old border case drew a rectangle
// around the box), and the shaft at the OOXML half-height.
func TestArrowRendersOutlineNotBox(t *testing.T) {
	fc := NewFontCache()
	if pickInstalledFont(fc) == "" {
		t.Skip("no fonts installed on this machine")
	}
	p := New()
	slide := p.GetActiveSlide()
	sh := slide.AddAutoShape()
	sh.SetAutoShapeType(AutoShapeArrowRight)
	sh.SetPosition(914400, 914400) // 72pt
	sh.SetSize(2857500, 1905000)   // 225x150pt
	sh.SetSolidFill(NewColor("C00000"))
	b := NewBorder()
	b.Style = BorderSolid
	b.Width = 2
	b.Color = NewColor("000000")
	sh.SetBorder(b)

	opts := DefaultRenderOptions()
	opts.Width = 720 // 1px per pt
	img, err := p.SlideToImage(0, opts)
	if err != nil {
		t.Fatalf("render: %v", err)
	}
	isRed := func(x, y int) bool {
		r, g, bl, _ := img.At(x, y).RGBA()
		return r>>8 > 120 && g>>8 < 90 && bl>>8 < 90
	}
	isDark := func(x, y int) bool {
		r, g, bl, _ := img.At(x, y).RGBA()
		return r>>8 < 110 && g>>8 < 110 && bl>>8 < 110
	}
	// Box (72,72)-(297,222); ss=150, head 75 → base x=222; shaft 36pt tall
	// → y 108..186.
	if !isRed(150, 147) {
		t.Error("no fill in the shaft centre (150,147)")
	}
	if !isRed(280, 147) {
		t.Error("no fill at the head tip (280,147)")
	}
	// Above the shaft, left of the head: inside the box, outside the shape.
	if isRed(150, 90) {
		t.Error("fill above the shaft (150,90) — bounding-box fill")
	}
	// The old shaft was 0.4*h (y 117..177); the law gives 0.5*h. A pixel
	// between 108 and 117 must be filled now and was empty before.
	if !isRed(150, 112) {
		t.Error("shaft does not reach the OOXML half-height (150,112)")
	}
	// Border: no ink on the box's top edge midway (the old rectangle border
	// put a 2pt line there); the outline's shaft-top edge at y=108 has it.
	if isDark(160, 73) || isDark(160, 72) {
		t.Error("border ink on the bounding-box top edge — rectangle border")
	}
	if !isDark(160, 108) && !isDark(160, 109) {
		t.Error("no border on the shaft's top edge (160,108)")
	}
}
