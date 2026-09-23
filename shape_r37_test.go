package gopresentation

import (
	"math"
	"testing"
)

// The r37 shape semantics tests.
//
// Deck 00022823's slides 15/20/21 use bentUpArrow arrows (five instances,
// one with custom adj values) and a snip2DiagRect picture frame; the
// renderer had no path for either preset and painted solid rectangles.
// The bentUpArrow formulas were pinned against PowerPoint COM exports of
// nine adj variants (see out_deck measure_r37_var.py):
//
//	shaft thickness      = min(adj1, 0.5)·ss
//	head base width      = min(adj2·2, 1)·ss
//	head triangle height = min(adj3, 0.5)·ss
//
// with ss = min(w,h), the head's base right corner on the shape's right
// edge, and the vertical arm centred on the head's base midpoint.

// r := renderer for the pure-geometry helpers.
func r37renderer() *renderer {
	return &renderer{}
}

func TestBentUpArrowDefaultGeometry(t *testing.T) {
	// 200x200 box, default adj (25000, 25000, 25000): ss=200, st=50,
	// tw=100, th=50. Head base right corner on the right edge, apex at
	// (150, 0), arm x 125..175, bar top at y=150.
	pts := r37renderer().bentUpArrowPoints(0, 0, 200, 200, nil)
	want := []fpoint{
		{0, 200},   // bar bottom-left
		{175, 200}, // bar bottom-right
		{175, 50},  // arm right at head base
		{200, 50},  // head base right corner
		{150, 0},   // apex
		{100, 50},  // head base left corner
		{125, 50},  // arm left at head base
		{125, 150}, // arm left at bar top
		{0, 150},   // bar top-left
	}
	if len(pts) != len(want) {
		t.Fatalf("points = %d, want %d: %v", len(pts), len(want), pts)
	}
	for i, p := range want {
		if math.Abs(pts[i].x-p.x) > 0.01 || math.Abs(pts[i].y-p.y) > 0.01 {
			t.Errorf("point[%d] = (%v,%v), want (%v,%v)", i, pts[i].x, pts[i].y, p.x, p.y)
		}
	}
}

func TestBentUpArrowAdjustments(t *testing.T) {
	// adj1=60000 caps the shaft at 0.5·ss; adj2=60000 caps the head width
	// at 1·ss (adj2·2 overflows); adj3=60000 caps the head height at 0.5·ss.
	pts := r37renderer().bentUpArrowPoints(0, 0, 200, 200, map[string]int{
		"adj1": 60000, "adj2": 60000, "adj3": 60000,
	})
	// st=100, tw=200, th=100: cx=100, apex (100,0), base corners (0,100) and
	// (200,100), arm x 50..150, bar top at y=100.
	want := []fpoint{
		{0, 200}, {150, 200}, {150, 100}, {200, 100}, {100, 0}, {0, 100}, {50, 100}, {50, 100}, {0, 100},
	}
	if len(pts) != len(want) {
		t.Fatalf("points = %d, want %d: %v", len(pts), len(want), pts)
	}
	for i, p := range want {
		if math.Abs(pts[i].x-p.x) > 0.01 || math.Abs(pts[i].y-p.y) > 0.01 {
			t.Errorf("point[%d] = (%v,%v), want (%v,%v)", i, pts[i].x, pts[i].y, p.x, p.y)
		}
	}

	// A custom adj1=39000 (deck slide20 instance): shaft = 0.39·ss with no cap.
	// cx = 200 - tw/2 = 150 with default adj2; arm left = 150-39 = 111;
	// bar top = 200-78 = 122.
	pts = r37renderer().bentUpArrowPoints(0, 0, 200, 200, map[string]int{"adj1": 39000})
	if math.Abs(pts[7].x-111) > 0.01 || math.Abs(pts[7].y-122) > 0.01 {
		t.Errorf("arm-left/bar-top point = (%v,%v), want (111,122)", pts[7].x, pts[7].y)
	}
}

func TestSnip2DiagRectGeometry(t *testing.T) {
	// Default adj (16667): both diagonal corners snipped by ss/6.
	pts := r37renderer().snip2DiagRectPoints(0, 0, 300, 200, nil)
	want := []fpoint{
		{0, 0},            // top-left (square)
		{300 - 33.333, 0}, // top-right snip start
		{300, 33.333},     // top-right snip end
		{300, 200},        // bottom-right (square)
		{33.333, 200},     // bottom-left snip end
		{0, 200 - 33.333}, // bottom-left snip start
	}
	if len(pts) != len(want) {
		t.Fatalf("points = %d, want %d: %v", len(pts), len(want), pts)
	}
	for i, p := range want {
		if math.Abs(pts[i].x-p.x) > 0.01 || math.Abs(pts[i].y-p.y) > 0.01 {
			t.Errorf("point[%d] = (%v,%v), want (%v,%v)", i, pts[i].x, pts[i].y, p.x, p.y)
		}
	}

	// adj1=40000 enlarges only the bottom-left snip.
	pts = r37renderer().snip2DiagRectPoints(0, 0, 300, 200, map[string]int{"adj1": 40000})
	if math.Abs(pts[4].x-80) > 0.01 || math.Abs(pts[5].y-120) > 0.01 {
		t.Errorf("bottom-left snip = (%v,%v)/(?,%v), want x=80 / y=120", pts[4].x, pts[4].y, pts[5].y)
	}
	if math.Abs(pts[1].x-(300-33.333)) > 0.01 {
		t.Errorf("top-right snip must stay at default: %v", pts[1].x)
	}
}

// TestBentUpArrowRendersShapeNotRectangle: the render sentinel. The preset
// used to fall through to the default case and paint the bounding box solid.
// A 200x200pt arrow at (72,72)pt renders with fill inside the bar and the
// head, and — the actual regression — no fill in the box's empty top corners.
func TestBentUpArrowRendersShapeNotRectangle(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()
	sh := slide.AddAutoShape()
	sh.SetAutoShapeType(AutoShapeBentUpArrow)
	sh.SetPosition(914400, 914400) // 72pt
	sh.SetSize(2540000, 2540000)   // 200pt
	sh.SetSolidFill(NewColor("C00000"))

	opts := DefaultRenderOptions()
	opts.Width = 720 // 1px per pt
	img, err := p.SlideToImage(0, opts)
	if err != nil {
		t.Fatalf("render: %v", err)
	}
	isRed := func(x, y int) bool {
		r, g, b, _ := img.At(x, y).RGBA()
		return r>>8 > 120 && g>>8 < 90 && b>>8 < 90
	}
	anyRed := func(x, y int) bool {
		for dx := -2; dx <= 2; dx++ {
			for dy := -2; dy <= 2; dy++ {
				if isRed(x+dx, y+dy) {
					return true
				}
			}
		}
		return false
	}
	// Bar interior: y 222..272, x 72..247 (st=50, cx=222, armR=247).
	if !anyRed(120, 250) {
		t.Error("no fill in the horizontal bar (120,250) — shape not painted")
	}
	// Head triangle interior near the apex (222,72), e.g. (222,80).
	if !anyRed(222, 80) {
		t.Error("no fill near the head apex (222,80)")
	}
	// The box's top corners stay empty: bar only spans y>=222 and the head
	// triangle narrows to the apex at x=222.
	if anyRed(85, 85) {
		t.Error("fill in the empty top-left corner — bounding box painted")
	}
	if anyRed(260, 85) {
		t.Error("fill in the empty top-right corner — bounding box painted")
	}
}
