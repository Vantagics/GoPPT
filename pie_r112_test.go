package gopresentation

import (
	"image"
	"math"
	"testing"
)

func TestPieWedgePointsAngles(t *testing.T) {
	adj := map[string]int{"adj1": 9000000, "adj2": 16200000}
	pts := pieWedgePoints(0, 0, 400, 400, adj)
	if len(pts) < 3 {
		t.Fatalf("wedge collapsed: %d points", len(pts))
	}
	// First point is the centre; the rest walk the arc from 150° to 270°.
	cx, cy := 200.0, 200.0
	if math.Abs(pts[0].x-cx) > 1 || math.Abs(pts[0].y-cy) > 1 {
		t.Errorf("first point not the centre: %v", pts[0])
	}
	last := pts[len(pts)-1]
	// 270° on the screen clock: x = cx, y = cy - r … with y down, sin(270°)=-1 → cy - r.
	if math.Abs(last.x-cx) > 2 || math.Abs(last.y-(cy-200)) > 2 {
		t.Errorf("arc end at %v, want (200, 0) for 270°", last)
	}
	// The midpoint of the sweep sits at 210°: cos=−0.866, sin=−0.5.
	mid := pts[len(pts)/2]
	if cx+200*math.Cos(210*math.Pi/180)-mid.x > 6 {
		t.Errorf("arc drifts from the ellipse: mid=%v", mid)
	}
}

func TestPieWedgeWrapAround(t *testing.T) {
	// adj2 30° < adj1 270°: the sweep wraps through 360° — 120° of arc.
	pts := pieWedgePoints(0, 0, 400, 400, map[string]int{"adj1": 16200000, "adj2": 1800000})
	// First arc point = circle at 270° (top of the ellipse).
	if got := pts[1]; math.Abs(got.x-200) > 2 || got.y > 2 {
		t.Errorf("arc start %v, want (200, 0)", got)
	}
	// Last arc point at 30°, screen clock (y down): (cx + r*0.866, cy + r*0.5) —
	// down-right of the centre.
	last := pts[len(pts)-1]
	if math.Abs(last.x-(200+200*0.866)) > 3 || math.Abs(last.y-(200+100)) > 3 {
		t.Errorf("arc end %v, want (373, 300)", last)
	}
}

// The prst="pie" preset renders as a filled wedge (formerly the bounding-rect
// default, which stacked three slides' slices into block cards).
func TestPiePresetRendersWedge(t *testing.T) {
	img := image.NewRGBA(image.Rect(0, 0, 200, 200))
	r := &renderer{img: img, scaleX: 1, scaleY: 1}
	as := &AutoShape{
		BaseShape: BaseShape{fill: NewFill().SetSolid(Color{ARGB: "FF4472C4"})},
		shapeType: AutoShapePie,
	}
	as.offsetX, as.offsetY, as.width, as.height = 0, 0, 200, 200
	as.adjustValues = map[string]int{"adj1": 9000000, "adj2": 16200000}
	r.renderAutoShapeFill(as, 0, 0, 200, 200)
	// A point inside the wedge's fan (210° at r≈80: (38, 60)) is filled blue.
	if got := img.RGBAAt(38, 60); got.A == 0 {
		t.Fatalf("wedge interior not painted: %+v", got)
	} else if got.B < got.R || got.B == 0 {
		t.Errorf("wedge interior wrong colour: %+v", got)
	}
	// 0° (right of centre) lies outside a 150°..270° wedge — must keep the
	// background (the fill is transparent, not painted over).
	if out := img.RGBAAt(195, 100); out.A != 0 && out.B > out.R && out.G > out.R {
		t.Errorf("0° region painted though outside the wedge: %+v", out)
	}
}
