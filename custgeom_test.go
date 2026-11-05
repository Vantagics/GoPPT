package gopresentation

import "testing"

func TestEvalCustGeomGuides(t *testing.T) {
	guides := []custGuide{
		{"f0", "val 10800000"},
		{"f1", "val 5400000"},
		{"f2", "val 180"},
		{"f3", "val w"},
		{"f4", "val h"},
		{"f5", "val 0"},
		{"f6", "val 1567160"},
		{"f8", "val 99515"},
		{"f26", "+- 0 0 -90"},
		{"f27", "*/ f3 1 1567160"},
		{"f28", "*/ f4 1 995146"},
		{"f32", "*/ f26 f0 1"},
		{"f35", "*/ f32 1 f2"},
		{"f50", "+- f35 0 f1"},
		{"fA", "pin 500 0 100"},
		{"fB", "pin -5 0 100"},
		{"fC", "pin 55 0 100"},
		{"fD", "?: 1 77 42"},
		{"fE", "?: 0 77 42"},
		{"fF", "abs -13"},
		{"fG", "min 9 3 6"},
		{"fH", "max 9 3 6"},
		{"fI", "ss"},
	}
	vals := evalCustGeomGuides(guides, 1567160, 995146)
	cases := map[string]float64{
		"f5":  0,
		"f8":  99515,
		"f26": 90,            // 0+0-(-90)
		"f27": 1,             // w/pathW
		"f28": 1,             // h/pathH
		"f32": 90 * 10800000, // f26*f0
		"f35": 90 * 10800000 / 180,
		"f50": 90*10800000/180 - 5400000,
		"fA":  100,
		"fB":  0,
		"fC":  55,
		"fD":  77,
		"fE":  42,
		"fF":  13,
		"fG":  3,
		"fH":  9,
		"fI":  995146,
	}
	for name, want := range cases {
		if got := vals[name]; got != want {
			t.Errorf("%s = %v, want %v", name, got, want)
		}
	}
}

func TestEvalCustGeomGuidesPathSpaceFallback(t *testing.T) {
	// A path with no w/h attributes uses the shape box as its coordinate
	// space — the "w"/"h" guides then equal the shape ext.
	vals := evalCustGeomGuides(nil, 1566863, 995363)
	if vals["w"] != 1566863 || vals["h"] != 995363 {
		t.Errorf("w/h built-ins wrong: %v/%v", vals["w"], vals["h"])
	}
	if vals["ss"] != 995363 {
		t.Errorf("ss = %v, want 995363", vals["ss"])
	}
}

// The pasted-diagram freeform round trip: a gdLst-driven custGeom renders its
// fill and border where the pre-guide parser produced a degenerate empty
// polygon (only the shadow showed).
func TestCustGeomGuideDrivenFreeformRenders(t *testing.T) {
	r, _ := NewReader(ReaderPowerPoint2007)
	pres, err := r.Read(`F:/pptx_corpus/web/customGeo.pptx`)
	if err != nil {
		t.Skip("corpus deck not available")
	}
	slide := pres.slides[5]
	var box *RichTextShape
	for _, sh := range slide.shapes {
		if g, ok := sh.(*GroupShape); ok {
			for _, c := range g.shapes {
				if rt, ok := c.(*RichTextShape); ok && rt.name == "Freeform 9" {
					box = rt
				}
			}
		}
	}
	if box == nil {
		t.Fatal("Freeform 9 not found")
	}
	if box.customPath == nil || box.customPath.Width <= 0 {
		t.Fatalf("custom path space not resolved: %+v", box.customPath)
	}
	// The path points must span the declared coordinate space — the
	// guide-driven box path reaches the full width/height of its space.
	var minX, minY, maxX, maxY float64
	first := true
	for _, cmd := range box.customPath.Commands {
		for _, p := range cmd.Pts {
			if first || float64(p.X) < minX {
				minX = float64(p.X)
			}
			if first || float64(p.X) > maxX {
				maxX = float64(p.X)
			}
			if first || float64(p.Y) < minY {
				minY = float64(p.Y)
			}
			if first || float64(p.Y) > maxY {
				maxY = float64(p.Y)
			}
			first = false
		}
	}
	if maxX-minX < float64(box.customPath.Width)/2 {
		t.Errorf("path spans only %v..%v of width %d — guides unresolved?", minX, maxX, box.customPath.Width)
	}
	if maxY-minY < float64(box.customPath.Height)/2 {
		t.Errorf("path spans only %v..%v of height %d — guides unresolved?", minY, maxY, box.customPath.Height)
	}
}
