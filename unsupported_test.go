package gopresentation

import (
	"image"
	"image/color"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

// This file covers the "unsupported shape" behaviour: a construct the reader
// recognises but cannot draw must not disappear. Before this, a SmartArt frame
// left no trace at all, so a preview showed a blank region that looked the same
// as a correctly rendered empty shape — which is the worst possible failure
// mode, because a generation bug becomes invisible.

const diagramURI = "http://schemas.openxmlformats.org/drawingml/2006/diagram"

// shapeTypes summarises a shape list for failure messages, where the count alone
// would not say whether the frame came back as the wrong kind or not at all.
func shapeTypes(shapes []Shape) string {
	parts := make([]string, 0, len(shapes))
	for _, s := range shapes {
		parts = append(parts, s.GetType().String())
	}
	return strings.Join(parts, ", ")
}

// replaceGraphicData rewrites the single <a:graphicData> element in a slide's
// XML. That lets a test turn the fixture's chart frame into a SmartArt frame (or
// any other construct) without hand-writing a whole OOXML package.
func replaceGraphicData(t *testing.T, slideXML, uri, body string) string {
	t.Helper()
	start := strings.Index(slideXML, "<a:graphicData")
	if start < 0 {
		t.Fatal("slide XML has no <a:graphicData>; cannot build the fixture")
	}
	openEnd := strings.Index(slideXML[start:], ">")
	if openEnd < 0 {
		t.Fatal("malformed <a:graphicData>")
	}
	openEnd += start
	end := strings.Index(slideXML[openEnd:], "</a:graphicData>")
	if end < 0 {
		t.Fatal("unterminated <a:graphicData>")
	}
	end += openEnd
	const closeTag = "</a:graphicData>"
	return slideXML[:start] +
		`<a:graphicData uri="` + uri + `">` + body + closeTag +
		slideXML[end+len(closeTag):]
}

// writeTempPPTX writes raw package bytes to a temp .pptx and returns the path.
func writeTempPPTX(t *testing.T, data []byte) string {
	t.Helper()
	path := filepath.Join(t.TempDir(), "fixture.pptx")
	if err := os.WriteFile(path, data, 0o644); err != nil {
		t.Fatalf("write fixture: %v", err)
	}
	return path
}

// smartArtPPTX builds a package identical to the chart fixture except that the
// graphicFrame holds a SmartArt diagram instead of a chart.
func smartArtPPTX(t *testing.T) []byte {
	t.Helper()
	parts := zipParts(t, writeToBytes(t, chartSlidePresentation()))
	const slidePart = "ppt/slides/slide1.xml"
	if _, ok := parts[slidePart]; !ok {
		t.Fatalf("fixture package has no %s", slidePart)
	}
	// The referenced relationships do not need to exist: classification is by
	// the graphicData uri, which is what a real deck would carry.
	body := `<dgm:relIds xmlns:dgm="` + diagramURI + `"` +
		` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"` +
		` r:dm="rId2" r:lo="rId3" r:qs="rId4" r:cs="rId5"/>`
	parts[slidePart] = []byte(replaceGraphicData(t, string(parts[slidePart]), diagramURI, body))
	return buildZip(t, parts)
}

// firstUnsupported returns the first UnsupportedShape in shapes.
func firstUnsupported(shapes []Shape) *UnsupportedShape {
	for _, s := range shapes {
		if u, ok := s.(*UnsupportedShape); ok {
			return u
		}
	}
	return nil
}

func TestUnsupportedShapeTypeString(t *testing.T) {
	if got := ShapeTypeUnsupported.String(); got != "Unsupported" {
		t.Errorf("ShapeTypeUnsupported.String() = %q, want %q", got, "Unsupported")
	}
}

// TestReaderKeepsStandInForSmartArt is the core regression test: a SmartArt
// frame must come back as a visible stand-in carrying the original geometry,
// not vanish.
func TestReaderKeepsStandInForSmartArt(t *testing.T) {
	pres, err := Open(writeTempPPTX(t, smartArtPPTX(t)))
	if err != nil {
		t.Fatalf("Open: %v", err)
	}

	shapes := pres.GetAllSlides()[0].GetShapes()
	got := firstUnsupported(shapes)
	if got == nil {
		t.Fatalf("no UnsupportedShape among %d shapes [%s]; the SmartArt frame was dropped silently",
			len(shapes), shapeTypes(shapes))
	}
	if r := got.GetReason(); !strings.Contains(r, "SmartArt") {
		t.Errorf("GetReason() = %q, want it to name SmartArt", r)
	}
	if ct := got.GetContentType(); !strings.Contains(ct, "/diagram") {
		t.Errorf("GetContentType() = %q, want the diagram uri", ct)
	}
	if label := got.Label(); !strings.Contains(label, "SmartArt") {
		t.Errorf("Label() = %q, want it to mention SmartArt", label)
	}
	if got.GetType() != ShapeTypeUnsupported {
		t.Errorf("GetType() = %v, want ShapeTypeUnsupported", got.GetType())
	}

	// The frame geometry has to survive, otherwise the placeholder would not
	// land where the original graphic sat.
	if got.GetWidth() != 8000000 || got.GetHeight() != 4500000 {
		t.Errorf("placeholder geometry = %dx%d, want 8000000x4500000",
			got.GetWidth(), got.GetHeight())
	}
	if got.GetOffsetX() != 500000 || got.GetOffsetY() != 1000000 {
		t.Errorf("placeholder offset = (%d,%d), want (500000,1000000)",
			got.GetOffsetX(), got.GetOffsetY())
	}

	if all := pres.UnsupportedShapes(); len(all) != 1 {
		t.Errorf("UnsupportedShapes() returned %d entries, want 1", len(all))
	}
}

// TestReaderKeepsStandInForUnreadableChart covers the other half of the same
// problem: when the chart part is missing the frame is still in the file, so the
// reader must say so instead of rendering a blank chart area.
func TestReaderKeepsStandInForUnreadableChart(t *testing.T) {
	parts := zipParts(t, writeToBytes(t, chartSlidePresentation()))
	chartPart := firstPartMatching(parts, func(n string) bool {
		return strings.HasPrefix(n, "ppt/charts/")
	})
	if chartPart == "" {
		t.Fatal("fixture package has no chart part; the test would not exercise the path")
	}
	// Remove the target but keep the frame and its relationship, which is what a
	// truncated or hand-edited file looks like.
	delete(parts, chartPart)

	pres, err := Open(writeTempPPTX(t, buildZip(t, parts)))
	if err != nil {
		t.Fatalf("Open: %v", err)
	}

	shapes := pres.GetAllSlides()[0].GetShapes()
	got := firstUnsupported(shapes)
	if got == nil {
		t.Fatalf("no UnsupportedShape among %d shapes [%s]; an unreadable chart went unnoticed",
			len(shapes), shapeTypes(shapes))
	}
	if r := got.GetReason(); !strings.Contains(r, "could not be read") {
		t.Errorf("GetReason() = %q, want it to report the unreadable part", r)
	}
	if got.GetWidth() != 8000000 {
		t.Errorf("placeholder width = %d, want 8000000 (the original frame's width)", got.GetWidth())
	}
}

// TestNormalDeckHasNoUnsupportedShapes guards against false positives: a deck
// built entirely from supported shapes must report none, or the check would be
// useless in a build pipeline.
func TestNormalDeckHasNoUnsupportedShapes(t *testing.T) {
	pres, err := Open(writeTempPPTX(t, writeToBytes(t, chartSlidePresentation())))
	if err != nil {
		t.Fatalf("Open: %v", err)
	}
	if all := pres.UnsupportedShapes(); len(all) != 0 {
		t.Errorf("UnsupportedShapes() = %d entries [%s], want none for a fully supported deck",
			len(all), shapeTypes(pres.GetAllSlides()[0].GetShapes()))
	}
}

// TestUnsupportedShapesFoundInsideGroups checks the enumeration descends into
// groups, where a diagram is just as likely to hide.
func TestUnsupportedShapesFoundInsideGroups(t *testing.T) {
	p := New()
	g := NewGroupShape()
	inner := NewGroupShape()
	u := NewUnsupportedShape("SmartArt diagram")
	u.BaseShape.SetOffsetX(100).SetOffsetY(200).SetWidth(3000).SetHeight(4000)
	inner.AddShape(u)
	g.AddShape(inner)
	p.GetActiveSlide().AddShape(g)

	all := p.UnsupportedShapes()
	if len(all) != 1 {
		t.Fatalf("UnsupportedShapes() returned %d entries, want 1 (nested two groups deep)", len(all))
	}
	if all[0] != u {
		t.Errorf("UnsupportedShapes()[0] = %p, want the nested shape %p", all[0], u)
	}
}

// TestUnsupportedShapeRendersVisiblePlaceholder verifies the renderer draws
// something conspicuous rather than leaving the region blank. It checks for the
// placeholder's border colour specifically, so an unrelated mark on the slide
// cannot satisfy it.
func TestUnsupportedShapeRendersVisiblePlaceholder(t *testing.T) {
	fc := NewFontCache()
	opts := goldenOptions(fc)

	p := New()
	u := NewUnsupportedShape("SmartArt diagram")
	u.BaseShape.SetOffsetX(2000000).SetOffsetY(1500000)
	u.BaseShape.SetWidth(4000000).SetHeight(2500000)
	p.GetActiveSlide().AddShape(u)

	img, err := p.SlideToImage(0, opts)
	if err != nil {
		t.Fatalf("render: %v", err)
	}

	box := emuRect(t, p, 2000000, 1500000, 4000000, 2500000, opts.Width)
	if box.Empty() {
		t.Fatal("could not derive the placeholder rectangle")
	}

	if n := countInkIn(img, box); n == 0 {
		t.Fatalf("placeholder rendered nothing inside %v", box)
	}
	if n := countNearColorIn(img, box, unsupportedBorder, 60); n == 0 {
		t.Errorf("placeholder frame missing: no pixels near %v inside %v", unsupportedBorder, box)
	}
	// The fill must actually cover the box, not just outline it, so the region
	// reads as a placeholder rather than as a thin stray rectangle.
	if n := countNearColorIn(img, box.Inset(6), unsupportedFill, 12); n == 0 {
		t.Errorf("placeholder interior is not filled with %v inside %v", unsupportedFill, box.Inset(6))
	}
}

// TestUnsupportedShapeOutsideSlideIsHarmless checks a placeholder with bad
// geometry cannot panic or corrupt the canvas.
func TestUnsupportedShapeOutsideSlideIsHarmless(t *testing.T) {
	fc := NewFontCache()
	opts := goldenOptions(fc)

	for _, geom := range [][4]int64{
		{-9000000, -9000000, 4000000, 4000000}, // fully off-slide
		{1000, 1000, 0, 0},                     // zero size
		{1000, 1000, -5000, -5000},             // negative size
	} {
		p := New()
		u := NewUnsupportedShape("SmartArt diagram")
		u.BaseShape.SetOffsetX(geom[0]).SetOffsetY(geom[1])
		u.BaseShape.SetWidth(geom[2]).SetHeight(geom[3])
		p.GetActiveSlide().AddShape(u)

		if _, err := p.SlideToImage(0, opts); err != nil {
			t.Errorf("render with geometry %v: %v", geom, err)
		}
	}
}

// TestUnsupportedShapeIsSkippedByWriter documents the deliberate trade-off: the
// original XML is not retained, so writing the deck back out drops the
// construct. The read must still succeed and the rest of the slide must survive.
func TestUnsupportedShapeIsSkippedByWriter(t *testing.T) {
	pres, err := Open(writeTempPPTX(t, smartArtPPTX(t)))
	if err != nil {
		t.Fatalf("Open: %v", err)
	}
	before := len(pres.GetAllSlides()[0].GetShapes())

	path := filepath.Join(t.TempDir(), "roundtrip.pptx")
	w, err := NewWriter(pres, WriterPowerPoint2007)
	if err != nil {
		t.Fatalf("new writer: %v", err)
	}
	pw, ok := w.(*PPTXWriter)
	if !ok {
		t.Fatalf("writer type = %T, want *PPTXWriter", w)
	}
	if err := pw.Save(path); err != nil {
		t.Fatalf("save: %v", err)
	}

	again, err := Open(path)
	if err != nil {
		t.Fatalf("re-open: %v", err)
	}
	shapes := again.GetAllSlides()[0].GetShapes()
	if len(shapes) != before-1 {
		t.Errorf("shape count after round trip = %d, want %d (the stand-in is not written)",
			len(shapes), before-1)
	}
	if firstUnsupported(shapes) != nil {
		t.Error("the stand-in was written back out; it should be skipped")
	}
	// The title text box must survive, so dropping the diagram is not collateral.
	var texts int
	for _, s := range shapes {
		if s.GetType() == ShapeTypeRichText {
			texts++
		}
	}
	if texts == 0 {
		t.Error("the slide's text shape did not survive the round trip")
	}
}

// countNearColorIn counts pixels within tolerance of want inside rect. It is
// used instead of exact colour matching because anti-aliasing shades the edges.
func countNearColorIn(img image.Image, rect image.Rectangle, want color.RGBA, tolerance int) int {
	return countNearColorInExcept(img, rect, image.Rectangle{}, want, tolerance)
}

// countNearColorOutside counts pixels near want inside outer but outside inner.
// It is how the containment assertions say "spilled out of the frame".
func countNearColorOutside(img image.Image, outer, inner image.Rectangle, want color.RGBA, tolerance int) int {
	return countNearColorInExcept(img, outer, inner, want, tolerance)
}

func countNearColorInExcept(img image.Image, outer, inner image.Rectangle, want color.RGBA, tolerance int) int {
	outer = outer.Intersect(img.Bounds())
	n := 0
	near := func(a, b uint8) bool {
		d := int(a) - int(b)
		if d < 0 {
			d = -d
		}
		return d <= tolerance
	}
	for y := outer.Min.Y; y < outer.Max.Y; y++ {
		for x := outer.Min.X; x < outer.Max.X; x++ {
			if !inner.Empty() && image.Pt(x, y).In(inner) {
				continue
			}
			r, g, b, a := img.At(x, y).RGBA()
			if a>>8 < 8 {
				continue
			}
			if near(uint8(r>>8), want.R) && near(uint8(g>>8), want.G) && near(uint8(b>>8), want.B) {
				n++
			}
		}
	}
	return n
}

// TestUnsupportedPlaceholderLabelStaysInsideFrame pins the containment rule for
// small placeholders. The label is text the library generates rather than text
// from the document, so it must never run outside the frame: it used to be
// drawn at a fixed size with no clipping, so a small placeholder spilled its
// caption over its neighbours and made a correct deck look corrupted.
func TestUnsupportedPlaceholderLabelStaysInsideFrame(t *testing.T) {
	fc := NewFontCache()
	opts := goldenOptions(fc)

	cases := []struct {
		name  string
		box   [4]int64 // x, y, w, h in EMU
		drawn bool     // whether a legible label is expected in the box
	}{
		// Narrow: the natural-size label cannot fit, so it must shrink.
		{"narrow", [4]int64{2000000, 1500000, 1600000, 1200000}, true},
		// Tiny: even the smallest label is too wide and has to be cropped.
		{"tiny", [4]int64{2000000, 1500000, 900000, 700000}, false},
	}
	for _, c := range cases {
		t.Run(c.name, func(t *testing.T) {
			pres := New()
			u := NewUnsupportedShape("SmartArt diagram")
			u.BaseShape.SetOffsetX(c.box[0]).SetOffsetY(c.box[1])
			u.BaseShape.SetWidth(c.box[2]).SetHeight(c.box[3])
			pres.GetActiveSlide().AddShape(u)

			img, err := pres.SlideToImage(0, opts)
			if err != nil {
				t.Fatalf("render: %v", err)
			}
			box := emuRect(t, pres, c.box[0], c.box[1], c.box[2], c.box[3], opts.Width)
			if box.Dx() < 40 || box.Dy() < 30 {
				t.Fatalf("fixture box %v is too small to test containment meaningfully", box)
			}

			// Tolerance 60 keeps the amber frame out of the match, so only glyph
			// pixels count.
			if c.drawn {
				if n := countNearColorIn(img, box.Inset(4), unsupportedText, 60); n == 0 {
					t.Errorf("no label text inside %v; the shrink-to-fit path drew nothing", box)
				}
			}
			// Whatever it drew has to stay inside the frame.
			ring := image.Rect(box.Min.X-16, box.Min.Y-16, box.Max.X+16, box.Max.Y+16)
			if n := countNearColorOutside(img, ring, box, unsupportedText, 60); n != 0 {
				t.Errorf("label ink escaped the placeholder frame: %d px within %v but outside %v", n, ring, box)
			}
		})
	}
}

// TestPlaceholderBorderShrinksWithTheBox checks the frame weight adapts, so a
// tiny placeholder is not reduced to nothing but border.
func TestPlaceholderBorderShrinksWithTheBox(t *testing.T) {
	cases := []struct {
		w, h, want int
	}{
		{400, 300, 2}, // roomy: full weight
		{48, 48, 2},   // 48/12 = 4, still clamped to 2
		{24, 24, 2},   // 24/12 = 2
		{12, 12, 1},   // 12/12 = 1
		{4, 4, 1},     // never zero, or the frame would vanish
	}
	for _, c := range cases {
		if got := placeholderBorderWidth(c.w, c.h); got != c.want {
			t.Errorf("placeholderBorderWidth(%d, %d) = %d, want %d", c.w, c.h, got, c.want)
		}
	}
}
