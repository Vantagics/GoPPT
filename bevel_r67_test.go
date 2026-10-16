package gopresentation

import (
	"bytes"
	"image"
	"testing"
)

// A shape can carry its own <a:sp3d><a:bevelT> in its spPr — not a theme style
// reference, a literal bevel. PowerPoint lights the face and rim of such a
// shape, and SmartArt cached drawings lean on it heavily: every node of
// customGeo drawing10 (its column bodies and header bars) declares bevelT
// softRound w=25400 h=6350, and the lit sheen is most of what makes those
// columns read as raised panels in PowerPoint's export. The reader used to
// parse bevelT for connectors only, so shape-owned bevels vanished and the
// shapes rendered flat.

// bevelShapeXML is one rectangle with an explicit fill and — optionally — the
// shape-owned sp3d bevel. rect is the geometry every customGeo drawing10 node
// carries, which sends the shape down the rich-text draw path — the one the
// SmartArt panels actually render through.
func bevelShapeXML(withSp3d bool) string {
	sp3d := ""
	if withSp3d {
		sp3d = `<a:sp3d><a:bevelT w="25400" h="6350" prst="softRound"/></a:sp3d>`
	}
	return `<p:sp>
  <p:nvSpPr><p:cNvPr id="2" name="Box"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="500000" y="500000"/><a:ext cx="3000000" cy="2000000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="4BACC6"/></a:solidFill>
    ` + sp3d + `
  </p:spPr>
  <p:txBody>
    <a:bodyPr/>
    <a:lstStyle/>
    <a:p><a:r><a:rPr lang="en-US" sz="1800"/><a:t>item</a:t></a:r></a:p>
  </p:txBody>
</p:sp>`
}

func bevelDeck(t *testing.T, withSp3d bool) []byte {
	t.Helper()
	pres := New()
	pres.AddSlide(newSlide()) // any slide, so slide1.xml gets written
	parts := zipParts(t, writeToBytes(t, pres))
	parts["ppt/slides/slide1.xml"] = []byte(bodyPrSlide(bevelShapeXML(withSp3d)))
	return buildZip(t, parts)
}

// TestShapeOwnedBevelParsesIntoTheShape is the reader half: the bevel's EMU
// size has to reach the shape the same way the theme-reference bevel does,
// because the renderer gates its shading on it.
func TestShapeOwnedBevelParsesIntoTheShape(t *testing.T) {
	for _, tc := range []struct {
		withSp3d     bool
		wantW, wantH int64
	}{
		{true, 25400, 6350},
		{false, 0, 0},
	} {
		data := bevelDeck(t, tc.withSp3d)
		pres, err := ReadFrom(bytes.NewReader(data), int64(len(data)))
		if err != nil {
			t.Fatalf("read deck: %v", err)
		}
		shapes := pres.GetAllSlides()[0].GetShapes()
		if len(shapes) != 1 {
			t.Fatalf("want 1 shape, got %d", len(shapes))
		}
		a, ok := shapes[0].(*RichTextShape)
		if !ok {
			t.Fatalf("want *RichTextShape, got %T", shapes[0])
		}
		if a.bevelW != tc.wantW || a.bevelH != tc.wantH {
			t.Errorf("withSp3d=%v: bevel = (%d,%d), want (%d,%d)",
				tc.withSp3d, a.bevelW, a.bevelH, tc.wantW, tc.wantH)
		}
	}
}

// TestShapeOwnedBevelLightsTheCanvas is the render half: the same shape with
// and without its sp3d bevel must not draw the same pixels. Before the fix the
// bevel never parsed, the two renders were identical, and SmartArt panels drew
// flat against PowerPoint's lit ones.
func TestShapeOwnedBevelLightsTheCanvas(t *testing.T) {
	render := func(withSp3d bool) image.Image {
		data := bevelDeck(t, withSp3d)
		pres, err := ReadFrom(bytes.NewReader(data), int64(len(data)))
		if err != nil {
			t.Fatalf("read deck: %v", err)
		}
		return renderSlide0(t, pres, 1600)
	}
	flat := render(false)
	lit := render(true)

	box := lit.Bounds()
	diff := 0
	for y := box.Min.Y; y < box.Max.Y; y++ {
		for x := box.Min.X; x < box.Max.X; x++ {
			r1, g1, b1, _ := flat.At(x, y).RGBA()
			r2, g2, b2, _ := lit.At(x, y).RGBA()
			if r1 != r2 || g1 != g2 || b1 != b2 {
				diff++
			}
		}
	}
	// The bevel is a 25400 EMU rim (~4px at this width) plus face lighting on
	// a 525x350px panel: a handful of pixels would mean a rounding accident,
	// not a lit shape.
	if diff < 500 {
		t.Errorf("a beveled shape differs from a flat one in only %d pixels; the bevel never reached the canvas", diff)
	}
}
