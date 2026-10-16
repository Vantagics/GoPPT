package gopresentation

import (
	"bytes"
	"image"
	"image/png"
	"os"
	"strings"
	"testing"
)

// A SmartArt graphicFrame names four auxiliary parts (the data model, the
// layout, the quick style and the colour transform) that describe how to lay a
// diagram out. PowerPoint also caches a fifth — ppt/diagrams/drawingN.xml —
// holding the diagram's already-resolved geometry, hung off the *slide's*
// relationship list rather than off the dgm:relIds element. Reading that part
// is what turns a visible "unsupported" placeholder into the real diagram.

const nsDiagram = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
const nsDsp = "http://schemas.microsoft.com/office/drawing/2008/diagram"
const relDrawingType = "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing"

// diagramFrameXML is a graphicFrame that is nothing but a diagram reference.
func diagramFrameXML(x, y, cx, cy int) string {
	return `<p:graphicFrame>
  <p:nvGraphicFramePr>
    <p:cNvPr id="3" name="Diagram 1"/>
    <p:cNvGraphicFramePr/>
    <p:nvPr/>
  </p:nvGraphicFramePr>
  <p:xfrm>
    <a:off x="` + itoa(x) + `" y="` + itoa(y) + `"/>
    <a:ext cx="` + itoa(cx) + `" cy="` + itoa(cy) + `"/>
  </p:xfrm>
  <a:graphic>
    <a:graphicData uri="` + nsDiagram + `">
      <dgm:relIds xmlns:dgm="` + nsDiagram + `" xmlns:r="` + nsOfficeDocRels + `" r:dm="rId90" r:lo="rId91" r:qs="rId92" r:cs="rId93"/>
    </a:graphicData>
  </a:graphic>
</p:graphicFrame>`
}

// diagramDrawingXML is a two-node diagram in its resolved form. The element
// local names deliberately match the p:-prefixed shapes the slide scanner
// already reads, which is the whole reason it can be reused unchanged.
func diagramDrawingXML() string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<dsp:drawing xmlns:dsp="` + nsDsp + `" xmlns:a="` + nsDrawingML + `">
  <dsp:spTree>
    <dsp:nvGrpSpPr>
      <dsp:cNvPr id="0" name=""/>
      <dsp:cNvGrpSpPr/>
    </dsp:nvGrpSpPr>
    <dsp:grpSpPr/>
    <dsp:sp modelId="{11111111-1111-1111-1111-111111111111}">
      <dsp:nvSpPr>
        <dsp:cNvPr id="0" name="alpha"/>
        <dsp:cNvSpPr/>
      </dsp:nvSpPr>
      <dsp:spPr>
        <a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="500000"/></a:xfrm>
        <a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>
        <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
      </dsp:spPr>
      <dsp:txBody>
        <a:bodyPr/>
        <a:lstStyle/>
        <a:p><a:r><a:rPr lang="en-US" sz="1800"/><a:t>Alpha</a:t></a:r></a:p>
      </dsp:txBody>
    </dsp:sp>
    <dsp:sp modelId="{22222222-2222-2222-2222-222222222222}">
      <dsp:nvSpPr>
        <dsp:cNvPr id="0" name="beta"/>
        <dsp:cNvSpPr/>
      </dsp:nvSpPr>
      <dsp:spPr>
        <a:xfrm><a:off x="1000000" y="0"/><a:ext cx="1000000" cy="500000"/></a:xfrm>
        <a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>
        <a:solidFill><a:srgbClr val="0000FF"/></a:solidFill>
      </dsp:spPr>
      <dsp:txBody>
        <a:bodyPr/>
        <a:lstStyle/>
        <a:p><a:r><a:rPr lang="en-US" sz="1800"/><a:t>Beta</a:t></a:r></a:p>
      </dsp:txBody>
    </dsp:sp>
  </dsp:spTree>
</dsp:drawing>`
}

// diagramDeck builds a one-slide package whose only content is a diagram
// graphicFrame. withPart and relType let a test break exactly one thing.
func diagramDeck(t *testing.T, withPart bool, relType string) []byte {
	t.Helper()
	pres := New()
	pres.AddSlide(newSlide()) // any slide, so slide1.xml gets written
	parts := zipParts(t, writeToBytes(t, pres))
	parts["ppt/slides/slide1.xml"] = []byte(bodyPrSlide(diagramFrameXML(100000, 200000, 2000000, 1000000)))

	if withPart {
		parts["ppt/diagrams/drawing1.xml"] = []byte(diagramDrawingXML())
		ct := string(parts["[Content_Types].xml"])
		ct = strings.Replace(ct, "</Types>",
			`  <Override PartName="/ppt/diagrams/drawing1.xml" ContentType="application/vnd.ms-office.drawingml.diagramDrawing+xml"/>`+"\n</Types>", 1)
		parts["[Content_Types].xml"] = []byte(ct)
	}

	relPath := "ppt/slides/_rels/slide1.xml.rels"
	rel := string(parts[relPath])
	if withPart {
		rel = strings.Replace(rel, "</Relationships>",
			`  <Relationship Id="rId95" Type="`+relType+`" Target="../diagrams/drawing1.xml"/>`+"\n</Relationships>", 1)
	}
	parts[relPath] = []byte(rel)
	return buildZip(t, parts)
}

func readDiagramDeck(t *testing.T, data []byte) *Presentation {
	t.Helper()
	pres, err := ReadFrom(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatalf("read deck: %v", err)
	}
	return pres
}

func renderSlide0(t *testing.T, pres *Presentation, width int) image.Image {
	t.Helper()
	path := t.TempDir() + "/slide1.png"
	opts := DefaultRenderOptions()
	opts.Width = width
	if err := pres.SaveSlideAsImage(0, path, opts); err != nil {
		t.Fatalf("render: %v", err)
	}
	f, err := os.Open(path)
	if err != nil {
		t.Fatalf("open render: %v", err)
	}
	defer f.Close()
	img, err := png.Decode(f)
	if err != nil {
		t.Fatalf("decode render: %v", err)
	}
	return img
}

// inkBox returns the bounding box and the pixel count of the pixels satisfying
// match, read on straight (non-premultiplied) 8-bit channels.
func inkBox(img image.Image, match func(r, g, b uint8) bool) (image.Rectangle, int) {
	box := image.Rectangle{}
	n := 0
	b := img.Bounds()
	for y := b.Min.Y; y < b.Max.Y; y++ {
		for x := b.Min.X; x < b.Max.X; x++ {
			r, g, bl, a := img.At(x, y).RGBA()
			if a>>8 < 8 {
				continue
			}
			if match(uint8(r>>8), uint8(g>>8), uint8(bl>>8)) {
				n++
				box = box.Union(image.Rect(x, y, x+1, y+1))
			}
		}
	}
	return box, n
}

var isRed = func(r, g, b uint8) bool { return r > 180 && g < 90 && b < 90 }
var isBlue = func(r, g, b uint8) bool { return b > 180 && r < 90 && g < 90 }

// TestSmartArtRendersFromCachedDrawing is the positive case: the drawing part
// is present, so the frame must come back as a group holding the diagram's own
// nodes, positioned by the frame and scaled by the identity map.
func TestSmartArtRendersFromCachedDrawing(t *testing.T) {
	pres := readDiagramDeck(t, diagramDeck(t, true, relDrawingType))

	if bad := pres.UnsupportedShapes(); len(bad) != 0 {
		t.Fatalf("a diagram with a cached drawing part must not be unsupported, got %d: %v",
			len(bad), bad[0].Label())
	}

	shapes := pres.GetAllSlides()[0].GetShapes()
	if len(shapes) != 1 {
		t.Fatalf("want 1 top-level shape, got %d", len(shapes))
	}
	g, ok := shapes[0].(*GroupShape)
	if !ok {
		t.Fatalf("want a *GroupShape for the diagram, got %T", shapes[0])
	}
	if g.GetOffsetX() != 100000 || g.GetOffsetY() != 200000 {
		t.Errorf("group offset = (%d,%d), want the frame's (100000,200000)", g.GetOffsetX(), g.GetOffsetY())
	}
	if g.GetWidth() != 2000000 || g.GetHeight() != 1000000 {
		t.Errorf("group extent = %dx%d, want the frame's 2000000x1000000", g.GetWidth(), g.GetHeight())
	}
	// The child coordinate space must be the identity map, or every node lands
	// somewhere else: renderer.go computes
	// x' = off + (x - childOffX) * width / childExtX.
	if g.childOffX != 0 || g.childOffY != 0 {
		t.Errorf("childOff = (%d,%d), want (0,0) so the child space is the identity map", g.childOffX, g.childOffY)
	}
	if g.childExtX != 2000000 || g.childExtY != 1000000 {
		t.Errorf("childExt = %dx%d, want the frame's 2000000x1000000", g.childExtX, g.childExtY)
	}
	if n := g.GetShapeCount(); n != 2 {
		t.Fatalf("want the drawing's 2 nodes, got %d", n)
	}

	first, ok := g.GetShapes()[0].(*AutoShape)
	if !ok {
		t.Fatalf("diagram node 1 is %T, want *AutoShape", g.GetShapes()[0])
	}
	if first.GetOffsetX() != 0 || first.GetWidth() != 1000000 {
		t.Errorf("node 1 x=%d w=%d, want the drawing part's 0 and 1000000", first.GetOffsetX(), first.GetWidth())
	}
	if got := first.GetFill(); got == nil || got.Type != FillSolid ||
		got.Color.GetRed() != 0xFF || got.Color.GetGreen() != 0 {
		t.Errorf("node 1 fill = %+v, want solid red", got)
	}
	if txt := pres.ExtractText(); !strings.Contains(txt, "Alpha") || !strings.Contains(txt, "Beta") {
		t.Errorf("deck text = %q, want both diagram node labels", txt)
	}
}

// TestSmartArtDiagramIsDrawn checks the pixels, because the reader model can be
// right while the rasteriser ignores the group. Both nodes are asserted, so a
// renderer that drew only the first one cannot pass.
//
// It pins the *position* as well as the presence of ink, because a count alone
// has no teeth here: with childExt left at 0 the renderer's group transform is
// skipped entirely, every node still draws, and the pixel counts are
// unchanged — only the frame's origin is missing. The nodes must therefore land
// where the frame offset plus their own child offsets put them, which is what
// makes the identity child space observable from the rendered image.
func TestSmartArtDiagramIsDrawn(t *testing.T) {
	const width = 800
	pres := readDiagramDeck(t, diagramDeck(t, true, relDrawingType))
	img := renderSlide0(t, pres, width)

	redBox, red := inkBox(img, isRed)
	if red < 500 {
		t.Fatalf("red node ink = %d px, want a filled node (>=500)", red)
	}
	blueBox, blue := inkBox(img, isBlue)
	if blue < 500 {
		t.Fatalf("blue node ink = %d px, want a filled node (>=500)", blue)
	}

	// The slide is 9144000 x 6858000 EMU. The frame sits at (100000,200000)
	// and measures 2000000 x 1000000; the red node is the frame's left half and
	// the blue node its right half, so at width 800 that is
	//   red  x 8.75 .. 96.2, y 17.5 .. 61.2
	//   blue x 96.2 .. 183.6
	// Two pixels of slack absorb the antialiased rounded corners.
	const tol = 3
	check := func(what string, box image.Rectangle, x0, y0, x1, y1 int) {
		t.Helper()
		if box.Empty() {
			t.Fatalf("%s drew no ink", what)
		}
		if box.Min.X < x0-tol || box.Min.Y < y0-tol || box.Max.X > x1+tol || box.Max.Y > y1+tol {
			t.Errorf("%s ink box = %v, want it inside (%d,%d)-(%d,%d) +/- %d",
				what, box, x0, y0, x1, y1, tol)
		}
	}
	check("red node", redBox, 8, 17, 97, 62)
	check("blue node", blueBox, 95, 17, 184, 62)

	// The two nodes sit side by side, so their boxes must not run into one
	// another: that would mean the child x offsets were ignored.
	if redBox.Max.X > blueBox.Min.X+tol {
		t.Errorf("red node ends at x=%d but blue starts at x=%d; the two diagram "+
			"nodes should sit side by side, not on top of each other", redBox.Max.X, blueBox.Min.X)
	}
}

// TestDiagramWithoutCachedDrawingFallsBack keeps the visible stand-in for a
// deck whose diagrams were never cached. Silently drawing an empty box would
// read like a correct render.
func TestDiagramWithoutCachedDrawingFallsBack(t *testing.T) {
	pres := readDiagramDeck(t, diagramDeck(t, false, relDrawingType))
	bad := pres.UnsupportedShapes()
	if len(bad) != 1 {
		t.Fatalf("want 1 unsupported shape when the drawing part is absent, got %d", len(bad))
	}
	if !strings.Contains(bad[0].Label(), "SmartArt") {
		t.Errorf("label = %q, want it to name SmartArt", bad[0].Label())
	}
}

// TestDiagramNeedsTheDrawingRelationship is the reverse knife. The part is
// there and readable, so a reader that resolved it by guessing a conventional
// file name would still find it; only matching the relationship type gets the
// right answer. It also proves the fallback above is about a missing part, not
// merely about a frame with nothing to show.
func TestDiagramNeedsTheDrawingRelationship(t *testing.T) {
	pres := readDiagramDeck(t, diagramDeck(t, true, relDrawingType+"-wrong"))
	bad := pres.UnsupportedShapes()
	if len(bad) != 1 {
		t.Fatalf("a diagramDrawing relationship of the wrong type must not be used, got %d unsupported", len(bad))
	}
	if !strings.Contains(bad[0].Label(), "SmartArt") {
		t.Errorf("label = %q, want it to name SmartArt", bad[0].Label())
	}
}
