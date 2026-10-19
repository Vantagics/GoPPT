package gopresentation

import (
	"strings"
	"testing"
)

// A SmartArt drawing node (dsp:sp) carries TWO transforms: the shape geometry
// in dsp:spPr/a:xfrm, and the *text* rectangle in dsp:txXfrm — inset from the
// shape by the preset's text-rect formulas (for roundRect, a*29289/100000 per
// side, customGeo drawing4 node 2: shape 2279381x2240458 EMU against a text
// rect 2148139x2109216). The slide scanner matches "off"/"ext" on local name
// alone, so before the txXfrm gate the text rect silently overwrote the shape
// geometry: every diagram node shrank by the inset and drifted inward, which
// measured 0.94x scale plus a ~12px shift against PowerPoint COM exports on
// customGeo slides 10/11/13.
//
// The text rect is symmetric about the shape's centre, so ignoring it leaves
// centred node text where PowerPoint puts it; the bodyPr insets still govern
// the text box inside the shape.

// txxfrmDrawingXML is the r65 two-node diagram with a dsp:txXfrm added to each
// node — an inset rect the way PowerPoint writes it, deliberately far enough
// from the shape rect that any leak into the geometry is unmistakable.
func txxfrmDrawingXML() string {
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
      <dsp:txXfrm><a:off x="29289" y="14644"/><a:ext cx="941422" cy="470711"/></dsp:txXfrm>
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
      <dsp:txXfrm><a:off x="1029289" y="14644"/><a:ext cx="941422" cy="470711"/></dsp:txXfrm>
    </dsp:sp>
  </dsp:spTree>
</dsp:drawing>`
}

// txxfrmDeck is diagramDeck with the txXfrm-bearing drawing part.
func txxfrmDeck(t *testing.T) []byte {
	t.Helper()
	pres := New()
	pres.AddSlide(newSlide())
	parts := zipParts(t, writeToBytes(t, pres))
	parts["ppt/slides/slide1.xml"] = []byte(bodyPrSlide(diagramFrameXML(100000, 200000, 2000000, 1000000)))
	parts["ppt/diagrams/drawing1.xml"] = []byte(txxfrmDrawingXML())
	ct := string(parts["[Content_Types].xml"])
	ct = strings.Replace(ct, "</Types>",
		`  <Override PartName="/ppt/diagrams/drawing1.xml" ContentType="application/vnd.ms-office.drawingml.diagramDrawing+xml"/>`+"\n</Types>", 1)
	parts["[Content_Types].xml"] = []byte(ct)
	relPath := "ppt/slides/_rels/slide1.xml.rels"
	rel := string(parts[relPath])
	rel = strings.Replace(rel, "</Relationships>",
		`  <Relationship Id="rId95" Type="`+relDrawingType+`" Target="../diagrams/drawing1.xml"/>`+"\n</Relationships>", 1)
	parts[relPath] = []byte(rel)
	return buildZip(t, parts)
}

// TestSmartArtTxXfrmDoesNotOverrideShapeGeometry pins the reader: a node's
// geometry comes from dsp:spPr/a:xfrm, never from the dsp:txXfrm text rect
// that follows the text body.
func TestSmartArtTxXfrmDoesNotOverrideShapeGeometry(t *testing.T) {
	pres := readDiagramDeck(t, txxfrmDeck(t))
	shapes := pres.GetAllSlides()[0].GetShapes()
	if len(shapes) != 1 {
		t.Fatalf("want 1 top-level shape, got %d", len(shapes))
	}
	g, ok := shapes[0].(*GroupShape)
	if !ok {
		t.Fatalf("want a *GroupShape for the diagram, got %T", shapes[0])
	}
	nodes := g.GetShapes()
	if len(nodes) != 2 {
		t.Fatalf("want the drawing's 2 nodes, got %d", len(nodes))
	}
	for i, want := range []struct{ x, y, w, h int64 }{
		{0, 0, 1000000, 500000},
		{1000000, 0, 1000000, 500000},
	} {
		node, ok := nodes[i].(*AutoShape)
		if !ok {
			t.Fatalf("node %d is %T, want *AutoShape", i, nodes[i])
		}
		if node.GetOffsetX() != want.x || node.GetOffsetY() != want.y ||
			node.GetWidth() != want.w || node.GetHeight() != want.h {
			t.Errorf("node %d geometry = (%d,%d) %dx%d, want the spPr xfrm (%d,%d) %dx%d",
				i, node.GetOffsetX(), node.GetOffsetY(), node.GetWidth(), node.GetHeight(),
				want.x, want.y, want.w, want.h)
		}
	}
}
