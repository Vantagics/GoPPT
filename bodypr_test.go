package gopresentation

import (
	"bytes"
	"strings"
	"testing"
)

// The <a:bodyPr> tests.
//
// Three shape kinds carry a text body, and the reader parses the same
// attributes and child elements for all three while the renderer honours them
// for all three. Each writer used to build the element its own way, and the two
// bespoke emitters (AutoShape and PlaceholderShape) wrote a bare <a:bodyPr/> —
// so a shape read from a deck and written back lost its text insets, its text
// direction and its auto-fit mode, and an AutoShape lost the anchor PowerPoint
// then re-applied as "top".
//
// Every assertion below is structural, on the emitted XML text, because the
// round trip cannot see the problem: the reader matches on the local name and
// does not care whether an attribute is there, so a dropped attribute simply
// reads back as the default.

// bodyPrSlide wraps spTree content in the rest of a slide.
func bodyPrSlide(shapes string) string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="` + nsDrawingML + `" xmlns:r="` + nsOfficeDocRels + `" xmlns:p="` + nsPresentationML + `">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr/>
` + shapes + `
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr>
    <a:masterClrMapping/>
  </p:clrMapOvr>
</p:sld>`
}

// bodyPrTextBox is a text box whose <a:bodyPr> is whatever the test is about.
func bodyPrTextBox(bodyPr string) string {
	return `<p:sp>
  <p:nvSpPr><p:cNvPr id="2" name="TextBox"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="500000" y="500000"/><a:ext cx="3000000" cy="2000000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
  <p:txBody>
    ` + bodyPr + `
    <a:lstStyle/>
    <a:p><a:r><a:rPr lang="en-US" dirty="0" sz="1800"/><a:t>SIDE</a:t></a:r></a:p>
  </p:txBody>
</p:sp>`
}

// bodyPrAutoShape is an AutoShape: its geometry is not "rect", which is what
// makes the reader convert the temporary text body into one.
func bodyPrAutoShape(bodyPr, text string) string {
	return `<p:sp>
  <p:nvSpPr><p:cNvPr id="3" name="Rounded"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="500000" y="500000"/><a:ext cx="3000000" cy="2000000"/></a:xfrm>
    <a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>
  </p:spPr>
  <p:txBody>
    ` + bodyPr + `
    <a:lstStyle/>
    <a:p><a:r><a:rPr lang="en-US" dirty="0" sz="1800"/><a:t>` + text + `</a:t></a:r></a:p>
  </p:txBody>
</p:sp>`
}

// bodyPrPlaceholder is a slide-level placeholder.
func bodyPrPlaceholder(bodyPr string) string {
	return `<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="4" name="Body Placeholder"/>
    <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
    <p:nvPr><p:ph type="body" idx="1"/></p:nvPr>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="500000" y="500000"/><a:ext cx="3000000" cy="2000000"/></a:xfrm>
  </p:spPr>
  <p:txBody>
    ` + bodyPr + `
    <a:lstStyle/>
    <a:p><a:r><a:rPr lang="en-US" dirty="0" sz="1800"/><a:t>BODY</a:t></a:r></a:p>
  </p:txBody>
</p:sp>`
}

// bodyPrRoundTrip reads a one-slide package whose slide is the given shape, then
// writes it back, and returns both the presentation and the slide part the
// writer produced. The written part is what a structural assertion looks at.
func bodyPrRoundTrip(t *testing.T, shape string) (*Presentation, string) {
	t.Helper()
	parts := zipParts(t, writeToBytes(t, New()))
	parts["ppt/slides/slide1.xml"] = []byte(bodyPrSlide(shape))
	pkg := buildZip(t, parts)

	pres, err := (&PPTXReader{}).ReadFromReader(bytes.NewReader(pkg), int64(len(pkg)))
	if err != nil {
		t.Fatalf("read the package: %v", err)
	}
	if n := len(pres.GetAllSlides()); n != 1 {
		t.Fatalf("read %d slides, want 1", n)
	}
	out := zipParts(t, writeToBytes(t, pres))
	slide, ok := out["ppt/slides/slide1.xml"]
	if !ok {
		t.Fatal("the written package has no ppt/slides/slide1.xml")
	}
	return pres, string(slide)
}

// writtenBodyPr returns the first <a:bodyPr> element of a written slide part, so
// an assertion is about the element rather than about the whole document.
func writtenBodyPr(t *testing.T, slide string) string {
	t.Helper()
	i := strings.Index(slide, "<a:bodyPr")
	if i < 0 {
		t.Fatal("the written slide has no <a:bodyPr>")
	}
	rest := slide[i:]
	if j := strings.Index(rest, "</a:bodyPr>"); j >= 0 {
		return rest[:j+len("</a:bodyPr>")]
	}
	j := strings.Index(rest, "/>")
	return rest[:j+2]
}

// firstShape returns the first shape of the first slide.
func firstShape(t *testing.T, pres *Presentation) Shape {
	t.Helper()
	shapes := pres.GetAllSlides()[0].GetShapes()
	if len(shapes) == 0 {
		t.Fatal("the slide has no shapes")
	}
	return shapes[0]
}

// TestTextBoxBodyPrKeepsItsTextInsets covers the insets half of the text body.
// The reader parses lIns/tIns/rIns/bIns and the renderer pads the text with
// them, but the writer emitted none of the four, so saving a deck reset every
// custom inset back to PowerPoint's default padding.
func TestTextBoxBodyPrKeepsItsTextInsets(t *testing.T) {
	pres, slide := bodyPrRoundTrip(t, bodyPrTextBox(
		`<a:bodyPr lIns="0" tIns="12700" rIns="25400" bIns="38100"/>`))

	rt, ok := firstShape(t, pres).(*RichTextShape)
	if !ok {
		t.Fatalf("shape is %T, want *RichTextShape", firstShape(t, pres))
	}
	if !rt.insetsSet {
		t.Fatal("the reader dropped the text insets before the writer could see them")
	}
	if rt.insetLeft != 0 || rt.insetTop != 12700 || rt.insetRight != 25400 || rt.insetBottom != 38100 {
		t.Errorf("read insets = %d/%d/%d/%d, want 0/12700/25400/38100",
			rt.insetLeft, rt.insetTop, rt.insetRight, rt.insetBottom)
	}

	got := writtenBodyPr(t, slide)
	for _, want := range []string{`lIns="0"`, `tIns="12700"`, `rIns="25400"`, `bIns="38100"`} {
		if !strings.Contains(got, want) {
			t.Errorf("written %s has no %s: the text insets were read and drawn but not saved", got, want)
		}
	}
}

// TestBodyPrLeavesOutInsetsItWasNeverGiven guards the other direction: a shape
// whose <a:bodyPr> states no insets must not gain four of them. A placeholder
// inherits its padding from the layout, and pinning PowerPoint's defaults here
// would break that inheritance for good.
func TestBodyPrLeavesOutInsetsItWasNeverGiven(t *testing.T) {
	_, slide := bodyPrRoundTrip(t, bodyPrTextBox(`<a:bodyPr/>`))
	got := writtenBodyPr(t, slide)
	for _, unwanted := range []string{"lIns=", "tIns=", "rIns=", "bIns="} {
		if strings.Contains(got, unwanted) {
			t.Errorf("written %s states %s, but the shape never carried insets", got, unwanted)
		}
	}
}

// TestTextBoxBodyPrKeepsItsTextDirection covers vert. Vertical text is read into
// the model and the renderer rotates the text for it; the writer dropped it, so
// a saved deck came back horizontal.
func TestTextBoxBodyPrKeepsItsTextDirection(t *testing.T) {
	_, slide := bodyPrRoundTrip(t, bodyPrTextBox(`<a:bodyPr vert="vert270"/>`))
	if got := writtenBodyPr(t, slide); !strings.Contains(got, `vert="vert270"`) {
		t.Errorf("written %s has no vert=\"vert270\": the text direction was read and drawn but not saved", got)
	}

	// "horz" is the schema default, so it is left out rather than stated. The
	// reader stores whatever it read, and writing it back would be harmless but
	// pointless; what matters is that the element stays well formed.
	_, slide = bodyPrRoundTrip(t, bodyPrTextBox(`<a:bodyPr vert="horz"/>`))
	if got := writtenBodyPr(t, slide); strings.Contains(got, "vert=") {
		t.Errorf("written %s states vert for the default direction", got)
	}
}

// TestBodyPrRecordsTheAutoFitMode covers the auto-fit choice. The shrunken-text
// and resize-to-fit modes are read, and the renderer acts on them; the writer
// emitted <a:normAutofit> only when a font scale happened to be recorded, and
// <a:spAutoFit/> never.
func TestBodyPrRecordsTheAutoFitMode(t *testing.T) {
	cases := []struct {
		name     string
		bodyPr   string
		want     []string
		unwanted []string
	}{
		{
			name:   "resize shape to fit text",
			bodyPr: `<a:bodyPr><a:spAutoFit/></a:bodyPr>`,
			want:   []string{"<a:spAutoFit/>"},
		},
		{
			name:   "shrink text on overflow, already scaled",
			bodyPr: `<a:bodyPr><a:normAutofit fontScale="62500"/></a:bodyPr>`,
			want:   []string{`<a:normAutofit fontScale="62500"/>`},
		},
		{
			// "Shrink text on overflow" with nothing to shrink yet: PowerPoint
			// writes the element without the attribute, and the mode still has
			// to survive.
			name:   "shrink text on overflow, not yet scaled",
			bodyPr: `<a:bodyPr><a:normAutofit/></a:bodyPr>`,
			want:   []string{"<a:normAutofit/>"},
		},
		{
			// The zero value of AutoFitType is "the model was not told", which
			// is not the same as "do not auto-fit": <a:noAutofit/> would claim a
			// decision nobody made.
			name:     "no mode stated",
			bodyPr:   `<a:bodyPr/>`,
			unwanted: []string{"<a:spAutoFit/>", "<a:normAutofit", "<a:noAutofit/>"},
		},
	}

	for _, c := range cases {
		t.Run(c.name, func(t *testing.T) {
			_, slide := bodyPrRoundTrip(t, bodyPrTextBox(c.bodyPr))
			got := writtenBodyPr(t, slide)
			for _, want := range c.want {
				if !strings.Contains(got, want) {
					t.Errorf("written %s has no %s", got, want)
				}
			}
			for _, unwanted := range c.unwanted {
				if strings.Contains(got, unwanted) {
					t.Errorf("written %s states %s, which the shape was never told", got, unwanted)
				}
			}
		})
	}
}

// TestAutoShapeBodyPrKeepsItsAnchorWrapAndColumns covers the AutoShape emitter,
// which wrote a bare <a:bodyPr/>. The anchor is the damaging one: PowerPoint's
// default is "top", so a shape whose text sat in the middle jumped to the top of
// its box the moment the deck was saved.
func TestAutoShapeBodyPrKeepsItsAnchorWrapAndColumns(t *testing.T) {
	pres, slide := bodyPrRoundTrip(t, bodyPrAutoShape(
		`<a:bodyPr anchor="ctr" wrap="none" numCol="2"/>`, "INSIDE"))

	sh, ok := firstShape(t, pres).(*AutoShape)
	if !ok {
		t.Fatalf("shape is %T, want *AutoShape", firstShape(t, pres))
	}
	if sh.textAnchor != TextAnchorMiddle {
		t.Errorf("read anchor = %q, want %q", sh.textAnchor, TextAnchorMiddle)
	}
	if sh.wordWrap {
		t.Error("read wrap=\"none\" as wrapping text")
	}
	if sh.columns != 2 {
		t.Errorf("read numCol as %d, want 2", sh.columns)
	}

	got := writtenBodyPr(t, slide)
	for _, want := range []string{`anchor="ctr"`, `wrap="none"`, `numCol="2"`} {
		if !strings.Contains(got, want) {
			t.Errorf("written %s has no %s", got, want)
		}
	}
}

// TestAutoShapeBodyPrKeepsItsAutoFit covers the third body property the
// AutoShape conversion dropped. The mode is parsed into the temporary text body
// and the renderer sizes the text by it, so a shape that lost it was re-laid out
// with the wrong rule the next time the deck was opened.
func TestAutoShapeBodyPrKeepsItsAutoFit(t *testing.T) {
	pres, slide := bodyPrRoundTrip(t, bodyPrAutoShape(
		`<a:bodyPr><a:spAutoFit/></a:bodyPr>`, "FIT"))

	sh, ok := firstShape(t, pres).(*AutoShape)
	if !ok {
		t.Fatalf("shape is %T, want *AutoShape", firstShape(t, pres))
	}
	if sh.autoFit != AutoFitShape {
		t.Errorf("read autoFit = %v, want %v: the AutoShape conversion drops the mode "+
			"the bodyPr branch read into the temporary text body", sh.autoFit, AutoFitShape)
	}
	if got := writtenBodyPr(t, slide); !strings.Contains(got, "<a:spAutoFit/>") {
		t.Errorf("written %s has no <a:spAutoFit/>", got)
	}
}

// TestAutoShapeWrapNoneIsNotWrappedWhenDrawn is the rendering half of the same
// defect: the shape's text was drawn with wordWrap hard-coded true, so a deck
// that said wrap="none" still got wrapped text in the preview. The number of
// inked row bands is one per drawn line.
func TestAutoShapeWrapNoneIsNotWrappedWhenDrawn(t *testing.T) {
	fc := NewFontCache()
	requireAnyFont(t, fc)

	const text = "ALPHA BRAVO CHARLIE DELTA ECHO FOXTROT GOLF HOTEL"

	textLines := func(wrap string) int {
		pres, _ := bodyPrRoundTrip(t, bodyPrAutoShape(`<a:bodyPr wrap="`+wrap+`"/>`, text))
		opts := goldenOptions(fc)
		opts.Width = 800
		img, err := pres.SlideToImage(0, opts)
		if err != nil {
			t.Fatalf("render with wrap=%q: %v", wrap, err)
		}
		return len(inkRowBands(img))
	}

	if n := textLines("square"); n < 2 {
		t.Fatalf("wrap=\"square\" drew %d text line(s), want at least 2; the fixture no longer wraps", n)
	}
	if n := textLines("none"); n != 1 {
		t.Errorf("wrap=\"none\" drew %d text lines, want 1: the shape's <a:bodyPr wrap> "+
			"is being ignored when the text is drawn", n)
	}
}

// TestPlaceholderBodyPrKeepsItsAnchor covers the third carrier. The reader kept
// the anchor in a local that only the RichTextShape and AutoShape branches
// copied out, so a placeholder's anchor was lost on the way in as well as on the
// way out.
func TestPlaceholderBodyPrKeepsItsAnchor(t *testing.T) {
	pres, slide := bodyPrRoundTrip(t, bodyPrPlaceholder(`<a:bodyPr anchor="ctr"/>`))

	ph, ok := firstShape(t, pres).(*PlaceholderShape)
	if !ok {
		t.Fatalf("shape is %T, want *PlaceholderShape", firstShape(t, pres))
	}
	if ph.textAnchor != TextAnchorMiddle {
		t.Errorf("read anchor = %q, want %q: the bodyPr branch stores the anchor on the "+
			"shape, not only in the local the text-box branch reads", ph.textAnchor, TextAnchorMiddle)
	}
	if got := writtenBodyPr(t, slide); !strings.Contains(got, `anchor="ctr"`) {
		t.Errorf("written %s has no anchor=\"ctr\"", got)
	}
}

// TestAutoShapeKeepsItsRunFormatting covers the other half of the AutoShape
// emitter. The reader keeps both the runs and a flattened string, and the
// renderer draws the runs; the writer emitted the flattened string as one
// unformatted run, so every bold, size and colour inside a shape was gone after
// a save.
func TestAutoShapeKeepsItsRunFormatting(t *testing.T) {
	shape := `<p:sp>
  <p:nvSpPr><p:cNvPr id="3" name="Rounded"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="500000" y="500000"/><a:ext cx="3000000" cy="2000000"/></a:xfrm>
    <a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>
  </p:spPr>
  <p:txBody>
    <a:bodyPr/>
    <a:lstStyle/>
    <a:p>
      <a:pPr algn="ctr"/>
      <a:r>
        <a:rPr lang="en-US" dirty="0" sz="3200" b="1"><a:solidFill><a:srgbClr val="CC0000"/></a:solidFill></a:rPr>
        <a:t>BOLD</a:t>
      </a:r>
    </a:p>
  </p:txBody>
</p:sp>`

	pres, slide := bodyPrRoundTrip(t, shape)
	sh, ok := firstShape(t, pres).(*AutoShape)
	if !ok {
		t.Fatalf("shape is %T, want *AutoShape", firstShape(t, pres))
	}
	if len(sh.paragraphs) == 0 {
		t.Fatal("the reader kept no paragraphs for the shape")
	}
	if sh.text != "BOLD" {
		t.Errorf("flattened text = %q, want %q", sh.text, "BOLD")
	}

	for _, want := range []string{`b="1"`, `sz="3200"`, `CC0000`, `algn="ctr"`} {
		if !strings.Contains(slide, want) {
			t.Errorf("the written slide has no %s: a shape's text was saved as one "+
				"unformatted run even though the runs survived the read", want)
		}
	}
}
