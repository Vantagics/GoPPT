package gopresentation

import (
	"archive/zip"
	"bytes"
	"image"
	"image/color"
	"strings"
	"testing"
)

// Shape attributes that every layer but the writer carried.
//
// The reader and the renderer both handle a picture's crop, a picture's
// opacity, a preset geometry's adjustment values and a custom geometry path;
// the writer emitted none of them. Nothing in this library could see that,
// because a round trip through our own reader agrees with itself either way and
// the renderer draws from the model — where the values are present and correct.
// PowerPoint reads the part, so every one of those attributes came back as its
// default: an uncropped picture, an opaque picture, a rounded rectangle with
// the preset's corner radius, and a freeform shape that had become a rectangle.
//
// Two carriers of the same picture crop were also missed on the way *in*: a
// background picture, on a slide and on a layout, was read but its crop and
// opacity were not, because they are siblings of <a:blip> rather than children
// of a <p:pic>.
//
// Each test below asserts on the emitted XML text — with the a: prefix, since
// this library's own reader matches on the local name and would happily accept
// a bare <srcRect> that PowerPoint rejects — plus a render or a round trip
// where the symptom is visible.

// slidePartOfShapeRound returns the slide part of a one-slide presentation.
func slidePartOfShapeRound(t *testing.T, p *Presentation) string {
	t.Helper()
	part, ok := zipParts(t, writeToBytes(t, p))["ppt/slides/slide1.xml"]
	if !ok {
		t.Fatal("ppt/slides/slide1.xml was not written")
	}
	return string(part)
}

// --- picture crop and opacity ---

// TestPictureCropAndOpacityReachTheFile covers the writer. Both attributes were
// read from the file and applied by the renderer, so only the preview was ever
// right; saving dropped them.
func TestPictureCropAndOpacityReachTheFile(t *testing.T) {
	p := New()
	d := p.GetActiveSlide().CreateDrawingShape()
	d.SetImageData(smallPNG(t), "image/png")
	d.SetOffsetX(1000000).SetOffsetY(1000000)
	d.SetWidth(2000000).SetHeight(2000000)
	d.SetCrop(10000, 20000, 30000, 40000)
	d.SetAlphaValue(50000)

	slide := slidePartOfShapeRound(t, p)
	pic := blockIn(t, slide, "<p:pic>", "</p:pic>")

	if !strings.Contains(pic, `<a:srcRect l="10000" t="20000" r="30000" b="40000"/>`) {
		t.Errorf("the crop is not in the picture; the reader and the renderer both carry it, "+
			"so the writer has to. Got:\n%s", pic)
	}
	if !strings.Contains(pic, `<a:alphaModFix amt="50000"/>`) {
		t.Errorf("the opacity is not in the picture. Got:\n%s", pic)
	}
	// CT_BlipFillProperties fixes the order: blip, srcRect, stretch; the
	// opacity is a child of the blip.
	assertOrdered(t, "picture blipFill", blockIn(t, pic, "<p:blipFill>", "</p:blipFill>"),
		"<a:blip", "<a:alphaModFix", "</a:blip>", "<a:srcRect", "<a:stretch")
}

// TestPictureWithoutCropOrOpacityWritesNeither is the control: an uncropped,
// opaque picture must not gain either element. Writing an all-zero <a:srcRect>
// would be schemavalid but is not what PowerPoint writes, and an
// <a:alphaModFix amt="100000"/> would claim an opacity the model does not set.
func TestPictureWithoutCropOrOpacityWritesNeither(t *testing.T) {
	p := New()
	d := p.GetActiveSlide().CreateDrawingShape()
	d.SetImageData(smallPNG(t), "image/png")
	d.SetWidth(2000000).SetHeight(2000000)

	pic := blockIn(t, slidePartOfShapeRound(t, p), "<p:pic>", "</p:pic>")

	if strings.Contains(pic, "<a:srcRect") {
		t.Errorf("an uncropped picture gained a crop. Got:\n%s", pic)
	}
	if strings.Contains(pic, "<a:alphaModFix") {
		t.Errorf("an opaque picture gained an alphaModFix. Got:\n%s", pic)
	}
}

// TestPictureCropAndOpacitySurviveRoundTrip closes the loop: what the writer
// emits has to be what the reader reads back, at the same values.
func TestPictureCropAndOpacitySurviveRoundTrip(t *testing.T) {
	p := New()
	d := p.GetActiveSlide().CreateDrawingShape()
	d.SetImageData(smallPNG(t), "image/png")
	d.SetWidth(2000000).SetHeight(2000000)
	d.SetCrop(12500, 0, 0, 25000)
	d.SetAlphaValue(40000)

	back, ok := firstShapeOfType[*DrawingShape](roundTrip(t, p))
	if !ok {
		t.Fatal("the picture did not survive the round trip")
	}
	if got := back.GetCropLeft(); got != 12500 {
		t.Errorf("crop left = %d, want 12500", got)
	}
	if got := back.GetCropBottom(); got != 25000 {
		t.Errorf("crop bottom = %d, want 25000", got)
	}
	if got := back.GetAlphaValue(); got != 40000 {
		t.Errorf("alpha = %d, want 40000", got)
	}
}

// --- preset geometry adjustment values ---

// TestAutoShapeAdjustmentsReachTheFile covers the writer's hardcoded
// <a:avLst/>. A rounded rectangle's radius, an arrow's proportions and a
// chevron's point all live there, so a save reset every one of them to the
// preset default.
func TestAutoShapeAdjustmentsReachTheFile(t *testing.T) {
	p := New()
	a := p.GetActiveSlide().CreateAutoShape()
	a.SetAutoShapeType(AutoShapeRoundedRect)
	a.SetOffsetX(500000).SetOffsetY(500000)
	a.SetWidth(3000000).SetHeight(1500000)
	// Inserted out of order on purpose: the writer sorts the names, because a
	// map's iteration order is random and a file that differs between two saves
	// of the same model is a reproducibility bug.
	a.SetAdjustValue("adj2", 33000)
	a.SetAdjustValue("adj1", 25000)

	sp := blockIn(t, slidePartOfShapeRound(t, p), "<p:sp>", "</p:sp>")
	geom := blockIn(t, sp, "<a:prstGeom", "</a:prstGeom>")

	if !strings.Contains(geom, `<a:gd name="adj1" fmla="val 25000"/>`) ||
		!strings.Contains(geom, `<a:gd name="adj2" fmla="val 33000"/>`) {
		t.Errorf("the adjustment values are not in the shape. Got:\n%s", geom)
	}
	assertOrdered(t, "adjustment list", geom,
		`<a:prstGeom prst="roundRect">`, "<a:avLst>", "adj1", "adj2", "</a:avLst>", "</a:prstGeom>")
}

// TestConnectorAdjustmentsReachTheFile is the same defect at the connector
// call site: a bent connector's knee is an adjustment value, and the writer
// wrote an empty <a:avLst/> for it too.
func TestConnectorAdjustmentsReachTheFile(t *testing.T) {
	p := New()
	l := p.GetActiveSlide().CreateLineShape()
	l.SetConnectorType("bentConnector3")
	l.SetAdjustValue("adj1", 50000)
	l.SetAdjustValue("adj2", 25000)

	cxn := blockIn(t, slidePartOfShapeRound(t, p), "<p:cxnSp>", "</p:cxnSp>")

	if !strings.Contains(cxn, `<a:gd name="adj1" fmla="val 50000"/>`) {
		t.Errorf("the connector's adjustment values are not in the file. Got:\n%s", cxn)
	}
	assertOrdered(t, "connector geometry", cxn,
		`<a:prstGeom prst="bentConnector3">`, "<a:avLst>", "adj1", "adj2", "</a:prstGeom>")
}

// TestAdjustmentsSurviveRoundTrip pins the values the reader gives back, so the
// emitted list is not merely present but parsed as the same map.
func TestAdjustmentsSurviveRoundTrip(t *testing.T) {
	p := New()
	a := p.GetActiveSlide().CreateAutoShape()
	a.SetAutoShapeType(AutoShapeArrowRight)
	a.SetWidth(3000000).SetHeight(1500000)
	a.SetAdjustValue("adj1", 60000)
	a.SetAdjustValue("adj2", 40000)

	back, ok := firstShapeOfType[*AutoShape](roundTrip(t, p))
	if !ok {
		t.Fatal("the auto shape did not survive the round trip")
	}
	got := back.GetAdjustValues()
	if got["adj1"] != 60000 || got["adj2"] != 40000 {
		t.Errorf("adjustment values after a round trip = %v, want map[adj1:60000 adj2:40000]", got)
	}
}

// --- custom geometry ---

// freeformPath is a path using every command the reader understands, so one
// round trip exercises the whole serialiser.
func freeformPath() *CustomGeomPath {
	return &CustomGeomPath{
		Width:  1000,
		Height: 1000,
		Commands: []PathCommand{
			{Type: "moveTo", Pts: []PathPoint{{X: 0, Y: 0}}},
			{Type: "lnTo", Pts: []PathPoint{{X: 1000, Y: 0}}},
			{Type: "cubicBezTo", Pts: []PathPoint{{X: 800, Y: 200}, {X: 600, Y: 800}, {X: 0, Y: 1000}}},
			{Type: "quadBezTo", Pts: []PathPoint{{X: 100, Y: 900}, {X: 0, Y: 800}}},
			{Type: "arcTo", WR: 50, HR: 40, StAng: 0, SwAng: 5400000},
			{Type: "close"},
		},
	}
}

// TestFreeformShapeIsWrittenAsCustomGeometry covers the omission with the
// largest visual effect: a shape with a custGeom path was written with
// <a:prstGeom prst="rect">, so every freeform drawing came back as a rectangle.
func TestFreeformShapeIsWrittenAsCustomGeometry(t *testing.T) {
	p := New()
	s := p.GetActiveSlide().CreateRichTextShape()
	s.SetWidth(3000000).SetHeight(2000000)
	s.SetCustomPath(freeformPath())

	sp := blockIn(t, slidePartOfShapeRound(t, p), "<p:sp>", "</p:sp>")

	if !strings.Contains(sp, "<a:custGeom>") {
		t.Fatalf("the custom path is not in the shape; it was drawn by the renderer and read "+
			"by the reader, so a save has to keep it. Got:\n%s", sp)
	}
	if strings.Contains(sp, "<a:prstGeom") {
		t.Errorf("a shape with custom geometry also carries preset geometry, which is not "+
			"allowed: the two are alternatives. Got:\n%s", sp)
	}
	// CT_CustomGeometry2D fixes the child order, and pathLst comes last.
	assertOrdered(t, "custGeom", blockIn(t, sp, "<a:custGeom>", "</a:custGeom>"),
		"<a:avLst/>", "<a:gdLst/>", "<a:ahLst/>", "<a:cxnLst/>", "<a:rect ", "<a:pathLst>", "<a:path ", "</a:pathLst>")
}

// TestFreeformPathCommandsReachTheFile checks the commands themselves, since a
// path whose points are dropped is as wrong as one that is missing.
func TestFreeformPathCommandsReachTheFile(t *testing.T) {
	p := New()
	s := p.GetActiveSlide().CreateRichTextShape()
	s.SetWidth(3000000).SetHeight(2000000)
	s.SetCustomPath(freeformPath())

	path := blockIn(t, slidePartOfShapeRound(t, p), `<a:path w=`, "</a:path>")

	for _, want := range []string{
		`<a:moveTo><a:pt x="0" y="0"/></a:moveTo>`,
		`<a:lnTo><a:pt x="1000" y="0"/></a:lnTo>`,
		`<a:cubicBezTo><a:pt x="800" y="200"/><a:pt x="600" y="800"/><a:pt x="0" y="1000"/></a:cubicBezTo>`,
		`<a:quadBezTo><a:pt x="100" y="900"/><a:pt x="0" y="800"/></a:quadBezTo>`,
		`<a:arcTo wR="50" hR="40" stAng="0" swAng="5400000"/>`,
		`<a:close/>`,
	} {
		if !strings.Contains(path, want) {
			t.Errorf("path command %s is missing. Got:\n%s", want, path)
		}
	}
	if !strings.HasPrefix(path, `<a:path w="1000" h="1000">`) {
		t.Errorf("the path lost its coordinate space; got:\n%s", path)
	}
}

// TestConnectorFreeformPathIsWritten is the second custGeom call site: a
// freeform connector is a <p:cxnSp>, and its path was dropped the same way.
func TestConnectorFreeformPathIsWritten(t *testing.T) {
	p := New()
	l := p.GetActiveSlide().CreateLineShape()
	l.SetWidth(3000000).SetHeight(2000000)
	l.SetCustomPath(freeformPath())

	cxn := blockIn(t, slidePartOfShapeRound(t, p), "<p:cxnSp>", "</p:cxnSp>")

	if !strings.Contains(cxn, "<a:custGeom>") {
		t.Errorf("the connector's custom path is not in the file. Got:\n%s", cxn)
	}
	if strings.Contains(cxn, "<a:prstGeom") {
		t.Errorf("the connector wrote both custGeom and prstGeom. Got:\n%s", cxn)
	}
}

// TestFreeformPathSurvivesRoundTrip pins the parsed path against the one that
// was written. The reader appends points to the last command it saw, so a
// truncated command would come back as a different shape rather than as an
// error.
func TestFreeformPathSurvivesRoundTrip(t *testing.T) {
	p := New()
	s := p.GetActiveSlide().CreateRichTextShape()
	s.SetWidth(3000000).SetHeight(2000000)
	s.SetCustomPath(freeformPath())

	back, ok := firstShapeOfType[*RichTextShape](roundTrip(t, p))
	if !ok {
		t.Fatal("the freeform shape did not survive the round trip")
	}
	cp := back.GetCustomPath()
	if cp == nil {
		t.Fatal("the custom path was lost on the way back in")
	}
	if cp.Width != 1000 || cp.Height != 1000 {
		t.Errorf("path space = %dx%d, want 1000x1000", cp.Width, cp.Height)
	}
	want := freeformPath().Commands
	if len(cp.Commands) != len(want) {
		t.Fatalf("path has %d commands, want %d: %+v", len(cp.Commands), len(want), cp.Commands)
	}
	for i, w := range want {
		got := cp.Commands[i]
		if got.Type != w.Type || len(got.Pts) != len(w.Pts) {
			t.Errorf("command %d = %s with %d points, want %s with %d",
				i, got.Type, len(got.Pts), w.Type, len(w.Pts))
			continue
		}
		for j, p := range w.Pts {
			if got.Pts[j] != p {
				t.Errorf("command %d point %d = %+v, want %+v", i, j, got.Pts[j], p)
			}
		}
		if got.WR != w.WR || got.HR != w.HR || got.StAng != w.StAng || got.SwAng != w.SwAng {
			t.Errorf("command %d arc parameters = %d/%d/%d/%d, want %d/%d/%d/%d",
				i, got.WR, got.HR, got.StAng, got.SwAng, w.WR, w.HR, w.StAng, w.SwAng)
		}
	}
}

// TestCustomPathArrowEndsReachTheLine covers the arrow ends of a freeform line.
// They belong to the shape's <a:ln>, and the renderer draws them for a custom
// path, so dropping them turned a curved arrow into a plain curve.
func TestCustomPathArrowEndsReachTheLine(t *testing.T) {
	p := New()
	s := p.GetActiveSlide().CreateRichTextShape()
	s.SetWidth(3000000).SetHeight(2000000)
	s.SetCustomPath(freeformPath())
	s.GetBorder().SetSolidFill(ColorBlack).SetWidth(2)
	s.SetHeadEnd(&LineEnd{Type: ArrowTriangle, Width: "med", Length: "med"})
	s.SetTailEnd(&LineEnd{Type: ArrowArrow, Width: "lg", Length: "lg"})

	ln := blockIn(t, slidePartOfShapeRound(t, p), "<a:ln ", "</a:ln>")

	if !strings.Contains(ln, `<a:headEnd type="triangle"`) {
		t.Errorf("the head end is not on the shape's line. Got:\n%s", ln)
	}
	if !strings.Contains(ln, `<a:tailEnd type="arrow"`) {
		t.Errorf("the tail end is not on the shape's line. Got:\n%s", ln)
	}
	// CT_LineProperties: the fill, then the dash, then the ends.
	assertOrdered(t, "line properties", ln, "<a:solidFill>", "<a:headEnd", "<a:tailEnd")
}

// TestConnectorArrowEndsReachTheLine protects the same emitter on the connector
// path, where the ends live inside the connector's own <a:ln>. It is the
// control for the freeform case above: one arrow end, and no head end written
// for a shape that has none.
func TestConnectorArrowEndsReachTheLine(t *testing.T) {
	p := New()
	l := p.GetActiveSlide().CreateLineShape()
	l.SetTailEnd(&LineEnd{Type: ArrowArrow, Width: "med", Length: "med"})

	ln := blockIn(t, slidePartOfShapeRound(t, p), "<a:ln ", "</a:ln>")

	if !strings.Contains(ln, `<a:tailEnd type="arrow" w="med" len="med"/>`) {
		t.Errorf("the connector's arrow end is not on its line. Got:\n%s", ln)
	}
	if strings.Contains(ln, "<a:headEnd") {
		t.Errorf("a connector with no head end wrote one. Got:\n%s", ln)
	}
}

// --- background pictures ---

// backgroundPicturePackage assembles a one-slide package whose slide background
// is a picture. The blipFill body is inserted verbatim, so a test can choose
// whether the background carries a crop and an opacity.
func backgroundPicturePackage(t *testing.T, blipFillBody string) []byte {
	t.Helper()
	return buildZip(t, map[string][]byte{
		"ppt/presentation.xml": []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
			`<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"` +
			` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
			`<p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst></p:presentation>`),
		"ppt/_rels/presentation.xml.rels": []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
			`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
			`<Relationship Id="rId1"` +
			` Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"` +
			` Target="slides/slide1.xml"/></Relationships>`),
		"ppt/slides/slide1.xml": []byte(slideWithBackground(blipFillBody)),
		"ppt/slides/_rels/slide1.xml.rels": []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
			`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
			`<Relationship Id="rId1"` +
			` Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"` +
			` Target="media/image1.png"/></Relationships>`),
		"ppt/slides/media/image1.png": twoTonePNG(t),
	})
}

// slideWithBackground returns a slide whose <p:bg> holds the given blipFill
// body. The background is the only content, so the picture the reader prepends
// is the only shape.
func slideWithBackground(blipFillBody string) string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
		`<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"` +
		` xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"` +
		` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
		`<p:cSld>` +
		`<p:bg><p:bgPr>` + blipFillBody + `<a:effectLst/></p:bgPr></p:bg>` +
		`<p:spTree>` +
		`<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
		`<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>` +
		`<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>` +
		`</p:spTree></p:cSld></p:sld>`
}

// readOneSlidePackage reads a package and returns its single slide.
func readOneSlidePackage(t *testing.T, data []byte) *Slide {
	t.Helper()
	reader, err := NewReader(ReaderPowerPoint2007)
	if err != nil {
		t.Fatalf("new reader: %v", err)
	}
	pres, err := reader.ReadFromReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatalf("read package: %v", err)
	}
	if pres.GetSlideCount() != 1 {
		t.Fatalf("slide count = %d, want 1", pres.GetSlideCount())
	}
	slide, err := pres.GetSlide(0)
	if err != nil {
		t.Fatalf("get slide: %v", err)
	}
	return slide
}

// TestBackgroundPictureCropIsApplied covers the reader. A <p:bg> picture is
// turned into a full-slide drawing, but the crop and the opacity live as
// siblings of <a:blip> rather than inside a <p:pic>, and only the <p:pic> form
// collected them — so a cropped background hero image was drawn whole.
func TestBackgroundPictureCropIsApplied(t *testing.T) {
	pkg := backgroundPicturePackage(t,
		`<a:blipFill><a:blip r:embed="rId1"/><a:srcRect l="50000" t="0" r="0" b="0"/>`+
			`<a:stretch><a:fillRect/></a:stretch></a:blipFill>`)

	slide := readOneSlidePackage(t, pkg)
	shapes := slide.GetShapes()
	if len(shapes) == 0 {
		t.Fatal("the background picture was not read at all")
	}
	d, ok := shapes[0].(*DrawingShape)
	if !ok {
		t.Fatalf("first shape is %T, want the background picture", shapes[0])
	}
	if got := d.GetCropLeft(); got != 50000 {
		t.Errorf("background crop left = %d, want 50000: the crop is a sibling of <a:blip>, "+
			"not a child of <p:pic>", got)
	}
	if len(d.GetImageData()) == 0 {
		t.Fatal("the background picture has no image data")
	}
}

// TestBackgroundPictureCropIsVisibleInTheRender checks the symptom rather than
// the field: the fixture's picture is red on its left half and blue on its
// right, and cropping away the left half means no red may survive anywhere.
func TestBackgroundPictureCropIsVisibleInTheRender(t *testing.T) {
	pkg := backgroundPicturePackage(t,
		`<a:blipFill><a:blip r:embed="rId1"/><a:srcRect l="50000" t="0" r="0" b="0"/>`+
			`<a:stretch><a:fillRect/></a:stretch></a:blipFill>`)
	pres, err := (&PPTXReader{}).ReadFromReader(bytes.NewReader(pkg), int64(len(pkg)))
	if err != nil {
		t.Fatalf("read package: %v", err)
	}

	imgs, err := pres.SlidesToImages(goldenOptions(NewFontCache()))
	if err != nil {
		t.Fatalf("render: %v", err)
	}
	if len(imgs) != 1 {
		t.Fatalf("rendered %d slides, want 1", len(imgs))
	}
	img := imgs[0]
	b := img.Bounds()

	red := color.RGBA{R: 255, A: 255}
	blue := color.RGBA{B: 255, A: 255}
	left := image.Rect(b.Min.X, b.Min.Y, b.Min.X+b.Dx()/4, b.Min.Y+b.Dy())
	right := image.Rect(b.Min.X+b.Dx()*3/4, b.Min.Y, b.Max.X, b.Min.Y+b.Dy())

	// The surviving half is stretched across the whole slide, so both quarters
	// are blue; without the crop the left quarter is the red half.
	if n := countNearColorIn(img, left, red, 8); n != 0 {
		t.Errorf("%d red pixels survive in the left quarter: the background's crop was not applied", n)
	}
	if n := countNearColorIn(img, left, blue, 8); n == 0 {
		t.Error("no blue in the left quarter; the cropped background was not drawn")
	}
	if n := countNearColorIn(img, right, blue, 8); n == 0 {
		t.Error("no blue in the right quarter; the cropped background was not drawn")
	}
}

// TestLayoutBackgroundPictureKeepsItsCrop covers the third carrier: a layout's
// background picture. That parser returned as soon as it saw the <a:blip>, so
// the <a:srcRect> that follows it was never reached.
func TestLayoutBackgroundPictureKeepsItsCrop(t *testing.T) {
	pkg := buildZip(t, map[string][]byte{
		"ppt/media/image1.png": twoTonePNG(t),
	})
	zr, err := zip.NewReader(bytes.NewReader(pkg), int64(len(pkg)))
	if err != nil {
		t.Fatalf("open zip: %v", err)
	}
	rels := []xmlRelForRead{{ID: "rId1", Type: relTypeImage, Target: "../media/image1.png"}}

	doc := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
		`<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"` +
		` xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"` +
		` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
		`<p:cSld><p:bg><p:bgPr>` +
		`<a:blipFill><a:blip r:embed="rId1"/><a:srcRect l="50000" t="0" r="0" b="0"/>` +
		`<a:alphaModFix amt="60000"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>` +
		`<a:effectLst/></p:bgPr></p:bg>` +
		`<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
		`<p:grpSpPr/></p:spTree></p:cSld></p:sldLayout>`

	reader := &PPTXReader{}
	fill, ds := reader.parseLayoutBackground([]byte(doc), rels, zr, "ppt/slideLayouts/slideLayout1.xml", New())

	if fill != nil {
		t.Error("parseLayoutBackground returned a fill for a picture background")
	}
	if ds == nil {
		t.Fatal("the layout background picture was not read")
	}
	if got := ds.GetCropLeft(); got != 50000 {
		t.Errorf("layout background crop left = %d, want 50000", got)
	}
	if got := ds.GetAlphaValue(); got != 60000 {
		t.Errorf("layout background alpha = %d, want 60000", got)
	}
	if len(ds.GetImageData()) == 0 {
		t.Error("the layout background picture has no image data")
	}
}

// --- placeholders ---

// TestPlaceholderWithoutTypeOmitsTheAttribute pins a value PowerPoint rejects.
// A <p:ph> may omit its type — placeholders inheriting from the layout do — and
// the reader keeps the empty string, but the writer formatted type="%s"
// unconditionally. ST_PlaceholderType has no empty member, so the file was
// invalid; this library's reader accepts it, so no round trip could see it.
func TestPlaceholderWithoutTypeOmitsTheAttribute(t *testing.T) {
	p := New()
	ph := p.GetActiveSlide().CreatePlaceholderShape("")
	ph.SetPlaceholderIndex(7)
	ph.SetText("inherited")

	slide := slidePartOfShapeRound(t, p)

	if strings.Contains(slide, `type=""`) {
		t.Errorf("an empty placeholder type was written; ST_PlaceholderType has no empty member. "+
			"Got:\n%s", blockIn(t, slide, "<p:nvPr>", "</p:nvPr>"))
	}
	if !strings.Contains(slide, `<p:ph idx="7"/>`) {
		t.Errorf("the type-less placeholder lost its index. Got:\n%s",
			blockIn(t, slide, "<p:nvPr>", "</p:nvPr>"))
	}
}

// TestTypedPlaceholderStillWritesItsType is the control: the attribute is only
// omitted when there is nothing to write.
func TestTypedPlaceholderStillWritesItsType(t *testing.T) {
	p := New()
	ph := p.GetActiveSlide().CreatePlaceholderShape(PlaceholderBody)
	ph.SetPlaceholderIndex(1)
	ph.SetText("body")

	slide := slidePartOfShapeRound(t, p)

	if !strings.Contains(slide, `<p:ph type="body" idx="1"/>`) {
		t.Errorf("a typed placeholder did not write its type. Got:\n%s",
			blockIn(t, slide, "<p:nvPr>", "</p:nvPr>"))
	}
}

// firstShapeOfType returns the first shape of a one-slide presentation that has
// the wanted type, looking through groups as well.
func firstShapeOfType[T Shape](pres *Presentation) (T, bool) {
	var zero T
	if pres.GetSlideCount() == 0 {
		return zero, false
	}
	slides := pres.GetAllSlides()
	if len(slides) == 0 {
		return zero, false
	}
	if v, ok := findShapeOfType[T](slides[0].GetShapes()); ok {
		return v, true
	}
	return zero, false
}

func findShapeOfType[T Shape](shapes []Shape) (T, bool) {
	var zero T
	for _, s := range shapes {
		if v, ok := s.(T); ok {
			return v, true
		}
		if g, ok := s.(*GroupShape); ok {
			if v, ok := findShapeOfType[T](g.GetShapes()); ok {
				return v, true
			}
		}
	}
	return zero, false
}
