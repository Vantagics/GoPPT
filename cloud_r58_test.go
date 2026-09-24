package gopresentation

import (
	"bytes"
	"image"
	"image/color"
	"math"
	"testing"
)

// deck2 slide27 stacks two cloudCallout AutoShapes with circle-path gradient
// fills. Three things were broken, all pinned here:
//   - the renderer had no cloudCallout geometry at all, so both callouts fell
//     to the rect fallback (a rounded box instead of a puffy cloud);
//   - the ECMA-376 guide chain (POI presetShapeDefinitions.xml) evaluates
//     cat2/sat2 projections to walk three tail bubbles from the callout tip
//     toward the cloud — none of them were drawn;
//   - the tail bubble centres skipped the box offset (ox/oy) that the cloud
//     body applied, so on the direct-canvas path they landed near the origin
//     and off the image entirely.

const r58CloudBoxX = 2324595
const r58CloudBoxY = 1695171
const r58CloudBoxW = 3352800
const r58CloudBoxH = 1576387

const r58CloudXML = `<p:sp>
  <p:nvSpPr><p:cNvPr id="2" name="Cloud Callout"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="2324595" y="1695171"/><a:ext cx="3352800" cy="1576387"/></a:xfrm>
    <a:prstGeom prst="cloudCallout">
      <a:avLst>
        <a:gd name="adj1" fmla="val -53418"/>
        <a:gd name="adj2" fmla="val 56473"/>
      </a:avLst>
    </a:prstGeom>
    <a:solidFill><a:srgbClr val="D6E4F7"/></a:solidFill>
    <a:ln w="25400"><a:solidFill><a:srgbClr val="1F497D"/></a:solidFill></a:ln>
  </p:spPr>
  <p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/>
    <a:p><a:pPr algn="ctr"/><a:endParaRPr lang="en-US"/></a:p>
  </p:txBody>
</p:sp>`

// r58Render renders a slide whose spTree holds body, returning the image and
// the callout box in pixels.
func r58Render(t *testing.T, body string) (image.Image, int, int, int, int) {
	t.Helper()
	parts := zipParts(t, writeToBytes(t, New()))
	parts["ppt/slides/slide1.xml"] = []byte(bodyPrSlide(body))
	pkg := buildZip(t, parts)
	pres, err := (&PPTXReader{}).ReadFromReader(bytes.NewReader(pkg), int64(len(pkg)))
	if err != nil {
		t.Fatalf("read the package: %v", err)
	}
	const rw = 1600
	opts := DefaultRenderOptions()
	opts.Width = rw
	opts.FontCache = NewFontCache()
	img, err := pres.SlideToImage(0, opts)
	if err != nil {
		t.Fatalf("render: %v", err)
	}
	x0 := r58CloudBoxX * rw / 9144000
	y0 := r58CloudBoxY * rw / 9144000
	w := r58CloudBoxW * rw / 9144000
	h := r58CloudBoxH * rw / 9144000
	return img, x0, y0, w, h
}

// navyRingHits counts how many of 36 sampled angles around a circle carry the
// stroke colour — the bubble outlines must survive as closed rings.
func navyRingHits(t *testing.T, img interface{ At(x, y int) color.Color }, cx, cy, r float64) int {
	t.Helper()
	hits := 0
	for a := 0; a < 360; a += 10 {
		x := int(cx + r*math.Cos(float64(a)*math.Pi/180))
		y := int(cy + r*math.Sin(float64(a)*math.Pi/180))
		c := color.NRGBAModel.Convert(img.At(x, y)).(color.NRGBA)
		if c.B-c.R > 40 && c.B > 90 && c.R < 120 {
			hits++
		}
	}
	return hits
}

// TestCloudCalloutTailBubbles renders the deck's callout and checks all three
// tail bubbles draw as closed navy rings at the positions the guide chain
// predicts. With the missing-offset bug every bubble landed at the canvas
// origin (0/36 hits); with no geometry at all, nothing rendered.
func TestCloudCalloutTailBubbles(t *testing.T) {
	img, x0, y0, w, h := r58Render(t, r58CloudXML)

	// Guide chain for the tip: xPos = w/2 + w·adj1/1e5, yPos likewise.
	tipX := float64(x0) + float64(w)*(0.5-0.53418)
	tipY := float64(y0) + float64(h)*(0.5+0.56473)
	ss := math.Min(float64(w), float64(h))
	radii := []float64{ss * 600 / 21600, ss * 1200 / 21600, ss * 1800 / 21600}

	hits := navyRingHits(t, img, tipX, tipY, radii[0])
	if hits < 25 {
		t.Fatalf("tip bubble ring only %d/36 navy hits at (%.1f,%.1f) r=%.1f: the tail bubbles are missing or offset", hits, tipX, tipY, radii[0])
	}
	// The other two bubbles walk the tip→cloud line with growing radii; pin
	// the largest one, whose centre the chain puts at g23/g24 — verify via
	// the geometry helper rather than duplicating the whole chain here.
	_, tails, _ := cloudCalloutGeometry(x0, y0, w, h, map[string]int{"adj1": -53418, "adj2": 56473})
	if len(tails) != 3 {
		t.Fatalf("expected 3 tail bubbles, got %d", len(tails))
	}
	for i, c := range tails {
		if c.r != radii[i] {
			t.Fatalf("tail%d radius %.2f, want %.2f", i, c.r, radii[i])
		}
		h := navyRingHits(t, img, c.cx, c.cy, c.r)
		if h < 25 {
			t.Fatalf("tail%d ring only %d/36 navy hits at (%.1f,%.1f) r=%.1f", i, h, c.cx, c.cy, c.r)
		}
	}
}

// TestCloudCalloutGeometryGuides pins the guide-chain arithmetic: the tip
// bubble sits exactly at (xPos,yPos), radii follow ss·{600,1200,1800}/21600,
// and the chain walks away from the tip toward the cloud with growing steps.
func TestCloudCalloutGeometryGuides(t *testing.T) {
	w, h := 587, 276
	_, tails, details := cloudCalloutGeometry(0, 0, w, h, map[string]int{"adj1": -53418, "adj2": 56473})

	wantX := float64(w) * (0.5 - 0.53418)
	wantY := float64(h) * (0.5 + 0.56473)
	if math.Abs(tails[0].cx-wantX) > 0.5 || math.Abs(tails[0].cy-wantY) > 0.5 {
		t.Fatalf("tip bubble centre (%.2f,%.2f), want (%.2f,%.2f)", tails[0].cx, tails[0].cy, wantX, wantY)
	}
	ss := math.Min(float64(w), float64(h))
	for i, want := range []float64{ss * 600 / 21600, ss * 1200 / 21600, ss * 1800 / 21600} {
		if math.Abs(tails[i].r-want) > 1e-9 {
			t.Fatalf("tail%d radius %.4f, want %.4f", i, tails[i].r, want)
		}
	}
	d01 := math.Hypot(tails[1].cx-tails[0].cx, tails[1].cy-tails[0].cy)
	d12 := math.Hypot(tails[2].cx-tails[1].cx, tails[2].cy-tails[1].cy)
	if d12 <= d01 {
		t.Fatalf("bubble walk must accelerate away from the tip: d01=%.1f d12=%.1f", d01, d12)
	}
	if len(details) != 11 {
		t.Fatalf("expected 11 crease arcs, got %d", len(details))
	}
	// Crease arcs are stroke-only subpaths — open, so 2+ points each.
	for i, d := range details {
		if len(d) < 2 {
			t.Fatalf("crease arc %d degenerate: %d points", i, len(d))
		}
	}
}

// TestCloudCalloutWriterRoundTrip: the writer half is generic prstGeomXML,
// but it must be pinned — a save that drops prst="cloudCallout" or the avLst
// reverts the shape to a default cloud on reopen.
func TestCloudCalloutWriterRoundTrip(t *testing.T) {
	pres := New()
	slide := pres.GetActiveSlide()
	shape := NewAutoShape()
	shape.SetAutoShapeType(AutoShapeCloudCallout)
	shape.SetPosition(2324595, 1695171)
	shape.SetSize(3352800, 1576387)
	shape.SetAdjustValue("adj1", -53418)
	shape.SetAdjustValue("adj2", 56473)
	slide.AddShape(shape)

	parts := zipParts(t, writeToBytes(t, pres))
	xmlStr := string(parts["ppt/slides/slide1.xml"])
	for _, want := range []string{`prst="cloudCallout"`, `name="adj1" fmla="val -53418"`, `name="adj2" fmla="val 56473"`} {
		if !bytes.Contains([]byte(xmlStr), []byte(want)) {
			t.Fatalf("saved slide XML missing %q", want)
		}
	}

	pkg := buildZip(t, parts)
	pres2, err := (&PPTXReader{}).ReadFromReader(bytes.NewReader(pkg), int64(len(pkg)))
	if err != nil {
		t.Fatalf("re-read: %v", err)
	}
	s2, err := pres2.GetSlide(0)
	if err != nil {
		t.Fatalf("slide 0: %v", err)
	}
	var got *AutoShape
	for _, sh := range s2.GetShapes() {
		if as, ok := sh.(*AutoShape); ok && as.shapeType == AutoShapeCloudCallout {
			got = as
		}
	}
	if got == nil {
		t.Fatalf("cloudCallout lost on round-trip")
	}
	if got.adjustValues["adj1"] != -53418 || got.adjustValues["adj2"] != 56473 {
		t.Fatalf("adjustments lost on round-trip: %v", got.adjustValues)
	}
}
