package gopresentation

import (
	"image"
	"image/color"
	"strings"
	"testing"
)

// SVG rasterisation.
//
// The renderer has no SVG dependency, so svg.go implements the subset of SVG
// that real slides use for icon art and connector lines. These tests pin the two
// properties that matter: the supported subset draws real ink, and anything
// outside it is *refused with a reason* rather than approximated or silently
// dropped, because a mis-drawn icon is worse than a labelled placeholder.

// svgTestIcon is the fixture used throughout: a filled disc with a triangle
// punched over it, so fill, geometry and paint order all matter.
const svgTestIcon = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 16 16">` +
	`<circle cx="8" cy="8" r="7" fill="#336699"/>` +
	`<polygon points="8,3 13,12 3,12" fill="#FFFFFF"/>` +
	`</svg>`

// countColours returns how many pixels have exactly this opaque RGB.
func countColours(img *image.RGBA, want color.RGBA) int {
	n := 0
	b := img.Bounds()
	for y := b.Min.Y; y < b.Max.Y; y++ {
		for x := b.Min.X; x < b.Max.X; x++ {
			if img.RGBAAt(x, y) == want {
				n++
			}
		}
	}
	return n
}

// TestSVGRasterisesFilledShapes checks that the supported subset produces the
// colours the document asks for, at the sizes it asks for. A rasteriser that
// silently produced nothing would still return a valid image, so the assertion
// is on actual ink rather than on the absence of an error.
func TestSVGRasterisesFilledShapes(t *testing.T) {
	const size = 64
	img, err := renderSVG([]byte(svgTestIcon), size, size)
	if err != nil {
		t.Fatalf("renderSVG: %v", err)
	}
	if img.Bounds() != image.Rect(0, 0, size, size) {
		t.Fatalf("image bounds = %v, want %v", img.Bounds(), image.Rect(0, 0, size, size))
	}

	disc := color.RGBA{R: 0x33, G: 0x66, B: 0x99, A: 0xFF}
	hole := color.RGBA{R: 0xFF, G: 0xFF, B: 0xFF, A: 0xFF}
	if n := countColours(img, disc); n < 200 {
		t.Errorf("filled disc covers only %d exact pixels at %dpx; the circle was not filled", n, size)
	}
	if n := countColours(img, hole); n < 50 {
		t.Errorf("triangle covers only %d exact pixels at %dpx; the polygon was not filled", n, size)
	}
	// Corners are outside the disc, so they must stay transparent: the raster
	// must not fill its own background.
	for _, p := range []image.Point{{0, 0}, {size - 1, 0}, {0, size - 1}, {size - 1, size - 1}} {
		if _, _, _, a := img.At(p.X, p.Y).RGBA(); a != 0 {
			t.Errorf("corner %v has alpha %d, want 0: the SVG background must stay transparent", p, a>>8)
		}
	}
}

// TestSVGHonoursViewBoxScaling checks that the document's user units are mapped
// onto the requested raster rather than drawn at their nominal size. A 16-unit
// viewBox rendered into 128 px has to scale by 8.
func TestSVGHonoursViewBoxScaling(t *testing.T) {
	small, err := renderSVG([]byte(svgTestIcon), 32, 32)
	if err != nil {
		t.Fatalf("renderSVG 32px: %v", err)
	}
	large, err := renderSVG([]byte(svgTestIcon), 128, 128)
	if err != nil {
		t.Fatalf("renderSVG 128px: %v", err)
	}
	disc := color.RGBA{R: 0x33, G: 0x66, B: 0x99, A: 0xFF}
	ns, nl := countColours(small, disc), countColours(large, disc)
	if ns == 0 || nl == 0 {
		t.Fatalf("disc missing at one of the sizes: %d px at 32, %d px at 128", ns, nl)
	}
	// Area scales with the square of the linear factor; allow generous slack for
	// anti-aliasing at the edges.
	ratio := float64(nl) / float64(ns)
	if ratio < 8 || ratio > 24 {
		t.Errorf("disc area grew %.1fx from 32px to 128px, want roughly 16x: "+
			"the viewBox is not being scaled (%d px vs %d px)", ratio, ns, nl)
	}
}

// TestSVGRefusesConstructsOutsideTheSubset is the safety net for the whole
// rasteriser. Every case here would render *something* if the construct were
// quietly ignored, and the something would be wrong; the contract is to return
// an error naming the construct so the caller can draw a placeholder instead.
func TestSVGRefusesConstructsOutsideTheSubset(t *testing.T) {
	const head = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 16 16">`

	cases := []struct {
		name string
		doc  string
		want string
	}{
		{
			name: "text element",
			doc:  head + `<text x="1" y="8">hi</text></svg>`,
			want: "<text>",
		},
		{
			name: "embedded stylesheet",
			doc:  head + `<style>.a{fill:red}</style><rect width="16" height="16" fill="#f00"/></svg>`,
			want: "<style>",
		},
		{
			name: "use reference",
			doc:  head + `<use href="#x"/></svg>`,
			want: "<use>",
		},
		{
			name: "filter",
			doc:  head + `<filter id="f"/><rect width="16" height="16" fill="#f00"/></svg>`,
			want: "<filter>",
		},
		{
			name: "nested svg",
			doc:  head + `<svg x="0" y="0" width="8" height="8"><rect width="8" height="8" fill="#f00"/></svg></svg>`,
			want: "nested <svg>",
		},
		{
			name: "unknown element",
			doc:  head + `<foreignObject width="16" height="16"/></svg>`,
			want: "<foreignobject>",
		},
		{
			name: "unused gradient definition",
			doc: head + `<defs><linearGradient id="g"><stop offset="0" stop-color="#f00"/>` +
				`</linearGradient></defs><rect width="8" height="8" fill="#0f0"/></svg>`,
			// Gradients are inert until painted, and painted gradients are
			// refused by parsePaint, so a definition alone must not fail.
			want: "",
		},
		{
			name: "malformed xml",
			doc:  head + `<rect width="16" height="16"`,
			want: "cannot parse XML",
		},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			img, err := renderSVG([]byte(tc.doc), 32, 32)
			if tc.want == "" {
				if err != nil {
					t.Fatalf("renderSVG refused a supported document (%s): %v", tc.name, err)
				}
				if img == nil {
					t.Fatalf("renderSVG returned no image for %s", tc.name)
				}
				return
			}
			if err == nil {
				t.Fatalf("renderSVG accepted an out-of-subset document (%s) and returned %v",
					tc.name, img.Bounds())
			}
			if !strings.Contains(err.Error(), tc.want) {
				t.Errorf("error %q does not name the offending construct (want %q)", err, tc.want)
			}
		})
	}
}

// TestSVGWithNothingDrawableIsAnError covers the subtler half of the contract:
// a document that is inside the subset but paints nothing would produce a fully
// transparent raster, which is indistinguishable from a picture that is meant
// to be blank. That has to be an error too.
func TestSVGWithNothingDrawableIsAnError(t *testing.T) {
	docs := map[string]string{
		"empty svg": `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 16 16"></svg>`,
		"explicitly unpainted": `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 16 16">` +
			`<rect width="16" height="16" fill="none"/></svg>`,
		"no svg element": "this is not markup at all",
	}
	for name, doc := range docs {
		t.Run(name, func(t *testing.T) {
			if _, err := renderSVG([]byte(doc), 32, 32); err == nil {
				t.Error("renderSVG returned success for a document with nothing drawable")
			}
		})
	}
}

// TestSVGAppliesTheInitialFill checks SVG's initial paint state: the fill
// defaults to solid black, so geometry with no fill attribute is painted rather
// than skipped. Treating it as unpainted would drop most of a real icon, whose
// paths often rely on the default.
func TestSVGAppliesTheInitialFill(t *testing.T) {
	const doc = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 16 16">` +
		`<rect width="16" height="16"/></svg>`
	img, err := renderSVG([]byte(doc), 32, 32)
	if err != nil {
		t.Fatalf("renderSVG: %v", err)
	}
	black := color.RGBA{R: 0, G: 0, B: 0, A: 0xFF}
	if n := countColours(img, black); n == 0 {
		t.Error("a rect with no fill attribute painted nothing; SVG's initial fill is black")
	}
}

// TestSVGDoesNotDrawDefinitions checks that <defs> content stays out of the
// picture. Its children are definitions of reusable parts, drawn only where
// something references them; drawing them in place would sprinkle artwork the
// document deliberately keeps off the canvas across the picture.
func TestSVGDoesNotDrawDefinitions(t *testing.T) {
	const doc = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 16 16">` +
		`<defs><rect width="16" height="16" fill="#FF0000"/>` +
		`<path d="M0 0 L16 16 L0 16 Z" fill="#FF0000"/></defs>` +
		`<rect width="8" height="8" fill="#00FF00"/></svg>`
	img, err := renderSVG([]byte(doc), 32, 32)
	if err != nil {
		t.Fatalf("renderSVG: %v", err)
	}

	red := color.RGBA{R: 0xFF, G: 0, B: 0, A: 0xFF}
	if n := countColours(img, red); n != 0 {
		t.Errorf("%d pixels of definition colour were drawn; <defs> content is not content", n)
	}
	green := color.RGBA{R: 0, G: 0xFF, B: 0, A: 0xFF}
	if n := countColours(img, green); n == 0 {
		t.Error("the shape outside <defs> was not drawn")
	}
}

// TestSVGDefinitionsDoNotLeakPastTheirElement checks that the suppression is
// scoped: a shape after </defs> must be drawn again.
func TestSVGDefinitionsDoNotLeakPastTheirElement(t *testing.T) {
	const doc = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 16 16">` +
		`<defs><rect width="16" height="16" fill="#FF0000"/></defs>` +
		`<rect x="0" y="0" width="4" height="4" fill="#00FF00"/>` +
		`<rect x="4" y="4" width="4" height="4" fill="#0000FF"/></svg>`
	img, err := renderSVG([]byte(doc), 32, 32)
	if err != nil {
		t.Fatalf("renderSVG: %v", err)
	}
	red := color.RGBA{R: 0xFF, G: 0, B: 0, A: 0xFF}
	if n := countColours(img, red); n != 0 {
		t.Errorf("%d pixels of definition colour were drawn", n)
	}
	for _, want := range []color.RGBA{
		{R: 0, G: 0xFF, B: 0, A: 0xFF},
		{R: 0, G: 0, B: 0xFF, A: 0xFF},
	} {
		if n := countColours(img, want); n == 0 {
			t.Errorf("shape with fill #%02X%02X%02X was not drawn after </defs>", want.R, want.G, want.B)
		}
	}
}

// TestSVGRejectsUnusableTargetSize checks the guards that keep a hostile
// viewport from allocating an enormous raster.
func TestSVGRejectsUnusableTargetSize(t *testing.T) {
	for _, size := range [][2]int{{0, 10}, {10, 0}, {-1, 10}, {maxSVGDimension + 1, 10}} {
		if _, err := renderSVG([]byte(svgTestIcon), size[0], size[1]); err == nil {
			t.Errorf("renderSVG accepted target size %dx%d", size[0], size[1])
		}
	}
}

// TestLooksLikeSVGRecognisesRealDocuments checks the content sniff that decides
// whether to hand a picture to the rasteriser. It has to accept an SVG part
// whatever the declaration in front of it, and it must not claim a raster is
// SVG — the caller uses a true answer to choose a renderer, so a false positive
// would turn a decodable image into a placeholder.
func TestLooksLikeSVGRecognisesRealDocuments(t *testing.T) {
	yes := map[string]string{
		"bare svg":            `<svg xmlns="http://www.w3.org/2000/svg"/>`,
		"xml declaration":     `<?xml version="1.0" encoding="UTF-8"?><svg xmlns="http://www.w3.org/2000/svg"/>`,
		"leading whitespace":  "\n\t  " + `<svg xmlns="http://www.w3.org/2000/svg"/>`,
		"comment before root": `<?xml version="1.0"?><!-- generated --><svg xmlns="http://www.w3.org/2000/svg"/>`,
	}
	for name, data := range yes {
		if !looksLikeSVG([]byte(data)) {
			t.Errorf("looksLikeSVG(%s) = false, want true", name)
		}
	}

	no := map[string]string{
		"empty":        "",
		"plain text":   "hello",
		"json":         `{"a":1}`,
		"html":         `<!DOCTYPE html><html><body>hi</body></html>`,
		"png header":   "\x89PNG\r\n\x1a\n",
		"jpeg header":  "\xff\xd8\xff\xe0",
		"xml non-svg":  `<?xml version="1.0"?><root><child/></root>`,
		"svg in a tag": `<a>svg</a>`,
	}
	for name, data := range no {
		if looksLikeSVG([]byte(data)) {
			t.Errorf("looksLikeSVG(%s) = true, want false", name)
		}
	}
}

// TestSVGPictureRendersInkNotAPlaceholder is the end-to-end check: a picture
// holding SVG bytes must be drawn, not replaced by the unsupported-construct
// placeholder. The placeholder is identifiable by its amber border, which no
// part of the fixture uses.
func TestSVGPictureRendersInkNotAPlaceholder(t *testing.T) {
	pres := New()
	shape := pres.GetActiveSlide().CreateDrawingShape()
	shape.BaseShape.SetOffsetX(1000000).SetOffsetY(1000000)
	shape.BaseShape.SetWidth(3000000).SetHeight(3000000)
	shape.SetImageData([]byte(svgTestIcon), "image/svg+xml")

	const width = 640
	opts := DefaultRenderOptions()
	opts.Width = width
	img, err := pres.SlideToImage(0, opts)
	if err != nil {
		t.Fatalf("render: %v", err)
	}
	rgba := toRGBA(t, img)

	rect := emuRect(t, pres, 1000000, 1000000, 3000000, 3000000, width)
	if n := countInkIn(img, rect); n == 0 {
		t.Fatal("the SVG picture drew no ink at all")
	}
	placeholder := color.RGBA{R: 214, G: 141, B: 17, A: 255}
	if n := countColours(rgba, placeholder); n > 0 {
		t.Errorf("found %d pixels of the placeholder border colour: the SVG picture was "+
			"replaced by a placeholder instead of being rasterised", n)
	}
	if n := countColours(rgba, color.RGBA{R: 0x33, G: 0x66, B: 0x99, A: 0xFF}); n == 0 {
		t.Error("the picture's own fill colour is absent from the render")
	}
}

// TestUndecodablePictureShowsAPlaceholder is the negative side of the same
// contract: a picture the renderer genuinely cannot decode must be visible as a
// placeholder. Drawing nothing is what made the missing-SVG bug hard to see.
func TestUndecodablePictureShowsAPlaceholder(t *testing.T) {
	pres := New()
	shape := pres.GetActiveSlide().CreateDrawingShape()
	shape.BaseShape.SetOffsetX(1000000).SetOffsetY(1000000)
	shape.BaseShape.SetWidth(3000000).SetHeight(3000000)
	shape.SetImageData([]byte("not an image in any format"), "application/octet-stream")

	const width = 640
	opts := DefaultRenderOptions()
	opts.Width = width
	img, err := pres.SlideToImage(0, opts)
	if err != nil {
		t.Fatalf("render: %v", err)
	}
	placeholder := color.RGBA{R: 214, G: 141, B: 17, A: 255}
	if n := countColours(toRGBA(t, img), placeholder); n == 0 {
		t.Error("an undecodable picture drew no placeholder border, so the failure is invisible")
	}
}
