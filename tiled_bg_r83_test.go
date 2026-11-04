package gopresentation

import (
	"bytes"
	"image"
	"image/color"
	"image/png"
	"os"
	"path/filepath"
	"testing"
)

// A background blipFill with <a:tile/> repeats the bitmap at its native size
// instead of stretching it over the slide. The corpus deck whose brick wall
// carried a bare tile rendered one house-sized brick — 22% of the slide
// against Office's export — and after the fix lands under 1%.
func TestTiledBackgroundRepeatsTheBitmap(t *testing.T) {
	parts := zipParts(t, writeToBytes(t, New()))
	parts["ppt/slides/slide1.xml"] = []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:bg><p:bgPr><a:blipFill><a:blip r:embed="rId1"/><a:tile/></a:blipFill><a:effectLst/></p:bgPr></p:bg>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
    </p:spTree></p:cSld>
</p:sld>`)
	parts["ppt/slides/_rels/slide1.xml.rels"] = []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`)
	// A 4x4 image, opaque centre 2x2, transparent border: the tile must land
	// wherever the pattern lands and the background around each tile shows
	// the transparent border composited on the canvas — with draw.Draw's
	// crop semantics a naive tile paste leaves the rest of the cell bare.
	w := image.NewRGBA(image.Rect(0, 0, 4, 4))
	for y := 1; y < 3; y++ {
		for x := 1; x < 3; x++ {
			w.Set(x, y, color.RGBA{R: 220, G: 60, B: 20, A: 255})
		}
	}
	var pngBuf bytes.Buffer
	png.Encode(&pngBuf, w)
	parts["ppt/media/image1.png"] = pngBuf.Bytes()
	pkg := buildZip(t, parts)

	pres, err := (&PPTXReader{}).ReadFromReader(bytes.NewReader(pkg), int64(len(pkg)))
	if err != nil {
		t.Fatalf("read: %v", err)
	}
	opts := DefaultRenderOptions()
	opts.Width = 960
	path := filepath.Join(t.TempDir(), "slide.png")
	if err := pres.SaveSlideAsImage(0, path, opts); err != nil {
		t.Fatalf("render: %v", err)
	}
	f, err := os.Open(path)
	if err != nil {
		t.Fatal(err)
	}
	defer f.Close()
	img, err := png.Decode(f)
	if err != nil {
		t.Fatal(err)
	}
	// At width 960 the canvas is exactly 96 dpi over the 10in slide, so a
	// 4px image tile lands cell-for-cell on one render pixel: the pattern
	// must repeat through every cell — paste-and-crop semantics would blank
	// everything past the first cell.
	for _, cx := range []int{2, 6, 10, 22} {
		r, g, b, _ := img.At(cx, 2).RGBA()
		if r>>8 != 220 || g>>8 != 60 || b>>8 != 20 {
			t.Errorf("tile cell at x=%d: (%d,%d,%d), want the brick fill — the pattern stopped repeating", cx, r>>8, g>>8, b>>8)
		}
	}
}
