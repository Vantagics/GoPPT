package gopresentation

import (
	"image/png"
	"os"
	"path/filepath"
	"testing"
)

// The render canvas must match PowerPoint's own export pixel for pixel, or
// every downstream comparison starts with a size mismatch. PowerPoint's
// export rounds the derived height (1600 * 7562850/10077450 = 1200.756
// becomes 1201 rows); the renderer truncated, so a deck with this real-world
// slide size rendered 1600x1200 and no longer paired with the reference.
func TestCanvasHeightMatchesOfficeRounding_r75(t *testing.T) {
	p := New()
	layout := p.GetLayout()
	if layout == nil {
		t.Fatal("no layout")
	}
	layout.CX = 10077450
	layout.CY = 7562850

	opts := DefaultRenderOptions()
	opts.Width = 1600
	path := filepath.Join(t.TempDir(), "slide.png")
	if err := p.SaveSlideAsImage(0, path, opts); err != nil {
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
	if got := img.Bounds().Dy(); got != 1201 {
		t.Fatalf("canvas height = %d, want 1201 (rounded, as Office exports)", got)
	}
	if got := img.Bounds().Dx(); got != 1600 {
		t.Fatalf("canvas width = %d, want 1600", got)
	}
}
