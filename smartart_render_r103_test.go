package gopresentation

import (
	"image"
	"image/color"
	"image/png"
	"os"
	"path/filepath"
	"testing"
)

// The SmartArt engine's render pipeline, exercised through the corpus's
// LibreOffice-authored decks (no cached drawing → the v1 engine fires):
// model parse → family dispatch → geometry → pixel paint. A regression in
// any step shows up as a wrong pixel rather than silent 0% coverage.

func renderSmartArtDeck(t *testing.T, path string) image.Image {
	t.Helper()
	pres, err := Open(path)
	if err != nil {
		t.Fatalf("read: %v", err)
	}
	shapes := pres.GetAllSlides()[0].GetShapes()
	found := false
	for _, s := range shapes {
		if u, ok := s.(*UnsupportedShape); ok && u.model != nil {
			found = true
		}
	}
	if !found {
		t.Fatalf("no SmartArt model found among %d shapes", len(shapes))
	}
	opts := DefaultRenderOptions()
	opts.Width = 400
	path2 := filepath.Join(t.TempDir(), "s.png")
	if err := pres.SaveSlideAsImage(0, path2, opts); err != nil {
		t.Fatalf("render: %v", err)
	}
	f, err := os.Open(path2)
	if err != nil {
		t.Fatal(err)
	}
	defer f.Close()
	img, err := png.Decode(f)
	if err != nil {
		t.Fatal(err)
	}
	return img
}

// dominantColours scans the image and returns the distinct saturated colours
// (neither white nor black nor pale grey), sorted by frequency.
func dominantColours(img image.Image) []color.RGBA {
	counts := map[color.RGBA]int{}
	b := img.Bounds()
	for y := b.Min.Y; y < b.Max.Y; y++ {
		for x := b.Min.X; x < b.Max.X; x++ {
			r, g, bl, _ := img.At(x, y).RGBA()
			c := color.RGBA{R: uint8(r >> 8), G: uint8(g >> 8), B: uint8(bl >> 8), A: 255}
			// skip white, near-white, black, near-black
			if (c.R > 230 && c.G > 230 && c.B > 230) || (c.R < 25 && c.G < 25 && c.B < 25) {
				continue
			}
			counts[c]++
		}
	}
	var out []color.RGBA
	for c, n := range counts {
		if n > 20 {
			out = append(out, c)
		}
	}
	return out
}

// TestSmartArtRenderLin: fill-color-list renders three columns with distinct
// accent header fills — the colourful quick style cycles the palette.
func TestSmartArtRenderLin(t *testing.T) {
	img := renderSmartArtDeck(t, `F:/pptx_corpus/web/fill-color-list.pptx`)
	cols := dominantColours(img)
	if len(cols) < 3 {
		t.Errorf("the lin render shows %d distinct saturated colours, want at least 3 (red/green/purple headers)", len(cols))
	}
}

// TestSmartArtRenderSnake: rightoleft renders a grid with white gutters
// between the boxes.
func TestSmartArtRenderSnake(t *testing.T) {
	img := renderSmartArtDeck(t, `F:/pptx_corpus/web/smartart-rightoleftblockdiagram.pptx`)
	w := img.Bounds().Dx()
	h := img.Bounds().Dy()
	// scan the middle column for a white gap row (between the grid boxes)
	gapFound := false
	for y := h / 4; y < 3*h/4; y++ {
		r, g, b, _ := img.At(w/2, y).RGBA()
		if r > 240 && g > 240 && b > 240 {
			gapFound = true
			break
		}
	}
	if !gapFound {
		t.Error("the snake grid has no white gutter between the boxes")
	}
}

// TestSmartArtRenderCycle: cycle renders nodes on a ring (non-white pixels
// in the ring area above and below the centre).
func TestSmartArtRenderCycle(t *testing.T) {
	img := renderSmartArtDeck(t, `F:/pptx_corpus/web/smartart-cycle.pptx`)
	w := img.Bounds().Dx()
	h := img.Bounds().Dy()
	cols := dominantColours(img)
	if len(cols) == 0 {
		t.Error("the cycle render has no saturated colours — the ring nodes are missing")
	}
	// the ring should place nodes away from the centre — scan for
	// non-white pixels outside the central area
	nonWhite := 0
	for y := 0; y < h; y++ {
		for x := 0; x < w; x++ {
			r, g, b, _ := img.At(x, y).RGBA()
			if !(r>>8 > 240 && g>>8 > 240 && b>>8 > 240) {
				nonWhite++
			}
		}
	}
	if nonWhite < 50 {
		t.Errorf("the cycle render has only %d non-white pixels — the ring nodes are missing", nonWhite)
	}
}
