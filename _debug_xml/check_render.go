package main

import (
	"archive/zip"
	"bytes"
	"fmt"
	"image"
	"image/color"
	"image/png"
	"io"
	"os"

	_ "image/gif"
	_ "image/jpeg"

	gopresentation "github.com/VantageDataChat/GoPPT"
)

func main() {
	// 1. Render slide 40
	reader, _ := gopresentation.NewReader(gopresentation.ReaderPowerPoint2007)
	pres, _ := reader.Read("test.pptx")
	opts := gopresentation.DefaultRenderOptions()
	opts.Width = 1920
	pres.SaveSlideAsImage(39, "slide40_current.png", opts)

	// 2. Read the rendered image
	f, _ := os.Open("slide40_current.png")
	defer f.Close()
	img, _ := png.Decode(f)
	bounds := img.Bounds()
	w, h := bounds.Dx(), bounds.Dy()
	fmt.Printf("Rendered: %dx%d\n", w, h)

	// 3. Check specific icon positions (from XML EMU coords)
	// Slide is 12192000 EMU wide, 6858000 EMU tall
	type iconCheck struct {
		name string
		emuX, emuY, emuW, emuH int64
	}
	icons := []iconCheck{
		{"EMF-icon1(rId1)", 1003300, 4758055, 409575, 320675},
		{"EMF-icon2(rId2)", 1711960, 4757420, 409575, 319405},
		{"PNG-fire(rId3)", 1206500, 4879975, 387985, 253365},
		{"EMF-icon3(rId1)", 4961890, 4778375, 409575, 311785},
		{"EMF-icon4(rId2)", 5670550, 4777740, 409575, 310515},
		{"EMF-icon5(rId1)", 7554595, 4778375, 409575, 311785},
		{"EMF-icon6(rId2)", 8263255, 4777740, 409575, 310515},
	}

	for _, ic := range icons {
		px := int(float64(ic.emuX) / 12192000 * float64(w))
		py := int(float64(ic.emuY) / 6858000 * float64(h))
		pw := int(float64(ic.emuW) / 12192000 * float64(w))
		ph := int(float64(ic.emuH) / 6858000 * float64(h))

		nonWhite := 0
		gray200 := 0
		colored := 0
		total := 0
		for y := py; y < py+ph && y < h; y++ {
			for x := px; x < px+pw && x < w; x++ {
				cr, cg, cb, ca := img.At(x, y).RGBA()
				r8 := uint8(cr >> 8)
				g8 := uint8(cg >> 8)
				b8 := uint8(cb >> 8)
				a8 := uint8(ca >> 8)
				total++
				if a8 > 0 && !(r8 == 255 && g8 == 255 && b8 == 255) {
					nonWhite++
					if r8 == 200 && g8 == 200 && b8 == 200 {
						gray200++
					} else {
						colored++
					}
				}
			}
		}
		fmt.Printf("  %s at (%d,%d) %dx%d: total=%d nonWhite=%d gray=%d colored=%d\n",
			ic.name, px, py, pw, ph, total, nonWhite, gray200, colored)
	}

	// 4. Test EMF decoding directly
	fmt.Println("\n=== Direct EMF decode test ===")
	zr, _ := zip.OpenReader("test.pptx")
	defer zr.Close()
	emfFiles := map[string]string{
		"ppt/media/image95.emf": "EMF95",
		"ppt/media/image96.emf": "EMF96",
	}
	for _, zf := range zr.File {
		if name, ok := emfFiles[zf.Name]; ok {
			rc, _ := zf.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()

			// Try standard decode
			_, _, err := image.Decode(bytes.NewReader(data))
			fmt.Printf("  %s (%d bytes): standard decode err=%v\n", name, len(data), err)

			// Try metafile decode (this is what the renderer does)
			fc := gopresentation.NewFontCache()
			decoded := gopresentation.DecodeMetafileBitmapExported(data, fc)
			if decoded != nil {
				b := decoded.Bounds()
				// Count non-transparent pixels
				nonTrans := 0
				for y := b.Min.Y; y < b.Max.Y; y++ {
					for x := b.Min.X; x < b.Max.X; x++ {
						_, _, _, a := decoded.At(x, y).RGBA()
						if a > 0 {
							nonTrans++
						}
					}
				}
				fmt.Printf("    decoded: %dx%d, nonTransparent=%d/%d\n", b.Dx(), b.Dy(), nonTrans, b.Dx()*b.Dy())

				// Save decoded EMF as PNG for inspection
				outName := fmt.Sprintf("_debug_xml/%s_decoded.png", name)
				out, _ := os.Create(outName)
				png.Encode(out, decoded)
				out.Close()
				fmt.Printf("    saved to %s\n", outName)
			} else {
				fmt.Printf("    decoded: nil (failed)\n")
			}
		}
	}

	// 5. Check if there are any gray rectangles (failed decode markers)
	grayCount := 0
	for y := bounds.Min.Y; y < bounds.Max.Y; y++ {
		for x := bounds.Min.X; x < bounds.Max.X; x++ {
			cr, cg, cb, ca := img.At(x, y).RGBA()
			if uint8(cr>>8) == 200 && uint8(cg>>8) == 200 && uint8(cb>>8) == 200 && uint8(ca>>8) == 255 {
				grayCount++
			}
		}
	}
	fmt.Printf("\nTotal gray placeholder pixels: %d\n", grayCount)

	// 6. Sample some unique colors in the bottom icon area
	colorMap := map[color.RGBA]int{}
	for y := h * 65 / 100; y < h*75/100; y++ {
		for x := 0; x < w; x++ {
			cr, cg, cb, ca := img.At(x, y).RGBA()
			c := color.RGBA{uint8(cr >> 8), uint8(cg >> 8), uint8(cb >> 8), uint8(ca >> 8)}
			if c.A > 0 && c.R != 255 && c.G != 255 && c.B != 255 {
				colorMap[c]++
			}
		}
	}
	fmt.Printf("\nUnique non-white colors in icon area: %d\n", len(colorMap))
	// Top 10
	type cc struct {
		c color.RGBA
		n int
	}
	var sorted []cc
	for c, n := range colorMap {
		sorted = append(sorted, cc{c, n})
	}
	for i := 0; i < len(sorted); i++ {
		for j := i + 1; j < len(sorted); j++ {
			if sorted[j].n > sorted[i].n {
				sorted[i], sorted[j] = sorted[j], sorted[i]
			}
		}
	}
	for i, c := range sorted {
		if i >= 10 {
			break
		}
		fmt.Printf("  (%d,%d,%d,%d): %d\n", c.c.R, c.c.G, c.c.B, c.c.A, c.n)
	}
}
