package main

import (
	"fmt"
	"image"
	"image/png"
	"os"
	"sort"
	"strconv"
	"strings"
)

// inspectPixels reports what is actually in a rectangle of a rendered PNG.
//
// It exists because a preview is judged by eye, and an eye cannot tell a glyph
// filled with a 12%-opacity colour from one that was only outlined — both read
// as "a faint shape". A colour histogram settles it: a filled glyph contributes
// tens of thousands of pixels of one colour, an outline contributes a few
// hundred spread across antialiasing steps.
//
// A coarse ASCII map of the region is printed alongside, so the shape can be
// recognised without opening the image.
//
//	go run ./cmd/debug_preview --pix <file.png> [x y w h]
func inspectPixels(args []string) {
	if len(args) < 1 {
		fmt.Fprintln(os.Stderr, "usage: --pix <file.png> [x y w h]")
		os.Exit(2)
	}
	f, err := os.Open(args[0])
	if err != nil {
		fmt.Printf("open: %v\n", err)
		return
	}
	defer f.Close()
	img, err := png.Decode(f)
	if err != nil {
		fmt.Printf("decode: %v\n", err)
		return
	}
	b := img.Bounds()
	rect := b
	if len(args) >= 5 {
		var vals [4]int
		bad := false
		for i := range vals {
			v, err := strconv.Atoi(args[1+i])
			if err != nil {
				bad = true
				break
			}
			vals[i] = v
		}
		if bad {
			fmt.Fprintln(os.Stderr, "usage: --pix <file.png> [x y w h]  (x y w h must be integers)")
			os.Exit(2)
		}
		rect = image.Rect(b.Min.X+vals[0], b.Min.Y+vals[1], b.Min.X+vals[0]+vals[2], b.Min.Y+vals[1]+vals[3]).Intersect(b)
	}
	if rect.Empty() {
		fmt.Printf("region %v is outside %v\n", rect, b)
		return
	}
	fmt.Printf("image %dx%d  region %v (%dx%d)\n", b.Dx(), b.Dy(), rect, rect.Dx(), rect.Dy())

	counts := map[uint32]int{}
	for y := rect.Min.Y; y < rect.Max.Y; y++ {
		for x := rect.Min.X; x < rect.Max.X; x++ {
			counts[packRGBA8(img.At(x, y))]++
		}
	}
	type entry struct {
		c uint32
		n int
	}
	list := make([]entry, 0, len(counts))
	for c, n := range counts {
		list = append(list, entry{c, n})
	}
	sort.Slice(list, func(i, j int) bool {
		if list[i].n != list[j].n {
			return list[i].n > list[j].n
		}
		return list[i].c < list[j].c
	})
	total := rect.Dx() * rect.Dy()
	fmt.Printf("distinct colours: %d\n", len(counts))
	shown := list
	if len(shown) > 14 {
		shown = shown[:14]
	}
	for _, e := range shown {
		fmt.Printf("  %s %8d  %5.2f%%\n", formatRGBA8(e.c), e.n, float64(e.n)*100/float64(total))
	}

	// Legend: the most common colour is the background; every other frequent
	// colour gets a rune, assigned in descending count so the letters read as
	// "most ink first".
	legend := map[uint32]rune{}
	var order []uint32
	runes := []rune("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghij")
	for _, e := range list {
		if len(order) >= len(runes) {
			break
		}
		if float64(e.n)*100/float64(total) < 0.2 {
			break
		}
		legend[e.c] = runes[len(order)]
		order = append(order, e.c)
	}

	const cols, rows = 64, 24
	var sb strings.Builder
	for r := 0; r < rows; r++ {
		for col := 0; col < cols; col++ {
			x0 := rect.Min.X + col*rect.Dx()/cols
			x1 := rect.Min.X + (col+1)*rect.Dx()/cols
			y0 := rect.Min.Y + r*rect.Dy()/rows
			y1 := rect.Min.Y + (r+1)*rect.Dy()/rows
			if x1 <= x0 {
				x1 = x0 + 1
			}
			if y1 <= y0 {
				y1 = y0 + 1
			}
			cell := map[uint32]int{}
			for y := y0; y < y1 && y < rect.Max.Y; y++ {
				for x := x0; x < x1 && x < rect.Max.X; x++ {
					cell[packRGBA8(img.At(x, y))]++
				}
			}
			best, bestN := uint32(0), -1
			for c, n := range cell {
				if n > bestN || (n == bestN && c < best) {
					best, bestN = c, n
				}
			}
			if ch, ok := legend[best]; ok {
				sb.WriteRune(ch)
			} else {
				sb.WriteByte('.')
			}
		}
		sb.WriteByte('\n')
	}
	fmt.Printf("map (%d x %d cells):\n%s", cols, rows, sb.String())
	fmt.Printf("legend: . = colour below 0.2%% threshold\n")
	for _, c := range order {
		fmt.Printf("  %c = %s\n", legend[c], formatRGBA8(c))
	}
}

func packRGBA8(c interface {
	RGBA() (uint32, uint32, uint32, uint32)
}) uint32 {
	r, g, b, a := c.RGBA()
	return uint32(uint8(r>>8))<<24 | uint32(uint8(g>>8))<<16 | uint32(uint8(b>>8))<<8 | uint32(uint8(a>>8))
}

func formatRGBA8(c uint32) string {
	return fmt.Sprintf("#%06X a=%3d", c>>8, c&0xff)
}

// cropAndZoom writes a rectangle of a rendered PNG back out, magnified by an
// integer factor, so small text can be read.
//
// Judging a preview by eye means looking at it at whatever size the viewer
// chooses, and a display that downscales a 1600px render to 1080px smears 11pt
// text into something that reads as doubled or blurred — an artefact of the
// viewing, not of the render. Zooming the region of interest first removes that
// ambiguity.
//
//	go run ./cmd/debug_preview --crop <in.png> <out.png> [x y w h] [zoom]
func cropAndZoom(args []string) {
	if len(args) < 2 {
		fmt.Fprintln(os.Stderr, "usage: --crop <in.png> <out.png> [x y w h] [zoom]")
		os.Exit(2)
	}
	f, err := os.Open(args[0])
	if err != nil {
		fmt.Printf("open: %v\n", err)
		return
	}
	img, err := png.Decode(f)
	f.Close()
	if err != nil {
		fmt.Printf("decode: %v\n", err)
		return
	}
	b := img.Bounds()
	rect, zoom := b, 1
	if len(args) >= 6 {
		var vals [4]int
		bad := false
		for i := range vals {
			v, err := strconv.Atoi(args[2+i])
			if err != nil {
				bad = true
				break
			}
			vals[i] = v
		}
		if bad {
			fmt.Fprintln(os.Stderr, "usage: --crop <in.png> <out.png> [x y w h] [zoom]")
			os.Exit(2)
		}
		rect = image.Rect(b.Min.X+vals[0], b.Min.Y+vals[1],
			b.Min.X+vals[0]+vals[2], b.Min.Y+vals[1]+vals[3]).Intersect(b)
	}
	if len(args) >= 7 {
		if v, err := strconv.Atoi(args[6]); err == nil && v > 0 {
			zoom = v
		}
	}
	if rect.Empty() {
		fmt.Printf("region %v is outside %v\n", rect, b)
		return
	}
	out := image.NewRGBA(image.Rect(0, 0, rect.Dx()*zoom, rect.Dy()*zoom))
	for y := 0; y < out.Bounds().Dy(); y++ {
		for x := 0; x < out.Bounds().Dx(); x++ {
			out.Set(x, y, img.At(rect.Min.X+x/zoom, rect.Min.Y+y/zoom))
		}
	}
	h, err := os.Create(args[1])
	if err != nil {
		fmt.Printf("create: %v\n", err)
		return
	}
	defer h.Close()
	if err := png.Encode(h, out); err != nil {
		fmt.Printf("encode: %v\n", err)
		return
	}
	fmt.Printf("wrote %s: %v zoomed %dx -> %dx%d\n", args[1], rect, zoom, out.Bounds().Dx(), out.Bounds().Dy())
}
