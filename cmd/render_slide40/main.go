package main

import (
"archive/zip"
"fmt"
"image"
"image/color"
"image/png"
"io"
"os"

gp "github.com/VantageDataChat/GoPPT"
)

func main() {
// Decode and analyze both EMF files
zr, _ := zip.OpenReader("test.pptx")
defer zr.Close()
for _, zf := range zr.File {
if zf.Name == "ppt/media/image95.emf" || zf.Name == "ppt/media/image96.emf" {
rc, _ := zf.Open()
data, _ := io.ReadAll(rc)
rc.Close()
name := zf.Name[len("ppt/media/"):]

img := gp.DecodeEMFForTest(data)
if img == nil {
fmt.Printf("%s: nil\n", name)
continue
}
b := img.Bounds()
fmt.Printf("%s: %dx%d\n", name, b.Dx(), b.Dy())

// Analyze colors
colorMap := map[color.RGBA]int{}
transparent := 0
for y := b.Min.Y; y < b.Max.Y; y++ {
for x := b.Min.X; x < b.Max.X; x++ {
cr, cg, cb, ca := img.At(x, y).RGBA()
c := color.RGBA{uint8(cr >> 8), uint8(cg >> 8), uint8(cb >> 8), uint8(ca >> 8)}
if c.A == 0 {
transparent++
} else {
colorMap[c]++
}
}
}
fmt.Printf("  transparent=%d, unique colors=%d\n", transparent, len(colorMap))
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
fmt.Println("  Top colors:")
for i, c := range sorted {
if i >= 15 { break }
fmt.Printf("    RGBA(%d,%d,%d,%d): %d pixels\n", c.c.R, c.c.G, c.c.B, c.c.A, c.n)
}

// Save
outName := fmt.Sprintf("_debug_xml/%s.png", name[:len(name)-4])
out, _ := os.Create(outName)
png.Encode(out, img)
out.Close()
}
}

// Also render slide 40 and crop the icon area
reader, _ := gp.NewReader(gp.ReaderPowerPoint2007)
pres, _ := reader.Read("test.pptx")
opts := gp.DefaultRenderOptions()
opts.Width = 1920
pres.SaveSlideAsImage(39, "slide40_current.png", opts)

f, _ := os.Open("slide40_current.png")
defer f.Close()
rendered, _ := png.Decode(f)
rb := rendered.Bounds()
w, h := rb.Dx(), rb.Dy()

// Crop bottom icon area (y 60%-85%)
startY := h * 58 / 100
endY := h * 88 / 100
crop := image.NewRGBA(image.Rect(0, 0, w, endY-startY))
for y := startY; y < endY; y++ {
for x := 0; x < w; x++ {
cr, cg, cb, ca := rendered.At(x, y).RGBA()
off := (y-startY)*crop.Stride + x*4
crop.Pix[off] = uint8(cr >> 8)
crop.Pix[off+1] = uint8(cg >> 8)
crop.Pix[off+2] = uint8(cb >> 8)
crop.Pix[off+3] = uint8(ca >> 8)
}
}
out, _ := os.Create("_debug_xml/icon_area_crop.png")
png.Encode(out, crop)
out.Close()
fmt.Println("\nSaved icon area crop to _debug_xml/icon_area_crop.png")
}