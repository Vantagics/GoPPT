// Command zz_cmp compares two directories of slide renders pixel by pixel and
// says where they disagree, so a fidelity gap can be located before it is
// explained.
//
//	go run ./cmd/zz_cmp <dirA> <dirB> <outdir> [tolerance]
//
// dirA is the render under test, dirB the reference (PowerPoint's own export).
// Files are paired by name. For every pair it writes three things: a heat map
// (the reference dimmed, disagreement in red), a per-slide row of numbers, and
// a self-contained HTML sheet with the two renders side by side.
//
// Pixel equality is not the goal and is not attainable: PowerPoint and this
// library antialias text differently and resample pictures differently. What
// the numbers are for is *location* - which region of which slide, and how
// much of it - so a gap can be attributed to a shape or a subsystem rather
// than to "the render is slightly different".
package main

import (
	"fmt"
	"image"
	"image/color"
	"image/png"
	"os"
	"path/filepath"
	"sort"
	"strings"
)

const gridX, gridY = 16, 12

type result struct {
	name     string
	diffPct  float64
	meanDiff float64
	bbox     image.Rectangle
	w, h     int
	grid     [gridY][gridX]float64
	sizeErr  string
}

func main() {
	if len(os.Args) < 4 {
		fmt.Fprintln(os.Stderr, "usage: zz_cmp <dirA> <dirB> <outdir> [tolerance]")
		os.Exit(2)
	}
	dirA, dirB, outDir := os.Args[1], os.Args[2], os.Args[3]
	tol := 48.0
	if len(os.Args) > 4 {
		fmt.Sscanf(os.Args[4], "%f", &tol)
	}
	if err := os.MkdirAll(outDir, 0o755); err != nil {
		fmt.Fprintln(os.Stderr, "mkdir:", err)
		os.Exit(1)
	}

	names, _ := filepath.Glob(filepath.Join(dirA, "*.png"))
	sort.Strings(names)

	var res []result
	var html strings.Builder
	html.WriteString("<!doctype html><meta charset=utf-8><title>render comparison</title>\n")
	html.WriteString("<style>body{font:13px/1.5 system-ui;margin:16px;background:#fff;color:#111}" +
		"table{border-collapse:collapse;margin:0 0 22px}td,th{border:1px solid #ccc;padding:4px 8px;vertical-align:top}" +
		"th{background:#f2f2f2;text-align:left}img{display:block;background:#fafafa}" +
		".g{font:11px/1.05 monospace;white-space:pre;background:#fafafa;padding:4px;border:1px solid #ddd}" +
		".bad{background:#ffd6d6}.ok{background:#dff0d8}</style>\n")
	fmt.Fprintf(&html, "<h1>render comparison</h1><p>A = %s<br>B = %s<br>tolerance = %.0f</p>\n", dirA, dirB, tol)

	for _, pa := range names {
		base := filepath.Base(pa)
		pb := filepath.Join(dirB, base)
		a, errA := load(pa)
		b, errB := load(pb)
		if errA != nil || errB != nil {
			res = append(res, result{name: base, sizeErr: fmt.Sprintf("load A=%v B=%v", errA, errB)})
			continue
		}
		if !a.Bounds().Eq(b.Bounds()) {
			res = append(res, result{name: base, sizeErr: fmt.Sprintf("size %v vs %v", a.Bounds(), b.Bounds())})
			continue
		}
		r := compare(a, b, tol)
		r.name = base
		res = append(res, r)

		heatPath := filepath.Join(outDir, "heat_"+base)
		writeHeat(heatPath, a, b, tol)
		writeHTMLRow(&html, base, r, dirA, dirB, outDir)
	}

	// Summary, worst first: the slide to look at is the one that disagrees most.
	sort.Slice(res, func(i, j int) bool { return res[i].diffPct > res[j].diffPct })
	fmt.Printf("%-14s %7s %7s  %-24s %s\n", "slide", "diff%", "mean", "differing region", "worst cells")
	for _, r := range res {
		if r.sizeErr != "" {
			fmt.Printf("%-14s %s\n", r.name, r.sizeErr)
			continue
		}
		worst := worstCells(r.grid, 4)
		fmt.Printf("%-14s %7.2f %7.2f  %-24s %s\n", r.name, r.diffPct, r.meanDiff, boxPct(r), worst)
	}

	html.WriteString("<h2>all slides, worst first</h2><table><tr><th>slide</th><th>diff%</th><th>mean</th><th>bbox</th><th>grid</th></tr>\n")
	for _, r := range res {
		cls := "ok"
		if r.diffPct > 5 {
			cls = "bad"
		}
		if r.sizeErr != "" {
			fmt.Fprintf(&html, "<tr><td>%s</td><td colspan=4>%s</td></tr>\n", r.name, r.sizeErr)
			continue
		}
		fmt.Fprintf(&html, "<tr class=%q><td>%s</td><td>%.2f</td><td>%.2f</td><td>%v</td><td class=g>%s</td></tr>\n",
			cls, r.name, r.diffPct, r.meanDiff, r.bbox, gridText(r.grid))
	}
	html.WriteString("</table>\n")
	os.WriteFile(filepath.Join(outDir, "report.html"), []byte(html.String()), 0o644)
	fmt.Printf("\nwrote %s\n", filepath.Join(outDir, "report.html"))
}

func load(path string) (image.Image, error) {
	f, err := os.Open(path)
	if err != nil {
		return nil, err
	}
	defer f.Close()
	img, _, err := image.Decode(f)
	return img, err
}

func maxChanDiff(c1, c2 color.Color) float64 {
	r1, g1, b1, _ := c1.RGBA()
	r2, g2, b2, _ := c2.RGBA()
	d := func(x, y uint32) float64 {
		v := int(x>>8) - int(y>>8)
		if v < 0 {
			return float64(-v)
		}
		return float64(v)
	}
	m := d(r1, r2)
	if v := d(g1, g2); v > m {
		m = v
	}
	if v := d(b1, b2); v > m {
		m = v
	}
	return m
}

func compare(a, b image.Image, tol float64) result {
	bd := a.Bounds()
	w, h := bd.Dx(), bd.Dy()
	var r result
	// Tracked separately, and only set by a real hit: image.Rect normalises its
	// arguments, so seeding with Rect(w,h,-1,-1) silently yields the whole
	// image and every later comparison against it is a no-op.
	minX, minY, maxX, maxY := w, h, -1, -1
	total, sum, over := 0.0, 0.0, 0.0
	var cellHit [gridY][gridX]float64
	var cellAll [gridY][gridX]float64
	for y := 0; y < h; y++ {
		gy := y * gridY / h
		if gy >= gridY {
			gy = gridY - 1
		}
		for x := 0; x < w; x++ {
			gx := x * gridX / w
			if gx >= gridX {
				gx = gridX - 1
			}
			d := maxChanDiff(a.At(bd.Min.X+x, bd.Min.Y+y), b.At(bd.Min.X+x, bd.Min.Y+y))
			total++
			sum += d
			cellAll[gy][gx]++
			if d > tol {
				over++
				cellHit[gy][gx]++
				if x < minX {
					minX = x
				}
				if y < minY {
					minY = y
				}
				if x > maxX {
					maxX = x
				}
				if y > maxY {
					maxY = y
				}
			}
		}
	}
	r.diffPct = over / total * 100
	r.meanDiff = sum / total
	for gy := 0; gy < gridY; gy++ {
		for gx := 0; gx < gridX; gx++ {
			if cellAll[gy][gx] > 0 {
				r.grid[gy][gx] = cellHit[gy][gx] / cellAll[gy][gx] * 100
			}
		}
	}
	r.w, r.h = w, h
	if maxX >= 0 {
		r.bbox = image.Rect(minX, minY, maxX+1, maxY+1)
	}
	return r
}

// boxPct reports the differing region as a fraction of the slide, which is the
// form that survives a change of render width and is comparable across slides.
func boxPct(r result) string {
	if r.bbox.Empty() {
		return "none"
	}
	return fmt.Sprintf("x %d-%d%% y %d-%d%%",
		r.bbox.Min.X*100/r.w, (r.bbox.Max.X-1)*100/r.w,
		r.bbox.Min.Y*100/r.h, (r.bbox.Max.Y-1)*100/r.h)
}

// writeHeat draws B dimmed, with A's disagreement over it in red: the eye finds
// a region far faster in an image than in a grid of numbers.
func writeHeat(path string, a, b image.Image, tol float64) {
	bd := a.Bounds()
	out := image.NewRGBA(bd)
	w, h := bd.Dx(), bd.Dy()
	for y := 0; y < h; y++ {
		for x := 0; x < w; x++ {
			ca := a.At(bd.Min.X+x, bd.Min.Y+y)
			cb := b.At(bd.Min.X+x, bd.Min.Y+y)
			d := maxChanDiff(ca, cb)
			br, bg, bb, _ := cb.RGBA()
			gr := byte(br>>8) * 45 / 100
			gg := byte(bg>>8) * 45 / 100
			gb := byte(bb>>8) * 45 / 100
			if d > tol {
				t := d / 255
				if t > 1 {
					t = 1
				}
				out.Set(bd.Min.X+x, bd.Min.Y+y, color.RGBA{
					R: byte(float64(gr)*(1-t) + 255*t),
					G: byte(float64(gg) * (1 - t)),
					B: byte(float64(gb) * (1 - t)),
					A: 255,
				})
			} else {
				out.Set(bd.Min.X+x, bd.Min.Y+y, color.RGBA{gr, gg, gb, 255})
			}
		}
	}
	f, err := os.Create(path)
	if err != nil {
		return
	}
	defer f.Close()
	_ = png.Encode(f, out)
}

func writeHTMLRow(h *strings.Builder, name string, r result, dirA, dirB, outDir string) {
	fmt.Fprintf(h, "<h2>%s &nbsp; diff %.2f%% &nbsp; mean %.2f</h2>\n", name, r.diffPct, r.meanDiff)
	fmt.Fprintf(h, "<table><tr><th>%s (render under test)</th><th>%s (PowerPoint)</th><th>disagreement</th></tr>\n", name, name)
	fmt.Fprintf(h, "<tr><td><img src=%q width=520></td><td><img src=%q width=520></td><td><img src=%q width=520></td></tr>\n",
		relSrc(outDir, dirA, name),
		relSrc(outDir, dirB, name),
		"heat_"+name)
	fmt.Fprintf(h, "<tr><td colspan=3 class=g>%s</td></tr></table>\n", gridText(r.grid))
}

// relSrc returns the src attribute for an image living in dir, written from a
// page stored in outDir, so the report works no matter where the two input
// directories sit relative to it.
func relSrc(outDir, dir, name string) string {
	absOut, err := filepath.Abs(outDir)
	if err != nil {
		absOut = outDir
	}
	absDir, err := filepath.Abs(dir)
	if err != nil {
		absDir = dir
	}
	rel, err := filepath.Rel(absOut, filepath.Join(absDir, name))
	if err != nil {
		return filepath.ToSlash(filepath.Join(absDir, name))
	}
	return filepath.ToSlash(rel)
}

func gridText(g [gridY][gridX]float64) string {
	var b strings.Builder
	for gy := 0; gy < gridY; gy++ {
		for gx := 0; gx < gridX; gx++ {
			v := g[gy][gx]
			switch {
			case v < 0.5:
				b.WriteString(".")
			case v < 5:
				b.WriteString("-")
			case v < 20:
				b.WriteString("+")
			case v < 50:
				b.WriteString("*")
			default:
				b.WriteString("#")
			}
		}
		b.WriteString("\n")
	}
	return b.String()
}

func worstCells(g [gridY][gridX]float64, n int) string {
	type cell struct {
		v      float64
		gy, gx int
	}
	var cs []cell
	for gy := 0; gy < gridY; gy++ {
		for gx := 0; gx < gridX; gx++ {
			if g[gy][gx] > 0.5 {
				cs = append(cs, cell{g[gy][gx], gy, gx})
			}
		}
	}
	sort.Slice(cs, func(i, j int) bool { return cs[i].v > cs[j].v })
	if len(cs) > n {
		cs = cs[:n]
	}
	var b strings.Builder
	for _, c := range cs {
		fmt.Fprintf(&b, "c%d,r%d=%.0f%% ", c.gx, c.gy, c.v)
	}
	if b.Len() == 0 {
		return "-"
	}
	return strings.TrimSpace(b.String())
}
