// Command zz_corpus sweeps a directory of real-world .pptx decks and reports,
// per deck, everything a fidelity round needs before it opens a single pixel:
//
//	zz_corpus <dir> [-render] [-width N] [-out <rendir>] [-v]
//
// For each deck it reports the slide count, a census of the concrete shape
// types used (recursing into groups — that count is how many renderer
// branches the deck actually walks), the shapes the reader could not
// represent, and the pictures that carry no image data.
//
// The last two are the two silent failures worth catching at scale: an
// unresolved r:embed, and a construct that is quietly dropped rather than
// kept as a visible placeholder.
//
//	-render   additionally render every slide and report per-slide failures
//	-width    render width in pixels (default 800; small and fast)
//	-out      root for rendered slides (default <dir>/../corpus_render)
//	-v        list every distinct shape kind and every unsupported label
//
// It is deliberately a scanner and not a judge: it says where the coverage is,
// not whether a render is right. Attribution goes back to cmd/zz_inspect,
// cmd/zz_fill and the XML itself, or to a COM diff against PowerPoint.
package main

import (
	"flag"
	"fmt"
	"os"
	"path/filepath"
	"reflect"
	"sort"
	"strings"

	gopresentation "github.com/Vantagics/GoPPT"
)

type census struct {
	kinds  map[string]int // concrete Go type -> count
	total  int
	pics   int
	noData int // pictures with no image bytes attached
}

func newCensus() *census { return &census{kinds: map[string]int{}} }

func (c *census) walk(shapes []gopresentation.Shape) {
	for _, s := range shapes {
		if s == nil {
			continue
		}
		c.total++
		c.kinds[reflect.TypeOf(s).String()]++
		if d, ok := s.(*gopresentation.DrawingShape); ok {
			c.pics++
			if len(d.GetImageData()) == 0 {
				c.noData++
			}
		}
		if g, ok := s.(*gopresentation.GroupShape); ok {
			c.walk(g.GetShapes())
		}
	}
}

func (c *census) distinct() int { return len(c.kinds) }

// sortedKinds returns the type census ordered by frequency, then by name, so
// two runs of the same corpus are directly comparable.
func (c *census) sortedKinds() []string {
	out := make([]string, 0, len(c.kinds))
	for k := range c.kinds {
		out = append(out, k)
	}
	sort.Slice(out, func(i, j int) bool {
		if c.kinds[out[i]] != c.kinds[out[j]] {
			return c.kinds[out[i]] > c.kinds[out[j]]
		}
		return out[i] < out[j]
	})
	return out
}

func main() {
	dir := flag.String("dir", "", "directory of .pptx decks (or use as the positional argument)")
	render := flag.Bool("render", false, "render every slide as well")
	width := flag.Int("width", 800, "render width in pixels")
	outRoot := flag.String("out", "", "root directory for rendered slides")
	verbose := flag.Bool("v", false, "list shape kinds and unsupported labels")
	// Go's flag package stops at the first non-flag argument, so "dir -v"
	// would silently drop every flag after the directory. Pull the one
	// positional out by hand and hand the flags the rest.
	var positional string
	var flags []string
	for _, a := range os.Args[1:] {
		if !strings.HasPrefix(a, "-") && positional == "" {
			positional = a
			continue
		}
		flags = append(flags, a)
	}
	flag.CommandLine.Parse(flags)

	if *dir == "" {
		*dir = positional
	}
	if *dir == "" {
		fmt.Fprintln(os.Stderr, "usage: zz_corpus [-render] [-width N] [-out <rendir>] [-v] <dir>")
		os.Exit(2)
	}
	if *outRoot == "" {
		*outRoot = filepath.Join(filepath.Dir(strings.TrimRight(*dir, string(filepath.Separator))), "corpus_render")
	}

	entries, err := filepath.Glob(filepath.Join(*dir, "*.pptx"))
	if err != nil || len(entries) == 0 {
		fmt.Fprintf(os.Stderr, "no .pptx files under %s\n", *dir)
		os.Exit(1)
	}
	sort.Strings(entries)

	reader, err := gopresentation.NewReader(gopresentation.ReaderPowerPoint2007)
	if err != nil {
		fmt.Fprintf(os.Stderr, "new reader: %v\n", err)
		os.Exit(1)
	}

	var totalSlides, totalUnsup, totalNoData, totalRendered int
	var failed []string

	for _, path := range entries {
		name := filepath.Base(path)
		pres, err := reader.Read(path)
		if err != nil {
			fmt.Printf("%-46s READ-FAIL %v\n", trunc(name, 46), err)
			failed = append(failed, name)
			continue
		}

		c := newCensus()
		for _, s := range pres.GetAllSlides() {
			if s != nil {
				c.walk(s.GetShapes())
			}
		}
		unsup := pres.UnsupportedShapes()
		n := pres.GetSlideCount()
		totalSlides += n
		totalUnsup += len(unsup)
		totalNoData += c.noData

		layout := pres.GetLayout()
		dim := "?"
		if layout != nil {
			dim = fmt.Sprintf("%dx%d", layout.CX, layout.CY)
		}

		rendered, rerr := 0, ""
		if *render {
			outDir := filepath.Join(*outRoot, strings.TrimSuffix(name, ".pptx"))
			r, e := renderDeck(pres, outDir, *width)
			rendered, rerr = r, e
			totalRendered += r
		}

		fmt.Printf("%-46s slides=%-4d kinds=%-3d shapes=%-6d pics=%-5d noimgdata=%-4d unsup=%-4d %s",
			trunc(name, 46), n, c.distinct(), c.total, c.pics, c.noData, len(unsup), dim)
		if *render {
			fmt.Printf(" rendered=%d", rendered)
			if rerr != "" {
				fmt.Printf(" (%s)", rerr)
			}
		}
		fmt.Println()

		if *verbose {
			for _, k := range c.sortedKinds() {
				fmt.Printf("      %-34s %d\n", k, c.kinds[k])
			}
			labels := map[string]int{}
			for _, u := range unsup {
				labels[u.Label()]++
			}
			var ls []string
			for k := range labels {
				ls = append(ls, k)
			}
			sort.Strings(ls)
			for _, l := range ls {
				fmt.Printf("      UNSUP %-28s %d\n", trunc(l, 28), labels[l])
			}
		}
		pres.Close()
	}

	fmt.Printf("\n%s: %d decks, %d slides, %d unsupported shapes, %d pictures without image data, %d rendered\n",
		filepath.Base(strings.TrimRight(*dir, string(filepath.Separator))), len(entries),
		totalSlides, totalUnsup, totalNoData, totalRendered)
	if len(failed) > 0 {
		fmt.Printf("read failures (%d): %s\n", len(failed), strings.Join(failed, ", "))
	}
}

func renderDeck(pres *gopresentation.Presentation, outDir string, width int) (int, string) {
	if err := os.MkdirAll(outDir, 0o755); err != nil {
		return 0, "mkdir: " + err.Error()
	}
	opts := gopresentation.DefaultRenderOptions()
	opts.Width = width
	n := pres.GetSlideCount()
	ok, firstErr := 0, ""
	for i := 0; i < n; i++ {
		out := filepath.Join(outDir, fmt.Sprintf("slide%02d.png", i+1))
		if err := pres.SaveSlideAsImage(i, out, opts); err != nil {
			if firstErr == "" {
				firstErr = fmt.Sprintf("slide %d: %v", i+1, err)
			}
			continue
		}
		ok++
	}
	if ok != n {
		return ok, fmt.Sprintf("%d/%d ok, first error: %s", ok, n, firstErr)
	}
	return ok, ""
}

func trunc(s string, n int) string {
	r := []rune(s)
	if len(r) <= n {
		return s
	}
	return string(r[:n-1]) + "…"
}
