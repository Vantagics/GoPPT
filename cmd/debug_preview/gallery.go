package main

import (
	"bytes"
	"encoding/base64"
	"fmt"
	"html"
	"image"
	"image/png"
	"os"
	"path/filepath"
	"sort"
	"strconv"
	"strings"

	xdraw "golang.org/x/image/draw"
)

// writeGallery turns a directory of rendered PNGs into a single self-contained
// HTML contact sheet.
//
// The images are embedded as data URIs rather than referenced by relative path.
// A review sheet is opened from a location the generator does not control — a
// preview panel, a file manager, a mail client — and a page whose images resolve
// only when opened from one particular directory is a page that shows up blank,
// which is worse than a larger file. Thumbnails are downscaled first so the
// result stays a few megabytes rather than tens.
//
//	go run ./cmd/debug_preview --gallery <dir> <out.html> [thumb-width] [title]
func writeGallery(args []string) {
	if len(args) < 2 {
		fmt.Fprintln(os.Stderr, "usage: --gallery <dir> <out.html> [thumb-width] [title]")
		os.Exit(2)
	}
	dir, outPath := args[0], args[1]
	thumbW := 1280
	if len(args) >= 3 {
		if v, err := strconv.Atoi(args[2]); err == nil && v > 0 {
			thumbW = v
		}
	}
	title := "Preview render"
	if len(args) >= 4 {
		title = args[3]
	}

	entries, err := os.ReadDir(dir)
	if err != nil {
		fmt.Printf("read dir: %v\n", err)
		return
	}
	var names []string
	for _, e := range entries {
		if e.IsDir() || !strings.HasSuffix(strings.ToLower(e.Name()), ".png") {
			continue
		}
		names = append(names, e.Name())
	}
	// Slide order matters and is not alphabetical: slide2 must precede slide10.
	sort.Slice(names, func(i, j int) bool { return naturalLess(names[i], names[j]) })
	if len(names) == 0 {
		fmt.Printf("no .png files in %s\n", dir)
		return
	}

	var cards strings.Builder
	var totalBytes int
	for _, name := range names {
		path := filepath.Join(dir, name)
		data, err := os.ReadFile(path)
		if err != nil {
			fmt.Printf("read %s: %v\n", path, err)
			continue
		}
		thumb, err := pngThumbnail(data, thumbW)
		if err != nil {
			fmt.Printf("thumbnail %s: %v\n", path, err)
			continue
		}
		totalBytes += len(thumb)
		cards.WriteString(`    <figure class="card">` + "\n")
		fmt.Fprintf(&cards, "      <img alt=%q src=\"data:image/png;base64,%s\">\n",
			html.EscapeString(name), base64.StdEncoding.EncodeToString(thumb))
		label := strings.TrimSuffix(name, filepath.Ext(name))
		label = strings.TrimPrefix(label, "slide")
		fmt.Fprintf(&cards, "      <figcaption><span class=\"n\">%s</span> <span class=\"f\">%s</span></figcaption>\n",
			html.EscapeString(label), html.EscapeString(name))
		cards.WriteString("    </figure>\n")
	}

	page := galleryHTML(title, dir, len(names), thumbW, cards.String())
	if err := os.WriteFile(outPath, []byte(page), 0o644); err != nil {
		fmt.Printf("write %s: %v\n", outPath, err)
		return
	}
	info, _ := os.Stat(outPath)
	mb := float64(info.Size()) / (1 << 20)
	fmt.Printf("wrote %s: %d slides, thumbnails %d px wide, embedded %.1f MB -> %.1f MB html\n",
		outPath, len(names), thumbW, float64(totalBytes)/(1<<20), mb)
}

// pngThumbnail decodes PNG data and returns a PNG re-encoded at width px,
// preserving aspect ratio. Images already narrower than the target are returned
// unchanged so a small render is not needlessly resampled.
func pngThumbnail(data []byte, width int) ([]byte, error) {
	src, err := png.Decode(bytes.NewReader(data))
	if err != nil {
		return nil, err
	}
	b := src.Bounds()
	if b.Dx() <= width {
		return data, nil
	}
	h := b.Dy() * width / b.Dx()
	if h < 1 {
		h = 1
	}
	dst := image.NewRGBA(image.Rect(0, 0, width, h))
	xdraw.CatmullRom.Scale(dst, dst.Bounds(), src, b, xdraw.Over, nil)

	var buf bytes.Buffer
	if err := png.Encode(&buf, dst); err != nil {
		return nil, err
	}
	return buf.Bytes(), nil
}

// naturalLess orders names by their numeric run, so slide2 sorts before slide10.
func naturalLess(a, b string) bool {
	na, sa := splitTrailingNumber(a)
	nb, sb := splitTrailingNumber(b)
	if sa == sb {
		return na < nb
	}
	return sa < sb
}

func splitTrailingNumber(name string) (int, string) {
	base := strings.TrimSuffix(name, filepath.Ext(name))
	i := len(base)
	for i > 0 && base[i-1] >= '0' && base[i-1] <= '9' {
		i--
	}
	if i == len(base) {
		return 0, base
	}
	n, err := strconv.Atoi(base[i:])
	if err != nil {
		return 0, base
	}
	return n, base[:i]
}

func galleryHTML(title, dir string, count, thumbW int, cards string) string {
	return `<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>` + html.EscapeString(title) + `</title>
<style>
  :root {
    --bg: #f6f7f9; --fg: #16202b; --muted: #5b6b7c;
    --card: #ffffff; --line: #dde3ea; --accent: #0a6cff;
  }
  @media (prefers-color-scheme: dark) {
    :root {
      --bg: #12171d; --fg: #e8edf3; --muted: #93a3b4;
      --card: #1a212a; --line: #2b3641; --accent: #62a3ff;
    }
  }
  * { box-sizing: border-box; }
  body {
    margin: 0; padding: 32px 28px 56px;
    background: var(--bg); color: var(--fg);
    font: 15px/1.6 -apple-system, "Segoe UI", "Microsoft YaHei", system-ui, sans-serif;
  }
  header { max-width: 1180px; margin: 0 auto 26px; }
  h1 { margin: 0 0 6px; font-size: 21px; letter-spacing: .2px; }
  .meta { color: var(--muted); font-size: 13.5px; }
  .meta code {
    background: var(--card); border: 1px solid var(--line);
    border-radius: 5px; padding: 1px 6px; font-size: 12.5px;
  }
  .grid {
    max-width: 1180px; margin: 0 auto;
    display: grid; gap: 22px;
    grid-template-columns: repeat(auto-fill, minmax(430px, 1fr));
  }
  .card {
    margin: 0; background: var(--card);
    border: 1px solid var(--line); border-radius: 10px;
    overflow: hidden; box-shadow: 0 1px 2px rgba(0,0,0,.05);
  }
  .card img { display: block; width: 100%; height: auto; background: #fff; }
  figcaption {
    display: flex; align-items: baseline; gap: 8px;
    padding: 9px 13px; border-top: 1px solid var(--line);
    font-size: 12.5px; color: var(--muted);
  }
  .n { font-weight: 700; color: var(--accent); font-size: 13.5px; }
  .f { font-family: ui-monospace, Consolas, monospace; }
  footer {
    max-width: 1180px; margin: 34px auto 0;
    color: var(--muted); font-size: 12.5px;
  }
</style>
</head>
<body>
<header>
  <h1>` + html.EscapeString(title) + `</h1>
  <div class="meta">
    ` + strconv.Itoa(count) + ` slides &middot; thumbnails ` + strconv.Itoa(thumbW) + ` px wide, embedded so this page opens anywhere &middot; source <code>` + html.EscapeString(dir) + `</code>
  </div>
</header>
<div class="grid">
` + cards + `</div>
<footer>Rendered by GoPPT. Every image is embedded in this file; there are no external references.</footer>
</body>
</html>
`
}
