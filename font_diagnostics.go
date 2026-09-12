package gopresentation

import (
	"fmt"
	"sort"
	"strings"
	"sync"
)

// FontFallbackKind describes how one font request was resolved.
type FontFallbackKind int

const (
	// FontResolved means the requested font was installed and used as-is.
	FontResolved FontFallbackKind = iota
	// FontSubstituted means the requested font was missing and a different
	// installed font was used instead.
	FontSubstituted
	// FontMissing means no installed font matched and the built-in bitmap
	// fallback was used. Text rendered this way ignores the requested size and
	// non-ASCII glyphs (CJK in particular) come out as blank boxes, so a
	// preview showing tofu can be traced to a missing font rather than to a
	// problem with the document.
	FontMissing
)

func (k FontFallbackKind) String() string {
	switch k {
	case FontResolved:
		return "resolved"
	case FontSubstituted:
		return "substituted"
	case FontMissing:
		return "missing"
	}
	return fmt.Sprintf("FontFallbackKind(%d)", int(k))
}

// FontUsage records how one font request resolved during a render.
type FontUsage struct {
	// Requested is the font name the document asked for (Font.Name).
	Requested string
	// Used is the font actually used. Empty when Kind is FontMissing, because
	// the built-in bitmap font is not a named font.
	Used string
	// Kind says whether the request was satisfied exactly, substituted, or
	// unsatisfied.
	Kind FontFallbackKind
	// Bold and Italic echo the requested style.
	Bold, Italic bool
}

func (u FontUsage) String() string {
	style := ""
	switch {
	case u.Bold && u.Italic:
		style = " (bold italic)"
	case u.Bold:
		style = " (bold)"
	case u.Italic:
		style = " (italic)"
	}
	if u.Kind == FontMissing {
		return fmt.Sprintf("%s%s -> <missing>", u.Requested, style)
	}
	return fmt.Sprintf("%s%s -> %s", u.Requested, style, u.Used)
}

// FontDiagnostics collects the font requests that could not be resolved exactly
// during a render, so a caller can tell "the document is broken" apart from
// "this machine lacks the font". It is safe for concurrent use.
//
// A single FontDiagnostics may be shared across renders; entries accumulate.
type FontDiagnostics struct {
	mu    sync.Mutex
	seen  map[string]FontUsage
	order []string
}

// NewFontDiagnostics creates an empty report.
func NewFontDiagnostics() *FontDiagnostics {
	return &FontDiagnostics{seen: make(map[string]FontUsage)}
}

func fontUsageKey(u FontUsage) string {
	return fmt.Sprintf("%s|%t|%t", strings.ToLower(u.Requested), u.Bold, u.Italic)
}

func (d *FontDiagnostics) record(u FontUsage) {
	if d == nil {
		return
	}
	key := fontUsageKey(u)
	d.mu.Lock()
	defer d.mu.Unlock()
	if d.seen == nil {
		d.seen = make(map[string]FontUsage)
	}
	if _, ok := d.seen[key]; !ok {
		d.order = append(d.order, key)
	}
	d.seen[key] = u
}

// Usages returns every recorded font usage, in first-seen order.
//
// Only requests that were *not* satisfied exactly are recorded: fonts that
// resolved as asked are skipped, because reporting them would bury the handful
// of rows a caller actually acts on. So every element has Kind either
// FontSubstituted or FontMissing, and a non-empty result always means the
// rendered text differs from what the document asked for.
func (d *FontDiagnostics) Usages() []FontUsage {
	if d == nil {
		return nil
	}
	d.mu.Lock()
	defer d.mu.Unlock()
	out := make([]FontUsage, 0, len(d.order))
	for _, k := range d.order {
		out = append(out, d.seen[k])
	}
	return out
}

func (d *FontDiagnostics) withKind(kind FontFallbackKind) []FontUsage {
	if d == nil {
		return nil
	}
	var out []FontUsage
	for _, u := range d.Usages() {
		if u.Kind == kind {
			out = append(out, u)
		}
	}
	return out
}

// Substituted returns the fonts that were missing but replaced by another
// installed font.
func (d *FontDiagnostics) Substituted() []FontUsage { return d.withKind(FontSubstituted) }

// Missing returns the fonts that could not be satisfied at all and fell back to
// the built-in bitmap font.
func (d *FontDiagnostics) Missing() []FontUsage { return d.withKind(FontMissing) }

// UsedBitmapFallback reports whether any text fell back to the built-in bitmap
// font, which is the condition that makes CJK text render as blank boxes.
func (d *FontDiagnostics) UsedBitmapFallback() bool {
	return len(d.Missing()) > 0
}

// MissingNames returns the distinct requested font names that were not found,
// sorted for stable output.
func (d *FontDiagnostics) MissingNames() []string {
	return distinctNames(append(d.Missing(), d.Substituted()...))
}

func distinctNames(usages []FontUsage) []string {
	set := make(map[string]struct{}, len(usages))
	for _, u := range usages {
		if u.Requested != "" {
			set[u.Requested] = struct{}{}
		}
	}
	out := make([]string, 0, len(set))
	for name := range set {
		out = append(out, name)
	}
	sort.Strings(out)
	return out
}

// Summary renders a one-line description suitable for a log message. It returns
// an empty string when every requested font was found.
func (d *FontDiagnostics) Summary() string {
	if d == nil {
		return ""
	}
	subs := d.Substituted()
	misses := d.Missing()
	if len(subs) == 0 && len(misses) == 0 {
		return ""
	}

	var parts []string
	if names := distinctNames(misses); len(names) > 0 {
		parts = append(parts, "missing fonts (rendered with built-in bitmap fallback): "+strings.Join(names, ", "))
	}
	if len(subs) > 0 {
		lines := make([]string, 0, len(subs))
		for _, u := range subs {
			lines = append(lines, u.String())
		}
		sort.Strings(lines)
		parts = append(parts, "substituted: "+strings.Join(lines, "; "))
	}
	return strings.Join(parts, " | ")
}

// --- fallback chains --------------------------------------------------------

// defaultFontFallbackChain is tried, in order, when a document requests a font
// that is not installed. CJK-capable fonts come first so that East Asian text
// keeps rendering, with a few ubiquitous Latin fonts at the end.
var defaultFontFallbackChain = []string{
	"Microsoft YaHei", "SimSun", "SimHei", "NSimSun",
	"Yu Gothic", "Meiryo", "MS Gothic",
	"Malgun Gothic", "Gulim",
	"Noto Sans CJK SC", "Noto Sans SC", "WenQuanYi Micro Hei",
	"Arial", "Helvetica", "DejaVu Sans",
}

// defaultCJKFallbackChain omits Latin-only fonts. It is used for runs that are
// known to contain CJK characters, where a Latin-only face would only produce
// missing-glyph boxes.
var defaultCJKFallbackChain = []string{
	"Microsoft YaHei", "SimSun", "SimHei", "NSimSun",
	"Yu Gothic", "Meiryo", "MS Gothic",
	"Malgun Gothic", "Gulim",
	"Noto Sans CJK SC", "Noto Sans SC", "WenQuanYi Micro Hei",
}

// mergeFallbackChain puts the caller-supplied fonts in front of the built-in
// list, skipping duplicates and empty entries.
func mergeFallbackChain(configured, builtin []string) []string {
	if len(configured) == 0 {
		return builtin
	}
	out := make([]string, 0, len(configured)+len(builtin))
	seen := make(map[string]struct{}, len(configured)+len(builtin))
	for _, list := range [][]string{configured, builtin} {
		for _, name := range list {
			name = strings.TrimSpace(name)
			if name == "" {
				continue
			}
			key := strings.ToLower(name)
			if _, dup := seen[key]; dup {
				continue
			}
			seen[key] = struct{}{}
			out = append(out, name)
		}
	}
	return out
}
