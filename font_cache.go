package gopresentation

import (
	"encoding/binary"
	"fmt"
	"os"
	"path/filepath"
	"runtime"
	"strings"
	"sync"

	"golang.org/x/image/font"
	"golang.org/x/image/font/opentype"
	"golang.org/x/image/font/sfnt"
)

// fontKey uniquely identifies a font face by name, size, bold, and italic.
type fontKey struct {
	name   string
	size   float64
	bold   bool
	italic bool
}

// glyphCoverageKey identifies a cached "this font has this glyph" answer. Size
// is deliberately absent: glyph coverage is a property of the font file, not of
// the rasterisation size.
type glyphCoverageKey struct {
	name   string
	r      rune
	bold   bool
	italic bool
}

// lowerFontName is strings.ToLower with no allocation when the name is already
// lowercase, which is the common case and sits on the per-text-run path. Names
// containing non-ASCII bytes fall through to strings.ToLower so Unicode case
// folding behaves exactly as before.
func lowerFontName(s string) string {
	for i := 0; i < len(s); i++ {
		c := s[i]
		if c >= 0x80 || (c >= 'A' && c <= 'Z') {
			return strings.ToLower(s)
		}
	}
	return s
}

// FontCache manages TrueType font loading and face caching.
// It searches system font directories and user-specified directories
// for .ttf and .otf files, then caches parsed fonts and rendered faces.
//
// A FontCache is safe for concurrent use, so one cache can back a pool of
// parallel renders. Constructing it scans the font directories, so build it once
// and pass it in RenderOptions.FontCache rather than letting each render make
// its own.
type FontCache struct {
	mu           sync.RWMutex
	dirs         []string                  // directories to search for fonts
	fonts        map[string]*opentype.Font // lowercase font name -> parsed font
	faces        map[fontKey]font.Face     // cached render faces (HintingFull)
	measureFaces map[fontKey]font.Face     // cached measure faces (HintingNone)
	// vitals caches the raw vertical metrics (OS/2 usWinAscent/usWinDescent,
	// head unitsPerEm) per parsed font, read straight from the font bytes at
	// registration time. Go's font.Face.Metrics reports the hhea values
	// instead, and for many fonts those are not what a renderer uses.
	vitals map[*opentype.Font]fontVitals
	// misses caches requests that matched no installed font. Without it a
	// document that names an uninstalled font re-runs the whole style-variant
	// search (and, when nothing matches at all, the entire fallback chain) for
	// every text run on every slide. It is dropped whenever a font is
	// registered, so a font loaded after a failed lookup is still found.
	misses map[fontKey]struct{}
	// coverage caches "does this font have a real glyph for this rune". A font
	// can be installed and resolvable by name yet still lack the characters a
	// document needs, so a successful lookup says only that the *face* exists,
	// not that it can draw the text. The renderer asks this on every CJK run,
	// and a deck repeats the same characters across slides, so the answer is
	// memoised. Dropped whenever a font is registered.
	coverage map[glyphCoverageKey]bool
	scanned  bool
}

// NewFontCache creates a FontCache that searches the given directories
// plus the OS default font directories.
func NewFontCache(extraDirs ...string) *FontCache {
	dirs := append(systemFontDirs(), extraDirs...)
	return &FontCache{
		dirs:         dirs,
		fonts:        make(map[string]*opentype.Font),
		faces:        make(map[fontKey]font.Face),
		measureFaces: make(map[fontKey]font.Face),
		vitals:       make(map[*opentype.Font]fontVitals),
		misses:       make(map[fontKey]struct{}),
		coverage:     make(map[glyphCoverageKey]bool),
	}
}

// GetFace returns a font.Face for the given font properties.
// It tries to find a matching TrueType font; returns nil if not found.
func (fc *FontCache) GetFace(name string, sizePt float64, bold, italic bool) font.Face {
	fc.ensureScanned()

	key := fontKey{name: lowerFontName(name), size: sizePt, bold: bold, italic: italic}

	fc.mu.RLock()
	if face, ok := fc.faces[key]; ok {
		fc.mu.RUnlock()
		return face
	}
	_, miss := fc.misses[key]
	fc.mu.RUnlock()
	if miss {
		return nil
	}

	// Try to find the font with style variants
	f := fc.findFont(name, bold, italic)
	if f == nil {
		fc.rememberMiss(key)
		return nil
	}

	var style opentype.FaceOptions
	style.Size = sizePt
	style.DPI = 72
	style.Hinting = font.HintingFull

	face, err := opentype.NewFace(f, &style)
	if err != nil {
		return nil
	}

	fc.mu.Lock()
	fc.faces[key] = face
	fc.mu.Unlock()
	return face
}

// GetMeasureFace returns a font.Face with HintingNone for text measurement.
// PowerPoint's text layout engine uses unhinted (ideal) glyph metrics for
// line wrapping and text measurement. Using HintingNone produces glyph
// advances that match PowerPoint's DirectWrite layout, so wrapping occurs
// at the same character positions.
func (fc *FontCache) GetMeasureFace(name string, sizePt float64, bold, italic bool) font.Face {
	fc.ensureScanned()

	key := fontKey{name: lowerFontName(name), size: sizePt, bold: bold, italic: italic}

	fc.mu.RLock()
	if face, ok := fc.measureFaces[key]; ok {
		fc.mu.RUnlock()
		return face
	}
	_, miss := fc.misses[key]
	fc.mu.RUnlock()
	if miss {
		return nil
	}

	f := fc.findFont(name, bold, italic)
	if f == nil {
		fc.rememberMiss(key)
		return nil
	}

	face, err := opentype.NewFace(f, &opentype.FaceOptions{
		Size:    sizePt,
		DPI:     72,
		Hinting: font.HintingNone,
	})
	if err != nil {
		return nil
	}

	fc.mu.Lock()
	fc.measureFaces[key] = face
	fc.mu.Unlock()
	return face
}

// rememberMiss records that no installed font matched a request, so the
// style-variant search is not repeated. Only "nothing matched" is cached, never
// a face-construction failure: the render and measure paths build their faces
// with different hinting, and the two share this set, so caching an error from
// one of them could wrongly suppress the other. The set is cleared whenever a
// font is registered, so a font loaded after a failed lookup is still found.
func (fc *FontCache) rememberMiss(key fontKey) {
	fc.mu.Lock()
	if fc.misses == nil {
		fc.misses = make(map[fontKey]struct{})
	}
	fc.misses[key] = struct{}{}
	fc.mu.Unlock()
}

// resetDerivedCachesLocked drops the caches derived from the registered font
// set. It must be called after any font registration: otherwise a font loaded
// after a failed lookup stays invisible for the lifetime of the cache, and a
// glyph-coverage answer computed against the old set keeps suppressing (or
// wrongly permitting) that font. The caller must hold fc.mu.
func (fc *FontCache) resetDerivedCachesLocked() {
	fc.misses = make(map[fontKey]struct{})
	fc.coverage = make(map[glyphCoverageKey]bool)
}

// findFont looks up a parsed font by name, trying style-specific variants first.
func (fc *FontCache) findFont(name string, bold, italic bool) *opentype.Font {
	f, _ := fc.findFontKey(name, bold, italic)
	return f
}

// findFontKey is findFont plus the cache key that actually matched, so callers
// can report which font was substituted. The key is empty when nothing matched.
func (fc *FontCache) findFontKey(name string, bold, italic bool) (*opentype.Font, string) {
	fc.mu.RLock()
	defer fc.mu.RUnlock()
	return fc.findFontKeyLocked(lowerFontName(name), bold, italic)
}

// FindFontName reports which registered font would satisfy a request for name,
// or "" if no installed font matches. It is intended for diagnostics: it lets a
// caller tell whether a document's font is actually available before rendering.
func (fc *FontCache) FindFontName(name string, bold, italic bool) string {
	fc.ensureScanned()
	_, key := fc.findFontKey(name, bold, italic)
	return key
}

// HasFont reports whether a font matching name is available.
func (fc *FontCache) HasFont(name string, bold, italic bool) bool {
	return fc.FindFontName(name, bold, italic) != ""
}

// CoversRune reports whether the font matching name has a real glyph for r.
//
// HasFont answers "is this font installed", which is not the same question as
// "can it draw this text", and the difference is what produces tofu. A face
// that is missing a character does not fail: it silently renders the font's
// .notdef glyph, which for most Latin fonts is an empty rectangle. So a run
// that names a Latin font for East Asian text — exactly what a generator that
// copies one font-family list into both <a:latin> and <a:ea> emits — looks
// perfectly resolved while drawing nothing but boxes.
//
// Glyph index 0 is .notdef, so a zero index means the character is absent.
// Answers are memoised per (font, style, rune); see FontCache.coverage.
func (fc *FontCache) CoversRune(name string, bold, italic bool, r rune) bool {
	if name == "" {
		return false
	}
	fc.ensureScanned()

	key := glyphCoverageKey{name: lowerFontName(name), r: r, bold: bold, italic: italic}
	fc.mu.RLock()
	if covered, ok := fc.coverage[key]; ok {
		fc.mu.RUnlock()
		return covered
	}
	fc.mu.RUnlock()

	covered := false
	if f := fc.findFont(name, bold, italic); f != nil {
		if idx, err := f.GlyphIndex(nil, r); err == nil && idx != 0 {
			covered = true
		}
	}

	fc.mu.Lock()
	if fc.coverage == nil {
		fc.coverage = make(map[glyphCoverageKey]bool)
	}
	fc.coverage[key] = covered
	fc.mu.Unlock()
	return covered
}

// findFontKeyLocked is findFontKey without locking. The caller must hold fc.mu.
func (fc *FontCache) findFontKeyLocked(lower string, bold, italic bool) (*opentype.Font, string) {
	// Try style-specific names: Windows uses "arialbd", "arialbi", "ariali" etc.
	if bold && italic {
		for _, suffix := range []string{" bold italic", "bi", " bolditalic", "z"} {
			if f, ok := fc.fonts[lower+suffix]; ok {
				return f, lower + suffix
			}
		}
	}
	if bold {
		for _, suffix := range []string{" bold", "bd", "b"} {
			if f, ok := fc.fonts[lower+suffix]; ok {
				return f, lower + suffix
			}
		}
	}
	if italic {
		for _, suffix := range []string{" italic", "i", " it"} {
			if f, ok := fc.fonts[lower+suffix]; ok {
				return f, lower + suffix
			}
		}
	}

	// Fall back to base name
	if f, ok := fc.fonts[lower]; ok {
		return f, lower
	}

	// Try Chinese font name alias
	if alias, ok := chineseFontAliases[lower]; ok {
		return fc.findFontKeyLocked(alias, bold, italic)
	}

	return nil, ""
}

// LoadFont manually loads a TrueType/OpenType font file and registers it under the given name.
// Returns an error if the file exceeds maxFontFileSize.
//
// A font registered here takes precedence over one the directory scan finds
// under the same name. Without that ordering the scan, which runs lazily on the
// first face lookup, would silently replace a deliberately supplied font with
// whatever the machine happens to have installed — defeating the point of
// bundling a font to make rendering reproducible. The scan is therefore run
// before the registration, which can make the first LoadFont pay for the
// directory walk that the first render would have triggered anyway.
func (fc *FontCache) LoadFont(name string, path string) error {
	info, err := os.Stat(path)
	if err != nil {
		return err
	}
	if info.Size() > maxFontFileSize {
		return fmt.Errorf("font file too large: %d bytes (max %d)", info.Size(), maxFontFileSize)
	}
	data, err := os.ReadFile(path)
	if err != nil {
		return err
	}
	f, err := opentype.Parse(data)
	if err != nil {
		return err
	}
	fc.ensureScanned()
	fc.mu.Lock()
	fc.fonts[lowerFontName(name)] = f
	fc.recordVitals(f, data, 0)
	fc.registerByFamilyName(f)
	// A newly registered font may satisfy a request that previously missed.
	fc.resetDerivedCachesLocked()
	fc.mu.Unlock()
	return nil
}

// LoadFontData registers a TrueType/OpenType font from raw bytes.
//
// Like LoadFont, a font registered here takes precedence over one found by the
// directory scan under the same name.
func (fc *FontCache) LoadFontData(name string, data []byte) error {
	f, err := opentype.Parse(data)
	if err != nil {
		return err
	}
	fc.ensureScanned()
	fc.mu.Lock()
	fc.fonts[lowerFontName(name)] = f
	fc.recordVitals(f, data, 0)
	fc.registerByFamilyName(f)
	fc.resetDerivedCachesLocked()
	fc.mu.Unlock()
	return nil
}

func (fc *FontCache) ensureScanned() {
	fc.mu.RLock()
	scanned := fc.scanned
	fc.mu.RUnlock()
	if scanned {
		return
	}

	fc.mu.Lock()
	defer fc.mu.Unlock()
	if fc.scanned {
		return
	}
	fc.scanned = true

	for _, dir := range fc.dirs {
		fc.scanDir(dir)
	}
}

// maxFontScanDepth limits recursive directory traversal when scanning for fonts.
const maxFontScanDepth = 3

// maxFontFileSize limits the size of individual font files loaded into memory.
const maxFontFileSize = 20 << 20 // 20 MB

func (fc *FontCache) scanDir(dir string) {
	fc.scanDirDepth(dir, 0)
}

func (fc *FontCache) scanDirDepth(dir string, depth int) {
	if depth > maxFontScanDepth {
		return
	}
	entries, err := os.ReadDir(dir)
	if err != nil {
		return
	}
	for _, entry := range entries {
		if entry.IsDir() {
			fc.scanDirDepth(filepath.Join(dir, entry.Name()), depth+1)
			continue
		}
		name := entry.Name()
		lower := lowerFontName(name)
		isTTC := strings.HasSuffix(lower, ".ttc") || strings.HasSuffix(lower, ".otc")
		isSingle := strings.HasSuffix(lower, ".ttf") || strings.HasSuffix(lower, ".otf")
		if !isTTC && !isSingle {
			continue
		}

		path := filepath.Join(dir, name)

		// Check file size before reading
		info, err := entry.Info()
		if err != nil || info.Size() > maxFontFileSize {
			continue
		}

		data, err := os.ReadFile(path)
		if err != nil {
			continue
		}

		if isTTC {
			fc.loadCollection(data, lower)
		} else {
			fc.loadSingleFont(data, lower)
		}
	}
}

// loadSingleFont parses a single TTF/OTF font and registers it by both
// filename and internal family name.
func (fc *FontCache) loadSingleFont(data []byte, lowerFilename string) {
	f, err := opentype.Parse(data)
	if err != nil {
		return
	}
	baseName := strings.TrimSuffix(lowerFilename, filepath.Ext(lowerFilename))
	fc.fonts[baseName] = f
	fc.recordVitals(f, data, 0)
	// Also register by the font's internal family name
	fc.registerByFamilyName(f)
}

// loadCollection parses a TTC/OTC font collection and registers each font
// by its internal family name.
func (fc *FontCache) loadCollection(data []byte, lowerFilename string) {
	coll, err := opentype.ParseCollection(data)
	if err != nil {
		return
	}
	n := coll.NumFonts()
	for i := 0; i < n; i++ {
		f, err := coll.Font(i)
		if err != nil {
			continue
		}
		// Collection members share the file but carry their own table
		// directory at the i-th offset of the TTC header.
		base := 0
		if len(data) >= 16 {
			if off := int(binary.BigEndian.Uint32(data[12+4*i : 16+4*i])); off > 0 && off < len(data) {
				base = off
			}
		}
		fc.recordVitals(f, data, base)
		// Register first font also by base filename for backward compat
		if i == 0 {
			baseName := strings.TrimSuffix(lowerFilename, filepath.Ext(lowerFilename))
			fc.fonts[baseName] = f
		}
		fc.registerByFamilyName(f)
	}
}

// chineseFontAliases maps Chinese font names to their English equivalents.
// This allows PPTX files that reference fonts by Chinese name to find them
// in the cache where they're registered by English family name.
var chineseFontAliases = map[string]string{
	"宋体":      "simsun",
	"黑体":      "simhei",
	"微软雅黑":    "microsoft yahei",
	"微软雅黑 ui": "microsoft yahei ui",
	"楷体":      "kaiti",
	"仿宋":      "fangsong",
	"新宋体":     "nsimsun",
	"等线":      "dengxian",
	"华文细黑":    "stxihei",
	"华文黑体":    "stheiti",
	"华文楷体":    "stkaiti",
	"华文宋体":    "stsong",
	"华文仿宋":    "stfangsong",
	"华文中宋":    "stzhongsong",
	"方正舒体":    "fzshuti",
	"方正姚体":    "fzyaoti",
	"隶书":      "lisu",
	"幼圆":      "youyuan",
}

// fontVitals holds the raw vertical metrics a renderer needs, read directly
// from the font's tables. upem is head.unitsPerEm; winAsc/winDesc are
// OS/2.usWinAscent/usWinDescent in font units.
type fontVitals struct {
	upem    int
	winAsc  int
	winDesc int
}

// parseFontVitals extracts the vertical metrics from raw sfnt data. base is
// the offset of the font's table directory within data (0 for a single font,
// the TTC inner offset for a collection member). OS/2 is preferred because
// that is what GDI/DirectWrite render from; a font too old to carry one falls
// back to hhea. The OS/2 offsets are version-stable: usWinAscent/usWinDescent
// sit at 74/76 in every OS/2 version.
func parseFontVitals(data []byte, base int) (fontVitals, bool) {
	var v fontVitals
	if len(data) < base+12 {
		return v, false
	}
	numTables := int(binary.BigEndian.Uint16(data[base+4 : base+6]))
	var os2, head, hhea int
	for i := 0; i < numTables; i++ {
		off := base + 12 + 16*i
		if off+16 > len(data) {
			break
		}
		switch string(data[off : off+4]) {
		case "OS/2":
			os2 = int(binary.BigEndian.Uint32(data[off+8 : off+12]))
		case "head":
			head = int(binary.BigEndian.Uint32(data[off+8 : off+12]))
		case "hhea":
			hhea = int(binary.BigEndian.Uint32(data[off+8 : off+12]))
		}
	}
	if head > 0 && head+20 <= len(data) {
		v.upem = int(binary.BigEndian.Uint16(data[head+18 : head+20]))
	}
	if os2 > 0 && os2+78 <= len(data) {
		v.winAsc = int(binary.BigEndian.Uint16(data[os2+74 : os2+76]))
		v.winDesc = int(binary.BigEndian.Uint16(data[os2+76 : os2+78]))
	} else if hhea > 0 && hhea+10 <= len(data) {
		// No OS/2 (ancient font): hhea ascender/descender are the next best
		// approximation of what a rasteriser would use.
		asc := int(int16(binary.BigEndian.Uint16(data[hhea+4 : hhea+6])))
		desc := int(int16(binary.BigEndian.Uint16(data[hhea+6 : hhea+8])))
		if asc > 0 {
			v.winAsc = asc
		}
		if desc < 0 {
			v.winDesc = -desc
		}
	}
	if v.upem <= 0 || v.winAsc <= 0 || v.winDesc < 0 {
		return v, false
	}
	return v, true
}

// recordVitals caches the vertical metrics of a parsed font alongside its raw
// bytes. base is the table-directory offset (0 for single fonts, the TTC inner
// offset for collection members).
func (fc *FontCache) recordVitals(f *opentype.Font, data []byte, base int) {
	if v, ok := parseFontVitals(data, base); ok {
		if fc.vitals == nil {
			fc.vitals = make(map[*opentype.Font]fontVitals)
		}
		fc.vitals[f] = v
	}
}

// WinVerticalMetrics returns the GDI/DirectWrite vertical metrics of the font
// that would render name — OS/2 usWinAscent/usWinDescent — scaled to sizePx
// pixels per em. PowerPoint places the baseline winAscent below the line top
// and advances the line winAscent+winDescent, while Go's font.Face.Metrics
// reports the hhea values; for Calibri those differ by a fifth of an em
// (hhea ascent 0.75em vs win ascent 0.952em), which is a 13px baseline error
// on 28pt text. ok is false when no font matched or its tables said nothing
// usable, in which case the caller should fall back to face.Metrics.
func (fc *FontCache) WinVerticalMetrics(name string, sizePx float64, bold, italic bool) (asc, desc float64, ok bool) {
	if name == "" || sizePx <= 0 {
		return 0, 0, false
	}
	fc.ensureScanned()
	f := fc.findFont(name, bold, italic)
	if f == nil {
		return 0, 0, false
	}
	fc.mu.RLock()
	v, cached := fc.vitals[f]
	fc.mu.RUnlock()
	if !cached {
		return 0, 0, false
	}
	scale := sizePx / float64(v.upem)
	return float64(v.winAsc) * scale, float64(v.winDesc) * scale, true
}

// registerByFamilyName extracts the font family name from the font's name
// table and registers it in the cache.
//
// A family name does not identify a file. calibri.ttf, calibrib.ttf,
// calibrii.ttf and calibriz.ttf all report the family "Calibri" and differ only
// in their subfamily, so registering the family name for each of them leaves
// whichever one the directory walk reached last holding it. On Windows that is
// calibriz.ttf, Calibri Bold Italic, so every request for Calibri found a font —
// findFontKeyLocked returns as soon as the key exists, so nothing reported a
// substitution — and the whole document was drawn bold italic. Worse, the
// winner is decided by directory enumeration order, so the same deck rendered
// differently on different machines and neither matched PowerPoint.
//
// The bare family name is therefore the regular face's, and a styled face is
// registered under the suffixed names findFontKeyLocked already looks for
// ("calibri bold", "calibri bold italic"). A styled face still claims the bare
// name when nothing else has, so that a family with only a bold file installed
// stays reachable rather than resolving to nothing.
func (fc *FontCache) registerByFamilyName(f *opentype.Font) {
	familyName, err := f.Name(nil, sfnt.NameIDFamily)
	if err != nil || familyName == "" {
		return
	}
	lower := strings.ToLower(familyName)

	subfamily, subErr := f.Name(nil, sfnt.NameIDSubfamily)
	subfamily = strings.TrimSpace(subfamily)
	regular := subErr != nil || subfamily == "" || strings.EqualFold(subfamily, "Regular")

	if regular {
		fc.fonts[lower] = f
	} else {
		if _, taken := fc.fonts[lower]; !taken {
			fc.fonts[lower] = f
		}
		fc.fonts[lower+" "+strings.ToLower(subfamily)] = f
	}

	// Also register by full name (e.g. "Microsoft YaHei Bold")
	fullName, err := f.Name(nil, sfnt.NameIDFull)
	if err == nil && fullName != "" {
		fc.fonts[strings.ToLower(fullName)] = f
	}
}

// systemFontDirs returns OS-specific font directories.
func systemFontDirs() []string {
	switch runtime.GOOS {
	case "windows":
		windir := os.Getenv("WINDIR")
		if windir == "" {
			windir = `C:\Windows`
		}
		localAppData := os.Getenv("LOCALAPPDATA")
		dirs := []string{filepath.Join(windir, "Fonts")}
		if localAppData != "" {
			dirs = append(dirs, filepath.Join(localAppData, "Microsoft", "Windows", "Fonts"))
		}
		return dirs
	case "darwin":
		home, _ := os.UserHomeDir()
		dirs := []string{
			"/System/Library/Fonts",
			"/Library/Fonts",
		}
		if home != "" {
			dirs = append(dirs, filepath.Join(home, "Library", "Fonts"))
		}
		return dirs
	default: // linux, freebsd, etc.
		home, _ := os.UserHomeDir()
		dirs := []string{
			"/usr/share/fonts",
			"/usr/local/share/fonts",
		}
		if home != "" {
			dirs = append(dirs, filepath.Join(home, ".local", "share", "fonts"))
			dirs = append(dirs, filepath.Join(home, ".fonts"))
		}
		return dirs
	}
}
