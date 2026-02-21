package gopresentation

import (
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

// FontCache manages TrueType font loading and face caching.
// It searches system font directories and user-specified directories
// for .ttf and .otf files, then caches parsed fonts and rendered faces.
type FontCache struct {
	mu           sync.RWMutex
	dirs         []string                  // directories to search for fonts
	fonts        map[string]*opentype.Font // lowercase font name -> parsed font
	faces        map[fontKey]font.Face     // cached render faces (HintingFull)
	measureFaces map[fontKey]font.Face     // cached measure faces (HintingNone)
	scanned      bool
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
	}
}

// GetFace returns a font.Face for the given font properties.
// It tries to find a matching TrueType font; returns nil if not found.
func (fc *FontCache) GetFace(name string, sizePt float64, bold, italic bool) font.Face {
	fc.ensureScanned()

	key := fontKey{name: strings.ToLower(name), size: sizePt, bold: bold, italic: italic}

	fc.mu.RLock()
	if face, ok := fc.faces[key]; ok {
		fc.mu.RUnlock()
		return face
	}
	fc.mu.RUnlock()

	// Try to find the font with style variants
	f := fc.findFont(name, bold, italic)
	if f == nil {
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

	key := fontKey{name: strings.ToLower(name), size: sizePt, bold: bold, italic: italic}

	fc.mu.RLock()
	if face, ok := fc.measureFaces[key]; ok {
		fc.mu.RUnlock()
		return face
	}
	fc.mu.RUnlock()

	f := fc.findFont(name, bold, italic)
	if f == nil {
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

// findFont looks up a parsed font by name, trying style-specific variants first.
func (fc *FontCache) findFont(name string, bold, italic bool) *opentype.Font {
	fc.mu.RLock()
	defer fc.mu.RUnlock()

	lower := strings.ToLower(name)

	// Try style-specific names: Windows uses "arialbd", "arialbi", "ariali" etc.
	if bold && italic {
		for _, suffix := range []string{" bold italic", "bi", " bolditalic", "z"} {
			if f, ok := fc.fonts[lower+suffix]; ok {
				return f
			}
		}
	}
	if bold {
		for _, suffix := range []string{" bold", "bd", "b"} {
			if f, ok := fc.fonts[lower+suffix]; ok {
				return f
			}
		}
	}
	if italic {
		for _, suffix := range []string{" italic", "i", " it"} {
			if f, ok := fc.fonts[lower+suffix]; ok {
				return f
			}
		}
	}

	// Fall back to base name
	if f, ok := fc.fonts[lower]; ok {
		return f
	}

	// Try Chinese font name alias
	if alias, ok := chineseFontAliases[lower]; ok {
		return fc.findFontByKey(alias, bold, italic)
	}

	return nil
}

// findFontByKey looks up a font by its already-lowercased key, with style variants.
func (fc *FontCache) findFontByKey(lower string, bold, italic bool) *opentype.Font {
	if bold && italic {
		for _, suffix := range []string{" bold italic", "bi", " bolditalic", "z"} {
			if f, ok := fc.fonts[lower+suffix]; ok {
				return f
			}
		}
	}
	if bold {
		for _, suffix := range []string{" bold", "bd", "b"} {
			if f, ok := fc.fonts[lower+suffix]; ok {
				return f
			}
		}
	}
	if italic {
		for _, suffix := range []string{" italic", "i", " it"} {
			if f, ok := fc.fonts[lower+suffix]; ok {
				return f
			}
		}
	}
	if f, ok := fc.fonts[lower]; ok {
		return f
	}
	return nil
}


// LoadFont manually loads a TrueType/OpenType font file and registers it under the given name.
// Returns an error if the file exceeds maxFontFileSize.
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
	fc.mu.Lock()
	fc.fonts[strings.ToLower(name)] = f
	fc.registerByFamilyName(f)
	fc.mu.Unlock()
	return nil
}

// LoadFontData registers a TrueType/OpenType font from raw bytes.
func (fc *FontCache) LoadFontData(name string, data []byte) error {
	f, err := opentype.Parse(data)
	if err != nil {
		return err
	}
	fc.mu.Lock()
	fc.fonts[strings.ToLower(name)] = f
	fc.registerByFamilyName(f)
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
		lower := strings.ToLower(name)
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

// registerByFamilyName extracts the font family name from the font's name
// table and registers it in the cache.
func (fc *FontCache) registerByFamilyName(f *opentype.Font) {
	familyName, err := f.Name(nil, sfnt.NameIDFamily)
	if err == nil && familyName != "" {
		fc.fonts[strings.ToLower(familyName)] = f
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
