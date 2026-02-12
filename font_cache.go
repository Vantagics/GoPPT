package gopresentation

import (
	"os"
	"path/filepath"
	"runtime"
	"strings"
	"sync"

	"golang.org/x/image/font"
	"golang.org/x/image/font/opentype"
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
	mu       sync.RWMutex
	dirs     []string            // directories to search for fonts
	fonts    map[string]*opentype.Font // lowercase font name -> parsed font
	faces    map[fontKey]font.Face     // cached faces
	scanned  bool
}

// NewFontCache creates a FontCache that searches the given directories
// plus the OS default font directories.
func NewFontCache(extraDirs ...string) *FontCache {
	dirs := append(systemFontDirs(), extraDirs...)
	return &FontCache{
		dirs:  dirs,
		fonts: make(map[string]*opentype.Font),
		faces: make(map[fontKey]font.Face),
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
		for _, suffix := range []string{" bold", "bd", "b", " bold"} {
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

	return nil
}


// LoadFont manually loads a TrueType/OpenType font file and registers it under the given name.
func (fc *FontCache) LoadFont(name string, path string) error {
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

func (fc *FontCache) scanDir(dir string) {
	entries, err := os.ReadDir(dir)
	if err != nil {
		return
	}
	for _, entry := range entries {
		if entry.IsDir() {
			// Recurse one level into subdirectories (e.g. Windows font families)
			fc.scanDir(filepath.Join(dir, entry.Name()))
			continue
		}
		name := entry.Name()
		lower := strings.ToLower(name)
		if !strings.HasSuffix(lower, ".ttf") && !strings.HasSuffix(lower, ".otf") {
			continue
		}

		path := filepath.Join(dir, name)
		data, err := os.ReadFile(path)
		if err != nil {
			continue
		}
		f, err := opentype.Parse(data)
		if err != nil {
			continue
		}

		// Register by filename without extension, lowercased
		baseName := strings.TrimSuffix(lower, filepath.Ext(lower))
		fc.fonts[baseName] = f
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
