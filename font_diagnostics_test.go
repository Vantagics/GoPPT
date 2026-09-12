package gopresentation

import (
	"os"
	"path/filepath"
	"strings"
	"sync"
	"testing"

	"golang.org/x/image/font"
	"golang.org/x/image/font/opentype"
)

// emptyFontCache returns a cache with no fonts at all, so font resolution falls
// all the way through to the bitmap fallback.
func emptyFontCache() *FontCache {
	return &FontCache{
		fonts:        make(map[string]*opentype.Font),
		faces:        make(map[fontKey]font.Face),
		measureFaces: make(map[fontKey]font.Face),
		scanned:      true,
	}
}

// pickInstalledFont returns the cache key of some installed font, chosen
// deterministically so test output is stable. Empty when no fonts are found.
func pickInstalledFont(fc *FontCache) string {
	fc.ensureScanned()
	best := ""
	for k := range fc.fonts {
		if best == "" || k < best {
			best = k
		}
	}
	return best
}

const bogusFontName = "Zz No Such Font 12345"

// textSlideWithFont builds a one-slide presentation with a single text run in
// the given font, which is enough to drive font resolution.
func textSlideWithFont(fontName, text string) *Presentation {
	p := New()
	shape := p.GetActiveSlide().CreateRichTextShape()
	shape.BaseShape.SetOffsetX(400000).SetOffsetY(400000)
	shape.BaseShape.SetWidth(8000000).SetHeight(2000000)
	run := shape.CreateTextRun(text)
	run.GetFont().SetName(fontName).SetSize(24)
	return p
}

func fontRenderOptions(fc *FontCache) *RenderOptions {
	opts := DefaultRenderOptions()
	opts.Width = 640
	opts.FontCache = fc
	return opts
}

func TestFontCacheFindFontName(t *testing.T) {
	fc := NewFontCache()

	if got := fc.FindFontName(bogusFontName, false, false); got != "" {
		t.Errorf("FindFontName(%q) = %q, want empty", bogusFontName, got)
	}
	if fc.HasFont(bogusFontName, false, false) {
		t.Errorf("HasFont(%q) = true, want false", bogusFontName)
	}

	installed := pickInstalledFont(fc)
	if installed == "" {
		t.Skip("no fonts installed on this machine")
	}
	if !fc.HasFont(installed, false, false) {
		t.Errorf("HasFont(%q) = false, want true", installed)
	}
}

// TestFontDiagnosticsSilentWhenFontInstalled checks the report stays empty when
// nothing went wrong, so a non-empty report is a reliable signal.
func TestFontDiagnosticsSilentWhenFontInstalled(t *testing.T) {
	fc := NewFontCache()
	installed := pickInstalledFont(fc)
	if installed == "" {
		t.Skip("no fonts installed on this machine")
	}

	diag := NewFontDiagnostics()
	opts := fontRenderOptions(fc)
	opts.FontDiagnostics = diag

	if _, err := textSlideWithFont(installed, "Hello world").SlideToImage(0, opts); err != nil {
		t.Fatalf("SlideToImage: %v", err)
	}
	if usages := diag.Usages(); len(usages) != 0 {
		t.Errorf("Usages() = %v, want none", usages)
	}
	if s := diag.Summary(); s != "" {
		t.Errorf("Summary() = %q, want empty", s)
	}
	if diag.UsedBitmapFallback() {
		t.Error("UsedBitmapFallback() = true, want false")
	}
}

// TestFontDiagnosticsReportsMissingFont is the case that matters for preview:
// a document asking for an unavailable font must be reported, not silently
// rendered with whatever the machine happens to have.
func TestFontDiagnosticsReportsMissingFont(t *testing.T) {
	fc := NewFontCache()
	if pickInstalledFont(fc) == "" {
		t.Skip("no fonts installed on this machine")
	}

	diag := NewFontDiagnostics()
	opts := fontRenderOptions(fc)
	opts.FontDiagnostics = diag

	if _, err := textSlideWithFont(bogusFontName, "Hello world").SlideToImage(0, opts); err != nil {
		t.Fatalf("SlideToImage: %v", err)
	}

	names := diag.MissingNames()
	if len(names) != 1 || !strings.EqualFold(names[0], bogusFontName) {
		t.Fatalf("MissingNames() = %v, want [%s]", names, bogusFontName)
	}
	usages := diag.Usages()
	if len(usages) == 0 {
		t.Fatal("no font usages recorded for an unavailable font")
	}
	for _, u := range usages {
		if u.Kind == FontResolved {
			t.Errorf("unavailable font reported as resolved: %+v", u)
		}
		if !strings.EqualFold(u.Requested, bogusFontName) {
			t.Errorf("Requested = %q, want %q", u.Requested, bogusFontName)
		}
	}
	if s := diag.Summary(); !strings.Contains(s, bogusFontName) {
		t.Errorf("Summary() = %q, should mention %q", s, bogusFontName)
	}
}

// TestConfiguredFontFallbackChain verifies a caller-supplied chain is honoured
// ahead of the built-in one, which is how a project pins CJK text to a font it
// actually ships or installs.
func TestConfiguredFontFallbackChain(t *testing.T) {
	fc := NewFontCache()
	installed := pickInstalledFont(fc)
	if installed == "" {
		t.Skip("no fonts installed on this machine")
	}

	diag := NewFontDiagnostics()
	opts := fontRenderOptions(fc)
	opts.FontDiagnostics = diag
	opts.FontFallback = []string{"Zz Also Missing", installed}

	if _, err := textSlideWithFont(bogusFontName, "Hello world").SlideToImage(0, opts); err != nil {
		t.Fatalf("SlideToImage: %v", err)
	}

	subs := diag.Substituted()
	if len(subs) == 0 {
		t.Fatalf("no substitution recorded; usages = %v", diag.Usages())
	}
	for _, u := range subs {
		if !strings.EqualFold(u.Used, installed) {
			t.Errorf("Used = %q, want the configured fallback %q", u.Used, installed)
		}
	}
}

// TestOnFontFallbackCallback verifies the callback fires with the same events
// the report collects.
func TestOnFontFallbackCallback(t *testing.T) {
	fc := NewFontCache()
	if pickInstalledFont(fc) == "" {
		t.Skip("no fonts installed on this machine")
	}

	var mu sync.Mutex
	var got []FontUsage
	opts := fontRenderOptions(fc)
	opts.OnFontFallback = func(u FontUsage) {
		mu.Lock()
		got = append(got, u)
		mu.Unlock()
	}

	if _, err := textSlideWithFont(bogusFontName, "Hello world").SlideToImage(0, opts); err != nil {
		t.Fatalf("SlideToImage: %v", err)
	}

	mu.Lock()
	defer mu.Unlock()
	if len(got) == 0 {
		t.Fatal("OnFontFallback was never called for an unavailable font")
	}
	for _, u := range got {
		if u.Kind == FontResolved {
			t.Errorf("callback reported a resolved font: %+v", u)
		}
	}
}

// TestFontDiagnosticsMissingIsObservable checks the reported failure mode that
// produces tofu: when no installed font matches, the request must surface as
// missing rather than as a silent, plausible-looking render.
func TestFontDiagnosticsMissingIsObservable(t *testing.T) {
	r := &renderer{
		fontCache:    emptyFontCache(),
		fontFallback: defaultFontFallbackChain,
		cjkFallback:  defaultCJKFallbackChain,
		fontDiag:     NewFontDiagnostics(),
		scaleX:       1,
	}
	f := NewFont()
	f.SetName("SimSun").SetSize(18)

	face, used, kind := r.resolveFace(f, r.fontSizePixels(f), false)
	if face != nil || used != "" || kind != FontMissing {
		t.Errorf("resolveFace with an empty cache = (%v,%q,%v), want (nil,\"\",FontMissing)", face, used, kind)
	}
	if r.getFace(f) == nil {
		t.Error("getFace should still return the built-in bitmap face")
	}
	if !r.fontDiag.UsedBitmapFallback() {
		t.Error("UsedBitmapFallback() = false, want true")
	}
	if names := r.fontDiag.MissingNames(); len(names) != 1 || names[0] != "SimSun" {
		t.Errorf("MissingNames() = %v, want [SimSun]", names)
	}
}

func TestMergeFallbackChain(t *testing.T) {
	got := mergeFallbackChain([]string{" A ", "", "a", "B"}, []string{"b", "C"})
	want := []string{"A", "B", "C"}
	if len(got) != len(want) {
		t.Fatalf("mergeFallbackChain = %v, want %v", got, want)
	}
	for i := range want {
		if got[i] != want[i] {
			t.Fatalf("mergeFallbackChain = %v, want %v", got, want)
		}
	}
	if got := mergeFallbackChain(nil, want); len(got) != len(want) {
		t.Errorf("mergeFallbackChain(nil, ...) = %v, want the builtin list", got)
	}
}

func TestFontDiagnosticsConcurrentRecord(t *testing.T) {
	diag := NewFontDiagnostics()
	var wg sync.WaitGroup
	for i := 0; i < 32; i++ {
		wg.Add(1)
		go func(i int) {
			defer wg.Done()
			diag.record(FontUsage{
				Requested: "Font" + string(rune('A'+i%4)),
				Kind:      FontMissing,
			})
		}(i)
	}
	wg.Wait()
	if got := len(diag.Usages()); got != 4 {
		t.Errorf("Usages() length = %d, want 4 distinct entries", got)
	}
}

// firstInstalledFontData returns the raw bytes of an installed font file so a
// test can register it in a hand-built cache. It skips when none is readable.
func firstInstalledFontData(t *testing.T) []byte {
	t.Helper()
	for _, dir := range systemFontDirs() {
		entries, err := os.ReadDir(dir)
		if err != nil {
			continue
		}
		for _, entry := range entries {
			if entry.IsDir() || !strings.HasSuffix(strings.ToLower(entry.Name()), ".ttf") {
				continue
			}
			info, err := entry.Info()
			if err != nil || info.Size() > maxFontFileSize {
				continue
			}
			data, err := os.ReadFile(filepath.Join(dir, entry.Name()))
			if err != nil {
				continue
			}
			if _, err := opentype.Parse(data); err == nil {
				return data
			}
		}
	}
	t.Skip("no readable .ttf font file on this machine")
	return nil
}

// TestNegativeCacheDoesNotChangeResolution checks the miss cache is purely an
// optimisation: lookups must return exactly what an uncached lookup would, and a
// miss for one name must not affect any other name.
func TestNegativeCacheDoesNotChangeResolution(t *testing.T) {
	fc := NewFontCache()
	for i := 0; i < 3; i++ {
		if face := fc.GetFace(bogusFontName, 14, false, false); face != nil {
			t.Fatalf("GetFace(%q) returned a face on iteration %d", bogusFontName, i)
		}
	}
	if fc.HasFont(bogusFontName, false, false) {
		t.Errorf("HasFont(%q) = true for an absent font", bogusFontName)
	}

	real := pickInstalledFont(fc)
	if real == "" {
		t.Skip("no system fonts installed")
	}
	if face := fc.GetFace(real, 14, false, false); face == nil {
		t.Errorf("GetFace(%q) = nil after recording an unrelated miss", real)
	}
	// The render and measure paths share the miss set, so a hit must stay a hit
	// on both.
	if face := fc.GetMeasureFace(real, 14, false, false); face == nil {
		t.Errorf("GetMeasureFace(%q) = nil after recording an unrelated miss", real)
	}
	// A shared set also means a name known to be absent misses on both paths,
	// even the one that has never been asked for it before.
	if face := fc.GetMeasureFace(bogusFontName, 14, false, false); face != nil {
		t.Errorf("GetMeasureFace(%q) returned a face for an absent font", bogusFontName)
	}
}

// TestFontCacheNegativeCacheIsInvalidatedByLoadFont guards the one way the miss
// cache could go wrong: remembering a failure is only safe until something new
// is registered, so a font loaded after a failed lookup must become visible.
func TestFontCacheNegativeCacheIsInvalidatedByLoadFont(t *testing.T) {
	data := firstInstalledFontData(t)
	const name = "Zz Deliberately Unregistered Font"

	fc := emptyFontCache()
	if face := fc.GetFace(name, 12, false, false); face != nil {
		t.Fatalf("lookup of %q unexpectedly succeeded", name)
	}
	// The second lookup takes the cached-miss path.
	if face := fc.GetFace(name, 12, false, false); face != nil {
		t.Fatalf("cached miss returned a face for %q", name)
	}
	if err := fc.LoadFontData(name, data); err != nil {
		t.Fatalf("LoadFontData: %v", err)
	}
	if face := fc.GetFace(name, 12, false, false); face == nil {
		t.Error("a font loaded after a miss stayed invisible: the negative cache was not invalidated")
	}

	// The measure-face path keeps its own map, so it needs the same guarantee.
	mc := emptyFontCache()
	if face := mc.GetMeasureFace(name, 12, false, false); face != nil {
		t.Fatalf("measure lookup of %q unexpectedly succeeded", name)
	}
	if err := mc.LoadFontData(name, data); err != nil {
		t.Fatalf("LoadFontData: %v", err)
	}
	if face := mc.GetMeasureFace(name, 12, false, false); face == nil {
		t.Error("measure-face path: a font loaded after a miss stayed invisible")
	}
}

// TestFontCacheConcurrentUse backs the documented "safe for concurrent use"
// guarantee: a pool of renders sharing one cache must not race or corrupt
// results. Run under -race this covers both the face caches and the miss set.
func TestFontCacheConcurrentUse(t *testing.T) {
	fc := NewFontCache()
	real := pickInstalledFont(fc)
	if real == "" {
		t.Skip("no system fonts installed")
	}

	var wg sync.WaitGroup
	for i := 0; i < 16; i++ {
		wg.Add(1)
		go func(i int) {
			defer wg.Done()
			// Half the workers ask for an installed font, half for an absent
			// one, so the read, write and miss paths are all exercised.
			name := real
			if i%2 == 1 {
				name = bogusFontName
			}
			for j := 0; j < 20; j++ {
				size := float64(10 + j)
				face := fc.GetFace(name, size, false, false)
				if name == real && face == nil {
					t.Errorf("GetFace(%q, %v) = nil for an installed font", real, size)
					return
				}
				fc.GetMeasureFace(name, size, false, false)
			}
		}(i)
	}
	wg.Wait()
}

// TestManualFontRegistrationSurvivesTheDirectoryScan pins the precedence
// between the two ways a font enters a cache.
//
// A caller who supplies a font file is asking for that specific build of it, so
// the registration must outrank a later directory scan. The order matters
// because the scan is lazy: it runs on the first face lookup, which is normally
// the first render — long after startup code has registered its fonts. A scan
// that overwrote the registration would silently swap in whatever the machine
// has installed under the same name, which is exactly the non-determinism
// bundling a font is meant to remove.
func TestManualFontRegistrationSurvivesTheDirectoryScan(t *testing.T) {
	scanned, supplied := twoDistinctFontDatas(t)
	const name = "zzollision" // the base name the scan will register `scanned` under

	dir := t.TempDir()
	if err := os.WriteFile(filepath.Join(dir, name+".ttf"), scanned, 0o644); err != nil {
		t.Fatalf("write font fixture: %v", err)
	}

	fc := NewFontCache(dir)
	if err := fc.LoadFontData(name, supplied); err != nil {
		t.Fatalf("LoadFontData: %v", err)
	}
	// Nothing has scanned the directory yet; this lookup is what triggers it.
	face := fc.GetFace(name, 20, false, false)
	if face == nil {
		t.Fatal("lookup of the registered font returned no face")
	}

	want := faceMetricsFor(t, name, supplied)
	if got := face.Metrics(); got != want {
		t.Errorf("the directory scan replaced the supplied font: metrics = %+v, want %+v (the supplied font)", got, want)
	}
}

// faceMetricsFor registers data in a throwaway cache and returns the metrics of
// the face a lookup then produces. Two different fonts give different metrics at
// the same size, so a metrics comparison says which font a lookup resolved to.
func faceMetricsFor(t *testing.T, name string, data []byte) font.Metrics {
	t.Helper()
	fc := emptyFontCache()
	if err := fc.LoadFontData(name, data); err != nil {
		t.Fatalf("LoadFontData: %v", err)
	}
	face := fc.GetFace(name, 20, false, false)
	if face == nil {
		t.Fatal("a registered font did not produce a face")
	}
	return face.Metrics()
}

// twoDistinctFontDatas returns two installed font files whose metrics differ at
// the same size, so a test can tell them apart. It skips when the machine does
// not offer such a pair.
func twoDistinctFontDatas(t *testing.T) ([]byte, []byte) {
	t.Helper()
	const wantMetrics = 4 // metric variants to collect before comparing

	var candidates [][]byte
	var metrics []font.Metrics
	for _, dir := range systemFontDirs() {
		entries, err := os.ReadDir(dir)
		if err != nil {
			continue
		}
		for _, entry := range entries {
			if len(candidates) >= wantMetrics {
				break
			}
			if entry.IsDir() || !strings.HasSuffix(strings.ToLower(entry.Name()), ".ttf") {
				continue
			}
			info, err := entry.Info()
			if err != nil || info.Size() > maxFontFileSize {
				continue
			}
			data, err := os.ReadFile(filepath.Join(dir, entry.Name()))
			if err != nil {
				continue
			}
			if _, err := opentype.Parse(data); err != nil {
				continue
			}
			m := faceMetricsFor(t, "zz candidate", data)
			for i, seen := range metrics {
				if seen != m {
					return candidates[i], data
				}
			}
			candidates = append(candidates, data)
			metrics = append(metrics, m)
		}
	}
	t.Skip("no two installed fonts with different metrics on this machine")
	return nil, nil
}

// BenchmarkFontLookupMiss measures the cost of repeatedly asking for a font
// that is not installed, which is what a preview pays on a machine lacking the
// document's fonts. Each sub-benchmark uses one name and style, so the only
// variable is how expensive an unresolved lookup is.
//
// Measured on an AMD Ryzen 7 8745HS (Windows, with the font scan warmed), with
// and without the miss cache:
//
//	plain        169 ns  48 B  2 allocs  ->  ~86 ns  24 B  1 alloc
//	bold         241 ns  48 B  2 allocs  ->  ~109 ns 24 B  1 alloc
//	bold italic  524 ns  96 B  3 allocs  ->  ~91 ns  24 B  1 alloc
//
// The saving grows with the style because bold/italic probes several
// name-suffix variants (arialbd, arialbi, ariali, ...) per miss.
func BenchmarkFontLookupMiss(b *testing.B) {
	cases := []struct {
		name         string
		bold, italic bool
	}{
		{"plain", false, false},
		{"bold", true, false},
		{"bold-italic", true, true},
	}
	for _, c := range cases {
		b.Run(c.name, func(b *testing.B) {
			fc := NewFontCache()
			// Warm the directory scan: it parses every installed font file and
			// would otherwise dominate the measurement as a one-off cost.
			fc.ensureScanned()
			fc.GetFace("Zz Warmup Font", 14, false, false)
			b.ReportAllocs()
			b.ResetTimer()
			for i := 0; i < b.N; i++ {
				if face := fc.GetFace(bogusFontName, 14, c.bold, c.italic); face != nil {
					b.Fatal("lookup unexpectedly succeeded")
				}
			}
		})
	}
}
