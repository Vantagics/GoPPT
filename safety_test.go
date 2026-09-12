package gopresentation

import (
	"archive/zip"
	"bytes"
	"errors"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

// --- recover mechanics ------------------------------------------------------

// panicky deliberately trips a runtime panic (write to a nil map).
func panicky() {
	var m map[string]int
	m["boom"] = 1
}

func TestRecoverToErrorConvertsPanic(t *testing.T) {
	var err error
	func() {
		defer recoverToError(&err, "testOp")
		panicky()
	}()

	if err == nil {
		t.Fatal("panic was not converted to an error")
	}
	var pe *PanicError
	if !errors.As(err, &pe) {
		t.Fatalf("error type = %T, want *PanicError", err)
	}
	if pe.Op != "testOp" {
		t.Errorf("Op = %q, want %q", pe.Op, "testOp")
	}
	if len(pe.StackTrace()) == 0 {
		t.Error("no stack trace captured")
	}
}

// TestRecoverToErrorKeepsPriorError verifies that a panic during cleanup does
// not hide the error the function had already decided to return.
func TestRecoverToErrorKeepsPriorError(t *testing.T) {
	prior := errors.New("prior failure")
	var err error
	func() {
		defer recoverToError(&err, "testOp")
		err = prior
		panicky()
	}()

	if !errors.Is(err, prior) {
		t.Errorf("errors.Is(err, prior) = false; err = %v", err)
	}
	var pe *PanicError
	if !errors.As(err, &pe) {
		t.Fatalf("error type = %T, want *PanicError", err)
	}
}

// TestRecoverToErrorRepanicsWithoutTarget verifies the guard refuses to swallow
// a panic when there is nowhere to report it.
func TestRecoverToErrorRepanicsWithoutTarget(t *testing.T) {
	defer func() {
		if r := recover(); r == nil {
			t.Error("panic should propagate when errp is nil")
		}
	}()
	func() {
		defer recoverToError(nil, "testOp")
		panicky()
	}()
}

func TestErrIsPanic(t *testing.T) {
	if _, ok := ErrIsPanic(errors.New("plain")); ok {
		t.Error("plain error reported as a panic")
	}
	if _, ok := ErrIsPanic(nil); ok {
		t.Error("nil error reported as a panic")
	}
	v, ok := ErrIsPanic(&PanicError{Op: "x", Value: "boom"})
	if !ok || v != "boom" {
		t.Errorf("ErrIsPanic = (%v, %v), want (boom, true)", v, ok)
	}
}

func TestPanicErrorUnwrap(t *testing.T) {
	sentinel := errors.New("sentinel")
	if !errors.Is(&PanicError{Op: "x", Value: sentinel}, sentinel) {
		t.Error("errors.Is should find an error panic value")
	}
	if !errors.Is(&PanicError{Op: "x", Value: "s", Prior: sentinel}, sentinel) {
		t.Error("errors.Is should find the prior error")
	}
}

// --- public entry points on a nil presentation ------------------------------

func TestNilPresentationEntryPointsDoNotPanic(t *testing.T) {
	var p *Presentation
	dir := t.TempDir()

	if p.GetSlideCount() != 0 {
		t.Error("GetSlideCount on nil presentation should be 0")
	}
	if p.GetAllSlides() != nil {
		t.Error("GetAllSlides on nil presentation should be nil")
	}
	if _, err := p.GetSlide(0); err == nil {
		t.Error("GetSlide on nil presentation = nil error")
	}
	if _, err := p.SlideToImage(0, nil); err == nil {
		t.Error("SlideToImage on nil presentation = nil error")
	}
	if _, err := p.SlidesToImages(nil); err == nil {
		t.Error("SlidesToImages on nil presentation = nil error")
	}
	if err := p.SaveSlideAsImage(0, filepath.Join(dir, "a.png"), nil); err == nil {
		t.Error("SaveSlideAsImage on nil presentation = nil error")
	}
	if err := p.SaveSlidesAsImages(filepath.Join(dir, "s%d.png"), nil); err == nil {
		t.Error("SaveSlidesAsImages on nil presentation = nil error")
	}
	if err := p.Save(filepath.Join(dir, "a.pptx")); err == nil {
		t.Error("Save on nil presentation = nil error")
	}
}

// --- malformed packages -----------------------------------------------------

// zipParts unpacks a package into a name -> content map.
func zipParts(t *testing.T, data []byte) map[string][]byte {
	t.Helper()
	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatalf("open zip: %v", err)
	}
	parts := make(map[string][]byte, len(zr.File))
	for _, f := range zr.File {
		rc, err := f.Open()
		if err != nil {
			t.Fatalf("open %s: %v", f.Name, err)
		}
		buf := new(bytes.Buffer)
		if _, err := buf.ReadFrom(rc); err != nil {
			t.Fatalf("read %s: %v", f.Name, err)
		}
		rc.Close()
		parts[f.Name] = buf.Bytes()
	}
	return parts
}

// buildZip repacks name -> content into a zip archive.
func buildZip(t *testing.T, parts map[string][]byte) []byte {
	t.Helper()
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	for name, content := range parts {
		w, err := zw.Create(name)
		if err != nil {
			t.Fatalf("create %s: %v", name, err)
		}
		if _, err := w.Write(content); err != nil {
			t.Fatalf("write %s: %v", name, err)
		}
	}
	if err := zw.Close(); err != nil {
		t.Fatalf("close zip: %v", err)
	}
	return buf.Bytes()
}

// firstPartMatching returns the name of the first part satisfying pred.
func firstPartMatching(parts map[string][]byte, pred func(string) bool) string {
	for name := range parts {
		if pred(name) {
			return name
		}
	}
	return ""
}

// TestMalformedPackagesDoNotPanic is the regression test for untrusted input:
// whatever the corruption, the public API must either succeed or return an
// error, never panic out of the library.
func TestMalformedPackagesDoNotPanic(t *testing.T) {
	valid := writeToBytes(t, chartSlidePresentation())
	base := zipParts(t, valid)

	if firstPartMatching(base, func(n string) bool { return len(n) > 11 && n[:11] == "ppt/charts/" }) == "" {
		t.Fatal("fixture package has no chart part; the test would not cover chart parsing")
	}

	// mutate returns a copy of the fixture with one corruption applied.
	mutate := func(f func(map[string][]byte)) map[string][]byte {
		clone := make(map[string][]byte, len(base))
		for k, v := range base {
			dup := make([]byte, len(v))
			copy(dup, v)
			clone[k] = dup
		}
		f(clone)
		return clone
	}

	corruptions := []struct {
		name string
		run  func(parts map[string][]byte)
	}{
		{"delete content types", func(m map[string][]byte) { delete(m, "[Content_Types].xml") }},
		{"delete presentation", func(m map[string][]byte) { delete(m, "ppt/presentation.xml") }},
		{"delete all slide rels", func(m map[string][]byte) {
			for n := range m {
				if len(n) > 10 && n[:10] == "ppt/slides" && len(n) > 5 && n[len(n)-5:] == ".rels" {
					delete(m, n)
				}
			}
		}},
		{"garbage presentation xml", func(m map[string][]byte) {
			m["ppt/presentation.xml"] = []byte("<<<not xml at all")
		}},
		{"garbage slide xml", func(m map[string][]byte) {
			n := firstPartMatching(m, func(s string) bool {
				return len(s) > 16 && s[:16] == "ppt/slides/slide" && s[len(s)-4:] == ".xml"
			})
			if n != "" {
				m[n] = []byte("<p:sld><p:cSld><p:spTree>")
			}
		}},
		{"garbage chart xml", func(m map[string][]byte) {
			n := firstPartMatching(m, func(s string) bool { return len(s) > 11 && s[:11] == "ppt/charts/" })
			if n != "" {
				m[n] = []byte("<c:chartSpace><c:plotArea><c:barChart>")
			}
		}},
		{"empty chart xml", func(m map[string][]byte) {
			n := firstPartMatching(m, func(s string) bool { return len(s) > 11 && s[:11] == "ppt/charts/" })
			if n != "" {
				m[n] = nil
			}
		}},
		{"chart rel points at missing part", func(m map[string][]byte) {
			n := firstPartMatching(m, func(s string) bool { return len(s) > 11 && s[:11] == "ppt/charts/" })
			delete(m, n)
		}},
		{"rels contain path traversal", func(m map[string][]byte) {
			for name := range m {
				if len(name) > 5 && name[len(name)-5:] == ".rels" {
					m[name] = []byte(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
						`<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../../../../etc/passwd"/>` +
						`</Relationships>`)
					break
				}
			}
		}},
		{"nul bytes in slide xml", func(m map[string][]byte) {
			n := firstPartMatching(m, func(s string) bool {
				return len(s) > 16 && s[:16] == "ppt/slides/slide" && s[len(s)-4:] == ".xml"
			})
			if n != "" {
				m[n] = bytes.Repeat([]byte{0x00}, 512)
			}
		}},
		{"truncated slide xml", func(m map[string][]byte) {
			n := firstPartMatching(m, func(s string) bool {
				return len(s) > 16 && s[:16] == "ppt/slides/slide" && s[len(s)-4:] == ".xml"
			})
			if n != "" && len(m[n]) > 100 {
				m[n] = m[n][:100]
			}
		}},
		{"oversized attribute values", func(m map[string][]byte) {
			n := firstPartMatching(m, func(s string) bool {
				return len(s) > 16 && s[:16] == "ppt/slides/slide" && s[len(s)-4:] == ".xml"
			})
			if n != "" {
				m[n] = append(m[n], bytes.Repeat([]byte("<a:x "), 5000)...)
			}
		}},
	}

	for _, tc := range corruptions {
		t.Run(tc.name, func(t *testing.T) {
			pkg := buildZip(t, mutate(tc.run))

			// ReadFromReader must not panic.
			pres, err := (&PPTXReader{}).ReadFromReader(bytes.NewReader(pkg), int64(len(pkg)))
			if err == nil && pres != nil {
				// If it did parse, rendering must not panic either.
				opts := DefaultRenderOptions()
				opts.Width = 320
				if _, rerr := pres.SlidesToImages(opts); rerr != nil {
					t.Logf("render returned error (acceptable): %v", rerr)
				}
			}

			// Open must not panic either.
			path := filepath.Join(t.TempDir(), "broken.pptx")
			if werr := os.WriteFile(path, pkg, 0o644); werr != nil {
				t.Fatalf("write fixture: %v", werr)
			}
			if pres2, oerr := Open(path); oerr == nil && pres2 != nil {
				opts := DefaultRenderOptions()
				opts.Width = 320
				_, _ = pres2.SlidesToImages(opts)
			}
		})
	}
}

// --- group nesting ----------------------------------------------------------

// nestedGroupPresentation builds a deck with depth levels of group nested inside
// group, plus a text run at the bottom of the chain.
func nestedGroupPresentation(depth int, marker string) *Presentation {
	p := New()
	group := p.GetActiveSlide().CreateGroupShape()
	group.BaseShape.SetOffsetX(100000).SetOffsetY(100000)
	group.BaseShape.SetWidth(8000000).SetHeight(6000000)

	for i := 1; i < depth; i++ {
		child := NewGroupShape()
		child.BaseShape.SetOffsetX(100000).SetOffsetY(100000)
		child.BaseShape.SetWidth(8000000).SetHeight(6000000)
		group.AddShape(child)
		group = child
	}

	leaf := NewRichTextShape()
	leaf.BaseShape.SetOffsetX(200000).SetOffsetY(200000)
	leaf.BaseShape.SetWidth(4000000).SetHeight(1000000)
	leaf.CreateTextRun(marker)
	group.AddShape(leaf)

	return p
}

// groupDepth returns the depth of the deepest group reachable from shapes.
func groupDepth(shapes []Shape) int {
	deepest := 0
	for _, s := range shapes {
		g, ok := s.(*GroupShape)
		if !ok {
			continue
		}
		if d := 1 + groupDepth(g.shapes); d > deepest {
			deepest = d
		}
	}
	return deepest
}

// treeContainsText reports whether any text run under shapes holds want.
func treeContainsText(shapes []Shape, want string) bool {
	for _, s := range shapes {
		switch v := s.(type) {
		case *GroupShape:
			if treeContainsText(v.shapes, want) {
				return true
			}
		case *RichTextShape:
			if paragraphsContain(v.paragraphs, want) {
				return true
			}
		case *PlaceholderShape:
			if paragraphsContain(v.paragraphs, want) {
				return true
			}
		}
	}
	return false
}

func paragraphsContain(paragraphs []*Paragraph, want string) bool {
	for _, para := range paragraphs {
		for _, elem := range para.elements {
			if tr, ok := elem.(*TextRun); ok && strings.Contains(tr.text, want) {
				return true
			}
		}
	}
	return false
}

// roundTrip writes a presentation to a package and reads it back, which is the
// path a preview tool actually takes.
func roundTrip(t *testing.T, p *Presentation) *Presentation {
	t.Helper()
	pkg := writeToBytes(t, p)
	pres, err := (&PPTXReader{}).ReadFromReader(bytes.NewReader(pkg), int64(len(pkg)))
	if err != nil {
		t.Fatalf("read back: %v", err)
	}
	return pres
}

// TestGroupNestingDepthIsCapped covers the one input that can take the process
// down in a way recover cannot intercept. Nesting depth is chosen by the file,
// and both the renderer and the writer recurse over the resulting tree, so a
// package full of <p:grpSp> elements can exhaust the goroutine stack while
// drawing or saving — a fatal error, not a panic, so the fail-closed boundary
// would not catch it. The reader therefore caps how deep a tree it builds.
//
// The cap is a policy rather than a crash reproduction, so this asserts the
// policy: a package nested past the cap must come back no deeper than it.
func TestGroupNestingDepthIsCapped(t *testing.T) {
	pres := roundTrip(t, nestedGroupPresentation(4*maxGroupDepth, "DEEP-LEAF-MARKER"))
	shapes := pres.GetAllSlides()[0].GetShapes()

	depth := groupDepth(shapes)
	if depth == 0 {
		t.Fatal("the round trip lost every group")
	}
	if depth > maxGroupDepth {
		t.Errorf("group depth = %d, want <= %d: a deeply nested package reached the renderer intact", depth, maxGroupDepth)
	}
	if treeContainsText(shapes, "DEEP-LEAF-MARKER") {
		t.Error("the leaf below the cap survived, so the levels past the cap were not dropped")
	}

	// The trimmed tree must still render and save.
	opts := DefaultRenderOptions()
	opts.Width = 160
	if _, err := pres.SlidesToImages(opts); err != nil {
		t.Errorf("render after trimming: %v", err)
	}
	if err := pres.Save(filepath.Join(t.TempDir(), "trimmed.pptx")); err != nil {
		t.Errorf("save after trimming: %v", err)
	}
}

// TestShallowGroupNestingIsUnaffected is the control for the cap: a depth
// PowerPoint actually produces must survive a round trip intact, leaf text
// included, so the cap is not quietly trimming real documents.
func TestShallowGroupNestingIsUnaffected(t *testing.T) {
	const depth = 5

	pres := roundTrip(t, nestedGroupPresentation(depth, "SHALLOW-LEAF-MARKER"))
	shapes := pres.GetAllSlides()[0].GetShapes()

	if got := groupDepth(shapes); got != depth {
		t.Errorf("group depth = %d, want %d", got, depth)
	}
	if !treeContainsText(shapes, "SHALLOW-LEAF-MARKER") {
		t.Error("the leaf text was lost, so the cap is trimming documents it should leave alone")
	}
}

// TestNonZipInputsDoNotPanic covers inputs that are not packages at all.
func TestNonZipInputsDoNotPanic(t *testing.T) {
	inputs := map[string][]byte{
		"empty":          {},
		"random bytes":   []byte("not a zip file, just some text"),
		"zip magic only": {0x50, 0x4B, 0x03, 0x04},
		"nul bytes":      bytes.Repeat([]byte{0x00}, 1024),
	}
	for name, data := range inputs {
		t.Run(name, func(t *testing.T) {
			if _, err := (&PPTXReader{}).ReadFromReader(bytes.NewReader(data), int64(len(data))); err == nil {
				t.Error("expected an error for non-package input")
			}
			path := filepath.Join(t.TempDir(), "x.pptx")
			if err := os.WriteFile(path, data, 0o644); err != nil {
				t.Fatalf("write: %v", err)
			}
			if _, err := Open(path); err == nil {
				t.Error("expected an error for non-package input")
			}
		})
	}

	// A zip that is structurally fine but is not a presentation.
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	w, _ := zw.Create("hello.txt")
	_, _ = w.Write([]byte("hi"))
	zw.Close()
	if _, err := (&PPTXReader{}).ReadFromReader(bytes.NewReader(buf.Bytes()), int64(buf.Len())); err == nil {
		t.Error("expected an error for a zip that is not a presentation")
	}
}
