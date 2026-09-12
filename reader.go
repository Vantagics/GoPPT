package gopresentation

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"strings"
)

// Reader is the interface for presentation readers.
type Reader interface {
	Read(path string) (*Presentation, error)
	ReadFromReader(r io.ReaderAt, size int64) (*Presentation, error)
}

// ReaderType represents the input format.
type ReaderType string

const (
	ReaderPowerPoint2007 ReaderType = "PowerPoint2007"
)

// NewReader creates a reader for the given format.
func NewReader(format ReaderType) (Reader, error) {
	switch format {
	case ReaderPowerPoint2007:
		return &PPTXReader{}, nil
	default:
		return nil, fmt.Errorf("unsupported reader format: %s", format)
	}
}

// PPTXReader reads PPTX files.
type PPTXReader struct{}

// Read reads a presentation from a file path.
//
// Read never panics on malformed input: a recovered panic is returned as a
// *PanicError.
func (r *PPTXReader) Read(path string) (pres *Presentation, err error) {
	defer recoverToError(&err, "PPTXReader.Read")
	f, err := os.Open(path)
	if err != nil {
		return nil, fmt.Errorf("failed to open file: %w", err)
	}
	defer f.Close()

	info, err := f.Stat()
	if err != nil {
		return nil, fmt.Errorf("failed to stat file: %w", err)
	}

	return r.ReadFromReader(f, info.Size())
}

// ReadFromReader reads a presentation from an io.ReaderAt.
//
// ReadFromReader never panics on malformed input: a recovered panic is returned
// as a *PanicError.
func (r *PPTXReader) ReadFromReader(reader io.ReaderAt, size int64) (pres *Presentation, err error) {
	defer recoverToError(&err, "PPTXReader.ReadFromReader")
	if size <= 0 {
		return nil, fmt.Errorf("invalid reader size: %d", size)
	}
	if size > int64(maxZipInputSize) {
		return nil, fmt.Errorf("file size %d exceeds maximum allowed (%d bytes)", size, maxZipInputSize)
	}

	zr, err := zip.NewReader(reader, size)
	if err != nil {
		return nil, fmt.Errorf("failed to open zip: %w", err)
	}

	if len(zr.File) > maxZipEntries {
		return nil, fmt.Errorf("zip archive contains too many entries (%d > %d)", len(zr.File), maxZipEntries)
	}

	pres = &Presentation{
		properties:             NewDocumentProperties(),
		presentationProperties: NewPresentationProperties(),
		slides:                 make([]*Slide, 0),
		slideMasters:           make([]*SlideMaster, 0),
		layout:                 NewDocumentLayout(),
	}

	// Read core properties (non-fatal: missing properties are acceptable)
	_ = r.readCoreProperties(zr, pres)

	// Read the extended properties and any custom properties (non-fatal)
	_ = r.readAppProperties(zr, pres)
	_ = r.readCustomProperties(zr, pres)

	// Read theme colors (non-fatal)
	r.readThemeColors(zr, pres)

	// Read presentation.xml to get slide list and layout
	slideRels, err := r.readPresentation(zr, pres)
	if err != nil {
		return nil, err
	}

	// Read presentation relationships
	presRels, err := r.readRelationships(zr, "ppt/_rels/presentation.xml.rels")
	if err != nil {
		return nil, err
	}

	// Read the package-level comment author table once; every slide's comments
	// reference it by id. Reading it per slide would repeat the same parse.
	commentAuthors := r.readCommentAuthors(zr)

	// Read slides
	for _, relID := range slideRels {
		target := ""
		for _, rel := range presRels {
			if rel.ID == relID {
				target = rel.Target
				break
			}
		}
		if target == "" {
			continue
		}

		// Normalize path
		if !strings.HasPrefix(target, "ppt/") {
			target = "ppt/" + target
		}

		slide, err := r.readSlide(zr, target, pres, commentAuthors)
		if err != nil {
			return nil, fmt.Errorf("failed to read slide %s: %w", target, err)
		}
		pres.slides = append(pres.slides, slide)
	}

	return pres, nil
}

// maxZipEntrySize is the maximum allowed size for a single file extracted from a ZIP.
// This prevents zip bomb attacks. 50 MB is generous for any legitimate PPTX part.
const maxZipEntrySize = 50 << 20 // 50 MB

// maxZipInputSize caps the size of the archive handed to the reader. It bounds
// the memory the parser can be asked to look at, not the amount it extracts:
// per-entry extraction is bounded separately by maxZipEntrySize.
const maxZipInputSize = 200 << 20 // 200 MB

// maxZipEntries is the maximum number of files allowed in a ZIP archive.
const maxZipEntries = 10000

// readFileFromZip returns the decompressed contents of one part of a package.
//
// Reading a presentation looks parts up hundreds of times — every slide, its
// relationships, and every embedded image and chart — so the lookup is done
// through zip.Reader.Open, which searches the reader's name index instead of
// walking every entry. A linear scan here would cost O(parts) per lookup.
//
// Two size checks are applied: the uncompressed size the archive declares, to
// reject an oversized part before spending CPU on it, and the number of bytes
// actually produced, because the declared size is attacker-controlled metadata
// and may understate what the entry expands to.
func readFileFromZip(zr *zip.Reader, name string) ([]byte, error) {
	if len(zr.File) > maxZipEntries {
		return nil, fmt.Errorf("zip archive contains too many entries (%d > %d)", len(zr.File), maxZipEntries)
	}
	rc, declaredSize, err := openZipPart(zr, name)
	if err != nil {
		return nil, err
	}
	defer rc.Close()

	if declaredSize > int64(maxZipEntrySize) {
		return nil, fmt.Errorf("file %s exceeds maximum allowed size (%d bytes)", name, maxZipEntrySize)
	}
	data, err := io.ReadAll(io.LimitReader(rc, int64(maxZipEntrySize)+1))
	if err != nil {
		return nil, fmt.Errorf("failed to read %s from zip: %w", name, err)
	}
	if int64(len(data)) > int64(maxZipEntrySize) {
		return nil, fmt.Errorf("file %s actual size exceeds maximum allowed size", name)
	}
	return data, nil
}

// openZipPart opens a single part and reports its declared uncompressed size,
// or -1 when the size is unavailable.
//
// zip.Reader.Open only accepts fs-valid paths, and normalises the separators of
// the names it indexes, so an archive written with backslashes still resolves.
// Names it refuses — a relationship target containing "..", say — fall back to
// the exact string match the reader has always used, which is also what keeps
// the behaviour identical for archives that carry such names.
func openZipPart(zr *zip.Reader, name string) (io.ReadCloser, int64, error) {
	if f, err := zr.Open(name); err == nil {
		if info, statErr := f.Stat(); statErr == nil {
			return f, info.Size(), nil
		}
		return f, -1, nil
	}
	for _, f := range zr.File {
		if f.Name != name {
			continue
		}
		rc, err := f.Open()
		if err != nil {
			return nil, 0, fmt.Errorf("failed to open %s in zip: %w", name, err)
		}
		return rc, int64(f.UncompressedSize64), nil
	}
	return nil, 0, fmt.Errorf("file not found in zip: %s", name)
}

// --- Relationship reading ---

type xmlRelForRead struct {
	ID         string `xml:"Id,attr"`
	Type       string `xml:"Type,attr"`
	Target     string `xml:"Target,attr"`
	TargetMode string `xml:"TargetMode,attr"`
}

type xmlRelsForRead struct {
	XMLName       xml.Name        `xml:"Relationships"`
	Relationships []xmlRelForRead `xml:"Relationship"`
}

func (r *PPTXReader) readRelationships(zr *zip.Reader, path string) ([]xmlRelForRead, error) {
	data, err := readFileFromZip(zr, path)
	if err != nil {
		return nil, nil // relationships file may not exist
	}

	var rels xmlRelsForRead
	if err := xml.Unmarshal(data, &rels); err != nil {
		return nil, fmt.Errorf("failed to parse relationships %s: %w", path, err)
	}
	return rels.Relationships, nil
}
