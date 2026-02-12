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

// zipIndex builds a map from file name to *zip.File for O(1) lookups.
func zipIndex(zr *zip.Reader) map[string]*zip.File {
	m := make(map[string]*zip.File, len(zr.File))
	for _, f := range zr.File {
		m[f.Name] = f
	}
	return m
}

// Read reads a presentation from a file path.
func (r *PPTXReader) Read(path string) (*Presentation, error) {
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
func (r *PPTXReader) ReadFromReader(reader io.ReaderAt, size int64) (*Presentation, error) {
	if size <= 0 {
		return nil, fmt.Errorf("invalid reader size: %d", size)
	}
	if size > int64(maxZipTotalSize) {
		return nil, fmt.Errorf("file size %d exceeds maximum allowed (%d bytes)", size, maxZipTotalSize)
	}

	zr, err := zip.NewReader(reader, size)
	if err != nil {
		return nil, fmt.Errorf("failed to open zip: %w", err)
	}

	if len(zr.File) > maxZipEntries {
		return nil, fmt.Errorf("zip archive contains too many entries (%d > %d)", len(zr.File), maxZipEntries)
	}

	pres := &Presentation{
		properties:             NewDocumentProperties(),
		presentationProperties: NewPresentationProperties(),
		slides:                 make([]*Slide, 0),
		slideMasters:           make([]*SlideMaster, 0),
		layout:                 NewDocumentLayout(),
	}

	// Read core properties (non-fatal: missing properties are acceptable)
	_ = r.readCoreProperties(zr, pres)

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

		slide, err := r.readSlide(zr, target, pres)
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

// maxZipTotalSize is the cumulative limit for all extracted content from a single ZIP.
const maxZipTotalSize = 200 << 20 // 200 MB

// maxZipEntries is the maximum number of files allowed in a ZIP archive.
const maxZipEntries = 10000

func readFileFromZip(zr *zip.Reader, name string) ([]byte, error) {
	if len(zr.File) > maxZipEntries {
		return nil, fmt.Errorf("zip archive contains too many entries (%d > %d)", len(zr.File), maxZipEntries)
	}
	for _, f := range zr.File {
		if f.Name == name {
			if f.UncompressedSize64 > maxZipEntrySize {
				return nil, fmt.Errorf("file %s exceeds maximum allowed size (%d bytes)", name, maxZipEntrySize)
			}
			rc, err := f.Open()
			if err != nil {
				return nil, fmt.Errorf("failed to open %s in zip: %w", name, err)
			}
			defer rc.Close()
			data, err := io.ReadAll(io.LimitReader(rc, int64(maxZipEntrySize)+1))
			if err != nil {
				return nil, fmt.Errorf("failed to read %s from zip: %w", name, err)
			}
			if int64(len(data)) > int64(maxZipEntrySize) {
				return nil, fmt.Errorf("file %s actual size exceeds maximum allowed size", name)
			}
			return data, nil
		}
	}
	return nil, fmt.Errorf("file not found in zip: %s", name)
}

// --- Relationship reading ---

type xmlRelForRead struct {
	ID         string `xml:"Id,attr"`
	Type       string `xml:"Type,attr"`
	Target     string `xml:"Target,attr"`
	TargetMode string `xml:"TargetMode,attr"`
}

type xmlRelsForRead struct {
	XMLName       xml.Name         `xml:"Relationships"`
	Relationships []xmlRelForRead  `xml:"Relationship"`
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
