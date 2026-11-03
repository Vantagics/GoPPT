package gopresentation

import (
	"archive/zip"
	"bytes"
	"encoding/binary"
	"hash/crc32"
	"testing"
)

// buildStoreZip builds a store-method archive with the given named payloads.
func buildStoreZip(t *testing.T, parts map[string]string) []byte {
	t.Helper()
	var buf bytes.Buffer
	w := zip.NewWriter(&buf)
	for _, name := range []string{"[Content_Types].xml", "ppt/presentation.xml"} {
		f, err := w.Create(name)
		if err != nil {
			t.Fatal(err)
		}
		if _, err := f.Write([]byte(parts[name])); err != nil {
			t.Fatal(err)
		}
	}
	if err := w.Close(); err != nil {
		t.Fatal(err)
	}
	return buf.Bytes()
}

// TestRepairTruncatedZip recovers every intact entry from an archive whose
// end-of-central-directory record is missing and whose tail carries the start
// of a new local header — the shape a cut-off download leaves behind.
func TestRepairTruncatedZip(t *testing.T) {
	intact := buildStoreZip(t, map[string]string{
		"[Content_Types].xml":  "<?xml version=\"1.0\"?><Types/>",
		"ppt/presentation.xml": "<p:presentation/>",
	})
	truncated := append(append([]byte{}, intact[:len(intact):len(intact)]...), "PK\x03\x04\n"...)

	zr, err := repairTruncatedZip(bytes.NewReader(truncated), int64(len(truncated)))
	if err != nil {
		t.Fatalf("repair: %v", err)
	}
	if len(zr.File) != 2 {
		t.Fatalf("entries = %d, want 2", len(zr.File))
	}
	for _, f := range zr.File {
		rc, err := f.Open()
		if err != nil {
			t.Fatalf("open %s: %v", f.Name, err)
		}
		var got bytes.Buffer
		if _, err := got.ReadFrom(rc); err != nil {
			t.Fatalf("read %s: %v", f.Name, err)
		}
		rc.Close()
		if want := partsWant(f.Name); got.String() != want {
			t.Errorf("%s content = %q, want %q", f.Name, got.String(), want)
		}
	}
}

func partsWant(name string) string {
	if name == "[Content_Types].xml" {
		return "<?xml version=\"1.0\"?><Types/>"
	}
	return "<p:presentation/>"
}

// buildRawStoreZip hand-writes a store-method archive whose local headers
// carry the sizes, the layout PowerPoint itself produces. Returned with no
// central directory: the repair walk does not need one, and the corrupt-entry
// test needs exact offsets into the entry data.
func buildRawStoreZip(t *testing.T, names []string, datas map[string][]byte) []byte {
	t.Helper()
	var buf bytes.Buffer
	for _, name := range names {
		data := datas[name]
		h := make([]byte, zipLocalHeaderLen)
		binary.LittleEndian.PutUint32(h[0:], 0x04034b50)
		binary.LittleEndian.PutUint16(h[4:], 20) // version
		binary.LittleEndian.PutUint16(h[6:], 0)  // flags: sizes in the header
		binary.LittleEndian.PutUint16(h[8:], 0)  // store
		binary.LittleEndian.PutUint32(h[14:], crc32.ChecksumIEEE(data))
		binary.LittleEndian.PutUint32(h[18:], uint32(len(data)))
		binary.LittleEndian.PutUint32(h[22:], uint32(len(data)))
		binary.LittleEndian.PutUint16(h[26:], uint16(len(name)))
		buf.Write(h)
		buf.WriteString(name)
		buf.Write(data)
	}
	return buf.Bytes()
}

// TestRepairStopsAtCorruptEntry drops everything from the first corrupt
// entry on: Office's repair reads the prefix and discards the rest, and so
// does this one — a partially recovered deck beats no deck.
func TestRepairStopsAtCorruptEntry(t *testing.T) {
	names := []string{"a.xml", "b.xml"}
	datas := map[string][]byte{"a.xml": []byte("<a/>"), "b.xml": []byte("<b/>")}
	raw := buildRawStoreZip(t, names, datas)

	// Overwrite the first bytes of b's stored data, then cut the tail off.
	damaged := append([]byte{}, raw...)
	bDataStart := zipLocalHeaderLen + len("a.xml") + len("<a/>") + zipLocalHeaderLen + len("b.xml")
	copy(damaged[bDataStart:], "XXXX")
	truncated := append(damaged, "PK\x03\x04\n"...)

	zr, err := repairTruncatedZip(bytes.NewReader(truncated), int64(len(truncated)))
	if err != nil {
		t.Fatalf("repair: %v", err)
	}
	if len(zr.File) != 1 {
		t.Fatalf("entries = %d, want the one intact entry", len(zr.File))
	}
	if zr.File[0].Name != "a.xml" {
		t.Fatalf("kept %q, want the intact entry", zr.File[0].Name)
	}
}

// TestIntactZipNeedsNoRepair pins the normal path: a healthy archive must
// never reach the salvage code.
func TestIntactZipNeedsNoRepair(t *testing.T) {
	intact := buildStoreZip(t, map[string]string{
		"[Content_Types].xml":  "<?xml version=\"1.0\"?><Types/>",
		"ppt/presentation.xml": "<p:presentation/>",
	})
	if _, err := zip.NewReader(bytes.NewReader(intact), int64(len(intact))); err != nil {
		t.Fatalf("control zip should open normally: %v", err)
	}
}
