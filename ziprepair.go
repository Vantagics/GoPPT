package gopresentation

import (
	"archive/zip"
	"bytes"
	"compress/flate"
	"encoding/binary"
	"fmt"
	"hash/crc32"
	"io"
	"time"
)

// The end-of-central-directory record is the only thing archive/zip trusts, so
// an archive whose tail was cut off — a truncated download, a half-written
// file — reads as "not a zip at all", even though every part up to the cut is
// intact. PowerPoint repairs such packages by walking the local file headers
// from the front and reading whatever they still declare; repairTruncatedZip
// does the same, so a deck Office opens is not a read failure here.
//
// Repair is a last resort: it only runs when the normal open has already
// failed, and it rebuilds the parts into a fresh in-memory archive the
// standard reader can index. The walk stops at the first unusable header —
// a cut-off name, a corrupt deflate stream, a checksum mismatch — and drops
// everything from there on, exactly as Office drops it.

const (
	zipLocalHeaderSig = 0x04034b50 // "PK\x03\x04"
	zipLocalHeaderLen = 30
)

// repairTruncatedZip rebuilds a readable archive from a truncated one. An
// error means the walk found nothing usable and the caller should report the
// original open error.
func repairTruncatedZip(reader io.ReaderAt, size int64) (*zip.Reader, error) {
	if size > int64(maxZipInputSize) {
		return nil, fmt.Errorf("repair refused: size %d exceeds maximum (%d bytes)", size, maxZipInputSize)
	}
	buf, err := io.ReadAll(io.NewSectionReader(reader, 0, size))
	if err != nil {
		return nil, err
	}

	var out bytes.Buffer
	w := zip.NewWriter(&out)
	entries := 0
	for off := 0; off+zipLocalHeaderLen <= len(buf); {
		if binary.LittleEndian.Uint32(buf[off:]) != zipLocalHeaderSig {
			off++ // padding or a stray descriptor; keep looking before giving up
			continue
		}
		flags := binary.LittleEndian.Uint16(buf[off+6:])
		method := binary.LittleEndian.Uint16(buf[off+8:])
		wantCRC := binary.LittleEndian.Uint32(buf[off+14:])
		csize := int64(binary.LittleEndian.Uint32(buf[off+18:]))
		nlen := int(binary.LittleEndian.Uint16(buf[off+26:]))
		xlen := int(binary.LittleEndian.Uint16(buf[off+28:]))
		nameStart := off + zipLocalHeaderLen
		dataStart := int64(nameStart + nlen + xlen)
		if nlen == 0 || dataStart > int64(len(buf)) {
			break // the header itself is cut off; nothing reliable after it
		}
		name := string(buf[nameStart : nameStart+nlen])
		if flags&0x8 != 0 {
			// Sizes live in a trailing data descriptor. A deflate stream is
			// self-terminating, so it can be inflated blind and the descriptor
			// read where the stream ends; stored data has no such boundary.
			if method != zip.Deflate {
				break
			}
			content, _ := inflateBlind(buf[dataStart:])
			if content == nil {
				break
			}
			next := findDescriptor(buf, dataStart, content)
			if next < 0 {
				break
			}
			if err := writeRepairedEntry(w, name, content); err != nil {
				return nil, err
			}
			entries++
			off = int(next)
			continue
		}
		if dataStart+csize > int64(len(buf)) {
			break // declared data runs past the cut: the tail entry is lost
		}
		if entries >= maxZipEntries {
			break
		}

		raw := buf[dataStart : dataStart+csize]
		var content []byte
		switch method {
		case zip.Store:
			content = raw
		case zip.Deflate:
			if inflated, infErr := inflateBounded(raw); infErr == nil {
				content = inflated
			}
		}
		if content == nil || crc32.ChecksumIEEE(content) != wantCRC {
			break // corrupt stream or checksum: nothing reliable after this
		}
		if err := writeRepairedEntry(w, name, content); err != nil {
			return nil, err
		}
		entries++
		off = int(dataStart + csize)
	}
	if entries == 0 {
		return nil, fmt.Errorf("repair found no usable entries")
	}
	if err := w.Close(); err != nil {
		return nil, err
	}

	// The rebuilt bytes stay reachable through the returned reader, which
	// holds the bytes.Reader they were indexed from.
	repaired := out.Bytes()
	return zip.NewReader(bytes.NewReader(repaired), int64(len(repaired)))
}

// inflateBounded inflates one entry's stream and fails when it produces more
// than maxZipEntrySize bytes: a declared size is absent here, so the bound is
// the only guard against a hostile stream.
func inflateBounded(raw []byte) ([]byte, error) {
	fr := flate.NewReader(bytes.NewReader(raw))
	defer fr.Close()
	out, err := io.ReadAll(io.LimitReader(fr, int64(maxZipEntrySize)+1))
	if err != nil {
		return nil, err
	}
	if len(out) > maxZipEntrySize {
		return nil, fmt.Errorf("entry exceeds maximum size")
	}
	return out, nil
}

// countingReader counts the bytes a blind inflation consumed. Kept for
// diagnostics; the descriptor search below does not rely on it, because the
// flate decoder reads ahead in buffered chunks and overshoots the stream end.
type countingReader struct {
	r *bytes.Reader
	n int64
}

func (c *countingReader) Read(p []byte) (int, error) {
	n, err := c.r.Read(p)
	c.n += int64(n)
	return n, err
}

// inflateBlind inflates a stream whose compressed length is unknown. The
// flate format is self-terminating, so the bytes produced are the entry's
// whole content even though the decoder may have buffered past its end.
func inflateBlind(raw []byte) ([]byte, int64) {
	cr := &countingReader{r: bytes.NewReader(raw)}
	fr := flate.NewReader(cr)
	defer fr.Close()
	out, err := io.ReadAll(io.LimitReader(fr, int64(maxZipEntrySize)+1))
	if err != nil || len(out) > maxZipEntrySize {
		return nil, 0
	}
	return out, cr.n
}

// findDescriptor locates the data descriptor that follows a bit-3 entry and
// returns the offset just past it, or -1. The flate decoder overshoots the
// stream end, so the descriptor is found by search, not position: the first
// candidate whose declared compressed size matches the distance from the
// data start and whose CRC matches the recovered content. The signature is
// optional, so both layouts are accepted.
func findDescriptor(buf []byte, dataStart int64, content []byte) int64 {
	crc := crc32.ChecksumIEEE(content)
	end := int64(len(buf))
	limit := dataStart + 65536
	for p := dataStart; p+12 <= end && p < limit; p++ {
		rest := buf[p:]
		if bytes.HasPrefix(rest, []byte("PK\x07\x08")) {
			if binary.LittleEndian.Uint32(rest[8:]) == uint32(p-dataStart) &&
				binary.LittleEndian.Uint32(rest[4:]) == crc {
				return p + 16
			}
			continue
		}
		if binary.LittleEndian.Uint32(rest[4:]) == uint32(p-dataStart) &&
			binary.LittleEndian.Uint32(rest[0:]) == crc {
			return p + 12
		}
	}
	return -1
}

// writeRepairedEntry stores one recovered part as-is. The content is already
// decompressed, so the rebuilt copy carries it uncompressed: re-compressing
// recovered bytes buys nothing and store entries cannot fail.
func writeRepairedEntry(w *zip.Writer, name string, content []byte) error {
	h := &zip.FileHeader{Name: name, Method: zip.Store, Modified: time.Unix(0, 0)}
	f, err := w.CreateHeader(h)
	if err != nil {
		return err
	}
	_, err = f.Write(content)
	return err
}
