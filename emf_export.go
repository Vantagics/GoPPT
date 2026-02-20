package gopresentation

import "image"

// DecodeEMFForTest is an exported wrapper for testing EMF decoding.
func DecodeEMFForTest(data []byte) image.Image {
	return decodeEMFBitmap(data)
}
