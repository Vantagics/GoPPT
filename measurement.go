package gopresentation

import "math"

// EMU (English Metric Units) conversion helpers.
// 1 inch = 914400 EMU, 1 point = 12700 EMU, 1 cm = 360000 EMU.

const (
	emuPerInch       = 914400
	emuPerPoint      = 12700
	emuPerCentimeter = 360000
	emuPerMillimeter = 36000
	// maxEMU is the maximum safe EMU value to prevent overflow.
	maxEMU = math.MaxInt64 / 2
)

// Inch converts inches to EMU. Clamps to safe range.
func Inch(n float64) int64 {
	return clampEMU(n * emuPerInch)
}

// Point converts points to EMU.
func Point(n float64) int64 {
	return clampEMU(n * emuPerPoint)
}

// Centimeter converts centimeters to EMU.
func Centimeter(n float64) int64 {
	return clampEMU(n * emuPerCentimeter)
}

// Millimeter converts millimeters to EMU.
func Millimeter(n float64) int64 {
	return clampEMU(n * emuPerMillimeter)
}

// EMUToInch converts EMU to inches.
func EMUToInch(emu int64) float64 {
	return float64(emu) / emuPerInch
}

// EMUToPoint converts EMU to points.
func EMUToPoint(emu int64) float64 {
	return float64(emu) / emuPerPoint
}

// EMUToCentimeter converts EMU to centimeters.
func EMUToCentimeter(emu int64) float64 {
	return float64(emu) / emuPerCentimeter
}

// EMUToMillimeter converts EMU to millimeters.
func EMUToMillimeter(emu int64) float64 {
	return float64(emu) / emuPerMillimeter
}

// clampEMU converts a float64 to int64, clamping to prevent overflow.
func clampEMU(v float64) int64 {
	if v > float64(maxEMU) {
		return maxEMU
	}
	if v < -float64(maxEMU) {
		return -maxEMU
	}
	return int64(v)
}
