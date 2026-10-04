// Package picture reads source images and encodes the StdPicture payload used
// by MS-OFORMS picture properties.
package picture

import (
	"bytes"
	"encoding/binary"
	"errors"
	"fmt"
	"image/jpeg"
	"io"
	"math"
	"os"
	"path/filepath"
	"strings"
)

const (
	maxAssetBytes  = 16 << 20
	maxAssetPixels = 16_000_000
	stdPictureTag  = uint32(0x746c)
)

var stdPictureCLSID = [16]byte{
	0x04, 0x52, 0xe3, 0x0b, 0x91, 0x8f, 0xce, 0x11,
	0x9d, 0xe3, 0x00, 0xaa, 0x00, 0x4b, 0xb8, 0x51,
}

// Asset contains a validated BMP or JPEG source image. Data is the original
// file payload, while Format is normalized to "bmp" or "jpeg".
type Asset struct {
	Data          []byte
	Format        string
	Width, Height int
}

// Load reads a project-root-relative image without following links outside
// root. The byte and decoded-pixel limits apply before returning the asset.
func Load(root, path string) (asset Asset, resultErr error) {
	normalized := strings.ReplaceAll(path, `\`, "/")
	if strings.TrimSpace(path) == "" || strings.HasPrefix(normalized, "/") || strings.Contains(normalized, ":") || !filepath.IsLocal(filepath.FromSlash(normalized)) {
		return Asset{}, fmt.Errorf("picture path must be project-root-relative")
	}
	clean := filepath.Clean(filepath.FromSlash(normalized))
	if clean == "." || !filepath.IsLocal(clean) {
		return Asset{}, fmt.Errorf("picture path must name a file inside the project root")
	}
	rootHandle, err := os.OpenRoot(root)
	if err != nil {
		return Asset{}, fmt.Errorf("open project root: %w", err)
	}
	defer func() {
		if closeErr := rootHandle.Close(); closeErr != nil && resultErr == nil {
			resultErr = fmt.Errorf("close project root: %w", closeErr)
		}
	}()

	file, err := rootHandle.Open(clean)
	if err != nil {
		return Asset{}, fmt.Errorf("open picture %q: %w", path, err)
	}
	defer func() {
		if closeErr := file.Close(); closeErr != nil && resultErr == nil {
			resultErr = fmt.Errorf("close picture %q: %w", path, closeErr)
		}
	}()
	info, err := file.Stat()
	if err != nil {
		return Asset{}, fmt.Errorf("stat picture %q: %w", path, err)
	}
	if !info.Mode().IsRegular() {
		return Asset{}, fmt.Errorf("picture %q is not a regular file", path)
	}
	if info.Size() <= 0 || info.Size() > maxAssetBytes {
		return Asset{}, fmt.Errorf("picture %q must be between 1 byte and %d bytes", path, maxAssetBytes)
	}
	data, err := io.ReadAll(io.LimitReader(file, maxAssetBytes+1))
	if err != nil {
		return Asset{}, fmt.Errorf("read picture %q: %w", path, err)
	}
	if len(data) == 0 || len(data) > maxAssetBytes {
		return Asset{}, fmt.Errorf("picture %q exceeds the %d-byte limit", path, maxAssetBytes)
	}
	asset, err = inspect(data)
	if err != nil {
		return Asset{}, fmt.Errorf("picture %q: %w", path, err)
	}
	asset.Data = data
	return asset, nil
}

// Encode wraps a validated source image in the StdPicture wire format used by
// MS-OFORMS GuidAndPicture records.
func Encode(data []byte) ([]byte, error) {
	if _, err := inspect(data); err != nil {
		return nil, err
	}
	raw := make([]byte, 0, len(stdPictureCLSID)+8)
	raw = append(raw, stdPictureCLSID[:]...)
	raw = binary.LittleEndian.AppendUint32(raw, stdPictureTag)
	raw = binary.LittleEndian.AppendUint32(raw, uint32(len(data)))
	raw = append(raw, data...)
	if _, err := Decode(raw); err != nil {
		return nil, fmt.Errorf("validate encoded StdPicture: %w", err)
	}
	return raw, nil
}

// Decode validates and unwraps a persisted StdPicture payload.
func Decode(raw []byte) (Asset, error) {
	if len(raw) < len(stdPictureCLSID)+8 {
		return Asset{}, errors.New("truncated GuidAndPicture payload")
	}
	if !bytes.Equal(raw[:len(stdPictureCLSID)], stdPictureCLSID[:]) {
		return Asset{}, errors.New("unexpected StdPicture CLSID")
	}
	if tag := binary.LittleEndian.Uint32(raw[16:20]); tag != stdPictureTag {
		return Asset{}, fmt.Errorf("unexpected StdPicture preamble %#08x", tag)
	}
	size := binary.LittleEndian.Uint32(raw[20:24])
	if uint64(size) != uint64(len(raw)-24) {
		return Asset{}, fmt.Errorf("StdPicture size %d does not match payload length %d", size, len(raw)-24)
	}
	asset, err := inspect(raw[24:])
	if err != nil {
		return Asset{}, err
	}
	asset.Data = bytes.Clone(raw[24:])
	return asset, nil
}

func inspect(data []byte) (Asset, error) {
	if len(data) == 0 || len(data) > maxAssetBytes {
		return Asset{}, fmt.Errorf("image must be between 1 byte and %d bytes", maxAssetBytes)
	}
	if len(data) >= 2 && data[0] == 'B' && data[1] == 'M' {
		width, height, err := inspectBMP(data)
		if err != nil {
			return Asset{}, err
		}
		return Asset{Format: "bmp", Width: width, Height: height}, nil
	}
	if len(data) >= 2 && data[0] == 0xff && data[1] == 0xd8 {
		config, err := jpeg.DecodeConfig(bytes.NewReader(data))
		if err != nil {
			return Asset{}, fmt.Errorf("invalid JPEG image: %w", err)
		}
		if err := validateDimensions(config.Width, config.Height); err != nil {
			return Asset{}, err
		}
		if _, err := jpeg.Decode(bytes.NewReader(data)); err != nil {
			return Asset{}, fmt.Errorf("invalid JPEG image: %w", err)
		}
		return Asset{Format: "jpeg", Width: config.Width, Height: config.Height}, nil
	}
	return Asset{}, errors.New("image format must be BMP or JPEG")
}

func inspectBMP(data []byte) (int, int, error) {
	if len(data) < 54 {
		return 0, 0, errors.New("truncated BMP file header")
	}
	declaredSize := uint64(binary.LittleEndian.Uint32(data[2:6]))
	pixelOffset := uint64(binary.LittleEndian.Uint32(data[10:14]))
	headerSize := binary.LittleEndian.Uint32(data[14:18])
	if headerSize != 40 {
		return 0, 0, errors.New("only 40-byte BITMAPINFOHEADER BMP files are supported")
	}
	width := int64(int32(binary.LittleEndian.Uint32(data[18:22])))
	height := int64(int32(binary.LittleEndian.Uint32(data[22:26])))
	if width <= 0 || height == 0 || height == math.MinInt32 {
		return 0, 0, errors.New("BMP dimensions must be positive and non-zero")
	}
	if height < 0 {
		height = -height
	}
	widthInt, heightInt := int(width), int(height)
	if err := validateDimensions(widthInt, heightInt); err != nil {
		return 0, 0, err
	}
	if binary.LittleEndian.Uint16(data[26:28]) != 1 {
		return 0, 0, errors.New("BMP plane count must be 1")
	}
	bits := binary.LittleEndian.Uint16(data[28:30])
	if bits != 24 && bits != 32 {
		return 0, 0, fmt.Errorf("only 24-bit and 32-bit BMP files are supported (got %d-bit)", bits)
	}
	compression := binary.LittleEndian.Uint32(data[30:34])
	if compression != 0 {
		return 0, 0, fmt.Errorf("only uncompressed BI_RGB BMP files are supported (got compression %d)", compression)
	}
	headerEnd := uint64(14 + headerSize)
	if pixelOffset < headerEnd || pixelOffset > uint64(len(data)) {
		return 0, 0, errors.New("BMP pixel offset is outside the image payload")
	}
	rowBytes := ((uint64(widthInt)*uint64(bits) + 31) / 32) * 4
	pixelBytes := rowBytes * uint64(heightInt)
	if pixelOffset+pixelBytes > uint64(len(data)) {
		return 0, 0, errors.New("truncated BMP pixel array")
	}
	if declaredSize != 0 && (declaredSize > uint64(len(data)) || declaredSize < pixelOffset+pixelBytes) {
		return 0, 0, errors.New("BMP declared file size does not contain the pixel array")
	}
	imageSize := uint64(binary.LittleEndian.Uint32(data[34:38]))
	if imageSize != 0 && imageSize < pixelBytes {
		return 0, 0, errors.New("BMP declared image size is smaller than the pixel array")
	}
	return widthInt, heightInt, nil
}

func validateDimensions(width, height int) error {
	if width <= 0 || height <= 0 || uint64(width)*uint64(height) > maxAssetPixels {
		return fmt.Errorf("image dimensions %dx%d exceed the %d-pixel limit", width, height, maxAssetPixels)
	}
	return nil
}
