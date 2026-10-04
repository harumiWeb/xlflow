package picture

import (
	"bytes"
	"encoding/binary"
	"os"
	"os/exec"
	"path/filepath"
	"runtime"
	"testing"
)

const excelAuthoredAssetDir = "../compiler/testdata/pictures-excel-authored"

func TestLoadAcceptsBMPAndJPEGWithWindowsSeparators(t *testing.T) {
	root := t.TempDir()
	if err := os.Mkdir(filepath.Join(root, "assets"), 0o755); err != nil {
		t.Fatal(err)
	}
	for _, name := range []string{"logo.bmp", "logo.jpg"} {
		data, err := os.ReadFile(filepath.Join(excelAuthoredAssetDir, name))
		if err != nil {
			t.Fatal(err)
		}
		if err := os.WriteFile(filepath.Join(root, "assets", name), data, 0o600); err != nil {
			t.Fatal(err)
		}
		asset, err := Load(root, `assets\`+name)
		if err != nil {
			t.Fatalf("Load(%q): %v", name, err)
		}
		wantFormat := "bmp"
		if filepath.Ext(name) == ".jpg" {
			wantFormat = "jpeg"
		}
		if asset.Format != wantFormat || asset.Width != 24 || asset.Height != 16 || !bytes.Equal(asset.Data, data) {
			t.Fatalf("Load(%q) = %#v, want format %s and 24x16 source bytes", name, asset, wantFormat)
		}
	}
}

func TestEncodeDecodeBMPAndJPEGStdPicture(t *testing.T) {
	for _, name := range []string{"logo.bmp", "logo.jpg"} {
		t.Run(name, func(t *testing.T) {
			data, err := os.ReadFile(filepath.Join(excelAuthoredAssetDir, name))
			if err != nil {
				t.Fatal(err)
			}
			raw, err := Encode(data)
			if err != nil {
				t.Fatal(err)
			}
			if !bytes.Equal(raw[:16], stdPictureCLSID[:]) || binary.LittleEndian.Uint32(raw[16:20]) != stdPictureTag || binary.LittleEndian.Uint32(raw[20:24]) != uint32(len(data)) {
				t.Fatalf("StdPicture envelope has wrong CLSID, preamble, or byte length: %x", raw[:24])
			}
			asset, err := Decode(raw)
			if err != nil {
				t.Fatal(err)
			}
			wantFormat := "bmp"
			if filepath.Ext(name) == ".jpg" {
				wantFormat = "jpeg"
			}
			if asset.Format != wantFormat || asset.Width != 24 || asset.Height != 16 || !bytes.Equal(asset.Data, data) {
				t.Fatalf("Decode did not return owned source payload: %#v", asset)
			}
			asset.Data[0] ^= 0xff
			if !bytes.Equal(raw[24:], data) {
				t.Fatal("Decode.Data aliases the encoded input")
			}
		})
	}
}

func TestDecodeRejectsMalformedStdPictureAndUnsupportedBMP(t *testing.T) {
	data, err := os.ReadFile(filepath.Join(excelAuthoredAssetDir, "logo.bmp"))
	if err != nil {
		t.Fatal(err)
	}
	valid, err := Encode(data)
	if err != nil {
		t.Fatal(err)
	}
	for _, test := range []struct {
		name string
		edit func([]byte) []byte
	}{
		{"short", func(raw []byte) []byte { return raw[:23] }},
		{"clsid", func(raw []byte) []byte { raw[0] ^= 1; return raw }},
		{"preamble", func(raw []byte) []byte { raw[16] = 0; return raw }},
		{"length", func(raw []byte) []byte { raw[20]++; return raw }},
		{"trailing", func(raw []byte) []byte { return append(raw, 0) }},
		{"BMP-bit-depth", func(raw []byte) []byte { raw[24+28] = 16; raw[24+29] = 0; return raw }},
		{"BMP-compression", func(raw []byte) []byte { raw[24+30] = 1; return raw }},
	} {
		t.Run(test.name, func(t *testing.T) {
			if _, err := Decode(test.edit(bytes.Clone(valid))); err == nil {
				t.Fatal("Decode accepted malformed or unsupported payload")
			}
		})
	}
}

func TestLoadRejectsAbsoluteAndEscapingPaths(t *testing.T) {
	root := t.TempDir()
	for _, name := range []string{`..\outside.bmp`, `../outside.bmp`, `/outside.bmp`, `\outside.bmp`, `C:\outside.bmp`, `C:outside.bmp`, `\\server\share\outside.bmp`} {
		t.Run(name, func(t *testing.T) {
			if _, err := Load(root, name); err == nil {
				t.Fatalf("Load accepted non-project-relative path %q", name)
			}
		})
	}
}

func TestLoadRejectsJunctionEscape(t *testing.T) {
	if runtime.GOOS != "windows" {
		t.Skip("Windows junction confinement")
	}
	parent := t.TempDir()
	root, outside := filepath.Join(parent, "root"), filepath.Join(parent, "outside")
	for _, dir := range []string{root, outside} {
		if err := os.Mkdir(dir, 0o755); err != nil {
			t.Fatal(err)
		}
	}
	data, err := os.ReadFile(filepath.Join(excelAuthoredAssetDir, "logo.bmp"))
	if err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(outside, "logo.bmp"), data, 0o600); err != nil {
		t.Fatal(err)
	}
	junction := filepath.Join(root, "escape")
	if output, err := exec.Command("cmd.exe", "/c", "mklink", "/J", junction, outside).CombinedOutput(); err != nil {
		t.Skipf("cannot create Windows junction: %v (%s)", err, output)
	}
	if _, err := Load(root, `escape\logo.bmp`); err == nil {
		t.Fatal("Load followed a junction outside the project root")
	}
}

func TestLoadEnforcesFileAndPixelLimits(t *testing.T) {
	root := t.TempDir()
	tooLarge := make([]byte, maxAssetBytes+1)
	if err := os.WriteFile(filepath.Join(root, "large.bmp"), tooLarge, 0o600); err != nil {
		t.Fatal(err)
	}
	if _, err := Load(root, "large.bmp"); err == nil {
		t.Fatal("Load accepted an asset over 16 MiB")
	}
	data := make([]byte, 54)
	copy(data, "BM")
	binary.LittleEndian.PutUint32(data[2:6], uint32(len(data)))
	binary.LittleEndian.PutUint32(data[10:14], 54)
	binary.LittleEndian.PutUint32(data[14:18], 40)
	binary.LittleEndian.PutUint32(data[18:22], 4001)
	binary.LittleEndian.PutUint32(data[22:26], 4000)
	binary.LittleEndian.PutUint16(data[26:28], 1)
	binary.LittleEndian.PutUint16(data[28:30], 24)
	if err := os.WriteFile(filepath.Join(root, "pixels.bmp"), data, 0o600); err != nil {
		t.Fatal(err)
	}
	if _, err := Load(root, "pixels.bmp"); err == nil {
		t.Fatal("Load accepted an asset over 16 million pixels")
	}
}
