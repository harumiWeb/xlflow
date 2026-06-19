package ovba

import (
	"bytes"
	"math/rand"
	"testing"
)

// randBytes returns n deterministic high-entropy (incompressible) bytes.
func randBytes(n int, seed int64) []byte {
	b := make([]byte, n)
	rand.New(rand.NewSource(seed)).Read(b)
	return b
}

// TestCompressRawPaddingBoundary is the C1 regression: a raw chunk is padded to
// 4096 bytes, so it must only be emitted for full 4096-byte chunks. A short
// final chunk that cannot compress below 4096 ([MS-OVBA] 12-bit CompressedChunkSize)
// has no lossless representation and must error rather than silently corrupt.
func TestCompressRawPaddingBoundary(t *testing.T) {
	t.Run("compressible_short_roundtrips", func(t *testing.T) {
		in := bytes.Repeat([]byte("ABCD"), 1000) // 4000B, repetitive -> compressed
		out, err := Compress(in)
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		got, err := Decompress(out)
		if err != nil {
			t.Fatalf("decompress: %v", err)
		}
		if !bytes.Equal(got, in) {
			t.Errorf("round-trip mismatch: in=%dB out=%dB", len(in), len(got))
		}
	})

	t.Run("incompressible_short_under_limit_roundtrips", func(t *testing.T) {
		// N=3000 <= 3640: worst-case body (N + N/8) stays under 4096 -> compressed.
		in := randBytes(3000, 1)
		out, err := Compress(in)
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		got, err := Decompress(out)
		if err != nil {
			t.Fatalf("decompress: %v", err)
		}
		if !bytes.Equal(got, in) {
			t.Errorf("round-trip mismatch: in=%dB out=%dB", len(in), len(got))
		}
	})

	t.Run("full_incompressible_chunk_roundtrips", func(t *testing.T) {
		// Exactly 4096B incompressible -> raw chunk, lossless (no padding).
		in := randBytes(4096, 2)
		out, err := Compress(in)
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		got, err := Decompress(out)
		if err != nil {
			t.Fatalf("decompress: %v", err)
		}
		if !bytes.Equal(got, in) {
			t.Errorf("round-trip mismatch: in=%dB out=%dB", len(in), len(got))
		}
	})

	t.Run("incompressible_short_over_limit_errors", func(t *testing.T) {
		// N=4000 in [3641,4095], incompressible: no lossless encoding -> error.
		in := randBytes(4000, 1)
		if _, err := Compress(in); err == nil {
			t.Fatal("expected error for incompressible short final chunk, got nil")
		}
	})
}
