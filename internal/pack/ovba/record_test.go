package ovba

import "testing"

func TestSizedRecord(t *testing.T) {
	// RecordId 0x0004 (NAME) + size 10 + "VBAProject".
	got := sizedRecord(0x0004, []byte("VBAProject"))
	want := []byte{
		0x04, 0x00, // id
		0x0A, 0x00, 0x00, 0x00, // size = 10
		'V', 'B', 'A', 'P', 'r', 'o', 'j', 'e', 'c', 't',
	}
	if !bytesEqual(got, want) {
		t.Errorf("sizedRecord = %x, want %x", got, want)
	}
}

func TestUTF16LE(t *testing.T) {
	got := utf16le("Sp")
	want := []byte{'S', 0x00, 'p', 0x00}
	if !bytesEqual(got, want) {
		t.Errorf("utf16le = %x, want %x", got, want)
	}
}
