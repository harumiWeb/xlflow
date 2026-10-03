package ovba

import (
	"bytes"
	"encoding/binary"
	"strings"
	"testing"
)

func projectReferenceFixture() []byte {
	// REFERENCEPROJECT shape and paths from MS-OVBA section 3.1.2.2.
	abs := []byte(`*\CC:\Example Path\Example-ReferencedProject.xls`)
	rel := []byte(`*\CExample-ReferencedProject.xls`)
	payload := binary.LittleEndian.AppendUint32(nil, uint32(len(abs)))
	payload = append(payload, abs...)
	payload = binary.LittleEndian.AppendUint32(payload, uint32(len(rel)))
	payload = append(payload, rel...)
	payload = binary.LittleEndian.AppendUint32(payload, 0x49A95F46)
	payload = binary.LittleEndian.AppendUint16(payload, 0x000D)
	return sizedRecord(0x000E, payload)
}

func TestReferenceAggregateSizesAndProjectBoundaries(t *testing.T) {
	registered, err := BuildProjectReferences([]RegisteredReferenceSpec{{Name: "stdole", LibID: `*\G{00020430-0000-0000-C000-000000000046}#2.0#0##`}}, 1252)
	if err != nil {
		t.Fatal(err)
	}
	info, err := BuildProjectInformation(ProjectInformationSpec{SysKind: 1, CodePage: 1252, Name: "Test"})
	if err != nil {
		t.Fatal(err)
	}
	modules, err := BuildProjectModules(nil, 1252)
	if err != nil {
		t.Fatal(err)
	}
	for _, raw := range [][]byte{registered, projectReferenceFixture()} {
		for _, aggregate := range []uint32{0, 1, 0xFFFFFFFF} {
			modified := bytes.Clone(raw)
			offset := 0
			for binary.LittleEndian.Uint16(modified[offset:]) == 0x0016 || binary.LittleEndian.Uint16(modified[offset:]) == 0x003E {
				offset += 6 + int(binary.LittleEndian.Uint32(modified[offset+2:]))
			}
			binary.LittleEndian.PutUint32(modified[offset+2:], aggregate)
			refs, err := ParseProjectReferences(modified, 1252)
			if err != nil || len(refs) != 1 || !bytes.Equal(refs[0].Raw, modified) {
				t.Fatalf("aggregate %x: %v", aggregate, err)
			}
			// A following registered group and module boundary must stay aligned.
			combined := append(bytes.Clone(modified), registered...)
			plain := append(append(bytes.Clone(info), combined...), modules...)
			dir, err := ParseDir(plain)
			if err != nil || !bytes.Equal(dir.RefsRaw, combined) || !bytes.Equal(dir.ProjectInfoRaw, info) {
				t.Fatalf("dir aggregate %x: %v", aggregate, err)
			}
		}
	}
	project := projectReferenceFixture()
	for n := 1; n < len(project); n++ {
		if _, err := ParseProjectReferences(project[:n], 1252); err == nil {
			t.Fatalf("accepted truncated PROJECT at %d", n)
		}
		if _, err := walkRecords(project[:n]); err == nil {
			t.Fatalf("walker accepted truncated PROJECT at %d", n)
		}
	}
	for _, offset := range []int{6, 10 + int(binary.LittleEndian.Uint32(project[6:]))} {
		bad := bytes.Clone(project)
		binary.LittleEndian.PutUint32(bad[offset:], 0xFFFFFFFF)
		if _, err := ParseProjectReferences(bad, 1252); err == nil {
			t.Fatal("accepted oversized PROJECT path")
		}
	}
}

func TestReferenceLIBIDGrammar(t *testing.T) {
	base := `*\G{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0##`
	for _, valid := range []string{base, strings.Replace(base, "#2.0#0#", "#a.FFFF#FFFFFFFF#", 1), strings.Replace(base, `*\G`, `*\H`, 1), base + strings.Repeat("a", 255), base + "name#with#hash"} {
		if err := validateReferenceLIBID([]byte(valid)); err != nil {
			t.Fatalf("valid LIBID rejected: %v", err)
		}
	}
	for _, invalid := range []string{
		strings.Replace(base, "#2.0#", "#bogus#", 1),
		strings.Replace(base, "#2.0#", "#10000.0#", 1),
		strings.Replace(base, "#2.0#", "#.0#", 1),
		strings.Replace(base, "#2.0#0#", "#2.0#FFFFFFFFF#", 1),
		strings.Replace(base, "#2.0#0#", "#2.0#Z#", 1),
		strings.Replace(base, `*\G`, `*\X`, 1),
		strings.Replace(base, "0D452EE1", "ZD452EE1", 1),
		strings.TrimSuffix(base, "#"), base + strings.Repeat("a", 256),
	} {
		raw, err := BuildProjectReferences([]RegisteredReferenceSpec{{Name: "MSForms", LibID: invalid}}, 1252)
		if err != nil {
			t.Fatal(err)
		}
		if _, err := ParseProjectReferences(raw, 1252); err == nil {
			t.Fatalf("malformed LIBID accepted: %s", invalid)
		}
	}
}
