package ovba

import (
	"bytes"
	"encoding/binary"
	"fmt"
)

// ProjectInformationSpec is the modeled PROJECTINFORMATION section of dir.
type ProjectInformationSpec struct {
	SysKind  uint32
	LCID     uint32
	CodePage uint16
	Name     string
}

// RegisteredReferenceSpec is one registered type-library reference.
type RegisteredReferenceSpec struct {
	Name  string
	LibID string
}

// BuildProjectInformation emits the required PROJECTINFORMATION records for a
// source-only project. Defaults that are not part of the public blank profile
// are intentionally centralized here and written in the order required by
// MS-OVBA.
func BuildProjectInformation(spec ProjectInformationSpec) ([]byte, error) {
	if spec.SysKind == 0 {
		return nil, fmt.Errorf("ovba: project SYSKIND must not be zero")
	}
	if spec.LCID == 0 {
		return nil, fmt.Errorf("ovba: project LCID must not be zero")
	}
	name, err := EncodeMBCS(spec.Name, spec.CodePage)
	if err != nil {
		return nil, fmt.Errorf("ovba: project name: %w", err)
	}
	if len(name) == 0 {
		return nil, fmt.Errorf("ovba: project name must not be empty")
	}

	var out []byte
	out = append(out, recU32(0x0001, 4, spec.SysKind)...)     // PROJECTSYSKIND
	out = append(out, recU32(0x004A, 4, 0x00000005)...)       // PROJECTCOMPATVERSION
	out = append(out, recU32(0x0002, 4, spec.LCID)...)        // PROJECTLCID
	out = append(out, recU32(0x0014, 4, 0x00000409)...)       // PROJECTLCIDINVOKE
	out = append(out, recU16(0x0003, 2, spec.CodePage)...)    // PROJECTCODEPAGE
	out = append(out, sizedRecord(0x0004, name)...)           // PROJECTNAME
	out = append(out, sizedRecord(0x0005, nil)...)            // PROJECTDOCSTRING
	out = append(out, sizedRecord(0x0040, nil)...)            // PROJECTDOCSTRINGUNICODE
	out = append(out, sizedRecord(0x0006, nil)...)            // PROJECTHELPFILEPATH
	out = append(out, sizedRecord(0x003D, nil)...)            // PROJECTHELPFILEPATH unicode copy
	out = append(out, recU32(0x0007, 4, 0)...)                // PROJECTHELPCONTEXT
	out = append(out, recU32(0x0008, 4, 0)...)                // PROJECTLIBFLAGS
	out = append(out, projectVersionRecord(0x6CDFF001, 7)...) // PROJECTVERSION
	out = append(out, sizedRecord(0x000C, nil)...)            // PROJECTCONSTANTS
	out = append(out, sizedRecord(0x003C, nil)...)            // PROJECTCONSTANTSUNICODE
	return out, nil
}

func projectVersionRecord(major uint32, minor uint16) []byte {
	out := make([]byte, 12)
	binary.LittleEndian.PutUint16(out[0:], 0x0009)
	binary.LittleEndian.PutUint32(out[2:], 0x00000004)
	binary.LittleEndian.PutUint32(out[6:], major)
	binary.LittleEndian.PutUint16(out[10:], minor)
	return out
}

// BuildProjectReferences emits registered-reference records without relying
// on bytes copied from an existing workbook.
func BuildProjectReferences(specs []RegisteredReferenceSpec, codepage uint16) ([]byte, error) {
	var out []byte
	for _, spec := range specs {
		name, err := EncodeMBCS(spec.Name, codepage)
		if err != nil {
			return nil, fmt.Errorf("ovba: reference %q name: %w", spec.Name, err)
		}
		libID, err := EncodeMBCS(spec.LibID, codepage)
		if err != nil {
			return nil, fmt.Errorf("ovba: reference %q LIBID: %w", spec.Name, err)
		}
		if len(name) == 0 || len(libID) == 0 {
			return nil, fmt.Errorf("ovba: reference name and LIBID must not be empty")
		}
		out = append(out, sizedRecord(0x0016, name)...)               // REFERENCENAME
		out = append(out, sizedRecord(0x003E, utf16le(spec.Name))...) // REFERENCENAMEUNICODE

		payload := make([]byte, 4+len(libID)+6)
		binary.LittleEndian.PutUint32(payload, uint32(len(libID)))
		copy(payload[4:], libID)
		// Reserved1 (uint32) and Reserved2 (uint16) remain zero.
		out = append(out, sizedRecord(0x000D, payload)...) // REFERENCEREGISTERED
	}
	return out, nil
}

// BuildProjectText emits a deterministic textual PROJECT stream for a fresh
// source-only project.
func BuildProjectText(name string, specs []ProjectComponentSpec, codepage uint16) ([]byte, error) {
	encodedName, err := EncodeMBCS(name, codepage)
	if err != nil {
		return nil, fmt.Errorf("ovba: PROJECT name: %w", err)
	}
	if len(encodedName) == 0 {
		return nil, fmt.Errorf("ovba: PROJECT name must not be empty")
	}

	var out bytes.Buffer
	// These deterministic unprotected-project values are an internally
	// consistent set observed from Excel. CMG/DPB/GC include a project-key
	// checksum, so they must remain paired with this project ID.
	out.WriteString("ID=\"{61CB3C72-4521-44A9-8538-FAE59C5B1642}\"\r\n")
	for _, spec := range specs {
		if !isProjectComponentKind(spec.Kind) {
			return nil, fmt.Errorf("PROJECT component %q has unknown kind %q", spec.Name, spec.Kind)
		}
		encoded, err := EncodeMBCS(spec.Name, codepage)
		if err != nil {
			return nil, fmt.Errorf("PROJECT component %q: %w", spec.Name, err)
		}
		out.WriteString(spec.Kind)
		out.WriteByte('=')
		out.Write(encoded)
		if spec.Kind == "Document" {
			out.WriteString("/&H00000000")
		}
		out.WriteString("\r\n")
	}
	out.WriteString("Name=\"")
	out.Write(encodedName)
	out.WriteString("\"\r\nHelpContextID=\"0\"\r\nVersionCompatible32=\"393222000\"\r\n")
	out.WriteString("CMG=\"7577CB4035B139B139B139B139\"\r\n")
	out.WriteString("DPB=\"EAE854D3C8D4C8D4C8\"\r\n")
	out.WriteString("GC=\"5F5DE16E23969997999766\"\r\n\r\n")
	out.WriteString("[Host Extender Info]\r\n")
	out.WriteString("&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000\r\n\r\n")
	out.WriteString("[Workspace]\r\n")
	for _, spec := range specs {
		encoded, _ := EncodeMBCS(spec.Name, codepage)
		out.Write(encoded)
		if spec.Kind == "Document" {
			out.WriteString("=0, 0, 0, 0, C\r\n")
		} else {
			out.WriteString("=32, 32, 1872, 983, Z\r\n")
		}
	}
	return out.Bytes(), nil
}

// BuildProjectWM emits the MBCS/Unicode module-name mapping for PROJECTwm.
func BuildProjectWM(specs []ProjectComponentSpec, codepage uint16) ([]byte, error) {
	var out []byte
	for _, spec := range specs {
		mbcs, err := EncodeMBCS(spec.Name, codepage)
		if err != nil {
			return nil, fmt.Errorf("ovba: PROJECTwm component %q: %w", spec.Name, err)
		}
		out = append(out, mbcs...)
		out = append(out, 0)
		out = append(out, utf16le(spec.Name)...)
		out = append(out, 0, 0)
	}
	// A final empty Unicode name terminates the mapping after the last MBCS/
	// Unicode pair. Together with that pair's UTF-16 NUL this leaves the exact
	// five-zero suffix emitted by Excel.
	return append(out, 0, 0), nil
}
