# MS-OFORMS reader contract

This document defines xlflow's internal, pure-Go reader contract for Microsoft
Forms designer state stored in `vbaProject.bin`. The architectural rationale
for file-level VBA project handling remains in ADR-0012. The user-facing
UserForm schema remains the `xlflow.userform` `FormSpec`; the model described
here is a separate binary persistence model.

## Scope and entry point

`internal/vba/userforms/oforms.ReadForm` reads one top-level UserForm storage
from an already-open `internal/pack/cfb.Container`. Callers supply the storage
path and the VBA project's `PROJECTCODEPAGE`. The reader does not use Excel,
COM, VBIDE, or Windows APIs and does not silently fall back to an Excel-backed
path.

The reader owns the following structures:

- the top-level and nested `f` and `o` streams;
- `FormControl`, `FormSiteData`, `SiteDepthsAndTypes`, and
  `OleSiteConcreteControl` records;
- property masks, aligned DataBlock and ExtraDataBlock fields, StreamData,
  `TextProps`, geometry, names, IDs, and object-stream extents;
- `\x03VBFrame`, `\x01CompObj`, optional `x` streams, CFB storage metadata,
  and nested container storages;
- raw strings, padding, pictures, arrays, record tails, and streams required by
  a future lossless writer.

`DiscoverForms` reports root-level non-`VBA` storages containing an `f` stream.
Module classification remains the responsibility of the VBA project layer.

## Lossless binary model

Decoded values never replace their persistence bytes. Strings retain their
compression flag and original encoded bytes. Records retain masks, alignment
padding, opaque pictures and arrays, raw tails, and the complete source record.
Each container level retains its original streams and storage metadata.

The first known property tables cover the built-in classes represented by the
existing FormSpec contract: Label, TextBox, ComboBox, ListBox, CommandButton,
CheckBox, OptionButton, and Frame. The shared reader also recognizes Image,
MorphData, SpinButton, TabStrip, ToggleButton, ScrollBar, Form/Page, and
MultiPage cache indices so their structural boundaries and container ownership
can be validated.

An unsupported class remains an opaque control only when its site record and
`ObjectStreamSize` unambiguously delimit all of its bytes. Unknown fields
inside a known record are accepted only when its declared record boundary still
reconciles exactly; otherwise parsing fails.

This package does not project into FormSpec, publish files during `pull`, edit
records, or serialize streams. Those are separate stages covered by the
follow-up UserForm issues.

## Encoding and geometry

Compressed MS-OFORMS strings and `\x03VBFrame` use the VBA project code page.
Unsupported code pages and undecodable byte sequences fail. Uncompressed
strings use little-endian UTF-16 and an odd byte count fails. CompObj identity
strings use the same supplied project code page while their complete bytes are
also retained.

Control positions and sizes remain signed HIMETRIC integers in the binary
model. Conversion to FormSpec points belongs to the projection layer.

## Structural validation

The reader returns `oforms.ErrMalformed`, wrapped by an `oforms.ParseError`
carrying storage path, stream, byte offset, and structure, when any of these
invariants fail:

- required top-level `f`, `o`, `\x03VBFrame`, or `\x01CompObj` streams are
  missing;
- a version, length-prefixed field, record boundary, or alignment step leaves
  its enclosing stream;
- expanded `SiteDepthsAndTypes` entries do not equal `CountOfSites`;
- `CountOfBytes` does not end at the expected site-data boundary;
- the sum of site `ObjectStreamSize` values differs from the `o` stream size;
- a container site lacks its matching `i<site-id>` storage, a non-container
  claims one, two storages claim the same ID, or a child storage has no owner;
- a known property record leaves unexplained bytes outside an explicitly
  supported opaque tail.

The resource limits are 64 MiB per designer stream, 65,535 sites per form, and
64 nested container levels. Counts and lengths are checked against remaining
input before allocating or slicing.

## Verification

The committed Excel-authored `p4_form.bin` fixture covers a simple form and
known control property record. `p6_nested_form.bin` covers Frame, MultiPage,
Page, and nested container storages. Focused corruption tests cover missing and
mis-sized `o` streams, inconsistent depth runs, orphan storages, unsupported
code pages, and structurally bounded opaque controls. Native Go fuzz targets
exercise both the public CFB-to-form reader and the `f` stream parser.
