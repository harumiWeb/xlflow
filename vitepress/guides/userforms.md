# Manage UserForms

Use `form snapshot` to capture Designer state, `form build` to create from a persisted specification, and `form export-image` to verify the rendered result. In sidecar mode, Designer structure lives in `src/forms/specs/` and code-behind in `src/forms/code/`.

```bash
xlflow form snapshot UserForm1 --json
xlflow form build src/forms/specs/UserForm1.yaml --json
xlflow form export-image UserForm1 --json
```

Run `lint` and `form build` preflight before opening Excel. Do not edit generated `.frm` code-behind independently when sidecar mode is authoritative; see the [UserForm specification](../reference/userform-spec).

## MultiPage and TabStrip

Keep the authored `controls` list flat and connect a `MultiPage` to its `Page`
controls with `parentId`; put page content controls under each `Page`. A
`TabStrip` is separate: it has its own `tabs` array and does not contain
controls. Each tab has a required `name` and may define `caption`,
`controlTipText`, `tag`, `accelerator`, `enabled`, and `visible`. Omit `tabs`
to leave an existing collection unspecified; use `tabs: []` to author an empty
collection.

`selectedIndex` is zero-based for a nonempty Page or Tab collection. The
normalized no-selection value is `-1`, including when the collection is empty.
Page geometry comes from its owning MultiPage, so do not author Page
`left`/`top`/`width`/`height`. See the [UserForm specification](../reference/userform-spec)
for the complete schema, snapshot shape, and generation details.

## Image pictures

An Image control can load a project-root-relative BMP or JPEG asset in pure-Go
`pack` and `push --backend file`:

```yaml
- id: brand-mark
  name: BrandMark
  type: Image
  picture:
    path: src/forms/assets/brand.jpg
```

Assets are confined to the project after resolving symlinks and junctions, and
are limited to 16 MiB and 16 million pixels. Use `picture: { remove: true }`
to clear an existing picture. Omitting `picture` preserves it. File pull stores
supported embedded images in `src/forms/assets/<full-SHA-256>.bmp|jpg`; those
files are not automatically deleted. File push loads, validates, and
fingerprints referenced assets only; unreferenced assets are retained and
ignored. Excel may normalize JPEG input to BMP when saving. Excel-backed `form build` and apply APIs reject `picture` until
they implement the same contract. See the [UserForm specification](../reference/userform-spec)
for invalid values and full field details.
