# FormSpec examples

Use these complete specs as starting points. [forms.md](forms.md) is the main
create/edit guide; [forms-properties.md](forms-properties.md) lists supported
properties. Keep only cataloged fields.

Geometry is parent-relative points: root controls use the UserForm; children
use their Frame or Page. Keep IDs stable; `parentId` refers to an `id`. Page
bounds come from the MultiPage, so omit Page geometry.

## Contact form

Shows Label, TextBox, CheckBox, and CommandButton with typed appearance fields;
`textAlign` is `1` left, `2` center, or `3` right.

```yaml
schemaVersion: 1
kind: xlflow.userform
basis: designer
coordinateSystem: points
form:
  name: ContactForm
  caption: Contact
  width: 320
  height: 190
  font:
    name: Segoe UI
    size: 9
  foreColor:
    kind: rgb
    value: "#202020"
  backColor:
    kind: rgb
    value: "#F4F6FA"
controls:
  - id: contact-title
    name: Title
    type: Label
    left: 12
    top: 10
    width: 220
    height: 24
    caption: Contact us
    font:
      name: Segoe UI
      size: 14
      bold: true
    foreColor:
      kind: rgb
      value: "#174A7E"
    textAlign: 1
  - id: name-label
    name: NameLabel
    type: Label
    left: 12
    top: 48
    width: 76
    height: 20
    caption: Name
    textAlign: 1
  - id: name-input
    name: NameInput
    type: TextBox
    left: 96
    top: 46
    width: 204
    height: 22
    font:
      name: Segoe UI
      size: 9
    foreColor:
      kind: rgb
      value: "#202020"
    backColor:
      kind: rgb
      value: "#FFFFFF"
    textAlign: 1
  - id: email-label
    name: EmailLabel
    type: Label
    left: 12
    top: 80
    width: 76
    height: 20
    caption: Email
    textAlign: 1
  - id: email-input
    name: EmailInput
    type: TextBox
    left: 96
    top: 78
    width: 204
    height: 22
    textAlign: 1
  - id: updates-opt-in
    name: UpdatesOptIn
    type: CheckBox
    left: 12
    top: 112
    width: 220
    height: 22
    caption: Send me occasional updates
    value: false
    textAlign: 1
  - id: submit-contact
    name: SubmitContact
    type: CommandButton
    left: 216
    top: 148
    width: 84
    height: 26
    caption: Send
    font:
      name: Segoe UI
      size: 9
      bold: true
    foreColor:
      kind: rgb
      value: "#FFFFFF"
    backColor:
      kind: rgb
      value: "#1769AA"
```

## OptionButtons grouped by a Frame

A shared Frame parent groups OptionButtons; positions are Frame-relative, with
no `groupName` property.

```yaml
schemaVersion: 1
kind: xlflow.userform
basis: designer
coordinateSystem: points
form:
  name: ContactPreference
  caption: Contact preference
  width: 280
  height: 150
controls:
  - id: contact-method-frame
    name: ContactMethodFrame
    type: Frame
    left: 12
    top: 12
    width: 244
    height: 100
    caption: Preferred contact method
  - id: contact-by-email
    parentId: contact-method-frame
    name: ContactByEmail
    type: OptionButton
    left: 12
    top: 22
    width: 190
    height: 22
    caption: Email
    value: true
  - id: contact-by-phone
    parentId: contact-method-frame
    name: ContactByPhone
    type: OptionButton
    left: 12
    top: 54
    width: 190
    height: 22
    caption: Phone
    value: false
```

## MultiPage, Pages, Image, and optional TabStrip

Before applying this spec, create `src/forms/assets/brand.bmp` under the project
root. Use a valid project-relative BMP or JPEG asset path.

```yaml
schemaVersion: 1
kind: xlflow.userform
basis: designer
coordinateSystem: points
form:
  name: ProductDetails
  caption: Product details
  width: 440
  height: 280
controls:
  - id: detail-pages
    name: DetailPages
    type: MultiPage
    left: 12
    top: 12
    width: 296
    height: 212
    selectedIndex: 0
  - id: overview-page
    parentId: detail-pages
    name: OverviewPage
    type: Page
    caption: Overview
  - id: overview-label
    parentId: overview-page
    name: OverviewLabel
    type: Label
    left: 12
    top: 12
    width: 180
    height: 20
    caption: Product overview
  - id: brand-page
    parentId: detail-pages
    name: BrandPage
    type: Page
    caption: Brand
  - id: brand-image
    parentId: brand-page
    name: BrandImage
    type: Image
    left: 12
    top: 12
    width: 112
    height: 72
    picture:
      path: src/forms/assets/brand.bmp
    pictureSizeMode: 3
  # Optional standalone TabStrip; its tabs are records, not Page controls.
  - id: detail-tabs
    name: DetailTabs
    type: TabStrip
    left: 324
    top: 12
    width: 104
    height: 90
    selectedIndex: 0
    tabs:
      - name: Summary
        caption: Summary
      - name: Artwork
        caption: Artwork
```

The TabStrip is independent; connect its selection to Pages in form code. Omit
it if unused. On edits, omitted `tabs` preserves them; `tabs: []` removes them.

## Adapting these specs

- The `controls` list replaces the collection; omitted existing controls are removed.
- Omitted properties on retained controls preserve values; explicit `""`, `false`,
  and `0` are authored. See [forms.md](forms.md) and [forms-properties.md](forms-properties.md).
- Add only cataloged fields. MSForms names like `default`, `cancel`, `multiLine`,
  and `groupName` are not implied or valid arbitrary YAML fields.
