# Excel-authored Image picture fixture

This developer-only fixture was captured with
`scripts/test-userform-pictures-e2e.ps1 -Phase create` in real Excel. The
capture used Excel 16.0 build 17932 on Windows 10.0.22631. The gate closed
Excel and confirmed that its owned process exited before preserving these
files.

`logo.bmp` and `logo.jpg` are both 24 x 16 pixel source images. Excel loaded
them into `ImageBmp` and a nested `FrameMain.ImageJpeg`, saved the workbook,
and reopened it. Both controls reported StdPicture type 1 and dimensions 508 x
339 HIMETRIC before save and after reopening. The sentinel macro returned
`issue-912-ok`.

Excel persisted both pictures as BMP payloads, including the JPEG-authored
control. The observed StdPicture resource begins with GUID
`0452e30b918fce119de300aa004bb851`, preamble `6c740000`, a payload length of
`0x4b6`, and `BM` image data. Thus this fixture proves that Excel accepts both
source files through `LoadPicture` and normalizes the saved resource to BMP.
It does not prove direct JPEG StdPicture encoding from an xlflow-generated
artifact; that is covered by a separate blank-pack run described below.

The full canonical Issue #912 gate subsequently verified native-JPEG picture
authoring through file push, blank pack, and template pack in real Excel, with
source asset paths unavailable during workbook verification. Evidence is in
`tmp_workspaces/issue-912-file-push-20261004-040504-de315a` and the three
helper workspaces:

- `tmp_workspaces/issue-912-pictures-20261004-040532-dc56f3` (file push)
- `tmp_workspaces/issue-912-pictures-20261004-040537-73f28c` (blank pack)
- `tmp_workspaces/issue-912-pictures-20261004-040541-6d63b5` (template pack)

On Excel 16.0 build 17932 / Windows 10.0.22631, both Image controls reported
StdPicture type 1, positive 508 x 339 HIMETRIC dimensions, picture size mode
3, and alignment 3 before save and after SaveAs/reopen. The sentinel was
`issue-912-ok`, and owned Excel cleanup was confirmed in all three runs. This
supersedes the earlier blank-pack-only evidence note; the Excel-authored
fixture above still records JPEG input normalized to BMP and remains separate
from native-JPEG persistence evidence.

`baseline.bin` is the Excel-authored project and `normalized.bin` is its
save/reopen project. A normalized BMP extracted from the JPEG-authored control
is pixel-equivalent to the original JPEG but is not byte-equal to it. Tests of
the no-op/template carry path should preserve the original persisted picture
resource identity; tests that decode and re-encode should compare equivalent
pixels and Designer behavior rather than the original asset bytes.

The recorded environment and before/after observations are in
`environment.json`. The test is developer-only and is not run by ordinary
tests or CI.
