# UserForm event catalogue provenance

`catalogue.json` is a compact projection of the Windows Excel 16.0 VBE
`CreateEventProc` oracle run in
`tmp_workspaces/userform-events-20261016/vbe-events.json`. It contains the 222
declarations accepted by the VBE for the 15 observed event-host types. The
signature projection replaces only the concrete object name before the event
name with `{object}`; argument types, `ByVal`, implicit `ByRef`, and all other
signature text remain as observed.

The ordinary default flags come from the observed TypeLib event metadata in
`tmp_workspaces/userform-events-20261016/typelib-events.json` (the default flag
is `FUNCFLAG_FDEFAULT`, value 512). Image is the confirmed exception: its
hidden Click and DblClick events are rejected by `CreateEventProc`, while an
actual VBE UI double-click generated `BeforeDragOver`, recorded in
`tmp_workspaces/userform-events-20261016/image-default-code.txt`. Accordingly,
Image exposes its six accepted public events and marks `BeforeDragOver` as its
default. The root UserForm default is Click.

The verification used Excel 16.0 x64 and included the VBE-specific extender
events BeforeUpdate, AfterUpdate, Enter and Exit. All 222 generated declarations
were imported together into a real form, compiled and retained through
save/export. Verification-only local declarations kept otherwise empty event
procedures from being removed by Excel during saving. Initialize, TextBox and
ComboBox Change, CheckBox Click and MultiPage Change also produced asserted
worksheet sentinel values. Both sidecar and frm source modes were exercised;
frm insertion preserved the Designer header and FRX bytes.

The oracle runs confirmed cleanup of their owned Excel processes, and the
managed verification session was stopped after saving. Local commands, results
and limitations are recorded in
`tmp_workspaces/userform-events-20261016/verification.md`. These ignored local
artifacts are supporting evidence; the committed catalogue and regression tests
are the portable contract and do not require Excel at runtime.
