# UserForm binary corpus checks

The current Linux regression set consumes the committed Excel-authored
projects under `../compiler/testdata` plus the historical `p4_form.bin` and
`p6_nested_form.bin` CFB fixtures. It checks lossless Designer stream/storage
replay, logical FormSpec round trips, and a 512-control generated stress case.

The committed Excel-authored cases cover common controls, Frame and nested
Frame, MultiPage with nested Page controls, empty MultiPage state, Japanese
text, and property edits. A standalone picture/resource-bearing Designer and
a fixture with non-ASCII control names are not in this corpus yet. The 512
control stress case is compiler-generated; it is not represented as an
Excel-authored fixture.

The regression checks use xlflow's Go reader, serializer, projection, and
compiler against the committed fixtures. Go tests and CI require no Python
environment or external parser dependency.
