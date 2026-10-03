package oforms

import (
	"os"
	"testing"

	"github.com/harumiWeb/xlflow/internal/pack/cfb"
)

// Developer inspection reads saved bytes only; Excel capture is external.
func TestInspectMultiPageEvidence(t *testing.T) {
	path := os.Getenv("XLFLOW_MULTIPAGE_EVIDENCE")
	if path == "" {
		t.Skip("developer-only saved binary inspection")
	}
	body, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	c, err := cfb.Open(body)
	if err != nil {
		t.Fatal(err)
	}
	names := DiscoverForms(c)
	for _, name := range names {
		f, err := ReadForm(c, name, 932)
		if err != nil {
			t.Fatal(err)
		}
		for _, l := range f.Levels {
			t.Logf("LEVEL %s meta=%x props=%v size=%v x=%x comp=%x trailing=%x", l.Path, l.StorageMeta.CLSID, l.Record.Values, l.Record.Sizes, l.XRaw, l.CompObjRaw, l.TrailingRaw)
			for _, control := range l.Controls {
				t.Logf("CONTROL %s class=%d id=%d site=%v pos=%v", control.Name, control.CLSIDCacheIndex, control.ID, control.Site.Values, control.Site.Position)
				if control.Record != nil && control.Level == nil {
					t.Logf("RECORD %s values=%v size=%v arrays=%x tail=%x text=%v", control.Name, control.Record.Values, control.Record.Sizes, control.Record.Arrays, control.Record.TailRaw, control.Record.TextProps)
				}
			}
		}
	}
}
