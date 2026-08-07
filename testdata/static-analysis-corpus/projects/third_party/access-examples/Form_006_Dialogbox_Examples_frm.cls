VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_006_Dialogbox_Examples_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdBrowse_Click()

    Me.txtDir.Value = GetDirectory("Please Select a directory", Me)

End Sub

Private Sub cmdClose_Click()
On Error Resume Next

    DoCmd.Close

End Sub

Private Sub cmdDBBrowse_Click()

    Me.txtDB.Value = FindFile(, , "Database Files", "*.mdb")

End Sub

Private Sub cmdXLBrowse_Click()

    Me.txtXLFIle.Value = FindFile()

End Sub
