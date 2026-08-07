VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_980_Data_Dictionary_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Creates a Data Dictionary of Access Database
'
' Ver.  Date            Author              Details
' 1.00  21 MAY 2004     Anthony Duguid      Initial version
'--------------------------------------------------------------------------------------------------------------------
Option Explicit

Dim NameOfNewTable As String
Private Sub Cancel_Click()

    DoCmd.Close

End Sub
Private Sub Proceed_Click()
On Error GoTo Err_Proceed_Click

    If DocumentTables() = True Then
        If MsgBox("Documentation completed successfully!" & vbCrLf & vbCrLf & "Do you wish to view the '" & NameOfNewTable & "' table now?", vbYesNo, "Database Properties") = vbYes Then
            DoCmd.OpenTable (NameOfNewTable)
        End If
    Else
        MsgBox "Documentation cancelled.", vbInformation, "Database Properties"
    End If
    
    DoCmd.Close acForm, Me.Name

Exit_Proceed_Click:
    Exit Sub

Err_Proceed_Click:
    MsgBox Err.Description
    Resume Exit_Proceed_Click
    
End Sub

Public Function DocumentTables() As Variant
On Error GoTo Err_DocumentTables
    
    DocumentTables = False
    
    Dim MyDB As DAO.Database
    Dim NewTable As DAO.TableDef, NewField As DAO.Field, NewSet As DAO.Recordset
    Dim t As Integer, f As Integer, NewProp As Property
    Dim NumberOfFields As Integer, Speedo As Variant
    
    Speedo = SysCmd(acSysCmdSetStatus, "Preparing tables")
        
    Set MyDB = DBEngine.Workspaces(0).Databases(0)
    MyDB.TableDefs.Refresh      'Otherwise the db may think there are several of these tables
    NameOfNewTable = [ChosenTableName]
        
   'Delete result table if it already exists
    For t = 0 To MyDB.TableDefs.Count - 1
        If MyDB.TableDefs(t).Name = NameOfNewTable Then
            If MsgBox("The table '" & NameOfNewTable & "' already exists." & Chr$(10) & Chr$(13) & Chr$(10) & Chr$(13) & "Do you wish to overwrite it?", vbOKCancel, "Table already exists") = vbOK Then
                MyDB.TableDefs.Delete NameOfNewTable
                Exit For
            Else
                GoTo Exit_DocumentTables
            End If
        End If
    Next t
    
   'Work out how many fields have to be added and start
    NumberOfFields = 0
    For t = 0 To MyDB.TableDefs.Count - 1
        If MyDB.TableDefs(t).Name <> NameOfNewTable Or Not MyDB.TableDefs(t).Name Like "MSys*" Or Not MyDB.TableDefs(t).Name = "ORACLE_ALL_WRDS_tbl" Then
             For f = 0 To MyDB.TableDefs(t).Fields.Count - 1
                  NumberOfFields = NumberOfFields + 1
             Next f
        End If
    Next t
    Speedo = SysCmd(SYSCMD_INITMETER, "Documenting fields", NumberOfFields)
    DoCmd.Hourglass True
    
   'Create table to accept results of documentation
    Set NewTable = MyDB.CreateTableDef(NameOfNewTable)
       'Create fields
       
        Set NewField = NewTable.CreateField("TBL_NME", dbText, 255)
        NewTable.Fields.Append NewField
        
        Set NewField = NewTable.CreateField("FLD_NME", dbText, 255)
        NewTable.Fields.Append NewField
        
        If [ShowType] Then
            Set NewField = NewTable.CreateField("FLD_TYPE", dbText, 20)
            NewTable.Fields.Append NewField
    
        End If
        If [ShowSize] Then
            Set NewField = NewTable.CreateField("FLD_SIZE", dbByte, 100)
'            Set [FLD_SIZE].Properties("Description") = NewTable.CreateProperty("Description") = "Field Size"
            NewTable.Fields.Append NewField
               ' NewTable.CreateProperty("Description", , "FLD_SIZE") = "Field Size"
              'NewTable![FLD_SIZE].Properties("Description") = "Field Size" 'to document the field
        End If
      
        If [ShowDescription] Then
            Set NewField = NewTable.CreateField("FLD_DESC", dbText, 255)
            NewTable.Fields.Append NewField
        End If
         If [ShowOrdinalPosition] Then
            Set NewField = NewTable.CreateField("FLD_POS", dbLong, 8)
            NewTable.Fields.Append NewField
        End If
         If [ShowCollatingOrder] Then
            Set NewField = NewTable.CreateField("FLD_COLL_ORD", dbText, 20)
            NewTable.Fields.Append NewField
        End If
        If [ShowFormat] Then
            Set NewField = NewTable.CreateField("FLD_FORMAT", dbText, 20)
            NewTable.Fields.Append NewField
        End If
        If [ShowInputMask] Then
            Set NewField = NewTable.CreateField("FLD_MASK", dbText, 20)
            NewTable.Fields.Append NewField
        End If
        If ShowCaption Then
            Set NewField = NewTable.CreateField("FLD_CAPTION", dbText, 100)
            NewTable.Fields.Append NewField
        End If
        If [ShowDefaultValue] Then
            Set NewField = NewTable.CreateField("FLD_DEF_VAL", dbText, 20)
            NewField.AllowZeroLength = -1
            NewTable.Fields.Append NewField
        End If
        If [ShowValidationRule] Then
            Set NewField = NewTable.CreateField("FLD_VALD_RULE", dbText, 20)
            NewField.AllowZeroLength = -1
            NewTable.Fields.Append NewField
        End If
        If [ShowValidationText] Then
            Set NewField = NewTable.CreateField("FLD_VALD_TXT", dbText, 20)
            NewField.AllowZeroLength = -1
            NewTable.Fields.Append NewField
        End If
        If [ShowRequired] Then
            Set NewField = NewTable.CreateField("FLD_REQ", dbText, 20)
            NewField.AllowZeroLength = -1
            NewTable.Fields.Append NewField
        End If
        If [ShowAllowZeroLength] Then
            Set NewField = NewTable.CreateField("FLD_ZERO", dbText, 20)
            NewField.AllowZeroLength = -1
            NewTable.Fields.Append NewField
        End If
        If [ShowIndexed] Then
            Set NewField = NewTable.CreateField("FLD_INDEX", dbText, 20)
            NewTable.Fields.Append NewField
        End If

        
        Set NewField = NewTable.CreateField("FLD_SPACES", dbBoolean)
        NewTable.Fields.Append NewField
        ' FLD_ORC_MTCH
        
         If [ShowDateRun] Then
            Set NewField = NewTable.CreateField("DTE_RUN", dbDate, 30)
            NewTable.Fields.Append NewField
        End If
    MyDB.TableDefs.Append NewTable
    
   'Open new table to accept data
    Set NewSet = MyDB.OpenRecordset(NewTable.Name, DB_OPEN_DYNASET)
    
    NumberOfFields = 0      'Reset counter
   'For each table which is not a system table or the target table
    For t = 0 To MyDB.TableDefs.Count - 1
        If MyDB.TableDefs(t).Name <> NewTable.Name And Not MyDB.TableDefs(t).Name Like "MSys*" Then
           'For each field in the table
            For f = 0 To MyDB.TableDefs(t).Fields.Count - 1
                NumberOfFields = NumberOfFields + 1
                Speedo = SysCmd(SYSCMD_UPDATEMETER, NumberOfFields)
                NewSet.AddNew
                NewSet![TBL_NME] = MyDB.TableDefs(t).Fields(f).SourceTable
                NewSet![FLD_NME] = MyDB.TableDefs(t).Fields(f).Name
    Dim strTempfld As String
                strTempfld = MyDB.TableDefs(t).Fields(f).Name
                If [ShowType] Then NewSet![FLD_TYPE] = FieldType(MyDB.TableDefs(t).Fields(f).Type, [UsePlainEnglish])
               'FieldType function converts integer value of type to a text description which is either standard AccessSpeak (false), or PlainEnglish (true)
                If [ShowSize] Then NewSet![FLD_SIZE] = MyDB.TableDefs(t).Fields(f).Size
                If [ShowDescription] Then NewSet![FLD_DESC] = MyDB.TableDefs(t).Fields(f).Properties("Description")
                If [ShowOrdinalPosition] Then NewSet![FLD_POS] = MyDB.TableDefs(t).Fields(f).OrdinalPosition
                If [ShowCollatingOrder] Then NewSet![FLD_COLL_ORD] = MyDB.TableDefs(t).Fields(f).CollatingOrder
                If [ShowFormat] Then NewSet![FLD_FORMAT] = MyDB.TableDefs(t).Fields(f).Properties("Format")
                If [ShowDefaultValue] Then NewSet![FLD_DEF_VAL] = MyDB.TableDefs(t).Fields(f).DefaultValue
                If [ShowRequired] Then NewSet![FLD_REQ] = MyDB.TableDefs(t).Fields(f).Required
                If [ShowAllowZeroLength] Then NewSet![FLD_ZERO] = MyDB.TableDefs(t).Fields(f).AllowZeroLength
                'If [ShowIndexed] Then NewSet![FLD_INDEX] = CheckFieldIndex(MyDb.TableDefs(t).Name, MyDb.TableDefs(t).Fields(f).Name)
                If [ShowDateRun] Then NewSet![DTE_RUN] = Now()                      'Current Date
                'NewSet![FLD_ORC_MTCH] = OracleWrd(strTempfld)                       'Oracle Word Matches
                NewSet.Update
                
               'The following lines are required somewhere but I don't know
               'if this is for each database or each table or what
                'Set NewProp = MyDb.TableDefs(T).Fields(F).CreateProperty("Description")
                'MyDb.TableDefs(T).Fields(F).Properties.Append NewProp
                'Set NewProp = MyDb.TableDefs(T).Fields(F).CreateProperty("Caption")
                'MyDb.TableDefs(T).Fields(F).Properties.Append NewProp
                
            Next f
        End If
    Next t
    NewSet.Close
    MyDB.Close
    
    DocumentTables = True

Exit_DocumentTables:
    Speedo = SysCmd(SYSCMD_REMOVEMETER)
    DoCmd.Hourglass False
    Exit Function

Err_DocumentTables:
    Select Case Err.Number

        Case 3270       'Property doesn't exits so just ignore
                        'E.g. a caption for a date field
            Resume Next
        Case 3315       'Can't be a zero length string
                        'If ignore, nothing is stored so this is OK
            Resume Next
        Case Else
            MsgBox "Unexpected Error - " & Err.Description & " (" & Err.Number & ")"
            Resume Exit_DocumentTables

    End Select

End Function


Function FieldType(intType As Integer, PlainEnglish As Variant) As String

    Select Case intType
        Case dbBoolean
           ' If PlainEnglish Then
                FieldType = "BOOLEAN"
           ' Else
           '     FieldType = "dbBoolean"
           ' End If
        Case dbByte
           ' If PlainEnglish Then
                FieldType = "BYTE"
           ' Else
           '     FieldType = "dbByte"
           ' End If
        Case dbInteger
            'If PlainEnglish Then
                FieldType = "INTEGER"
            'Else
            '    FieldType = "dbInteger"
            'End If
        Case dbLong
            'If PlainEnglish Then
                FieldType = "LONG"
            'Else
            '    FieldType = "dbLong"
            'End If
        Case dbCurrency
            'If PlainEnglish Then
                FieldType = "CURRENCY"
           ' Else
              '  FieldType = "dbCurrency"
          '  End If
        Case dbSingle
            'If PlainEnglish Then
                FieldType = "SINGLE"
           ' Else
            '    FieldType = "dbSingle"
           ' End If
        Case dbDouble
            'If PlainEnglish Then
                FieldType = "DOUBLE"
           ' Else
             '   FieldType = "dbDouble"
          '  End If
        Case dbDate
            'If PlainEnglish Then
                FieldType = "DATE"
           ' Else
           '     FieldType = "dbDate"
          '  End If
        Case dbText
            'If PlainEnglish Then
                FieldType = "VARCHAR2"
            'Else
            '    FieldType = "dbText"
           ' End If
        Case dbLongBinary
           ' If PlainEnglish Then
                FieldType = "OLE object"
          '  Else
          '      FieldType = "dbLongBinary"
           ' End If
        Case dbMemo
            'If PlainEnglish Then
                FieldType = "Memo"
            'Else
            '    FieldType = "dbMemo"
           ' End If
        Case dbGUID
            FieldType = "dbGUID"
        Case dbDecimal
            FieldType = "DECIMAL"
    End Select

End Function

Function SPACES(ByVal strTextVal As String)
        
Dim intPtr As Integer
   
            intPtr = InStr(strTextVal, " ")
            If intPtr > 0 Then
                'MsgBox "Oracle does not like spaces"
                'Produce the value for a textbox here "spaces = true"
                SPACES = True
            End If

End Function
