Attribute VB_Name = "BStorage"
Option Compare Database
Option Explicit
'
' BStorage V1.1.3
' Managing, writing, and reading of encrypted data to and from table fields.
'
' (c) Gustav Brock, Cactus Data ApS, CPH
' https://github.com/GustavBrock/VBA.Cryptography
'
' Requires:
'   Module  BCrypt
'

' Append a binary field to an existing table.
' Generic function.
' Returns True if the binary field exists or was created.
'
' 2022-04-01. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function AppendBinaryField( _
    ByVal Database As DAO.Database, _
    ByVal TableName As String, _
    ByVal FieldName As String, _
    Optional ByVal FieldSize As Long) _
    As Boolean
    
    Dim Success     As Boolean
    
    Success = AppendStorageField(Database, TableName, FieldName, dbBinary, FieldSize)
    
    AppendBinaryField = Success
    
End Function

' Append a long binary (BLOB - Binary Large OBjects) field to an existing table.
' Generic function.
' Returns True if the long binary field exists or was created.
'
' 2022-04-01. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function AppendLongBinaryField( _
    ByVal Database As DAO.Database, _
    ByVal TableName As String, _
    ByVal FieldName As String) _
    As Boolean
    
    Dim Success     As Boolean
    
    Success = AppendStorageField(Database, TableName, FieldName, dbLongBinary)
    
    AppendLongBinaryField = Success
    
End Function

' Append a long text (Memo) field to an existing table.
' Generic function.
' Returns True if the long text field exists or was created.
'
' 2022-04-01. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function AppendLongTextField( _
    ByVal Database As DAO.Database, _
    ByVal TableName As String, _
    ByVal FieldName As String) _
    As Boolean
    
    Dim Success     As Boolean
    
    Success = AppendStorageField(Database, TableName, FieldName, dbMemo)
    
    AppendLongTextField = Success
    
End Function

' Append a short text field to an existing table.
' Generic function.
' Returns True if the short text field exists or was created.
'
' 2022-04-01. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function AppendShortTextField( _
    ByVal Database As DAO.Database, _
    ByVal TableName As String, _
    ByVal FieldName As String, _
    Optional ByVal FieldSize As Long) _
    As Boolean
    
    Dim Success     As Boolean
    
    Success = AppendStorageField(Database, TableName, FieldName, dbText, FieldSize)
    
    AppendShortTextField = Success
    
End Function

' Append to an existing table a field suitable for storing encrypted data.
' Generic function.
' Returns True if the storage field exists as specified or was created.
'
' The field type can be any of these:
'
'   dbBinary        Binary
'   dbLongBinary    Binary (BLOB - Binary Large OBjects)
'   dbText          Short Text
'   dbMemo          Long Text (Memo)
'
' The field size can be specified for Binary and Short Text only:
'
'   Maximum value Binary:   510 (Default)
'   Minimum value Binary:     1
'   Maximum value Binary:   255 (Default)
'   Minimum value text:       1
'
' 2022-03-17. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function AppendStorageField( _
    ByVal Database As DAO.Database, _
    ByVal TableName As String, _
    ByVal FieldName As String, _
    ByVal FieldType As DAO.DataTypeEnum, _
    Optional ByVal FieldSize As Long) _
    As Boolean
    
    ' Field size limits.
    Const MaximumSizeBinary As Long = 510
    Const MaximumSizeText   As Long = 255
    Const MinimumSize       As Long = 1
    Const DefaultSize       As Long = 0
    
    Dim Table       As DAO.TableDef
    Dim Field       As DAO.Field
    Dim Success     As Boolean
    
    ' Validate field type.
    Select Case FieldType
        Case dbBinary, dbLongBinary, dbText, dbMemo
            ' OK.
        Case Else
            ' Ignore other field types.
            Exit Function
    End Select
            
    ' Validate database and object names.
    If Database Is Nothing Or TableName = "" Or FieldName = "" Then
        ' Nothing to do. Exit.
        Exit Function
    End If
    
    ' Validate table name.
    For Each Table In Database.TableDefs
        If Table.Name = TableName Then
            Exit For
        End If
    Next
    If Table Is Nothing Then
        ' Table name not found in this database.
        Exit Function
    End If
    
    ' Validate field name.
    For Each Field In Table.Fields
        If Field.Name = FieldName Then
            Exit For
        End If
    Next
    
    If Field Is Nothing Then
        ' Create the field.
        Set Field = Table.CreateField(FieldName, FieldType)
        If FieldSize = DefaultSize Then
            Success = True
        ElseIf FieldSize >= MinimumSize Then
            ' Set the size of the field for relevant field types.
            Select Case FieldType
                Case dbBinary
                    If FieldSize <= MaximumSizeBinary Then
                        Field.Size = FieldSize
                        Success = True
                    End If
                Case dbText
                    If FieldSize <= MaximumSizeText Then
                        Field.Size = FieldSize
                        Success = True
                    End If
            End Select
        End If
        If Success = True Then
            Success = False
            Table.Fields.Append Field
            Success = True
        End If
    ElseIf Field.Type = FieldType Then
        ' The field exists.
        If Field.Size = FieldSize Then
            Success = True
        ElseIf FieldSize = DefaultSize Then
            Select Case Field.Type
                Case dbBinary
                    Success = (Field.Size = MaximumSizeBinary)
                Case dbText
                    Success = (Field.Size = MaximumSizeText)
            End Select
        End If
    End If
    
    Set Field = Nothing
    Set Table = Nothing
    
    AppendStorageField = Success
    
End Function

' Read, for an ID, the (casted) byte array from a field of a table in the specified database.
' Returns the value as a byte array if the ID is found.
' Returns an empty byte array, if the ID is not found.
'
' 2022-04-02. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function ReadBinaryField( _
    ByVal Database As DAO.Database, _
    ByVal TableName As String, _
    ByVal FieldName As String, _
    ByVal Id As Long) _
    As Byte()
    
    Dim Records             As DAO.Recordset
    
    Dim Sql                 As String
    Dim Data()              As Byte
    
    ' Validate database and object names.
    If Database Is Nothing Or TableName = "" Or FieldName = "" Then
        ' Nothing to do. Exit.
        Exit Function
    End If
    
    ' Wrap table and field names in brackets.
    TableName = Replace("[{0}]", "{0}", TableName)
    FieldName = Replace("[{0}]", "{0}", FieldName)
    
    Sql = "Select " & FieldName & " From " & TableName & " Where Id = " & Id & ""
    Set Records = CurrentDb.OpenRecordset(Sql, dbOpenDynaset, dbReadOnly)
    
    If Records.RecordCount = 1 Then
        Select Case Records(FieldName).Type
            Case dbBinary, dbLongBinary, dbMemo, dbText
                ' Read data as is.
                If Not IsNull(Records(FieldName).Value) Then
                    Data = Records(FieldName).Value
                End If
        End Select
    Else
        ' Return an empty array.
    End If
    Records.Close
    
    ReadBinaryField = Data
    
End Function

' Read, for an ID, the text from a field of a table in the specified database.
' Returns the value as text if the ID is found.
' Returns an empty string, if the ID is not found.
'
' 2022-04-02. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function ReadTextField( _
    ByVal Database As DAO.Database, _
    ByVal TableName As String, _
    ByVal FieldName As String, _
    ByVal Id As Long) _
    As String
    
    Dim Records             As DAO.Recordset
    
    Dim Sql                 As String
    Dim Text                As String
    
    ' Validate database and object names.
    If Database Is Nothing Or TableName = "" Or FieldName = "" Then
        ' Nothing to do. Exit.
        Exit Function
    End If
    
    ' Wrap table and field names in brackets.
    TableName = Replace("[{0}]", "{0}", TableName)
    FieldName = Replace("[{0}]", "{0}", FieldName)
    
    Sql = "Select " & FieldName & " From " & TableName & " Where Id = " & Id & ""
    Set Records = CurrentDb.OpenRecordset(Sql, dbOpenDynaset, dbReadOnly)
    
    If Records.RecordCount = 1 Then
        Select Case Records(FieldName).Type
            Case dbMemo, dbText
                ' Read data as is.
                If Not IsNull(Records(FieldName).Value) Then
                    Text = Records(FieldName).Value
                End If
        End Select
    Else
        ' Return an empty string.
    End If
    Records.Close
    
    ReadTextField = Text
    
End Function

' Save, for an ID, a byte array to a binary field of a table in the specified database.
' If argument Data() is an empty array, Null will be saved if possible.
' Returns True if success, False if not.
'
' 2022-04-05. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function SaveBinaryField( _
    ByVal Database As DAO.Database, _
    ByVal TableName As String, _
    ByVal FieldName As String, _
    ByVal Id As Long, _
    ByRef Data() As Byte) _
    As Boolean
    
    Dim Records             As DAO.Recordset
    
    Dim Sql                 As String
    Dim Skip                As Boolean
    Dim Success             As Boolean
    
    ' Wrap table and field names in brackets.
    TableName = Replace("[{0}]", "{0}", TableName)
    FieldName = Replace("[{0}]", "{0}", FieldName)
    
    Sql = "Select " & FieldName & " From " & TableName & " Where Id = " & Id & ""
    Set Records = CurrentDb.OpenRecordset(Sql, dbOpenDynaset)
    
    If Records.RecordCount = 1 Then
        Select Case Records(FieldName).Type
            Case dbBinary, dbLongBinary
                If IsNull(Records(FieldName).Value) And StrPtr(Data) = 0 Then
                    ' Nothing to update.
                    Skip = True
                ElseIf StrPtr(Data) = 0 Then
                    ' Set field value to Null.
                ElseIf Records(FieldName).Type = dbBinary And (UBound(Data) - LBound(Data) + 1 > Records(FieldName).Size) Then
                    ' Array is too large to save in this field.
                    Skip = True
                End If
                
                If Not Skip Then
                    Records.Edit
                    If StrPtr(Data) = 0 Then
                        If Not Records(FieldName).Required Then
                            Records(FieldName).Value = Null
                        End If
                    Else
                        Records(FieldName).Value = Data
                    End If
                    Records.Update
                End If
                
                If IsNull(Records(FieldName).Value) And StrPtr(Data) = 0 Then
                    Success = True
                ElseIf UBound(Records(FieldName).Value) - LBound(Records(FieldName).Value) = UBound(Data) - LBound(Data) Then
                    Success = True
                End If
        End Select
    End If
    Records.Close
    
    SaveBinaryField = Success
    
End Function

' Save, for an ID, a text value to a text field of a table in the specified database.
' If argument Text is an empty string, Null will be saved if possible.
' Returns True if success, False if not.
'
' 2022-04-02. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function SaveTextField( _
    ByVal Database As DAO.Database, _
    ByVal TableName As String, _
    ByVal FieldName As String, _
    ByVal Id As Long, _
    ByVal Text As String) _
    As Boolean
    
    Dim Records             As DAO.Recordset
    
    Dim Sql                 As String
    Dim Success             As Boolean
    
    ' Wrap table and field names in brackets.
    TableName = Replace("[{0}]", "{0}", TableName)
    FieldName = Replace("[{0}]", "{0}", FieldName)
    
    Sql = "Select " & FieldName & " From " & TableName & " Where Id = " & Id & ""
    Set Records = CurrentDb.OpenRecordset(Sql, dbOpenDynaset)
    
    If Records.RecordCount = 1 Then
        Select Case Records(FieldName).Type
            Case dbText, dbMemo
                If Nz(Records(FieldName).Value) = Text Then
                    ' Nothing to update.
                ElseIf Records(FieldName).Type = dbText And Len(Text) > Records(FieldName).Size Then
                    ' Text is too long to save in this field.
                Else
                    Records.Edit
                    If Text = "" Then
                        If Not Records(FieldName).Required Then
                            Records(FieldName).Value = Null
                        ElseIf Records(FieldName).AllowZeroLength Then
                            Records(FieldName).Value = ""
                        End If
                    Else
                        Records(FieldName).Value = Text
                    End If
                    Records.Update
                End If
                Success = (Nz(Records(FieldName).Value) = Text)
        End Select
    End If
    Records.Close
    
    SaveTextField = Success
    
End Function

' Decrypt a Base64 encoded encrypted text directly from a Short Text or Long Text table field
' using a passed key.
' Accepts Null values for both data and key.
'
' If success, the decrypted text is returned.
' If no data, no key, or wrong key, Null is returned.
'
' Typical usage:
'
'   Parameters [Key] Text ( 255 );
'   Select *, VDecryptBase64([TextFieldName], [Key]) As Content
'   From TableName
'
' 2022-04-07. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDecryptBase64( _
    ByVal EncryptedText As Variant, _
    Key As Variant) _
    As Variant

    Dim Text        As Variant
    
    If Nz(Key, "") = "" Or Nz(EncryptedText, "") = "" Then
        ' Either no key or no data.
        Text = Null
    Else
        ' Return decrypted data as text.
        Text = Decrypt(EncryptedText, Key)
        If Text = "" Then
            ' Decryption failed.
            Text = Null
        End If
    End If
    
    VDecryptBase64 = Text
    
End Function

' Decrypt an encrypted text directly from a Binary or Long Binary (BLOB) table field
' using a passed key.
' Accepts Null values for both data and key.
'
' If success, the decrypted text is returned.
' If no data, no key, or wrong key, Null is returned.
'
' Typical usage:
'
'   Parameters [Key] Text ( 255 );
'   Select *, VDecryptBinary([BinaryFieldName], [Key]) As Content
'   From TableName
'
' 2022-04-07. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDecryptBinary( _
    ByVal EncryptedData As Variant, _
    Key As Variant) _
    As Variant

    Dim Data()      As Byte
    Dim Text        As Variant
    
    If Nz(Key, "") = "" Or IsNull(EncryptedData) Then
        ' Either no key or no data.
        Text = Null
    ElseIf VarType(Key) <> vbString Or VarType(EncryptedData) <> vbString Then
        ' Key is not text, or - as a binary field will be read as text - ,
        ' EncryptedData holds invalid data.
        Text = Null
    Else
        ' Arguments are valid. Attempt to decrypt the data.
        DecryptData (EncryptedData), (Key), Data
        If StrPtr(Data) = 0 Then
            ' Decryption failed.
            Text = Null
        Else
            ' Return decrypted data as text.
            Text = CStr(Data)
        End If
    End If
    
    VDecryptBinary = Text
    
End Function

' Encrypt and Base64 encode a text directly to a Short Text or Long Text table field
' using a passed key.
' Accepts Null values for both data and key.
'
' If success, the encrypted text is returned.
' If no data, no key, or wrong key, Null is returned.
'
' Typical usage (update query):
'
'   Parameters
'       TableNameId Long,
'       [Text] LongText,
'       [Key] LongText;
'   Update
'       TableName
'   Set
'       TableName.[Date] = Date(),
'       TableName.Content = [Text],
'       TableName.ContentBase64 = VEncryptBase64([Text],[Key])
'       Where TableNameId = [Id];
'
' 2022-04-07. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VEncryptBase64( _
    ByVal Text As Variant, _
    Key As Variant) _
    As Variant

    Dim EncryptedText   As Variant
    
    If Nz(Key, "") = "" Or Nz(Text, "") = "" Then
        ' Either no key or no data.
        EncryptedText = Null
    Else
        ' Return encrypted data as Base64 encoded text.
        EncryptedText = Encrypt(Text, Key)
        If EncryptedText = "" Then
            ' Encryption failed.
            EncryptedText = Null
        End If
    End If
    
    VEncryptBase64 = EncryptedText
    
End Function

' Encrypt a text directly to a Binary or Long Binary (BLOB) table field
' using a passed key.
' Accepts Null values for both text and key.
'
' If success, the encrypted data is returned.
' If no data, no key, or wrong key, Null is returned.
'
' Typical usage:
'
' Typical usage (update query):
'
'   Parameters
'       TableNameId Long,
'       [Text] LongText,
'       [Key] LongText;
'   Update
'       TableName
'   Set
'       TableName.[Date] = Date(),
'       TableName.Content = [Text],
'       TableName.ContentBinary = VEncryptBinary([Text],[Key])
'       Where TableNameId = [Id];
'
' 2022-04-07. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VEncryptBinary( _
    ByVal Text As Variant, _
    Key As Variant) _
    As Variant

    Dim EncryptedData() As Byte
    Dim Data            As Variant
    
    If Nz(Key, "") = "" Or Nz(Text) = "" Then
        ' Either no key or no data.
        Data = Null
    ElseIf VarType(Key) <> vbString Or VarType(Text) <> vbString Then
        ' Key or Text is not text.
        Data = Null
    Else
        ' Arguments are valid. Attempt to encrypt the text.
        If EncryptData((Text), (Key), EncryptedData) = True Then
            ' Success.
            Data = EncryptedData
        Else
            ' Decryption failed.
            Data = Null
        End If
    End If
    
    VEncryptBinary = Data
    
End Function

