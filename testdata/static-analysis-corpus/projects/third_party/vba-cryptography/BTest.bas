Attribute VB_Name = "BTest"
Option Compare Database
Option Explicit
'
' BTest V1.1.2
' Various tests for or using the functions from BCrypt.
'
' (c) Gustav Brock, Cactus Data ApS, CPH
' https://github.com/GustavBrock/VBA.Cryptography
'
' Requires:
'   Module  BCrypt
'

' Verify a formula for calculating the encrypted byte lengths of small strings.
' Small strings are those that can be stored as "Short Text".
'
' 2022-02-18. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function TestEncryptedLength()

    Dim TextData()      As Byte
    Dim KeyData()       As Byte
    Dim EncryptedData() As Byte
    Dim Index           As Integer
    Dim LastLength      As Integer
    Dim ThisLength      As Integer
    Dim ByteLength      As Integer
    
    KeyData = "Any Key"
    
    ' List headers.
    Debug.Print "Lengths:"
    Debug.Print "Text", "Encrypted", "Calculated"
    
    ' Check for up to and including the maximum length of "Short Text", 255 characters.
    For Index = 1 To 260
        TextData = Space(Index)
        EncryptData TextData, KeyData, EncryptedData
        ' True length.
        ByteLength = 1 + UBound(EncryptedData)
        ' Calculated length.
        ThisLength = EncryptedByteLength(Index)
        
        ' List lengths.
        If LastLength < ByteLength Then
            Debug.Print Index, ByteLength, ThisLength
            LastLength = ByteLength
        End If
    Next
    
End Function

