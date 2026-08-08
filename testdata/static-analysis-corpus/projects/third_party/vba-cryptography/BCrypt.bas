Attribute VB_Name = "BCrypt"
Option Compare Binary
Option Explicit
'
' BCrypt V1.2.1
' Hashing, encrypting, and decrypting of text using the BCrypt API.
'
' (c) Gustav Brock, Cactus Data ApS, CPH
' https://github.com/GustavBrock/VBA.Cryptography
'
'
' Credit:
'
' Based on code at Stack Overflow published 2021-04-28 (with later edits):
'   https://stackoverflow.com/questions/67294035/basic-encrypting-of-a-text-file/67294779#comment122708972_67294779
' by Stack Overflow user Erik A:
'   https://stackoverflow.com/users/7296893/erik-a
'
' CNG API documentation:
'   Cryptography API: Next Generation
'   https://learn.microsoft.com/en-us/windows/win32/seccng/cng-portal


' API declarations.
'
Private Declare PtrSafe Function BCryptOpenAlgorithmProvider Lib "bcrypt.dll" ( _
    ByRef phAlgorithm As LongPtr, _
    ByVal pszAlgId As LongPtr, _
    ByVal pszImplementation As LongPtr, _
    ByVal dwFlags As Long) _
    As Long
    
Private Declare PtrSafe Function BCryptCloseAlgorithmProvider Lib "bcrypt.dll" ( _
    ByVal hAlgorithm As LongPtr, _
    ByVal dwFlags As Long) _
    As Long
    
Private Declare PtrSafe Function BCryptHash Lib "bcrypt.dll" ( _
    ByVal hAlgorithm As LongPtr, _
    ByVal pbSecret As LongPtr, _
    ByVal cbSecret As Long, _
    ByVal pbInput As LongPtr, _
    ByVal cbInput As Long, _
    ByVal pbOutput As LongPtr, _
    ByVal cbOutput As Long) _
    As Long

Private Declare PtrSafe Function BCryptGetProperty Lib "bcrypt.dll" ( _
    ByVal hObject As LongPtr, _
    ByVal pszProperty As LongPtr, _
    ByRef pbOutput As Any, _
    ByVal cbOutput As Long, _
    ByRef pcbResult As Long, _
    ByVal dfFlags As Long) _
    As Long
    
Private Declare PtrSafe Function BCryptSetProperty Lib "bcrypt.dll" ( _
    ByVal hObject As LongPtr, _
    ByVal pszProperty As LongPtr, _
    ByRef pbInput As Any, _
    ByVal cbInput As Long, _
    ByVal dfFlags As Long) _
    As Long
    
Private Declare PtrSafe Function BCryptCreateHash Lib "bcrypt.dll" ( _
    ByVal hAlgorithm As LongPtr, _
    ByRef phHash As LongPtr, _
    ByRef pbHashObject As Any, _
    ByVal cbHashObject As Long, _
    ByVal pbSecret As LongPtr, _
    ByVal cbSecret As Long, _
    ByVal dwFlags As Long) _
    As Long

Private Declare PtrSafe Function BCryptHashData Lib "bcrypt.dll" ( _
    ByVal hHash As LongPtr, _
    ByRef pbInput As Any, _
    ByVal cbInput As Long, _
    Optional ByVal dwFlags As Long = 0) _
    As Long

Private Declare PtrSafe Function BCryptFinishHash Lib "bcrypt.dll" ( _
    ByVal hHash As LongPtr, _
    ByRef pbOutput As Any, _
    ByVal cbOutput As Long, _
    ByVal dwFlags As Long) _
    As Long

Private Declare PtrSafe Function BCryptDestroyHash Lib "bcrypt.dll" ( _
    ByVal hHash As LongPtr) _
    As Long

Private Declare PtrSafe Function BCryptGenRandom Lib "bcrypt.dll" ( _
    ByVal hAlgorithm As LongPtr, _
    ByRef pbBuffer As Any, _
    ByVal cbBuffer As Long, _
    ByVal dwFlags As Long) _
    As Long

Private Declare PtrSafe Function BCryptGenerateSymmetricKey Lib "bcrypt.dll" ( _
    ByVal hAlgorithm As LongPtr, _
    ByRef hKey As LongPtr, _
    ByRef pbKeyObject As Any, _
    ByVal cbKeyObject As Long, _
    ByRef pbSecret As Any, _
    ByVal cbSecret As Long, _
    ByVal dwFlags As Long) _
    As Long

Private Declare PtrSafe Function BCryptEncrypt Lib "bcrypt.dll" ( _
    ByVal hKey As LongPtr, _
    ByRef pbInput As Any, _
    ByVal cbInput As Long, _
    ByRef pPaddingInfo As Any, _
    ByRef pbIV As Any, _
    ByVal cbIV As Long, _
    ByRef pbOutput As Any, _
    ByVal cbOutput As Long, _
    ByRef pcbResult As Long, _
    ByVal dwFlags As Long) _
    As Long

Private Declare PtrSafe Function BCryptDecrypt Lib "bcrypt.dll" ( _
    ByVal hKey As LongPtr, _
    ByRef pbInput As Any, _
    ByVal cbInput As Long, _
    ByRef pPaddingInfo As Any, _
    ByRef pbIV As Any, _
    ByVal cbIV As Long, _
    ByRef pbOutput As Any, _
    ByVal cbOutput As Long, _
    ByRef pcbResult As Long, _
    ByVal dwFlags As Long) _
    As Long

Private Declare PtrSafe Function BCryptDestroyKey Lib "bcrypt.dll" ( _
    ByVal hKey As LongPtr) _
    As Long

Private Declare PtrSafe Sub RtlMoveMemory Lib "Kernel32.dll" ( _
    ByRef Destination As Any, _
    ByRef Source As Any, _
    ByVal Length As LongPtr)
    
Private Declare PtrSafe Sub RtlZeroMemory Lib "Kernel32.dll" ( _
    ByRef Destination As Any, _
    ByVal Length As LongPtr)

Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpWideCharStr As LongPtr, _
    ByVal cchWideChar As Long, _
    ByVal lpMultiByteStr As LongPtr, _
    ByVal cbMultiByte As Long, _
    ByVal lpDefaultChar As LongPtr, _
    ByVal lpUsedDefaultChar As LongPtr) _
    As Long


' Constants for API.
'
Private Const BcryptAlgHandleHmacFlag   As Long = &H8
Private Const BcryptBlockPadding        As Long = 1
Private Const StatusNtSuccess           As Long = 0
Private Const CpUtf8                    As Long = 65001


' User defined types.
'
Private Type QuadSextet
    Sextet1 As Byte
    Sextet2 As Byte
    Sextet3 As Byte
    Sextet4 As Byte
End Type


' Enums.
'
' Allowed BCrypt hash algorithms.
Public Enum BcHashAlgorithm
    [_First] = 1
    bcMd2 = 1
    bcMd4 = 2
    bcMd5 = 3
    [_FirstHmac] = 4
    bcSha1 = 4
    bcSha256 = 5
    bcSha384 = 6
    bcSha512 = 7
    [_Last] = 7
    [_LastHmac] = 7
End Enum

' Allowed BCrypt random algorithms.
Public Enum BcRandomAlgorithm
    [_First] = 1
    bcRng = 1
    bcFips186DsaRng = 2
    [_Last] = 2
End Enum

' Utilised BCrypt encryption algorithms.
Public Enum BcEncryptionAlgorithm
    [_First] = 1
    bcAes = 1
    [_Last] = 1
End Enum


' Default application values.
'
' Default encryption algorithm.
Private Const DefaultEncryptionAlgorithm    As Long = BcEncryptionAlgorithm.bcAes
' Default hash algorithm.
Private Const DefaultBcHashAlgorithm        As Long = BcHashAlgorithm.bcSha256
' Default random algorithm.
Private Const DefaultBcRandomAlgorithm      As Long = BcRandomAlgorithm.bcRng


' Application contants.
'
' Maximum size (byte count) of a Binary field of an Access table.
Private Const MaximumBinaryFieldSize        As Integer = 510
' Maximum size (character count) of a Short Text field of an Access table.
Private Const MaximumTextFieldSize          As Integer = 255

' Convert and return a Base64 encoded string as a byte array.
' Generic function.
'
' To be called from function Decrypt.
'
' 2021-10-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function Base64Bytes( _
    ByVal Text64 As String) _
    As Byte()
    
    Dim Index           As Long
    Dim Tail            As Long
    Dim Sextets         As QuadSextet
    Dim Data()          As Byte
    Dim DataLength      As Long
    
    If Right(Text64, 2) = "==" Then
        Tail = 2
    ElseIf Right(Text64, 1) = "=" Then
        Tail = 1
    End If
    DataLength = Len(Text64) * 3 \ 4 - Tail

    ReDim Data(0 To DataLength - 1)
    For Index = LBound(Data) To UBound(Data)
        Select Case Index Mod 3
            Case 0
                Sextets = Base64QuadSextet(Mid(Text64, (Index \ 3) * 4 + 1, 4))
                Data(Index) = Sextets.Sextet1 * 2 ^ 2 + (Sextets.Sextet2 \ 2 ^ 4)
            Case 1
                Data(Index) = (Sextets.Sextet2 * 2 ^ 4 And 255) + Sextets.Sextet3 \ 2 ^ 2
            Case 2
                Data(Index) = (Sextets.Sextet3 * 2 ^ 6 And 255) + Sextets.Sextet4
        End Select
    Next
    
    Base64Bytes = Data

End Function

' Convert and return a Base64 encoded string as a QuadSextet.
' Generic function.
'
' To be called from function Base64Bytes.
'
' 2021-10-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function Base64QuadSextet( _
    ByVal Text64 As String) _
    As QuadSextet
    
    Const Base64Table   As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="
    
    Dim Sextets         As QuadSextet
    
    Sextets.Sextet1 = InStr(Base64Table, Mid(Text64, 1, 1)) - 1
    Sextets.Sextet2 = InStr(Base64Table, Mid(Text64, 2, 1)) - 1
    Sextets.Sextet3 = InStr(Base64Table, Mid(Text64, 3, 1)) - 1
    Sextets.Sextet4 = InStr(Base64Table, Mid(Text64, 4, 1)) - 1
    
    Base64QuadSextet = Sextets
    
End Function

' Return the literal encryption algorithm name determined by
' the passed value of enum BcEncryptionAlgorithm.
'
' To be called from functions CngEncrypt and CngDecrypt.
'
' 2021-10-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function BcryptEncryptionAlgorithm( _
    ByVal BcryptEncryptionAlgorithmId As BcEncryptionAlgorithm) _
    As String
    
    Dim EncryptionAlgorithmName As String
    
    ' Note: EncryptionAlgorithmName must be in UPPERCASE.
    Select Case BcryptEncryptionAlgorithmId
        Case BcEncryptionAlgorithm.bcAes
            EncryptionAlgorithmName = "AES"
    End Select
    
    BcryptEncryptionAlgorithm = EncryptionAlgorithmName

End Function

' Return the literal hash algorithm name determined by
' the passed value of enum BcHashAlgorithm.
'
' To be called from functions CngHash and IsBcryptHashAlgorithm.
'
' 2021-10-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function BcryptHashAlgorithm( _
    ByVal BcryptHashAlgorithmId As BcHashAlgorithm) _
    As String
    
    Dim HashAlgorithmName       As String
    
    ' Note: HashAlgorithmName must be in UPPERCASE.
    Select Case BcryptHashAlgorithmId
        Case BcHashAlgorithm.bcMd2
            HashAlgorithmName = "MD2"
        Case BcHashAlgorithm.bcMd4
            HashAlgorithmName = "MD4"
        Case BcHashAlgorithm.bcMd5
            HashAlgorithmName = "MD5"
        Case BcHashAlgorithm.bcSha1
            HashAlgorithmName = "SHA1"
        Case BcHashAlgorithm.bcSha256
            HashAlgorithmName = "SHA256"
        Case BcHashAlgorithm.bcSha384
            HashAlgorithmName = "SHA384"
        Case BcHashAlgorithm.bcSha512
            HashAlgorithmName = "SHA512"
    End Select
    
    BcryptHashAlgorithm = HashAlgorithmName

End Function

' Return the literal random algorithm name determined by
' the passed value of enum BcRandomAlgorithm.
'
' To be called from functions CngRandom and IsBcryptRandomAlgorithm.
'
' 2021-10-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function BcryptRandomAlgorithm( _
    ByVal BcryptRandomAlgorithmId As BcRandomAlgorithm) _
    As String
    
    Dim RandomAlgorithmName     As String
    
    ' Note: RandomAlgorithmName must be in UPPERCASE.
    Select Case BcryptRandomAlgorithmId
        Case BcRandomAlgorithm.bcRng
            RandomAlgorithmName = "RNG"
        Case BcRandomAlgorithm.bcFips186DsaRng
            RandomAlgorithmName = "FIPS186DSARNG"
    End Select
    
    BcryptRandomAlgorithm = RandomAlgorithmName

End Function

' Convert and return a byte array as a Base64 encoded string.
' Generic function.
'
' To be called from functions Encrypt, Hash, and TextBase64.
'
' 2022-02-20. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function ByteBase64( _
    ByRef Data() As Byte) _
    As String

    Const Base64Table   As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
    
    Dim Sextets         As QuadSextet
    Dim Index           As Long
    Dim Text            As String
    Dim Number          As Long
    
    If StrPtr(Data) = 0 Then
        ' No data. Nothing to do.
    Else
        Number = (1 + (UBound(Data) \ 3)) * 4
        Text = String(Number, "=")
        
        For Index = 0 To (UBound(Data) - 2) \ 3
            Select Case (UBound(Data) - 2)
                Case -2
                    Sextets = BytesQuadSextet(Data(Index * 3))
                Case -1
                    Sextets = BytesQuadSextet(Data(Index * 3), Data(Index * 3 + 1))
                Case Else
                    Sextets = BytesQuadSextet(Data(Index * 3), Data(Index * 3 + 1), Data(Index * 3 + 2))
            End Select
            Mid(Text, (Index * 4) + 1, 1) = Mid(Base64Table, Sextets.Sextet1 + 1, 1)
            Mid(Text, (Index * 4) + 2, 1) = Mid(Base64Table, Sextets.Sextet2 + 1, 1)
            If UBound(Data) >= 1 Then
                Mid(Text, (Index * 4) + 3, 1) = Mid(Base64Table, Sextets.Sextet3 + 1, 1)
            End If
            If UBound(Data) >= 2 Then
                Mid(Text, (Index * 4) + 4, 1) = Mid(Base64Table, Sextets.Sextet4 + 1, 1)
            End If
        Next
        
        If UBound(Data) >= 2 Then
            Select Case (1 + UBound(Data)) Mod 3
                Case Is = 2
                    ' Leave one trailing equal sign.
                    Sextets = BytesQuadSextet(Data(Index * 3), Data(Index * 3 + 1))
                    Mid(Text, (Index * 4) + 1, 1) = Mid(Base64Table, Sextets.Sextet1 + 1, 1)
                    Mid(Text, (Index * 4) + 2, 1) = Mid(Base64Table, Sextets.Sextet2 + 1, 1)
                    Mid(Text, (Index * 4) + 3, 1) = Mid(Base64Table, Sextets.Sextet3 + 1, 1)
                Case Is = 1
                    ' Leave two trailing equal signs.
                    Sextets = BytesQuadSextet(Data(Index * 3))
                    Mid(Text, (Index * 4) + 1, 1) = Mid(Base64Table, Sextets.Sextet1 + 1, 1)
                    Mid(Text, (Index * 4) + 2, 1) = Mid(Base64Table, Sextets.Sextet2 + 1, 1)
                Case Is = 0
                    ' Leave no trailing equal sign.
            End Select
        End If
    End If
    
    ByteBase64 = Text
    
End Function

' Convert and return one, two, or three byte arrays as a QuadSextet.
' Generic function.
'
' To be called from function ByteBase64.
'
' 2021-10-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function BytesQuadSextet( _
    ByVal Byte1 As Byte, _
    Optional ByVal Byte2 As Byte, _
    Optional ByVal Byte3 As Byte) _
    As QuadSextet
    
    Dim Sextets         As QuadSextet
    
    Sextets.Sextet1 = Byte1 \ 4
    Sextets.Sextet2 = (((Byte1 * 2 ^ 6) And 255) \ 4) + Byte2 \ (2 ^ 4)
    Sextets.Sextet3 = (((Byte2 * 2 ^ 4) And 255) \ 4) + Byte3 \ (2 ^ 6)
    Sextets.Sextet4 = (((Byte3 * 2 ^ 2) And 255) \ 4)
    
    BytesQuadSextet = Sextets
    
End Function

' Decrypt data at pointer Data using AES encryption, IVectorInput and Secret.
'
' To be called from function DecryptData.
'
'   Arguments:
'       Data:               Memory pointer to data.
'       DataLength:         Length of byte array to decrypt.
'       IVectorInput:       Memory pointer to IVector.
'       IVectorInputLength: Length of byte array for the IVector.
'       Secret:             Memory pointer to 128-bits secret.
'       SecretLength:       Length of byte array for the secret for AES encrypting.
'   Output:
'       Byte array containing decrypted data.
'
' Original code by Erik A, 2019.
' 2026-01-08. Gustav Brock, Cactus Data ApS, CPH.
'
Private Function CngDecrypt( _
    ByVal Data As LongPtr, _
    ByVal DataLength As Long, _
    ByVal IVectorInput As LongPtr, _
    ByVal IVectorInputLength As Long, _
    ByVal Secret As LongPtr, _
    ByVal SecretLength As Long) _
    As Byte()
    
    Const Implementation    As LongPtr = 0
    Const Flags             As Long = 0
    Const Result            As Long = 0
    
    Dim Algorithm           As LongPtr
    Dim AlgorithmId         As String
    Dim Property            As String
    Dim KeyObject()         As Byte
    Dim KeyObjectLength     As Long
    Dim IVector()           As Byte
    Dim IVectorLength       As Long
    Dim Value               As String
    Dim Key                 As LongPtr
    Dim PlainText()         As Byte
    Dim PlainTextLength     As Long
    Dim ResultLength        As Long     ' Not used.
    Dim Status              As Long

    ' Open algorithm provider.
    AlgorithmId = BcryptEncryptionAlgorithm(DefaultEncryptionAlgorithm) & vbNullChar
    Status = BCryptOpenAlgorithmProvider(Algorithm, StrPtr(AlgorithmId), Implementation, Flags)

    If Status = StatusNtSuccess Then
        ' Allocate memory to hold KeyObject.
        Property = "ObjectLength" & vbNullChar
        Status = BCryptGetProperty(Algorithm, StrPtr(Property), KeyObjectLength, LenB(KeyObjectLength), Result, Flags)
    End If
    
    If Status = StatusNtSuccess Then
        ReDim KeyObject(0 To KeyObjectLength - 1)
        ' Calculate the block length for the IVector.
        Property = "BlockLength" & vbNullChar
        Status = BCryptGetProperty(Algorithm, StrPtr(Property), IVectorLength, LenB(IVectorLength), Result, Flags)
        If IVectorLength > IVectorInputLength Then
            Debug.Print "IVector lengths:", IVectorLength, IVectorInputLength
            Status = Not StatusNtSuccess
        End If
    End If
    
    If Status = StatusNtSuccess Then
        ' Resize and copy the initialization vector.
        ReDim IVector(0 To IVectorLength - 1)
        RtlMoveMemory IVector(0), ByVal IVectorInput, IVectorLength

        ' Set chaining mode.
        Property = "ChainingMode" & vbNullChar
        Value = "ChainingModeCBC" & vbNullChar
        Status = BCryptSetProperty(Algorithm, StrPtr(Property), ByVal StrPtr(Value), LenB(Value), Flags)
    End If

    If Status = StatusNtSuccess Then
        ' Create KeyObject using secret.
        If BCryptGenerateSymmetricKey(Algorithm, Key, KeyObject(0), KeyObjectLength, ByVal Secret, SecretLength, Flags) = StatusNtSuccess Then
            ' Calculate output buffer size and allocate output buffer.
            If BCryptDecrypt(Key, ByVal Data, DataLength, ByVal 0, IVector(0), IVectorLength, ByVal 0, 0, PlainTextLength, BcryptBlockPadding) = StatusNtSuccess Then
                ReDim PlainText(0 To PlainTextLength - 1)
                ' Decrypt the data.
                If BCryptDecrypt(Key, ByVal Data, DataLength, ByVal 0, IVector(0), IVectorLength, PlainText(0), PlainTextLength, ResultLength, BcryptBlockPadding) = StatusNtSuccess Then
                    CngDecrypt = PlainText
                End If
            End If
        End If
    End If
    
    ' Clean up.
    If Key <> 0 Then
        BCryptDestroyKey Key
    End If
    If Algorithm <> 0 Then
        BCryptCloseAlgorithmProvider Algorithm, Flags
    End If

    If KeyObjectLength > 0 Then
        RtlZeroMemory KeyObject(0), KeyObjectLength
    End If
    If IVectorLength > 0 Then
        RtlZeroMemory IVector(0), IVectorLength
    End If

End Function

' Encrypt data at pointer Data using AES encryption, IVectorInput and Secret.
'
' To be called from functions CngEncryptW and EncryptData.
'
'   Arguments:
'       Data:               Memory pointer to data.
'       DataLength:         Length of byte array to encrypt.
'       IVectorInput:       Memory pointer to IVector.
'       IVectorInputLength: Length of byte array for the IVector.
'       Secret:             Memory pointer to 128-bits secret.
'       SecretLength:       Length of byte array for the secret for AES encrypting.
'   Output:
'       Byte array containing encrypted data.
'
' Original code by Erik A, 2019.
' 2026-01-08. Gustav Brock, Cactus Data ApS, CPH.
'
Private Function CngEncrypt( _
    ByVal Data As LongPtr, _
    ByVal DataLength As Long, _
    ByVal IVectorInput As LongPtr, _
    ByVal IVectorInputLength As Long, _
    ByVal Secret As LongPtr, _
    ByVal SecretLength As Long) _
    As Byte()
    
    Const Implementation    As LongPtr = 0
    Const Flags             As Long = 0
    Const Result            As Long = 0
    
    Dim Algorithm           As LongPtr
    Dim AlgorithmId         As String
    Dim Property            As String
    Dim KeyObject()         As Byte
    Dim KeyObjectLength     As Long
    Dim IVector()           As Byte
    Dim IVectorLength       As Long
    Dim Value               As String
    Dim Key                 As LongPtr
    Dim CipherText()        As Byte
    Dim CipherTextLength    As Long
    Dim ResultLength        As Long     ' Not used.
    Dim Status              As Long

    ' Open algorithm provider.
    AlgorithmId = BcryptEncryptionAlgorithm(DefaultEncryptionAlgorithm) & vbNullChar
    Status = BCryptOpenAlgorithmProvider(Algorithm, StrPtr(AlgorithmId), Implementation, Flags)

    If Status = StatusNtSuccess Then
        ' Allocate memory to hold KeyObject.
        Property = "ObjectLength" & vbNullChar
        Status = BCryptGetProperty(Algorithm, StrPtr(Property), KeyObjectLength, LenB(KeyObjectLength), Result, Flags)
    End If
    
    If Status = StatusNtSuccess Then
        ReDim KeyObject(0 To KeyObjectLength - 1)
        ' Check that block length = 128 bits.
        Property = "BlockLength" & vbNullChar
        Status = BCryptGetProperty(Algorithm, StrPtr(Property), IVectorLength, LenB(IVectorLength), Result, Flags)
        If IVectorLength > IVectorInputLength Then
            Debug.Print "IVector lengths:", IVectorLength, IVectorInputLength
            Status = Not StatusNtSuccess
        End If
    End If
    
    If Status = StatusNtSuccess Then
        ' Resize and copy the initialization vector.
        ReDim IVector(0 To IVectorLength - 1)
        RtlMoveMemory IVector(0), ByVal IVectorInput, IVectorLength

        ' Set chaining mode.
        Property = "ChainingMode" & vbNullChar
        Value = "ChainingModeCBC" & vbNullChar
        Status = BCryptSetProperty(Algorithm, StrPtr(Property), ByVal StrPtr(Value), LenB(Value), Flags)
    End If

    If Status = StatusNtSuccess Then
        ' Create KeyObject using secret.
        If BCryptGenerateSymmetricKey(Algorithm, Key, KeyObject(0), KeyObjectLength, ByVal Secret, SecretLength, Flags) = StatusNtSuccess Then
            ' Calculate output buffer size and allocate output buffer.
            If BCryptEncrypt(Key, ByVal Data, DataLength, ByVal 0, IVector(0), IVectorLength, ByVal 0, 0, CipherTextLength, BcryptBlockPadding) = StatusNtSuccess Then
                ReDim CipherText(0 To CipherTextLength - 1)
                ' Encrypt the data.
                If BCryptEncrypt(Key, ByVal Data, DataLength, ByVal 0, IVector(0), IVectorLength, CipherText(0), CipherTextLength, ResultLength, BcryptBlockPadding) = StatusNtSuccess Then
                    ' Output the encrypted data.
                    CngEncrypt = CipherText
                End If
            End If
        End If
    End If
    
    ' Clean up.
    If Key <> 0 Then
        BCryptDestroyKey Key
    End If
    If Algorithm <> 0 Then
        BCryptCloseAlgorithmProvider Algorithm, Flags
    End If

    If KeyObjectLength > 0 Then
        RtlZeroMemory KeyObject(0), KeyObjectLength
    End If
    If IVectorLength > 0 Then
        RtlZeroMemory IVector(0), IVectorLength
    End If

End Function

' Create a hash of the data at pointer Data by using the Next Generation Cryptography API
' and returns the hash as a byte array.
'
' To be called from functions DecryptData and HashData.
'
' Allowed algorithms (NB: Hash algorithms only; check OS support):
'   https://docs.microsoft.com/en-us/windows/desktop/SecCNG/cng-algorithm-identifiers
'
' Loosely based on the method described at:
'   https://docs.microsoft.com/en-us/windows/desktop/SecCNG/creating-a-hash-with-cng
'
' Original code by Erik A, 2019.
' 2026-01-08. Gustav Brock, Cactus Data ApS, CPH.
'
Private Function CngHash( _
    ByVal Data As LongPtr, _
    ByVal DataLength As Long, _
    Optional ByVal BcryptHashAlgorithmId As BcHashAlgorithm = BcHashAlgorithm.bcSha256) _
    As Byte()
    
    Const Implementation    As LongPtr = 0
    Const Flags             As Long = 0
    Const Secret            As Long = 0
    Const SecretLength      As Long = 0
    Const Result            As Long = 0
    
    Dim Algorithm           As LongPtr
    Dim AlgorithmId         As String
    Dim Property            As String
    Dim HashObject()        As Byte
    Dim HashObjectLength    As Long
    Dim HashDigestLength    As Long
    Dim Hash                As LongPtr
    Dim Values()            As Byte
    Dim Status              As Long
    
    ' Validate requested hash algorithm.
    If Not IsBcryptHashAlgorithmId(BcryptHashAlgorithmId) Then
        ' Invalid hash algorithm id. Use default hash algorithm.
        BcryptHashAlgorithmId = DefaultBcHashAlgorithm
    End If
    
    ' Open algorithm provider.
    AlgorithmId = BcryptHashAlgorithm(BcryptHashAlgorithmId) & vbNullChar
    Status = BCryptOpenAlgorithmProvider(Algorithm, StrPtr(AlgorithmId), Implementation, Flags)
    
    If Status = StatusNtSuccess Then
        ' Determine hash object size and allocate memory.
        Property = "ObjectLength" & vbNullChar
        Status = BCryptGetProperty(Algorithm, StrPtr(Property), HashObjectLength, LenB(HashObjectLength), Result, Flags)
    End If
    
    If Status = StatusNtSuccess Then
        ReDim HashObject(0 To HashObjectLength - 1)
        ' Determine hash digest size and allocate memory.
        Property = "HashDigestLength" & vbNullChar
        Status = BCryptGetProperty(Algorithm, StrPtr(Property), HashDigestLength, LenB(HashDigestLength), Result, Flags)
    End If
    
    If Status = StatusNtSuccess Then
        ReDim Values(0 To HashDigestLength - 1)
        ' Create the hash object.
        If BCryptCreateHash(Algorithm, Hash, HashObject(0), HashObjectLength, Secret, SecretLength, Flags) = StatusNtSuccess Then
            ' Hash the data.
            If BCryptHashData(Hash, ByVal Data, DataLength) = StatusNtSuccess Then
                ' Get the hash.
                If BCryptFinishHash(Hash, Values(0), HashDigestLength, Flags) = StatusNtSuccess Then
                    ' Return the hash.
                    CngHash = Values
                End If
            End If
        End If
    End If
    
    ' Clean up.
    If Algorithm <> 0 Then
        BCryptCloseAlgorithmProvider Algorithm, Flags
    End If
    If Hash <> 0 Then
        BCryptDestroyHash Hash
    End If

End Function

' Create a HMAC (Hash-Based Message Authentication Code) hash of the data at pointer Data by
' using the Next Generation Cryptography API.
' Returns the hash as a byte array.
'
' To be called from function HmacData.
'
' Allowed algorithms (NB: HMAC Hash algorithms only; check OS support):
'   https://docs.microsoft.com/en-us/windows/desktop/SecCNG/cng-algorithm-identifiers
'
' Loosely based on the method described at:
'   https://docs.microsoft.com/en-us/windows/desktop/SecCNG/creating-a-hash-with-cng
'
' 2026-01-08. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function CngHmac( _
    ByVal TextData As LongPtr, _
    ByVal TextDataLength As Long, _
    ByVal KeyData As LongPtr, _
    ByVal KeyDataLength As Long, _
    Optional ByVal BcryptHashAlgorithmId As BcHashAlgorithm = BcHashAlgorithm.bcSha256) _
    As Byte()
    
    Const Implementation    As LongPtr = 0
    Const Flags             As Long = 0
    
    Dim Data()              As Byte
    
    Dim Algorithm           As LongPtr
    Dim AlgorithmId         As String
    Dim DataLength          As Long
    Dim Status              As Long

    ' Validate requested hash algorithm.
    If Not IsBcryptHashHmacAlgorithmId(BcryptHashAlgorithmId) Then
        ' Invalid hash algorithm for HMAC. Use default hash algorithm.
        BcryptHashAlgorithmId = DefaultBcHashAlgorithm
    End If

    ' Open algorithmprovider as HMAC.
    AlgorithmId = BcryptHashAlgorithm(BcryptHashAlgorithmId)
    Status = BCryptOpenAlgorithmProvider(Algorithm, StrPtr(AlgorithmId), Implementation, BcryptAlgHandleHmacFlag)
    
    If Status = StatusNtSuccess Then
        ' Determine digest size based on the algorithm id.
        DataLength = HashByteLength(BcryptHashAlgorithmId)
        ReDim Data(0 To DataLength - 1)
        ' Create the HMAC hash value.
        If BCryptHash(Algorithm, KeyData, KeyDataLength, TextData, TextDataLength, VarPtr(Data(0)), DataLength) = StatusNtSuccess Then
            ' Return the HMAC hash.
            CngHmac = Data
        End If
    End If

    ' Clean up.
    If Algorithm <> 0 Then
        BCryptCloseAlgorithmProvider Algorithm, Flags
    End If

End Function

' Fill data at pointer Data with random bytes.
'
' To be called from function RandomData.
'
' Allowed algorithms (NB: Random algorithms only; check OS support):
'   https://docs.microsoft.com/en-us/windows/desktop/SecCNG/cng-algorithm-identifiers
'
' Original code by Erik A, 2019.
' 2021-10-24. Gustav Brock, Cactus Data ApS, CPH.
'
Private Sub CngRandom( _
    ByRef Data As LongPtr, _
    ByRef BufferSize As Long, _
    Optional ByVal BcryptRandomAlgorithmId As BcRandomAlgorithm = BcRandomAlgorithm.bcRng)

    Const Implementation    As LongPtr = 0
    Const Flags             As Long = 0
    
    Dim Algorithm           As LongPtr
    Dim AlgorithmId         As String
    
    ' Validate requested randomise algorithm.
    If Not IsBcryptRandomAlgorithmId(BcryptRandomAlgorithmId) Then
        ' Invalid randomise algorithm id. Use default randomise algorithm.
        BcryptRandomAlgorithmId = DefaultBcRandomAlgorithm
    End If
    
    ' Open crypto provider.
    AlgorithmId = BcryptRandomAlgorithm(BcryptRandomAlgorithmId) & vbNullChar
    BCryptOpenAlgorithmProvider Algorithm, StrPtr(AlgorithmId), Implementation, Flags
    
    ' Fill byte array with random data and return size of buffer.
    BCryptGenRandom Algorithm, ByVal Data, BufferSize, Flags
    
    ' Clean up.
    BCryptCloseAlgorithmProvider Algorithm, Flags
    
End Sub

' Decrypt a Base64 encoded string encrypted using AES and a key.
' Return the decrypted and decoded text as a plain string.
'
' Example:
'   EncryptedText = "6uLffExuQmAi/oI3AzCLZTRZfv1XL6kl01z4hJ5y1MWXHgFACj3XhvboF/rNU89znrX1d5btmCbRK9dAjjjlKxTDJMImQr3YGiscMDvn/YtjKmc8nFuR65IU9vEn4a0Rca72k55cZXjKzOGMpbZ/6A=="
'   Key = "Have a Cigar"
'   Text = Decrypt(EncryptedText, Key)
'   Text -> Careful with that axe, Eugene!
'
' Original code by Erik A, 2019.
' 2021-10-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function Decrypt( _
    ByVal EncryptedText As String, _
    ByVal Key As String) _
    As String

    Dim EncryptedData()     As Byte
    Dim TextData()          As Byte
    
    If EncryptedText = "" Or Key = "" Then
        ' Nothing to do.
    Else
        ' Convert the Base64 encoded string to a byte array.
        EncryptedData = Base64Bytes(EncryptedText)
        If DecryptData(EncryptedData, (Key), TextData) = True Then
            ' Success.
        Else
            ' Invalid EncryptedData or wrong key.
        End If
    End If
    
    Decrypt = TextData

End Function

' Decrypt an AES encrypted byte array using a key passed as another byte array.
' Return by reference the decrypted data as a byte array.
' Return True if success.
'
' To be called from function Decrypt.
'
' Original code by Erik A, 2019.
' 2026-01-08. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DecryptData( _
    ByRef EncryptedData() As Byte, _
    ByRef KeyData() As Byte, _
    ByRef DecryptedData() As Byte) _
    As Boolean
    
    Const SizeLength        As Long = 4
    Const IVectorLength     As Long = 16
    Const SecretLength      As Long = 16

    Dim DataLength          As Long
    Dim DataHash()          As Byte
    Dim KeyHash()           As Byte
    Dim IVector             As LongPtr
    Dim Data()              As Byte
    Dim Index               As Long
    Dim HashError           As Boolean
    Dim Result              As Boolean
    
    If StrPtr(EncryptedData) = 0 Or StrPtr(KeyData) = 0 Then
        ' Either no data or no key. Nothing to do.
    ElseIf EncryptedByteLength(UBound(EncryptedData) - LBound(EncryptedData) + 1) = 0 Then
        ' Data length is below minimum.
    Else
        ' Get the SHA1 hash of the key. Could be modified to use SHA256.
        KeyHash = HashData(KeyData, bcSha1)
        ReDim Preserve KeyHash(0 To SecretLength - 1)
        ' Get the pointer to IVector. The last 16 bytes is IVector.
        IVector = VarPtr(EncryptedData(UBound(EncryptedData) - LBound(EncryptedData) + 1 - IVectorLength))
        ' Decrypt the data.
        Data = CngDecrypt( _
            VarPtr(EncryptedData(0)), UBound(EncryptedData) - LBound(EncryptedData) + 1 - IVectorLength, _
            IVector, IVectorLength, _
            VarPtr(KeyHash(0)), SecretLength)
        
        ' Check that CngDecrypt did return some data.
        If StrPtr(Data) = 0 Then
            ' No data. Most likely, a wrong key was passed.
        ElseIf UBound(Data) <= SizeLength Then
            ' No useful data returned. A minimum of five bytes is expected.
        Else
            ' Get the data length.
            RtlMoveMemory DataLength, Data(0), SizeLength
            
            ' Check if length is valid, with invalid key length = random data.
            If DataLength > (UBound(Data) - LBound(Data) + 1 - SizeLength) Or DataLength < 0 Then
                ' No useful data.
                ' Leave DecryptedData empty, and return False.
            Else
                ' Get the SHA1 hash of the decrypted data. Could be modified to use SHA256.
                DataHash = CngHash(VarPtr(Data(SizeLength)), DataLength, bcSha1)
                ' Verify the hash.
                For Index = LBound(DataHash) To UBound(DataHash)
                    If DataHash(Index) <> Data(SizeLength + DataLength + Index) Then
                        ' Stored hash not equal to hash with decrypted data, key incorrect, or encrypted data has been tampered with.
                        ' Leave DecryptedData empty, and return False.
                        HashError = True
                        Exit For
                    End If
                Next
                If Not HashError Then
                    ' Initialize output array.
                    ReDim DecryptedData(0 To DataLength - 1)
                    ' Copy data to output array.
                    RtlMoveMemory DecryptedData(0), Data(SizeLength), DataLength
                    Result = True
                End If
            End If
        End If
    End If
    
    DecryptData = Result
    
End Function

' Return for an avaliable length of an encrypted byte array
' the possible maximum length of a byte array to be encrypted.
'
' To be called from function FitByteField.
'
' Examples:
'   Available encrypted     Length plain maximum
'   0                         0
'  47                         0
'  48                         1
'  63                         1
'  64                         4
'  79                         4
'  80                        12
'  95                        12
'  96                        20
' 111                        20
' 112                        28
' 127                        28
' 128                        36
' 143                        36
' 144                        44
' 159                        44
' 160                        52
' 175                        52
' 176                        60
' 191                        60
' 192                        68
' 207                        68
' 208                        76
' 223                        76
' 224                        84
' 239                        84
' 240                        92
' 255                        92
' 256                       100
' 480                       212
' 495                       212
' 496                       220
' Largest option for a Binary field of Access
' holding the encrypted byte array:
' 510                       220
' 511                       220
' 512                       228
'
' 2021-10-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DecryptedByteLength( _
    ByVal EncryptedByteLength As Long) _
    As Long
    
    Dim Length      As Long
    
    Select Case EncryptedByteLength
        Case Is < 48
            Length = 0
        Case Is < 64
            Length = 1
        Case Else
            Length = Int(EncryptedByteLength / 16) * 8 - 28
    End Select
    
    DecryptedByteLength = Length

End Function

' Return for an available length of an encrypted and Base64 encoded string
' the possible maximum length of a string to be encrypted.
'
' To be called from function FitTextField.
'
' Examples:
'   Available encrypted     Length plain maximum
'     0                       0
'    63                       0
'    64                       3
'    87                       3
'    88                      11
'   107                      11
'   108                      19
'   127                      19
'   128                      27
'   151                      27
'   152                      35
'   171                      35
'   172                      43
'   191                      43
'   192                      51
'   215                      51
'   216                      59
'   235                      59
' Largest option for a Short Text field of Access
' holding the decrypted string:
'   236                      67
'   255                      67
'   256                      75
'   747                     251
' Largest option for a Short Text field of Access
' holding the encrypted string:
'   748                     259
'
' 2021-10-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DecryptedTextLength( _
    ByVal EncryptedTextLength As Long) _
    As Long
    
    Dim Length      As Long
    Dim Length64    As Long
    Dim Delta20     As Long
    Dim Delta       As Long
    
    Select Case EncryptedTextLength
        Case Is < 64
            Length = 0
        Case Is < 64 + 24
            Length = 3
        Case Else
            Length64 = -Int(-EncryptedTextLength / 64)
            Delta20 = -Int(-((Length64 * 64 - EncryptedTextLength) / 20))
            Delta = (Delta20 \ 4) * 3 + Delta20 Mod 4
            Length = (Length64 - 1) * 24 - (Delta * 8) - 4 + 7
    End Select
    
    DecryptedTextLength = Length

End Function

' Encrypt a string using AES and a key.
' Return the encrypted text as a Base64 encoded string.
'
' Example:
'   Text = "Careful with that axe, Eugene!"
'   Key = "Have a Cigar"
'   EncryptedText = Encrypt(Text, Key)
'   EncryptedText -> 6uLffExuQmAi/oI3AzCLZTRZfv1XL6kl01z4hJ5y1MWXHgFACj3XhvboF/rNU89znrX1d5btmCbRK9dAjjjlKxTDJMImQr3YGiscMDvn/YtjKmc8nFuR65IU9vEn4a0Rca72k55cZXjKzOGMpbZ/6A==
'
' Note: Length of the encrypted string can be predetermined by the function EncryptedTextLength:
'   ' Use Text from example above.
'   Length = EncryptedTextLength(Len(Text))
'   Length -> 152
'
' Original code by Erik A, 2019.
' 2021-10-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function Encrypt( _
    ByVal Text As String, _
    ByVal Key As String) _
    As String

    Dim EncryptedData()     As Byte
    Dim EncryptedText       As String
    
    If Text = "" Or Key = "" Then
        ' Nothing to do.
    Else
        If EncryptData((Text), (Key), EncryptedData) = True Then
            ' Success.
            ' Convert the byte array to a Base64 encoded string.
            EncryptedText = ByteBase64(EncryptedData)
        Else
            ' Missing Text or Key.
        End If
    End If
    
    Encrypt = EncryptedText

End Function

' Encrypt a byte array using AES encryption and a KeyData passed as another byte array.
' Return by reference the encrypted data as a byte array.
' Return True if success.
'
' To be called from function Encrypt.
'
' NOTE:
'   Even when passed the same arguments (TextData and KeyData), the returned and
'   encrypted data will be unique for every call.
'
' Original code by Erik A, 2019.
' 2026-01-08. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function EncryptData( _
    ByRef TextData() As Byte, _
    ByRef KeyData() As Byte, _
    ByRef EncryptedData() As Byte) _
    As Boolean
    
    Const SizeLength        As Long = 4
    Const IVectorLength     As Long = 16
    Const SecretLength      As Long = 16
    
    Dim KeyHash()           As Byte
    Dim InputHash()         As Byte
    Dim Data()              As Byte
    Dim DataLength          As Long
    Dim IVector()           As Byte
    Dim InputHashLength     As Long
    Dim Result              As Boolean
    
    ' Get SHA1 hash of the data and of the KeyData. Could be modified to use SHA256.
    InputHash = HashData(TextData, bcSha1)
    If StrPtr(InputHash) = 0 Then
        ' Empty data. Exit with success.
        Result = True
    Else
        InputHashLength = UBound(InputHash) + 1
        ' Get the SHA1 hash of the key. Could be modified to use SHA256.
        KeyHash = HashData(KeyData, bcSha1)
        ReDim Preserve KeyHash(0 To SecretLength - 1)
    End If
    
    If StrPtr(InputHash) = 0 Or StrPtr(KeyHash) = 0 Then
        ' Either no data or no KeyData. Nothing to do.
    Else
        DataLength = UBound(TextData) - LBound(TextData) + 1
        
        ' Data size is: Long (4 bytes) + DataLength + SHA1 (20 bytes)
        ReDim Data(0 To SizeLength + DataLength + InputHashLength - 1)
        ' Append length (in bytes) to start of array.
        RtlMoveMemory Data(0), DataLength, SizeLength
        ' Append data.
        RtlMoveMemory Data(SizeLength), TextData(LBound(TextData)), DataLength
        ' Append hash of the data.
        RtlMoveMemory Data(SizeLength + DataLength), InputHash(0), InputHashLength
        
        ' Generate IVector.
        IVector = RandomData(IVectorLength)
        ' Encrypt data.
        EncryptedData = CngEncrypt( _
            VarPtr(Data(0)), SizeLength + DataLength + InputHashLength, _
            VarPtr(IVector(0)), IVectorLength, _
            VarPtr(KeyHash(0)), SecretLength)
        ' Deallocate copy made to encrypt.
        Erase Data
        ' Extend encrypted data to append IVector.
        ReDim Preserve EncryptedData(LBound(EncryptedData) To UBound(EncryptedData) + IVectorLength)
        ' Append IVector.
        RtlMoveMemory EncryptedData(UBound(EncryptedData) - LBound(EncryptedData) + 1 - IVectorLength), IVector(0), IVectorLength
        Result = True
    End If
    
    EncryptData = Result
    
End Function

' Return the byte length of a string of the length DecryptedTextLength
' encrypted with function EncryptData.
'
' To be called from function FitByteField.
'
' Example:
'   Text = "Careful with that axe, Eugene!"
'   DecryptedTextLength = Len(Text)     ' = 30
'   Length = EncryptedByteLength(DecryptedTextLength)
'   Length -> 112
'
' Example data:
'
' Length plain  Length encrypted
'   0             0
'   1            48
'   4            64
'  12            80
'  20            96
'  28           112
'  36           128
'  44           144
'  52           160
'  60           176
'  68           192
'  76           208
'  84           224
'  92           240
' 100           256
' 108           272
' 116           288
' 124           304
' 132           320
' 140           336
' 148           352
' 156           368
' 164           384
' 172           400
' 180           416
' 188           432
' 196           448
' 204           464
' 212           480
' 220           496
' 227           496
' 227 characters is the largest string to encrypt, if the
' encrypted byte array must fit a Binary field of Access.
' 228           512
' 236           528
' 244           544
' 252           560
' The maximum length of an Access Short Text field is 255 characters.
' 255           560
' 260           576
'
' 2022-02-18. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function EncryptedByteLength( _
    ByVal DecryptedTextLength As Long) _
    As Long
    
    Dim Length  As Long

    If DecryptedTextLength > 0 Then
        Length = 48 + ((DecryptedTextLength + 4) \ 8) * 16
    End If
    
    EncryptedByteLength = Length

End Function

' Return the length of a string encrypted and Base64 encoded with function Encrypt.
'
' To be called from function FitTextField.
'
' Example:
'   Text = "Careful with that axe, Eugene!"
'   DecryptedTextLength = Len(Text)     ' = 30
'   Length = EncryptedTextLength(DecryptedTextLength)
'   Length -> 152
'
' Example data:
'
' Length plain  Length encrypted
'   0             0
'   1            64
'   4            88
'  12           108
'  20           128
'  28           152
'  36           172
'  44           192
'  52           216
'  60           236
'  67           236
' 67 characters is the largest string to encrypt, if the
' encrypted string must fit a Short Text field of Access.
'  68           256
'  76           280
'  84           300
'  92           320
' 100           344
' 108           364
' 116           384
' 124           408
' 132           428
' 140           448
' 148           472
' 156           492
' 164           512
' 172           536
' 180           556
' 188           576
' 196           600
' 204           620
' 212           640
' 220           664
' 228           684
' 236           704
' 244           728
' 252           748
' The maximum length of an Access Short Text field is 255 characters.
' 255           748
' 260           768
'
' 2021-10-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function EncryptedTextLength( _
    ByVal DecryptedTextLength As Long) _
    As Long
    
    Dim Length      As Long
    
    If DecryptedTextLength > 0 Then
        Length = 64 - Int(-(DecryptedTextLength - 3) / 8) * 20 - Int(-(DecryptedTextLength - 3) / 24) * 4
    End If
    
    EncryptedTextLength = Length

End Function

' Check if the length of the encrypted value of argument Text will fit in
' a table field (of data type Binary) having the size of argument FieldSize.
' Return True if it will fit, False if not.
'
' Optionally, if argument ChopText is True:
'   - chop Text to fit the field size
'   - return the chopped text by reference
'   - return True for any length of Text
'
' 2022-02-18. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function FitByteField( _
    ByRef Text As String, _
    Optional ByVal FieldSize As Integer = MaximumBinaryFieldSize, _
    Optional ByVal ChopText As Boolean) _
    As Boolean
    
    Dim Length      As Long
    Dim Success     As Boolean
    
    If Text = "" Then
        Success = True
    Else
        Length = EncryptedByteLength(Len(Text))
        If Length < FieldSize Then
            Success = True
        ElseIf ChopText = True Then
            Text = Left(Text, DecryptedByteLength(Length))
            Success = True
        End If
    End If
    
    FitByteField = Success
    
End Function

' Check if the length of the encrypted value of argument Text will fit in
' a table field (of data type Text) having the size of argument FieldSize.
' Return True if it will fit, False if not.
'
' Optionally, if argument ChopText is True:
'   - chop Text to fit the field size
'   - return the chopped text by reference
'   - return True for any length of Text
'
' 2021-10-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function FitTextField( _
    ByRef Text As String, _
    Optional ByVal FieldSize As Integer = MaximumTextFieldSize, _
    Optional ByVal ChopText As Boolean) _
    As Boolean
    
    Dim Length      As Long
    Dim Success     As Boolean
    
    If Text = "" Then
        Success = True
    Else
        Length = EncryptedTextLength(Len(Text))
        If Length < FieldSize Then
            Success = True
        ElseIf ChopText = True Then
            Text = Left(Text, DecryptedTextLength(Length))
            Success = True
        End If
    End If
    
    FitTextField = Success
    
End Function

' Return a Base64 encoded hash of a string using the specified hash algorithm.
' By default, hash algorithm SHA256 is used.
'
' By default, the Text value is processed as Unicode.
' Optionally, the Text is processed as ANSI which is required for, say, OAuth2 tokens.
' Reference:
'   https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/strconv-function
'
' Example:
'   Text = "Get your filthy hands off my desert."
'   Value = Hash(Text)
'   Value -> "AIPgWDlQLv7bvLdg7Oa78dyRbC0tStuEXJRk0MMehOc="
'
' Length of the generated Base64 encoded hash string:
'
'   Encoding    Length
'   MD2         24
'   MD4         24
'   MD5         24
'   SHA1        28
'   SHA256      44      ' Default.
'   SHA384      64
'   SHA512      88
'
' 2024-08-08. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function Hash( _
    ByVal Text As String, _
    Optional ByVal BcryptHashAlgorithmId As BcHashAlgorithm = BcHashAlgorithm.bcSha256, _
    Optional ByVal AnsiEncoding As Boolean) _
    As String
    
    Dim TextData()          As Byte
    
    Dim HashBase64          As String
    
    If Text = "" Then
        ' No data. Nothing to do.
    Else
        If AnsiEncoding Then
            TextData = StrConv(Text, vbFromUnicode)
        Else
            TextData = Text
        End If
        HashBase64 = ByteBase64(HashData(TextData, BcryptHashAlgorithmId))
    End If
    
    Hash = HashBase64
    
End Function

' Return the byte count of a hash from function HashData.
'
' Default count is 32.
' Maximum count is 64.
' Minimum count is 16.
'
' 2022-02-02. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function HashByteLength( _
    Optional ByVal BcryptHashAlgorithmId As BcHashAlgorithm = bcSha256) _
    As Integer

    Dim Length  As Integer
    
    Select Case BcryptHashAlgorithmId
        Case bcMd2, bcMd4, bcMd5
            Length = 16
        Case bcSha1
            Length = 20
        Case bcSha256
            Length = 32
        Case bcSha384
            Length = 48
        Case bcSha512
            Length = 64
    End Select
    
    HashByteLength = Length
    
End Function

' Create and return the hash of a byte array as another byte array
' using the specified hash algorithm.
' By default, hash algorithm SHA256 is used.
'
' To be called from functions EncryptData, DecryptData, and Hash.
'
' 2021-10-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function HashData( _
    ByRef TextData() As Byte, _
    Optional ByVal BcryptHashAlgorithmId As BcHashAlgorithm = BcHashAlgorithm.bcSha256) _
    As Byte()
    
    Dim Data()          As Byte
    
    Dim TextDataLength  As Long
    
    If StrPtr(TextData) = 0 Then
        ' Array TextData is empty.
    Else
        TextDataLength = UBound(TextData) - LBound(TextData) + 1
        Data = CngHash(VarPtr(TextData(LBound(TextData))), TextDataLength, BcryptHashAlgorithmId)
    End If
    
    HashData = Data
    
End Function

' Return the fixed length of a Base64 encoded hash from function Hash
' using the specified hash algorithm.
'
' Default length is 44.
' Maximum length is 88.
' Minimum length is 24.
'
' 2022-02-02. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function HashTextLength( _
    Optional ByVal BcryptHashAlgorithmId As BcHashAlgorithm = bcSha256) _
    As Integer

    Dim Length  As Integer
    
    Select Case BcryptHashAlgorithmId
        Case bcMd2, bcMd4, bcMd5
            Length = 24
        Case bcSha1
            Length = 28
        Case bcSha256
            Length = 44
        Case bcSha384
            Length = 64
        Case bcSha512
            Length = 88
    End Select
    
    HashTextLength = Length
    
End Function

' Return a Base64 encoded HMAC (Hash-Based Message Authentication Code) hash of a string
' using the specified hash algorithm.
' By default, hash algorithm SHA256 is used.
'
' Example:
'   Text = "Several Species of Small Furry Animals Gathered Together in a Cave and Grooving With a Pict Lyrics."
'   Key = "Sysyphus"
'   Value = Hmac(Text, Key)
'   Value -> "KoSbRMsshQwgNKi9H2uG5NDyl+qdzNM6s2tG8AM9wk8="
'   Value = Hmac(Text, Key, bcSha512)
'   Value -> "+B2HzGp8bWfdE1WihvobFQEbAX91zdDA8fiGXFuAT4fjjBquEU55K5ooSOO3jJYN8NPLdySQf5DVR1y31rHkhQ=="
'
' Length of the generated Base64 encoded hash string:
'
'   Encoding    Length
'   SHA1        28
'   SHA256      44      ' Default.
'   SHA384      64
'   SHA512      88
'
' 2025-11-25. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function Hmac( _
    ByVal Text As String, _
    ByVal Key As String, _
    Optional ByVal BcryptHashAlgorithmId As BcHashAlgorithm = BcHashAlgorithm.bcSha256) _
    As String
    
    Dim TextData()          As Byte
    Dim KeyData()           As Byte
    
    Dim HmacBase64          As String
    
    If Text = "" Or Key = "" Then
        ' No data. Nothing to do.
    Else
        ' UTF-8 encode Text and Key.
        TextData = TextUtf8Bytes(Text)
        KeyData = TextUtf8Bytes(Key)
        HmacBase64 = ByteBase64(HmacData(TextData, KeyData, BcryptHashAlgorithmId))
    End If
    
    Hmac = HmacBase64
    
End Function

' Create and return the HMAC (Hash-Based Message Authentication Code) hash of a byte array as
' another byte array using the specified hash algorithm.
' By default, hash algorithm SHA256 is used.
'
' To be called from function Hmac.
'
' 2025-11-26. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function HmacData( _
    ByRef TextData() As Byte, _
    ByRef KeyData() As Byte, _
    Optional ByVal BcryptHashAlgorithmId As BcHashAlgorithm = BcHashAlgorithm.bcSha256) _
    As Byte()
    
    Dim Data()          As Byte
    
    Dim TextDataLength  As Long
    Dim KeyDataLength   As Long
    
    If StrPtr(TextData) = 0 Then
        ' Array TextData is empty.
    ElseIf StrPtr(KeyData) = 0 Then
        ' Array KeyData is empty.
    Else
        TextDataLength = UBound(TextData) - LBound(TextData) + 1
        KeyDataLength = UBound(KeyData) - LBound(KeyData) + 1
        Data = CngHmac(VarPtr(TextData(LBound(TextData))), TextDataLength, VarPtr(KeyData(LBound(KeyData))), KeyDataLength, BcryptHashAlgorithmId)
    End If
    
    HmacData = Data
    
End Function

' Return True if the passed text value represents a value of
' enum BcHashAlgorithm.
' Note: To validate, HashAlgorithm must be in UPPERCASE.
'
' 2021-10-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function IsBcryptHashAlgorithm( _
    ByVal HashAlgorithm As String) _
    As Boolean
    
    Dim Index           As BcHashAlgorithm
    Dim Result          As Boolean
    
    For Index = BcHashAlgorithm.[_First] To BcHashAlgorithm.[_Last]
        If BcryptHashAlgorithm(Index) = HashAlgorithm Then
            Result = True
            Exit For
        End If
    Next
    
    IsBcryptHashAlgorithm = Result
    
End Function

' Return True if the passed value of enum BcHashAlgorithm is valid.
'
' To be called from function CngHash.
'
' 2021-10-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function IsBcryptHashAlgorithmId( _
    ByVal HashAlgorithmId As BcHashAlgorithm) _
    As Boolean
    
    Dim Result          As Boolean
    
    If BcHashAlgorithm.[_First] <= HashAlgorithmId And HashAlgorithmId <= BcHashAlgorithm.[_Last] Then
        Result = True
    End If
    
    IsBcryptHashAlgorithmId = Result
    
End Function

' Return True if the passed text value represents a value of
' enum BcHashAlgorithm valid for HMAC.
' Note: To validate, HashAlgorithm must be in UPPERCASE.
'
' 2025-11-25. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function IsBcryptHashHmacAlgorithm( _
    ByVal HashAlgorithm As String) _
    As Boolean
    
    Dim Index           As BcHashAlgorithm
    Dim Result          As Boolean
    
    For Index = BcHashAlgorithm.[_FirstHmac] To BcHashAlgorithm.[_LastHmac]
        If BcryptHashAlgorithm(Index) = HashAlgorithm Then
            Result = True
            Exit For
        End If
    Next
    
    IsBcryptHashHmacAlgorithm = Result
    
End Function

' Return True if the passed value of enum BcHashAlgorithm is valid for HMAC.
'
' To be called from function CngHmac.
'
' 2025-11-25. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function IsBcryptHashHmacAlgorithmId( _
    ByVal HashAlgorithmId As BcHashAlgorithm) _
    As Boolean
    
    Dim Result          As Boolean
    
    If BcHashAlgorithm.[_FirstHmac] <= HashAlgorithmId And HashAlgorithmId <= BcHashAlgorithm.[_LastHmac] Then
        Result = True
    End If
    
    IsBcryptHashHmacAlgorithmId = Result
    
End Function

' Return True if the passed text value represents a value of
' enum BcRandomAlgorithm.
' Note: To validate, RandomAlgorithm must be in UPPERCASE.
'
' 2021-10-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function IsBcryptRandomAlgorithm( _
    ByVal RandomAlgorithm As String) _
    As Boolean
    
    Dim Index           As BcRandomAlgorithm
    Dim Result          As Boolean
    
    For Index = BcRandomAlgorithm.[_First] To BcRandomAlgorithm.[_Last]
        If BcryptRandomAlgorithm(Index) = RandomAlgorithm Then
            Result = True
            Exit For
        End If
    Next
    
    IsBcryptRandomAlgorithm = Result
    
End Function

' Return True if the passed value of enum BcRandomAlgorithm
' is valid.
'
' To be called from function CngRandom.
'
' 2021-10-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function IsBcryptRandomAlgorithmId( _
    ByVal RandomAlgorithmId As BcRandomAlgorithm) _
    As Boolean
    
    Dim Result          As Boolean
    
    If BcRandomAlgorithm.[_First] <= RandomAlgorithmId And RandomAlgorithmId <= BcRandomAlgorithm.[_Last] Then
        Result = True
    End If
    
    IsBcryptRandomAlgorithmId = Result
    
End Function

' Return a random string of the length specified using the
' specified random algorithm.
' By default, random algorithm RNG is used.
' Optionally, random algorithm Fips186DSARng can be specified.
'
' Optionally, if argument TrueBase64 is True, a random byte array of
' the length specified will be created and returned Base64 encoded.
'
' Examples:
'   Value = Random(23)
'   Value -> "ZXq2fue8QZ+d83Lw0T4Mi68"
'
'   Value = Random(23, , True)
'   Value -> "M/KlMXO8l2FP9Kh0GtWcAr63a4UVqZw="
'
' 2022-02-20. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function Random( _
    ByVal TextSize As Long, _
    Optional ByVal BcryptRandomAlgorithmId As BcRandomAlgorithm = BcRandomAlgorithm.bcRng, _
    Optional ByVal TrueBase64 As Boolean) _
    As String
    
    Dim Text    As String
    
    If TextSize <= 0 Then
        ' Nothing to do.
    ElseIf TrueBase64 Then
        Text = ByteBase64(RandomData(TextSize, BcryptRandomAlgorithmId))
    ElseIf TextSize < 4 Then
        Text = Left(ByteBase64(RandomData(TextSize, BcryptRandomAlgorithmId)), TextSize)
    Else
        Text = Left(ByteBase64(RandomData(TextSize * 0.75, BcryptRandomAlgorithmId)), TextSize)
    End If
    
    Random = Text
    
End Function

' Return a byte array with random bytes using the specified random algorithm.
' By default, random algorithm RNG is used.
'
' To be called from function Random or EncryptData.
'
' Allowed algorithms (NB: Random algorithms only; check OS support):
'   https://docs.microsoft.com/en-us/windows/desktop/SecCNG/cng-algorithm-identifiers
'
' 2022-02-20. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function RandomData( _
    ByVal DataLength As Long, _
    Optional ByVal BcryptRandomAlgorithmId As BcRandomAlgorithm = BcRandomAlgorithm.bcRng)

    Dim Data()      As Byte
    
    If DataLength <= 0 Then
        ' Nothing to do.
    Else
        ReDim Data(0 To DataLength - 1)
        CngRandom VarPtr(Data(LBound(Data))), DataLength, BcryptRandomAlgorithmId
    End If
    
    RandomData = Data
    
End Function

' Convert and return a plain string as a Base64 encoded string.
' Generic function.
'
' 2021-10-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function TextBase64( _
    ByVal Text As String) _
    As String
    
    Dim Text64          As String

    Text64 = ByteBase64(StrConv(Text, vbFromUnicode))
    
    TextBase64 = Text64

End Function

' Convert a string to an UTF-8 encoded byte array.
' Returns the converted string as an array.
'
' 2025-11-25. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function TextUtf8Bytes( _
    ByVal Text As String) _
    As Byte()
    
    Const Flags             As Long = 0
    Const DefaultChar       As LongPtr = 0
    Const UsedDefaultChar   As LongPtr = 0
    Const MultiByte         As Long = 0
    Const MultiByteStr      As LongPtr = 0
    
    Dim Data()              As Byte
    
    Dim Length              As Long
    
    Length = WideCharToMultiByte(CpUtf8, Flags, StrPtr(Text), Len(Text), MultiByteStr, MultiByte, DefaultChar, UsedDefaultChar)
    If Length > 0 Then
        ReDim Data(0 To Length - 1)
        Call WideCharToMultiByte(CpUtf8, Flags, StrPtr(Text), Len(Text), VarPtr(Data(0)), Length, DefaultChar, UsedDefaultChar)
    End If
    
    TextUtf8Bytes = Data
    
End Function

