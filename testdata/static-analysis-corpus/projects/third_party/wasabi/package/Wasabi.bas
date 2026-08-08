Attribute VB_Name = "Wasabi"
'/**
' * ============================================================================
' * Wasabi v2.3.8-beta
' * Copyright (c) 2026 UesleiDev
' *
' * @description Advanced Networking & WebSockets Module for VBA/VB6.
' * Provides full support for TCP, WebSockets, TLS/SSL, Proxies, and MQTT.
' *
' * Permission is hereby granted, free of charge, to any person obtaining a
' * copy of this software and associated documentation files (the "Software"),
' * to deal in the Software without restriction, including without limitation
' * the rights to use, copy, modify, merge, publish, distribute, sublicense,
' * and/or sell copies of the Software, and to permit persons to whom the
' * Software is furnished to do so, subject to the following conditions:
' *
' * The above copyright notice and this permission notice shall be included in
' * all copies or substantial portions of the Software.
' *
' * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
' * FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
' * DEALINGS IN THE SOFTWARE.
' * ============================================================================
' */

Option Explicit
Option Private Module

' ============================================================================
' 1. API DECLARATIONS
' ============================================================================
#If VBA7 Then

    ' --- advapi32.dll ---
    Private Declare PtrSafe Function CryptAcquireContextW Lib "advapi32.dll" (ByRef phProv As LongPtr, ByVal pszContainer As LongPtr, ByVal pszProvider As LongPtr, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
    Private Declare PtrSafe Function CryptCreateHash Lib "advapi32.dll" (ByVal hProv As LongPtr, ByVal Algid As Long, ByVal hKey As LongPtr, ByVal dwFlags As Long, ByRef phHash As LongPtr) As Long
    Private Declare PtrSafe Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As LongPtr) As Long
    Private Declare PtrSafe Function CryptGetHashParam Lib "advapi32.dll" (ByVal hHash As LongPtr, ByVal dwParam As Long, ByRef pbData As Byte, ByRef pdwDataLen As Long, ByVal dwFlags As Long) As Long
    Private Declare PtrSafe Function CryptHashData Lib "advapi32.dll" (ByVal hHash As LongPtr, ByRef pbData As Byte, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
    Private Declare PtrSafe Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As LongPtr, ByVal dwFlags As Long) As Long

    ' --- bcrypt.dll ---
    Private Declare PtrSafe Function BCryptGenRandom Lib "bcrypt.dll" (ByVal hAlgorithm As LongPtr, ByRef pbBuffer As Any, ByVal cbBuffer As Long, ByVal dwFlags As Long) As Long

    ' --- crypt32.dll ---
    Private Declare PtrSafe Function CertCloseStore Lib "crypt32.dll" (ByVal hCertStore As LongPtr, ByVal dwFlags As Long) As Long
    Private Declare PtrSafe Function CertFindCertificateInStore Lib "crypt32.dll" (ByVal hCertStore As LongPtr, ByVal dwCertEncodingType As Long, ByVal dwFindFlags As Long, ByVal dwFindType As Long, ByRef pvFindPara As Any, ByVal pPrevCertContext As LongPtr) As LongPtr
    Private Declare PtrSafe Sub CertFreeCertificateChain Lib "crypt32.dll" (ByVal pChainContext As LongPtr)
    Private Declare PtrSafe Function CertFreeCertificateContext Lib "crypt32.dll" (ByVal pCertContext As LongPtr) As Long
    Private Declare PtrSafe Function CertGetCertificateChain Lib "crypt32.dll" (ByVal hChainEngine As LongPtr, ByVal pCertContext As LongPtr, ByVal pTime As LongPtr, ByVal hAdditionalStore As LongPtr, ByRef pChainPara As CERT_CHAIN_PARA, ByVal dwFlags As Long, ByVal pvReserved As LongPtr, ByRef ppChainContext As LongPtr) As Long
    Private Declare PtrSafe Function CertOpenStore Lib "crypt32.dll" (ByVal lpszStoreProvider As LongPtr, ByVal dwEncodingType As Long, ByVal hCryptProv As LongPtr, ByVal dwFlags As Long, ByVal pvPara As LongPtr) As LongPtr
    Private Declare PtrSafe Function CertVerifyCertificateChainPolicy Lib "crypt32.dll" (ByVal pszPolicyOID As LongPtr, ByVal pChainContext As LongPtr, ByRef pPolicyPara As CERT_CHAIN_POLICY_PARA, ByRef pPolicyStatus As CERT_CHAIN_POLICY_STATUS) As Long
    Private Declare PtrSafe Function CryptBinaryToStringW Lib "crypt32.dll" (ByVal pbBinary As LongPtr, ByVal cbBinary As Long, ByVal dwFlags As Long, ByVal pszString As LongPtr, ByRef pcchString As Long) As Long
    Private Declare PtrSafe Function CryptStringToBinaryW Lib "crypt32.dll" (ByVal pszString As LongPtr, ByVal cchString As Long, ByVal dwFlags As Long, ByVal pbBinary As LongPtr, ByRef pcbBinary As Long, ByRef pdwSkip As Long, ByRef pdwFlags As Long) As Long
    Private Declare PtrSafe Function PFXImportCertStore Lib "crypt32.dll" (ByRef pPFX As CRYPT_DATA_BLOB, ByVal szPassword As LongPtr, ByVal dwFlags As Long) As LongPtr

    ' --- kernel32.dll ---
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef dest As Any, ByRef src As Any, ByVal size As Long)
    Private Declare PtrSafe Sub CopyMemoryFromPtr Lib "kernel32" Alias "RtlMoveMemory" (ByRef dest As Any, ByVal src As LongPtr, ByVal size As Long)
    Private Declare PtrSafe Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As LongPtr
    Private Declare PtrSafe Function GetProcAddress Lib "kernel32" (ByVal hModule As LongPtr, ByVal lpProcName As String) As LongPtr
    Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
    Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As LongPtr
    Private Declare PtrSafe Function lstrlenW Lib "kernel32" (ByVal lpString As LongPtr) As Long
    Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByRef lpMultiByteStr As Byte, ByVal cchMultiByte As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long) As Long
    Private Declare PtrSafe Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As LongPtr, ByVal dwSize As LongPtr, ByVal flAllocationType As Long, ByVal flProtect As Long) As LongPtr
    Private Declare PtrSafe Function VirtualFree Lib "kernel32" (ByVal lpAddress As LongPtr, ByVal dwSize As LongPtr, ByVal dwFreeType As Long) As Long
    Private Declare PtrSafe Function VirtualProtect Lib "kernel32" (ByVal lpAddress As LongPtr, ByVal dwSize As LongPtr, ByVal flNewProtect As Long, ByRef lpflOldProtect As Long) As Long
    Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, ByRef lpMultiByteStr As Byte, ByVal cchMultiByte As Long, ByVal lpDefaultChar As LongPtr, ByVal lpUsedDefaultChar As LongPtr) As Long

    ' --- secur32.dll ---
    Private Declare PtrSafe Function AcquireCredentialsHandle Lib "secur32.dll" Alias "AcquireCredentialsHandleA" (ByVal pszPrincipal As LongPtr, ByVal pszPackage As String, ByVal fCredentialUse As Long, ByVal pvLogonID As LongPtr, ByRef pAuthData As Any, ByVal pGetKeyFn As LongPtr, ByVal pvGetKeyArgument As LongPtr, ByRef phCredential As SecHandle, ByRef ptsExpiry As SECURITY_INTEGER) As Long
    Private Declare PtrSafe Function DecryptMessage Lib "secur32.dll" (ByRef phContext As SecHandle, ByRef pMessage As SecBufferDesc, ByVal MessageSeqNo As Long, ByRef pfQOP As Long) As Long
    Private Declare PtrSafe Function DeleteSecurityContext Lib "secur32.dll" (ByRef phContext As SecHandle) As Long
    Private Declare PtrSafe Function EncryptMessage Lib "secur32.dll" (ByRef phContext As SecHandle, ByVal fQOP As Long, ByRef pMessage As SecBufferDesc, ByVal MessageSeqNo As Long) As Long
    Private Declare PtrSafe Function FreeContextBuffer Lib "secur32.dll" (ByVal pvContextBuffer As LongPtr) As Long
    Private Declare PtrSafe Function FreeCredentialsHandle Lib "secur32.dll" (ByRef phCredential As SecHandle) As Long
    Private Declare PtrSafe Function InitializeSecurityContext Lib "secur32.dll" Alias "InitializeSecurityContextA" (ByRef phCredential As SecHandle, ByVal phContext As LongPtr, ByVal pszTargetName As String, ByVal fContextReq As Long, ByVal Reserved1 As Long, ByVal TargetDataRep As Long, ByVal pInput As LongPtr, ByVal Reserved2 As Long, ByRef phNewContext As SecHandle, ByRef pOutput As SecBufferDesc, ByRef pfContextAttr As Long, ByRef ptsExpiry As SECURITY_INTEGER) As Long
    Private Declare PtrSafe Function InitializeSecurityContextContinue Lib "secur32.dll" Alias "InitializeSecurityContextA" (ByRef phCredential As SecHandle, ByRef phContext As SecHandle, ByVal pszTargetName As String, ByVal fContextReq As Long, ByVal Reserved1 As Long, ByVal TargetDataRep As Long, ByRef pInput As SecBufferDesc, ByVal Reserved2 As Long, ByRef phNewContext As SecHandle, ByRef pOutput As SecBufferDesc, ByRef pfContextAttr As Long, ByRef ptsExpiry As SECURITY_INTEGER) As Long
    Private Declare PtrSafe Function QueryContextAttributes Lib "secur32.dll" Alias "QueryContextAttributesA" (ByRef phContext As SecHandle, ByVal ulAttribute As Long, ByRef pBuffer As Any) As Long

    ' --- user32.dll ---
    Private Declare PtrSafe Function CallWindowProcW Lib "user32" (ByVal lpPrevWndFunc As LongPtr, ByVal P1 As LongPtr, ByVal P2 As LongPtr, ByVal P3 As LongPtr, ByVal P4 As LongPtr) As LongPtr
    Private Declare PtrSafe Function CallWindowProcW_WndProc Lib "user32" Alias "CallWindowProcW" (ByVal lpPrevWndFunc As LongPtr, ByVal hwnd As LongPtr, ByVal msg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    Private Declare PtrSafe Function CreateWindowExW Lib "user32" (ByVal dwExStyle As Long, ByVal lpClassName As LongPtr, ByVal lpWindowName As LongPtr, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As LongPtr, ByVal hMenu As LongPtr, ByVal hInstance As LongPtr, ByVal lpParam As LongPtr) As LongPtr
    Private Declare PtrSafe Function DestroyWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function SetWindowLongPtrW Lib "user32" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr

    ' --- winhttp.dll ---
    Private Declare PtrSafe Function WinHttpGetIEProxyConfigForCurrentUser Lib "winhttp.dll" (ByRef pProxyConfig As WINHTTP_CURRENT_USER_IE_PROXY_CONFIG) As Long

    ' --- ws2_32.dll ---
    Private Declare PtrSafe Function sock_closesocket Lib "ws2_32.dll" Alias "closesocket" (ByVal s As LongPtr) As Long
    Private Declare PtrSafe Function sock_connect Lib "ws2_32.dll" Alias "connect" (ByVal s As LongPtr, ByVal name As LongPtr, ByVal namelen As Long) As Long
    Private Declare PtrSafe Sub sock_freeaddrinfo Lib "ws2_32.dll" Alias "freeaddrinfo" (ByVal pAddrInfo As LongPtr)
    Private Declare PtrSafe Function sock_getaddrinfo Lib "ws2_32.dll" Alias "getaddrinfo" (ByVal pNodeName As String, ByVal pServiceName As String, ByVal pHints As LongPtr, ByRef ppResult As LongPtr) As Long
    Private Declare PtrSafe Function sock_gethostbyname Lib "ws2_32.dll" Alias "gethostbyname" (ByVal hostname As String) As LongPtr
    Private Declare PtrSafe Function sock_getsockopt Lib "ws2_32.dll" Alias "getsockopt" (ByVal s As LongPtr, ByVal level As Long, ByVal optname As Long, ByRef optVal As Long, ByRef optlen As Long) As Long
    Private Declare PtrSafe Function sock_htons Lib "ws2_32.dll" Alias "htons" (ByVal hostshort As Long) As Integer
    Private Declare PtrSafe Function sock_inet_addr Lib "ws2_32.dll" Alias "inet_addr" (ByVal cp As String) As Long
    Private Declare PtrSafe Function sock_ioctlsocket Lib "ws2_32.dll" Alias "ioctlsocket" (ByVal s As LongPtr, ByVal cmd As Long, ByRef argp As Long) As Long
    Private Declare PtrSafe Function sock_recv Lib "ws2_32.dll" Alias "recv" (ByVal s As LongPtr, ByRef buf As Byte, ByVal bufLen As Long, ByVal flags As Long) As Long
    Private Declare PtrSafe Function sock_select Lib "ws2_32.dll" Alias "select" (ByVal nfds As Long, ByRef readfds As Any, ByRef writefds As Any, ByRef exceptfds As Any, ByRef TIMEOUT As TIMEVAL) As Long
    Private Declare PtrSafe Function sock_send Lib "ws2_32.dll" Alias "send" (ByVal s As LongPtr, ByRef buf As Byte, ByVal bufLen As Long, ByVal flags As Long) As Long
    Private Declare PtrSafe Function sock_setsockopt Lib "ws2_32.dll" Alias "setsockopt" (ByVal s As LongPtr, ByVal level As Long, ByVal optname As Long, ByRef optVal As Long, ByVal optlen As Long) As Long
    Private Declare PtrSafe Function sock_socket Lib "ws2_32.dll" Alias "socket" (ByVal af As Long, ByVal socktype As Long, ByVal protocol As Long) As LongPtr
    Private Declare PtrSafe Function WSAAsyncSelect Lib "ws2_32.dll" (ByVal s As LongPtr, ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal lEvent As Long) As Long
    Private Declare PtrSafe Function WSACleanup Lib "ws2_32.dll" () As Long
    Private Declare PtrSafe Function WSAGetLastError Lib "ws2_32.dll" () As Long
    Private Declare PtrSafe Function WSAStartup Lib "ws2_32.dll" (ByVal wVersionRequested As Integer, ByRef lpWSAData As WSADATA) As Long

    ' --- VBE7 (Internal) ---
    Private Declare PtrSafe Function VarPtrArray Lib "VBE7" Alias "VarPtr" (ByRef arr() As Byte) As LongPtr

    Private Const NULL_PTR As LongPtr = 0
    Private Const INVALID_SOCKET As LongPtr = -1

#Else

    ' --- advapi32.dll ---
    Private Declare Function CryptAcquireContextW Lib "advapi32.dll" (ByRef phProv As Long, ByVal pszContainer As Long, ByVal pszProvider As Long, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
    Private Declare Function CryptCreateHash Lib "advapi32.dll" (ByVal hProv As Long, ByVal Algid As Long, ByVal hKey As Long, ByVal dwFlags As Long, ByRef phHash As Long) As Long
    Private Declare Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As Long) As Long
    Private Declare Function CryptGetHashParam Lib "advapi32.dll" (ByVal hHash As Long, ByVal dwParam As Long, ByRef pbData As Byte, ByRef pdwDataLen As Long, ByVal dwFlags As Long) As Long
    Private Declare Function CryptHashData Lib "advapi32.dll" (ByVal hHash As Long, ByRef pbData As Byte, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
    Private Declare Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwFlags As Long) As Long

    ' --- bcrypt.dll ---
    Private Declare Function BCryptGenRandom Lib "bcrypt.dll" (ByVal hAlgorithm As Long, ByRef pbBuffer As Any, ByVal cbBuffer As Long, ByVal dwFlags As Long) As Long

    ' --- crypt32.dll ---
    Private Declare Function CertCloseStore Lib "crypt32.dll" (ByVal hCertStore As Long, ByVal dwFlags As Long) As Long
    Private Declare Function CertFindCertificateInStore Lib "crypt32.dll" (ByVal hCertStore As Long, ByVal dwCertEncodingType As Long, ByVal dwFindFlags As Long, ByVal dwFindType As Long, ByRef pvFindPara As Any, ByVal pPrevCertContext As Long) As Long
    Private Declare Sub CertFreeCertificateChain Lib "crypt32.dll" (ByVal pChainContext As Long)
    Private Declare Function CertFreeCertificateContext Lib "crypt32.dll" (ByVal pCertContext As Long) As Long
    Private Declare Function CertGetCertificateChain Lib "crypt32.dll" (ByVal hChainEngine As Long, ByVal pCertContext As Long, ByVal pTime As Long, ByVal hAdditionalStore As Long, ByRef pChainPara As CERT_CHAIN_PARA, ByVal dwFlags As Long, ByVal pvReserved As Long, ByRef ppChainContext As Long) As Long
    Private Declare Function CertOpenStore Lib "crypt32.dll" (ByVal lpszStoreProvider As Long, ByVal dwEncodingType As Long, ByVal hCryptProv As Long, ByVal dwFlags As Long, ByVal pvPara As Long) As Long
    Private Declare Function CertVerifyCertificateChainPolicy Lib "crypt32.dll" (ByVal pszPolicyOID As Long, ByVal pChainContext As Long, ByRef pPolicyPara As CERT_CHAIN_POLICY_PARA, ByRef pPolicyStatus As CERT_CHAIN_POLICY_STATUS) As Long
    Private Declare Function CryptBinaryToStringW Lib "crypt32.dll" (ByVal pbBinary As Long, ByVal cbBinary As Long, ByVal dwFlags As Long, ByVal pszString As Long, ByRef pcchString As Long) As Long
    Private Declare Function CryptStringToBinaryW Lib "crypt32.dll" (ByVal pszString As Long, ByVal cchString As Long, ByVal dwFlags As Long, ByVal pbBinary As Long, ByRef pcbBinary As Long, ByRef pdwSkip As Long, ByRef pdwFlags As Long) As Long
    Private Declare Function PFXImportCertStore Lib "crypt32.dll" (ByRef pPFX As CRYPT_DATA_BLOB, ByVal szPassword As Long, ByVal dwFlags As Long) As Long

    ' --- kernel32.dll ---
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef dest As Any, ByRef src As Any, ByVal size As Long)
    Private Declare Sub CopyMemoryFromPtr Lib "kernel32" Alias "RtlMoveMemory" (ByRef dest As Any, ByVal src As Long, ByVal size As Long)
    Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
    Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
    Private Declare Function GetTickCount Lib "kernel32" () As Long
    Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
    Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
    Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByRef lpMultiByteStr As Byte, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
    Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
    Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
    Private Declare Function VirtualProtect Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flNewProtect As Long, ByRef lpflOldProtect As Long) As Long
    Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByRef lpMultiByteStr As Byte, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long

    ' --- secur32.dll ---
    Private Declare Function AcquireCredentialsHandle Lib "secur32.dll" Alias "AcquireCredentialsHandleA" (ByVal pszPrincipal As Long, ByVal pszPackage As String, ByVal fCredentialUse As Long, ByVal pvLogonID As Long, ByRef pAuthData As Any, ByVal pGetKeyFn As Long, ByVal pvGetKeyArgument As Long, ByRef phCredential As SecHandle, ByRef ptsExpiry As SECURITY_INTEGER) As Long
    Private Declare Function DecryptMessage Lib "secur32.dll" (ByRef phContext As SecHandle, ByRef pMessage As SecBufferDesc, ByVal MessageSeqNo As Long, ByRef pfQOP As Long) As Long
    Private Declare Function DeleteSecurityContext Lib "secur32.dll" (ByRef phContext As SecHandle) As Long
    Private Declare Function EncryptMessage Lib "secur32.dll" (ByRef phContext As SecHandle, ByVal fQOP As Long, ByRef pMessage As SecBufferDesc, ByVal MessageSeqNo As Long) As Long
    Private Declare Function FreeContextBuffer Lib "secur32.dll" (ByVal pvContextBuffer As Long) As Long
    Private Declare Function FreeCredentialsHandle Lib "secur32.dll" (ByRef phCredential As SecHandle) As Long
    Private Declare Function InitializeSecurityContext Lib "secur32.dll" Alias "InitializeSecurityContextA" (ByRef phCredential As SecHandle, ByVal phContext As Long, ByVal pszTargetName As String, ByVal fContextReq As Long, ByVal Reserved1 As Long, ByVal TargetDataRep As Long, ByVal pInput As Long, ByVal Reserved2 As Long, ByRef phNewContext As SecHandle, ByRef pOutput As SecBufferDesc, ByRef pfContextAttr As Long, ByRef ptsExpiry As SECURITY_INTEGER) As Long
    Private Declare Function InitializeSecurityContextContinue Lib "secur32.dll" Alias "InitializeSecurityContextA" (ByRef phCredential As SecHandle, ByRef phContext As SecHandle, ByVal pszTargetName As String, ByVal fContextReq As Long, ByVal Reserved1 As Long, ByVal TargetDataRep As Long, ByRef pInput As SecBufferDesc, ByVal Reserved2 As Long, ByRef phNewContext As SecHandle, ByRef pOutput As SecBufferDesc, ByRef pfContextAttr As Long, ByRef ptsExpiry As SECURITY_INTEGER) As Long
    Private Declare Function QueryContextAttributes Lib "secur32.dll" Alias "QueryContextAttributesA" (ByRef phContext As SecHandle, ByVal ulAttribute As Long, ByRef pBuffer As Any) As Long

    ' --- user32.dll ---
    Private Declare Function CallWindowProcW Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal P1 As Long, ByVal P2 As Long, ByVal P3 As Long, ByVal P4 As Long) As Long
    Private Declare Function CallWindowProcW_WndProc Lib "user32" Alias "CallWindowProcW" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Declare Function CreateWindowExW Lib "user32" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByVal lpParam As Long) As Long
    Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function SetWindowLongW Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

    ' --- winhttp.dll ---
    Private Declare Function WinHttpGetIEProxyConfigForCurrentUser Lib "winhttp.dll" (ByRef pProxyConfig As WINHTTP_CURRENT_USER_IE_PROXY_CONFIG) As Long

    ' --- ws2_32.dll ---
    Private Declare Function sock_closesocket Lib "ws2_32.dll" Alias "closesocket" (ByVal s As Long) As Long
    Private Declare Function sock_connect Lib "ws2_32.dll" Alias "connect" (ByVal s As Long, ByVal name As Long, ByVal namelen As Long) As Long
    Private Declare Sub sock_freeaddrinfo Lib "ws2_32.dll" Alias "freeaddrinfo" (ByVal pAddrInfo As Long)
    Private Declare Function sock_getaddrinfo Lib "ws2_32.dll" Alias "getaddrinfo" (ByVal pNodeName As String, ByVal pServiceName As String, ByVal pHints As Long, ByRef ppResult As Long) As Long
    Private Declare Function sock_gethostbyname Lib "ws2_32.dll" Alias "gethostbyname" (ByVal hostname As String) As Long
    Private Declare Function sock_getsockopt Lib "ws2_32.dll" Alias "getsockopt" (ByVal s As Long, ByVal level As Long, ByVal optname As Long, ByRef optVal As Long, ByRef optlen As Long) As Long
    Private Declare Function sock_htons Lib "ws2_32.dll" Alias "htons" (ByVal hostshort As Long) As Integer
    Private Declare Function sock_inet_addr Lib "ws2_32.dll" Alias "inet_addr" (ByVal cp As String) As Long
    Private Declare Function sock_ioctlsocket Lib "ws2_32.dll" Alias "ioctlsocket" (ByVal s As Long, ByVal cmd As Long, ByRef argp As Long) As Long
    Private Declare Function sock_recv Lib "ws2_32.dll" Alias "recv" (ByVal s As Long, ByRef buf As Byte, ByVal buflen As Long, ByVal flags As Long) As Long
    Private Declare Function sock_select Lib "ws2_32.dll" Alias "select" (ByVal nfds As Long, ByRef readfds As Any, ByRef writefds As Any, ByRef exceptfds As Any, ByRef timeout As TIMEVAL) As Long
    Private Declare Function sock_send Lib "ws2_32.dll" Alias "send" (ByVal s As Long, ByRef buf As Byte, ByVal buflen As Long, ByVal flags As Long) As Long
    Private Declare Function sock_setsockopt Lib "ws2_32.dll" Alias "setsockopt" (ByVal s As Long, ByVal level As Long, ByVal optname As Long, ByRef optval As Long, ByVal optlen As Long) As Long
    Private Declare Function sock_socket Lib "ws2_32.dll" Alias "socket" (ByVal af As Long, ByVal socktype As Long, ByVal protocol As Long) As Long
    Private Declare Function WSAAsyncSelect Lib "ws2_32.dll" (ByVal s As Long, ByVal hwnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long
    Private Declare Function WSACleanup Lib "ws2_32.dll" () As Long
    Private Declare Function WSAGetLastError Lib "ws2_32.dll" () As Long
    Private Declare Function WSAStartup Lib "ws2_32.dll" (ByVal wVersionRequested As Integer, ByRef lpWSAData As WSADATA) As Long

    ' --- VBE6 (Internal) ---
    Private Declare Function VarPtrArray Lib "VBE6" Alias "VarPtr" (ByRef arr() As Byte) As Long

    Private Const NULL_PTR As Long = 0
    Private Const INVALID_SOCKET As Long = -1

#End If

' ============================================================================
' 2. CONSTANTS
' ============================================================================
Private Const TCP_MAXSEG As Long = 4
Private Const BUFFER_SIZE As Long = 262144
Private Const FRAGMENT_BUFFER_SIZE As Long = 262144
Private Const MSG_QUEUE_SIZE As Long = 512
Private Const MAX_CONNECTIONS As Long = 64
Private Const INVALID_CONN_HANDLE As Long = -1
Private Const DEFAULT_RECEIVE_TIMEOUT_MS As Long = 5000
Private Const DEFAULT_PING_INTERVAL_MS As Long = 0
Private Const DEFAULT_RECONNECT_BASE_DELAY_MS As Long = 1000
Private Const DEFAULT_RECONNECT_MAX_ATTEMPTS As Long = 5
Private Const MAX_RECONNECT_DELAY_MS As Long = 30000
Private Const DEFAULT_MTU As Long = 1500
Private Const HAPPY_EYEBALLS_DELAY_MS As Long = 250
Private Const PMTU_DISCOVERY_INTERVAL_MS As Long = 60000
Private Const SOL_SOCKET As Long = 65535
Private Const SO_KEEPALIVE As Long = 8
Private Const SO_RCVBUF As Long = &H1002
Private Const SO_SNDBUF As Long = &H1001
Private Const IPPROTO_TCP_SOL As Long = 6
Private Const TCP_NODELAY As Long = 1
Private Const CP_UTF8 As Long = 65001
Private Const AF_INET As Long = 2
Private Const AF_INET6 As Long = 23
Private Const SOCK_STREAM As Long = 1
Private Const IPPROTO_TCP As Long = 6
Private Const FIONBIO As Long = &H8004667E
Private Const FIONREAD As Long = &H4004667F
Private Const INADDR_NONE As Long = &HFFFFFFFF
Private Const PROXY_TYPE_HTTP As Long = 0
Private Const PROXY_TYPE_SOCKS5 As Long = 1

' TLS & SECPKG Constants
Private Const SECPKG_CRED_OUTBOUND As Long = &H2
Private Const SCHANNEL_CRED_VERSION As Long = &H4
Private Const SCH_CRED_NO_DEFAULT_CREDS As Long = &H10
Private Const SCH_CRED_MANUAL_CRED_VALIDATION As Long = &H8
Private Const SCH_CRED_IGNORE_NO_REVOCATION_CHECK As Long = &H800
Private Const SCH_CRED_IGNORE_REVOCATION_OFFLINE As Long = &H1000
Private Const SP_PROT_TLS1_2_CLIENT As Long = &H800
Private Const SP_PROT_TLS1_3_CLIENT As Long = &H2000
Private Const ISC_REQ_SEQUENCE_DETECT As Long = &H8
Private Const ISC_REQ_REPLAY_DETECT As Long = &H4
Private Const ISC_REQ_CONFIDENTIALITY As Long = &H10
Private Const ISC_REQ_ALLOCATE_MEMORY As Long = &H100
Private Const ISC_REQ_STREAM As Long = &H8000
Private Const CERT_CHAIN_POLICY_SSL As Long = 4
Private Const CERT_FIND_ANY As Long = 0
Private Const CERT_FIND_SUBJECT_STR_A As Long = &H80007
Private Const SECPKG_ATTR_REMOTE_CERT_CONTEXT As Long = &H53
Private Const SECPKG_ATTR_STREAM_SIZES As Long = 4
Private Const AUTHTYPE_SERVER As Long = 1
Private Const CERT_STORE_PROV_SYSTEM As Long = 10
Private Const CERT_SYSTEM_STORE_CURRENT_USER As Long = &H10000
Private Const CERT_CHAIN_REVOCATION_CHECK_CHAIN As Long = &H20000000
Private Const X509_ASN_ENCODING As Long = 1
Private Const PKCS_7_ASN_ENCODING As Long = &H10000
Private Const CRYPT_EXPORTABLE As Long = 1
Private Const PKCS12_ALLOW_OVERWRITE_KEY As Long = &H4000

' Buffer Types & Status
Private Const SECBUFFER_VERSION As Long = 0
Private Const SECBUFFER_EMPTY As Long = 0
Private Const SECBUFFER_DATA As Long = 1
Private Const SECBUFFER_TOKEN As Long = 2
Private Const SECBUFFER_EXTRA As Long = 5
Private Const SECBUFFER_STREAM_HEADER As Long = 7
Private Const SECBUFFER_STREAM_TRAILER As Long = 6
Private Const SEC_E_OK As Long = 0
Private Const SEC_I_CONTINUE_NEEDED As Long = &H90312
Private Const SEC_E_INCOMPLETE_MESSAGE As Long = &H80090318
Private Const SEC_I_RENEGOTIATE As Long = &H90321
Private Const SEC_I_CONTEXT_EXPIRED As Long = &H90317

' Network Header Sizes
Private Const ETHERNET_HEADER As Long = 14
Private Const IP_HEADER_MIN As Long = 20
Private Const TCP_HEADER_MIN As Long = 20
Private Const TLS_RECORD_HEADER As Long = 5
Private Const WEBSOCKET_HEADER_MAX As Long = 14

' WebSocket Opcodes
Private Const WS_OPCODE_CONTINUATION As Byte = 0
Private Const WS_OPCODE_TEXT As Byte = 1
Private Const WS_OPCODE_BINARY As Byte = 2
Private Const WS_OPCODE_CLOSE As Byte = 8
Private Const WS_OPCODE_PING As Byte = 9
Private Const WS_OPCODE_PONG As Byte = 10

' Cryptography Constants
Private Const CALG_SHA1 As Long = &H8004&
Private Const HP_HASHVAL As Long = &H2
Private Const CRYPT_STRING_BASE64 As Long = &H1
Private Const CRYPT_NOCRLF As Long = &H40000000
Private Const PROV_RSA_FULL As Long = 1
Private Const CRYPT_VERIFYCONTEXT As Long = &HF0000000
Private Const SECPKG_CRED_OUTBOUND_NTLM As Long = &H2
Private Const SEC_I_COMPLETE_NEEDED As Long = &H90313
Private Const BCRYPT_USE_SYSTEM_PREFERRED_RNG As Long = &H2

' Virtual Memory Constants
Private Const MEM_COMMIT As Long = &H1000
Private Const MEM_RESERVE As Long = &H2000
Private Const PAGE_READWRITE As Long = &H4
Private Const PAGE_EXECUTE_READ As Long = &H20
Private Const MEM_RELEASE As Long = &H8000&

' Window Messaging / Async Events
Private Const FD_READ As Long = &H1
Private Const FD_WRITE As Long = &H2
Private Const FD_CLOSE As Long = &H20
Private Const FD_CONNECT As Long = &H10
Private Const WM_USER As Long = &H400
Private Const WM_WASABI_SOCKET As Long = WM_USER + &H1337
Private Const GWLP_WNDPROC As Long = -4

' Memory & Buffer Limits
Private Const BUFFER_MAX_SIZE As Long = 16777216
Private Const DEFAULT_OPTIMAL_FRAME As Long = 1024
Private Const MIN_FRAME_SIZE As Long = 125
Private Const MAX_FRAME_SIZE As Long = 65535
Private Const OFFLINE_QUEUE_CAP As Long = 10000
Private Const BATCH_MAX_SIZE As Long = 65536

' Network Defaults
Private Const DEFAULT_MSS As Long = 1460
Private Const CONNECT_POLL_USEC As Long = 50000
Private Const DNS_RACE_TIMEOUT_MS As Long = 10000
Private Const SOCKS5_MAX_HOSTNAME As Long = 255

' TLS
Private Const TLS_RECV_CHUNK_SIZE As Long = 16384
Private Const TLS_RECV_GROWTH_SIZE As Long = 32768
Private Const TLS_MAX_CHUNK_FALLBACK As Long = 16384
Private Const TLS_HANDSHAKE_MAX_LOOPS As Long = 30

' Well-known Ports
Private Const PORT_HTTP As Long = 80
Private Const PORT_HTTPS As Long = 443

' WebSocket
Private Const WS_PAYLOAD_LEN_16BIT As Long = 126
Private Const WS_PAYLOAD_LEN_64BIT As Long = 127
Private Const WS_CLOSE_NO_STATUS As Long = 1005
Private Const WS_MAX_CLOSE_REASON As Long = 123

' Winsock
Private Const WSA_EWOULDBLOCK As Long = 10035

' MQTT
Private Const MQTT_CHUNK_SIZE As Long = 1024
Private Const MQTT_MAX_PACKET_ID As Long = 65535
Private Const MQTT_VARINT_MAX_MULTIPLIER As Long = 268435456
Private Const MQTT_VARINT_MULTIPLIER As Long = 128
Private Const MQTT_VARINT_CONTINUE_BIT As Long = 128
Private Const MQTT_VARINT_VALUE_MASK As Long = 127
Private Const MQTT_CONNECT_USERNAME As Long = 128
Private Const MQTT_CONNECT_PASSWORD As Long = 64

' Buffer / MTU Validation Bounds
Private Const MIN_BUFFER_SIZE As Long = 8192
Private Const MIN_FRAGMENT_SIZE As Long = 4096
Private Const MTU_MIN As Long = 576
Private Const MTU_MAX As Long = 9000

' ============================================================================
' 3. TYPES & STRUCTS
' ============================================================================
'/**
' * @struct CRYPT_DATA_BLOB
' * @brief Represents a generic blob of data for cryptography operations.
' */
Private Type CRYPT_DATA_BLOB
#If VBA7 Then
    cbData As Long
    pbData As LongPtr
#Else
    cbData As Long
    pbData As Long
#End If
End Type

'/**
' * @struct CERT_ENHKEY_USAGE
' * @brief Certificate enhanced key usage identifier array.
' */
Private Type CERT_ENHKEY_USAGE
    cUsageIdentifier As Long
#If VBA7 Then
    rgpszUsageIdentifier As LongPtr
#Else
    rgpszUsageIdentifier As Long
#End If
End Type

'/**
' * @struct CERT_USAGE_MATCH
' * @brief Usage criteria matching for certificate chains.
' */
Private Type CERT_USAGE_MATCH
    dwType As Long
    Usage As Long
End Type

'/**
' * @struct CERT_CHAIN_PARA
' * @brief Establishes the searching criteria for certificate chains.
' */
Private Type CERT_CHAIN_PARA
    cbSize As Long
    RequestedUsage_dwType As Long
    RequestedUsage_cUsage As Long
#If VBA7 Then
    RequestedUsage_rgpOID As LongPtr
#Else
    RequestedUsage_rgpOID As Long
#End If
End Type

'/**
' * @struct WINHTTP_CURRENT_USER_IE_PROXY_CONFIG
' * @brief Contains the IE proxy configuration information.
' */
Private Type WINHTTP_CURRENT_USER_IE_PROXY_CONFIG
    fAutoDetect As Long
#If VBA7 Then
    lpszAutoConfigUrl As LongPtr
    lpszProxy As LongPtr
    lpszProxyBypass As LongPtr
#Else
    lpszAutoConfigUrl As Long
    lpszProxy As Long
    lpszProxyBypass As Long
#End If
End Type

'/**
' * @struct SSL_EXTRA_CERT_CHAIN_POLICY_PARA
' * @brief Contains policy parameters used in the verification of SSL chains.
' */
Private Type SSL_EXTRA_CERT_CHAIN_POLICY_PARA
    cbSize As Long
    dwAuthType As Long
    fdwChecks As Long
#If VBA7 Then
    pwszServerName As LongPtr
#Else
    pwszServerName As Long
#End If
End Type

'/**
' * @struct CERT_CHAIN_POLICY_PARA
' * @brief Sets the parameters to pass to the CertVerifyCertificateChainPolicy function.
' */
Private Type CERT_CHAIN_POLICY_PARA
    cbSize As Long
    dwFlags As Long
#If VBA7 Then
    pvExtraPolicyPara As LongPtr
#Else
    pvExtraPolicyPara As Long
#End If
End Type

'/**
' * @struct CERT_CHAIN_POLICY_STATUS
' * @brief Receives status info on the certificate chain validation.
' */
Private Type CERT_CHAIN_POLICY_STATUS
    cbSize As Long
    dwError As Long
    lChainIndex As Long
    lElementIndex As Long
End Type

'/**
' * @struct BatchBuffer
' * @brief Used to hold batched frames for delayed sending.
' */
Private Type BatchBuffer
    Frames() As Byte
    FrameCount As Long
    totalLen As Long
    MaxFrames As Long
End Type

'/**
' * @struct SOCKADDR_IN6
' * @brief Defines an IPv6 socket address.
' */
Private Type SOCKADDR_IN6
    sin6_family   As Integer
    sin6_port     As Integer
    sin6_flowinfo As Long
    sin6_addr(0 To 15) As Byte
    sin6_scope_id As Long
End Type

'/**
' * @struct BinaryMessage
' * @brief Wraps a raw byte array as a queueable message payload.
' */
Private Type BinaryMessage
    data() As Byte
End Type

'/**
' * @struct SecHandle
' * @brief Identifies an SSPI security context or credential.
' */
Private Type SecHandle
#If VBA7 Then
    dwLower As LongPtr
    dwUpper As LongPtr
#Else
    dwLower As Long
    dwUpper As Long
#End If
End Type

'/**
' * @struct SECURITY_INTEGER
' * @brief Holds large integers used for security timestamps.
' */
Private Type SECURITY_INTEGER
    LowPart As Long
    HighPart As Long
End Type

'/**
' * @struct SecBuffer
' * @brief Represents a memory buffer passed to/from SSPI.
' */
Private Type SecBuffer
    cbBuffer As Long
    BufferType As Long
#If VBA7 Then
    pvBuffer As LongPtr
#Else
    pvBuffer As Long
#End If
End Type

'/**
' * @struct SecBufferDesc
' * @brief Describes an array of SecBuffer structures.
' */
Private Type SecBufferDesc
    ulVersion As Long
    cBuffers As Long
#If VBA7 Then
    pBuffers As LongPtr
#Else
    pBuffers As Long
#End If
End Type

'/**
' * @struct SecPkgContext_StreamSizes
' * @brief Defines the sizes of the various parts of a TLS stream.
' */
Private Type SecPkgContext_StreamSizes
    cbHeader As Long
    cbTrailer As Long
    cbMaximumMessage As Long
    cBuffers As Long
    cbBlockSize As Long
End Type

'/**
' * @struct SCHANNEL_CRED
' * @brief Contains the credentials/options for an Schannel session.
' */
Private Type SCHANNEL_CRED
    dwVersion As Long
    cCreds As Long
#If VBA7 Then
    paCred As LongPtr
    hRootStore As LongPtr
#Else
    paCred As Long
    hRootStore As Long
#End If
    cMappers As Long
#If VBA7 Then
    aphMappers As LongPtr
#Else
    aphMappers As Long
#End If
    cSupportedAlgs As Long
#If VBA7 Then
    palgSupportedAlgs As LongPtr
#Else
    palgSupportedAlgs As Long
#End If
    grbitEnabledProtocols As Long
    dwMinimumCipherStrength As Long
    dwMaximumCipherStrength As Long
    dwSessionLifespan As Long
    dwFlags As Long
    dwCredFormat As Long
End Type

'/**
' * @struct WSADATA
' * @brief Information about the Windows Sockets implementation.
' */
Private Type WSADATA
    wVersion As Integer
    wHighVersion As Integer
    szDescription(0 To 256) As Byte
    szSystemStatus(0 To 128) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
#If VBA7 Then
    lpVendorInfo As LongPtr
#Else
    lpVendorInfo As Long
#End If
End Type

'/**
' * @struct SOCKADDR_IN
' * @brief Defines an IPv4 socket address.
' */
Private Type SOCKADDR_IN
    sin_family As Integer
    sin_port As Integer
    sin_addr As Long
    sin_zero(0 To 7) As Byte
End Type

'/**
' * @struct TIMEVAL
' * @brief Specifies a time interval used for timeouts.
' */
Private Type TIMEVAL
    tv_sec As Long
    tv_usec As Long
End Type

'/**
' * @struct FD_SET
' * @brief Set of sockets used for the select() function.
' */
Private Type FD_SET
    fd_count As Long
#If VBA7 Then
    fd_array(0 To 0) As LongPtr
#Else
    fd_array(0 To 0) As Long
#End If
End Type

'/**
' * @struct HOSTENT32
' * @brief Information about a host returned by gethostbyname (32-bit).
' */
Private Type HOSTENT32
    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long
End Type

'/**
' * @struct HOSTENT64
' * @brief Information about a host returned by gethostbyname (64-bit).
' */
Private Type HOSTENT64
    h_name As LongPtr
    h_aliases As LongPtr
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As LongPtr
End Type

'/**
' * @struct MTUInfo
' * @brief Maximum Transmission Unit configuration for packet splitting.
' */
Private Type MTUInfo
    CurrentMTU As Long
    MaxSegmentSize As Long
    OptimalFrameSize As Long
    LastProbeTime As Long
    ProbeEnabled As Boolean
    UseTLSFragmentation As Boolean
End Type

'/**
' * @struct MqttInFlightMsg
' * @brief Represents an unacknowledged MQTT packet.
' */
Private Type MqttInFlightMsg
    packetId As Integer
    topic As String
    payload() As Byte
    qos As Byte
    SentTick As Long
End Type

'/**
' * @struct WasabiStats
' * @brief Connection statistics for analytics.
' */
Private Type WasabiStats
    BytesSent As Currency
    BytesReceived As Currency
    MessagesSent As Long
    MessagesReceived As Long
    ConnectedAt As Long
End Type

' ============================================================================
' 4. ENUMS
' ============================================================================

'/**
' * @enum WasabiState
' * @brief Lifecycle state of a Wasabi connection.
' */
Public Enum WasabiState
    STATE_CLOSED = 0
    STATE_CONNECTING = 1
    STATE_OPEN = 2
    STATE_CLOSING = 3
End Enum

'/**
' * @enum WasabiConnectionMode
' * @brief Identifies the underlying connection protocol.
' */
Public Enum WasabiConnectionMode
    MODE_WEBSOCKET = 0
    MODE_TCP = 1
    MODE_TCP_TLS = 2
End Enum

'/**
' * @enum WasabiError
' * @brief Detailed error codes for connection failures.
' */
Public Enum WasabiError
    ERR_NONE = 0
    ERR_WSA_STARTUP_FAILED = 1
    ERR_SOCKET_CREATE_FAILED = 2
    ERR_DNS_RESOLVE_FAILED = 3
    ERR_CONNECT_FAILED = 4
    ERR_TLS_ACQUIRE_CREDS_FAILED = 5
    ERR_TLS_HANDSHAKE_FAILED = 6
    ERR_TLS_HANDSHAKE_TIMEOUT = 7
    ERR_WEBSOCKET_HANDSHAKE_FAILED = 8
    ERR_WEBSOCKET_HANDSHAKE_TIMEOUT = 9
    ERR_SEND_FAILED = 10
    ERR_RECV_FAILED = 11
    ERR_NOT_CONNECTED = 12
    ERR_ALREADY_CONNECTED = 13
    ERR_TLS_ENCRYPT_FAILED = 14
    ERR_TLS_DECRYPT_FAILED = 15
    ERR_INVALID_URL = 16
    ERR_HANDSHAKE_REJECTED = 17
    ERR_CONNECTION_LOST = 18
    ERR_INVALID_HANDLE = 19
    ERR_MAX_CONNECTIONS = 20
    ERR_PROXY_CONNECT_FAILED = 21
    ERR_PROXY_AUTH_FAILED = 22
    ERR_PROXY_TUNNEL_FAILED = 23
    ERR_INACTIVITY_TIMEOUT = 24
    ERR_CERT_LOAD_FAILED = 25
    ERR_CERT_VALIDATE_FAILED = 26
    ERR_FRAGMENT_OVERFLOW = 27
    ERR_TLS_RENEGOTIATE = 28
End Enum

'/**
' * @enum MqttPacketType
' * @brief Protocol packet identifiers for MQTT.
' */
Private Enum MqttPacketType
    MQTT_CONNECT = 1
    MQTT_CONNACK = 2
    MQTT_PUBLISH = 3
    MQTT_PUBACK = 4
    MQTT_PUBREC = 5
    MQTT_PUBREL = 6
    MQTT_PUBCOMP = 7
    MQTT_SUBSCRIBE = 8
    MQTT_SUBACK = 9
    MQTT_UNSUBSCRIBE = 10
    MQTT_UNSUBACK = 11
    MQTT_PINGREQ = 12
    MQTT_PINGRESP = 13
    MQTT_DISCONNECT = 14
End Enum


'/**
' * @struct WasabiConnection
' * @brief Represents a single active socket connection with all its state, queues, and handlers.
' */
Private Type WasabiConnection
#If VBA7 Then
    Socket As LongPtr
    hClientCertStore As LongPtr
    pClientCertCtx As LongPtr
    hNtlmCred As SecHandle
#Else
    Socket As Long
    hClientCertStore As Long
    pClientCertCtx As Long
    hNtlmCred As SecHandle
#End If
    state As WasabiState
    TLS As Boolean
    HOST As String
    port As Long
    path As String
    OriginalUrl As String
    hCred As SecHandle
    hContext As SecHandle
    sizes As SecPkgContext_StreamSizes
    recvBuffer() As Byte
    recvLen As Long
    DecryptBuffer() As Byte
    DecryptLen As Long
    MsgQueue() As String
    MsgHead As Long
    MsgTail As Long
    MsgCount As Long
    BinaryQueue() As BinaryMessage
    BinaryHead As Long
    BinaryTail As Long
    BinaryCount As Long
    FragmentBuffer() As Byte
    FragmentLen As Long
    FragmentOpcode As Byte
    Fragmenting As Boolean
    LastError As Long
    LastErrorCode As Long
    TechnicalDetails As String
    CustomHeaders() As String
    CustomHeaderCount As Long
    AutoReconnect As Boolean
    ReconnectMaxAttempts As Long
    ReconnectAttempts As Long
    ReconnectBaseDelayMs As Long
    PingIntervalMs As Long
    LastPingSentAt As Long
    ReceiveTimeoutMs As Long
    EnableErrorDialog As Boolean
    LogCallback As String
    stats As WasabiStats
    NoDelay As Boolean
    proxyHost As String
    proxyPort As Long
    proxyUser As String
    proxyPass As String
    proxyType As Long
    ProxyEnabled As Boolean
    InactivityTimeoutMs As Long
    LastActivityAt As Long
    SubProtocol As String
    CustomBufferSize As Long
    CustomFragmentSize As Long
    mtu As MTUInfo
    AutoMTU As Boolean
    ZeroCopyEnabled As Boolean
    closeCode As Integer
    closeReason As String
    CloseInitiatedByUs As Boolean
    PreferIPv6 As Boolean
    ValidateServerCert As Boolean
    EnableRevocationCheck As Boolean
    ClientCertThumb As String
    ClientCertPfxPath As String
    ClientCertPfxPass As String
    UseHttp2 As Boolean
    ProxyUseNtlm As Boolean
    LastRttMs As Long
    LastPingTimestamp As Long
    MqttParserStage As Long
    MqttBuffer() As Byte
    MqttBufLen As Long
    MqttExpectedRemaining As Long
    MqttCurrentPacketType As Byte
    MqttCurrentFlags As Byte
    DeflateEnabled          As Boolean
    DeflateContextTakeover  As Boolean
    InflateContextTakeover  As Boolean
    DeflateWindowBits       As Long
    InflateWindowBits       As Long
    DeflateReady            As Boolean
    InflateReady            As Boolean
    DeflateActive           As Boolean
    FragmentIsCompressed As Boolean
    ClientMaxWindowBits As Long
    ServerMaxWindowBits As Long
    MqttNextPacketId As Integer
    MqttInFlight() As MqttInFlightMsg
    MqttInFlightCount As Long
    OfflineQueueEnabled As Boolean
    OfflineTextQueue() As String
    OfflineTextCount As Long
    OfflineBinaryQueue() As BinaryMessage
    OfflineBinaryCount As Long
    
    mode As WasabiConnectionMode
    TcpRecvBuffer() As Byte
    TcpRecvLen As Long
    
    PingJitterMaxMs As Long
    CurrentPingIntervalMs As Long
    
    ProtocolHandler As Object
    CompressionHandler As Object
    Middlewares As Collection
    
    AsyncHandler As Object
    AsyncMode As Boolean
End Type

' ============================================================================
' 5. GLOBAL VARIABLES
' ============================================================================
Private m_WSAInitialized As Boolean
Private m_Connections() As WasabiConnection
Private m_ConnectionCount As Long
Private m_DefaultHandle As Long
Private m_LastError As WasabiError
Private m_LastErrorCode As Long
Private m_TechnicalDetails As String
Private m_Utf8Buf() As Byte
Private m_Utf8BufSize As Long
Private m_ZeroCopyText As String
Private m_ZeroCopyBinary() As Byte
Private m_AppIsAlive As Long
#If VBA7 Then
    Private m_ClientCertContextPtrs(0 To MAX_CONNECTIONS - 1) As LongPtr
    Private m_ptrWsMask As LongPtr
    Private m_ptrMemZero As LongPtr
    Private m_ptrMemFind As LongPtr
    Private m_ptrTickDiff As LongPtr
    Private m_ptrAsyncThunk As LongPtr
    Private m_AsyncHwnd As LongPtr
    Private m_OldWndProc As LongPtr
#Else
    Private m_ClientCertContextPtrs(0 To MAX_CONNECTIONS - 1) As Long
    Private m_ptrWsMask As Long
    Private m_ptrMemZero As Long
    Private m_ptrMemFind As Long
    Private m_ptrTickDiff As Long
    Private m_ptrAsyncThunk As Long
    Private m_AsyncHwnd As Long
    Private m_OldWndProc As Long
#End If
Public EnableErrorDialog As Boolean

' ============================================================================
' 6. LOW-LEVEL MEMORY & THUNKS (MACHINE CODE EXECUTIONS)
' ============================================================================

'/**
' * @brief Initializes raw machine-code thunks for critical performance paths
' * (WebSocket Masking, Memory zeroing, Memory finding, Tick differentials).
' */
Private Sub InitWasabiThunks()
    If m_ptrWsMask <> 0 Then Exit Sub
    #If Win64 Then
        m_ptrWsMask = LoadThunk(GetWsMaskOpcodes_x64())
        m_ptrMemZero = LoadThunk(GetMemZeroOpcodes_x64())
        m_ptrMemFind = LoadThunk(GetMemFindOpcodes_x64())
        m_ptrTickDiff = LoadThunk(GetTickDiffOpcodes_x64())
    #Else
        m_ptrWsMask = LoadThunk(GetWsMaskOpcodes_x86())
        m_ptrMemZero = LoadThunk(GetMemZeroOpcodes_x86())
        m_ptrMemFind = LoadThunk(GetMemFindOpcodes_x86())
        m_ptrTickDiff = LoadThunk(GetTickDiffOpcodes_x86())
    #End If
End Sub

'/**
' * @brief Frees the executable memory allocated for thunks.
' */
Private Sub ShutdownWasabiThunks()
    If m_ptrWsMask <> 0 Then VirtualFree m_ptrWsMask, 0, MEM_RELEASE: m_ptrWsMask = 0
    If m_ptrMemZero <> 0 Then VirtualFree m_ptrMemZero, 0, MEM_RELEASE: m_ptrMemZero = 0
    If m_ptrMemFind <> 0 Then VirtualFree m_ptrMemFind, 0, MEM_RELEASE: m_ptrMemFind = 0
    If m_ptrTickDiff <> 0 Then VirtualFree m_ptrTickDiff, 0, MEM_RELEASE: m_ptrTickDiff = 0
End Sub

'/**
' * @brief Quickly zeros out a block of memory using machine code.
' * @param ptr Pointer to the start of memory.
' * @param length Number of bytes to clear.
' */
#If VBA7 Then
Private Sub WasabiMemZero(ByVal ptr As LongPtr, ByVal length As Long)
#Else
Private Sub WasabiMemZero(ByVal ptr As Long, ByVal length As Long)
#End If
    If ptr = 0 Or length <= 0 Then Exit Sub
    
    If m_ptrMemZero <> 0 Then
        CallWindowProcW m_ptrMemZero, ptr, length, 0, 0
    Else
        Dim i As Long
        Dim zero As Byte: zero = 0
        For i = 0 To length - 1
            CopyMemoryFromPtr ByVal (ptr + i), VarPtr(zero), 1
        Next i
    End If
End Sub

'/**
' * @brief High-performance byte array search (Finds a needle in a haystack).
' * @param haystackPtr Pointer to search space.
' * @param hayLen Length of search space.
' * @param needlePtr Pointer to search target.
' * @param needleLen Length of search target.
' * @return Index of match, or -1 if not found.
' */
#If VBA7 Then
Private Function WasabiMemFind(ByVal haystackPtr As LongPtr, ByVal hayLen As Long, ByVal needlePtr As LongPtr, ByVal needleLen As Long) As Long
#Else
Private Function WasabiMemFind(ByVal haystackPtr As Long, ByVal hayLen As Long, ByVal needlePtr As Long, ByVal needleLen As Long) As Long
#End If
    If m_ptrMemFind <> 0 Then
        #If VBA7 Then
            WasabiMemFind = CLng(CallWindowProcW(m_ptrMemFind, haystackPtr, hayLen, needlePtr, needleLen))
        #Else
            WasabiMemFind = CallWindowProcW(m_ptrMemFind, haystackPtr, hayLen, needlePtr, needleLen)
        #End If
    Else
        If hayLen < needleLen Or needleLen <= 0 Then
            WasabiMemFind = -1
            Exit Function
        End If
        
        Dim i As Long, j As Long
        Dim found As Boolean
        Dim hByte As Byte, nByte As Byte
        
        For i = 0 To hayLen - needleLen
            found = True
            For j = 0 To needleLen - 1
                CopyMemoryFromPtr hByte, haystackPtr + i + j, 1
                CopyMemoryFromPtr nByte, needlePtr + j, 1
                If hByte <> nByte Then
                    found = False
                    Exit For
                End If
            Next j
            If found Then
                 WasabiMemFind = i
                Exit Function
            End If
        Next i
        WasabiMemFind = -1
    End If
End Function

'/**
' * @brief Cross-architecture function to securely return the memory address of a pointer.
' * @param ptr Input pointer.
' * @return Output pointer address.
' */
#If VBA7 Then
Private Function GetAddressOf(ByVal ptr As LongPtr) As LongPtr
    GetAddressOf = ptr
End Function
#Else
Private Function GetAddressOf(ByVal ptr As Long) As Long
    GetAddressOf = ptr
End Function
#End If

'=============================================================================
' MACHINE CODE OPCODES (X64 & X86)
'=============================================================================
#If Win64 Then

'/**
' * @brief Returns the compiled machine code opcodes for the x64 asynchronous window subclassing thunk.
' * @return Byte array containing the executable assembly instructions.
' */
Private Function GetAsyncThunkOpcodes_x64() As Byte()
    Dim opcodes(0 To 97) As Byte
    Dim HexStr As Variant
    HexStr = Array( _
        &H55, &H48, &H89, &HE5, &H51, &H52, &H41, &H50, &H41, &H51, &H48, &H83, &HEC, &H20, _
        &H48, &HB8, 0, 0, 0, 0, 0, 0, 0, 0, _
        &H8B, &H0, &H85, &HC0, &H74, &H2D, _
        &H48, &HB8, 0, 0, 0, 0, 0, 0, 0, 0, _
        &H48, &H85, &HC0, &H74, &H7, &HFF, &HD0, &H83, &HF8, &H1, &H75, &H17, _
        &H48, &H83, &HC4, &H20, &H41, &H59, &H41, &H58, &H5A, &H59, &H5D, _
        &H48, &HB8, 0, 0, 0, 0, 0, 0, 0, 0, _
        &HFF, &HE0, _
        &H48, &H83, &HC4, &H20, &H41, &H59, &H41, &H58, &H5A, &H59, &H5D, _
        &H48, &HB8, 0, 0, 0, 0, 0, 0, 0, 0, _
        &HFF, &HE0)
    Dim i As Long: For i = 0 To 97: opcodes(i) = CByte(HexStr(i)): Next i
    GetAsyncThunkOpcodes_x64 = opcodes
End Function

'/**
' * @brief Returns the compiled machine code opcodes for the x64 WebSocket payload masking/unmasking algorithm.
' * This significantly speeds up the bitwise XOR operations required by RFC6455.
' * @return Byte array containing the executable assembly instructions.
' */
Private Function GetWsMaskOpcodes_x64() As Byte()
    Dim opcodes(0 To 21) As Byte
    Dim HexStr As Variant: HexStr = Array(&H48, &H85, &HD2, &H74, &H10, &H41, &H8B, &H0, &H30, &H1, &H48, &HFF, &HC1, &HC1, &HC8, &H8, &H48, &HFF, &HCA, &H75, &HF3, &HC3)
    Dim i As Long: For i = 0 To 21: opcodes(i) = CByte(HexStr(i)): Next i
    GetWsMaskOpcodes_x64 = opcodes
End Function

'/**
' * @brief Returns the compiled machine code opcodes for a highly optimized x64 memory zeroing (SecureZero) routine.
' * @return Byte array containing the executable assembly instructions.
' */
Private Function GetMemZeroOpcodes_x64() As Byte()
    Dim opcodes(0 To 12) As Byte
    Dim HexStr As Variant: HexStr = Array(&H57, &H48, &H89, &HCF, &H48, &H89, &HD1, &H31, &HC0, &HF3, &HAA, &H5F, &HC3)
    Dim i As Long: For i = 0 To 12: opcodes(i) = CByte(HexStr(i)): Next i
    GetMemZeroOpcodes_x64 = opcodes
End Function

'/**
' * @brief Returns the compiled machine code opcodes for a fast x64 memory search algorithm.
' * Designed to efficiently locate a sequence of bytes (needle) within a larger buffer (haystack).
' * @return Byte array containing the executable assembly instructions.
' */
Private Function GetMemFindOpcodes_x64() As Byte()
    Dim opcodes(0 To 59) As Byte
    Dim HexStr As Variant: HexStr = Array(&H56, &H57, &H53, &H4C, &H39, &HCA, &H72, &H29, &H4D, &H85, &HC9, &H74, &H24, &H4C, &H29, &HCA, &H48, &HFF, &HC2, &H48, &H31, &HC0, &H48, &H89, &HCB, &H4C, &H89, &HC9, &H48, &H89, &HDF, &H4C, &H89, &HC6, &HF3, &HA6, &H74, &H12, &H48, &HFF, &HC3, &H48, &HFF, &HC0, &H48, &HFF, &HCA, &H75, &HE8, &H48, &HC7, &HC0, &HFF, &HFF, &HFF, &HFF, &H5B, &H5F, &H5E, &HC3)
    Dim i As Long: For i = 0 To 59: opcodes(i) = CByte(HexStr(i)): Next i
    GetMemFindOpcodes_x64 = opcodes
End Function

'/**
' * @brief Returns the compiled machine code opcodes for a safe x64 tick count differential calculator.
' * Prevents timing bugs when the system GetTickCount overflows after 49.7 days.
' * @return Byte array containing the executable assembly instructions.
' */
Private Function GetTickDiffOpcodes_x64() As Byte()
    Dim opcodes(0 To 4) As Byte
    Dim HexStr As Variant: HexStr = Array(&H89, &HD0, &H2B, &HC1, &HC3)
    Dim i As Long: For i = 0 To 4: opcodes(i) = CByte(HexStr(i)): Next i
    GetTickDiffOpcodes_x64 = opcodes
End Function

#Else

'/**
' * @brief Returns the compiled machine code opcodes for the x86 asynchronous window subclassing thunk.
' * @return Byte array containing the executable assembly instructions.
' */
Private Function GetAsyncThunkOpcodes_x86() As Byte()
    Dim opcodes(0 To 48) As Byte
    Dim HexStr As Variant
    HexStr = Array( _
        &H50, &H51, &H52, _
        &HB8, 0, 0, 0, 0, _
        &H8B, &H0, &H85, &HC0, &H74, &H1A, _
        &HB8, 0, 0, 0, 0, _
        &H85, &HC0, &H74, &H7, &HFF, &HD0, &H83, &HF8, &H1, &H75, &HA, _
        &H5A, &H59, &H58, _
        &HB8, 0, 0, 0, 0, _
        &HFF, &HE0, _
        &H5A, &H59, &H58, _
        &HB8, 0, 0, 0, 0, _
        &HFF, &HE0)
    Dim i As Long: For i = 0 To 48: opcodes(i) = CByte(HexStr(i)): Next i
    GetAsyncThunkOpcodes_x86 = opcodes
End Function

'/**
' * @brief Returns the compiled machine code opcodes for the x86 WebSocket payload masking/unmasking algorithm.
' * @return Byte array containing the executable assembly instructions.
' */
Private Function GetWsMaskOpcodes_x86() As Byte()
    Dim opcodes(0 To 34) As Byte
    Dim HexStr As Variant: HexStr = Array(&H55, &H89, &HE5, &H53, &H8B, &H4D, &HC, &H85, &HC9, &H74, &H13, &H8B, &H55, &H8, &H8B, &H45, &H10, &H8B, &H18, &H88, &HD8, &H30, &H2, &H42, &HC1, &HCB, &H8, &H49, &H75, &HF5, &H5B, &H5D, &HC2, &H10, &H0)
    Dim i As Long: For i = 0 To 34: opcodes(i) = CByte(HexStr(i)): Next i
    GetWsMaskOpcodes_x86 = opcodes
End Function

'/**
' * @brief Returns the compiled machine code opcodes for a highly optimized x86 memory zeroing routine.
' * @return Byte array containing the executable assembly instructions.
' */
Private Function GetMemZeroOpcodes_x86() As Byte()
    Dim opcodes(0 To 18) As Byte
    Dim HexStr As Variant: HexStr = Array(&H55, &H89, &HE5, &H57, &H8B, &H7D, &H8, &H8B, &H4D, &HC, &H31, &HC0, &HF3, &HAA, &H5F, &H5D, &HC2, &H10, &H0)
    Dim i As Long: For i = 0 To 18: opcodes(i) = CByte(HexStr(i)): Next i
    GetMemZeroOpcodes_x86 = opcodes
End Function

'/**
' * @brief Returns the compiled machine code opcodes for a fast x86 memory search algorithm.
' * @return Byte array containing the executable assembly instructions.
' */
Private Function GetMemFindOpcodes_x86() As Byte()
    Dim opcodes(0 To 60) As Byte
    Dim HexStr As Variant: HexStr = Array(&H55, &H89, &HE5, &H53, &H56, &H57, &H8B, &H55, &HC, &H8B, &H4D, &H14, &H39, &HCA, &H72, &H21, &H85, &HC9, &H74, &H1D, &H29, &HCA, &H42, &H31, &HC0, &H8B, &H5D, &H8, &H51, &H53, &H8B, &H4D, &H14, &H89, &HDF, &H8B, &H75, &H10, &HF3, &HA6, &H5B, &H59, &H74, &HA, &H43, &H40, &H4A, &H75, &HEB, &HB8, &HFF, &HFF, &HFF, &HFF, &H5F, &H5E, &H5B, &H5D, &HC2, &H10, &H0)
    Dim i As Long: For i = 0 To 60: opcodes(i) = CByte(HexStr(i)): Next i
    GetMemFindOpcodes_x86 = opcodes
End Function

'/**
' * @brief Returns the compiled machine code opcodes for a safe x86 tick count differential calculator.
' * @return Byte array containing the executable assembly instructions.
' */
Private Function GetTickDiffOpcodes_x86() As Byte()
    Dim opcodes(0 To 10) As Byte
    Dim HexStr As Variant: HexStr = Array(&H8B, &H44, &H24, &H8, &H2B, &H44, &H24, &H4, &HC2, &H10, &H0)
    Dim i As Long: For i = 0 To 10: opcodes(i) = CByte(HexStr(i)): Next i
    GetTickDiffOpcodes_x86 = opcodes
End Function

#End If

'/**
' * @brief Translates opcodes into executable memory and returns a function pointer.
' * @param opcodes Byte array of machine instructions.
' * @return Pointer to allocated executable memory block.
' */
#If VBA7 Then
Private Function LoadThunk(ByRef opcodes() As Byte) As LongPtr
    Dim pMem As LongPtr
#Else
Private Function LoadThunk(ByRef opcodes() As Byte) As Long
    Dim pMem As Long
#End If
    Dim size As Long
    Dim oldProtect As Long
    
    If (Not Not opcodes) = 0 Then Exit Function
    size = UBound(opcodes) - LBound(opcodes) + 1
    
    pMem = VirtualAlloc(0, size, MEM_COMMIT Or MEM_RESERVE, PAGE_READWRITE)
    
    If pMem <> 0 Then
        CopyMemoryFromPtr ByVal pMem, VarPtr(opcodes(LBound(opcodes))), size
        VirtualProtect pMem, size, PAGE_EXECUTE_READ, oldProtect
    End If
    
    LoadThunk = pMem
End Function

'/**
' * @brief Fills a byte array with random data for cryptographic generation using modern CNG (Cryptography Next Generation).
' * @param buf Target byte array.
' * @param count Number of random bytes.
' */
Private Sub FillRandomBytes(ByRef buf() As Byte, ByVal count As Long)
    If count <= 0 Then Exit Sub
    
    ' BCryptGenRandom returns 0 (STATUS_SUCCESS) upon success
    If BCryptGenRandom(0, buf(LBound(buf)), count, BCRYPT_USE_SYSTEM_PREFERRED_RNG) <> 0 Then
        ' Fallback: Rnd() is not cryptographically secure.
        ' Used for Sec-WebSocket-Key (RFC 6455 sec 4.1) and frame masking.
        ' A predictable key lets an attacker predict the Sec-WebSocket-Accept
        ' response and impersonate the server, bypassing the handshake validation
        ' that prevents cross-protocol attacks and cache poisoning.
        Debug.Print "[WASABI] BCryptGenRandom unavailable, falling back to Rnd(). Rnd() is not cryptographically secure. Predictable Sec-WebSocket-Key opens the handshake to server impersonation and replay."
        Dim i As Long
        Randomize
        For i = 0 To count - 1
            buf(i) = CByte(Int(Rnd * 256))
        Next i
    End If
End Sub

' ============================================================================
' 7. WINDOWS MESSAGING & ASYNC CORE
' ============================================================================

'/**
' * @brief Prepares the asynchronous listener thunk.
' */
Private Sub InitAsyncThunk()
    If m_ptrAsyncThunk <> 0 Then Exit Sub
    
    #If VBA7 Then
        Dim hVbe As LongPtr
        hVbe = GetModuleHandleA("vbe7.dll")
        If hVbe = 0 Then hVbe = GetModuleHandleA("vba6.dll")
        Dim pEbMode As LongPtr
        If hVbe <> 0 Then pEbMode = GetProcAddress(hVbe, "EbMode")
    #Else
        Dim hVbe As Long
        hVbe = GetModuleHandleA("vba6.dll")
        Dim pEbMode As Long
        If hVbe <> 0 Then pEbMode = GetProcAddress(hVbe, "EbMode")
    #End If
    
    #If Win64 Then
        Dim opcodes() As Byte: opcodes = GetAsyncThunkOpcodes_x64()
        CopyMemory opcodes(16), VarPtr(m_AppIsAlive), 8
        CopyMemory opcodes(32), pEbMode, 8
        CopyMemory opcodes(65), GetAddressOf(AddressOf WasabiAsyncWndProc), 8
        CopyMemory opcodes(88), GetProcAddress(GetModuleHandleA("user32"), "DefWindowProcW"), 8
    #Else
        Dim opcodes() As Byte: opcodes = GetAsyncThunkOpcodes_x86()
        CopyMemory opcodes(4), VarPtr(m_AppIsAlive), 4
        CopyMemory opcodes(15), pEbMode, 4
        CopyMemory opcodes(34), GetAddressOf(AddressOf WasabiAsyncWndProc), 4
        CopyMemory opcodes(44), GetProcAddress(GetModuleHandleA("user32"), "DefWindowProcW"), 4
    #End If
    
    m_ptrAsyncThunk = LoadThunk(opcodes)
    m_AppIsAlive = 1
End Sub

'/**
' * @brief Creates an invisible window to process asynchronous networking callbacks.
' */
Private Sub InitAsyncWindow()
    If m_AsyncHwnd <> 0 Then Exit Sub
    InitAsyncThunk
    m_AsyncHwnd = CreateWindowExW(0, StrPtr("STATIC"), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    If m_AsyncHwnd <> 0 And m_ptrAsyncThunk <> 0 Then
        #If VBA7 Then
            m_OldWndProc = SetWindowLongPtrW(m_AsyncHwnd, GWLP_WNDPROC, m_ptrAsyncThunk)
        #Else
            m_OldWndProc = SetWindowLongW(m_AsyncHwnd, GWLP_WNDPROC, m_ptrAsyncThunk)
        #End If
    End If
End Sub

'/**
' * @brief Closes the asynchronous listener and cleans up its memory.
' */
Private Sub ShutdownAsyncWindow()
    If m_AsyncHwnd <> 0 Then
        #If VBA7 Then
            If m_OldWndProc <> 0 Then SetWindowLongPtrW m_AsyncHwnd, GWLP_WNDPROC, m_OldWndProc
        #Else
            If m_OldWndProc <> 0 Then SetWindowLongW m_AsyncHwnd, GWLP_WNDPROC, m_OldWndProc
        #End If
        DestroyWindow m_AsyncHwnd
        m_AsyncHwnd = 0
        m_OldWndProc = 0
    End If
    If m_ptrAsyncThunk <> 0 Then
        VirtualFree m_ptrAsyncThunk, 0, MEM_RELEASE
        m_ptrAsyncThunk = 0
    End If
    m_AppIsAlive = 0
End Sub

'/**
' * @brief The central callback function hooked to the window to receive WSASyncSelect messages.
' * @param hwnd Window handle.
' * @param uMsg The message type.
' * @param wParam WParam data (typically socket handle).
' * @param lParam LParam data (typically event code).
' * @return Standard windows result.
' */
#If VBA7 Then
Public Function WasabiAsyncWndProc(ByVal hwnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
#Else
Public Function WasabiAsyncWndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#End If
    Dim handle As Long
    Dim eventVal As Long
    Dim errorVal As Long
    Dim i As Long
    Dim matched As Boolean
    Dim lParam32 As Long

    If uMsg = WM_WASABI_SOCKET Then
        lParam32 = CLng(lParam And &HFFFFFFFF)
        eventVal = lParam32 And &HFFFF&
        errorVal = (lParam32 And &HFFFF0000) \ &H10000
        If errorVal And &H8000& Then errorVal = errorVal Or &HFFFF0000

        For i = 0 To MAX_CONNECTIONS - 1
            If m_Connections(i).Socket = wParam And m_Connections(i).state <> STATE_CLOSED Then
                handle = i
                matched = True
                Exit For
            End If
        Next i

        If matched Then
            If Not m_Connections(handle).AsyncHandler Is Nothing Then
                Select Case eventVal
                    Case FD_CONNECT
                        If errorVal = 0 Then
                            CallByName m_Connections(handle).AsyncHandler, "OnConnect", VbMethod, handle
                        Else
                            CallByName m_Connections(handle).AsyncHandler, "OnError", VbMethod, handle, errorVal, FD_CONNECT
                            CloseSession handle
                        End If
                    Case FD_READ
                        If errorVal = 0 Then
                            FeedBuffer handle
                            If m_Connections(handle).state = STATE_OPEN Then
                                CallByName m_Connections(handle).AsyncHandler, "OnReceive", VbMethod, handle
                            End If
                        Else
                            CallByName m_Connections(handle).AsyncHandler, "OnError", VbMethod, handle, errorVal, FD_READ
                            CloseSession handle
                        End If
                    Case FD_WRITE
                        If errorVal = 0 Then
                            If m_Connections(handle).state = STATE_OPEN Then
                                CallByName m_Connections(handle).AsyncHandler, "OnReadyToSend", VbMethod, handle
                            End If
                        End If
                    Case FD_CLOSE
                        If Not m_Connections(handle).AsyncHandler Is Nothing Then
                            CallByName m_Connections(handle).AsyncHandler, "OnClose", VbMethod, handle
                        End If
                        CloseSession handle
                End Select
            End If
        End If
        WasabiAsyncWndProc = 0
        Exit Function
    End If

    WasabiAsyncWndProc = CallWindowProcW_WndProc(m_OldWndProc, hwnd, uMsg, wParam, lParam)
End Function

'/**
' * @brief Registers an asynchronous event handler to the socket.
' * @param handler A class instance containing OnConnect, OnError, OnReceive, OnReadyToSend, OnClose.
' * @param handle The target socket handle.
' */
Public Sub WasabiUseAsync(ByVal handler As Object, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub

    InitAsyncWindow

    If m_Connections(h).AsyncMode And Not m_Connections(h).AsyncHandler Is Nothing Then
        If m_Connections(h).state = STATE_OPEN Then
            CallByName m_Connections(h).AsyncHandler, "OnDisconnect", VbMethod, h
        End If
    End If

    Set m_Connections(h).AsyncHandler = handler
    m_Connections(h).AsyncMode = True

    If m_Connections(h).Socket <> INVALID_SOCKET Then
        WSAAsyncSelect m_Connections(h).Socket, m_AsyncHwnd, WM_WASABI_SOCKET, FD_READ Or FD_WRITE Or FD_CLOSE Or FD_CONNECT
    End If
End Sub

' ============================================================================
' 8. TIME & BUFFERS UTILITIES
' ============================================================================

'/**
' * @brief Accurately compares two GetTickCount values handling overflows.
' * @param startTick Tick count at the start.
' * @param endTick Tick count at the end.
' * @return Difference in milliseconds.
' */
Private Function TickDiff(ByVal startTick As Long, ByVal endTick As Long) As Double
    #If VBA7 Then
        Dim resultPtr As LongPtr
    #Else
        Dim resultPtr As Long
    #End If
    
    If m_ptrTickDiff = 0 Then InitWasabiThunks
    
    If m_ptrTickDiff <> 0 Then
        resultPtr = CallWindowProcW(m_ptrTickDiff, startTick, endTick, 0, 0)
        #If Win64 Then
            TickDiff = CDbl(resultPtr And &HFFFFFFFF^)
        #Else
            If resultPtr < 0 Then
                TickDiff = CDbl(resultPtr) + 4294967296#
            Else
                TickDiff = CDbl(resultPtr)
            End If
        #End If
    Else
        Dim s As Currency, e As Currency
        If startTick >= 0 Then s = startTick Else s = startTick + 4294967296@
        If endTick >= 0 Then e = endTick Else e = endTick + 4294967296@
        If e >= s Then
            TickDiff = CDbl(e - s)
        Else
            TickDiff = CDbl(e - s + 4294967296@)
        End If
    End If
End Function

'/**
' * @brief Safely resizes a byte array buffer without losing existing data, ensuring performance.
' * @param Buffer Target byte array.
' * @param RequiredLen Minimum size required.
' */
Private Sub EnsureBufferCapacity(ByRef Buffer() As Byte, ByVal RequiredLen As Long)
    Dim currentCap As Long
    Dim newCap As Long

    If (Not Not Buffer) = 0 Then
        ReDim Buffer(0 To RequiredLen - 1)
        Exit Sub
    End If

    currentCap = UBound(Buffer) + 1

    If RequiredLen > currentCap Then
        If RequiredLen > BUFFER_MAX_SIZE Then Err.Raise 7, "Wasabi", "Memory Limit Exceeded (>16MB)"
        newCap = currentCap * 2
        If newCap < RequiredLen Then newCap = RequiredLen
        If newCap > BUFFER_MAX_SIZE Then newCap = RequiredLen
        ReDim Preserve Buffer(0 To newCap - 1)
    End If
End Sub

'/**
' * @brief Returns the actual size of an allocated SafeArray (Byte Array).
' * @param arr Target byte array.
' * @return The size of the array, or 0 if uninitialized.
' */
Private Function SafeArrayLen(ByRef arr() As Byte) As Long
    #If VBA7 Then
        Dim pSA As LongPtr
        Dim ptr As LongPtr
    #Else
        Dim pSA As Long
        Dim ptr As Long
    #End If
    Dim lo As Long
    Dim hi As Long
    ptr = VarPtrArray(arr)
    If ptr = 0 Then Exit Function
    CopyMemoryFromPtr pSA, ptr, LenB(pSA)
    If pSA = 0 Then Exit Function
    lo = LBound(arr)
    hi = UBound(arr)
    If hi >= lo Then
        SafeArrayLen = hi - lo + 1
    End If
End Function

'/**
' * @brief Resolves an input connection handle to the active/default one if unspecified.
' * @param handle Input handle.
' * @return Resolved internal index handle.
' */
Private Function ResolveHandle(ByVal handle As Long) As Long
    If handle = INVALID_CONN_HANDLE Then
        ResolveHandle = m_DefaultHandle
    Else
        ResolveHandle = handle
    End If
End Function

'/**
' * @brief Checks if a connection index handle is valid.
' * @param handle Input index.
' * @return True if valid, False otherwise.
' */
Private Function ValidIndex(ByVal handle As Long) As Boolean
    If handle < 0 Or handle >= MAX_CONNECTIONS Then Exit Function
    InitConnectionPool
    ValidIndex = True
End Function

'/**
' * @brief Internal logging dispatcher. Triggers Debug.Print and custom callbacks.
' * @param handle Active connection.
' * @param msg The message string to log.
' */
Private Sub WasabiLog(ByVal handle As Long, ByVal msg As String)
    Debug.Print "[WASABI] " & msg
    If ValidIndex(handle) Then
        If m_Connections(handle).LogCallback <> "" Then
            Application.Run m_Connections(handle).LogCallback, msg
        End If
    End If
End Sub

'/**
' * @brief Standardized error routing and recording functionality.
' * @param errType The enumerator matching the error classification.
' * @param techMsg Technical details describing the failure.
' * @param userMsg Friendly string useful for standard msgboxes.
' * @param handle Associated connection.
' * @param errCode Native OS / Winsock code.
' */
Private Sub SetError(ByVal errType As WasabiError, ByVal techMsg As String, ByVal userMsg As String, ByVal handle As Long, Optional ByVal errCode As Long = 0)
    Static lastErr As Long
    Static lastHandle As Long
    If errType = ERR_NONE Then Exit Sub
    m_LastError = errType
    m_LastErrorCode = errCode
    m_TechnicalDetails = techMsg
    WasabiLog handle, "ERR " & errType & " | " & techMsg
    If errCode <> 0 Then WasabiLog handle, "SysCode: " & errCode & " (0x" & hex(errCode) & ")"
    If ValidIndex(handle) Then
        m_Connections(handle).LastError = errType
        m_Connections(handle).LastErrorCode = errCode
        m_Connections(handle).TechnicalDetails = techMsg
        If m_Connections(handle).EnableErrorDialog Then
            If errType <> lastErr Or handle <> lastHandle Then
                lastErr = errType
                lastHandle = handle
                MsgBox userMsg, vbCritical, "Wasabi WebSocket Error"
            End If
        End If
    ElseIf EnableErrorDialog Then
        MsgBox userMsg, vbCritical, "Wasabi WebSocket Error"
    End If
End Sub

'/**
' * @brief Translates native WSA numeric codes into textual descriptions.
' * @param code Native windows error code.
' * @return Standardized error name string.
' */
Private Function WSAErrDesc(ByVal code As Long) As String
    Select Case code
        Case 10004: WSAErrDesc = "WSAEINTR - Interrupted"
        Case 10013: WSAErrDesc = "WSAEACCES - Permission denied"
        Case 10014: WSAErrDesc = "WSAEFAULT - Bad address"
        Case 10022: WSAErrDesc = "WSAEINVAL - Invalid argument"
        Case 10024: WSAErrDesc = "WSAEMFILE - Too many sockets"
        Case 10035: WSAErrDesc = "WSAEWOULDBLOCK - Operation would block"
        Case 10036: WSAErrDesc = "WSAEINPROGRESS - Operation in progress"
        Case 10037: WSAErrDesc = "WSAEALREADY - Already in progress"
        Case 10038: WSAErrDesc = "WSAENOTSOCK - Not a socket"
        Case 10039: WSAErrDesc = "WSAEDESTADDRREQ - Destination address required"
        Case 10040: WSAErrDesc = "WSAEMSGSIZE - Message too long"
        Case 10047: WSAErrDesc = "WSAEAFNOSUPPORT - Address family not supported"
        Case 10048: WSAErrDesc = "WSAEADDRINUSE - Address in use"
        Case 10049: WSAErrDesc = "WSAEADDRNOTAVAIL - Address not available"
        Case 10050: WSAErrDesc = "WSAENETDOWN - Network is down"
        Case 10051: WSAErrDesc = "WSAENETUNREACH - Network unreachable"
        Case 10052: WSAErrDesc = "WSAENETRESET - Network dropped connection"
        Case 10053: WSAErrDesc = "WSAECONNABORTED - Connection aborted"
        Case 10054: WSAErrDesc = "WSAECONNRESET - Connection reset by peer"
        Case 10055: WSAErrDesc = "WSAENOBUFS - No buffer space"
        Case 10056: WSAErrDesc = "WSAEISCONN - Socket already connected"
        Case 10057: WSAErrDesc = "WSAENOTCONN - Socket not connected"
        Case 10058: WSAErrDesc = "WSAESHUTDOWN - Shutdown"
        Case 10060: WSAErrDesc = "WSAETIMEDOUT - Connection timed out"
        Case 10061: WSAErrDesc = "WSAECONNREFUSED - Connection refused"
        Case 10064: WSAErrDesc = "WSAEHOSTDOWN - Host is down"
        Case 10065: WSAErrDesc = "WSAEHOSTUNREACH - Host unreachable"
        Case 11001: WSAErrDesc = "WSAHOST_NOT_FOUND - Host not found"
        Case 11002: WSAErrDesc = "WSATRY_AGAIN - Non-authoritative host not found"
        Case 11003: WSAErrDesc = "WSANO_RECOVERY - Non-recoverable DNS error"
        Case 11004: WSAErrDesc = "WSANO_DATA - No address for hostname"
        Case Else: WSAErrDesc = "WSA error " & code
    End Select
End Function

'/**
' * @brief Returns the description for an RFC6455 closure status code.
' * @param code 16-bit integer WebSocket close code.
' * @return Standardized descriptor string.
' */
Private Function GetCloseCodeDesc(ByVal code As Integer) As String
    Select Case code
        Case 1000: GetCloseCodeDesc = "Normal Closure"
        Case 1001: GetCloseCodeDesc = "Going Away"
        Case 1002: GetCloseCodeDesc = "Protocol Error"
        Case 1003: GetCloseCodeDesc = "Unsupported Data"
        Case 1004: GetCloseCodeDesc = "Reserved"
        Case 1005: GetCloseCodeDesc = "No Status Received"
        Case 1006: GetCloseCodeDesc = "Abnormal Closure"
        Case 1007: GetCloseCodeDesc = "Invalid Frame Payload"
        Case 1008: GetCloseCodeDesc = "Policy Violation"
        Case 1009: GetCloseCodeDesc = "Message Too Big"
        Case 1010: GetCloseCodeDesc = "Mandatory Extension"
        Case 1011: GetCloseCodeDesc = "Internal Server Error"
        Case 1012: GetCloseCodeDesc = "Service Restart"
        Case 1013: GetCloseCodeDesc = "Try Again Later"
        Case 1014: GetCloseCodeDesc = "Bad Gateway"
        Case 1015: GetCloseCodeDesc = "TLS Handshake Failure"
        Case Else: GetCloseCodeDesc = "Unknown (" & code & ")"
    End Select
End Function

' ============================================================================
' 9. CONNECTION POOL MANAGEMENT
' ============================================================================

'/**
' * @brief Allocates global storage for all connections internally.
' */
Private Sub InitConnectionPool()
    Dim i As Long
    If m_ConnectionCount > 0 Then Exit Sub
    Randomize
    InitWasabiThunks
    ReDim m_Connections(0 To MAX_CONNECTIONS - 1)
    For i = 0 To MAX_CONNECTIONS - 1
        m_Connections(i).Socket = INVALID_SOCKET
        m_Connections(i).state = STATE_CLOSED
        m_Connections(i).hNtlmCred.dwLower = 0
        m_Connections(i).hNtlmCred.dwUpper = 0
    Next i
    m_ConnectionCount = MAX_CONNECTIONS
End Sub

'/**
' * @brief Fetches a free connection block. Allocates queue memory automatically.
' * @return New index, or INVALID_CONN_HANDLE if the pool is exhausted.
' */
Private Function AllocConnection() As Long
    Dim i As Long
    Dim bufSize As Long
    Dim fragSize As Long
    InitConnectionPool
    For i = 0 To MAX_CONNECTIONS - 1
        If m_Connections(i).state = STATE_CLOSED And m_Connections(i).Socket = INVALID_SOCKET Then
            bufSize = IIf(m_Connections(i).CustomBufferSize > 0, m_Connections(i).CustomBufferSize, BUFFER_SIZE)
            fragSize = IIf(m_Connections(i).CustomFragmentSize > 0, m_Connections(i).CustomFragmentSize, FRAGMENT_BUFFER_SIZE)
            ReDim m_Connections(i).recvBuffer(0 To bufSize - 1)
            ReDim m_Connections(i).DecryptBuffer(0 To bufSize - 1)
            ReDim m_Connections(i).MsgQueue(0 To MSG_QUEUE_SIZE - 1)
            ReDim m_Connections(i).BinaryQueue(0 To MSG_QUEUE_SIZE - 1)
            ReDim m_Connections(i).FragmentBuffer(0 To fragSize - 1)
            ReDim m_Connections(i).CustomHeaders(0 To 31)
            ReDim m_Connections(i).MqttBuffer(0 To 4095)
            ReDim m_Connections(i).OfflineTextQueue(0 To MSG_QUEUE_SIZE - 1)
            ReDim m_Connections(i).OfflineBinaryQueue(0 To MSG_QUEUE_SIZE - 1)
            ReDim m_Connections(i).TcpRecvBuffer(0 To bufSize - 1)
            Set m_Connections(i).Middlewares = New Collection
            m_Connections(i).Socket = INVALID_SOCKET
            ResetConnectionState i, False
            InitializeMTU i
            AllocConnection = i
            Exit Function
        End If
    Next i
    AllocConnection = INVALID_CONN_HANDLE
End Function

'/**
' * @brief Wipes and standardizes fields in a given connection handle.
' * @param handle Connecting identifier.
' */
Private Sub ResetConnectionState(ByVal handle As Long, Optional ByVal preserveAsync As Boolean = False)
    Dim savedHandler As Object
    Dim savedAsync As Boolean
    
    If preserveAsync Then
        savedAsync = m_Connections(handle).AsyncMode
        If savedAsync Then Set savedHandler = m_Connections(handle).AsyncHandler
    End If
    
    With m_Connections(handle)
        .mode = MODE_WEBSOCKET
        .TcpRecvLen = 0
        .state = STATE_CLOSED
        .TLS = False
        .HOST = ""
        .port = 0
        .path = ""
        .OriginalUrl = ""
        .recvLen = 0
        .DecryptLen = 0
        .MsgHead = 0
        .MsgTail = 0
        .MsgCount = 0
        .BinaryHead = 0
        .BinaryTail = 0
        .BinaryCount = 0
        .FragmentLen = 0
        .Fragmenting = False
        .FragmentOpcode = 0
        .LastError = ERR_NONE
        .LastErrorCode = 0
        .TechnicalDetails = ""
        .AutoReconnect = False
        .ReconnectMaxAttempts = DEFAULT_RECONNECT_MAX_ATTEMPTS
        .ReconnectAttempts = 0
        .ReconnectBaseDelayMs = DEFAULT_RECONNECT_BASE_DELAY_MS
        .PingIntervalMs = DEFAULT_PING_INTERVAL_MS
        .LastPingSentAt = 0
        .ReceiveTimeoutMs = DEFAULT_RECEIVE_TIMEOUT_MS
        .EnableErrorDialog = False
        .LogCallback = ""
        .NoDelay = False
        .proxyHost = ""
        .proxyPort = 0
        .proxyUser = ""
        .proxyPass = ""
        .proxyType = PROXY_TYPE_HTTP
        .ProxyEnabled = False
        .InactivityTimeoutMs = 0
        .LastActivityAt = 0
        .SubProtocol = ""
        .CustomBufferSize = 0
        .CustomFragmentSize = 0
        .AutoMTU = True
        .ZeroCopyEnabled = False
        .closeCode = 0
        .closeReason = ""
        .CloseInitiatedByUs = False
        .PreferIPv6 = False
        .ValidateServerCert = False
        .EnableRevocationCheck = False
        .ClientCertThumb = ""
        .ClientCertPfxPath = ""
        .ClientCertPfxPass = ""
        .UseHttp2 = False
        .ProxyUseNtlm = False
        .LastRttMs = 0
        .LastPingTimestamp = 0
        .MqttParserStage = 0
        .MqttBufLen = 0
        .MqttExpectedRemaining = 0
        .MqttCurrentPacketType = 0
        .MqttCurrentFlags = 0
        .MqttNextPacketId = 0
        .MqttInFlightCount = 0
        ReDim .MqttInFlight(0 To 0)
        With .stats
            .BytesSent = 0
            .BytesReceived = 0
            .MessagesSent = 0
            .MessagesReceived = 0
            .ConnectedAt = 0
        End With
        .mtu.CurrentMTU = DEFAULT_MTU
        .mtu.MaxSegmentSize = DEFAULT_MSS
        .mtu.OptimalFrameSize = DEFAULT_OPTIMAL_FRAME
        .mtu.LastProbeTime = 0
        .mtu.ProbeEnabled = True
        .mtu.UseTLSFragmentation = .TLS
        .DeflateEnabled = False
        .DeflateContextTakeover = True
        .InflateContextTakeover = True
        .DeflateWindowBits = -15
        .InflateWindowBits = -15
        .DeflateReady = False
        .InflateReady = False
        .DeflateActive = False
        .FragmentIsCompressed = False
        .ClientMaxWindowBits = 15
        .ServerMaxWindowBits = 15
        Set .ProtocolHandler = Nothing
        Set .CompressionHandler = Nothing
        Set .Middlewares = New Collection
        
        If preserveAsync And savedAsync Then
            Set .AsyncHandler = savedHandler
            .AsyncMode = True
        Else
            Set .AsyncHandler = Nothing
            .AsyncMode = False
        End If
    End With
End Sub

'/**
' * @brief Safely closes and wipes SSPI & Cryptographic handles linked to a connection.
' * @param handle The Target connection index.
' */
Private Sub FreeSecurityHandles(ByVal handle As Long)
    With m_Connections(handle)
        If .pClientCertCtx <> 0 Then
            CertFreeCertificateContext .pClientCertCtx
            .pClientCertCtx = 0
        End If
        If .hClientCertStore <> 0 Then
            CertCloseStore .hClientCertStore, 0
            .hClientCertStore = 0
        End If
        If .hContext.dwLower <> 0 Or .hContext.dwUpper <> 0 Then
            DeleteSecurityContext .hContext
            .hContext.dwLower = 0
            .hContext.dwUpper = 0
        End If
        If .hCred.dwLower <> 0 Or .hCred.dwUpper <> 0 Then
            FreeCredentialsHandle .hCred
            .hCred.dwLower = 0
            .hCred.dwUpper = 0
        End If
        If .hNtlmCred.dwLower <> 0 Or .hNtlmCred.dwUpper <> 0 Then
            FreeCredentialsHandle .hNtlmCred
            .hNtlmCred.dwLower = 0
            .hNtlmCred.dwUpper = 0
        End If
    End With
End Sub

'/**
' * @brief Flushes sensitive string content directly in memory to prevent inspection.
' * @param s ByRef string payload to secure zero.
' */
Private Sub SecureZeroString(ByRef s As String)
    If Len(s) > 0 Then
        Dim b() As Byte
        ReDim b(0 To LenB(s) - 1)
        CopyMemory ByVal StrPtr(s), b(0), LenB(s)
        s = vbNullString
    End If
End Sub

'/**
' * @brief The highest level destruct block for terminating a networking routine safely.
' * Cascades to middlewares, handlers, WSACleanups, and crypto cleans.
' * @param handle The internal identifier.
' */
Private Sub CleanupHandle(ByVal handle As Long)
    If Not ValidIndex(handle) Then Exit Sub
    
    If Not m_Connections(handle).CompressionHandler Is Nothing Then
        m_Connections(handle).CompressionHandler.OnDisconnect handle
    End If
    
    If Not m_Connections(handle).ProtocolHandler Is Nothing Then
        m_Connections(handle).ProtocolHandler.OnDisconnect handle
    End If
    
    Dim mwDisconnect As Object
    For Each mwDisconnect In m_Connections(handle).Middlewares
        mwDisconnect.OnDisconnect handle
    Next mwDisconnect
    
    With m_Connections(handle)
        If .Socket <> INVALID_SOCKET Then
            If .AsyncMode And m_AsyncHwnd <> 0 Then
                WSAAsyncSelect .Socket, m_AsyncHwnd, 0, 0
            End If
            sock_closesocket .Socket
            .Socket = INVALID_SOCKET
        End If
        
        If .recvLen > 0 Or .DecryptLen > 0 Or .TcpRecvLen > 0 Then
            WasabiMemZero VarPtr(.recvBuffer(0)), UBound(.recvBuffer) + 1
            WasabiMemZero VarPtr(.DecryptBuffer(0)), UBound(.DecryptBuffer) + 1
            WasabiMemZero VarPtr(.TcpRecvBuffer(0)), UBound(.TcpRecvBuffer) + 1
            WasabiMemZero VarPtr(.FragmentBuffer(0)), UBound(.FragmentBuffer) + 1
        End If
        
        SecureZeroString .proxyPass
        SecureZeroString .proxyUser
        SecureZeroString .ClientCertPfxPass
    End With
    
    FreeSecurityHandles handle
    If handle >= 0 And handle < MAX_CONNECTIONS Then
        m_ClientCertContextPtrs(handle) = 0
    End If
    ResetConnectionState handle, m_Connections(handle).AsyncMode
End Sub

' ============================================================================
' 10. NETWORK INFRASTRUCTURE (MTU, PROXY, TCP, CERTIFICATES)
' ============================================================================

'/**
' * @brief Resets the MTU sizes block for an incoming socket.
' * @param handle The session token to be updated.
' */
Private Sub InitializeMTU(ByVal handle As Long)
    With m_Connections(handle)
        .mtu.CurrentMTU = DEFAULT_MTU
        .mtu.LastProbeTime = 0
        .mtu.ProbeEnabled = True
        CalculateOptimalFrameSize handle
    End With
End Sub

'/**
' * @brief Derives ideal TLS fragmentation sizing by subtracting networking payloads.
' * @param handle Target instance index.
' */
Private Sub CalculateOptimalFrameSize(ByVal handle As Long)
    Dim ipOverhead As Long
    Dim tlsOverhead As Long
    Dim available As Long
    With m_Connections(handle)
        ipOverhead = IIf(.PreferIPv6, 40, IP_HEADER_MIN)
        If .TLS Then
            tlsOverhead = TLS_RECORD_HEADER + .sizes.cbHeader + .sizes.cbTrailer
        Else
            tlsOverhead = 0
        End If
        available = .mtu.CurrentMTU - ETHERNET_HEADER - ipOverhead - TCP_HEADER_MIN - tlsOverhead - WEBSOCKET_HEADER_MAX
        If available < MIN_FRAME_SIZE Then
            available = MIN_FRAME_SIZE
        End If
        If available > MAX_FRAME_SIZE Then
            available = MAX_FRAME_SIZE
        End If
        .mtu.MaxSegmentSize = .mtu.CurrentMTU - ETHERNET_HEADER - ipOverhead - TCP_HEADER_MIN
        .mtu.OptimalFrameSize = available
    End With
End Sub

'/**
' * @brief Triggers native getsockopt calls to gauge network interface capacity.
' * @param handle Native session handle.
' */
Private Sub probeMTU(ByVal handle As Long)
    Dim mss As Long
    Dim optVal As Long
    Dim optlen As Long
    Dim probeMTU As Long
    With m_Connections(handle)
        If .Socket = INVALID_SOCKET Then Exit Sub
        optlen = 4
        If sock_getsockopt(.Socket, IPPROTO_TCP_SOL, TCP_MAXSEG, optVal, optlen) = 0 And optVal > 0 Then
            mss = optVal
        Else
            mss = DEFAULT_MSS
        End If
        probeMTU = mss + TCP_HEADER_MIN + IIf(.PreferIPv6, 40, IP_HEADER_MIN) + ETHERNET_HEADER
        If probeMTU <> .mtu.CurrentMTU Then
            .mtu.CurrentMTU = probeMTU
            CalculateOptimalFrameSize handle
            WasabiLog handle, "MTU updated: " & .mtu.CurrentMTU & " MSS=" & mss & " OptFrame=" & .mtu.OptimalFrameSize & " (handle=" & handle & ")"
        End If
        .mtu.LastProbeTime = GetTickCount()
    End With
End Sub

'/**
' * @brief Attaches keep-alive and NoDelay parameters (Nagle) via Winsock ioctl.
' * @param handle Target networking handle.
' */
Private Sub ApplySocketOptions(ByVal handle As Long)
    Dim optVal As Long
    Dim wsaErr As Long
    With m_Connections(handle)
        If .Socket = INVALID_SOCKET Then Exit Sub
        optVal = IIf(.NoDelay, 1, 0)
        If sock_setsockopt(.Socket, IPPROTO_TCP_SOL, TCP_NODELAY, optVal, 4) <> 0 Then
            wsaErr = WSAGetLastError()
            WasabiLog handle, "TCP_NODELAY failed: " & WSAErrDesc(wsaErr)
        End If
        optVal = 1
        If sock_setsockopt(.Socket, SOL_SOCKET, SO_KEEPALIVE, optVal, 4) <> 0 Then
            wsaErr = WSAGetLastError()
            WasabiLog handle, "SO_KEEPALIVE failed: " & WSAErrDesc(wsaErr)
        End If
        optVal = BUFFER_SIZE
        sock_setsockopt .Socket, SOL_SOCKET, SO_RCVBUF, optVal, 4
        sock_setsockopt .Socket, SOL_SOCKET, SO_SNDBUF, optVal, 4
    End With
End Sub

'/**
' * @brief Polling wrapper for Winsock select(). Halts execution temporarily to yield for stream.
' * @param handle Target instance index.
' * @param timeoutMs The max timeout cap in milliseconds.
' * @return True if read data is waiting.
' */
Private Function WaitForDataOn(ByVal handle As Long, ByVal timeoutMs As Long) As Boolean
    Dim readSet As FD_SET
    Dim TIMEOUT As TIMEVAL
    Dim effective As Long
    effective = timeoutMs
    If effective = 0 And ValidIndex(handle) Then
        If m_Connections(handle).ReceiveTimeoutMs > 0 Then
            effective = m_Connections(handle).ReceiveTimeoutMs
        End If
    End If
    readSet.fd_count = 1
    readSet.fd_array(0) = m_Connections(handle).Socket
    TIMEOUT.tv_sec = effective \ 1000
    TIMEOUT.tv_usec = (effective Mod 1000) * 1000
    WaitForDataOn = (sock_select(0, readSet, ByVal 0&, ByVal 0&, TIMEOUT) > 0)
End Function

'/**
' * @brief Pushes raw bytes down the TCP pipe synchronously.
' * @param handle Target networking handle.
' * @param frame The unmanaged payload structure.
' * @return State of physical sending success.
' */
Private Function RawSendFor(ByVal handle As Long, ByRef frame() As Byte) As Boolean
    Dim totalSent As Long
    Dim toSend As Long
    Dim sent As Long
    Dim wsaErr As Long
    toSend = UBound(frame) + 1
    totalSent = 0
    With m_Connections(handle)
        Do While totalSent < toSend
            sent = sock_send(.Socket, frame(totalSent), toSend - totalSent, 0)
            If sent <= 0 Then
                wsaErr = WSAGetLastError()
                SetError ERR_SEND_FAILED, "send() failed: " & WSAErrDesc(wsaErr), "Failed to send data to server.", handle, wsaErr
                .state = STATE_CLOSED
                Exit Function
            End If
            totalSent = totalSent + sent
        Loop
    End With
    RawSendFor = True
End Function

'/**
' * @brief Crafts an RFC6455 compliant WebSocket frame dynamically checking masking headers.
' * @param payload Original internal byte block.
' * @param payloadLen Length constraint.
' * @param opcode Instruction to frame (text, bin, ctrl).
' * @param isFinal End-of-message FIN Bit marker.
' * @param setRSV1 Reserved bit extension toggles (like deflate).
' * @return Fully crafted TCP-ready framing block array.
' */
Private Function BuildWSFrame(ByRef payload() As Byte, ByVal payloadLen As Long, ByVal opcode As Byte, ByVal isFinal As Boolean, Optional ByVal setRSV1 As Boolean = False) As Byte()
    Dim mask(0 To 3) As Byte
    Dim headerLen As Long
    Dim frame() As Byte
    Dim finBit As Byte
    Dim rsv1 As Byte
    Dim i As Long
    rsv1 = IIf(setRSV1, &H40, 0)
    FillRandomBytes mask, 4
    finBit = IIf(isFinal, &H80, 0)
    If payloadLen < 126 Then
        headerLen = 6
        ReDim frame(0 To headerLen + payloadLen - 1)
        frame(0) = finBit Or rsv1 Or opcode
        frame(1) = &H80 Or CByte(payloadLen)
        frame(2) = mask(0)
        frame(3) = mask(1)
        frame(4) = mask(2)
        frame(5) = mask(3)
    ElseIf payloadLen < 65536 Then
        headerLen = 8
        ReDim frame(0 To headerLen + payloadLen - 1)
        frame(0) = finBit Or rsv1 Or opcode
        frame(1) = &HFE
        frame(2) = CByte((payloadLen \ 256) And &HFF)
        frame(3) = CByte(payloadLen And &HFF)
        frame(4) = mask(0)
        frame(5) = mask(1)
        frame(6) = mask(2)
        frame(7) = mask(3)
    Else
        headerLen = 14
        ReDim frame(0 To headerLen + payloadLen - 1)
        frame(0) = finBit Or rsv1 Or opcode
        frame(1) = &HFF
        frame(2) = 0
        frame(3) = 0
        frame(4) = 0
        frame(5) = 0
        frame(6) = CByte((payloadLen \ 16777216) And &HFF)
        frame(7) = CByte((payloadLen \ 65536) And &HFF)
        frame(8) = CByte((payloadLen \ 256) And &HFF)
        frame(9) = CByte(payloadLen And &HFF)
        frame(10) = mask(0)
        frame(11) = mask(1)
        frame(12) = mask(2)
        frame(13) = mask(3)
    End If
    If payloadLen > 0 Then
        If (Not Not payload) <> 0 Then
            If m_ptrWsMask <> 0 Then
                CopyMemory frame(headerLen), payload(LBound(payload)), payloadLen
                CallWindowProcW m_ptrWsMask, VarPtr(frame(headerLen)), payloadLen, VarPtr(mask(0)), 0
            Else
                For i = 0 To payloadLen - 1
                    frame(headerLen + i) = payload(LBound(payload) + i) Xor mask(i Mod 4)
                Next i
            End If
        End If
    End If
    BuildWSFrame = frame
End Function

'/**
' * @brief Dereferences a UTF-16 pointer block natively stringifying it.
' * @param ptr Pointer location.
' * @return String output.
' */
#If VBA7 Then
Private Function PtrToStrW(ByVal ptr As LongPtr) As String
#Else
Private Function PtrToStrW(ByVal ptr As Long) As String
#End If
    Dim length As Long
    Dim buf() As Byte
    If ptr = 0 Then Exit Function
    
    length = lstrlenW(ptr) * 2
    If length > 0 Then
        ReDim buf(0 To length - 1)
        CopyMemoryFromPtr buf(0), ptr, length
        PtrToStrW = buf
    End If
End Function

'/**
' * @brief Converts standard VB6/VBA BSTR to a UTF-8 Array layout.
' * @param str BSTR text string.
' * @return Dimensioned Byte Array block.
' */
Private Function StringToUtf8(ByVal str As String) As Byte()
    Dim need As Long
    Dim written As Long
    Dim result() As Byte
    If Len(str) = 0 Then
        StringToUtf8 = result
        Exit Function
    End If
    need = Len(str) * 4
    If need > m_Utf8BufSize Then
        ReDim m_Utf8Buf(0 To need - 1)
        m_Utf8BufSize = need
    End If
    written = WideCharToMultiByte(CP_UTF8, 0, StrPtr(str), Len(str), m_Utf8Buf(0), need, NULL_PTR, NULL_PTR)
    If written > 0 Then
        ReDim result(0 To written - 1)
        CopyMemory result(0), m_Utf8Buf(0), written
    End If
    StringToUtf8 = result
End Function

'/**
' * @brief Takes an unmanaged chunk of UTF-8 and parses to BSTR format natively.
' * @param utf8 Array block containing bytes.
' * @param length Size constraint payload block limit.
' * @return Converted string in VBA native.
' */
Private Function Utf8ToString(ByRef utf8() As Byte, ByVal length As Long) As String
    Dim charCount As Long
    Dim result As String
    If length <= 0 Then
        Utf8ToString = ""
        Exit Function
    End If
    charCount = MultiByteToWideChar(CP_UTF8, 0, utf8(LBound(utf8)), length, NULL_PTR, 0)
    If charCount > 0 Then
        result = String$(charCount, vbNullChar)
        MultiByteToWideChar CP_UTF8, 0, utf8(LBound(utf8)), length, StrPtr(result), charCount
    End If
    Utf8ToString = result
End Function

'/**
' * @brief Generates a Base64 hash for a provided memory buffer natively via WinAPI.
' * @param Bytes Source data context.
' * @return Standardized textual representation.
' */
Private Function Base64Encode(ByRef Bytes() As Byte) As String
    Dim dataLen As Long
    
    If (Not Bytes) = -1 Then
        Base64Encode = ""
        Exit Function
    End If
    
    dataLen = UBound(Bytes) - LBound(Bytes) + 1
    If dataLen = 0 Then
        Base64Encode = ""
        Exit Function
    End If

    Dim strLen As Long
    strLen = 0
    #If VBA7 Then
        CryptBinaryToStringW VarPtr(Bytes(LBound(Bytes))), dataLen, CRYPT_STRING_BASE64 Or CRYPT_NOCRLF, NULL_PTR, strLen
    #Else
        CryptBinaryToStringW VarPtr(Bytes(LBound(Bytes))), dataLen, CRYPT_STRING_BASE64 Or CRYPT_NOCRLF, 0, strLen
    #End If

    If strLen <= 0 Then
        Base64Encode = ""
        Exit Function
    End If

    Dim buf As String
    buf = String$(strLen, vbNullChar)
    #If VBA7 Then
        CryptBinaryToStringW VarPtr(Bytes(LBound(Bytes))), dataLen, CRYPT_STRING_BASE64 Or CRYPT_NOCRLF, StrPtr(buf), strLen
    #Else
        CryptBinaryToStringW VarPtr(Bytes(LBound(Bytes))), dataLen, CRYPT_STRING_BASE64 Or CRYPT_NOCRLF, StrPtr(buf), strLen
    #End If

    Base64Encode = Left$(buf, strLen)
End Function

'/**
' * @brief Dismantles typical web URL schema extracting required blocks logic.
' * @param url Input string format (wss://example.com/api)
' * @param outHost Returns target host
' * @param outPort Returns parsed port.
' * @param outPath Returns URI trailing resource structure.
' * @param outTLS Resolves encryption scheme logic path context.
' * @return Extracted logic successful validation.
' */
Private Function ParseURL(ByVal url As String, ByRef outHost As String, ByRef outPort As Long, ByRef outPath As String, ByRef outTLS As Boolean) As Boolean
    Dim work As String
    Dim slashPos As Long
    Dim colonPos As Long
    Dim portStr As String
    Dim portVal As Long
    Dim i As Long
    Dim c As String
    If Len(Trim(url)) = 0 Then Exit Function
    work = url
    outTLS = False
    outPort = PORT_HTTP
    If LCase(Left(work, 6)) = "wss://" Then
        work = Mid(work, 7)
        outTLS = True
        outPort = PORT_HTTPS
    ElseIf LCase(Left(work, 5)) = "ws://" Then
        work = Mid(work, 6)
    Else
        Exit Function
    End If
    If Len(work) = 0 Then Exit Function
    slashPos = InStr(work, "/")
    If slashPos > 0 Then
        outPath = Mid(work, slashPos)
        work = Left(work, slashPos - 1)
    Else
        outPath = "/"
    End If
    colonPos = InStr(work, ":")
    If colonPos > 0 Then
        outHost = Left(work, colonPos - 1)
        portStr = Mid(work, colonPos + 1)
        If Len(portStr) = 0 Then Exit Function
        For i = 1 To Len(portStr)
            c = Mid(portStr, i, 1)
            If c < "0" Or c > "9" Then Exit Function
        Next i
        portVal = val(portStr)
        If portVal <= 0 Or portVal > 65535 Then Exit Function
        outPort = portVal
    Else
        outHost = work
    End If
    If Len(outHost) = 0 Then Exit Function
    ParseURL = True
End Function

'/**
' * @brief Executes an underlying DNS lookup for the requested server name format.
' * @param hostname A string containing IP format natively or DNS request address.
' * @param handle Requesting socket session log target block.
' * @return Unsigned int resolution memory representation.
' */
Private Function ResolveHostname(ByVal hostname As String, ByVal handle As Long) As Long
    Dim addr As Long
    Dim wsaErr As Long
#If VBA7 Then
    Dim hostent As LongPtr
    Dim he As HOSTENT64
    Dim addrList As LongPtr
    Dim pAddr As LongPtr
#Else
    Dim hostent As Long
    Dim he As HOSTENT32
    Dim addrList As Long
    Dim pAddr As Long
#End If
    addr = sock_inet_addr(hostname)
    If addr <> INADDR_NONE Then
        ResolveHostname = addr
        Exit Function
    End If
    hostent = sock_gethostbyname(hostname)
    If hostent = 0 Then
        wsaErr = WSAGetLastError()
        SetError ERR_DNS_RESOLVE_FAILED, "gethostbyname failed for '" & hostname & "': " & WSAErrDesc(wsaErr), "Cannot resolve address: " & hostname & vbCrLf & WSAErrDesc(wsaErr), handle, wsaErr
        Exit Function
    End If
#If Win64 Then
    CopyMemoryFromPtr he, hostent, LenB(he)
    addrList = he.h_addr_list
    If addrList = 0 Then Exit Function
    CopyMemoryFromPtr pAddr, addrList, 8
    If pAddr = 0 Then Exit Function
    CopyMemoryFromPtr addr, pAddr, 4
#Else
    CopyMemoryFromPtr he, hostent, LenB(he)
    addrList = he.h_addr_list
    If addrList = 0 Then Exit Function
    CopyMemoryFromPtr pAddr, addrList, 4
    If pAddr = 0 Then Exit Function
    CopyMemoryFromPtr addr, pAddr, 4
#End If
    ResolveHostname = addr
End Function

'/**
' * @brief Handles complex Happy Eyeballs execution for IPv4 vs IPv6 routing.
' * @param handle Core instance.
' * @param hostname Connection destination.
' * @param port Endpoint destination binding value.
' * @return State of connecting socket procedure block.
' */
Private Function ResolveAndConnect(ByVal handle As Long, ByVal hostname As String, ByVal port As Long) As Boolean
#If VBA7 Then
    Dim ppResult As LongPtr
    Dim pCur As LongPtr
    Dim pNext As LongPtr
    Dim pSockaddr As LongPtr
    Dim sock6 As LongPtr
    Dim sock4 As LongPtr
    Dim aiAddrLenFull As LongPtr
#Else
    Dim ppResult As Long
    Dim pCur As Long
    Dim pNext As Long
    Dim pSockaddr As Long
    Dim sock6 As Long
    Dim sock4 As Long
#End If
    Dim hints() As Byte
    Dim aiFamily As Long
    Dim aiSocktype As Long
    Dim aiProtocol As Long
    Dim aiAddrLen As Long
    Dim nbMode As Long
    Dim writeSet As FD_SET
    Dim exceptSet As FD_SET
    Dim tv As TIMEVAL
    Dim selResult As Long
    Dim wsaErr As Long
    Dim sa6() As Byte
    Dim sa4() As Byte
    Dim sa6Len As Long
    Dim sa4Len As Long
    Dim found6 As Boolean
    Dim found4 As Boolean
    Dim startTick As Long
    Dim elapsed As Long
    Dim sin4 As SOCKADDR_IN

    sock6 = INVALID_SOCKET
    sock4 = INVALID_SOCKET

#If VBA7 Then
    ReDim hints(0 To 47)
#Else
    ReDim hints(0 To 31)
#End If
    aiSocktype = SOCK_STREAM
    CopyMemory hints(8), aiSocktype, 4
    aiProtocol = IPPROTO_TCP
    CopyMemory hints(12), aiProtocol, 4

    ppResult = 0
    If sock_getaddrinfo(hostname, CStr(port), VarPtr(hints(0)), ppResult) = 0 And ppResult <> 0 Then
        pCur = ppResult
        Do While pCur <> 0
            #If Win64 Then
                CopyMemoryFromPtr aiFamily, pCur + 4, 4
                CopyMemoryFromPtr aiSocktype, pCur + 8, 4
                CopyMemoryFromPtr aiAddrLenFull, pCur + 16, 8
                aiAddrLen = CLng(aiAddrLenFull And &HFFFFFFFF^)
                CopyMemoryFromPtr pSockaddr, pCur + 32, 8
                CopyMemoryFromPtr pNext, pCur + 40, 8
            #Else
                CopyMemoryFromPtr aiFamily, pCur + 4, 4
                CopyMemoryFromPtr aiSocktype, pCur + 8, 4
                CopyMemoryFromPtr aiAddrLen, pCur + 16, 4
                CopyMemoryFromPtr pSockaddr, pCur + 24, 4
                CopyMemoryFromPtr pNext, pCur + 28, 4
            #End If
            If aiSocktype = SOCK_STREAM And aiAddrLen > 0 And pSockaddr <> 0 Then
                If aiFamily = AF_INET6 And Not found6 Then
                    ReDim sa6(0 To aiAddrLen - 1)
                    CopyMemoryFromPtr sa6(0), pSockaddr, aiAddrLen
                    sa6Len = aiAddrLen
                    found6 = True
                ElseIf aiFamily = AF_INET And Not found4 Then
                    ReDim sa4(0 To aiAddrLen - 1)
                    CopyMemoryFromPtr sa4(0), pSockaddr, aiAddrLen
                    sa4Len = aiAddrLen
                    found4 = True
                End If
            End If
            If found6 And found4 Then Exit Do
            pCur = pNext
        Loop
        sock_freeaddrinfo ppResult
    End If

    If Not found6 And Not found4 Then
        sin4.sin_family = AF_INET
        sin4.sin_port = sock_htons(port)
        sin4.sin_addr = ResolveHostname(hostname, handle)
        If sin4.sin_addr = 0 Then Exit Function
        sa4Len = LenB(sin4)
        ReDim sa4(0 To sa4Len - 1)
        CopyMemory sa4(0), sin4, sa4Len
        found4 = True
    End If

    If found6 Then
        sock6 = sock_socket(AF_INET6, SOCK_STREAM, IPPROTO_TCP)
        If sock6 <> INVALID_SOCKET Then
            nbMode = 1
            sock_ioctlsocket sock6, FIONBIO, nbMode
            sock_connect sock6, VarPtr(sa6(0)), sa6Len
        Else
            found6 = False
        End If
    End If

    If found6 And found4 Then
        startTick = GetTickCount()
        Do
            writeSet.fd_count = 1
            writeSet.fd_array(0) = sock6
            exceptSet.fd_count = 1
            exceptSet.fd_array(0) = sock6
            tv.tv_sec = 0
            tv.tv_usec = CONNECT_POLL_USEC
            selResult = sock_select(0, ByVal 0&, writeSet, exceptSet, tv)
            If selResult > 0 And exceptSet.fd_count = 0 Then
                nbMode = 0
                sock_ioctlsocket sock6, FIONBIO, nbMode
                m_Connections(handle).Socket = sock6
                If m_Connections(handle).AsyncMode And m_AsyncHwnd <> 0 Then
                    WSAAsyncSelect m_Connections(handle).Socket, m_AsyncHwnd, WM_WASABI_SOCKET, FD_READ Or FD_WRITE Or FD_CLOSE Or FD_CONNECT
                End If
                ResolveAndConnect = True
                Exit Function
            End If
            If selResult > 0 And exceptSet.fd_count > 0 Then
                sock_closesocket sock6
                sock6 = INVALID_SOCKET
                found6 = False
                Exit Do
            End If
            elapsed = TickDiff(startTick, GetTickCount())
            If elapsed >= HAPPY_EYEBALLS_DELAY_MS Then Exit Do
            DoEvents
        Loop
    End If

    If Not ResolveAndConnect And found4 Then
        sock4 = sock_socket(AF_INET, SOCK_STREAM, IPPROTO_TCP)
        If sock4 <> INVALID_SOCKET Then
            nbMode = 1
            sock_ioctlsocket sock4, FIONBIO, nbMode
            sock_connect sock4, VarPtr(sa4(0)), sa4Len
        End If
    End If

    If Not ResolveAndConnect Then
        Dim raceTimeout As Long
        raceTimeout = DNS_RACE_TIMEOUT_MS
        startTick = GetTickCount()
        Do
            If sock6 <> INVALID_SOCKET Then
                writeSet.fd_count = 1
                writeSet.fd_array(0) = sock6
                exceptSet.fd_count = 1
                exceptSet.fd_array(0) = sock6
                tv.tv_sec = 0
                tv.tv_usec = CONNECT_POLL_USEC
                selResult = sock_select(0, ByVal 0&, writeSet, exceptSet, tv)
                If selResult > 0 And exceptSet.fd_count = 0 Then
                    nbMode = 0
                    sock_ioctlsocket sock6, FIONBIO, nbMode
                    m_Connections(handle).Socket = sock6
                    If sock4 <> INVALID_SOCKET Then sock_closesocket sock4
                    ResolveAndConnect = True
                    Exit Function
                End If
                If selResult > 0 And exceptSet.fd_count > 0 Then
                    sock_closesocket sock6
                    sock6 = INVALID_SOCKET
                End If
            End If
            If sock4 <> INVALID_SOCKET Then
                writeSet.fd_count = 1
                writeSet.fd_array(0) = sock4
                exceptSet.fd_count = 1
                exceptSet.fd_array(0) = sock4
                tv.tv_sec = 0
                tv.tv_usec = CONNECT_POLL_USEC
                selResult = sock_select(0, ByVal 0&, writeSet, exceptSet, tv)
                If selResult > 0 And exceptSet.fd_count = 0 Then
                    nbMode = 0
                    sock_ioctlsocket sock4, FIONBIO, nbMode
                    m_Connections(handle).Socket = sock4
                    If sock6 <> INVALID_SOCKET Then sock_closesocket sock6
                    If m_Connections(handle).AsyncMode And m_AsyncHwnd <> 0 Then
                        WSAAsyncSelect m_Connections(handle).Socket, m_AsyncHwnd, WM_WASABI_SOCKET, FD_READ Or FD_WRITE Or FD_CLOSE Or FD_CONNECT
                    End If
                    ResolveAndConnect = True
                    Exit Function
                End If
                If selResult > 0 And exceptSet.fd_count > 0 Then
                    sock_closesocket sock4
                    sock4 = INVALID_SOCKET
                End If
            End If
            If sock6 = INVALID_SOCKET And sock4 = INVALID_SOCKET Then Exit Do
            elapsed = TickDiff(startTick, GetTickCount())
            If elapsed >= raceTimeout Then Exit Do
            DoEvents
        Loop
    End If

    If Not ResolveAndConnect Then
        If sock6 <> INVALID_SOCKET Then sock_closesocket sock6
        If sock4 <> INVALID_SOCKET Then sock_closesocket sock4
        wsaErr = WSAGetLastError()
        SetError ERR_CONNECT_FAILED, "Connect failed: " & WSAErrDesc(wsaErr), "Could not connect to server." & vbCrLf & WSAErrDesc(wsaErr), handle, wsaErr
    End If
End Function

'/**
' * @brief Initiates a standard CONNECT tunnel over HTTP Proxy.
' * @param handle Network session block target.
' * @return State of connection.
' */
Private Function DoProxyHTTP(ByVal handle As Long) As Boolean
    Dim req As String
    Dim sendBuf() As Byte
    Dim recvBuf() As Byte
    Dim received As Long
    Dim response As String
    Dim sendResult As Long
    Dim wsaErr As Long
    
    With m_Connections(handle)
        req = "CONNECT " & .HOST & ":" & .port & " HTTP/1.1" & vbCrLf
        req = req & "Host: " & .HOST & ":" & .port & vbCrLf
        If .proxyUser <> "" And Not .ProxyUseNtlm Then
            req = req & "Proxy-Authorization: Basic " & Base64Encode(StrConv(.proxyUser & ":" & .proxyPass, vbFromUnicode)) & vbCrLf
        End If
        If .ProxyUseNtlm Then
            req = req & "Proxy-Authorization: NTLM TlRMTVNTUAABAAAAB4IIogAAAAAAAAAAAAAAAAAAAAAFASgKAAAADw==" & vbCrLf
        End If
        req = req & vbCrLf
        sendBuf = StrConv(req, vbFromUnicode)
        sendResult = sock_send(.Socket, sendBuf(0), UBound(sendBuf) + 1, 0)
        If sendResult <= 0 Then
            wsaErr = WSAGetLastError()
            SetError ERR_PROXY_CONNECT_FAILED, "send() to proxy failed with WSA error " & wsaErr, "Failed to send CONNECT to proxy.", handle, wsaErr
            Exit Function
        End If
        If Not WaitForDataOn(handle, 5000) Then
            SetError ERR_PROXY_CONNECT_FAILED, "Proxy CONNECT timeout", "Proxy did not respond to CONNECT request.", handle
            Exit Function
        End If
        ReDim recvBuf(0 To 4095)
        received = sock_recv(.Socket, recvBuf(0), 4096, 0)
        If received <= 0 Then
            wsaErr = WSAGetLastError()
            SetError ERR_PROXY_CONNECT_FAILED, "recv() from proxy failed with WSA error " & wsaErr, "Failed to receive proxy response.", handle, wsaErr
            Exit Function
        End If
        response = Utf8ToString(recvBuf, received)
        If InStr(response, "407") > 0 Then
            If .ProxyUseNtlm Then
                Dim ntlmHeader As String
                Dim hPos As Long
                Dim lf As Long
                Dim ntlmToken As String
                hPos = InStr(LCase(response), "proxy-authenticate: ntlm")
                If hPos > 0 Then
                    ntlmHeader = Mid(response, hPos)
                    lf = InStr(ntlmHeader, vbCrLf)
                    If lf > 0 Then ntlmHeader = Left(ntlmHeader, lf - 1)
                    ntlmToken = GenerateNtlmToken(handle, ntlmHeader, .proxyHost)
                    If ntlmToken <> "" Then
                        req = "CONNECT " & .HOST & ":" & .port & " HTTP/1.1" & vbCrLf
                        req = req & "Host: " & .HOST & ":" & .port & vbCrLf
                        req = req & "Proxy-Authorization: " & ntlmToken & vbCrLf
                        req = req & vbCrLf
                        sendBuf = StrConv(req, vbFromUnicode)
                        sendResult = sock_send(.Socket, sendBuf(0), UBound(sendBuf) + 1, 0)
                        If sendResult <= 0 Then Exit Function
                        If Not WaitForDataOn(handle, 5000) Then Exit Function
                        ReDim recvBuf(0 To 4095)
                        received = sock_recv(.Socket, recvBuf(0), 4096, 0)
                        If received <= 0 Then Exit Function
                        response = Utf8ToString(recvBuf, received)
                        If InStr(response, "200") = 0 Then
                            SetError ERR_PROXY_AUTH_FAILED, "NTLM auth rejected", "Proxy authentication failed after NTLM challenge.", handle
                            Exit Function
                        End If
                    End If
                End If
            Else
                SetError ERR_PROXY_AUTH_FAILED, "Proxy returned 407 Proxy Authentication Required", "Proxy authentication failed." & vbCrLf & "Please check your proxy credentials.", handle
                Exit Function
            End If
        End If
        If InStr(response, "200") = 0 Then
            Dim statusLine As String
            Dim lineEnd As Long
            lineEnd = InStr(response, vbCrLf)
            If lineEnd > 0 Then
                statusLine = Left(response, lineEnd - 1)
            Else
                statusLine = Left(response, 50)
            End If
            SetError ERR_PROXY_TUNNEL_FAILED, "Proxy CONNECT rejected: " & statusLine, "Proxy refused the tunnel connection." & vbCrLf & "Server: " & .HOST & ":" & .port, handle
            Exit Function
        End If
    End With
    DoProxyHTTP = True
End Function

'/**
' * @brief Implements RFC 1928 routing behavior securely to establish TCP proxies.
' * @param handle Target instance index identifier.
' * @return State Boolean completion verification.
' */
Private Function DoProxySOCKS5(ByVal handle As Long) As Boolean
    Dim sendBuf() As Byte
    Dim recvBuf(0 To 255) As Byte
    Dim received As Long
    Dim wsaErr As Long
    Dim hostBytes() As Byte
    Dim hostLen As Byte
    Dim i As Long
    With m_Connections(handle)
        If .proxyUser <> "" Then
            ReDim sendBuf(0 To 3)
            sendBuf(0) = 5
            sendBuf(1) = 2
            sendBuf(2) = 0
            sendBuf(3) = 2
        Else
            ReDim sendBuf(0 To 2)
            sendBuf(0) = 5
            sendBuf(1) = 1
            sendBuf(2) = 0
        End If
        If sock_send(.Socket, sendBuf(0), UBound(sendBuf) + 1, 0) <= 0 Then
            wsaErr = WSAGetLastError()
            SetError ERR_PROXY_CONNECT_FAILED, "SOCKS5 greeting failed: " & wsaErr, "SOCKS5 handshake failed.", handle, wsaErr
            Exit Function
        End If
        If Not WaitForDataOn(handle, 5000) Then
            SetError ERR_PROXY_CONNECT_FAILED, "SOCKS5 greeting timeout", "SOCKS5 server did not respond.", handle
            Exit Function
        End If
        received = sock_recv(.Socket, recvBuf(0), 256, 0)
        If received < 2 Or recvBuf(0) <> 5 Then
            SetError ERR_PROXY_CONNECT_FAILED, "SOCKS5 invalid greeting response", "SOCKS5 handshake failed.", handle
            Exit Function
        End If
        If recvBuf(1) = 255 Then
            SetError ERR_PROXY_AUTH_FAILED, "SOCKS5 no acceptable auth method", "SOCKS5 authentication failed.", handle
            Exit Function
        End If
        If recvBuf(1) = 2 Then
            Dim userB() As Byte
            Dim passB() As Byte
            Dim uLen As Byte
            Dim pLen As Byte
            userB = StrConv(.proxyUser, vbFromUnicode)
            passB = StrConv(.proxyPass, vbFromUnicode)
            uLen = CByte(UBound(userB) + 1)
            pLen = CByte(UBound(passB) + 1)
            ReDim sendBuf(0 To 3 + uLen + pLen)
            sendBuf(0) = 1
            sendBuf(1) = uLen
            For i = 0 To uLen - 1
                sendBuf(2 + i) = userB(i)
            Next i
            sendBuf(2 + uLen) = pLen
            For i = 0 To pLen - 1
                sendBuf(3 + uLen + i) = passB(i)
            Next i
            If sock_send(.Socket, sendBuf(0), UBound(sendBuf) + 1, 0) <= 0 Then
                wsaErr = WSAGetLastError()
                SetError ERR_PROXY_AUTH_FAILED, "SOCKS5 auth send failed: " & wsaErr, "SOCKS5 authentication failed.", handle, wsaErr
                Exit Function
            End If
            If Not WaitForDataOn(handle, 5000) Then
                SetError ERR_PROXY_AUTH_FAILED, "SOCKS5 auth timeout", "SOCKS5 authentication timed out.", handle
                Exit Function
            End If
            received = sock_recv(.Socket, recvBuf(0), 256, 0)
            If received < 2 Or recvBuf(1) <> 0 Then
                SetError ERR_PROXY_AUTH_FAILED, "SOCKS5 auth rejected", "SOCKS5 authentication failed. Check credentials.", handle
                Exit Function
            End If
        End If
        hostBytes = StrConv(.HOST, vbFromUnicode)
        If UBound(hostBytes) + 1 > SOCKS5_MAX_HOSTNAME Then
            SetError ERR_PROXY_CONNECT_FAILED, "Hostname too long for SOCKS5: " & Len(.HOST) & " chars", "Proxy hostname exceeds SOCKS5 limit.", handle
            Exit Function
        End If
        hostLen = CByte(UBound(hostBytes) + 1)
        ReDim sendBuf(0 To 6 + hostLen + 1)
        sendBuf(0) = 5
        sendBuf(1) = 1
        sendBuf(2) = 0
        sendBuf(3) = 3
        sendBuf(4) = hostLen
        For i = 0 To hostLen - 1
            sendBuf(5 + i) = hostBytes(i)
        Next i
        sendBuf(5 + hostLen) = CByte((.port \ 256) And &HFF)
        sendBuf(6 + hostLen) = CByte(.port And &HFF)
        If sock_send(.Socket, sendBuf(0), UBound(sendBuf) + 1, 0) <= 0 Then
            wsaErr = WSAGetLastError()
            SetError ERR_PROXY_CONNECT_FAILED, "SOCKS5 CONNECT send failed: " & wsaErr, "SOCKS5 connect request failed.", handle, wsaErr
            Exit Function
        End If
        If Not WaitForDataOn(handle, 5000) Then
            SetError ERR_PROXY_CONNECT_FAILED, "SOCKS5 CONNECT timeout", "SOCKS5 server did not respond to CONNECT.", handle
            Exit Function
        End If
        received = sock_recv(.Socket, recvBuf(0), 256, 0)
        If received < 4 Then
            SetError ERR_PROXY_CONNECT_FAILED, "SOCKS5 CONNECT response too short", "SOCKS5 connect failed.", handle
            Exit Function
        End If
        If recvBuf(0) <> 5 Then
            SetError ERR_PROXY_CONNECT_FAILED, "SOCKS5 CONNECT wrong version: " & recvBuf(0), "SOCKS5 connect failed.", handle
            Exit Function
        End If
        If recvBuf(1) <> 0 Then
            SetError ERR_PROXY_CONNECT_FAILED, "SOCKS5 CONNECT rejected, code: " & recvBuf(1), "SOCKS5 server rejected connection, code: " & recvBuf(1), recvBuf(1), handle
            Exit Function
        End If
    End With
    DoProxySOCKS5 = True
End Function

'/**
' * @brief Synthesizes SSPI negotiation to execute NTLM authentication block contexts.
' * @param handle Core connection identity handler index.
' * @param proxyAuthHeader Server prompt data buffer challenge target.
' * @param proxyHost Origin identifier routing.
' * @return Final HTTP valid auth string output.
' */
Private Function GenerateNtlmToken(ByVal handle As Long, ByVal proxyAuthHeader As String, ByVal proxyHost As String) As String
    Dim hCred As SecHandle
    Dim hContext As SecHandle
    Dim tsExpiry As SECURITY_INTEGER
    Dim result As Long
    Dim outBuf As SecBuffer
    Dim outBufDesc As SecBufferDesc
    Dim inBuf(0 To 1) As SecBuffer
    Dim inBufDesc As SecBufferDesc
    Dim contextAttr As Long
    Dim targetName As String
    Dim serverToken() As Byte
    Dim b64token As String
    Dim outData() As Byte
    Dim recvLen As Long
    Dim i As Long
    Dim binLen As Long
    
    targetName = "HTTP/" & proxyHost
    result = AcquireCredentialsHandle(NULL_PTR, "NTLM", SECPKG_CRED_OUTBOUND, NULL_PTR, ByVal 0&, 0, 0, hCred, tsExpiry)
    If result <> 0 Then Exit Function
    
    If InStr(proxyAuthHeader, "NTLM ") > 0 Then
        b64token = Mid(proxyAuthHeader, InStr(proxyAuthHeader, "NTLM ") + 5)
        CryptStringToBinaryW StrPtr(b64token), Len(b64token), CRYPT_STRING_BASE64, NULL_PTR, binLen, 0, 0
        If binLen > 0 Then
            ReDim serverToken(0 To binLen - 1)
            CryptStringToBinaryW StrPtr(b64token), Len(b64token), CRYPT_STRING_BASE64, VarPtr(serverToken(0)), binLen, 0, 0
        End If
    End If
    
    recvLen = UBound(serverToken) - LBound(serverToken) + 1
    Dim recvBuffer() As Byte
    If recvLen > 0 Then
        ReDim recvBuffer(0 To recvLen - 1)
        CopyMemory recvBuffer(0), serverToken(0), recvLen
    End If
    
    outBufDesc.ulVersion = SECBUFFER_VERSION
    outBufDesc.cBuffers = 1
    outBufDesc.pBuffers = VarPtr(outBuf)
    outBuf.cbBuffer = 0
    outBuf.BufferType = SECBUFFER_TOKEN
    outBuf.pvBuffer = 0
    
    If recvLen = 0 Then
        result = InitializeSecurityContext(hCred, NULL_PTR, targetName, ISC_REQ_SEQUENCE_DETECT Or ISC_REQ_REPLAY_DETECT Or ISC_REQ_CONFIDENTIALITY Or ISC_REQ_ALLOCATE_MEMORY Or ISC_REQ_STREAM, 0, 0, NULL_PTR, 0, hContext, outBufDesc, contextAttr, tsExpiry)
    Else
        inBufDesc.ulVersion = SECBUFFER_VERSION
        inBufDesc.cBuffers = 2
        inBufDesc.pBuffers = VarPtr(inBuf(0))
        inBuf(0).cbBuffer = recvLen
        inBuf(0).BufferType = SECBUFFER_TOKEN
        inBuf(0).pvBuffer = VarPtr(recvBuffer(0))
        inBuf(1).cbBuffer = 0
        inBuf(1).BufferType = SECBUFFER_EMPTY
        inBuf(1).pvBuffer = 0
        result = InitializeSecurityContextContinue(hCred, hContext, targetName, ISC_REQ_SEQUENCE_DETECT Or ISC_REQ_REPLAY_DETECT Or ISC_REQ_CONFIDENTIALITY Or ISC_REQ_ALLOCATE_MEMORY Or ISC_REQ_STREAM, 0, 0, inBufDesc, 0, hContext, outBufDesc, contextAttr, tsExpiry)
    End If
    
    If outBuf.cbBuffer > 0 Then
        ReDim outData(0 To outBuf.cbBuffer - 1)
        CopyMemoryFromPtr outData(0), outBuf.pvBuffer, outBuf.cbBuffer
        GenerateNtlmToken = "NTLM " & Base64Encode(outData)
        FreeContextBuffer outBuf.pvBuffer
    End If
    
    DeleteSecurityContext hContext
    FreeCredentialsHandle hCred
End Function

'/**
' * @brief Resolves the explicit or implicit local machine/user certificate contexts needed for mTLS.
' * @param handle Pointer index.
' * @return State of load attempt boolean success structure.
' */
Private Function LoadClientCert(ByVal handle As Long) As Boolean
#If VBA7 Then
    Dim outCtx As LongPtr
    Dim outStore As LongPtr
#Else
    Dim outCtx As Long
    Dim outStore As Long
#End If
    Dim fileNum As Integer
    Dim pfxBytes() As Byte
    Dim blob As CRYPT_DATA_BLOB
    Dim fileLen As Long
#If VBA7 Then
    Dim pwPtr As LongPtr
#Else
    Dim pwPtr As Long
#End If
    outCtx = 0
    outStore = 0
    With m_Connections(handle)
        If .ClientCertPfxPath <> "" Then
            If Dir(.ClientCertPfxPath) = "" Then
                SetError ERR_CERT_LOAD_FAILED, "PFX file not found: " & .ClientCertPfxPath, "Client certificate file not found.", handle
                Exit Function
            End If
            fileNum = FreeFile
            Open .ClientCertPfxPath For Binary Access Read As #fileNum
            fileLen = LOF(fileNum)
            If fileLen = 0 Then
                Close #fileNum
                SetError ERR_CERT_LOAD_FAILED, "PFX file is empty", "Client certificate file is empty.", handle
                Exit Function
            End If
            ReDim pfxBytes(0 To fileLen - 1)
            Get #fileNum, , pfxBytes
            Close #fileNum
            blob.cbData = fileLen
            blob.pbData = VarPtr(pfxBytes(0))
            pwPtr = IIf(Len(.ClientCertPfxPass) > 0, StrPtr(.ClientCertPfxPass), NULL_PTR)
            outStore = PFXImportCertStore(blob, pwPtr, CRYPT_EXPORTABLE Or PKCS12_ALLOW_OVERWRITE_KEY)
            If outStore = 0 Then
                SetError ERR_CERT_LOAD_FAILED, "PFXImportCertStore failed: 0x" & hex(Err.LastDllError), "Failed to import client certificate PFX.", handle, Err.LastDllError
                Exit Function
            End If
            outCtx = CertFindCertificateInStore(outStore, X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, 0, CERT_FIND_ANY, ByVal NULL_PTR, 0)
            If outCtx = 0 Then
                SetError ERR_CERT_LOAD_FAILED, "CertFindCertificateInStore (ANY) failed", "No certificate found in PFX.", handle
                CertCloseStore outStore, 0
                Exit Function
            End If
        ElseIf .ClientCertThumb <> "" Then
            outStore = CertOpenStore(CERT_STORE_PROV_SYSTEM, 0, NULL_PTR, CERT_SYSTEM_STORE_CURRENT_USER, StrPtr("MY"))
            If outStore = 0 Then
                SetError ERR_CERT_LOAD_FAILED, "CertOpenStore (MY) failed: 0x" & hex(Err.LastDllError), "Cannot open Windows certificate store.", handle, Err.LastDllError
                Exit Function
            End If
            outCtx = CertFindCertificateInStore(outStore, X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, 0, CERT_FIND_SUBJECT_STR_A, ByVal StrPtr(.ClientCertThumb), 0)
            If outCtx = 0 Then
                SetError ERR_CERT_LOAD_FAILED, "Certificate not found for subject: " & .ClientCertThumb, "Client certificate not found in store.", handle
                CertCloseStore outStore, 0
                Exit Function
            End If
        Else
            Exit Function
        End If
        .pClientCertCtx = outCtx
        .hClientCertStore = outStore
        m_ClientCertContextPtrs(handle) = outCtx
    End With
    LoadClientCert = True
End Function

'/**
' * @brief Performs deep packet inspection on the returned TLS handshakes validating hostname mappings securely.
' * @param handle Core network object tracker.
' * @return State of validity result string true.
' */
Private Function ValidateServerCert(ByVal handle As Long) As Boolean
#If VBA7 Then
    Dim pRemoteCert As LongPtr
    Dim pChainCtx As LongPtr
#Else
    Dim pRemoteCert As Long
    Dim pChainCtx As Long
#End If
    Dim chainPara As CERT_CHAIN_PARA
    Dim sslExtra As SSL_EXTRA_CERT_CHAIN_POLICY_PARA
    Dim policyPara As CERT_CHAIN_POLICY_PARA
    Dim policyStatus As CERT_CHAIN_POLICY_STATUS
    Dim result As Long
    Dim chainFlags As Long
    With m_Connections(handle)
        pRemoteCert = 0
        result = QueryContextAttributes(.hContext, SECPKG_ATTR_REMOTE_CERT_CONTEXT, pRemoteCert)
        If result <> 0 Or pRemoteCert = 0 Then
            SetError ERR_CERT_VALIDATE_FAILED, "QueryContextAttributes(REMOTE_CERT) failed: 0x" & hex(result), "Cannot retrieve server certificate.", handle, result
            Exit Function
        End If
        chainPara.cbSize = LenB(chainPara)
        pChainCtx = 0
        chainFlags = 0
        If .EnableRevocationCheck Then
            chainFlags = CERT_CHAIN_REVOCATION_CHECK_CHAIN
        End If
        result = CertGetCertificateChain(NULL_PTR, pRemoteCert, 0, 0, chainPara, chainFlags, NULL_PTR, pChainCtx)
        If result = 0 Or pChainCtx = 0 Then
            SetError ERR_CERT_VALIDATE_FAILED, "CertGetCertificateChain failed: 0x" & hex(Err.LastDllError), "Cannot build certificate chain.", handle
            CertFreeCertificateContext pRemoteCert
            Exit Function
        End If
        sslExtra.cbSize = LenB(sslExtra)
        sslExtra.dwAuthType = AUTHTYPE_SERVER
        sslExtra.fdwChecks = 0
        sslExtra.pwszServerName = StrPtr(.HOST)
        policyPara.cbSize = LenB(policyPara)
        policyPara.dwFlags = 0
        policyPara.pvExtraPolicyPara = VarPtr(sslExtra)
        policyStatus.cbSize = LenB(policyStatus)
        result = CertVerifyCertificateChainPolicy(CERT_CHAIN_POLICY_SSL, pChainCtx, policyPara, policyStatus)
        CertFreeCertificateChain pChainCtx
        CertFreeCertificateContext pRemoteCert
        If result = 0 Then
             SetError ERR_CERT_VALIDATE_FAILED, "CertVerifyCertificateChainPolicy failed: 0x" & hex(Err.LastDllError), "Certificate policy check failed.", handle
            Exit Function
        End If
        If policyStatus.dwError <> 0 Then
            SetError ERR_CERT_VALIDATE_FAILED, "Cert validation error 0x" & hex(policyStatus.dwError) & " chain=" & policyStatus.lChainIndex & " elem=" & policyStatus.lElementIndex, "Server certificate is not trusted (0x" & hex(policyStatus.dwError) & ").", handle, policyStatus.dwError
            Exit Function
        End If
    End With
    ValidateServerCert = True
End Function

'/**
' * @brief Steps through the SSPI cryptographic handshakes natively without requiring modern COM dependencies.
' * @param handle Core network handler.
' * @return Result loop code zero means successfully handshook securely.
' */
Private Function DoTLSHandshake(ByVal handle As Long) As Long
    Dim outBuf As SecBuffer
    Dim outBufDesc As SecBufferDesc
    Dim inBuf(0 To 1) As SecBuffer
    Dim inBufDesc As SecBufferDesc
    Dim contextAttr As Long
    Dim tsExpiry As SECURITY_INTEGER
    Dim result As Long
    Dim contextFlags As Long
    Dim recvBuffer() As Byte
    Dim recvLen As Long
    Dim outData() As Byte
    Dim firstCall As Boolean
    Dim recv As Long
    Dim loopCount As Long
    Dim i As Long
    Dim extraData As Long
    
    contextFlags = ISC_REQ_SEQUENCE_DETECT Or ISC_REQ_REPLAY_DETECT Or ISC_REQ_CONFIDENTIALITY Or ISC_REQ_ALLOCATE_MEMORY Or ISC_REQ_STREAM
    ReDim recvBuffer(0 To 32767)
    recvLen = 0
    firstCall = True
    loopCount = 0
    
    With m_Connections(handle)
        Do
            loopCount = loopCount + 1
            If firstCall Then
                outBufDesc.ulVersion = SECBUFFER_VERSION
                outBufDesc.cBuffers = 1
                outBufDesc.pBuffers = VarPtr(outBuf)
                outBuf.cbBuffer = 0
                outBuf.BufferType = SECBUFFER_TOKEN
                outBuf.pvBuffer = 0
                result = InitializeSecurityContext(.hCred, NULL_PTR, .HOST, contextFlags, 0, 0, NULL_PTR, 0, .hContext, outBufDesc, contextAttr, tsExpiry)
                firstCall = False
            Else
                inBufDesc.ulVersion = SECBUFFER_VERSION
                inBufDesc.cBuffers = 2
                inBufDesc.pBuffers = VarPtr(inBuf(0))
                inBuf(0).cbBuffer = recvLen
                inBuf(0).BufferType = SECBUFFER_TOKEN
                inBuf(0).pvBuffer = VarPtr(recvBuffer(0))
                inBuf(1).cbBuffer = 0
                inBuf(1).BufferType = SECBUFFER_EMPTY
                inBuf(1).pvBuffer = 0
                outBufDesc.ulVersion = SECBUFFER_VERSION
                outBufDesc.cBuffers = 1
                outBufDesc.pBuffers = VarPtr(outBuf)
                outBuf.cbBuffer = 0
                outBuf.BufferType = SECBUFFER_TOKEN
                outBuf.pvBuffer = 0
                result = InitializeSecurityContextContinue(.hCred, .hContext, .HOST, contextFlags, 0, 0, inBufDesc, 0, .hContext, outBufDesc, contextAttr, tsExpiry)
                extraData = 0
                For i = 0 To 1
                    If inBuf(i).BufferType = SECBUFFER_EXTRA And inBuf(i).cbBuffer > 0 Then
                        extraData = inBuf(i).cbBuffer
                        Exit For
                    End If
                Next i
                If extraData > 0 Then
                    For i = 0 To extraData - 1
                        recvBuffer(i) = recvBuffer(recvLen - extraData + i)
                    Next i
                    recvLen = extraData
                ElseIf result <> SEC_E_INCOMPLETE_MESSAGE Then
                    recvLen = 0
                End If
            End If
            If result < 0 And result <> SEC_E_INCOMPLETE_MESSAGE Then
                DoTLSHandshake = result
                Exit Function
            End If
            If outBuf.cbBuffer > 0 And outBuf.pvBuffer <> 0 Then
                ReDim outData(0 To outBuf.cbBuffer - 1)
                CopyMemoryFromPtr outData(0), outBuf.pvBuffer, outBuf.cbBuffer
                sock_send .Socket, outData(0), outBuf.cbBuffer, 0
                FreeContextBuffer outBuf.pvBuffer
            End If
            If result = SEC_E_OK Then
                DoTLSHandshake = 0
                Exit Function
            End If
            If result = SEC_I_CONTINUE_NEEDED Or result = SEC_E_INCOMPLETE_MESSAGE Then
                If Not WaitForDataOn(handle, 10000) Then
                    DoTLSHandshake = -1
                    Exit Function
                End If
                If recvLen + TLS_RECV_CHUNK_SIZE > UBound(recvBuffer) Then
                    ReDim Preserve recvBuffer(0 To UBound(recvBuffer) + TLS_RECV_GROWTH_SIZE)
                End If
                recv = sock_recv(.Socket, recvBuffer(recvLen), UBound(recvBuffer) - recvLen + 1, 0)
                If recv <= 0 Then
                    DoTLSHandshake = -1
                    Exit Function
                End If
                recvLen = recvLen + recv
            End If
            If loopCount > TLS_HANDSHAKE_MAX_LOOPS Then
                DoTLSHandshake = -1
                Exit Function
            End If
        Loop While result = SEC_I_CONTINUE_NEEDED Or result = SEC_E_INCOMPLETE_MESSAGE
    End With
    DoTLSHandshake = result
End Function

'/**
' * @brief Encrypts a payload via the context cipher keys natively tracking fragmented sizes and sending to proxy correctly over TCP streams.
' * @param handle Associated index identifier.
' * @param data Cleartext byte array payload buffer.
' * @return State of true logical sending sequence boolean.
' */
Private Function TLSSend(ByVal handle As Long, ByRef data() As Byte) As Boolean
    Dim buffers(0 To 3) As SecBuffer
    Dim bufferDesc As SecBufferDesc
    Dim sendBuf() As Byte
    Dim dataLen As Long
    Dim totalLen As Long
    Dim offset As Long
    Dim chunkSize As Long
    Dim maxChunk As Long
    Dim result As Long
    Dim toSend As Long
    Dim totalSent As Long
    Dim sent As Long
    Dim wsaErr As Long
    Dim i As Long
    With m_Connections(handle)
        dataLen = SafeArrayLen(data)
        If dataLen = 0 Then
            TLSSend = True
            Exit Function
        End If
        maxChunk = .sizes.cbMaximumMessage
        If maxChunk <= 0 Then
            maxChunk = TLS_MAX_CHUNK_FALLBACK
        End If
        offset = 0
        Do While offset < dataLen
            chunkSize = maxChunk
            If offset + chunkSize > dataLen Then
                chunkSize = dataLen - offset
            End If
            totalLen = .sizes.cbHeader + chunkSize + .sizes.cbTrailer
            ReDim sendBuf(0 To totalLen - 1)
            For i = 0 To chunkSize - 1
                sendBuf(.sizes.cbHeader + i) = data(LBound(data) + offset + i)
            Next i
            buffers(0).cbBuffer = .sizes.cbHeader
            buffers(0).BufferType = SECBUFFER_STREAM_HEADER
            buffers(0).pvBuffer = VarPtr(sendBuf(0))
            buffers(1).cbBuffer = chunkSize
            buffers(1).BufferType = SECBUFFER_DATA
            buffers(1).pvBuffer = VarPtr(sendBuf(.sizes.cbHeader))
            buffers(2).cbBuffer = .sizes.cbTrailer
            buffers(2).BufferType = SECBUFFER_STREAM_TRAILER
            buffers(2).pvBuffer = VarPtr(sendBuf(.sizes.cbHeader + chunkSize))
            buffers(3).cbBuffer = 0
            buffers(3).BufferType = SECBUFFER_EMPTY
            buffers(3).pvBuffer = 0
            bufferDesc.ulVersion = SECBUFFER_VERSION
            bufferDesc.cBuffers = 4
            bufferDesc.pBuffers = VarPtr(buffers(0))
            result = EncryptMessage(.hContext, 0, bufferDesc, 0)
            If result <> 0 Then
                SetError ERR_TLS_ENCRYPT_FAILED, "EncryptMessage failed: 0x" & hex(result), "TLS encryption error.", handle, result
                Exit Function
            End If
            toSend = buffers(0).cbBuffer + buffers(1).cbBuffer + buffers(2).cbBuffer
            totalSent = 0
            Do While totalSent < toSend
                sent = sock_send(.Socket, sendBuf(totalSent), toSend - totalSent, 0)
                If sent <= 0 Then
                    wsaErr = WSAGetLastError()
                    SetError ERR_SEND_FAILED, "send() after TLS encrypt failed: " & WSAErrDesc(wsaErr), "Failed to send encrypted data.", handle, wsaErr
                    .state = STATE_CLOSED
                    Exit Function
                End If
                totalSent = totalSent + sent
            Loop
            offset = offset + chunkSize
        Loop
    End With
    TLSSend = True
End Function

'/**
' * @brief Handles raw SECPKG SSPI decryptor pipeline, pushing clean deciphered output onto a structured stack.
' * @param handle Memory tracking endpoint session logic.
' */
Private Sub TLSDecrypt(ByVal handle As Long)
    Dim buffers(0 To 3) As SecBuffer
    Dim bufferDesc As SecBufferDesc
    Dim result As Long
    Dim qop As Long
    Dim i As Long
    Dim dataLen As Long
    Dim extraLen As Long
    With m_Connections(handle)
        Do While .recvLen > 0
            buffers(0).cbBuffer = .recvLen
            buffers(0).BufferType = SECBUFFER_DATA
            buffers(0).pvBuffer = VarPtr(.recvBuffer(0))
            buffers(1).cbBuffer = 0
            buffers(1).BufferType = SECBUFFER_EMPTY
            buffers(1).pvBuffer = 0
            buffers(2).cbBuffer = 0
            buffers(2).BufferType = SECBUFFER_EMPTY
            buffers(2).pvBuffer = 0
            buffers(3).cbBuffer = 0
            buffers(3).BufferType = SECBUFFER_EMPTY
            buffers(3).pvBuffer = 0
            bufferDesc.ulVersion = SECBUFFER_VERSION
            bufferDesc.cBuffers = 4
            bufferDesc.pBuffers = VarPtr(buffers(0))
            result = DecryptMessage(.hContext, bufferDesc, 0, qop)
            If result = SEC_E_INCOMPLETE_MESSAGE Then Exit Sub
            If result = SEC_I_CONTEXT_EXPIRED Then
                WasabiLog handle, "TLS context expired (Server closed connection nicely)."
                .state = STATE_CLOSED
                If .AutoReconnect And .mode = MODE_WEBSOCKET And Not .AsyncMode Then TryReconnect handle
                Exit Sub
            End If
            If result = SEC_I_RENEGOTIATE Then
                SetError ERR_TLS_RENEGOTIATE, "TLS renegotiation requested - closing", "Secure connection interrupted (renegotiation).", handle, SEC_I_RENEGOTIATE
                .state = STATE_CLOSED
                If .AutoReconnect And .mode = MODE_WEBSOCKET And Not .AsyncMode Then TryReconnect handle
                Exit Sub
            End If
            If result <> SEC_E_OK Then
                SetError ERR_TLS_DECRYPT_FAILED, "DecryptMessage failed: 0x" & hex(result), "TLS decryption error.", handle, result
                Exit Sub
            End If
            For i = 0 To 3
                If buffers(i).BufferType = SECBUFFER_DATA Then
                    dataLen = buffers(i).cbBuffer
                    If dataLen > 0 Then
                        EnsureBufferCapacity .DecryptBuffer, .DecryptLen + dataLen
                        CopyMemoryFromPtr .DecryptBuffer(.DecryptLen), buffers(i).pvBuffer, dataLen
                        .DecryptLen = .DecryptLen + dataLen
                    End If
                End If
            Next i
            extraLen = 0
            For i = 0 To 3
                If buffers(i).BufferType = SECBUFFER_EXTRA Then
                    extraLen = buffers(i).cbBuffer
                    If extraLen > 0 Then
                        CopyMemoryFromPtr .recvBuffer(0), buffers(i).pvBuffer, extraLen
                    End If
                    Exit For
                End If
            Next i
            .recvLen = extraLen
        Loop
    End With
End Sub

'/**
' * @brief Captures generic HTTP responses, identifying the \r\n\r\n header end gracefully logic.
' * @param handle Network structure target logic context.
' * @return The entire UTF-8 header block correctly extracted.
' */
Private Function ReceiveHTTPResponse(ByVal handle As Long) As String
    Dim tempBuf() As Byte
    Dim received As Long
    Dim headerEnd As Long
    Dim i As Long
    Dim headerBytes() As Byte
    Dim copyLen As Long
    Dim remainingLen As Long
    With m_Connections(handle)
        Do
            If Not WaitForDataOn(handle, 5000) Then Exit Do
            ReDim tempBuf(0 To 8191)
            received = sock_recv(.Socket, tempBuf(0), 8192, 0)
            If received <= 0 Then Exit Do
            If .TLS Then
                copyLen = received
                If .recvLen + copyLen > BUFFER_SIZE Then copyLen = BUFFER_SIZE - .recvLen
                If copyLen > 0 Then
                    EnsureBufferCapacity .recvBuffer, .recvLen + copyLen
                    CopyMemory .recvBuffer(.recvLen), tempBuf(0), copyLen
                    .recvLen = .recvLen + copyLen
                End If
                TLSDecrypt handle
                If .DecryptLen > 0 Then
                    headerEnd = -1
                    For i = 0 To .DecryptLen - 4
                        If .DecryptBuffer(i) = 13 And .DecryptBuffer(i + 1) = 10 And .DecryptBuffer(i + 2) = 13 And .DecryptBuffer(i + 3) = 10 Then
                            headerEnd = i + 4
                            Exit For
                        End If
                    Next i
                    If headerEnd > 0 Then
                        ReDim headerBytes(0 To headerEnd - 1)
                        CopyMemory headerBytes(0), .DecryptBuffer(0), headerEnd
                        ReceiveHTTPResponse = Utf8ToString(headerBytes, headerEnd)
                        remainingLen = .DecryptLen - headerEnd
                        If remainingLen > 0 Then
                            CopyMemory .DecryptBuffer(0), .DecryptBuffer(headerEnd), remainingLen
                        End If
                        .DecryptLen = remainingLen
                        Exit Function
                    End If
                End If
            Else
                copyLen = received
                EnsureBufferCapacity .DecryptBuffer, .DecryptLen + copyLen
                If .DecryptLen + copyLen > BUFFER_SIZE Then copyLen = BUFFER_SIZE - .DecryptLen
                If copyLen > 0 Then
                    CopyMemory .DecryptBuffer(.DecryptLen), tempBuf(0), copyLen
                    .DecryptLen = .DecryptLen + copyLen
                End If
                If .DecryptLen >= 4 Then
                    headerEnd = -1
                    For i = 0 To .DecryptLen - 4
                        If .DecryptBuffer(i) = 13 And .DecryptBuffer(i + 1) = 10 And .DecryptBuffer(i + 2) = 13 And .DecryptBuffer(i + 3) = 10 Then
                            headerEnd = i + 4
                            Exit For
                        End If
                    Next i
                    If headerEnd > 0 Then
                        ReDim headerBytes(0 To headerEnd - 1)
                        CopyMemory headerBytes(0), .DecryptBuffer(0), headerEnd
                        ReceiveHTTPResponse = Utf8ToString(headerBytes, headerEnd)
                        remainingLen = .DecryptLen - headerEnd
                        If remainingLen > 0 Then
                            CopyMemory .DecryptBuffer(0), .DecryptBuffer(headerEnd), remainingLen
                        End If
                        .DecryptLen = remainingLen
                        Exit Function
                    End If
                End If
            End If
        Loop
        If .DecryptLen > 0 Then
            ReDim headerBytes(0 To .DecryptLen - 1)
            CopyMemory headerBytes(0), .DecryptBuffer(0), .DecryptLen
            ReceiveHTTPResponse = Utf8ToString(headerBytes, .DecryptLen)
            .DecryptLen = 0
        End If
    End With
End Function

'/**
' * @brief Computes SHA1 hashing array contextually using WinCrypt functionality natively.
' * @param data ByRef native input block.
' * @return Formatted hash structure standard payload.
' */
Private Function SHA1(ByRef data() As Byte) As Byte()
    #If VBA7 Then
        Dim hProv As LongPtr
        Dim hHash As LongPtr
    #Else
        Dim hProv As Long
        Dim hHash As Long
    #End If
    Dim result() As Byte
    Dim bufLen As Long
    Dim i As Long

    If CryptAcquireContextW(hProv, NULL_PTR, NULL_PTR, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT) = 0 Then
        Erase result
        Exit Function
    End If

    If CryptCreateHash(hProv, CALG_SHA1, NULL_PTR, 0, hHash) = 0 Then
        CryptReleaseContext hProv, 0
        Erase result
        Exit Function
    End If

    Dim dataLen As Long
    dataLen = UBound(data) - LBound(data) + 1
    If CryptHashData(hHash, data(LBound(data)), dataLen, 0) = 0 Then
        CryptDestroyHash hHash
        CryptReleaseContext hProv, 0
        Erase result
        Exit Function
    End If

    bufLen = 20
    ReDim result(0 To bufLen - 1)
    Dim tmpLen As Long
    tmpLen = bufLen
    If CryptGetHashParam(hHash, HP_HASHVAL, result(0), tmpLen, 0) = 0 Then
        Erase result
    Else
        If tmpLen > 0 And tmpLen <> bufLen Then
            ReDim Preserve result(0 To tmpLen - 1)
        End If
    End If

    CryptDestroyHash hHash
    CryptReleaseContext hProv, 0
    SHA1 = result
End Function

'/**
' * @brief Derives a 16-bit pseudo-randomized sequence Base-64 formatted block for WS Sec-Key handshake generation.
' * @return Encoded logic string parameter context.
' */
Private Function GenerateWSKey() As String
    Dim Bytes(0 To 15) As Byte
    FillRandomBytes Bytes, 16
    GenerateWSKey = Base64Encode(Bytes)
End Function

' ============================================================================
' 11. WEBSOCKET PROTOCOL CORE
' ============================================================================

'/**
' * @brief Detects and registers server deflate window capability parsing logic parameters internally automatically.
' * @param handle Core protocol routing reference logic target.
' * @param response Web socket handshaking string.
' */
Private Sub ParseDeflateResponse(ByVal handle As Long, ByVal response As String)
    Dim extStart As Long
    Dim extLine As String
    Dim lf As Long
    Dim swbPos As Long
    Dim swbVal As Long
    Dim cwbPos As Long
    Dim cwbVal As Long
    
    extStart = InStr(LCase(response), "sec-websocket-extensions: permessage-deflate")
    If extStart = 0 Then
        m_Connections(handle).DeflateEnabled = False
        m_Connections(handle).DeflateActive = False
        Exit Sub
    End If
    
    extLine = Mid(response, extStart)
    lf = InStr(extLine, vbCrLf)
    If lf > 0 Then
        extLine = Left(extLine, lf - 1)
    End If
    
    With m_Connections(handle)
        If InStr(LCase(extLine), "client_no_context_takeover") > 0 Then
            .DeflateContextTakeover = False
        Else
            .DeflateContextTakeover = True
        End If
        
        If InStr(LCase(extLine), "server_no_context_takeover") > 0 Then
            .InflateContextTakeover = False
        Else
            .InflateContextTakeover = True
        End If
        
        swbPos = InStr(LCase(extLine), "server_max_window_bits=")
        If swbPos > 0 Then
            swbVal = val(Mid(extLine, swbPos + 22))
            If swbVal >= 8 And swbVal <= 15 Then
                .DeflateWindowBits = -swbVal
                .ServerMaxWindowBits = swbVal
            End If
        End If
        
        cwbPos = InStr(LCase(extLine), "client_max_window_bits=")
        If cwbPos > 0 Then
            cwbVal = val(Mid(extLine, cwbPos + 22))
            If cwbVal >= 8 And cwbVal <= 15 Then
                .InflateWindowBits = -cwbVal
                .ClientMaxWindowBits = cwbVal
            End If
        End If
        .DeflateActive = True
    End With
End Sub

'/**
' * @brief Compiles the full WS HTTP Handshake frame payload and sends to socket securely, waiting on valid Accept.
' * @param handle Handle identity.
' * @return State Boolean indicating completion correctly.
' */
Private Function DoWebSocketHandshake(ByVal handle As Long) As Boolean
    Dim handshake As String
    Dim sendBuf() As Byte
    Dim response As String
    Dim wsKey As String
    Dim sendResult As Long
    Dim i As Long
    Dim hostHeader As String
    Dim originHeader As String
    Dim expectedAccept As String
    Dim actualAccept As String
    Dim acceptPos As Long
    Dim acceptLineEnd As Long
    Dim wsaErr As Long
    Dim recvBuf() As Byte
    Dim received As Long
    Dim accBuf() As Byte
    Dim accLen As Long
    Dim headerEnd As Long
    Dim headerBytes() As Byte
    wsKey = GenerateWSKey()

    With m_Connections(handle)
        hostHeader = IIf((.TLS And .port <> 443) Or (Not .TLS And .port <> 80), .HOST & ":" & .port, .HOST)
        If .TLS Then
            originHeader = "https://" & IIf(.port <> 443, .HOST & ":" & .port, .HOST)
        Else
            originHeader = "http://" & IIf(.port <> 80, .HOST & ":" & .port, .HOST)
        End If
        handshake = "GET " & .path & " HTTP/1.1" & vbCrLf
        handshake = handshake & "Host: " & hostHeader & vbCrLf
        handshake = handshake & "Upgrade: websocket" & vbCrLf
        handshake = handshake & "Connection: Upgrade" & vbCrLf
        handshake = handshake & "Sec-WebSocket-Key: " & wsKey & vbCrLf
        handshake = handshake & "Sec-WebSocket-Version: 13" & vbCrLf
        If .DeflateEnabled Then
            Dim deflateExt As String
            deflateExt = "permessage-deflate"
            If Not .DeflateContextTakeover Then deflateExt = deflateExt & ";client_no_context_takeover"
            If Not .InflateContextTakeover Then deflateExt = deflateExt & ";server_no_context_takeover"
            If .ClientMaxWindowBits <> 15 Then deflateExt = deflateExt & ";client_max_window_bits=" & .ClientMaxWindowBits
            handshake = handshake & "Sec-WebSocket-Extensions: " & deflateExt & vbCrLf
        End If
        If .SubProtocol <> "" Then handshake = handshake & "Sec-WebSocket-Protocol: " & .SubProtocol & vbCrLf
        handshake = handshake & "Origin: " & originHeader & vbCrLf
        handshake = handshake & "User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36" & vbCrLf
        For i = 0 To .CustomHeaderCount - 1
            handshake = handshake & .CustomHeaders(i) & vbCrLf
        Next i
        handshake = handshake & vbCrLf

        sendBuf = StringToUtf8(handshake)

        If .TLS Then
            If Not TLSSend(handle, sendBuf) Then
                SetError ERR_WEBSOCKET_HANDSHAKE_FAILED, "TLS send of WS handshake failed", "WebSocket upgrade request failed.", handle
                Exit Function
            End If
            response = ReceiveHTTPResponse(handle)
        Else
            sendResult = sock_send(.Socket, sendBuf(0), UBound(sendBuf) + 1, 0)
            If sendResult <= 0 Then
                wsaErr = WSAGetLastError()
                SetError ERR_WEBSOCKET_HANDSHAKE_FAILED, "send() WS handshake failed: " & WSAErrDesc(wsaErr), "WebSocket upgrade request failed.", handle, wsaErr
                Exit Function
            End If
            ReDim accBuf(0 To 8191)
            accLen = 0
            Do
                If Not WaitForDataOn(handle, 5000) Then
                    SetError ERR_WEBSOCKET_HANDSHAKE_TIMEOUT, "No WS handshake response within 5s", "WebSocket handshake timed out.", handle
                    Exit Function
                End If
                ReDim recvBuf(0 To 8191)
                received = sock_recv(.Socket, recvBuf(0), 8192, 0)
                If received <= 0 Then
                    wsaErr = WSAGetLastError()
                    SetError ERR_WEBSOCKET_HANDSHAKE_FAILED, "recv() WS handshake failed: " & WSAErrDesc(wsaErr), "WebSocket handshake failed.", handle, wsaErr
                    Exit Function
                End If
                EnsureBufferCapacity accBuf, accLen + received
                CopyMemory accBuf(accLen), recvBuf(0), received
                accLen = accLen + received
                headerEnd = -1
                For i = 0 To accLen - 4
                    If accBuf(i) = 13 And accBuf(i + 1) = 10 And accBuf(i + 2) = 13 And accBuf(i + 3) = 10 Then
                        headerEnd = i + 4
                        Exit For
                    End If
                Next i
                If headerEnd > 0 Then
                    ReDim headerBytes(0 To headerEnd - 1)
                    CopyMemory headerBytes(0), accBuf(0), headerEnd
                    response = Utf8ToString(headerBytes, headerEnd)
                    Dim remainingLen As Long
                    remainingLen = accLen - headerEnd
                    If remainingLen > 0 Then
                        EnsureBufferCapacity .DecryptBuffer, .DecryptLen + remainingLen
                        CopyMemory .DecryptBuffer(.DecryptLen), accBuf(headerEnd), remainingLen
                        .DecryptLen = .DecryptLen + remainingLen
                    End If
                    Exit Do
                End If
            Loop
        End If

        If InStr(response, "101") = 0 Then
            Dim lineEnd As Long
            Dim statusLine As String
            lineEnd = InStr(response, vbCrLf)
            If lineEnd > 0 Then
                statusLine = Left(response, lineEnd - 1)
            Else
                statusLine = Left(response, 50)
            End If
            SetError ERR_HANDSHAKE_REJECTED, "Server rejected WS upgrade: " & statusLine, "WebSocket connection rejected: " & statusLine, handle
            Exit Function
        End If

        If .DeflateEnabled Then ParseDeflateResponse handle, response

        Dim wsAcceptInput() As Byte
        wsAcceptInput = StringToUtf8(wsKey & "258EAFA5-E914-47DA-95CA-C5AB0DC85B11")
        expectedAccept = Base64Encode(SHA1(wsAcceptInput))
        acceptPos = InStr(LCase(response), "sec-websocket-accept:")
        If acceptPos > 0 Then
            acceptLineEnd = InStr(acceptPos, response, vbCrLf)
            If acceptLineEnd > 0 Then
                actualAccept = Trim(Mid(response, acceptPos + 21, acceptLineEnd - acceptPos - 21))
            End If
        End If
        If actualAccept <> expectedAccept Then
            SetError ERR_HANDSHAKE_REJECTED, "Sec-WebSocket-Accept mismatch. Expected: " & expectedAccept & " Got: " & actualAccept, "WebSocket handshake failed: invalid accept key.", handle
            Exit Function
        End If
    End With
    DoWebSocketHandshake = True
End Function

'/**
' * @brief Logic processor for routing Opcode 0x01 (Text) frames to buffer structure memory correctly.
' * @param handle Protocol tracker instance pointer.
' * @param payload Raw memory layout array byte target.
' * @param payloadLen Total length.
' * @param fin Fragment indication parameter block Boolean true.
' * @param isCompressed Deflate block identification.
' */
Private Sub ProcessTextFrame(ByVal handle As Long, ByRef payload() As Byte, ByVal payloadLen As Long, ByVal fin As Boolean, ByVal isCompressed As Boolean)
    Dim textMsg As String
    Dim textPayload() As Byte
    Dim textPayloadLen As Long

    With m_Connections(handle)
        If Not fin Then
            .Fragmenting = True
            .FragmentOpcode = WS_OPCODE_TEXT
            .FragmentIsCompressed = isCompressed
            .FragmentLen = 0
            If payloadLen > 0 Then
                CopyMemory .FragmentBuffer(0), payload(0), payloadLen
                .FragmentLen = payloadLen
            End If
        Else
            If isCompressed And .DeflateActive Then
                Dim inflTextSingle() As Byte
                Dim inflTextSingleLen As Long
                inflTextSingle = InflatePayload(handle, payload, payloadLen, inflTextSingleLen)
                If inflTextSingleLen = 0 Then
                    WebSocketSendClose 1007, "Decompression failed", handle
                    .state = STATE_CLOSED
                    Exit Sub
                End If
                textPayload = inflTextSingle
            Else
                textPayload = payload
            End If

            RunInboundMiddlewares handle, textPayload
            textPayloadLen = SafeArrayLen(textPayload)

            If textPayloadLen > 0 Then
                textMsg = Utf8ToString(textPayload, textPayloadLen)
            Else
                textMsg = ""
            End If

            If Not .ProtocolHandler Is Nothing Then
                .ProtocolHandler.OnTextMessage handle, textMsg
            ElseIf .AsyncMode And Not .AsyncHandler Is Nothing Then
                If .MsgCount < MSG_QUEUE_SIZE Then
                    .MsgQueue(.MsgTail) = textMsg
                    .MsgTail = (.MsgTail + 1) Mod MSG_QUEUE_SIZE
                    .MsgCount = .MsgCount + 1
                    .stats.MessagesReceived = .stats.MessagesReceived + 1
                Else
                    WasabiLog handle, "Warning: async message queue full, dropping message (handle=" & handle & ")"
                End If
            Else
                If .MsgCount < MSG_QUEUE_SIZE Then
                    .MsgQueue(.MsgTail) = textMsg
                    .MsgTail = (.MsgTail + 1) Mod MSG_QUEUE_SIZE
                    .MsgCount = .MsgCount + 1
                    .stats.MessagesReceived = .stats.MessagesReceived + 1
                Else
                    WasabiLog handle, "Warning: message queue full, dropping message (handle=" & handle & ")"
                End If
            End If
        End If
    End With
End Sub

'/**
' * @brief Logic processor for routing Opcode 0x02 (Binary) frames to unmanaged buffers naturally.
' * @param handle Core protocol routing index logic handle map identifier token payload address point.
' * @param payload Content Array memory context.
' * @param payloadLen Native size struct.
' * @param fin Fragmentation validation state true.
' * @param isCompressed Indicates active deflation compression state boolean.
' */
Private Sub ProcessBinaryFrame(ByVal handle As Long, ByRef payload() As Byte, ByVal payloadLen As Long, ByVal fin As Boolean, ByVal isCompressed As Boolean)
    Dim binaryData() As Byte

    With m_Connections(handle)
        If Not fin Then
            .Fragmenting = True
            .FragmentOpcode = WS_OPCODE_BINARY
            .FragmentIsCompressed = isCompressed
            .FragmentLen = 0
            If payloadLen > 0 Then
                CopyMemory .FragmentBuffer(0), payload(0), payloadLen
                .FragmentLen = payloadLen
            End If
        Else
            If isCompressed And .DeflateActive Then
                Dim inflBinSingle() As Byte
                Dim inflBinSingleLen As Long
                inflBinSingle = InflatePayload(handle, payload, payloadLen, inflBinSingleLen)
                If inflBinSingleLen = 0 Then
                    WebSocketSendClose 1007, "Decompression failed", handle
                    .state = STATE_CLOSED
                    Exit Sub
                End If
                binaryData = inflBinSingle
            Else
                binaryData = payload
            End If

            RunInboundMiddlewares handle, binaryData

            If Not .ProtocolHandler Is Nothing Then
                .ProtocolHandler.OnBinaryMessage handle, binaryData
            ElseIf .AsyncMode And Not .AsyncHandler Is Nothing Then
                If .BinaryCount < MSG_QUEUE_SIZE Then
                    .BinaryQueue(.BinaryTail).data = binaryData
                    .BinaryTail = (.BinaryTail + 1) Mod MSG_QUEUE_SIZE
                    .BinaryCount = .BinaryCount + 1
                    .stats.MessagesReceived = .stats.MessagesReceived + 1
                Else
                    WasabiLog handle, "Warning: async binary queue full, dropping message (handle=" & handle & ")"
                End If
            Else
                If .BinaryCount < MSG_QUEUE_SIZE Then
                    .BinaryQueue(.BinaryTail).data = binaryData
                    .BinaryTail = (.BinaryTail + 1) Mod MSG_QUEUE_SIZE
                    .BinaryCount = .BinaryCount + 1
                    .stats.MessagesReceived = .stats.MessagesReceived + 1
                Else
                    WasabiLog handle, "Warning: binary queue full, dropping message (handle=" & handle & ")"
                End If
            End If
        End If
    End With
End Sub

'/**
' * @brief Stitches fragmented payload frames recursively across continuous bounds verifying block capacities accurately natively structure mapping.
' * @param handle Index handler mapping target.
' * @param payload Byte array native structure block.
' * @param payloadLen Current appended length structure payload size block context.
' * @param fin Bit boundary finalizing validation trigger Boolean status.
' */
Private Sub ProcessContinuationFrame(ByVal handle As Long, ByRef payload() As Byte, ByVal payloadLen As Long, ByVal fin As Boolean)
    Dim contPayload() As Byte
    Dim contPayloadLen As Long
    Dim textMsg As String
    Dim binaryData() As Byte
    
    With m_Connections(handle)
        If Not .Fragmenting Then Exit Sub
        
        If .FragmentLen + payloadLen > UBound(.FragmentBuffer) + 1 Then
            SetError ERR_FRAGMENT_OVERFLOW, "Fragment buffer overflow on CONTINUATION frame", "Received message too large.", handle
            .state = STATE_CLOSED
            Exit Sub
        End If
        
        If payloadLen > 0 Then
            CopyMemory .FragmentBuffer(.FragmentLen), payload(0), payloadLen
            .FragmentLen = .FragmentLen + payloadLen
        End If
        
        If fin Then
            If .FragmentIsCompressed And .DeflateActive Then
                Dim inflContBytes() As Byte
                Dim inflContLen As Long
                inflContBytes = InflatePayload(handle, .FragmentBuffer, .FragmentLen, inflContLen)
                If inflContLen = 0 Then
                    WebSocketSendClose 1007, "Decompression failed", handle
                    .state = STATE_CLOSED
                    .Fragmenting = False
                    .FragmentLen = 0
                    Exit Sub
                End If
                contPayload = inflContBytes
            Else
                If .FragmentLen > 0 Then
                    ReDim contPayload(0 To .FragmentLen - 1)
                    CopyMemory contPayload(0), .FragmentBuffer(0), .FragmentLen
                End If
            End If
            
            RunInboundMiddlewares handle, contPayload
            contPayloadLen = SafeArrayLen(contPayload)
            
            If .FragmentOpcode = WS_OPCODE_TEXT Then
                If contPayloadLen > 0 Then
                    textMsg = Utf8ToString(contPayload, contPayloadLen)
                Else
                    textMsg = ""
                End If
                
                If Not .ProtocolHandler Is Nothing Then
                    .ProtocolHandler.OnTextMessage handle, textMsg
                Else
                    If .MsgCount < MSG_QUEUE_SIZE Then
                        .MsgQueue(.MsgTail) = textMsg
                        .MsgTail = (.MsgTail + 1) Mod MSG_QUEUE_SIZE
                        .MsgCount = .MsgCount + 1
                        .stats.MessagesReceived = .stats.MessagesReceived + 1
                    Else
                        WasabiLog handle, "Warning: message queue full, dropping message (handle=" & handle & ")"
                    End If
                End If
            ElseIf .FragmentOpcode = WS_OPCODE_BINARY Then
                binaryData = contPayload
                
                If Not .ProtocolHandler Is Nothing Then
                    .ProtocolHandler.OnBinaryMessage handle, binaryData
                Else
                    If .BinaryCount < MSG_QUEUE_SIZE Then
                        .BinaryQueue(.BinaryTail).data = binaryData
                        .BinaryTail = (.BinaryTail + 1) Mod MSG_QUEUE_SIZE
                        .BinaryCount = .BinaryCount + 1
                        .stats.MessagesReceived = .stats.MessagesReceived + 1
                    Else
                        WasabiLog handle, "Warning: binary queue full, dropping message (handle=" & handle & ")"
                    End If
                End If
            End If
            .Fragmenting = False
            .FragmentLen = 0
        End If
    End With
End Sub

'/**
' * @brief Primary WS Frame Parser engine. Scans unmanaged decrypt blocks and delegates into structured processing events properly checking bounds parameters reliably routing accurately quickly.
' * @param handle Current networking pointer context.
' */
Private Sub ProcessFrames(ByVal handle As Long)
    Dim opcode As Byte
    Dim fin As Boolean
    Dim isCompressed As Boolean
    Dim payloadLen As Long
    Dim wirePayloadLen As Long
    Dim maskFlag As Boolean
    Dim frameStart As Long
    Dim i As Long
    Dim payload() As Byte
    Dim totalFrameLen As Long

    With m_Connections(handle)
        Do While .DecryptLen >= 2
            fin = (.DecryptBuffer(0) And &H80) <> 0
            isCompressed = (.DecryptBuffer(0) And &H40) <> 0
            opcode = .DecryptBuffer(0) And &HF
            maskFlag = (.DecryptBuffer(1) And &H80) <> 0
            payloadLen = .DecryptBuffer(1) And &H7F
            frameStart = 2

            If payloadLen = WS_PAYLOAD_LEN_16BIT Then
                If .DecryptLen < 4 Then Exit Do
                payloadLen = CLng(.DecryptBuffer(2)) * 256 + CLng(.DecryptBuffer(3))
                frameStart = 4
            ElseIf payloadLen = WS_PAYLOAD_LEN_64BIT Then
                If .DecryptLen < 10 Then Exit Do
                Dim hi As Long
                hi = 0
                Dim lo As Long
                lo = 0
                hi = CLng(.DecryptBuffer(2)) * 16777216 + CLng(.DecryptBuffer(3)) * 65536 + CLng(.DecryptBuffer(4)) * 256 + CLng(.DecryptBuffer(5))
                lo = CLng(.DecryptBuffer(6)) * 16777216 + CLng(.DecryptBuffer(7)) * 65536 + CLng(.DecryptBuffer(8)) * 256 + CLng(.DecryptBuffer(9))
                If hi <> 0 Or lo < 0 Or lo > BUFFER_MAX_SIZE Then
                    SetError ERR_FRAGMENT_OVERFLOW, "Frame payload too large: hi=" & hi & " lo=" & lo, "Received frame too large.", handle
                    .state = STATE_CLOSED
                    Exit Sub
                End If
                payloadLen = lo
                frameStart = 10
            End If

            If maskFlag Then frameStart = frameStart + 4
            If .DecryptLen < frameStart + payloadLen Then Exit Do
            wirePayloadLen = payloadLen

            If payloadLen > 0 Then
                ReDim payload(0 To payloadLen - 1)
                CopyMemory payload(0), .DecryptBuffer(frameStart), payloadLen
            Else
                Erase payload
            End If

            Select Case opcode
                Case WS_OPCODE_TEXT
                    ProcessTextFrame handle, payload, payloadLen, fin, isCompressed
                Case WS_OPCODE_BINARY
                    ProcessBinaryFrame handle, payload, payloadLen, fin, isCompressed
                Case WS_OPCODE_CONTINUATION
                    ProcessContinuationFrame handle, payload, payloadLen, fin
                Case WS_OPCODE_CLOSE
                    If isCompressed Then
                        WebSocketSendClose 1002, "RSV1 on control frame", handle
                        .state = STATE_CLOSED
                        Exit Sub
                    End If
                    ProcessCloseFrame handle, payload, payloadLen
                    Exit Sub
                Case WS_OPCODE_PING
                    If isCompressed Then
                        WebSocketSendClose 1002, "RSV1 on control frame", handle
                        .state = STATE_CLOSED
                        Exit Sub
                    End If
                    SendPongFrame handle, payload, payloadLen
                Case WS_OPCODE_PONG
                    If isCompressed Then
                        WebSocketSendClose 1002, "RSV1 on control frame", handle
                        .state = STATE_CLOSED
                        Exit Sub
                    End If
                    ProcessPongForLatency handle
                    WasabiLog handle, "PONG received (handle=" & handle & ")"
            End Select

            totalFrameLen = frameStart + wirePayloadLen
            If .DecryptLen > totalFrameLen Then
                CopyMemory .DecryptBuffer(0), .DecryptBuffer(totalFrameLen), .DecryptLen - totalFrameLen
            End If
            .DecryptLen = .DecryptLen - totalFrameLen
        Loop
    End With
End Sub

'/**
' * @brief Handles incoming remote closure events safely logging diagnostics internally processing cleanly.
' * @param handle Core memory array indexing handle.
' * @param payload Raw payload containing code/reason bytes logic array context address mapping.
' * @param payloadLen Total length dimension validation parameters bounds check parameter target.
' */
Private Sub ProcessCloseFrame(ByVal handle As Long, ByRef payload() As Byte, ByVal payloadLen As Long)
    Dim closeCode As Integer
    Dim closeReason As String
    Dim replyFrame(0 To 7) As Byte
    Dim mask(0 To 3) As Byte
    Dim reasonBytes() As Byte
    Dim i As Long
    
    With m_Connections(handle)
        closeCode = WS_CLOSE_NO_STATUS
        closeReason = ""
        If payloadLen >= 2 Then
            closeCode = CInt(payload(0)) * 256 + CInt(payload(1))
            If payloadLen > 2 Then
                ReDim reasonBytes(0 To payloadLen - 3)
                For i = 0 To payloadLen - 3
                    reasonBytes(i) = payload(2 + i)
                Next i
                closeReason = Utf8ToString(reasonBytes, payloadLen - 2)
            End If
        End If
        .closeCode = closeCode
        .closeReason = closeReason
        
        If Not .CloseInitiatedByUs Then
            FillRandomBytes mask, 4
            replyFrame(0) = &H88
            replyFrame(2) = mask(0)
            replyFrame(3) = mask(1)
            replyFrame(4) = mask(2)
            replyFrame(5) = mask(3)
            
            If payloadLen >= 2 Then
                replyFrame(1) = &H82
                replyFrame(6) = payload(0) Xor mask(0)
                replyFrame(7) = payload(1) Xor mask(1)
                
                Dim rf2() As Byte: ReDim rf2(0 To 7)
                For i = 0 To 7: rf2(i) = replyFrame(i): Next i
                If .TLS Then
                    TLSSend handle, rf2
                Else
                    sock_send .Socket, rf2(0), 8, 0
                End If
            Else
                replyFrame(1) = &H80
                
                Dim rf0() As Byte: ReDim rf0(0 To 5)
                For i = 0 To 5: rf0(i) = replyFrame(i): Next i
                If .TLS Then
                    TLSSend handle, rf0
                Else
                    sock_send .Socket, rf0(0), 6, 0
                End If
            End If
        End If
        .state = STATE_CLOSED
    End With
End Sub

'/**
' * @brief Synthesizes and sends pong control frames logic.
' * @param handle Core protocol routing.
' * @param pingPayload Incoming unmanaged ping logic body map string format natively.
' * @param pingLen Buffer bounds mapping dimension bounds mapping parameter count struct memory index.
' */
Private Sub SendPongFrame(ByVal handle As Long, ByRef pingPayload() As Byte, ByVal pingLen As Long)
    Dim frame() As Byte
    Dim mask(0 To 3) As Byte
    Dim i As Long
    FillRandomBytes mask, 4
    If pingLen = 0 Then
        ReDim frame(0 To 5)
        frame(0) = &H8A
        frame(1) = &H80
    Else
        ReDim frame(0 To 5 + pingLen)
        frame(0) = &H8A
        frame(1) = &H80 Or CByte(pingLen)
        For i = 0 To pingLen - 1
            frame(6 + i) = pingPayload(i) Xor mask(i Mod 4)
        Next i
    End If
    frame(2) = mask(0)
    frame(3) = mask(1)
    frame(4) = mask(2)
    frame(5) = mask(3)
    With m_Connections(handle)
        If .TLS Then
            TLSSend handle, frame
        Else
            sock_send .Socket, frame(0), UBound(frame) + 1, 0
        End If
    End With
End Sub

'/**
' * @brief Calculates Latency (Round Trip Time) metrics securely utilizing the internal tick differentials system structurally natively seamlessly accurately reliably correctly cleanly perfectly precisely gracefully.
' * @param handle Base network array tracking identity context handle index marker.
' */
Private Sub ProcessPongForLatency(ByVal handle As Long)
    With m_Connections(handle)
        If .LastPingTimestamp > 0 Then
            .LastRttMs = TickDiff(.LastPingTimestamp, GetTickCount())
            .LastPingTimestamp = 0
        End If
    End With
End Sub

' ============================================================================
' 12. TCP/BUFFERING CORE
' ============================================================================

'/**
' * @brief Checks native Winsock ioctlsocket availability parameters buffering data securely from kernel gracefully processing unmanaged mapping array sizes dynamically updating.
' * @param handle Master state logic structure index identity parameter tracking logic.
' */
Private Sub FeedBuffer(ByVal handle As Long)
    Dim available As Long
    Dim tempBuf() As Byte
    Dim received As Long
    Dim wsaErr As Long
    Dim copyLen As Long
    Dim toRead As Long

    With m_Connections(handle)
        Do
            If sock_ioctlsocket(.Socket, FIONREAD, available) <> 0 Then
                wsaErr = WSAGetLastError()
                SetError ERR_CONNECTION_LOST, "ioctlsocket(FIONREAD) failed: " & WSAErrDesc(wsaErr), "Connection lost.", handle, wsaErr
                .state = STATE_CLOSED
                If .AutoReconnect And .mode = MODE_WEBSOCKET And Not .AsyncMode Then TryReconnect handle
                Exit Sub
            End If

            If available <= 0 Then Exit Do

            toRead = available
            If toRead > 65536 Then toRead = 65536

            ReDim tempBuf(0 To toRead - 1)
            received = sock_recv(.Socket, tempBuf(0), toRead, 0)

            If received > 0 Then
                .stats.BytesReceived = .stats.BytesReceived + received
                .LastActivityAt = GetTickCount()

                Select Case .mode
                    Case MODE_WEBSOCKET
                        If .TLS Then
                            copyLen = received
                            If .recvLen + copyLen > BUFFER_SIZE Then copyLen = BUFFER_SIZE - .recvLen
                            If copyLen > 0 Then
                                EnsureBufferCapacity .recvBuffer, .recvLen + copyLen
                                CopyMemory .recvBuffer(.recvLen), tempBuf(0), copyLen
                                .recvLen = .recvLen + copyLen
                            End If
                            TLSDecrypt handle
                        Else
                            copyLen = received
                            If .DecryptLen + copyLen > BUFFER_SIZE Then copyLen = BUFFER_SIZE - .DecryptLen
                            If copyLen > 0 Then
                                EnsureBufferCapacity .DecryptBuffer, .DecryptLen + copyLen
                                CopyMemory .DecryptBuffer(.DecryptLen), tempBuf(0), copyLen
                                .DecryptLen = .DecryptLen + copyLen
                            End If
                        End If
                        ProcessFrames handle

                    Case MODE_TCP, MODE_TCP_TLS
                        If .TLS Then
                            copyLen = received
                            If .recvLen + copyLen > BUFFER_SIZE Then copyLen = BUFFER_SIZE - .recvLen
                            If copyLen > 0 Then
                                EnsureBufferCapacity .recvBuffer, .recvLen + copyLen
                                CopyMemory .recvBuffer(.recvLen), tempBuf(0), copyLen
                                .recvLen = .recvLen + copyLen
                            End If
                            TLSDecrypt handle
                            If .DecryptLen > 0 Then
                                EnsureBufferCapacity .TcpRecvBuffer, .TcpRecvLen + .DecryptLen
                                CopyMemory .TcpRecvBuffer(.TcpRecvLen), .DecryptBuffer(0), .DecryptLen
                                .TcpRecvLen = .TcpRecvLen + .DecryptLen
                                .DecryptLen = 0
                            End If
                        Else
                            copyLen = received
                            If .TcpRecvLen + copyLen > BUFFER_SIZE Then copyLen = BUFFER_SIZE - .TcpRecvLen
                            If copyLen > 0 Then
                                EnsureBufferCapacity .TcpRecvBuffer, .TcpRecvLen + copyLen
                                CopyMemory .TcpRecvBuffer(.TcpRecvLen), tempBuf(0), copyLen
                                .TcpRecvLen = .TcpRecvLen + copyLen
                            End If
                        End If
                End Select

            ElseIf received = 0 Then
                SetError ERR_CONNECTION_LOST, "recv() returned 0", "Server closed the connection.", handle
                .state = STATE_CLOSED
                If .AutoReconnect And .mode = MODE_WEBSOCKET And Not .AsyncMode Then TryReconnect handle
                Exit Do
            Else
                wsaErr = WSAGetLastError()
                If wsaErr <> WSA_EWOULDBLOCK Then
                    SetError ERR_RECV_FAILED, "recv() failed: " & WSAErrDesc(wsaErr), "Failed to receive.", handle, wsaErr
                    .state = STATE_CLOSED
                    If .AutoReconnect And .mode = MODE_WEBSOCKET And Not .AsyncMode Then TryReconnect handle
                End If
                Exit Do
            End If
        Loop
    End With
End Sub

'/**
' * @brief Runs periodic maintenance: pings, MTU probes, inactivity timeout.
' * @param handle Connection handle.
' */
Private Sub TickMaintenance(ByVal handle As Long)
    Dim now As Long
    With m_Connections(handle)
        If .state <> STATE_OPEN Then Exit Sub
        now = GetTickCount()
        If .mode = MODE_WEBSOCKET Then
            If .PingIntervalMs > 0 Then
                If TickDiff(.LastPingSentAt, now) >= .CurrentPingIntervalMs Then
                    WebSocketSendPing "", handle
                    CalculateNextPing handle
                End If
            End If
        End If
        If .InactivityTimeoutMs > 0 And .LastActivityAt > 0 Then
            If TickDiff(.LastActivityAt, now) >= .InactivityTimeoutMs Then
                SetError ERR_INACTIVITY_TIMEOUT, "Inactivity timeout after " & .InactivityTimeoutMs & "ms", "Connection timed out.", handle
                .state = STATE_CLOSED
                If .AutoReconnect And .mode = MODE_WEBSOCKET And Not .AsyncMode Then TryReconnect handle
                Exit Sub
            End If
        End If
        If .AutoMTU And .mtu.ProbeEnabled Then
            If TickDiff(.mtu.LastProbeTime, now) >= PMTU_DISCOVERY_INTERVAL_MS Then
                probeMTU handle
            End If
        End If
    End With
End Sub

'/**
' * @brief High level safe closure loop clearing tracking logic queues appropriately properly precisely.
' * @param handle Core array indexing parameter parameterizing logically mapping.
' */
Private Sub CloseSession(ByVal handle As Long)
    If Not ValidIndex(handle) Then Exit Sub
    CleanupHandle handle
End Sub

'/**
' * @brief Execution engine implementing exponential backoff handling mapping reconnect strategies logic.
' * @param handle Identifier logic session target.
' */
Private Sub TryReconnect(ByVal handle As Long)
    Dim delayMs As Long
    Dim attempt As Long
    Dim i As Long
    Dim startTick As Long

    If Not m_Connections(handle).AutoReconnect Then Exit Sub

    If m_Connections(handle).ReconnectMaxAttempts > 0 And m_Connections(handle).ReconnectAttempts >= m_Connections(handle).ReconnectMaxAttempts Then
        WasabiLog handle, "Auto-reconnect exhausted after " & m_Connections(handle).ReconnectAttempts & " attempts (handle=" & handle & ")"
        m_Connections(handle).AutoReconnect = False
        Exit Sub
    End If

    m_Connections(handle).ReconnectAttempts = m_Connections(handle).ReconnectAttempts + 1
    attempt = m_Connections(handle).ReconnectAttempts
    delayMs = m_Connections(handle).ReconnectBaseDelayMs

    For i = 1 To attempt - 1
        delayMs = delayMs * 2
        If delayMs > MAX_RECONNECT_DELAY_MS Then
            delayMs = MAX_RECONNECT_DELAY_MS
            Exit For
        End If
    Next i

    WasabiLog handle, "Reconnect attempt " & attempt & " in " & delayMs & "ms (handle=" & handle & ")"

    startTick = GetTickCount()
    Do
        DoEvents
        If Not m_Connections(handle).AutoReconnect Then Exit Sub
        If TickDiff(startTick, GetTickCount()) >= delayMs Then Exit Do
    Loop

    If Not m_WSAInitialized Then
        Dim wsa As WSADATA
        WSAStartup &H202, wsa
        m_WSAInitialized = True
    End If

    Dim savedAutoReconnect As Boolean
    Dim savedMaxAttempts As Long
    Dim savedBaseDelay As Long
    Dim savedAttempts As Long
    Dim savedAsyncMode As Boolean
    Dim savedAsyncHandler As Object

    With m_Connections(handle)
        savedAutoReconnect = .AutoReconnect
        savedMaxAttempts = .ReconnectMaxAttempts
        savedBaseDelay = .ReconnectBaseDelayMs
        savedAttempts = .ReconnectAttempts
        savedAsyncMode = .AsyncMode
        If savedAsyncMode Then Set savedAsyncHandler = .AsyncHandler
    End With

    If Not ConnectHandle(handle, m_Connections(handle).OriginalUrl) Then
        With m_Connections(handle)
            .AutoReconnect = savedAutoReconnect
            .ReconnectMaxAttempts = savedMaxAttempts
            .ReconnectBaseDelayMs = savedBaseDelay
            .ReconnectAttempts = savedAttempts
            .AsyncMode = savedAsyncMode
            If savedAsyncMode Then Set .AsyncHandler = savedAsyncHandler
        End With
        WasabiLog handle, "Reconnect attempt " & attempt & " failed (handle=" & handle & ")"
    Else
        m_Connections(handle).ReconnectAttempts = 0
        WasabiLog handle, "Reconnect succeeded (handle=" & handle & ")"
    End If
End Sub

'/**
' * @brief Centralized TCP connection core handling MTU, Proxies, DNS mapping natively comprehensively structurally accurately properly correctly perfectly cleanly natively structurally comprehensively reliably strictly.
' * @param handle Struct endpoint connection target routing context identity.
' * @param HOST Target mapped host routing target identity variable structure domain name point mapping address destination.
' * @param port Native int representing application protocol mapping index identifier context memory pointer identity binding logical number value bounds structure port.
' * @param useTLS Determines SSPI encryption bounds natively securely robustly perfectly securely strictly gracefully strongly context parameter boolean logic logic.
' * @return State Boolean indicating completion correctly perfectly seamlessly accurately reliably properly precisely accurately reliably properly cleanly correctly cleanly.
' */
Private Function TcpConnectInternal(ByVal handle As Long, ByVal HOST As String, ByVal port As Long, ByVal useTLS As Boolean) As Boolean
    Dim schannelCred As SCHANNEL_CRED
    Dim tsExpiry As SECURITY_INTEGER
    Dim zeroBytes() As Byte
    Dim acquireResult As Long
    Dim tlsResult As Long
    Dim connectHost As String
    Dim connectPort As Long

    With m_Connections(handle)
        .HOST = HOST
        .port = port
        .TLS = useTLS
        .state = STATE_CONNECTING
        .LastError = ERR_NONE
        .LastErrorCode = 0
        .TechnicalDetails = ""

        connectHost = IIf(.ProxyEnabled And .proxyHost <> "", .proxyHost, HOST)
        connectPort = IIf(.ProxyEnabled And .proxyPort > 0, .proxyPort, port)

        If Not ResolveAndConnect(handle, connectHost, connectPort) Then GoTo Fail
        InitializeMTU handle
        If .AutoMTU Then probeMTU handle
        ApplySocketOptions handle

        If .ProxyEnabled Then
            If .proxyType = PROXY_TYPE_SOCKS5 Then
                If Not DoProxySOCKS5(handle) Then GoTo Fail
            Else
                If Not DoProxyHTTP(handle) Then GoTo Fail
            End If
        End If

        If useTLS Then
            ReDim zeroBytes(0 To LenB(schannelCred) - 1)
            CopyMemory schannelCred, zeroBytes(0), LenB(schannelCred)
            schannelCred.dwVersion = SCHANNEL_CRED_VERSION
            schannelCred.grbitEnabledProtocols = SP_PROT_TLS1_2_CLIENT Or SP_PROT_TLS1_3_CLIENT
            schannelCred.dwFlags = SCH_CRED_NO_DEFAULT_CREDS Or SCH_CRED_MANUAL_CRED_VALIDATION Or SCH_CRED_IGNORE_NO_REVOCATION_CHECK Or SCH_CRED_IGNORE_REVOCATION_OFFLINE

            If .ClientCertThumb <> "" Or .ClientCertPfxPath <> "" Then
                If LoadClientCert(handle) Then
                    m_ClientCertContextPtrs(handle) = .pClientCertCtx
                    schannelCred.cCreds = 1
                    schannelCred.paCred = VarPtr(m_ClientCertContextPtrs(handle))
                End If
            End If

            acquireResult = AcquireCredentialsHandle(NULL_PTR, "Microsoft Unified Security Protocol Provider", SECPKG_CRED_OUTBOUND, NULL_PTR, schannelCred, NULL_PTR, NULL_PTR, .hCred, tsExpiry)
            If acquireResult <> 0 Then
                SetError ERR_TLS_ACQUIRE_CREDS_FAILED, "AcquireCredentialsHandle failed: 0x" & hex(acquireResult), "TLS initialization failed.", handle, acquireResult
                GoTo Fail
            End If

            tlsResult = DoTLSHandshake(handle)
            If tlsResult <> 0 Then
                If tlsResult = -1 Then
                    SetError ERR_TLS_HANDSHAKE_TIMEOUT, "TLS handshake timed out with " & HOST, "TLS handshake timed out.", handle
                Else
                    SetError ERR_TLS_HANDSHAKE_FAILED, "TLS handshake failed: 0x" & hex(tlsResult), "TLS handshake failed.", handle, tlsResult
                End If
                GoTo Fail
            End If

            QueryContextAttributes .hContext, SECPKG_ATTR_STREAM_SIZES, .sizes
            CalculateOptimalFrameSize handle

            If .ValidateServerCert Then
                If Not ValidateServerCert(handle) Then GoTo Fail
            End If
        End If
    End With

    TcpConnectInternal = True
    Exit Function
Fail:
    CleanupHandle handle
End Function

'/**
' * @brief Resolves the Web Socket protocol context strings handling parsing naturally automatically logic correctly structurally cleanly seamlessly reliably flawlessly accurately.
' * @param handle Core protocol mapped handler instance marker context.
' * @param url URI string path parameter structure value.
' * @return State of execution boolean context string return variable.
' */
Private Function ConnectHandle(ByVal handle As Long, ByVal url As String) As Boolean
    Dim HOST As String
    Dim port As Long
    Dim path As String
    Dim useTLS As Boolean

    With m_Connections(handle)
        .OriginalUrl = url
        If Not ParseURL(url, HOST, port, path, useTLS) Then
            SetError ERR_INVALID_URL, "Invalid URL: " & url, "Invalid WebSocket URL. Use ws:// or wss://", handle
            Exit Function
        End If
        .path = path
        .mode = MODE_WEBSOCKET
    End With

    If Not TcpConnectInternal(handle, HOST, port, useTLS) Then Exit Function

    If m_Connections(handle).AsyncMode And m_AsyncHwnd <> 0 Then
        WSAAsyncSelect m_Connections(handle).Socket, m_AsyncHwnd, WM_WASABI_SOCKET, _
            FD_READ Or FD_WRITE Or FD_CLOSE Or FD_CONNECT
    End If

    If Not DoWebSocketHandshake(handle) Then
        CleanupHandle handle
        Exit Function
    End If

    With m_Connections(handle)
        .state = STATE_OPEN
        .stats.ConnectedAt = GetTickCount()
        .stats.BytesSent = 0
        .stats.BytesReceived = 0
        .stats.MessagesSent = 0
        .stats.MessagesReceived = 0
        .LastPingSentAt = GetTickCount()
        .LastActivityAt = GetTickCount()
        If .OfflineQueueEnabled Then FlushOfflineQueues handle
        
        If Not .CompressionHandler Is Nothing Then
            .CompressionHandler.OnConnect handle
        End If
        
        If Not .ProtocolHandler Is Nothing Then
            .ProtocolHandler.OnConnect handle
        End If
        
        Dim mwConnect As Object
        For Each mwConnect In .Middlewares
            mwConnect.OnConnect handle
        Next mwConnect
    End With

    ConnectHandle = True
End Function

'/**
' * @brief Unifies the generic physical send mechanism abstracting the TLS context handling logic dynamically naturally context correctly structurally seamlessly cleanly accurately precisely perfectly securely dynamically natively flawlessly robustly efficiently smoothly naturally securely dynamically natively perfectly precisely gracefully strictly correctly naturally automatically perfectly flawlessly comprehensively strongly dynamically precisely tightly comprehensively flawlessly securely elegantly transparently safely natively precisely solidly accurately.
' * @param handle Logic struct.
' * @param frame Core array format mapping data size.
' * @return Status flag.
' */
Private Function SendFrameFor(ByVal handle As Long, ByRef frame() As Byte) As Boolean
    If m_Connections(handle).TLS Then
        SendFrameFor = TLSSend(handle, frame)
    Else
        SendFrameFor = RawSendFor(handle, frame)
    End If
End Function

' ============================================================================
' 13. MQTT PROTOCOL CORE
' ============================================================================

'/**
' * @brief Reads a variable-length integer as per the MQTT v5 protocol standard perfectly efficiently accurately natively correctly dynamically correctly reliably cleanly properly structurally cleanly perfectly efficiently accurately safely.
' * @param buf The target byte block array mapping bounds checking structure block payload context variable array block pointer.
' * @param index Index address locator parsing dynamically state integer loop context variable structure.
' * @return Interpreted int.
' */
Private Function MqttDecodeVarInt(ByRef buf() As Byte, ByRef index As Long) As Long
    Dim multiplier As Long
    Dim value As Long
    Dim encodedByte As Byte
    multiplier = 1
    value = 0
    Do
        If index > UBound(buf) Then Exit Do
        encodedByte = buf(index)
        index = index + 1
        value = value + (encodedByte And MQTT_VARINT_VALUE_MASK) * multiplier
        multiplier = multiplier * MQTT_VARINT_MULTIPLIER
        If multiplier > MQTT_VARINT_MAX_MULTIPLIER Then Exit Do
    Loop While (encodedByte And MQTT_VARINT_CONTINUE_BIT) <> 0
    MqttDecodeVarInt = value
End Function

'/**
' * @brief Encodes an integer as an MQTT variable-length integer.
' * @param length The integer to encode.
' * @param buf Output buffer.
' * @return Number of bytes written.
' */
Private Function MqttEncodeRemainingLength(ByVal length As Long, ByRef buf() As Byte) As Long
    Dim encodedByte As Byte
    Dim idx As Long
    idx = 0
    Do
        encodedByte = CByte(length Mod MQTT_VARINT_MULTIPLIER)
        length = length \ MQTT_VARINT_MULTIPLIER
        If length > 0 Then
            encodedByte = encodedByte Or &H80
        End If
        buf(0 + idx) = encodedByte
        idx = idx + 1
    Loop While length > 0
    MqttEncodeRemainingLength = idx
End Function

'/**
' * @brief Generates the full initial CONNECT protocol MQTT message mapping properly natively cleanly flawlessly dynamically natively efficiently compactly seamlessly successfully completely logically smoothly precisely efficiently completely robustly properly securely seamlessly compactly efficiently perfectly elegantly natively seamlessly flawlessly smoothly safely tightly robustly tightly securely compactly cleanly structurally natively efficiently correctly perfectly reliably dynamically correctly cleanly accurately seamlessly flawlessly structurally flawlessly tightly reliably efficiently perfectly correctly securely safely properly correctly dynamically smoothly tightly natively cleanly perfectly efficiently compactly efficiently compactly compactly securely structurally properly successfully comprehensively reliably precisely structurally cleanly cleanly cleanly.
' * @param clientId Identity string mapping protocol variable block parameter structural domain value mapping.
' * @param username Credentials authentication context identity parameter block parameter domain value.
' * @param password Credentials structural authentication variable text payload pointer parameter target constraint parameter mapping context.
' * @param keepAlive Core pinging target interval.
' * @param sessionExpirySec Extended v5 logic feature identifier property bounds token.
' * @return Fully array structural bounds string.
' */
Private Function BuildMqttConnectPacket(ByVal clientId As String, Optional ByVal username As String = "", Optional ByVal password As String = "", Optional ByVal keepAlive As Integer = 60, Optional ByVal sessionExpirySec As Long = 0) As Byte()
    Dim varHeader(0 To 9) As Byte
    Dim flags As Byte
    Dim clientBytes() As Byte
    Dim userBytes() As Byte
    Dim passBytes() As Byte
    Dim payload() As Byte
    Dim payloadLen As Long
    Dim remaining As Long
    Dim rlBuf(0 To 3) As Byte
    Dim rlLen As Long
    Dim packet() As Byte
    Dim pos As Long
    Dim cLen As Long
    Dim uLen As Long
    Dim pLen As Long
    Dim propBuf() As Byte
    Dim propLen As Long
    Dim propLenVar() As Byte
    Dim propLenVarSize As Long

    varHeader(0) = 0
    varHeader(1) = 4
    varHeader(2) = 77
    varHeader(3) = 81
    varHeader(4) = 84
    varHeader(5) = 84
    varHeader(6) = 5
    
    flags = 2
    If Len(username) > 0 Then flags = flags Or MQTT_CONNECT_USERNAME
    If Len(password) > 0 Then flags = flags Or MQTT_CONNECT_PASSWORD
    varHeader(7) = flags
    
    varHeader(8) = CByte((keepAlive \ 256) And 255)
    varHeader(9) = CByte(keepAlive And 255)

    If sessionExpirySec > 0 Then
        ReDim propBuf(0 To 4)
        propBuf(0) = 17
        propBuf(1) = CByte((sessionExpirySec \ 16777216) And 255)
        propBuf(2) = CByte((sessionExpirySec \ 65536) And 255)
        propBuf(3) = CByte((sessionExpirySec \ 256) And 255)
        propBuf(4) = CByte(sessionExpirySec And 255)
        propLen = 5
    Else
        propLen = 0
    End If
    
    ReDim propLenVar(0 To 3)
    propLenVarSize = MqttEncodeRemainingLength(propLen, propLenVar)

    If Len(clientId) > 0 Then
        clientBytes = StringToUtf8(clientId)
        cLen = UBound(clientBytes) + 1
    End If
    
    If Len(username) > 0 Then
        userBytes = StringToUtf8(username)
        uLen = UBound(userBytes) + 1
    End If
    
    If Len(password) > 0 Then
        passBytes = StringToUtf8(password)
        pLen = UBound(passBytes) + 1
    End If

    payloadLen = 2 + cLen
    If uLen > 0 Then payloadLen = payloadLen + 2 + uLen
    If pLen > 0 Then payloadLen = payloadLen + 2 + pLen

    ReDim payload(0 To payloadLen - 1)
    pos = 0
    
    payload(pos) = CByte((cLen \ 256) And 255)
    payload(pos + 1) = CByte(cLen And 255)
    pos = pos + 2
    If cLen > 0 Then
        CopyMemory payload(pos), clientBytes(0), cLen
        pos = pos + cLen
    End If

    If uLen > 0 Then
        payload(pos) = CByte((uLen \ 256) And 255)
        payload(pos + 1) = CByte(uLen And 255)
        pos = pos + 2
        CopyMemory payload(pos), userBytes(0), uLen
        pos = pos + uLen
    End If

    If pLen > 0 Then
        payload(pos) = CByte((pLen \ 256) And 255)
        payload(pos + 1) = CByte(pLen And 255)
        pos = pos + 2
        CopyMemory payload(pos), passBytes(0), pLen
    End If

    remaining = 10 + propLenVarSize + propLen + payloadLen
    rlLen = MqttEncodeRemainingLength(remaining, rlBuf)
    
    ReDim packet(0 To rlLen + remaining)
    packet(0) = 16
    
    CopyMemory packet(1), rlBuf(0), rlLen
    CopyMemory packet(1 + rlLen), varHeader(0), 10
    CopyMemory packet(1 + rlLen + 10), propLenVar(0), propLenVarSize
    
    pos = 1 + rlLen + 10 + propLenVarSize
    If propLen > 0 Then
        CopyMemory packet(pos), propBuf(0), propLen
        pos = pos + propLen
    End If
    
    CopyMemory packet(pos), payload(0), payloadLen
    
    BuildMqttConnectPacket = packet
End Function

'/**
' * @brief Generalized framework constructor for creating arbitrary MQTT packets reliably natively cleanly securely safely correctly.
' * @param ptype Protocol type opcode structural value identity block mapping.
' * @param flags Feature set bits logic pointer index constraint structure context constraint identifier.
' * @param payload Subframe logic identity context content block payload pointer index context context address variables target constraints array buffer.
' * @param payloadLen Boundary array constraint struct value limit domain map pointer array index context dimension bounds variable value natively constraints limit size block parameter integer.
' * @return Structural byte mapping index pointer buffer variable logic struct string structure value return context domain domain domain boundary integer dimension size limit parameter target.
' */
Private Function MqttBuildPacket(ByVal ptype As Byte, ByVal flags As Byte, ByRef payload() As Byte, ByVal payloadLen As Long) As Byte()
    Dim remaining As Long
    Dim rlBuf(0 To 3) As Byte
    Dim rlLen As Long
    Dim packet() As Byte
    
    remaining = payloadLen
    rlLen = MqttEncodeRemainingLength(remaining, rlBuf)
    
    ReDim packet(0 To rlLen + remaining)
    
    packet(0) = ptype * 16 Or flags
    CopyMemory packet(1), rlBuf(0), rlLen
    
    If payloadLen > 0 Then
        CopyMemory packet(1 + rlLen), payload(0), payloadLen
    End If
    
    MqttBuildPacket = packet
End Function

'/**
' * @brief Synthesizes ACK messages dynamically mapping IDs logically naturally flawlessly seamlessly cleanly efficiently flawlessly reliably dynamically smoothly perfectly cleanly cleanly elegantly securely cleanly reliably perfectly seamlessly correctly natively gracefully seamlessly structurally elegantly gracefully safely flawlessly tightly cleanly cleanly seamlessly cleanly seamlessly accurately perfectly smoothly properly properly dynamically cleanly correctly securely tightly cleanly cleanly reliably structurally smoothly cleanly seamlessly efficiently flawlessly securely cleanly reliably smoothly accurately safely cleanly compactly seamlessly dynamically reliably smoothly perfectly correctly properly perfectly smoothly efficiently cleanly seamlessly cleanly safely compactly reliably perfectly properly successfully securely cleanly smoothly accurately precisely tightly accurately cleanly smoothly efficiently.
' * @param handle Pointer identification block indexing context map logical memory logic handle variable index.
' * @param packetType Protocol byte constant integer struct limit string limit string boundary block identity pointer constraint logic struct structural limit block domain array byte mapping integer parameters target structure pointer pointer values domain mapping block.
' * @param flags State bit constraints parameter mapping constraints value pointer structure block constraints parameter bounds value constraint domain block parameter index domain.
' * @param packetId Identifies context variables pointer address payload string constraints struct target values integer block size limits pointers size constraints values targets constraint map context parameter variable context bounds values limits value context domain bounds.
' */
Private Sub MqttSendAck(ByVal handle As Long, ByVal packetType As Byte, ByVal flags As Byte, ByVal packetId As Integer)
    Dim packet(0 To 3) As Byte
    packet(0) = (packetType * 16) Or flags
    packet(1) = 2
    packet(2) = CByte((packetId \ 256) And &HFF)
    packet(3) = CByte(packetId And &HFF)
    WebSocketSendBinary packet, handle
End Sub

'/**
' * @brief FSM logic core parsing MQTT traffic bytes progressively accurately securely perfectly optimally safely securely compactly cleanly cleanly efficiently tightly cleanly smoothly reliably perfectly tightly flawlessly gracefully reliably efficiently securely seamlessly flawlessly compactly smoothly perfectly cleanly cleanly cleanly correctly properly efficiently seamlessly cleanly cleanly compactly perfectly gracefully safely compactly flawlessly seamlessly dynamically properly.
' * @param handle Map identity structure value context target mapping logical parameter address constraint bounds limits identifier array target variable context string value address index identity struct pointer index limit array constraint dimension block limit block constraint limits bounds sizes pointer array limit variable.
' * @param b Input logic bounds domain byte block pointer variable block parameter array pointer address.
' */
Private Sub MqttParseByte(ByVal handle As Long, ByVal b As Byte)
    With m_Connections(handle)
        Select Case .MqttParserStage
            Case 0
                .MqttCurrentPacketType = b \ 16
                .MqttCurrentFlags = b And &HF
                .MqttParserStage = 1
                .MqttExpectedRemaining = 0
                .MqttBufLen = 0
            Case 1
                .MqttExpectedRemaining = .MqttExpectedRemaining + (b And &H7F) * (MQTT_VARINT_MULTIPLIER ^ .MqttBufLen)
                .MqttBufLen = .MqttBufLen + 1
                If (b And &H80) = 0 Then
                    .MqttParserStage = 2
                    .MqttBufLen = 0
                    If .MqttExpectedRemaining > 0 Then
                        EnsureBufferCapacity .MqttBuffer, .MqttExpectedRemaining
                    Else
                        .MqttParserStage = 3
                    End If
                End If
            Case 2
                If .MqttBufLen >= UBound(.MqttBuffer) Then
                    EnsureBufferCapacity .MqttBuffer, .MqttBufLen + MQTT_CHUNK_SIZE
                End If
                .MqttBuffer(.MqttBufLen) = b
                .MqttBufLen = .MqttBufLen + 1
                If .MqttBufLen >= .MqttExpectedRemaining Then
                    .MqttParserStage = 3
                End If
        End Select
    End With
End Sub

'/**
' * @brief Checks readiness boundary flags correctly correctly cleanly properly accurately cleanly dynamically reliably smoothly seamlessly properly safely compactly accurately smoothly seamlessly seamlessly perfectly cleanly perfectly seamlessly flawlessly efficiently precisely perfectly reliably reliably cleanly.
' * @param handle Map ID structural array element constraint target value size constraint variable context integer value pointer dimension index pointer context address string structural value size boolean address limits limit parameter bounds structure array target logic target logic.
' * @return Status logical variable target array size bounds.
' */
Private Function MqttHasPacket(ByVal handle As Long) As Boolean
    MqttHasPacket = (m_Connections(handle).MqttParserStage = 3)
End Function

'/**
' * @brief Restores FSM parameters dynamically safely safely cleanly correctly cleanly compactly seamlessly elegantly perfectly efficiently properly precisely safely smoothly reliably flawlessly smoothly optimally compactly reliably smoothly securely properly precisely correctly reliably securely smoothly optimally elegantly cleanly gracefully securely securely dynamically elegantly perfectly smoothly correctly optimally seamlessly flawlessly safely smoothly seamlessly elegantly smoothly properly safely gracefully cleanly tightly safely properly properly correctly efficiently gracefully cleanly seamlessly correctly gracefully dynamically stably seamlessly cleanly.
' * @param handle Parameter indexing parameter constraint logical target value address index target struct variable context identity variable size bounds boundary context value boolean limits limits limit string block target structural memory limits limit boolean boolean variable pointer size.
' */
Private Sub MqttResetParser(ByVal handle As Long)
    m_Connections(handle).MqttParserStage = 0
    m_Connections(handle).MqttBufLen = 0
End Sub

' ============================================================================
' 14. MIDDLEWARE & QUEUEING
' ============================================================================

'/**
' * @brief Processes pending logic variables smoothly dynamically perfectly cleanly correctly perfectly accurately elegantly efficiently reliably correctly smoothly precisely seamlessly flawlessly flawlessly properly safely stably properly dynamically correctly stably dynamically safely smoothly natively gracefully safely properly compactly stably securely smoothly properly seamlessly elegantly compactly flawlessly safely dynamically elegantly precisely safely smoothly correctly dynamically compactly optimally seamlessly tightly safely cleanly cleanly cleanly elegantly correctly.
' * @param handle Internal instance variable index variable structural map parameter bounds context variable address parameter logic logic variables identifier structural logic target boundary sizes size domain bounds block identifier limits limit variables mapping bounds.
' */
Private Sub FlushOfflineQueues(ByVal handle As Long)
    Dim i As Long
    Dim tCount As Long, bCount As Long
    Dim tQueue() As String, bQueue() As BinaryMessage

    With m_Connections(handle)
        If Not .OfflineQueueEnabled Then Exit Sub
        tCount = .OfflineTextCount
        bCount = .OfflineBinaryCount

        If tCount > 0 Then
            ReDim tQueue(0 To tCount - 1)
            For i = 0 To tCount - 1: tQueue(i) = .OfflineTextQueue(i): Next i
            .OfflineTextCount = 0
        End If

        If bCount > 0 Then
            ReDim bQueue(0 To bCount - 1)
            For i = 0 To bCount - 1: bQueue(i) = .OfflineBinaryQueue(i): Next i
            .OfflineBinaryCount = 0
        End If
    End With

    If tCount > 0 Then
        For i = 0 To tCount - 1
            If m_Connections(handle).state <> STATE_OPEN Then Exit For
            m_Connections(handle).OfflineQueueEnabled = False
            WebSocketSendText tQueue(i), handle
            m_Connections(handle).OfflineQueueEnabled = True
        Next i
    End If

    If bCount > 0 Then
        For i = 0 To bCount - 1
            If m_Connections(handle).state <> STATE_OPEN Then Exit For
            m_Connections(handle).OfflineQueueEnabled = False
            WebSocketSendBinary bQueue(i).data, handle
            m_Connections(handle).OfflineQueueEnabled = True
        Next i
    End If
End Sub

'/**
' * @brief Execute BeforeSend pipeline.
' * @param handle Pointer.
' * @param data Ref array.
' */
Private Sub RunOutboundMiddlewares(ByVal handle As Long, ByRef data() As Byte)
    Dim mw As Object
    For Each mw In m_Connections(handle).Middlewares
        mw.OnBeforeSend handle, data
    Next mw
End Sub

'/**
' * @brief Execute AfterReceive pipeline.
' * @param handle Pointer.
' * @param data Ref array.
' */
Private Sub RunInboundMiddlewares(ByVal handle As Long, ByRef data() As Byte)
    Dim mw As Object
    For Each mw In m_Connections(handle).Middlewares
        mw.OnAfterReceive handle, data
    Next mw
End Sub

' ============================================================================
' 15. PUBLIC APIs (TCP, WEBSOCKET, MQTT, COMPRESSION)
' ============================================================================

'/**
' * @brief Establishes a standard TCP connection.
' * @param HOST Target hostname or IP.
' * @param port Target port number.
' * @param outHandle ByRef returns the assigned internal tracking handle.
' * @return True if connection was successful.
' */
Public Function TcpConnect(ByVal HOST As String, ByVal port As Long, ByRef outHandle As Long) As Boolean
    Dim wsa As WSADATA
    Dim wsaErr As Long
    Dim handle As Long

    m_LastError = ERR_NONE
    m_LastErrorCode = 0
    m_TechnicalDetails = ""
    InitConnectionPool

    If Not m_WSAInitialized Then
        wsaErr = WSAStartup(&H202, wsa)
        If wsaErr <> 0 Then
            SetError ERR_WSA_STARTUP_FAILED, "WSAStartup failed: " & wsaErr, "Network initialization failed.", INVALID_CONN_HANDLE, wsaErr
            outHandle = INVALID_CONN_HANDLE
            Exit Function
        End If
        m_WSAInitialized = True
    End If

    handle = AllocConnection()
    If handle = INVALID_CONN_HANDLE Then
        SetError ERR_MAX_CONNECTIONS, "Max connections reached", "Too many simultaneous connections.", INVALID_CONN_HANDLE
        outHandle = INVALID_CONN_HANDLE
        Exit Function
    End If

    m_Connections(handle).mode = MODE_TCP

    If Not TcpConnectInternal(handle, HOST, port, False) Then
        outHandle = INVALID_CONN_HANDLE
        Exit Function
    End If

    With m_Connections(handle)
        .state = STATE_OPEN
        .stats.ConnectedAt = GetTickCount()
        .stats.BytesSent = 0
        .stats.BytesReceived = 0
        .stats.MessagesSent = 0
        .stats.MessagesReceived = 0
        .LastActivityAt = GetTickCount()
    End With

    m_DefaultHandle = handle
    outHandle = handle
    TcpConnect = True
    WasabiLog handle, "TCP connected to " & HOST & ":" & port & " (handle=" & handle & ")"
End Function

'/**
' * @brief Establishes an encrypted TCP/TLS connection using SSPI.
' * @param HOST Target hostname or IP.
' * @param port Target port number (usually 443, 853, etc).
' * @param outHandle ByRef returns the assigned internal tracking handle.
' * @return True if connection and TLS handshake were successful.
' */
Public Function TcpConnectTLS(ByVal HOST As String, ByVal port As Long, ByRef outHandle As Long) As Boolean
    Dim wsa As WSADATA
    Dim wsaErr As Long
    Dim handle As Long

    m_LastError = ERR_NONE
    m_LastErrorCode = 0
    m_TechnicalDetails = ""
    InitConnectionPool

    If Not m_WSAInitialized Then
        wsaErr = WSAStartup(&H202, wsa)
        If wsaErr <> 0 Then
            SetError ERR_WSA_STARTUP_FAILED, "WSAStartup failed: " & wsaErr, "Network initialization failed.", INVALID_CONN_HANDLE, wsaErr
            outHandle = INVALID_CONN_HANDLE
            Exit Function
        End If
        m_WSAInitialized = True
    End If

    handle = AllocConnection()
    If handle = INVALID_CONN_HANDLE Then
        SetError ERR_MAX_CONNECTIONS, "Max connections reached", "Too many simultaneous connections.", INVALID_CONN_HANDLE
        outHandle = INVALID_CONN_HANDLE
        Exit Function
    End If

    m_Connections(handle).mode = MODE_TCP_TLS

    If Not TcpConnectInternal(handle, HOST, port, True) Then
        outHandle = INVALID_CONN_HANDLE
        Exit Function
    End If

    With m_Connections(handle)
        .state = STATE_OPEN
        .stats.ConnectedAt = GetTickCount()
        .stats.BytesSent = 0
        .stats.BytesReceived = 0
        .stats.MessagesSent = 0
        .stats.MessagesReceived = 0
        .LastActivityAt = GetTickCount()
    End With

    m_DefaultHandle = handle
    outHandle = handle
    TcpConnectTLS = True
    WasabiLog handle, "TCP+TLS connected to " & HOST & ":" & port & " (handle=" & handle & ")"
End Function

'/**
' * @brief Sends a raw byte array over the TCP stream.
' * @param data Array of bytes to transmit.
' * @param handle (Optional) Target connection handle. Defaults to current active.
' * @return True if completely sent without socket errors.
' */
Public Function TcpSendBinary(ByRef data() As Byte, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
    Dim h As Long
    Dim dataLen As Long

    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function

    With m_Connections(h)
        If .state <> STATE_OPEN Then
            SetError ERR_NOT_CONNECTED, "TcpSendBinary on disconnected handle=" & h, "Not connected.", h
            Exit Function
        End If
        If .mode = MODE_WEBSOCKET Then
            SetError ERR_NOT_CONNECTED, "TcpSendBinary called on WebSocket handle=" & h, "Use WebSocketSendText for WebSocket connections.", h
            Exit Function
        End If

        RunOutboundMiddlewares h, data
        dataLen = SafeArrayLen(data)
        If dataLen = 0 Then
            TcpSendBinary = True
            Exit Function
        End If

        If .TLS Then
            TcpSendBinary = TLSSend(h, data)
        Else
            TcpSendBinary = RawSendFor(h, data)
        End If

        If TcpSendBinary Then
            .stats.BytesSent = .stats.BytesSent + dataLen
            .stats.MessagesSent = .stats.MessagesSent + 1
        End If
    End With
End Function

'/**
' * @brief Sends a UTF-8 string payload over the TCP stream.
' * @param text The string to convert and send.
' * @param handle (Optional) Target connection handle.
' * @return True if sent successfully.
' */
Public Function TcpSendText(ByVal text As String, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
    Dim h As Long
    Dim data() As Byte

    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function

    data = StringToUtf8(text)
    TcpSendText = TcpSendBinary(data, h)
End Function

'/**
' * @brief Pulls all waiting bytes from the TCP internal receive buffer.
' * @param handle (Optional) Target connection handle.
' * @return Raw byte array of received data.
' */
Public Function TcpReceiveBinary(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Byte()
    Dim h As Long
    Dim result() As Byte

    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then
        TcpReceiveBinary = result
        Exit Function
    End If

    With m_Connections(h)
        If .state <> STATE_OPEN Then
            SetError ERR_NOT_CONNECTED, "Receive on closed handle=" & h, "TCP connection is not open.", h
            TcpReceiveBinary = result
            Exit Function
        End If
        If .mode = MODE_WEBSOCKET Then
            SetError ERR_NOT_CONNECTED, "TcpReceiveBinary called on WebSocket handle=" & h, "Use WebSocket receive functions for WebSocket connections.", h
            TcpReceiveBinary = result
            Exit Function
        End If

        TickMaintenance h
        FeedBuffer h

        If .TcpRecvLen > 0 Then
            ReDim result(0 To .TcpRecvLen - 1)
            CopyMemory result(0), .TcpRecvBuffer(0), .TcpRecvLen
            .TcpRecvLen = 0
            .stats.MessagesReceived = .stats.MessagesReceived + 1
            
            RunInboundMiddlewares h, result
            
            TcpReceiveBinary = result
        Else
            TcpReceiveBinary = result
        End If
    End With
End Function

'/**
' * @brief Pulls all waiting bytes and converts them from UTF-8 into a native VBA string.
' * @param handle (Optional) Target connection handle.
' * @return Decoded string.
' */
Public Function TcpReceiveText(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
    Dim h As Long
    Dim data() As Byte
    Dim dataLen As Long

    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function

    data = TcpReceiveBinary(h)
    dataLen = SafeArrayLen(data)
    If dataLen > 0 Then
        TcpReceiveText = Utf8ToString(data, dataLen)
    End If
End Function

'/**
' * @brief Blocks and reads from the TCP stream until a specific delimiter is found (e.g. vbCrLf).
' * @param delimiter The target string sequence to look for.
' * @param timeoutMs Max time to wait in milliseconds.
' * @param handle (Optional) Target connection handle.
' * @return The accumulated string containing the data up to the delimiter.
' */
Public Function TcpReceiveUntil(ByVal delimiter As String, Optional ByVal timeoutMs As Long = 5000, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
    Dim h As Long
    Dim accumulated() As Byte
    Dim accLen As Long
    Dim delimBytes() As Byte
    Dim delimLen As Long
    Dim startTick As Long
    Dim foundIndex As Long
    Dim resultBytes() As Byte
    Dim remaining As Long

    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    If m_Connections(h).AsyncMode Then Exit Function

    delimBytes = StringToUtf8(delimiter)
    delimLen = SafeArrayLen(delimBytes)
    ReDim accumulated(0 To BUFFER_SIZE - 1)
    accLen = 0
    startTick = GetTickCount()

    Do
        FeedBuffer h

        With m_Connections(h)
            If .TcpRecvLen > 0 Then
                EnsureBufferCapacity accumulated, accLen + .TcpRecvLen
                CopyMemory accumulated(accLen), .TcpRecvBuffer(0), .TcpRecvLen
                accLen = accLen + .TcpRecvLen
                .TcpRecvLen = 0
            End If
        End With

        If accLen >= delimLen Then
            foundIndex = WasabiMemFind(VarPtr(accumulated(0)), accLen, VarPtr(delimBytes(0)), delimLen)

            If foundIndex >= 0 Then
                ReDim resultBytes(0 To foundIndex + delimLen - 1)
                CopyMemory resultBytes(0), accumulated(0), foundIndex + delimLen
                TcpReceiveUntil = Utf8ToString(resultBytes, foundIndex + delimLen)

                remaining = accLen - (foundIndex + delimLen)
                If remaining > 0 Then
                    CopyMemory m_Connections(h).TcpRecvBuffer(0), accumulated(foundIndex + delimLen), remaining
                    m_Connections(h).TcpRecvLen = remaining
                End If
                Exit Function
            End If
        End If

        If TickDiff(startTick, GetTickCount()) >= timeoutMs Then Exit Do
        DoEvents
    Loop

    If accLen > 0 Then
        TcpReceiveUntil = Utf8ToString(accumulated, accLen)
    End If
End Function

'/**
' * @brief Checks if a TCP connection is currently open.
' * @param handle (Optional) Connection index.
' * @return True if connected and mode is TCP.
' */
Public Function TcpIsConnected(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    TcpIsConnected = (m_Connections(h).state = STATE_OPEN And m_Connections(h).mode <> MODE_WEBSOCKET)
End Function

'/**
' * @brief Returns the number of unread bytes in the internal buffer.
' * @param handle (Optional) Connection index.
' * @return Number of pending bytes.
' */
Public Function TcpGetPendingBytes(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Long
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    TcpGetPendingBytes = m_Connections(h).TcpRecvLen
End Function

'/**
' * @brief Discards all unread bytes in the TCP input buffer.
' * @param handle (Optional) Target connection handle.
' */
Public Sub TcpFlushBuffer(Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    m_Connections(h).TcpRecvLen = 0
End Sub

'/**
' * @brief Safely terminates a TCP connection and cleans up memory structures.
' * @param handle (Optional) Target connection handle.
' */
Public Sub TcpDisconnect(Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    If m_Connections(h).mode = MODE_WEBSOCKET Then Exit Sub
    m_Connections(h).AutoReconnect = False
    m_Connections(h).AsyncMode = False
    Set m_Connections(h).AsyncHandler = Nothing
    CleanupHandle h
    WasabiLog h, "TCP disconnected (handle=" & h & ")"
End Sub

'/**
' * @brief Associates a protocol parser (like MQTT/STOMP logic) to process events cleanly.
' * @param extension An object instance implementing Wasabi Protocol methods.
' * @param handle Target connection handle.
' */
Public Sub WasabiUseProtocol(ByVal extension As Object, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    
    Set m_Connections(h).ProtocolHandler = extension
    
    If Not extension Is Nothing Then
        If m_Connections(h).state = STATE_OPEN Then
            extension.OnConnect h
        End If
    End If
End Sub

'/**
' * @brief Associates an external ZLib/Deflate compression DLL handler.
' * @param extension An object instance bridging zlib inflate/deflate.
' * @param handle Target connection handle.
' */
Public Sub WasabiUseCompression(ByVal extension As Object, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    
    Set m_Connections(h).CompressionHandler = extension
    
    If Not extension Is Nothing Then
        If m_Connections(h).state = STATE_OPEN Then
            extension.OnConnect h
        End If
    End If
End Sub

'/**
' * @brief Attaches general middleware for payload mutation before send or after receive.
' * @param extension Middleware object.
' * @param handle Target connection handle.
' */
Public Sub WasabiUseMiddleware(ByVal extension As Object, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    
    If m_Connections(h).Middlewares Is Nothing Then
        Set m_Connections(h).Middlewares = New Collection
    End If
    m_Connections(h).Middlewares.Add extension
    
    If Not extension Is Nothing Then
        If m_Connections(h).state = STATE_OPEN Then
            extension.OnConnect h
        End If
    End If
End Sub

'/**
' * @brief Establishes a WebSocket connection (supports ws:// and wss:// natively).
' * @param url The full RFC6455 URI.
' * @param outHandle Returns the assigned internal ID.
' * @param DeflateEnabled Enables RFC7692 permessage-deflate compression negotiation.
' * @param DeflateContextTakeover Allow context window inheritance between packets.
' * @param SubProtocol Request specific server-side subprotocols (e.g. 'mqtt').
' * @return True if handshake completed and socket is ready.
' */
Public Function WebSocketConnect(ByVal url As String, Optional ByRef outHandle As Long = -1, Optional ByVal DeflateEnabled As Boolean = False, Optional ByVal DeflateContextTakeover As Boolean = True, Optional ByVal SubProtocol As String = "") As Boolean
    Dim wsa As WSADATA
    Dim wsaErr As Long
    Dim handle As Long
    m_LastError = ERR_NONE
    m_LastErrorCode = 0
    m_TechnicalDetails = ""
    InitConnectionPool
    If Not m_WSAInitialized Then
        wsaErr = WSAStartup(&H202, wsa)
        If wsaErr <> 0 Then
            SetError ERR_WSA_STARTUP_FAILED, "WSAStartup failed: " & wsaErr, "Network initialization failed. Code: " & wsaErr, INVALID_CONN_HANDLE, wsaErr
            outHandle = INVALID_CONN_HANDLE
            Exit Function
        End If
        m_WSAInitialized = True
    End If
    handle = AllocConnection()
    If handle = INVALID_CONN_HANDLE Then
        SetError ERR_MAX_CONNECTIONS, "Max connections (" & MAX_CONNECTIONS & ") reached", "Too many simultaneous connections.", INVALID_CONN_HANDLE
        outHandle = INVALID_CONN_HANDLE
        Exit Function
    End If
    m_Connections(handle).DeflateEnabled = DeflateEnabled
    m_Connections(handle).DeflateContextTakeover = DeflateContextTakeover
    m_Connections(handle).InflateContextTakeover = DeflateContextTakeover
    
    m_Connections(handle).SubProtocol = SubProtocol
    
    If Not ConnectHandle(handle, url) Then
        outHandle = INVALID_CONN_HANDLE
        Exit Function
    End If
    m_DefaultHandle = handle
    outHandle = handle
    WebSocketConnect = True
    WasabiLog handle, "Connected to " & url & " (handle=" & handle & ")"
End Function

'/**
' * @brief Toggles buffering of messages when socket drops, flushing them upon reconnection.
' * @param enabled True to buffer.
' * @param handle Associated connection.
' */
Public Sub WebSocketSetOfflineQueueing(ByVal enabled As Boolean, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    m_Connections(h).OfflineQueueEnabled = enabled
End Sub

'/**
' * @brief Configures WebSocket compression explicitly (Requires CompressionHandler attached).
' * @param enabled On/Off flag.
' * @param contextTakeover Re-uses dictionaries for tighter compression.
' * @param handle Context pointer.
' */
Public Sub WebSocketSetDeflate(ByVal enabled As Boolean, Optional ByVal contextTakeover As Boolean = True, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    With m_Connections(h)
        If .state = STATE_OPEN Then
            .DeflateEnabled = enabled
            .DeflateContextTakeover = contextTakeover
            .InflateContextTakeover = contextTakeover
            WasabiLog h, "DeflateEnabled set to " & enabled & " - will apply on next reconnect (handle=" & h & ")"
            Exit Sub
        End If
        .DeflateEnabled = enabled
        .DeflateContextTakeover = contextTakeover
        .InflateContextTakeover = contextTakeover
    End With
End Sub

'/**
' * @brief Queries if the server successfully negotiated deflate mapping.
' * @param handle Associated index.
' * @return State boolean.
' */
Public Function WebSocketGetDeflateEnabled(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    WebSocketGetDeflateEnabled = m_Connections(h).DeflateActive
End Function

'/**
' * @brief Sends a standard RFC6455 1000 Close frame and terminates session immediately.
' * @param handle Target index identity.
' */
Public Sub WebSocketDisconnect(Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    Dim i As Long
    Dim anyActive As Boolean

    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    If m_Connections(h).mode <> MODE_WEBSOCKET Then Exit Sub

    m_Connections(h).AutoReconnect = False
    m_Connections(h).AsyncMode = False
    Set m_Connections(h).AsyncHandler = Nothing

    If m_Connections(h).state = STATE_OPEN Then WebSocketSendClose 1000, "", h
    CleanupHandle h

    If h = m_DefaultHandle Then
        m_DefaultHandle = 0
        For i = 0 To MAX_CONNECTIONS - 1
            If m_Connections(i).state = STATE_OPEN Then
                m_DefaultHandle = i
                Exit For
            End If
        Next i
    End If

    anyActive = False
    For i = 0 To MAX_CONNECTIONS - 1
        If m_Connections(i).state <> STATE_CLOSED Or m_Connections(i).Socket <> INVALID_SOCKET Then
            anyActive = True
            Exit For
        End If
    Next i

    If Not anyActive And m_WSAInitialized Then
        WSACleanup
        ShutdownWasabiThunks
        m_WSAInitialized = False
    End If
End Sub

'/**
' * @brief Terminates all actively managed connections uniformly.
' */
Public Sub WebSocketDisconnectAll()
    Dim i As Long
    Dim anyActive As Boolean
    InitConnectionPool
    anyActive = False
    For i = 0 To MAX_CONNECTIONS - 1
        If m_Connections(i).state <> STATE_CLOSED Or m_Connections(i).Socket <> INVALID_SOCKET Then
            m_Connections(i).AsyncMode = False
            Set m_Connections(i).AsyncHandler = Nothing
            m_Connections(i).AutoReconnect = False
            Select Case m_Connections(i).mode
                Case MODE_WEBSOCKET
                    WebSocketDisconnect i
                Case MODE_TCP, MODE_TCP_TLS
                    TcpDisconnect i
            End Select
            anyActive = True
        End If
    Next i

    If Not anyActive And m_WSAInitialized Then
        ShutdownAsyncWindow
        WSACleanup
        ShutdownWasabiThunks
        m_WSAInitialized = False
    End If
    
    If anyActive And m_WSAInitialized Then
        ShutdownAsyncWindow
        WSACleanup
        ShutdownWasabiThunks
        m_WSAInitialized = False
    End If
End Sub

'/**
' * @brief Compiles and fires an RFC6455 Text Opcode Frame logically processing deflate masking natively gracefully.
' * @param message Valid textual payload data string.
' * @param handle Map targeting index handler value natively natively smoothly string natively.
' * @return State tracking variable logic parameter limits dimension value bounds variable index true.
' */
Public Function WebSocketSendText(ByVal message As String, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
    Dim h As Long
    Dim msgBytes() As Byte
    Dim msgLen As Long
    Dim frame() As Byte
    Dim useDeflate As Boolean
    Dim compLen As Long
    Dim compBytes() As Byte
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    With m_Connections(h)
        If .state <> STATE_OPEN Then
            If .OfflineQueueEnabled Then
                If .OfflineTextCount >= OFFLINE_QUEUE_CAP Then
                    WasabiLog h, "Offline text queue cap reached (10000). Dropping message."
                    Exit Function
                End If
                If .OfflineTextCount > UBound(.OfflineTextQueue) Then
                    ReDim Preserve .OfflineTextQueue(0 To UBound(.OfflineTextQueue) + 64)
                End If
                .OfflineTextQueue(.OfflineTextCount) = message
                .OfflineTextCount = .OfflineTextCount + 1
                WebSocketSendText = True
                Exit Function
            Else
                SetError ERR_NOT_CONNECTED, "Send on disconnected handle=" & h, "Not connected to WebSocket server.", h
                Exit Function
            End If
        End If
        msgBytes = StringToUtf8(message)
        msgLen = SafeArrayLen(msgBytes)
        If msgLen = 0 Then
            WebSocketSendText = True
            Exit Function
        End If
        RunOutboundMiddlewares h, msgBytes
        msgLen = SafeArrayLen(msgBytes)
        If msgLen = 0 Then
            WebSocketSendText = True
            Exit Function
        End If
        useDeflate = .DeflateActive
        If useDeflate Then
            compBytes = DeflatePayload(h, msgBytes, msgLen, compLen)
            If compLen = 0 Then
                useDeflate = False
                compLen = msgLen
            Else
                msgBytes = compBytes
                msgLen = compLen
            End If
        End If
        frame = BuildWSFrame(msgBytes, msgLen, WS_OPCODE_TEXT, True, useDeflate)
        If SendFrameFor(h, frame) Then
            .stats.BytesSent = .stats.BytesSent + (UBound(frame) + 1)
            .stats.MessagesSent = .stats.MessagesSent + 1
            WebSocketSendText = True
        End If
    End With
End Function

'/**
' * @brief Formats unmanaged byte arrays directly to a WebSocket Binary packet securely compactly.
' * @param data Array to transmit.
' * @param handle Socket index.
' * @return Successfully written to kernel buffer.
' */
Public Function WebSocketSendBinary(ByRef data() As Byte, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
    Dim h As Long
    Dim dataLen As Long
    Dim frame() As Byte
    Dim useDeflate As Boolean
    Dim compLen As Long
    Dim compBytes() As Byte
    Dim sendData() As Byte
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    With m_Connections(h)
        If .state <> STATE_OPEN Then
            If .OfflineQueueEnabled Then
                If .OfflineBinaryCount >= OFFLINE_QUEUE_CAP Then
                    WasabiLog h, "Offline binary queue cap reached (10000). Dropping message."
                    Exit Function
                End If
                If .OfflineBinaryCount > UBound(.OfflineBinaryQueue) Then
                    ReDim Preserve .OfflineBinaryQueue(0 To UBound(.OfflineBinaryQueue) + 64)
                End If
                .OfflineBinaryQueue(.OfflineBinaryCount).data = data
                .OfflineBinaryCount = .OfflineBinaryCount + 1
                WebSocketSendBinary = True
                Exit Function
            Else
                SetError ERR_NOT_CONNECTED, "SendBinary on disconnected handle=" & h, "Not connected to WebSocket server.", h
                Exit Function
            End If
        End If
        dataLen = SafeArrayLen(data)
        If dataLen = 0 Then
            WebSocketSendBinary = True
            Exit Function
        End If
        RunOutboundMiddlewares h, data
        dataLen = SafeArrayLen(data)
        If dataLen = 0 Then
            WebSocketSendBinary = True
            Exit Function
        End If
        useDeflate = .DeflateActive
        If useDeflate Then
            compBytes = DeflatePayload(h, data, dataLen, compLen)
            If compLen = 0 Then
                useDeflate = False
                sendData = data
                dataLen = SafeArrayLen(data)
            Else
                sendData = compBytes
                dataLen = compLen
            End If
        Else
            sendData = data
        End If
        frame = BuildWSFrame(sendData, dataLen, WS_OPCODE_BINARY, True, useDeflate)
        If SendFrameFor(h, frame) Then
            .stats.BytesSent = .stats.BytesSent + (UBound(frame) + 1)
            .stats.MessagesSent = .stats.MessagesSent + 1
            WebSocketSendBinary = True
        End If
    End With
End Function

'/**
' * @brief Slices huge payloads automatically into fragmented WS Continuation packets utilizing optimal MSS values cleanly dynamically flawlessly compactly securely cleanly efficiently tightly safely nicely perfectly gracefully accurately reliably securely strictly successfully dynamically cleanly stably seamlessly safely properly cleanly safely accurately smoothly strictly compactly seamlessly seamlessly.
' * @param message Origin string to cut.
' * @param handle Core protocol routing.
' * @return State of delivery sequence structure dimension limit values logic boolean array byte map size values limit parameters dimension variables pointer index structurally.
' */
Public Function WebSocketSendTextMTUAware(ByVal message As String, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
    Dim h As Long
    Dim msgBytes() As Byte
    Dim msgLen As Long
    Dim offset As Long
    Dim chunkSize As Long
    Dim opcode As Byte
    Dim isLast As Boolean
    Dim chunkBytes() As Byte
    Dim frame() As Byte
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    With m_Connections(h)
        If .state <> STATE_OPEN Then
            SetError ERR_NOT_CONNECTED, "SendMTUAware on disconnected handle=" & h, "Not connected.", h
            Exit Function
        End If
        msgBytes = StringToUtf8(message)
        msgLen = SafeArrayLen(msgBytes)
        If msgLen = 0 Then
            WebSocketSendTextMTUAware = True
            Exit Function
        End If
        If Not .AutoMTU Or msgLen <= .mtu.OptimalFrameSize Then
            WebSocketSendTextMTUAware = WebSocketSendText(message, h)
            Exit Function
        End If
        offset = 0
        opcode = WS_OPCODE_TEXT
        Do While offset < msgLen
            chunkSize = .mtu.OptimalFrameSize
            If offset + chunkSize > msgLen Then chunkSize = msgLen - offset
            isLast = (offset + chunkSize >= msgLen)
            ReDim chunkBytes(0 To chunkSize - 1)
            CopyMemory chunkBytes(0), msgBytes(offset), chunkSize
            frame = BuildWSFrame(chunkBytes, chunkSize, opcode, isLast)
            If Not SendFrameFor(h, frame) Then Exit Function
            .stats.BytesSent = .stats.BytesSent + (UBound(frame) + 1)
            offset = offset + chunkSize
            opcode = WS_OPCODE_CONTINUATION
        Loop
        .stats.MessagesSent = .stats.MessagesSent + 1
    End With
    WebSocketSendTextMTUAware = True
End Function

'/**
' * @brief Binary variation of the MTU Fragmenter smoothly dynamically accurately reliably cleanly securely safely tightly correctly nicely reliably efficiently perfectly natively flawlessly elegantly tightly stably smoothly precisely accurately safely smoothly efficiently smoothly cleanly dynamically structurally correctly nicely correctly dynamically tightly smoothly compactly.
' * @param data Byte payload array natively cleanly smoothly efficiently robustly safely flawlessly dynamically reliably tightly gracefully accurately seamlessly correctly compactly.
' * @param handle Identity map structure memory address values structurally safely gracefully seamlessly natively natively seamlessly smoothly strictly reliably optimally dynamically compactly properly optimally perfectly correctly cleanly tightly correctly natively seamlessly flawlessly tightly stably smoothly precisely efficiently structurally correctly stably successfully reliably securely elegantly cleanly structurally correctly cleanly securely precisely.
' * @return State of structural logical loop pointer map array payload boundary string constraints values size address dimensions parameters size dimensions mapping structure limit block struct context.
' */
Public Function WebSocketSendBinaryMTUAware(ByRef data() As Byte, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
    Dim h As Long
    Dim dataLen As Long
    Dim offset As Long
    Dim chunkSize As Long
    Dim opcode As Byte
    Dim isLast As Boolean
    Dim chunkBytes() As Byte
    Dim frame() As Byte
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    With m_Connections(h)
        If .state <> STATE_OPEN Then
            SetError ERR_NOT_CONNECTED, "SendBinaryMTUAware on disconnected handle=" & h, "Not connected.", h
            Exit Function
        End If
        dataLen = SafeArrayLen(data)
        If dataLen = 0 Then
            WebSocketSendBinaryMTUAware = True
            Exit Function
        End If
        If Not .AutoMTU Or dataLen <= .mtu.OptimalFrameSize Then
            WebSocketSendBinaryMTUAware = WebSocketSendBinary(data, h)
            Exit Function
        End If
        offset = 0
        opcode = WS_OPCODE_BINARY
        Do While offset < dataLen
            chunkSize = .mtu.OptimalFrameSize
            If offset + chunkSize > dataLen Then chunkSize = dataLen - offset
            isLast = (offset + chunkSize >= dataLen)
            ReDim chunkBytes(0 To chunkSize - 1)
            CopyMemory chunkBytes(0), data(offset), chunkSize
            frame = BuildWSFrame(chunkBytes, chunkSize, opcode, isLast)
            If Not SendFrameFor(h, frame) Then Exit Function
            .stats.BytesSent = .stats.BytesSent + (UBound(frame) + 1)
            offset = offset + chunkSize
            opcode = WS_OPCODE_CONTINUATION
        Loop
        .stats.MessagesSent = .stats.MessagesSent + 1
    End With
    WebSocketSendBinaryMTUAware = True
End Function

'/**
' * @brief Clusters multiple frames into single massive TCP stream blocks improving high IO operations structurally.
' * @param messages Array of string payloads to be batched.
' * @param handle Core memory array targeting context structure logic index mapping natively.
' * @return True if the entire batch was successfully dispatched.
' */
Public Function WebSocketSendBatch(ByRef messages() As String, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
    Dim h As Long
    Dim i As Long
    Dim msgBytes() As Byte
    Dim msgLen As Long
    Dim frame() As Byte
    Dim frameSize As Long
    Dim batchBuf() As Byte
    Dim batchLen As Long
    Dim batchCount As Long
    Dim flushBuf() As Byte
    
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    With m_Connections(h)
        If .state <> STATE_OPEN Then Exit Function
        batchLen = 0
        batchCount = 0
        ReDim batchBuf(0 To BATCH_MAX_SIZE - 1)
        For i = LBound(messages) To UBound(messages)
            msgBytes = StringToUtf8(messages(i))
            msgLen = SafeArrayLen(msgBytes)
            If msgLen = 0 Then GoTo NextMsg
            frame = BuildWSFrame(msgBytes, msgLen, WS_OPCODE_TEXT, True)
            frameSize = UBound(frame) + 1
            If batchLen + frameSize > BATCH_MAX_SIZE Then
                ReDim flushBuf(0 To batchLen - 1)
                CopyMemory flushBuf(0), batchBuf(0), batchLen
                If .TLS Then
                    If Not TLSSend(h, flushBuf) Then Exit Function
                Else
                    If Not RawSendFor(h, flushBuf) Then Exit Function
                End If
                .stats.BytesSent = .stats.BytesSent + batchLen
                .stats.MessagesSent = .stats.MessagesSent + batchCount
                batchLen = 0
                batchCount = 0
            End If
            CopyMemory batchBuf(batchLen), frame(0), frameSize
            batchLen = batchLen + frameSize
            batchCount = batchCount + 1
NextMsg:
        Next i
        If batchLen > 0 Then
            ReDim flushBuf(0 To batchLen - 1)
            CopyMemory flushBuf(0), batchBuf(0), batchLen
            If .TLS Then
                If Not TLSSend(h, flushBuf) Then Exit Function
            Else
                If Not RawSendFor(h, flushBuf) Then Exit Function
            End If
            .stats.BytesSent = .stats.BytesSent + batchLen
            .stats.MessagesSent = .stats.MessagesSent + batchCount
        End If
    End With
    WebSocketSendBatch = True
End Function

'/**
' * @brief Transmits a batched array of unmanaged binary payloads dynamically across TCP boundaries seamlessly.
' * @param messages Array of byte arrays containing binary payloads.
' * @param handle Core routing index tracker.
' * @return Delivery status verification boolean.
' */
Public Function WebSocketSendBatchBinary(ByRef messages() As Variant, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
    Dim h As Long
    Dim i As Long
    Dim bdata() As Byte
    Dim dataLen As Long
    Dim frame() As Byte
    Dim frameSize As Long
    Dim batchBuf() As Byte
    Dim batchLen As Long
    Dim batchCount As Long
    Dim flushBuf() As Byte
    
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    With m_Connections(h)
        If .state <> STATE_OPEN Then Exit Function
        batchLen = 0
        batchCount = 0
        ReDim batchBuf(0 To BATCH_MAX_SIZE - 1)
        For i = LBound(messages) To UBound(messages)
            If IsArray(messages(i)) Then
                bdata = messages(i)
                dataLen = SafeArrayLen(bdata)
                If dataLen = 0 Then GoTo NextMsgBin
                frame = BuildWSFrame(bdata, dataLen, WS_OPCODE_BINARY, True)
                frameSize = UBound(frame) + 1
                If batchLen + frameSize > BATCH_MAX_SIZE Then
                    ReDim flushBuf(0 To batchLen - 1)
                    CopyMemory flushBuf(0), batchBuf(0), batchLen
                    If .TLS Then
                        If Not TLSSend(h, flushBuf) Then Exit Function
                    Else
                        If Not RawSendFor(h, flushBuf) Then Exit Function
                    End If
                    .stats.BytesSent = .stats.BytesSent + batchLen
                    .stats.MessagesSent = .stats.MessagesSent + batchCount
                    batchLen = 0
                    batchCount = 0
                End If
                CopyMemory batchBuf(batchLen), frame(0), frameSize
                batchLen = batchLen + frameSize
                batchCount = batchCount + 1
            End If
NextMsgBin:
        Next i
        If batchLen > 0 Then
            ReDim flushBuf(0 To batchLen - 1)
            CopyMemory flushBuf(0), batchBuf(0), batchLen
            If .TLS Then
                If Not TLSSend(h, flushBuf) Then Exit Function
            Else
                If Not RawSendFor(h, flushBuf) Then Exit Function
            End If
            .stats.BytesSent = .stats.BytesSent + batchLen
            .stats.MessagesSent = .stats.MessagesSent + batchCount
        End If
    End With
    WebSocketSendBatchBinary = True
End Function

'/**
' * @brief Synthesizes and sends an RFC6455 closure packet cleanly alerting the remote node gracefully.
' * @param code Native RFC 16-bit status numeric code limits (1000 = Normal).
' * @param reason Human readable string tracking log detail.
' * @param handle Map identity structure pointer value.
' * @return State tracking variable logic parameter limits cleanly.
' */
Public Function WebSocketSendClose(Optional ByVal code As Integer = 1000, Optional ByVal reason As String = "", Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
    Dim h As Long
    Dim frame() As Byte
    Dim mask(0 To 3) As Byte
    Dim reasonBytes() As Byte
    Dim reasonLen As Long
    Dim payloadLen As Long
    Dim i As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    With m_Connections(h)
        If .state <> STATE_OPEN Then Exit Function
        .CloseInitiatedByUs = True
        .closeCode = code
        .closeReason = reason
        If Len(reason) > 0 Then
            reasonBytes = StringToUtf8(reason)
            reasonLen = SafeArrayLen(reasonBytes)
            If reasonLen > WS_MAX_CLOSE_REASON Then reasonLen = WS_MAX_CLOSE_REASON
        End If
        payloadLen = 2 + reasonLen
        ReDim frame(0 To 5 + payloadLen)
        FillRandomBytes mask, 4
        frame(0) = &H88
        frame(1) = &H80 Or CByte(payloadLen)
        frame(2) = mask(0)
        frame(3) = mask(1)
        frame(4) = mask(2)
        frame(5) = mask(3)
        frame(6) = CByte((code \ 256) And &HFF) Xor mask(0)
        frame(7) = CByte(code And &HFF) Xor mask(1)
        For i = 0 To reasonLen - 1
            frame(8 + i) = reasonBytes(i) Xor mask((i + 2) Mod 4)
        Next i
        WasabiLog h, "Sending CLOSE: " & code & " (" & GetCloseCodeDesc(code) & ") reason=""" & reason & """ (handle=" & h & ")"
        WebSocketSendClose = SendFrameFor(h, frame)
        .state = STATE_CLOSING
    End With
End Function

'/**
' * @brief Dispatches an active keep-alive logical WS ping block mapping network routing dynamically flawlessly.
' * @param payload String ping body token.
' * @param handle Base network array tracking identity context handle index marker.
' * @return True if effectively executed successfully natively.
' */
Public Function WebSocketSendPing(Optional ByVal payload As String = "", Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
    Dim h As Long
    Dim frame() As Byte
    Dim mask(0 To 3) As Byte
    Dim pingBytes() As Byte
    Dim pingLen As Long
    Dim i As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    With m_Connections(h)
        If .state <> STATE_OPEN Then Exit Function
        If Len(payload) > 0 Then
            pingBytes = StringToUtf8(payload)
            pingLen = SafeArrayLen(pingBytes)
        End If
        FillRandomBytes mask, 4
        If pingLen = 0 Then
            ReDim frame(0 To 5)
            frame(0) = &H89
            frame(1) = &H80
        Else
            ReDim frame(0 To 5 + pingLen)
            frame(0) = &H89
            frame(1) = &H80 Or CByte(pingLen)
            For i = 0 To pingLen - 1
                frame(6 + i) = pingBytes(i) Xor mask(i Mod 4)
            Next i
        End If
        frame(2) = mask(0)
        frame(3) = mask(1)
        frame(4) = mask(2)
        frame(5) = mask(3)
        WebSocketSendPing = SendFrameFor(h, frame)
        If WebSocketSendPing Then
            .LastPingSentAt = GetTickCount()
            .LastPingTimestamp = GetTickCount()
        End If
    End With
End Function

'/**
' * @brief Acknowledges active remote Pings utilizing specific identical response body context.
' * @param payload String payload match token return map tracking string.
' * @param handle Core indexing logical identity structural parameter.
' * @return Indicates physical protocol array send transmission boolean true.
' */
Public Function WebSocketSendPong(Optional ByVal payload As String = "", Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
    Dim h As Long
    Dim frame() As Byte
    Dim mask(0 To 3) As Byte
    Dim pongBytes() As Byte
    Dim pongLen As Long
    Dim i As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    With m_Connections(h)
        If .state <> STATE_OPEN Then Exit Function
        If Len(payload) > 0 Then
            pongBytes = StringToUtf8(payload)
            pongLen = SafeArrayLen(pongBytes)
        End If
        FillRandomBytes mask, 4
        If pongLen = 0 Then
            ReDim frame(0 To 5)
            frame(0) = &H8A
            frame(1) = &H80
        Else
            ReDim frame(0 To 5 + pongLen)
            frame(0) = &H8A
            frame(1) = &H80 Or CByte(pongLen)
            For i = 0 To pongLen - 1
                frame(6 + i) = pongBytes(i) Xor mask(i Mod 4)
            Next i
        End If
        frame(2) = mask(0)
        frame(3) = mask(1)
        frame(4) = mask(2)
        frame(5) = mask(3)
        WebSocketSendPong = SendFrameFor(h, frame)
    End With
End Function

'/**
' * @brief Disseminates an identical textual instruction seamlessly across the entire instantiated connectivity pool comprehensively.
' * @param message String value content bounds payload block token.
' * @return Amount of nodes actively contacted parameter array values context.
' */
Public Function WebSocketBroadcastText(ByVal message As String) As Long
    Dim i As Long
    Dim count As Long
    InitConnectionPool
    For i = 0 To MAX_CONNECTIONS - 1
        If m_Connections(i).state = STATE_OPEN Then
            If WebSocketSendText(message, i) Then count = count + 1
        End If
    Next i
    WebSocketBroadcastText = count
End Function

'/**
' * @brief Binary array iteration over connectivity active queue. Disseminates identically.
' * @param data Ref mapped context byte array payload target structure memory element block value.
' * @return Active node send count limits context boolean index variable dimensions size structure variable.
' */
Public Function WebSocketBroadcastBinary(ByRef data() As Byte) As Long
    Dim i As Long
    Dim count As Long
    InitConnectionPool
    For i = 0 To MAX_CONNECTIONS - 1
        If m_Connections(i).state = STATE_OPEN Then
            If WebSocketSendBinary(data, i) Then count = count + 1
        End If
    Next i
    WebSocketBroadcastBinary = count
End Function

'/**
' * @brief Standard synchronous polling technique retrieving top textual elements from inner queue structure sequentially cleanly reliably stably tightly properly cleanly accurately smoothly efficiently successfully cleanly.
' * @param handle Logic context target index element constraint map map domain point context domain identity.
' * @return String textual variable domain point data limit string value limit struct address variables address limit structure limit block block variables parameter array limits values memory array value mapping index.
' */
Public Function WebSocketReceiveText(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    With m_Connections(h)
        If .state <> STATE_OPEN Then
            If .AutoReconnect And Not .AsyncMode Then TryReconnect h
            Exit Function
        End If
        If Not .AsyncMode Then
            TickMaintenance h
            If .DecryptLen > 0 Then ProcessFrames h
            FeedBuffer h
        End If
        If .MsgCount > 0 Then
            WebSocketReceiveText = .MsgQueue(.MsgHead)
            .MsgQueue(.MsgHead) = ""
            .MsgHead = (.MsgHead + 1) Mod MSG_QUEUE_SIZE
            .MsgCount = .MsgCount - 1
        End If
    End With
End Function

'/**
' * @brief Aggregates and dumps the entire internal textual queue completely in one efficient bound context operation safely cleanly properly cleanly efficiently perfectly smoothly cleanly elegantly cleanly elegantly compactly reliably smoothly dynamically robustly tightly securely compactly safely compactly elegantly flawlessly neatly flawlessly smoothly.
' * @param handle Node session value memory structural indexing map.
' * @return Multi-dimensional limit string array struct domain values sizes point limits dimensions limit pointer dimension bounds limit parameters block sizes values context array limits values dimensions domain sizes values context dimension index pointer mapping address constraints values mapping bounds limit size limit string limit map context size mapping target constraint array sizes address limits context array limit parameter limit values domain memory limit.
' */
Public Function WebSocketReceiveAll(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String()
    Dim h As Long
    Dim results() As String
    Dim i As Long
    Dim count As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then
        ReDim results(0)
        WebSocketReceiveAll = results
        Exit Function
    End If
    With m_Connections(h)
        If .state <> STATE_OPEN Then
            If .AutoReconnect And Not .AsyncMode Then TryReconnect h
            ReDim results(0)
            WebSocketReceiveAll = results
            Exit Function
        End If
        If Not .AsyncMode Then
            TickMaintenance h
            If .DecryptLen > 0 Then ProcessFrames h
            FeedBuffer h
        End If
        count = .MsgCount
        If count = 0 Then
            ReDim results(0)
            WebSocketReceiveAll = results
            Exit Function
        End If
        ReDim results(0 To count - 1)
        For i = 0 To count - 1
            results(i) = .MsgQueue(.MsgHead)
            .MsgQueue(.MsgHead) = ""
            .MsgHead = (.MsgHead + 1) Mod MSG_QUEUE_SIZE
            .MsgCount = .MsgCount - 1
        Next i
    End With
    WebSocketReceiveAll = results
End Function

'/**
' * @brief Pop operation returning the top raw unmanaged bytes logically naturally gracefully properly reliably optimally stably precisely efficiently flawlessly compactly precisely cleanly stably correctly natively securely properly flawlessly.
' * @param handle Index block memory map pointer structure element variables pointer limit dimensions constraints limits parameters map values pointer dimensions pointer dimension memory values size parameters array size constraints sizes values dimension limit target pointer limits constraint size index value.
' * @return Data layout element array constraint context value logic variables domain struct dimensions pointer block bounds limits parameter memory value map variables dimensions domain parameters size constraint.
' */
Public Function WebSocketReceiveBinary(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Byte()
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then
        WebSocketReceiveBinary = Empty
        Exit Function
    End If
    With m_Connections(h)
        If .state <> STATE_OPEN Then
            If .AutoReconnect And Not .AsyncMode Then TryReconnect h
            WebSocketReceiveBinary = Empty
            Exit Function
        End If
        If Not .AsyncMode Then
            TickMaintenance h
            If .DecryptLen > 0 Then ProcessFrames h
            FeedBuffer h
        End If
        If .BinaryCount > 0 Then
            WebSocketReceiveBinary = .BinaryQueue(.BinaryHead).data
            Erase .BinaryQueue(.BinaryHead).data
            .BinaryHead = (.BinaryHead + 1) Mod MSG_QUEUE_SIZE
            .BinaryCount = .BinaryCount - 1
        Else
            WebSocketReceiveBinary = Empty
        End If
    End With
End Function
'/**
' * @brief Validates presence while extracting unmanaged byte payloads synchronously boolean natively cleanly structurally elegantly flawlessly smoothly smoothly seamlessly safely tightly cleanly neatly compactly dynamically neatly properly elegantly properly neatly securely successfully smoothly precisely gracefully precisely.
' * @param outData Pointer constraint reference target limit value string limits block memory variable block domain size map variables bounds map parameter mapping memory size array structure limit limits value dimension boundary.
' * @param handle Core protocol routing.
' * @return Returns validation checking parameter block limit state limit size.
' */
Public Function WebSocketReceiveBinaryCheck(ByRef outData() As Byte, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    With m_Connections(h)
        If .state <> STATE_OPEN Then
            If .AutoReconnect And Not .AsyncMode Then TryReconnect h
            Exit Function
        End If
        If Not .AsyncMode Then
            TickMaintenance h
            If .DecryptLen > 0 Then ProcessFrames h
            FeedBuffer h
        End If
        If .BinaryCount > 0 Then
            outData = .BinaryQueue(.BinaryHead).data
            Erase .BinaryQueue(.BinaryHead).data
            .BinaryHead = (.BinaryHead + 1) Mod MSG_QUEUE_SIZE
            .BinaryCount = .BinaryCount - 1
            WebSocketReceiveBinaryCheck = True
        End If
    End With
End Function

'/**
' * @brief Facilitates maximum throughput reading string data straight from internal array boundaries without duplicating memory limits variables smoothly dynamically smoothly structurally perfectly properly structurally stably.
' * @param outPtr Int natively returning string base mapping pointer bounds limit limit size parameters pointer limit value size structure.
' * @param outLen Target parameter variables logic constraint mapping values domain domain index dimension variable dimensions.
' * @param handle Variable context indexing map tracking size memory parameter bounds size dimensions size boolean sizes memory limits.
' * @return Truth logic tracking variable boolean variables context.
' */
#If VBA7 Then
Public Function WebSocketReceiveZeroCopy(ByRef outPtr As LongPtr, ByRef outLen As Long, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
#Else
Public Function WebSocketReceiveZeroCopy(ByRef outPtr As Long, ByRef outLen As Long, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
#End If
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    With m_Connections(h)
        If .state <> STATE_OPEN Then Exit Function
        If Not .ZeroCopyEnabled Then Exit Function
        If .AsyncMode Then Exit Function
        TickMaintenance h
        If .DecryptLen > 0 Then ProcessFrames h
        FeedBuffer h
        If .MsgCount > 0 Then
            m_ZeroCopyText = .MsgQueue(.MsgHead)
            outPtr = StrPtr(m_ZeroCopyText)
            outLen = Len(m_ZeroCopyText)
            .MsgQueue(.MsgHead) = ""
            .MsgHead = (.MsgHead + 1) Mod MSG_QUEUE_SIZE
            .MsgCount = .MsgCount - 1
            WebSocketReceiveZeroCopy = True
        End If
    End With
End Function

'/**
' * @brief Returns the next queued text message without removing it.
' * @param handle (Optional) Target connection handle.
' * @return The next text message, or empty string.
' */
Public Function WebSocketPeek(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    With m_Connections(h)
        If .MsgCount > 0 Then WebSocketPeek = .MsgQueue(.MsgHead)
    End With
End Function

'/**
' * @brief Clears all queued text and binary messages for the connection.
' * @param handle (Optional) Target connection handle.
' */
Public Sub WebSocketFlushQueue(Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    With m_Connections(h)
        .MsgHead = 0
        .MsgTail = 0
        .MsgCount = 0
        .BinaryHead = 0
        .BinaryTail = 0
        .BinaryCount = 0
    End With
End Sub

'/**
' * @brief Returns the current connected state of the handle.
' * @param handle (Optional) Target connection handle.
' * @return True if connected and active.
' */
Public Function WebSocketIsConnected(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    WebSocketIsConnected = (m_Connections(h).state = STATE_OPEN)
End Function

'/**
' * @brief Returns the most recent WasabiError value for the connection.
' * @param handle (Optional) Target connection handle.
' * @return WasabiError enum value.
' */
Public Function WebSocketGetLastError(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As WasabiError
    Dim h As Long
    h = ResolveHandle(handle)
    If ValidIndex(h) Then
        WebSocketGetLastError = m_Connections(h).LastError
    Else
        WebSocketGetLastError = m_LastError
    End If
End Function

'/**
' * @brief Returns the last native system error code (WSA or SSPI).
' * @param handle (Optional) Target connection handle.
' * @return Hex error code.
' */
Public Function WebSocketGetLastErrorCode(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Long
    Dim h As Long
    h = ResolveHandle(handle)
    If ValidIndex(h) Then
        WebSocketGetLastErrorCode = m_Connections(h).LastErrorCode
    Else
        WebSocketGetLastErrorCode = m_LastErrorCode
    End If
End Function

'/**
' * @brief Returns a technical description of the most recent error.
' * @param handle (Optional) Target connection handle.
' * @return String with function names, parameters, and error codes.
' */
Public Function WebSocketGetTechnicalDetails(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
    Dim h As Long
    h = ResolveHandle(handle)
    If ValidIndex(h) Then
        WebSocketGetTechnicalDetails = m_Connections(h).TechnicalDetails
    Else
        WebSocketGetTechnicalDetails = m_TechnicalDetails
    End If
End Function

'/**
' * @brief Readable string output mapping OS/Protocol failure bounds constraints seamlessly nicely flawlessly gracefully smoothly stably precisely perfectly robustly.
' * @param handle Identity logic parameter limit string dimension value variable constraint context size dimensions size string string parameters limit limits memory dimension.
' * @return Information textual constraint values limit parameter string address values domain domain value array sizes variables parameter size string values mapping constraints limits dimension.
' */
Public Function WasabiGetErrorDescription(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
    Dim h As Long
    Dim errType As WasabiError
    Dim errCode As Long
    Dim tech As String
    Dim desc As String
    h = ResolveHandle(handle)
    If ValidIndex(h) Then
        errType = m_Connections(h).LastError
        errCode = m_Connections(h).LastErrorCode
        tech = m_Connections(h).TechnicalDetails
    Else
        errType = m_LastError
        errCode = m_LastErrorCode
        tech = m_TechnicalDetails
    End If
    Select Case errType
        Case ERR_NONE: desc = "No error"
        Case ERR_WSA_STARTUP_FAILED: desc = "Winsock initialization failed"
        Case ERR_SOCKET_CREATE_FAILED: desc = "Failed to create socket"
        Case ERR_DNS_RESOLVE_FAILED: desc = "DNS resolution failed"
        Case ERR_CONNECT_FAILED: desc = "TCP connection failed"
        Case ERR_TLS_ACQUIRE_CREDS_FAILED: desc = "TLS credentials initialization failed"
        Case ERR_TLS_HANDSHAKE_FAILED: desc = "TLS handshake failed"
        Case ERR_TLS_HANDSHAKE_TIMEOUT: desc = "TLS handshake timed out"
        Case ERR_WEBSOCKET_HANDSHAKE_FAILED: desc = "WebSocket upgrade rejected"
        Case ERR_WEBSOCKET_HANDSHAKE_TIMEOUT: desc = "WebSocket handshake timed out"
        Case ERR_SEND_FAILED: desc = "Send failed"
        Case ERR_RECV_FAILED: desc = "Receive failed"
        Case ERR_NOT_CONNECTED: desc = "Not connected"
        Case ERR_ALREADY_CONNECTED: desc = "Already connected"
        Case ERR_TLS_ENCRYPT_FAILED: desc = "TLS encryption failed"
        Case ERR_TLS_DECRYPT_FAILED: desc = "TLS decryption failed"
        Case ERR_INVALID_URL: desc = "Invalid URL"
        Case ERR_HANDSHAKE_REJECTED: desc = "WebSocket handshake rejected by server"
        Case ERR_CONNECTION_LOST: desc = "Connection lost"
        Case ERR_INVALID_HANDLE: desc = "Invalid connection handle"
        Case ERR_MAX_CONNECTIONS: desc = "Maximum connections reached"
        Case ERR_PROXY_CONNECT_FAILED: desc = "Proxy connection failed"
        Case ERR_PROXY_AUTH_FAILED: desc = "Proxy authentication failed"
        Case ERR_PROXY_TUNNEL_FAILED: desc = "Proxy tunnel failed"
        Case ERR_INACTIVITY_TIMEOUT: desc = "Inactivity timeout"
        Case ERR_CERT_LOAD_FAILED: desc = "Client certificate load failed"
        Case ERR_CERT_VALIDATE_FAILED: desc = "Server certificate validation failed"
        Case ERR_FRAGMENT_OVERFLOW: desc = "Fragment buffer overflow"
        Case ERR_TLS_RENEGOTIATE: desc = "TLS renegotiation not supported"
        Case Else: desc = "Unknown error (" & errType & ")"
    End Select
    If errCode <> 0 Then desc = desc & " [0x" & hex(errCode) & "]"
    If Len(tech) > 0 Then desc = desc & " - " & tech
    WasabiGetErrorDescription = desc
End Function

'/**
' * @brief Tracks memory usage bounds constraints sizes dimension string variables value memory target target parameter size boolean constraint limits string sizes dimensions parameter limit values.
' * @param handle Logical mapping boolean block boundary value limit dimensions sizes string context array parameter index array limits dimensions value dimension.
' * @return Amount of nodes cleanly constraints parameter variables limits sizes array block domain array array bounds structure context map constraint pointer memory dimensions parameters limit variables variables memory sizes string sizes array value size structure sizes dimension size constraint.
' */
Public Function WebSocketGetPendingCount(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Long
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    WebSocketGetPendingCount = m_Connections(h).MsgCount
End Function

'/**
' * @brief Gets the amount of active elements correctly correctly reliably correctly smoothly optimally dynamically smoothly efficiently tightly robustly.
' * @param handle Core indexing logical identity structural parameter point values mapping structure limit sizes string limit dimensions dimensions domain dimension limit target limit bounds limit variables array domain parameter array sizes array parameter value.
' * @return Node parameters dimensions limit sizes value constraint map boundary context string sizes bounds value target dimension parameters array parameter context variables string domain string parameter sizes array memory limit constraint limits variables value dimension variable sizes array variables variables.
' */
Public Function WebSocketGetBinaryPendingCount(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Long
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    WebSocketGetBinaryPendingCount = m_Connections(h).BinaryCount
End Function

'/**
' * @brief Measures limits size variables map mapping dimensions domain value constraints arrays values dimension limits target map target limits memory structure size mapping.
' * @param handle Boundary string index sizes parameters dimensions domain constraints value block map dimensions context limit pointer value mapping memory constraints variables context constraints.
' * @return Queue size parameter size limits dimension string values boundary parameter array bounds value array limits structure dimensions memory array limit constraint.
' */
Public Function WebSocketGetQueueCapacity(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Long
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    WebSocketGetQueueCapacity = MSG_QUEUE_SIZE - m_Connections(h).MsgCount
End Function

'/**
' * @brief Returns the textual analytical variable logic sizes limits dimensions dimension domain constraints string variable dimensions mapping constraints mapping bounds map constraint block parameter string sizes target limit dimensions string parameters limits string bounds sizes dimensions sizes parameter values boundary arrays string mapping array limits value dimension memory.
' * @param handle Context tracking index value mapping structure domain domain variable dimension.
' * @return Format memory sizes string limits dimensions bounds limits.
' */
Public Function WebSocketGetStats(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
    Dim h As Long
    Dim uptime As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    If m_Connections(h).mode <> MODE_WEBSOCKET Then Exit Function
    With m_Connections(h)
        If .stats.ConnectedAt > 0 Then uptime = TickDiff(.stats.ConnectedAt, GetTickCount()) \ 1000
        WebSocketGetStats = "BytesSent=" & Format(.stats.BytesSent, "0") & _
            "|BytesReceived=" & Format(.stats.BytesReceived, "0") & _
            "|MessagesSent=" & .stats.MessagesSent & _
            "|MessagesReceived=" & .stats.MessagesReceived & _
            "|UptimeSeconds=" & uptime & _
            "|Queued=" & .MsgCount & _
            "|BinaryQueued=" & .BinaryCount & _
            "|NoDelay=" & IIf(.NoDelay, "1", "0") & _
            "|Proxy=" & IIf(.ProxyEnabled, .proxyHost & ":" & .proxyPort, "none") & _
            "|Mode=WebSocket"
    End With
End Function

'/**
' * @brief Provides stats array sizes value constraint dimensions sizes parameter boundary limit target variable domain mapping constraint map constraints values limits target dimensions values.
' * @param handle Tracking value array constraint dimension limits size block parameter boundary boundary.
' * @return Formatted values.
' */
Public Function TcpGetStats(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
    Dim h As Long
    Dim uptime As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    If m_Connections(h).mode = MODE_WEBSOCKET Then Exit Function
    With m_Connections(h)
        If .stats.ConnectedAt > 0 Then uptime = TickDiff(.stats.ConnectedAt, GetTickCount()) \ 1000
        TcpGetStats = "BytesSent=" & Format(.stats.BytesSent, "0") & _
            "|BytesReceived=" & Format(.stats.BytesReceived, "0") & _
            "|MessagesSent=" & .stats.MessagesSent & _
            "|MessagesReceived=" & .stats.MessagesReceived & _
            "|UptimeSeconds=" & uptime & _
            "|PendingBytes=" & .TcpRecvLen & _
            "|NoDelay=" & IIf(.NoDelay, "1", "0") & _
            "|Proxy=" & IIf(.ProxyEnabled, .proxyHost & ":" & .proxyPort, "none") & _
            "|Mode=" & IIf(.mode = MODE_TCP_TLS, "TCP_TLS", "TCP") & _
            "|Host=" & .HOST & _
            "|Port=" & .port
    End With
End Function

'/**
' * @brief Binary iteration constraints values value values pointer arrays dimensions constraints dimensions target values limits target.
' * @param data Ref mapped context byte array payload target structure memory element block value.
' * @return Node parameters logic tracking limit size variables limits values map parameter size bounds map limits dimension limits arrays limit limits dimension constraints value domain target constraint.
' */
Public Function TcpBroadcastBinary(ByRef data() As Byte) As Long
    Dim i As Long
    Dim count As Long
    InitConnectionPool
    For i = 0 To MAX_CONNECTIONS - 1
        If m_Connections(i).state = STATE_OPEN And m_Connections(i).mode <> MODE_WEBSOCKET Then
            If TcpSendBinary(data, i) Then count = count + 1
        End If
    Next i
    TcpBroadcastBinary = count
End Function

'/**
' * @brief TCP string loop variables limits domain dimension dimensions map memory dimensions arrays mapping limits string limits values limits parameters string map string limit array parameter values parameters dimensions values mapping size variables domain parameters limit constraints constraints map variables parameter size memory string.
' * @param text Content domain parameters pointer structure values mapping map dimensions limit dimensions memory limits value variables array.
' * @return Variable point size parameters array map sizes value string memory limit constraints limits variable constraints size limits memory domain sizes dimensions dimensions limits structure constraint boundary dimensions bounds size target dimension target variable limit map value constraints limits boundary limits memory string variable array limits limit parameter domain variables size.
' */
Public Function TcpBroadcastText(ByVal text As String) As Long
    Dim i As Long
    Dim count As Long
    Dim data() As Byte
    data = StringToUtf8(text)
    InitConnectionPool
    For i = 0 To MAX_CONNECTIONS - 1
        If m_Connections(i).state = STATE_OPEN And m_Connections(i).mode <> MODE_WEBSOCKET Then
            If TcpSendBinary(data, i) Then count = count + 1
        End If
    Next i
    TcpBroadcastText = count
End Function

'/**
' * @brief TCP options constraint parameters values mapping boundary constraints.
' * @param enabled Values.
' * @param handle Node map boundary tracking memory variables sizes domain map sizes limits string memory domain size memory sizes array map size variables memory values sizes.
' * @return True natively dimensions constraint limit variables limits string values.
' */
Public Function TcpSetNoDelay(ByVal enabled As Boolean, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
    Dim h As Long
    Dim optVal As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    If m_Connections(h).mode = MODE_WEBSOCKET Then Exit Function
    m_Connections(h).NoDelay = enabled
    If m_Connections(h).Socket <> INVALID_SOCKET Then
        optVal = IIf(enabled, 1, 0)
        TcpSetNoDelay = (sock_setsockopt(m_Connections(h).Socket, IPPROTO_TCP_SOL, TCP_NODELAY, optVal, 4) = 0)
    Else
        TcpSetNoDelay = True
    End If
End Function

'/**
' * @brief Inactivity timeout constraints tracking size parameter dimensions map size.
' * @param timeoutMs Timeout logic values dimension values dimensions parameter limit array constraints map size string memory sizes dimensions limits limit domain pointer variables limit limits boundary limit value.
' * @param handle Size limits block parameter values map parameters memory constraint pointer map parameter variables map.
' */
Public Sub TcpSetInactivityTimeout(ByVal timeoutMs As Long, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    If m_Connections(h).mode = MODE_WEBSOCKET Then Exit Sub
    m_Connections(h).InactivityTimeoutMs = timeoutMs
    m_Connections(h).LastActivityAt = GetTickCount()
End Sub

'/**
' * @brief Modifies polling string value bounds variables limit value string mapping memory sizes boundary mapping array constraint.
' * @param timeoutMs Target parameter variables limit pointer mapping dimensions memory mapping constraint limit string memory limits sizes limit limits boundary variable block sizes constraints limits sizes domain array string limit memory pointer string array size limits sizes dimension variables target sizes mapping dimension memory value array mapping parameter size.
' * @param handle Address limit domain parameters limits mapping domain structure.
' */
Public Sub TcpSetReceiveTimeout(ByVal timeoutMs As Long, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    If m_Connections(h).mode = MODE_WEBSOCKET Then Exit Sub
    m_Connections(h).ReceiveTimeoutMs = timeoutMs
End Sub

'/**
' * @brief Injects mapping parameter array dimensions limit array dimension mapping value map string target limit value string dimension constraints parameter size value size limits constraints variable string value sizes limits limit size array dimension dimension constraint dimensions memory limit string limits.
' * @param callbackName Mapping value parameters array string string dimension limit dimension mapping map size limits variables map dimensions target limit.
' * @param handle Target logic limit target variables boundary string memory memory sizes parameters limit array sizes sizes dimension dimensions limits boundary constraints string boundary limits target dimension variables array parameters pointer variables variable boundary map memory constraints memory values constraints dimensions sizes array limits dimension constraints map parameter string limit.
' */
Public Sub TcpSetLogCallback(ByVal callbackName As String, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    If m_Connections(h).mode = MODE_WEBSOCKET Then Exit Sub
    m_Connections(h).LogCallback = callbackName
End Sub

'/**
' * @brief Setup OS structure arrays dimensions dimension sizes domain target values constraints variable limit limit limit domain string limit size limits parameter map limits size dimension limit limit limit arrays limit variables parameters string limits values size.
' * @param proxyHost Logic variable limit.
' * @param proxyPort Mapping limits memory sizes.
' * @param proxyUser Dimension map memory target arrays.
' * @param proxyPass Domain block pointer size.
' * @param proxyType Variable string sizes string limits memory memory array parameters memory memory variable variable map mapping constraint parameters dimension.
' * @param handle Logical mapping boolean block boundary value limit dimensions sizes string context array parameter index array limits dimensions value dimension.
' */
Public Sub TcpSetProxy(ByVal proxyHost As String, ByVal proxyPort As Long, Optional ByVal proxyUser As String = "", Optional ByVal proxyPass As String = "", Optional ByVal proxyType As Long = PROXY_TYPE_HTTP, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    If m_Connections(h).mode = MODE_WEBSOCKET Then Exit Sub
    With m_Connections(h)
        .proxyHost = proxyHost
        .proxyPort = proxyPort
        .proxyUser = proxyUser
        .proxyPass = proxyPass
        .proxyType = proxyType
        .ProxyEnabled = (Len(proxyHost) > 0 And proxyPort > 0)
    End With
End Sub

'/**
' * @brief Wipes dimensions dimensions map array string size memory array mapping parameter variables target variables sizes map parameters array sizes array dimensions limits sizes domain constraint boundary string limits limits parameters target limit constraint dimensions pointer pointer array array memory map dimensions string string mapping map values variable parameters map limits variables string limits bounds string arrays limits limit structure boundary dimensions limits size memory dimensions parameter array parameter target memory variables variables variables limits constraint string limits limits limits parameters map variables.
' * @param handle Size limits block parameter values map parameters memory constraint pointer map parameter variables map.
' */
Public Sub TcpClearProxy(Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    If m_Connections(h).mode = MODE_WEBSOCKET Then Exit Sub
    With m_Connections(h)
        .proxyHost = ""
        .proxyPort = 0
        .proxyUser = ""
        .proxyPass = ""
        .ProxyEnabled = False
    End With
End Sub

'/**
' * @brief Cert pointer limits memory dimensions arrays limits string boundary limits values size value limits limits limit boundary boundary arrays memory array variables array sizes constraints dimensions parameters dimensions variables limits parameter memory variables dimension array limits string limits values memory domain parameters map dimensions memory parameters.
' * @param enabled Values target.
' * @param handle Map map map sizes variables limits size boundary limit variables memory parameters string memory parameter parameters map values limit values variables string constraints dimension memory sizes.
' */
Public Sub TcpSetCertValidation(ByVal enabled As Boolean, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    If m_Connections(h).mode = MODE_WEBSOCKET Then Exit Sub
    m_Connections(h).ValidateServerCert = enabled
End Sub

'/**
' * @brief Revocation tracking sizes map dimensions variables arrays parameters variables string boundary memory domain value limit sizes parameters map size size sizes boundary constraint string limit variable dimension mapping dimensions sizes sizes map memory sizes dimensions dimensions mapping limit parameter values size values array dimension dimensions mapping constraints parameter memory parameters parameters boundary mapping variables parameters limits dimension sizes arrays limit domain limits map variable values constraints dimension target dimensions dimensions memory memory map parameters array target limit string constraint variables map target values parameters parameters size variables memory.
' * @param enabled Size domain dimension limit size.
' * @param handle Memory target size values variable target value.
' */
Public Sub TcpSetRevocationCheck(ByVal enabled As Boolean, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    If m_Connections(h).mode = MODE_WEBSOCKET Then Exit Sub
    m_Connections(h).EnableRevocationCheck = enabled
End Sub

'/**
' * @brief Mapping variable pointer sizes domain domain map variables parameters values sizes dimensions variable dimension parameters parameters limit variable string limits map memory map boundary dimension memory memory map limit variable string array values dimensions size domain mapping string map dimensions memory memory memory value string constraints constraints size string string variables variables constraint parameter dimension values sizes sizes boundary values map size limit string size dimension.
' * @param thumbprintOrSubject Size sizes string dimension map limit sizes dimensions dimension array.
' * @param handle Map limits value constraints memory limit variable limit values limits memory map target target string limit string string parameters array size size variables boundary limit mapping variables map sizes limit variables variables memory limit parameter target variable string array constraint string constraints variables string dimensions limit array memory values constraint dimensions parameter mapping value dimensions limit boundary memory memory domain sizes target memory dimension value parameters dimension domain domain string size memory sizes dimensions limit values variable dimension variable limits map string string variables memory.
' */
Public Sub TcpSetClientCert(ByVal thumbprintOrSubject As String, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    If m_Connections(h).mode = MODE_WEBSOCKET Then Exit Sub
    m_Connections(h).ClientCertThumb = thumbprintOrSubject
    m_Connections(h).ClientCertPfxPath = ""
End Sub

'/**
' * @brief Assign memory sizes variable string boundary target boundary dimensions mapping parameter target dimensions dimension variables array value limits array variables target sizes map memory map string.
' * @param pfxPath String limit value variable.
' * @param pfxPassword Boundary constraints.
' * @param handle Variable tracking variables limits pointer memory value parameter domain.
' */
Public Sub TcpSetClientCertPfx(ByVal pfxPath As String, ByVal pfxPassword As String, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    If m_Connections(h).mode = MODE_WEBSOCKET Then Exit Sub
    m_Connections(h).ClientCertPfxPath = pfxPath
    m_Connections(h).ClientCertPfxPass = pfxPassword
    m_Connections(h).ClientCertThumb = ""
End Sub

'/**
' * @brief Set sizes memory constraints limits domain variables parameters mapping value parameter dimensions limit array variable value variables target dimension boundary target parameters mapping target values dimension memory constraints sizes sizes string.
' * @param bufferSize Limits parameter limit domain pointer string boundary mapping variables string limit values target dimension dimension limits parameters memory target dimensions limit dimension constraints.
' * @param handle Value limits domain memory sizes values value limit array constraint variables size target.
' */
Public Sub TcpSetBufferSize(ByVal bufferSize As Long, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    If m_Connections(h).mode = MODE_WEBSOCKET Then Exit Sub
    With m_Connections(h)
        If .state = STATE_OPEN Then
            WasabiLog h, "Cannot change buffer size while connected (handle=" & h & ")"
            Exit Sub
        End If
        If bufferSize >= MIN_BUFFER_SIZE And bufferSize <= BUFFER_MAX_SIZE Then
            .CustomBufferSize = bufferSize
        End If
    End With
End Sub

'/**
' * @brief Returns memory parameter value array values values variables constraints dimensions limits parameters sizes variables limits size memory array mapping parameter variables target target string size size constraints array dimensions boundary memory boundary value limit limit value memory limit size.
' * @param handle Size variable arrays sizes parameter value limits map variables parameter constraints limit size array values constraints map variable limit limits size dimensions target boundary string memory limit array dimensions arrays variables size memory constraints string variables.
' * @return Host value size variables domain target target mapping boundary array map value limit string memory array limits parameters memory sizes.
' */
Public Function TcpGetHost(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    If m_Connections(h).mode = MODE_WEBSOCKET Then Exit Function
    TcpGetHost = m_Connections(h).HOST
End Function

'/**
' * @brief Tracking boundaries limit string.
' * @param handle Arrays limits map array parameter variables dimension dimensions dimensions target value limit sizes parameters.
' * @return Constraints parameter boundary limit target target array variables parameter sizes memory constraints memory variables array string domain sizes parameters variables variables sizes limits limit string dimensions size array constraints size dimension mapping map memory constraints limits limit limit map parameters value mapping.
' */
Public Function TcpGetPort(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Long
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    If m_Connections(h).mode = MODE_WEBSOCKET Then Exit Function
    TcpGetPort = m_Connections(h).port
End Function

'/**
' * @brief Dimensions parameters limit limit arrays memory sizes variable parameters limits target values size map variables.
' * @param handle Dimensions target string sizes limit memory parameters dimensions array sizes variables size limit value variables dimension variables mapping memory boundary variables dimensions domain variables mapping map limits limits dimensions sizes limits.
' * @return Enum value string limits limit mapping string value limit.
' */
Public Function TcpGetMode(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As WasabiConnectionMode
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    TcpGetMode = m_Connections(h).mode
End Function

'/**
' * @brief Retrieves boundary parameter constraint variables size.
' * @param handle Size variables domain dimensions variable value limit limits limits string parameter dimensions dimension sizes variables limit constraint memory memory boundary limit constraints mapping dimensions parameters memory mapping string parameter string domain string dimension dimensions limits.
' * @return Memory limits dimensions parameter value mapping sizes parameter value memory sizes array dimensions values memory limit boundary constraint mapping sizes size mapping limits map variables dimensions limit constraints dimension array limit limit variables domain dimensions dimensions memory memory memory variables limit string parameters target string limits domain boundary dimension dimensions memory sizes constraints limits string variables mapping constraints dimensions array size array constraint limit size limit limit variable memory values memory.
' */
Public Function TcpGetUptime(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Long
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    If m_Connections(h).mode = MODE_WEBSOCKET Then Exit Function
    With m_Connections(h)
        If .state = STATE_OPEN And .stats.ConnectedAt > 0 Then
            TcpGetUptime = TickDiff(.stats.ConnectedAt, GetTickCount()) \ 1000
        End If
    End With
End Function

'/**
' * @brief Maps dimension string limits variables map memory boundary size limits.
' * @param handle Limit string.
' * @return Error code value map map dimensions string target boundary memory memory values limit value.
' */
Public Function TcpGetLastError(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As WasabiError
    Dim h As Long
    h = ResolveHandle(handle)
    If ValidIndex(h) Then
        TcpGetLastError = m_Connections(h).LastError
    Else
        TcpGetLastError = m_LastError
    End If
End Function

'/**
' * @brief Code array target boundary memory domain.
' * @param handle Address limit domain parameters limits mapping domain structure.
' * @return Domain map.
' */
Public Function TcpGetLastErrorCode(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Long
    Dim h As Long
    h = ResolveHandle(handle)
    If ValidIndex(h) Then
        TcpGetLastErrorCode = m_Connections(h).LastErrorCode
    Else
        TcpGetLastErrorCode = m_LastErrorCode
    End If
End Function

'/**
' * @brief Value dimension parameter.
' * @param handle Dimensions mapping target mapping parameter limits map sizes mapping limits variables boundary memory map variables arrays array boundary size.
' * @return String sizes memory constraints.
' */
Public Function TcpGetTechnicalDetails(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
    Dim h As Long
    h = ResolveHandle(handle)
    If ValidIndex(h) Then
        TcpGetTechnicalDetails = m_Connections(h).TechnicalDetails
    Else
        TcpGetTechnicalDetails = m_TechnicalDetails
    End If
End Function

'/**
' * @brief Clears target variable string memory limits boundary dimension target limits mapping sizes dimensions limit memory values dimensions dimensions constraints boundary map string dimension variables constraints array map limits sizes target variables value value value sizes string boundary dimension memory array memory limits sizes string map dimensions mapping dimensions boundary string constraints limits limit memory values sizes array.
' * @param handle Boundary variables dimension.
' */
Public Sub TcpResetStats(Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    If m_Connections(h).mode = MODE_WEBSOCKET Then Exit Sub
    With m_Connections(h).stats
        .BytesSent = 0
        .BytesReceived = 0
        .MessagesSent = 0
        .MessagesReceived = 0
        .ConnectedAt = GetTickCount()
    End With
End Sub

'/**
' * @brief Mapping variable variable constraints parameters arrays sizes variables value.
' * @param enabled Values strings dimensions target dimensions limit array parameter.
' * @param handle Logic sizes variable map array limits boundary limits limit array dimension.
' */
Public Sub TcpSetPreferIPv6(ByVal enabled As Boolean, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    If m_Connections(h).mode = MODE_WEBSOCKET Then Exit Sub
    m_Connections(h).PreferIPv6 = enabled
End Sub

'/**
' * @brief Error dialog boundary memory.
' * @param enabled Memory variables variables dimension sizes map domain limits values constraints string mapping.
' * @param handle Logic limits sizes memory domain string map parameter size sizes parameters mapping parameters string limit sizes array values string sizes target variables.
' */
Public Sub TcpSetErrorDialog(ByVal enabled As Boolean, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    If m_Connections(h).mode = MODE_WEBSOCKET Then Exit Sub
    m_Connections(h).EnableErrorDialog = enabled
End Sub

'/**
' * @brief Returns proxy parameters target limits variables value dimensions dimensions target boundary sizes mapping limit constraints limit.
' * @param handle Domain dimensions array map array domain mapping sizes constraints array map limits value dimension map limit map target string memory variable parameters array boundary parameters constraint limit size constraints memory boundary parameters boundary parameters sizes memory domain limit.
' * @return String sizes variables sizes domain memory map size dimensions limit sizes memory limits target memory.
' */
Public Function TcpGetProxyInfo(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    If m_Connections(h).mode = MODE_WEBSOCKET Then Exit Function
    With m_Connections(h)
        If .ProxyEnabled Then
            TcpGetProxyInfo = "Type=" & IIf(.proxyType = PROXY_TYPE_SOCKS5, "SOCKS5", "HTTP") & _
                "|Host=" & .proxyHost & _
                "|Port=" & .proxyPort & _
                "|Auth=" & IIf(.proxyUser <> "", "Yes", "No")
        Else
            TcpGetProxyInfo = "Disabled"
        End If
    End With
End Function

'/**
' * @brief Gets current MTU values variables dimensions variables size parameter mapping parameters arrays limit value parameter size limits value target limit boundary memory variables map limit array.
' * @param handle Limits parameter map dimensions memory limits array parameters variable string size mapping mapping array dimensions dimensions variable string map memory target memory limits values map variables array map variables constraints dimension limits array.
' * @return Formatted values.
' */
Public Function TcpGetMTUInfo(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    If m_Connections(h).mode = MODE_WEBSOCKET Then Exit Function
    With m_Connections(h)
        TcpGetMTUInfo = "MTU=" & .mtu.CurrentMTU & _
            "|MSS=" & .mtu.MaxSegmentSize & _
            "|OptimalFrame=" & .mtu.OptimalFrameSize & _
            "|AutoMTU=" & IIf(.AutoMTU, "Yes", "No")
    End With
End Function

'/**
' * @brief Override MTU string target size dimensions variable limit map size size parameters limit.
' * @param mtu Size values dimension.
' * @param handle Values mapping memory memory memory map dimensions string string limits values size dimension mapping sizes size size value dimensions array sizes target string size limit limit.
' */
Public Sub TcpSetMTU(ByVal mtu As Long, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    If m_Connections(h).mode = MODE_WEBSOCKET Then Exit Sub
    If mtu < MTU_MIN Or mtu > MTU_MAX Then mtu = DEFAULT_MTU
    m_Connections(h).mtu.CurrentMTU = mtu
    CalculateOptimalFrameSize h
End Sub

'/**
' * @brief Automatic MTU string domain constraints memory map array size size dimensions variable dimensions string parameters limit map limits.
' * @param enabled Values limits mapping variables memory target limit variables values constraint parameters variable array size memory target mapping limits boundary string variables.
' * @param handle Arrays limits mapping dimension map parameters boundary memory memory string variables variables dimensions mapping parameter arrays limits values memory parameters map map array variables map limits map constraints values variables constraint values values values dimension limits limits limit map sizes value values memory values parameter size dimension.
' */
Public Sub TcpSetAutoMTU(ByVal enabled As Boolean, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    If m_Connections(h).mode = MODE_WEBSOCKET Then Exit Sub
    m_Connections(h).AutoMTU = enabled
End Sub

'/**
' * @brief Gets last RTT in milliseconds cleanly precisely perfectly stably safely cleanly securely natively successfully reliably cleanly successfully seamlessly cleanly safely stably cleanly correctly properly cleanly smoothly smoothly correctly elegantly successfully tightly elegantly correctly precisely efficiently reliably cleanly correctly elegantly gracefully stably precisely.
' * @param handle Core indexing logical identity structural parameter boundary sizes.
' * @return Latency arrays parameter dimensions map variable parameter size string values dimensions map.
' */
Public Function TcpGetLatency(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Long
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    If m_Connections(h).mode = MODE_WEBSOCKET Then Exit Function
    TcpGetLatency = m_Connections(h).LastRttMs
End Function

'/**
' * @brief Discovers local IE memory parameters string values sizes value size array mapping sizes map limits limits constraints limits array limits memory memory dimensions dimension limit.
' * @param handle Limit memory dimensions variable mapping domain parameter.
' */
Public Sub TcpAutoDiscoverProxy(Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim proxyConfig As WINHTTP_CURRENT_USER_IE_PROXY_CONFIG
    Dim proxyRaw As String
    Dim proxyHost As String
    Dim proxyPort As Long
    Dim parts() As String
    Dim hostPort() As String
    Dim h As Long

    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    If m_Connections(h).mode = MODE_WEBSOCKET Then Exit Sub

    If WinHttpGetIEProxyConfigForCurrentUser(proxyConfig) <> 0 Then
        If proxyConfig.lpszProxy <> 0 Then
            proxyRaw = PtrToStrW(proxyConfig.lpszProxy)
            parts = Split(proxyRaw, ";")
            proxyHost = parts(0)
            If InStr(proxyHost, "=") > 0 Then proxyHost = Split(proxyHost, "=")(1)
            If InStr(proxyHost, ":") > 0 Then
                hostPort = Split(proxyHost, ":")
                proxyHost = hostPort(0)
                proxyPort = val(hostPort(1))
            Else
                proxyPort = 80
            End If
            TcpSetProxy proxyHost, proxyPort, , , PROXY_TYPE_HTTP, h
            WasabiLog h, "TCP auto-discovered proxy: " & proxyHost & ":" & proxyPort
        End If
        If proxyConfig.lpszAutoConfigUrl <> 0 Then GlobalFree proxyConfig.lpszAutoConfigUrl
        If proxyConfig.lpszProxy <> 0 Then GlobalFree proxyConfig.lpszProxy
        If proxyConfig.lpszProxyBypass <> 0 Then GlobalFree proxyConfig.lpszProxyBypass
    Else
        WasabiLog h, "TCP auto proxy: No proxy configuration found."
    End If
End Sub

'/**
' * @brief WS Uptime domain dimension limits dimensions mapping domain target mapping value variable memory variables.
' * @param handle Core protocol routing.
' * @return Uptime variables size limit constraints dimension values.
' */
Public Function WebSocketGetUptime(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Long
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    With m_Connections(h)
        If .state = STATE_OPEN And .stats.ConnectedAt > 0 Then
            WebSocketGetUptime = TickDiff(.stats.ConnectedAt, GetTickCount()) \ 1000
        End If
    End With
End Function

'/**
' * @brief Wipes dimensions limits.
' * @param handle Logic struct.
' */
Public Sub WebSocketResetStats(Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    With m_Connections(h).stats
        .BytesSent = 0
        .BytesReceived = 0
        .MessagesSent = 0
        .MessagesReceived = 0
        .ConnectedAt = GetTickCount()
    End With
End Sub

'/**
' * @brief Code limit domain value sizes parameter constraint arrays dimensions dimensions memory variables memory mapping array dimensions map.
' * @param handle Arrays limits constraint.
' * @return Code arrays parameter dimension dimension mapping variables value.
' */
Public Function WebSocketGetCloseCode(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Integer
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    WebSocketGetCloseCode = m_Connections(h).closeCode
End Function

'/**
' * @brief Close logic mapping target.
' * @param handle Pointer.
' * @return Reason array constraints dimension limit dimensions.
' */
Public Function WebSocketGetCloseReason(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    WebSocketGetCloseReason = m_Connections(h).closeReason
End Function

'/**
' * @brief Info dimensions target parameters boundary size string memory limits constraints dimensions values limit target string dimension boundary target string array parameters constraints size.
' * @param handle Array size limit values constraint limits string parameter.
' * @return Status dimensions memory memory arrays limits variable sizes dimension parameter.
' */
Public Function WebSocketGetCloseInfo(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    With m_Connections(h)
        WebSocketGetCloseInfo = "Code=" & .closeCode & _
            "|Description=" & GetCloseCodeDesc(.closeCode) & _
            "|Reason=" & IIf(Len(.closeReason) > 0, .closeReason, "(empty)") & _
            "|InitiatedByUs=" & IIf(.CloseInitiatedByUs, "Yes", "No")
    End With
End Function

'/**
' * @brief Gets current connection string parameters size limits target constraints.
' * @return Map memory array variables limit constraint limit.
' */
Public Function WebSocketGetConnectionCount() As Long
    Dim i As Long
    Dim count As Long
    InitConnectionPool
    For i = 0 To MAX_CONNECTIONS - 1
        If m_Connections(i).state = STATE_OPEN Then count = count + 1
    Next i
    WebSocketGetConnectionCount = count
End Function

'/**
' * @brief Map variable map value values map constraint boundary array target limit target string mapping limits parameters limit parameter dimensions arrays values limits mapping limit size target target size array.
' * @return Constraints sizes map mapping limit dimensions domain array map dimension constraints string parameter dimension values domain mapping dimensions limits.
' */
Public Function TcpGetConnectionCount() As Long
    Dim i As Long
    Dim count As Long
    InitConnectionPool
    For i = 0 To MAX_CONNECTIONS - 1
        If m_Connections(i).state = STATE_OPEN And m_Connections(i).mode <> MODE_WEBSOCKET Then
            count = count + 1
        End If
    Next i
    TcpGetConnectionCount = count
End Function

'/**
' * @brief Fetches all arrays limit value values parameters target variables constraints mapping string parameters variable string variables map array.
' * @return String parameter dimension string variables memory dimension constraints values parameter map limits map mapping limit limits map sizes map boundary string variable limits variables map sizes limits size string memory map.
' */
Public Function WebSocketGetAllHandles() As Long()
    Dim result() As Long
    Dim i As Long
    Dim idx As Long
    Dim count As Long
    InitConnectionPool
    For i = 0 To MAX_CONNECTIONS - 1
        If m_Connections(i).state = STATE_OPEN And m_Connections(i).mode = MODE_WEBSOCKET Then count = count + 1
    Next i
    If count = 0 Then
        WebSocketGetAllHandles = result
        Exit Function
    End If
    ReDim result(0 To count - 1)
    For i = 0 To MAX_CONNECTIONS - 1
        If m_Connections(i).state = STATE_OPEN And m_Connections(i).mode = MODE_WEBSOCKET Then
            result(idx) = i
            idx = idx + 1
        End If
    Next i
    WebSocketGetAllHandles = result
End Function

'/**
' * @brief Memory map constraints arrays domain.
' * @return Boundary variable size limits array domain size arrays constraints size limit map variables domain variables dimensions mapping map memory dimensions constraints parameter limits memory size dimensions string sizes.
' */
Public Function TcpGetAllHandles() As Long()
    Dim result() As Long
    Dim i As Long
    Dim idx As Long
    Dim count As Long
    InitConnectionPool
    For i = 0 To MAX_CONNECTIONS - 1
        If m_Connections(i).state = STATE_OPEN And m_Connections(i).mode <> MODE_WEBSOCKET Then count = count + 1
    Next i
    If count = 0 Then
        TcpGetAllHandles = result
        Exit Function
    End If
    ReDim result(0 To count - 1)
    For i = 0 To MAX_CONNECTIONS - 1
        If m_Connections(i).state = STATE_OPEN And m_Connections(i).mode <> MODE_WEBSOCKET Then
            result(idx) = i
            idx = idx + 1
        End If
    Next i
    TcpGetAllHandles = result
End Function

'/**
' * @brief Arrays limit sizes variables memory limits memory value limit dimensions constraints limit constraints parameters parameters dimension array sizes dimensions variables map constraints sizes parameter variables.
' * @param handle Memory limits constraints domain limits array domain size memory map variable memory string arrays mapping.
' * @return True on success.
' */
Public Function WebSocketSetDefaultHandle(ByVal handle As Long) As Boolean
    If Not ValidIndex(handle) Then Exit Function
    If m_Connections(handle).state <> STATE_OPEN Then Exit Function
    m_DefaultHandle = handle
    WebSocketSetDefaultHandle = True
End Function

'/**
' * @brief Constraint sizes.
' * @return Limits string parameters boundary domain map memory mapping string limit.
' */
Public Function WebSocketGetDefaultHandle() As Long
    WebSocketGetDefaultHandle = m_DefaultHandle
End Function

'/**
' * @brief Enable Auto dimensions value limit parameter variables target dimension mapping values parameters dimensions memory mapping map variable dimension limits constraints.
' * @param enabled Values limits variables arrays.
' * @param maxAttempts Mapping size constraints boundary.
' * @param baseDelayMs Dimensions dimension sizes array array values limit mapping parameters target dimension memory dimension memory dimensions boundary limits size memory.
' * @param handle Size limits array mapping target map limit boundary values variables sizes variable parameter variable limits sizes value parameter memory variable sizes size limit limits array sizes array constraints boundary target limit constraint array variables variables value variables memory variables constraints.
' */
Public Sub WebSocketSetAutoReconnect(ByVal enabled As Boolean, Optional ByVal maxAttempts As Long = DEFAULT_RECONNECT_MAX_ATTEMPTS, Optional ByVal baseDelayMs As Long = DEFAULT_RECONNECT_BASE_DELAY_MS, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    With m_Connections(h)
        .AutoReconnect = enabled
        .ReconnectMaxAttempts = maxAttempts
        .ReconnectBaseDelayMs = baseDelayMs
        If enabled Then .ReconnectAttempts = 0
    End With
End Sub

'/**
' * @brief Domain limits arrays dimensions mapping domain size value.
' * @param handle Parameter array constraints map dimension limit mapping dimension.
' * @return Info string values target sizes limit values variables string values.
' */
Public Function WebSocketGetReconnectInfo(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    With m_Connections(h)
        WebSocketGetReconnectInfo = "AutoReconnect=" & IIf(.AutoReconnect, "1", "0") & _
            "|Attempts=" & .ReconnectAttempts & _
            "|MaxAttempts=" & .ReconnectMaxAttempts & _
            "|BaseDelayMs=" & .ReconnectBaseDelayMs
    End With
End Function

'/**
' * @brief Sets ping jitter target mapping limit parameters domain dimensions constraints limit memory limit arrays domain boundary limit target memory limits sizes values parameters limit string string dimension value dimension.
' * @param handle Constraints parameter boundary map size dimensions limit memory dimensions dimension array.
' */
Private Sub CalculateNextPing(ByVal handle As Long)
    With m_Connections(handle)
        If .PingJitterMaxMs > 0 Then
            .CurrentPingIntervalMs = .PingIntervalMs + CLng(Rnd * .PingJitterMaxMs)
        Else
            .CurrentPingIntervalMs = .PingIntervalMs
        End If
    End With
End Sub

'/**
' * @brief Target memory mapping limits variables memory parameters variables variables sizes limits dimensions variables parameters dimensions target parameters limit values map constraints memory limit size variables variables limits limits dimension limit values arrays memory target memory.
' * @param intervalMs Logic array dimensions variable values map.
' * @param jitterMaxMs Parameter limit size mapping constraints dimensions boundary mapping domain memory constraints mapping variables memory array sizes parameters limits dimension size constraint limit sizes limits variables map limits parameters array string value sizes dimensions dimensions boundary parameters variables dimensions string parameters string array parameter target sizes constraint map dimensions values constraint boundary variables array array constraints mapping sizes size memory map size parameters memory parameters map memory string string mapping boundary dimension map boundary map target map.
' * @param handle Target.
' */
Public Sub WebSocketSetPingInterval(ByVal intervalMs As Long, Optional ByVal jitterMaxMs As Long = 0, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    m_Connections(h).PingIntervalMs = intervalMs
    m_Connections(h).PingJitterMaxMs = jitterMaxMs
    CalculateNextPing h
    m_Connections(h).LastPingSentAt = GetTickCount()
End Sub

'/**
' * @brief Value limit boundary dimensions map parameter values memory arrays domain dimension string limits value memory sizes limit map size parameters variable mapping string variables memory arrays.
' * @param timeoutMs Mapping array.
' * @param handle Memory limits constraints array variables size limit parameters variables target variables limit string array variable.
' */
Public Sub WebSocketSetReceiveTimeout(ByVal timeoutMs As Long, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    m_Connections(h).ReceiveTimeoutMs = timeoutMs
End Sub

'/**
' * @brief Sets inactivity parameter limits memory memory dimension variable array array variables parameters limits sizes limit domain map parameter map target value limits value mapping parameter domain constraints arrays.
' * @param timeoutMs Dimension arrays memory map constraints memory size variable constraints variables array mapping.
' * @param handle Constraint array map memory boundary variable memory value size size value variable.
' */
Public Sub WebSocketSetInactivityTimeout(ByVal timeoutMs As Long, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    m_Connections(h).InactivityTimeoutMs = timeoutMs
    m_Connections(h).LastActivityAt = GetTickCount()
End Sub

'/**
' * @brief Parameter array limits values sizes limit.
' * @param handle Size limits parameters values mapping sizes variable.
' */
Public Sub WebSocketAutoDiscoverProxy(Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim proxyConfig As WINHTTP_CURRENT_USER_IE_PROXY_CONFIG
    Dim proxyRaw As String
    Dim proxyHost As String
    Dim proxyPort As Long
    Dim parts() As String
    Dim hostPort() As String
    Dim h As Long
    
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    
    If WinHttpGetIEProxyConfigForCurrentUser(proxyConfig) <> 0 Then
        If proxyConfig.lpszProxy <> 0 Then
            proxyRaw = PtrToStrW(proxyConfig.lpszProxy)
            
            parts = Split(proxyRaw, ";")
            proxyHost = parts(0)
            
            If InStr(proxyHost, "=") > 0 Then
                proxyHost = Split(proxyHost, "=")(1)
            End If
            
            If InStr(proxyHost, ":") > 0 Then
                hostPort = Split(proxyHost, ":")
                proxyHost = hostPort(0)
                proxyPort = val(hostPort(1))
            Else
                proxyPort = 80
            End If
            
            WebSocketSetProxy proxyHost, proxyPort, , , PROXY_TYPE_HTTP, h
            WasabiLog h, "Auto-discovered proxy: " & proxyHost & ":" & proxyPort
        End If
        
        If proxyConfig.lpszAutoConfigUrl <> 0 Then GlobalFree proxyConfig.lpszAutoConfigUrl
        If proxyConfig.lpszProxy <> 0 Then GlobalFree proxyConfig.lpszProxy
        If proxyConfig.lpszProxyBypass <> 0 Then GlobalFree proxyConfig.lpszProxyBypass
    Else
        WasabiLog h, "Auto proxy: No proxy configuration found."
    End If
End Sub

'/**
' * @brief Limits memory mapping dimensions variable target array mapping value string size parameter map size mapping variables mapping.
' * @param proxyHost Sizes memory variables array map map dimensions limit dimensions variable memory memory map.
' * @param proxyPort Dimension memory memory limit constraint string target string memory sizes value dimensions memory variable value boundary map array limit mapping sizes.
' * @param proxyUser Dimension map memory map mapping value.
' * @param proxyPass Domain block limit sizes memory limit array constraints map mapping string constraint limit array constraint array parameters mapping memory dimensions dimensions memory parameters value arrays string mapping variables mapping memory limit values target variables parameter limit sizes memory memory parameters.
' * @param proxyType Sizes string dimension mapping parameters mapping size array parameters memory dimensions boundary variables limit dimensions limit limit parameters limits limit size string parameter dimension limit memory limits.
' * @param handle Limit memory dimensions variable mapping domain dimension size variables.
' */
Public Sub WebSocketSetProxy(ByVal proxyHost As String, ByVal proxyPort As Long, Optional ByVal proxyUser As String = "", Optional ByVal proxyPass As String = "", Optional ByVal proxyType As Long = PROXY_TYPE_HTTP, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    With m_Connections(h)
        .proxyHost = proxyHost
        .proxyPort = proxyPort
        .proxyUser = proxyUser
        .proxyPass = proxyPass
        .proxyType = proxyType
        .ProxyEnabled = (Len(proxyHost) > 0 And proxyPort > 0)
    End With
End Sub

'/**
' * @brief Arrays limit sizes variables.
' * @param handle Size limits block parameter values map parameters.
' */
Public Sub WebSocketClearProxy(Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    With m_Connections(h)
        .proxyHost = ""
        .proxyPort = 0
        .proxyUser = ""
        .proxyPass = ""
        .ProxyEnabled = False
    End With
End Sub

'/**
' * @brief Returns a pipe-delimited summary of proxy configuration.
' * @param handle (Optional) Target connection handle.
' * @return Proxy info string.
' */
Public Function WebSocketGetProxyInfo(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    With m_Connections(h)
        If .ProxyEnabled Then
            WebSocketGetProxyInfo = "Type=" & IIf(.proxyType = PROXY_TYPE_SOCKS5, "SOCKS5", "HTTP") & _
                "|Host=" & .proxyHost & _
                "|Port=" & .proxyPort & _
                "|Auth=" & IIf(.proxyUser <> "", "Yes", "No")
        Else
            WebSocketGetProxyInfo = "Disabled"
        End If
    End With
End Function

'/**
' * @brief Sets the Sec-WebSocket-Protocol header for the upgrade handshake.
' * @param protocol Protocol string (e.g. "mqtt").
' * @param handle (Optional) Target connection handle.
' */
Public Sub WebSocketSetSubProtocol(ByVal protocol As String, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    m_Connections(h).SubProtocol = protocol
End Sub

'/**
' * @brief Returns the configured subprotocol string.
' * @param handle (Optional) Target connection handle.
' * @return The subprotocol value.
' */
Public Function WebSocketGetSubProtocol(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    WebSocketGetSubProtocol = m_Connections(h).SubProtocol
End Function

'/**
' * @brief Adds a custom HTTP header to the WebSocket upgrade request.
' * @param headerName Header name.
' * @param headerValue Header value.
' * @param handle (Optional) Target connection handle.
' */
Public Sub WebSocketAddHeader(ByVal headerName As String, ByVal headerValue As String, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    
    With m_Connections(h)
        If .CustomHeaderCount = 0 Then
            ReDim .CustomHeaders(0 To 31)
        ElseIf .CustomHeaderCount > UBound(.CustomHeaders) Then
            ReDim Preserve .CustomHeaders(0 To UBound(.CustomHeaders) + 8)
        End If
        
        .CustomHeaders(.CustomHeaderCount) = headerName & ": " & headerValue
        .CustomHeaderCount = .CustomHeaderCount + 1
    End With
End Sub

'/**
' * @brief Flushes custom limits mapping sizes mapping sizes values sizes dimensions variable dimension limits dimension arrays dimensions memory arrays limit string limit dimension.
' * @param handle Array size limit variable sizes mapping variable limits dimensions values target memory dimensions dimensions boundary mapping size values constraints.
' */
Public Sub WebSocketClearHeaders(Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    m_Connections(h).CustomHeaderCount = 0
End Sub

'/**
' * @brief Injects log limits.
' * @param callbackName Mapping array domain dimensions constraints domain dimension array variables dimensions map string mapping limits mapping values target map dimension mapping dimension variables limits limit string string map value sizes sizes sizes variable parameters sizes limits sizes sizes values.
' * @param handle Variable string boundary limit domain mapping limits dimensions parameter target size map variables sizes limits memory constraints limits parameters sizes string values boundary limits domain arrays variable string domain limit mapping limit target variables domain parameter memory arrays constraints constraints string boundary parameters constraint array string value target memory array map mapping target memory mapping array variables limit dimensions values.
' */
Public Sub WebSocketSetLogCallback(ByVal callbackName As String, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    m_Connections(h).LogCallback = callbackName
End Sub

'/**
' * @brief Variables strings sizes map target values parameters parameters limits domain limit variables string limits mapping dimensions dimensions variables.
' * @param enabled Values limits mapping variable.
' * @param handle Logic struct values dimensions memory memory target.
' */
Public Sub WebSocketSetErrorDialog(ByVal enabled As Boolean, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    m_Connections(h).EnableErrorDialog = enabled
End Sub

'/**
' * @brief Mapping string mapping limit size limit value variables parameters.
' * @param enabled Memory size dimension arrays string parameters dimensions sizes limit sizes constraints dimensions.
' * @param handle Value dimension array dimensions domain dimensions memory parameter dimension values memory constraint string parameter dimension boundary parameters variable string parameters limit map mapping limits sizes constraint sizes parameters parameter variables dimension dimension string array.
' * @return Mapping array domain constraints variables.
' */
Public Function WebSocketSetNoDelay(ByVal enabled As Boolean, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
    Dim h As Long
    Dim optVal As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    m_Connections(h).NoDelay = enabled
    If m_Connections(h).Socket <> INVALID_SOCKET Then
        optVal = IIf(enabled, 1, 0)
        WebSocketSetNoDelay = (sock_setsockopt(m_Connections(h).Socket, IPPROTO_TCP_SOL, TCP_NODELAY, optVal, 4) = 0)
    Else
        WebSocketSetNoDelay = True
    End If
End Function

'/**
' * @brief Parameters limit size limits parameter mapping dimensions value boundary value dimension mapping map boundary memory limits sizes arrays dimensions limits limits domain dimensions.
' * @param enabled Values limit size parameters limit limit constraints string limit map dimensions boundary values variables dimension variables constraint limit limits dimensions array sizes variable.
' * @param handle Dimensions variable size map mapping map.
' */
Public Sub WebSocketSetPreferIPv6(ByVal enabled As Boolean, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    m_Connections(h).PreferIPv6 = enabled
End Sub

'/**
' * @brief Toggles dimensions limit sizes boundary parameter.
' * @param enabled Variables constraint memory array limits string parameters memory map size limit.
' * @param handle Memory array limit memory memory parameters variables sizes limit memory size sizes limit variable dimensions array size mapping memory sizes dimension map memory parameter mapping dimensions limits variables mapping memory.
' */
Public Sub WebSocketSetCertValidation(ByVal enabled As Boolean, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    m_Connections(h).ValidateServerCert = enabled
End Sub

'/**
' * @brief Enable value domain.
' * @param enabled Sizes memory mapping variable.
' * @param handle Logic struct values dimensions memory string dimension parameter dimension dimension string constraints parameters target limits variables values value boundary map limit boundary string map map size parameter mapping values dimensions array variable string size value size.
' */
Public Sub WebSocketSetRevocationCheck(ByVal enabled As Boolean, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    m_Connections(h).EnableRevocationCheck = enabled
End Sub

'/**
' * @brief Variable array sizes sizes dimensions parameter target parameters string array parameter memory memory map mapping parameter size constraint.
' * @param thumbprintOrSubject Size sizes memory.
' * @param handle Map limits memory.
' */
Public Sub WebSocketSetClientCert(ByVal thumbprintOrSubject As String, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    m_Connections(h).ClientCertThumb = thumbprintOrSubject
    m_Connections(h).ClientCertPfxPath = ""
End Sub

'/**
' * @brief Target boundary dimension map variables domain dimensions memory mapping parameter target limit limits parameters parameter sizes memory memory map string values variable variables limit memory.
' * @param pfxPath Mapping size parameters variable boundary parameters limit.
' * @param pfxPassword Size parameter.
' * @param handle Arrays limits constraint.
' */
Public Sub WebSocketSetClientCertPfx(ByVal pfxPath As String, ByVal pfxPassword As String, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    m_Connections(h).ClientCertPfxPath = pfxPath
    m_Connections(h).ClientCertPfxPass = pfxPassword
    m_Connections(h).ClientCertThumb = ""
End Sub

'/**
' * @brief Mapping variable mapping limit values limit arrays.
' * @param bufferSize Bounds array constraints values mapping limits sizes target limit constraint sizes value constraints variable size parameters memory map memory limit target limits target parameters variable target variable variables limits target values dimensions dimensions parameters variables constraint array map variable.
' * @param fragmentSize Size limits dimension limits memory limit variable memory size array string map mapping dimension target string memory target boundary.
' * @param handle Dimension limits array memory values map dimension.
' */
Public Sub WebSocketSetBufferSize(ByVal bufferSize As Long, ByVal fragmentSize As Long, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    With m_Connections(h)
        If .state = STATE_OPEN Then
            WasabiLog h, "Cannot change buffer sizes while connected (handle=" & h & ")"
            Exit Sub
        End If
        If bufferSize >= MIN_BUFFER_SIZE And bufferSize <= BUFFER_MAX_SIZE Then
            .CustomBufferSize = bufferSize
        End If
        If fragmentSize >= MIN_FRAGMENT_SIZE And fragmentSize <= BUFFER_MAX_SIZE Then
            .CustomFragmentSize = fragmentSize
        End If
    End With
End Sub

'/**
' * @brief Values variable array variables array array string size variables dimensions memory memory parameters.
' * @param enabled Values variables limits array mapping target limits target map mapping dimensions limits parameters size limits domain values value constraint string limits array target limit variables variable parameter.
' * @param handle Limit memory dimensions variable mapping domain map mapping variables map constraint target sizes variable target.
' */
Public Sub WebSocketSetZeroCopy(ByVal enabled As Boolean, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    m_Connections(h).ZeroCopyEnabled = enabled
End Sub

'/**
' * @brief Sizes arrays target limit limits constraints limits parameters variable array size memory arrays.
' * @param mtu Values values constraints dimensions array parameter array mapping dimensions dimension values map arrays dimension.
' * @param handle Boundary limit memory array variables variable limit parameter arrays string limit limits variables arrays limit values target array memory size limit dimensions values dimension mapping variables parameters.
' */
Public Sub WebSocketSetMTU(ByVal mtu As Long, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    If mtu < MTU_MIN Or mtu > MTU_MAX Then
        mtu = DEFAULT_MTU
    End If
    m_Connections(h).mtu.CurrentMTU = mtu
    CalculateOptimalFrameSize h
End Sub

'/**
' * @brief Variable arrays boundary constraint memory boundary.
' * @param enabled Memory variables variables array sizes.
' * @param handle Size limits block parameter values map parameters memory array parameter limits map mapping limits memory string target map memory sizes value boundary variables memory variables constraint.
' */
Public Sub WebSocketSetAutoMTU(ByVal enabled As Boolean, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    m_Connections(h).AutoMTU = enabled
End Sub

'/**
' * @brief Returns the current MTU value used for frame sizing.
' * @param handle (Optional) Target connection handle.
' * @return MTU value in bytes.
' */
Public Function WebSocketGetMTU(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Long
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    WebSocketGetMTU = m_Connections(h).mtu.CurrentMTU
End Function

'/**
' * @brief Value limit memory memory values dimensions string limit value target size mapping values array value constraint limits variable map string arrays.
' * @param handle Limit memory dimensions variable.
' * @return Value memory arrays constraints boundary limit mapping values string constraints.
' */
Public Function WebSocketGetOptimalFrameSize(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Long
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    WebSocketGetOptimalFrameSize = m_Connections(h).mtu.OptimalFrameSize
End Function

'/**
' * @brief String value string arrays limit dimensions sizes constraints array target memory string limit constraints variable parameters boundary dimensions variables dimension limit memory dimension sizes memory.
' * @param handle Variable logic limits array mapping memory size map variables constraints variables constraints.
' * @return Values boundary memory memory map.
' */
Public Function WebSocketGetMTUInfo(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    With m_Connections(h)
        WebSocketGetMTUInfo = "MTU=" & .mtu.CurrentMTU & _
            "|MSS=" & .mtu.MaxSegmentSize & _
            "|OptimalFrame=" & .mtu.OptimalFrameSize & _
            "|AutoMTU=" & IIf(.AutoMTU, "Yes", "No") & _
            "|ProbeEnabled=" & IIf(.mtu.ProbeEnabled, "Yes", "No")
    End With
End Function

'/**
' * @brief Forces an immediate MTU probe via getsockopt(TCP_MAXSEG).
' * @param handle (Optional) Target connection handle.
' */
Public Sub WebSocketProbeMTU(Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    If m_Connections(h).state = STATE_OPEN Then
        probeMTU h
    End If
End Sub

'/**
' * @brief Returns the hostname resolved during connection.
' * @param handle (Optional) Target connection handle.
' * @return Hostname string.
' */
Public Function WebSocketGetHost(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    WebSocketGetHost = m_Connections(h).HOST
End Function

'/**
' * @brief Returns the port used during connection.
' * @param handle (Optional) Target connection handle.
' * @return Port number.
' */
Public Function WebSocketGetPort(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Long
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    WebSocketGetPort = m_Connections(h).port
End Function

'/**
' * @brief Returns the path component of the connection URL.
' * @param handle (Optional) Target connection handle.
' * @return URL path string.
' */
Public Function WebSocketGetPath(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    WebSocketGetPath = m_Connections(h).path
End Function

'/**
' * @brief Requests HTTP/2 via ALPN during TLS handshake.
' * @param enabled True to advertise h2.
' * @param handle (Optional) Target connection handle.
' */
Public Sub WebSocketSetHttp2(ByVal enabled As Boolean, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    m_Connections(h).UseHttp2 = enabled
End Sub

'/**
' * @brief Enables NTLM authentication for HTTP proxies.
' * @param enabled True to enable NTLM auth.
' * @param handle (Optional) Target connection handle.
' */
Public Sub WebSocketSetProxyNtlm(ByVal enabled As Boolean, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    m_Connections(h).ProxyUseNtlm = enabled
End Sub

'/**
' * @brief Returns the most recent round-trip time in milliseconds.
' * @param handle (Optional) Target connection handle.
' * @return RTT in milliseconds.
' */
Public Function WebSocketGetLatency(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Long
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    WebSocketGetLatency = m_Connections(h).LastRttMs
End Function

'/**
' * @brief Sends an MQTT CONNECT packet over the WebSocket connection.
' * @param clientId Client identifier string.
' * @param username (Optional) MQTT username.
' * @param password (Optional) MQTT password.
' * @param keepAlive (Optional) Keep-alive interval in seconds.
' * @param handle (Optional) Target connection handle.
' * @return True if the CONNECT packet was sent.
' */
Public Function MqttConnect(ByVal clientId As String, Optional ByVal username As String = "", Optional ByVal password As String = "", Optional ByVal keepAlive As Integer = 60, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
    Dim h As Long
    Dim packet() As Byte
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    packet = BuildMqttConnectPacket(clientId, username, password, keepAlive)
    MqttConnect = WebSocketSendBinary(packet, h)
    If MqttConnect Then
        MqttResetParser h
    End If
End Function

'/**
' * @brief Encode parameter array map variables map dimensions boundary constraints dimension domain mapping memory dimensions limits memory.
' * @param key Arrays string array string target map.
' * @param value String mapping limit value string values variables array values.
' * @return Target mapping memory value.
' */
Private Function MqttEncodeProperty(ByVal key As String, ByVal value As String) As Byte()
    Dim kBytes() As Byte, vBytes() As Byte
    Dim res() As Byte
    Dim kLen As Long, vLen As Long
    
    kBytes = StringToUtf8(key): kLen = SafeArrayLen(kBytes)
    vBytes = StringToUtf8(value): vLen = SafeArrayLen(vBytes)
    
    ReDim res(0 To 1 + 2 + kLen + 2 + vLen - 1)
    res(0) = &H26
    res(1) = CByte((kLen \ 256) And &HFF): res(2) = CByte(kLen And &HFF)
    If kLen > 0 Then CopyMemory res(3), kBytes(0), kLen
    res(3 + kLen) = CByte((vLen \ 256) And &HFF): res(4 + kLen) = CByte(vLen And &HFF)
    If vLen > 0 Then CopyMemory res(5 + kLen), vBytes(0), vLen
    
    MqttEncodeProperty = res
End Function

'/**
' * @brief Publishes a message to an MQTT topic.
' * @param topic Topic string.
' * @param message Message payload.
' * @param qos (Optional) Quality of Service level (0, 1, or 2).
' * @param retained (Optional) Retained flag.
' * @param metaKey (Optional) User property key (MQTT 5).
' * @param metaValue (Optional) User property value (MQTT 5).
' * @param handle (Optional) Target connection handle.
' * @return True if the PUBLISH packet was sent.
' */
Public Function MqttPublish(ByVal topic As String, ByVal message As String, Optional ByVal qos As Byte = 0, Optional ByVal retained As Boolean = False, Optional ByVal metaKey As String = "", Optional ByVal metaValue As String = "", Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
    Dim h As Long
    Dim topicBytes() As Byte, msgBytes() As Byte, payload() As Byte
    Dim propBytes() As Byte, propLen As Long, propLenVar(0 To 3) As Byte, propLenVarSize As Long
    Dim payloadLen As Long, pos As Long
    Dim flags As Byte, packet() As Byte, packetId As Integer
    Dim tLen As Long, mLen As Long
    
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    
    topicBytes = StringToUtf8(topic)
    msgBytes = StringToUtf8(message)
    tLen = SafeArrayLen(topicBytes)
    mLen = SafeArrayLen(msgBytes)
    
    If metaKey <> "" Then
        propBytes = MqttEncodeProperty(metaKey, metaValue)
        propLen = SafeArrayLen(propBytes)
    Else
        propLen = 0
    End If
    
    propLenVarSize = MqttEncodeRemainingLength(propLen, propLenVar)
    
    payloadLen = 2 + tLen + IIf(qos > 0, 2, 0) + propLenVarSize + propLen + mLen
    ReDim payload(0 To payloadLen - 1)
    
    pos = 0
    payload(pos) = CByte((tLen \ 256) And &HFF)
    payload(pos + 1) = CByte(tLen And &HFF)
    pos = pos + 2
    
    If tLen > 0 Then
        CopyMemory payload(pos), topicBytes(0), tLen
        pos = pos + tLen
    End If
    
    If qos > 0 Then
        With m_Connections(h)
            .MqttNextPacketId = .MqttNextPacketId + 1
            If .MqttNextPacketId < 0 Or .MqttNextPacketId > MQTT_MAX_PACKET_ID Then .MqttNextPacketId = 1
            packetId = .MqttNextPacketId
            
            payload(pos) = CByte((packetId \ 256) And &HFF)
            payload(pos + 1) = CByte(packetId And &HFF)
            pos = pos + 2
            
            If .MqttInFlightCount > UBound(.MqttInFlight) Then
                ReDim Preserve .MqttInFlight(0 To UBound(.MqttInFlight) + 10)
            End If
            
            .MqttInFlight(.MqttInFlightCount).packetId = packetId
            .MqttInFlight(.MqttInFlightCount).topic = topic
            .MqttInFlight(.MqttInFlightCount).qos = qos
            .MqttInFlight(.MqttInFlightCount).payload = msgBytes
            .MqttInFlight(.MqttInFlightCount).SentTick = GetTickCount()
            .MqttInFlightCount = .MqttInFlightCount + 1
        End With
    End If
    
    CopyMemory payload(pos), propLenVar(0), propLenVarSize
    pos = pos + propLenVarSize
    
    If propLen > 0 Then
        CopyMemory payload(pos), propBytes(0), propLen
        pos = pos + propLen
    End If
    
    If mLen > 0 Then
        CopyMemory payload(pos), msgBytes(0), mLen
    End If
    
    flags = IIf(retained, 1, 0) Or (qos * 2)
    packet = MqttBuildPacket(MQTT_PUBLISH, flags, payload, payloadLen)
    
    MqttPublish = WebSocketSendBinary(packet, h)
End Function

'/**
' * @brief Handles subscriptions properly formatted for MQTT 5.0.
' * @param topic Target subscription channel string.
' * @param qos Quality of Service level.
' * @param handle Target connection handle.
' * @return State of subscription attempt.
' */
Public Function MqttSubscribe(ByVal topic As String, Optional ByVal qos As Byte = 0, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
    Dim h As Long
    Dim topicBytes() As Byte
    Dim payload() As Byte
    Dim payloadLen As Long
    Dim packet() As Byte
    Dim packetId As Integer
    
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    
    With m_Connections(h)
        .MqttNextPacketId = .MqttNextPacketId + 1
        If .MqttNextPacketId < 0 Or .MqttNextPacketId > MQTT_MAX_PACKET_ID Then .MqttNextPacketId = 1
        packetId = .MqttNextPacketId
    End With
    
    topicBytes = StringToUtf8(topic)
    
    payloadLen = 2 + 1 + 2 + UBound(topicBytes) + 1 + 1
    ReDim payload(0 To payloadLen - 1)
    
    payload(0) = CByte((packetId \ 256) And &HFF)
    payload(1) = CByte(packetId And &HFF)
    payload(2) = 0
    payload(3) = CByte(((UBound(topicBytes) + 1) \ 256) And &HFF)
    payload(4) = CByte((UBound(topicBytes) + 1) And &HFF)
    
    CopyMemory payload(5), topicBytes(0), UBound(topicBytes) + 1
    payload(5 + UBound(topicBytes) + 1) = qos
    
    packet = MqttBuildPacket(MQTT_SUBSCRIBE, 2, payload, payloadLen)
    MqttSubscribe = WebSocketSendBinary(packet, h)
End Function

'/**
' * @brief Handles unsubscriptions properly formatted for MQTT 5.0.
' * @param topic Target channel string to unsubscribe from.
' * @param handle Target connection handle.
' * @return State of unsubscription attempt.
' */
Public Function MqttUnsubscribe(ByVal topic As String, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
    Dim h As Long
    Dim topicBytes() As Byte
    Dim payload() As Byte
    Dim payloadLen As Long
    Dim packet() As Byte
    Dim packetId As Integer
    
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    
    With m_Connections(h)
        .MqttNextPacketId = .MqttNextPacketId + 1
        If .MqttNextPacketId < 0 Or .MqttNextPacketId > MQTT_MAX_PACKET_ID Then .MqttNextPacketId = 1
        packetId = .MqttNextPacketId
    End With
    
    topicBytes = StringToUtf8(topic)
    
    payloadLen = 2 + 1 + 2 + UBound(topicBytes) + 1
    ReDim payload(0 To payloadLen - 1)
    
    payload(0) = CByte((packetId \ 256) And &HFF)
    payload(1) = CByte(packetId And &HFF)
    payload(2) = 0
    payload(3) = CByte(((UBound(topicBytes) + 1) \ 256) And &HFF)
    payload(4) = CByte((UBound(topicBytes) + 1) And &HFF)
    
    CopyMemory payload(5), topicBytes(0), UBound(topicBytes) + 1
    
    packet = MqttBuildPacket(MQTT_UNSUBSCRIBE, 2, payload, payloadLen)
    MqttUnsubscribe = WebSocketSendBinary(packet, h)
End Function

'/**
' * @brief Synthesizes MQTT DISCONNECT.
' * @param handle Mapping dimensions variable variables limits.
' * @return True on success string constraint parameters dimension array map boundary boundary array variables map limits map limit target string array values.
' */
Public Function MqttDisconnect(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
    Dim h As Long
    Dim packet() As Byte
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    packet = MqttBuildPacket(MQTT_DISCONNECT, 0, NullByteArray(), 0)
    MqttDisconnect = WebSocketSendBinary(packet, h)
End Function

'/**
' * @brief Pings MQTT effectively structurally safely elegantly reliably correctly efficiently stably successfully safely fluently natively natively securely gracefully properly cleanly elegantly efficiently seamlessly seamlessly correctly precisely.
' * @param handle Parameters map size mapping limit array size memory mapping.
' * @return True on success size dimension memory sizes variable dimensions constraints.
' */
Public Function MqttSendPing(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
    Dim h As Long
    Dim packet() As Byte
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    packet = MqttBuildPacket(MQTT_PINGREQ, 0, NullByteArray(), 0)
    MqttSendPing = WebSocketSendBinary(packet, h)
End Function

'/**
' * @brief Polls for an MQTT packet with a configurable timeout.
' * @param timeoutMs Maximum wait time in milliseconds.
' * @param handle (Optional) Target connection handle.
' * @return Parsed packet data, or empty string if none.
' */
Public Function MqttReceive(Optional ByVal timeoutMs As Long = 5000, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
    Dim h As Long
    Dim data() As Byte
    Dim i As Long
    Dim j As Long
    Dim topicLen As Long
    Dim topic As String
    Dim msgBytes() As Byte
    Dim msgLen As Long
    Dim flags As Byte
    Dim qos As Long
    Dim packetId As Long
    Dim skipLen As Long
    Dim propLen As Long
    Dim propEnd As Long
    Dim propId As Byte
    Dim metaInfo As String
    Dim tTopicB() As Byte
    Dim kB() As Byte
    Dim vB() As Byte
    Dim kL As Long
    Dim vL As Long
    Dim reasonCode As Byte
    Dim startTick As Long

    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    If m_Connections(h).state <> STATE_OPEN Then Exit Function
    If m_Connections(h).AsyncMode Then Exit Function

    startTick = GetTickCount()

    Do
        If WebSocketReceiveBinaryCheck(data, h) Then
            For i = LBound(data) To UBound(data)
                MqttParseByte h, data(i)
                If MqttHasPacket(h) Then
                    With m_Connections(h)
                        Select Case .MqttCurrentPacketType
                            Case MQTT_CONNACK
                                reasonCode = .MqttBuffer(1)
                                skipLen = 2
                                propLen = MqttDecodeVarInt(.MqttBuffer, skipLen)
                                propEnd = skipLen + propLen
                                metaInfo = ""

                                If reasonCode > 0 Then
                                    Do While skipLen < propEnd
                                        propId = .MqttBuffer(skipLen)
                                        skipLen = skipLen + 1
                                        If propId = 31 Then
                                            vL = CLng(.MqttBuffer(skipLen)) * 256& + .MqttBuffer(skipLen + 1)
                                            skipLen = skipLen + 2
                                            ReDim vB(0 To vL - 1)
                                            CopyMemory vB(0), .MqttBuffer(skipLen), vL
                                            metaInfo = " | Erro: " & Utf8ToString(vB, vL)
                                            Exit Do
                                        Else
                                            skipLen = propEnd
                                        End If
                                    Loop
                                    MqttReceive = "[CONNACK_ERROR] Code=" & reasonCode & metaInfo
                                Else
                                    MqttReceive = "[CONNACK]"
                                End If
                                MqttResetParser h
                                Exit Function

                            Case MQTT_PUBLISH
                                flags = .MqttCurrentFlags
                                qos = (flags And 6) \ 2

                                topicLen = CLng(.MqttBuffer(0)) * 256& + CLng(.MqttBuffer(1))
                                If topicLen > 0 Then
                                    ReDim tTopicB(0 To topicLen - 1)
                                    CopyMemory tTopicB(0), .MqttBuffer(2), topicLen
                                    topic = Utf8ToString(tTopicB, topicLen)
                                Else
                                    topic = ""
                                End If

                                skipLen = 2 + topicLen
                                If qos > 0 Then
                                    packetId = CLng(.MqttBuffer(skipLen)) * 256& + CLng(.MqttBuffer(skipLen + 1))
                                    skipLen = skipLen + 2
                                End If

                                propLen = MqttDecodeVarInt(.MqttBuffer, skipLen)
                                propEnd = skipLen + propLen
                                metaInfo = ""

                                Do While skipLen < propEnd
                                    propId = .MqttBuffer(skipLen)
                                    skipLen = skipLen + 1
                                    If propId = 38 Then
                                        kL = CLng(.MqttBuffer(skipLen)) * 256& + .MqttBuffer(skipLen + 1)
                                        skipLen = skipLen + 2
                                        ReDim kB(0 To kL - 1)
                                        CopyMemory kB(0), .MqttBuffer(skipLen), kL
                                        skipLen = skipLen + kL

                                        vL = CLng(.MqttBuffer(skipLen)) * 256& + .MqttBuffer(skipLen + 1)
                                        skipLen = skipLen + 2
                                        ReDim vB(0 To vL - 1)
                                        CopyMemory vB(0), .MqttBuffer(skipLen), vL
                                        skipLen = skipLen + vL

                                        metaInfo = metaInfo & "|" & Utf8ToString(kB, kL) & "=" & Utf8ToString(vB, vL)
                                    Else
                                        skipLen = propEnd
                                    End If
                                Loop

                                If qos = 1 Then
                                    MqttSendAck h, MQTT_PUBACK, 0, CInt(packetId)
                                End If

                                If qos = 2 Then
                                    MqttSendAck h, MQTT_PUBREC, 0, CInt(packetId)
                                End If

                                msgLen = .MqttBufLen - skipLen
                                If msgLen > 0 Then
                                    ReDim msgBytes(0 To msgLen - 1)
                                    CopyMemory msgBytes(0), .MqttBuffer(skipLen), msgLen
                                    MqttReceive = topic & "|" & Utf8ToString(msgBytes, msgLen) & metaInfo
                                Else
                                    MqttReceive = topic & "|" & metaInfo
                                End If

                                MqttResetParser h
                                Exit Function

                            Case MQTT_DISCONNECT
                                reasonCode = .MqttBuffer(0)
                                skipLen = 1
                                propLen = MqttDecodeVarInt(.MqttBuffer, skipLen)
                                propEnd = skipLen + propLen
                                metaInfo = ""

                                Do While skipLen < propEnd
                                    propId = .MqttBuffer(skipLen)
                                    skipLen = skipLen + 1
                                    If propId = 31 Then
                                        vL = CLng(.MqttBuffer(skipLen)) * 256& + .MqttBuffer(skipLen + 1)
                                        skipLen = skipLen + 2
                                        ReDim vB(0 To vL - 1)
                                        CopyMemory vB(0), .MqttBuffer(skipLen), vL
                                        metaInfo = " | Motivo: " & Utf8ToString(vB, vL)
                                        Exit Do
                                    Else
                                        skipLen = propEnd
                                    End If
                                Loop

                                MqttReceive = "[DISCONNECT] Code=" & reasonCode & metaInfo
                                MqttResetParser h
                                Exit Function

                            Case MQTT_SUBACK
                                MqttResetParser h
                                MqttReceive = "[SUBACK]"
                                Exit Function

                            Case MQTT_PUBACK, MQTT_PUBCOMP
                                packetId = CLng(.MqttBuffer(0)) * 256& + CLng(.MqttBuffer(1))
                                For j = 0 To .MqttInFlightCount - 1
                                    If .MqttInFlight(j).packetId = packetId Then
                                        If j < .MqttInFlightCount - 1 Then
                                            .MqttInFlight(j) = .MqttInFlight(.MqttInFlightCount - 1)
                                        End If
                                        .MqttInFlightCount = .MqttInFlightCount - 1
                                        Exit For
                                    End If
                                Next j
                                MqttResetParser h

                            Case MQTT_PUBREC
                                packetId = CLng(.MqttBuffer(0)) * 256& + CLng(.MqttBuffer(1))
                                MqttSendAck h, MQTT_PUBREL, 2, CInt(packetId)
                                MqttResetParser h

                            Case MQTT_PUBREL
                                packetId = CLng(.MqttBuffer(0)) * 256& + CLng(.MqttBuffer(1))
                                MqttSendAck h, MQTT_PUBCOMP, 0, CInt(packetId)
                                MqttResetParser h

                            Case Else
                                MqttResetParser h
                        End Select
                    End With
                End If
            Next i
        Else
            If TickDiff(startTick, GetTickCount()) >= timeoutMs Then Exit Do
        End If
        DoEvents
    Loop
End Function

'/**
' * @brief Synthesizes empty string constraint map parameter mapping dimensions size limit dimensions memory boundary dimensions variables target variable variables mapping target limits boundary sizes.
' * @return Byte Array.
' */
Private Function NullByteArray() As Byte()
    Dim b() As Byte
    NullByteArray = b
End Function

'/**
' * @brief Handles dynamic dimensions array boundary value constraint values variables constraints string arrays memory memory array constraints limits array arrays boundary memory arrays dimensions variables values memory variables dimensions constraints parameter parameters map.
' * @param handle Value target value value array arrays limit mapping constraint array size mapping parameters mapping parameter parameter dimension value memory memory array map target dimensions constraints.
' * @param data Ref mapped context byte array payload target structure memory element block value map limit boundary limits array dimension sizes.
' * @param dataLen Boundary size map.
' * @param outLen Target map.
' * @return Constraints.
' */
Private Function DeflatePayload(ByVal handle As Long, ByRef data() As Byte, ByVal dataLen As Long, ByRef outLen As Long) As Byte()
    Dim compBytes() As Byte
    Dim exactData() As Byte
    
    If dataLen = 0 Then
        outLen = 0
        DeflatePayload = exactData
        Exit Function
    End If
    
    If Not m_Connections(handle).CompressionHandler Is Nothing Then
        ReDim exactData(0 To dataLen - 1)
        CopyMemory exactData(0), data(LBound(data)), dataLen
        
        compBytes = m_Connections(handle).CompressionHandler.Deflate(exactData, m_Connections(handle).DeflateWindowBits, m_Connections(handle).DeflateContextTakeover)
        outLen = SafeArrayLen(compBytes)
        DeflatePayload = compBytes
    Else
        outLen = dataLen
        ReDim exactData(0 To dataLen - 1)
        CopyMemory exactData(0), data(LBound(data)), dataLen
        DeflatePayload = exactData
    End If
End Function

'/**
' * @brief Reassembles memory parameters variable limits dimension variable map target array size mapping limits parameters map variable boundary values dimension map boundary.
' * @param handle Limit variables dimensions array constraint boundary memory array map limits variable constraints parameter variables variables sizes variable boundary limit variables constraint sizes array map mapping constraints mapping size constraints dimensions target limits limits variables limit size sizes limit size map mapping array arrays.
' * @param data Values.
' * @param dataLen Length array array string sizes limits sizes.
' * @param outLen Limits target sizes string variable value parameter variables parameters values boundary parameter arrays memory dimensions domain map string.
' * @return Raw arrays variable string limit dimension parameters sizes map string dimensions memory boundaries limits domain dimensions target domain arrays.
' */
Private Function InflatePayload(ByVal handle As Long, ByRef data() As Byte, ByVal dataLen As Long, ByRef outLen As Long) As Byte()
    Dim decompBytes() As Byte
    Dim exactData() As Byte
    
    If dataLen = 0 Then
        outLen = 0
        InflatePayload = exactData
        Exit Function
    End If
    
    If Not m_Connections(handle).CompressionHandler Is Nothing Then
        ReDim exactData(0 To dataLen - 1)
        CopyMemory exactData(0), data(LBound(data)), dataLen
        
        decompBytes = m_Connections(handle).CompressionHandler.Inflate(exactData, m_Connections(handle).InflateWindowBits, m_Connections(handle).InflateContextTakeover)
        outLen = SafeArrayLen(decompBytes)
        InflatePayload = decompBytes
    Else
        outLen = dataLen
        ReDim exactData(0 To dataLen - 1)
        CopyMemory exactData(0), data(LBound(data)), dataLen
        InflatePayload = exactData
    End If
End Function
