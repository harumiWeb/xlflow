VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ROneCOne"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private Declare PtrSafe Function DispCallFunc Lib "oleaut32.dll" ( _
    ByVal instancePointer As LongPtr, _
    ByVal functionPointer As LongPtr, _
    ByVal callingConvention As Long, _
    ByVal returnType As Integer, _
    ByVal argumentCount As Long, _
    ByVal argumentTypesPointer As LongPtr, _
    ByVal argumentPointersPointer As LongPtr, _
    ByVal resultPointer As LongPtr _
) As Long

Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    Destination As Any, _
    Source As Any, _
    ByVal length As LongPtr _
)

Private Declare PtrSafe Function GetTickCount64 Lib "kernel32" () As LongLong

' Windows CNG primitives back the Hash surface: every algorithm ships inside
' bcrypt.dll on every supported Windows, needs no registration, and computes
' at native speed. Passing null object buffers lets CNG manage its own state.
Private Declare PtrSafe Function BCryptOpenAlgorithmProvider Lib "bcrypt.dll" ( _
    ByRef algorithmHandle As LongPtr, ByVal algorithmId As LongPtr, _
    ByVal implementation As LongPtr, ByVal flags As Long) As Long
Private Declare PtrSafe Function BCryptCloseAlgorithmProvider Lib "bcrypt.dll" ( _
    ByVal algorithmHandle As LongPtr, ByVal flags As Long) As Long
Private Declare PtrSafe Function BCryptCreateHash Lib "bcrypt.dll" ( _
    ByVal algorithmHandle As LongPtr, ByRef hashHandle As LongPtr, _
    ByVal hashObjectBuffer As LongPtr, ByVal hashObjectLength As Long, _
    ByVal secretPointer As LongPtr, ByVal secretLength As Long, _
    ByVal flags As Long) As Long
Private Declare PtrSafe Function BCryptHashData Lib "bcrypt.dll" ( _
    ByVal hashHandle As LongPtr, ByVal dataPointer As LongPtr, _
    ByVal dataLength As Long, ByVal flags As Long) As Long
Private Declare PtrSafe Function BCryptFinishHash Lib "bcrypt.dll" ( _
    ByVal hashHandle As LongPtr, ByVal outputPointer As LongPtr, _
    ByVal outputLength As Long, ByVal flags As Long) As Long
Private Declare PtrSafe Function BCryptDestroyHash Lib "bcrypt.dll" ( _
    ByVal hashHandle As LongPtr) As Long
Private Declare PtrSafe Function BCryptGetProperty Lib "bcrypt.dll" ( _
    ByVal objectHandle As LongPtr, ByVal propertyName As LongPtr, _
    ByRef output As Long, ByVal outputLength As Long, _
    ByRef resultLength As Long, ByVal flags As Long) As Long
Private Declare PtrSafe Function BCryptGenRandom Lib "bcrypt.dll" ( _
    ByVal algorithmHandle As LongPtr, ByRef buffer As Byte, _
    ByVal bufferLength As Long, ByVal flags As Long) As Long

' The COM runtime's generator backs Guid.NewGuid with version 4 values.
Private Declare PtrSafe Function CoCreateGuid Lib "ole32" ( _
    ByRef guidBytes As Byte) As Long

Private Declare PtrSafe Sub Sleep Lib "kernel32" ( _
    ByVal milliseconds As Long _
)

' Windows owns the time zone database, so the DateTime surface delegates all
' local and universal conversions to kernel32. A null zone pointer selects the
' machine's current zone with daylight saving rules applied per instant.
Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

' Decoder state for the pure-VBA RFC 1951 inflate engine behind the zip
' surface: a cursor and bit buffer over the loaded archive bytes, plus the
' growing output. The archive itself stays in one shared working buffer.
Private Type InflateState
    position As Long
    limit As Long
    bitBuffer As Long
    bitCount As Long
    output() As Byte
    outPosition As Long
End Type

Private Declare PtrSafe Sub GetSystemTime Lib "kernel32" ( _
    ByRef systemTime As SYSTEMTIME)
Private Declare PtrSafe Function SystemTimeToTzSpecificLocalTime Lib "kernel32" ( _
    ByVal timeZonePointer As LongPtr, ByRef universalTime As SYSTEMTIME, _
    ByRef localTime As SYSTEMTIME) As Long
Private Declare PtrSafe Function TzSpecificLocalTimeToSystemTime Lib "kernel32" ( _
    ByVal timeZonePointer As LongPtr, ByRef localTime As SYSTEMTIME, _
    ByRef universalTime As SYSTEMTIME) As Long

' Interactive process sessions. Standard output and standard error ride
' anonymous pipes that PeekNamedPipe inspects without ever blocking.
' Standard input rides an overlapped named pipe, so even a payload larger
' than the pipe buffer pends instead of freezing Excel while the child
' catches up.
Private Declare PtrSafe Function CreatePipe Lib "kernel32" ( _
    ByRef readPipe As LongPtr, ByRef writePipe As LongPtr, _
    ByRef pipeAttributes As SECURITY_ATTRIBUTES, ByVal size As Long) As Long
Private Declare PtrSafe Function CreateNamedPipeW Lib "kernel32" ( _
    ByVal pipeName As LongPtr, ByVal openMode As Long, _
    ByVal pipeMode As Long, ByVal maxInstances As Long, _
    ByVal outBufferSize As Long, ByVal inBufferSize As Long, _
    ByVal defaultTimeOut As Long, _
    ByRef securityAttributes As SECURITY_ATTRIBUTES) As LongPtr
Private Declare PtrSafe Function CreateFileW Lib "kernel32" ( _
    ByVal fileName As LongPtr, ByVal desiredAccess As Long, _
    ByVal shareMode As Long, _
    ByRef securityAttributes As SECURITY_ATTRIBUTES, _
    ByVal creationDisposition As Long, ByVal flagsAndAttributes As Long, _
    ByVal templateFile As LongPtr) As LongPtr
Private Declare PtrSafe Function ConnectNamedPipe Lib "kernel32" ( _
    ByVal namedPipe As LongPtr, ByVal overlapped As LongPtr) As Long
Private Declare PtrSafe Function CreateProcessW Lib "kernel32" ( _
    ByVal applicationName As LongPtr, ByVal commandLine As LongPtr, _
    ByVal processAttributes As LongPtr, ByVal threadAttributes As LongPtr, _
    ByVal inheritHandles As Long, ByVal creationFlags As Long, _
    ByVal environment As LongPtr, ByVal currentDirectory As LongPtr, _
    ByRef startupInformation As STARTUPINFOW, _
    ByRef processInformation As PROCESS_INFORMATION) As Long
Private Declare PtrSafe Function PeekNamedPipe Lib "kernel32" ( _
    ByVal namedPipe As LongPtr, ByVal buffer As LongPtr, _
    ByVal bufferSize As Long, ByVal bytesRead As LongPtr, _
    ByRef totalBytesAvailable As Long, _
    ByVal bytesLeftThisMessage As LongPtr) As Long
Private Declare PtrSafe Function ReadFile Lib "kernel32" ( _
    ByVal fileHandle As LongPtr, ByVal buffer As LongPtr, _
    ByVal bytesToRead As Long, ByRef bytesRead As Long, _
    ByVal overlapped As LongPtr) As Long
Private Declare PtrSafe Function WriteFileOverlapped Lib "kernel32" _
    Alias "WriteFile" ( _
    ByVal fileHandle As LongPtr, ByVal buffer As LongPtr, _
    ByVal bytesToWrite As Long, ByVal bytesWritten As LongPtr, _
    ByRef overlapped As OVERLAPPED) As Long
Private Declare PtrSafe Function GetOverlappedResult Lib "kernel32" ( _
    ByVal fileHandle As LongPtr, ByRef overlapped As OVERLAPPED, _
    ByRef bytesTransferred As Long, ByVal waitFlag As Long) As Long
Private Declare PtrSafe Function CreateEventW Lib "kernel32" ( _
    ByVal eventAttributes As LongPtr, ByVal manualReset As Long, _
    ByVal initialState As Long, ByVal eventName As LongPtr) As LongPtr
Private Declare PtrSafe Function SetHandleInformation Lib "kernel32" ( _
    ByVal handleValue As LongPtr, ByVal mask As Long, _
    ByVal flags As Long) As Long
Private Declare PtrSafe Function GetExitCodeProcess Lib "kernel32" ( _
    ByVal processHandle As LongPtr, ByRef exitCode As Long) As Long
Private Declare PtrSafe Function TerminateProcess Lib "kernel32" ( _
    ByVal processHandle As LongPtr, ByVal exitCode As Long) As Long
Private Declare PtrSafe Function CloseHandle Lib "kernel32" ( _
    ByVal handleValue As LongPtr) As Long
Private Declare PtrSafe Function GetOEMCP Lib "kernel32" () As Long

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As LongPtr
    bInheritHandle As Long
End Type

Private Type STARTUPINFOW
    cb As Long
    lpReserved As LongPtr
    lpDesktop As LongPtr
    lpTitle As LongPtr
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As LongPtr
    hStdInput As LongPtr
    hStdOutput As LongPtr
    hStdError As LongPtr
End Type

Private Type PROCESS_INFORMATION
    hProcess As LongPtr
    hThread As LongPtr
    dwProcessId As Long
    dwThreadId As Long
End Type

Private Type OVERLAPPED
    Internal As LongPtr
    InternalHigh As LongPtr
    OffsetLow As Long
    OffsetHigh As Long
    hEvent As LongPtr
End Type

' MIT License
'
' Copyright (c) 2026 William Smith
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.

Private Const ROLE_EMPTY As Long = 0
Private Const ROLE_VALUE As Long = 1
Private Const ROLE_PARAMETER As Long = 2
Private Const ROLE_EXPRESSION As Long = 3
Private Const ROLE_LAMBDA As Long = 4
Private Const ROLE_METHOD As Long = 5
Private Const ROLE_COMPOSITION As Long = 6
Private Const ROLE_LIST As Long = 7
Private Const ROLE_QUERY As Long = 8
Private Const ROLE_PROCEDURE As Long = 9
Private Const ROLE_MULTICAST As Long = 10
Private Const ROLE_NATIVE As Long = 11
Private Const ROLE_REFERENCE As Long = 12
Private Const ROLE_EVENT As Long = 13
Private Const ROLE_TRY As Long = 14
Private Const ROLE_ERROR As Long = 15
Private Const ROLE_WHERE_BUILDER As Long = 16
Private Const ROLE_DICTIONARY As Long = 17
Private Const ROLE_HASH_SET As Long = 18
Private Const ROLE_COLLECTION As Long = 19
Private Const ROLE_NODE As Long = 20
Private Const ROLE_IMMUTABLE As Long = 21
Private Const ROLE_CONCURRENT As Long = 22
Private Const ROLE_ENTRY As Long = 23
Private Const ROLE_TASK_FACTORY As Long = 24
Private Const ROLE_TASK As Long = 25
Private Const ROLE_CANCELLATION_SOURCE As Long = 26
Private Const ROLE_CANCELLATION_TOKEN As Long = 27
Private Const ROLE_DATA_TABLE As Long = 28
Private Const ROLE_DATA_COLUMN As Long = 29
Private Const ROLE_DATA_ROW As Long = 30
Private Const ROLE_DATA_SET As Long = 31
Private Const ROLE_DATA_RELATION As Long = 32
Private Const ROLE_DATA_VIEW As Long = 33
Private Const ROLE_DB_CONNECTION As Long = 34
Private Const ROLE_DB_COMMAND As Long = 35
Private Const ROLE_DB_PARAMETER As Long = 36
Private Const ROLE_DB_READER As Long = 37
Private Const ROLE_DB_ADAPTER As Long = 38
Private Const ROLE_DB_TRANSACTION As Long = 39
Private Const ROLE_TASK_COMPLETION_SOURCE As Long = 40
Private Const ROLE_PROGRESS As Long = 41
Private Const ROLE_CANCELLATION_REGISTRATION As Long = 42
Private Const ROLE_DB_NULL As Long = 43
Private Const ROLE_USING As Long = 44
Private Const ROLE_HTTP_CLIENT As Long = 45
Private Const ROLE_HTTP_RESPONSE As Long = 46
Private Const ROLE_JSON_FACTORY As Long = 47
Private Const ROLE_FILE_FACTORY As Long = 48
Private Const ROLE_DIRECTORY_FACTORY As Long = 49
Private Const ROLE_PATH_FACTORY As Long = 50
Private Const ROLE_CSV_FACTORY As Long = 51
Private Const ROLE_PROCESS_FACTORY As Long = 52
Private Const ROLE_PROCESS_RESULT As Long = 53
Private Const ROLE_REGEX As Long = 54
Private Const ROLE_REGEX_MATCH As Long = 55
Private Const ROLE_HASH_FACTORY As Long = 56
Private Const ROLE_CONVERT_FACTORY As Long = 57
Private Const ROLE_DATETIME_FACTORY As Long = 58
Private Const ROLE_DATETIME As Long = 59
Private Const ROLE_TIMESPAN_FACTORY As Long = 60
Private Const ROLE_TIMESPAN As Long = 61
Private Const ROLE_STRINGS_FACTORY As Long = 62
Private Const ROLE_STRING_BUILDER As Long = 63
Private Const ROLE_GUID_FACTORY As Long = 64
Private Const ROLE_RANDOM_FACTORY As Long = 65
Private Const ROLE_XML_FACTORY As Long = 66
Private Const ROLE_XML_NODE As Long = 67
Private Const ROLE_URI_FACTORY As Long = 68
Private Const ROLE_WEB_UTILITY_FACTORY As Long = 69
Private Const ROLE_LOGGER As Long = 70
Private Const ROLE_FILE_WATCHER As Long = 71
Private Const ROLE_FILE_CHANGE As Long = 72
Private Const ROLE_ZIPFILE_FACTORY As Long = 73
Private Const ROLE_ZIP_ARCHIVE As Long = 74
Private Const ROLE_ZIP_ENTRY As Long = 75
Private Const ROLE_QUERYABLE As Long = 76
Private Const ROLE_PROCESS_SESSION As Long = 77

Private Const STARTF_USESTDHANDLES As Long = &H100&
Private Const CREATE_NO_WINDOW As Long = &H8000000
Private Const HANDLE_FLAG_INHERIT As Long = 1
Private Const STILL_ACTIVE As Long = 259
Private Const PIPE_ACCESS_OUTBOUND As Long = 2
Private Const FILE_FLAG_OVERLAPPED As Long = &H40000000
Private Const GENERIC_READ As Long = &H80000000
Private Const OPEN_EXISTING As Long = 3
Private Const ERROR_IO_INCOMPLETE As Long = 996
Private Const SESSION_PIPE_BYTES As Long = 65536

Private Const QUERY_DIALECT_TSQL As Long = 1
Private Const QUERY_DIALECT_ACE As Long = 2

Private Const TASK_WORK As Long = 1
Private Const TASK_COMPLETED As Long = 2
Private Const TASK_DELAY As Long = 3
Private Const TASK_WHEN_ALL As Long = 4
Private Const TASK_WHEN_ANY As Long = 5
Private Const TASK_CONTINUATION As Long = 6
Private Const TASK_PROMISE As Long = 7
Private Const TASK_PROVIDER_CALL As Long = 8
Private Const TASK_YIELD As Long = 9
Private Const TASK_WAIT_ASYNC As Long = 10
Private Const TASK_HTTP_SEND As Long = 11
Private Const TASK_DB_ASYNC As Long = 12
Private Const TASK_PROCESS_RUN As Long = 13
Private Const TASK_FILE_WATCH As Long = 14
Private Const TASK_SESSION As Long = 15

' A watch task rescans its folder at most this often while awaited.
Private Const WATCH_POLL_INTERVAL_MS As Long = 250

Private Const TASK_STATUS_CREATED As Long = 0
Private Const TASK_STATUS_RUNNING As Long = 1
Private Const TASK_STATUS_COMPLETED As Long = 2
Private Const TASK_STATUS_CANCELED As Long = 3
Private Const TASK_STATUS_FAULTED As Long = 4

Private Const ROW_STATE_DETACHED As Long = 0
Private Const ROW_STATE_UNCHANGED As Long = 1
Private Const ROW_STATE_ADDED As Long = 2
Private Const ROW_STATE_DELETED As Long = 3
Private Const ROW_STATE_MODIFIED As Long = 4

Private Const COLLECTION_DICTIONARY As Long = 1
Private Const COLLECTION_HASH_SET As Long = 2
Private Const COLLECTION_QUEUE As Long = 3
Private Const COLLECTION_STACK As Long = 4
Private Const COLLECTION_LINKED_LIST As Long = 5
Private Const COLLECTION_SORTED_LIST As Long = 6
Private Const COLLECTION_SORTED_DICTIONARY As Long = 7
Private Const COLLECTION_SORTED_SET As Long = 8
Private Const COLLECTION_ORDERED_DICTIONARY As Long = 9
Private Const COLLECTION_PRIORITY_QUEUE As Long = 10
Private Const COLLECTION_OBSERVABLE As Long = 11
Private Const COLLECTION_READ_ONLY As Long = 12
Private Const COLLECTION_KEYED As Long = 13
Private Const COLLECTION_CONCURRENT_DICTIONARY As Long = 14
Private Const COLLECTION_CONCURRENT_QUEUE As Long = 15
Private Const COLLECTION_CONCURRENT_STACK As Long = 16
Private Const COLLECTION_CONCURRENT_BAG As Long = 17
Private Const COLLECTION_BLOCKING As Long = 18
Private Const COLLECTION_IMMUTABLE_ARRAY As Long = 19
Private Const COLLECTION_IMMUTABLE_LIST As Long = 20
Private Const COLLECTION_IMMUTABLE_DICTIONARY As Long = 21
Private Const COLLECTION_IMMUTABLE_HASH_SET As Long = 22
Private Const COLLECTION_IMMUTABLE_QUEUE As Long = 23
Private Const COLLECTION_IMMUTABLE_STACK As Long = 24
Private Const COLLECTION_IMMUTABLE_SORTED_DICTIONARY As Long = 25
Private Const COLLECTION_IMMUTABLE_SORTED_SET As Long = 26
Private Const COLLECTION_LOOKUP As Long = 27
Private Const COLLECTION_GROUPING As Long = 28

Private Const CALLCONV_STDCALL As Long = 4
Private Const HRESULT_SUCCESS As Long = 0

Private Const OP_ADD As Long = 1
Private Const OP_SUBTRACT As Long = 2
Private Const OP_MULTIPLY As Long = 3
Private Const OP_DIVIDE As Long = 4
Private Const OP_MODULO As Long = 5
Private Const OP_NEGATE As Long = 6
Private Const OP_CONCAT As Long = 7
Private Const OP_EQUAL As Long = 8
Private Const OP_NOT_EQUAL As Long = 9
Private Const OP_LESS_THAN As Long = 10
Private Const OP_LESS_THAN_OR_EQUAL As Long = 11
Private Const OP_GREATER_THAN As Long = 12
Private Const OP_GREATER_THAN_OR_EQUAL As Long = 13
Private Const OP_AND_ALSO As Long = 14
Private Const OP_OR_ELSE As Long = 15
Private Const OP_NOT As Long = 16
Private Const OP_MEMBER_VALUE As Long = 17
Private Const OP_MEMBER_OBJECT As Long = 18
Private Const OP_STARTS_WITH As Long = 19
Private Const OP_ENDS_WITH As Long = 20
Private Const OP_CONTAINS_TEXT As Long = 21
Private Const OP_IS_NOTHING As Long = 22
Private Const OP_IS_NOT_NOTHING As Long = 23
Private Const OP_IS_NULL_OR_EMPTY As Long = 24
Private Const OP_MATCHES_PATTERN As Long = 25
Private Const OP_SAFE_MEMBER_VALUE As Long = 26
Private Const OP_SAFE_MEMBER_OBJECT As Long = 27
Private Const OP_STRING_EQUAL As Long = 28
Private Const OP_STRING_NOT_EQUAL As Long = 29
Private Const OP_IN_SEQUENCE As Long = 30
Private Const OP_NOT_IN_SEQUENCE As Long = 31
Private Const OP_ANY_MATCH As Long = 32
Private Const OP_ALL_MATCH As Long = 33
Private Const OP_NONE_MATCH As Long = 34

Private Const QUERY_WHERE As Long = 1
Private Const QUERY_SELECT As Long = 2
Private Const QUERY_TAKE As Long = 3
Private Const QUERY_SKIP As Long = 4
Private Const QUERY_DISTINCT As Long = 5
Private Const QUERY_ORDER As Long = 6
Private Const QUERY_ORDER_DESCENDING As Long = 7
Private Const QUERY_ORDER_BY As Long = 8
Private Const QUERY_ORDER_BY_DESCENDING As Long = 9
Private Const QUERY_THEN_BY As Long = 10
Private Const QUERY_THEN_BY_DESCENDING As Long = 11
Private Const QUERY_APPEND As Long = 12
Private Const QUERY_PREPEND As Long = 13
Private Const QUERY_REVERSE As Long = 14
Private Const QUERY_DISTINCT_BY As Long = 15
Private Const QUERY_WHERE_NOT As Long = 16
Private Const QUERY_TAKE_WHILE As Long = 17
Private Const QUERY_SKIP_WHILE As Long = 18
Private Const QUERY_TAKE_LAST As Long = 19
Private Const QUERY_SKIP_LAST As Long = 20
Private Const QUERY_CONCAT As Long = 21
Private Const QUERY_UNION As Long = 22
Private Const QUERY_INTERSECT As Long = 23
Private Const QUERY_EXCEPT As Long = 24
Private Const QUERY_DEFAULT_IF_EMPTY As Long = 25
Private Const QUERY_CHUNK As Long = 26
Private Const QUERY_SELECT_MANY As Long = 27
Private Const QUERY_CAST As Long = 28
Private Const QUERY_OF_TYPE As Long = 29
Private Const QUERY_UNION_BY As Long = 30
Private Const QUERY_INTERSECT_BY As Long = 31
Private Const QUERY_EXCEPT_BY As Long = 32

Private Const ERROR_INVALID_ARGUMENT As Long = vbObjectError + 4100
Private Const ERROR_INVALID_OPERATION As Long = vbObjectError + 4101
Private Const ERROR_ARITY_MISMATCH As Long = vbObjectError + 4102
Private Const ERROR_TYPE_MISMATCH As Long = vbObjectError + 4103
Private Const ERROR_UNBOUND_PARAMETER As Long = vbObjectError + 4104
Private Const ERROR_METHOD_INVOCATION As Long = vbObjectError + 4105
Private Const ERROR_MEMBER_ACCESS As Long = vbObjectError + 4106
Private Const ERROR_NATIVE_INVOCATION As Long = vbObjectError + 4107
Private Const ERROR_BYREF_INVOCATION As Long = vbObjectError + 4108
Private Const ERROR_TIMEOUT As Long = vbObjectError + 4109
Private Const ERROR_CANCELED As Long = vbObjectError + 4110
Private Const ERROR_HTTP_REQUEST As Long = vbObjectError + 4111
Private Const ERROR_JSON_PARSE As Long = vbObjectError + 4112
Private Const ERROR_IO As Long = vbObjectError + 4113
Private Const ERROR_CSV_PARSE As Long = vbObjectError + 4114
Private Const ERROR_PROCESS As Long = vbObjectError + 4115
Private Const ERROR_REGEX As Long = vbObjectError + 4116
Private Const ERROR_FORMAT As Long = vbObjectError + 4117
Private Const ERROR_XML As Long = vbObjectError + 4118
Private Const ERROR_ZIP As Long = vbObjectError + 4119
Private Const ERROR_QUERY As Long = vbObjectError + 4120

Private Const HTTP_PROG_ID As String = "WinHttp.WinHttpRequest.5.1"
Private Const HTTP_DEFAULT_TIMEOUT_MS As Long = 30000

' ADO enumeration values used by the native async provider path: the option
' flag that starts Open and Execute without blocking, and the State bits that
' report an operation still connecting, executing, or fetching.
Private Const ADO_ASYNC_OPTION As Long = 16
Private Const ADO_STATE_OPEN As Long = 1
Private Const ADO_STATE_CONNECTING As Long = 2
Private Const ADO_STATE_EXECUTING As Long = 4
Private Const ADO_STATE_FETCHING As Long = 8

' ADODB.Stream enumeration values for the file system layer: stream types,
' the overwrite save mode, and the prog-id used for text and byte transput.
Private Const STREAM_PROG_ID As String = "ADODB.Stream"
Private Const STREAM_TYPE_BINARY As Long = 1
Private Const STREAM_TYPE_TEXT As Long = 2
Private Const STREAM_SAVE_OVERWRITE As Long = 2

' The Windows Script Host shell launches processes for the Process surface;
' Exec exposes a pollable Status so the Task scheduler never blocks on it.
Private Const WSHELL_PROG_ID As String = "WScript.Shell"

' The in-box script engine's expression object backs the Regex surface.
Private Const REGEX_PROG_ID As String = "VBScript.RegExp"

' MSXML6 ships with every supported Windows and backs the Xml surface with
' secure parser defaults: DTDs prohibited, external references unresolved.
Private Const XML_PROG_ID As String = "MSXML2.DOMDocument.6.0"

' Windows CNG option flag: the algorithm provider operates in HMAC mode.
Private Const BCRYPT_HMAC_FLAG As Long = 8

' Windows CNG option flag: draw from the system-preferred random provider.
Private Const BCRYPT_USE_SYSTEM_RNG As Long = 2

' Logging severities in the Microsoft.Extensions.Logging order.
Private Const LOG_LEVEL_TRACE As Long = 0
Private Const LOG_LEVEL_DEBUG As Long = 1
Private Const LOG_LEVEL_INFORMATION As Long = 2
Private Const LOG_LEVEL_WARNING As Long = 3
Private Const LOG_LEVEL_ERROR As Long = 4
Private Const LOG_LEVEL_CRITICAL As Long = 5

' The DateTime surface stores instants as milliseconds since the Unix epoch
' 1970-01-01T00:00:00Z; the bounds pin the supported clock years 0100-9999.
Private Const MS_PER_DAY As Double = 86400000#
Private Const MS_PER_HOUR As Double = 3600000#
Private Const MS_PER_MINUTE As Double = 60000#
Private Const MS_PER_SECOND As Double = 1000#
Private Const MIN_DATETIME_MS As Double = -59011459200000#
Private Const MAX_DATETIME_MS As Double = 253402300799999#

' Reader state for the JSON parser: a 1-based cursor over the input plus a
' byte snapshot of the same text as UTF-16 code units, so scanning loops read
' character codes without allocating a one-character string per position.
' The parsing strategy is adapted, with its author's permission, from
' ModernJsonInVBA by William Smith Edward.
Private Type JsonReader
    text As String
    bytes() As Byte
    pos As Long
    textLen As Long
End Type

' Growable string buffer: VBA concatenation in a loop is quadratic, so JSON
' text is written with Mid$ assignment into a doubling buffer instead.
Private Type JsonTextBuilder
    buffer As String
    used As Long
    capacity As Long
End Type

Private mRole As Long
Private mOperator As Long
Private mExpressionOption As Long
Private mLeft As ROneCOne
Private mRight As ROneCOne
Private mNext As ROneCOne

Private mScalarValue As Variant
Private mObjectValue As Object
Private mHasObjectValue As Boolean

Private mExpectedType As VbVarType
Private mExpectedObjectType As String
Private mParameters As Collection
Private mIsByRef As Boolean

Private mTarget As Object
Private mMethodName As String
Private mMethodArity As Long
Private mIsActionDelegate As Boolean
Private mHasParameterSignature As Boolean
Private mHasReturnSignature As Boolean
Private mReturnType As VbVarType
Private mReturnObjectType As String
Private mInvocationList As Collection
Private mNativeAddress As LongPtr
Private mReferenceAddress As LongPtr

Private mTryBody As ROneCOne
Private mCatchActions As Collection
Private mCatchNumbers As Collection
Private mFinallyAction As ROneCOne
Private mErrorNumber As Long
Private mErrorSource As String
Private mErrorDescription As String
Private mErrorHelpFile As String
Private mErrorHelpContext As Long

Private mItems() As ROneCOne
Private mItemsCount As Long
Private mEnumerationValues As Collection
Private mElementType As VbVarType
Private mElementObjectType As String
Private mSource As ROneCOne
Private mSourceTable As Object
Private mQueryDelegate As ROneCOne
Private mQueryComparer As ROneCOne
Private mQueryValue As ROneCOne
Private mQueryCount As Long
Private mVersion As Long
Private mDataVersion As Long
Private mEnumerationVersion As Long
Private mRowsSnapshot As ROneCOne
Private mRowsSnapshotVersion As Long
Private mColumnsSnapshot As ROneCOne
Private mColumnsSnapshotVersion As Long
Private mTablesSnapshot As ROneCOne
Private mTablesSnapshotVersion As Long
Private mRelationsSnapshot As ROneCOne
Private mRelationsSnapshotVersion As Long
Private mPrimaryKeySnapshot As ROneCOne
Private mPrimaryKeySnapshotVersion As Long
Private mElementParameter As ROneCOne
Private mContextMemberPath As String
Private mCollectionKind As Long
Private mKeys() As ROneCOne
Private mKeysCount As Long
Private mPriorities() As ROneCOne
Private mPrioritiesCount As Long
Private mHashSlots() As Long
Private mHashCodes() As Long
Private mHashKeys() As ROneCOne
Private mHashValues() As ROneCOne
Private mHashCapacity As Long
Private mHashCount As Long
Private mHashDirty As Boolean
Private mKeyType As VbVarType
Private mKeyObjectType As String
Private mKeySelector As ROneCOne
Private mCollectionChanged As ROneCOne
Private mNodePrevious As ROneCOne
Private mNodeNext As ROneCOne
Private mAddingCompleted As Boolean
Private mTaskBody As ROneCOne
Private mTaskResult As ROneCOne
Private mTaskException As ROneCOne
Private mTaskChildren As Collection
Private mTaskStatus As Long
Private mTaskDeadline As LongLong
Private mTaskAdvancing As Boolean
Private mTaskWaitActive As Boolean
Private mCancellationSource As ROneCOne
Private mCancellationRequested As Boolean
Private mName As String
Private mOriginalItems() As ROneCOne
Private mOriginalItemsCount As Long
Private mPrimaryKey As Collection
Private mDataIndex As ROneCOne
Private mDataIndexDirty As Boolean
Private mUniqueIndexes As Collection
Private mDataTables As Collection
Private mDataRelations As Collection
Private mRelationParentColumns As Collection
Private mRelationChildColumns As Collection
Private mRelationParentIndex As ROneCOne
Private mRelationChildIndex As ROneCOne
Private mAllowDBNull As Boolean
Private mUnique As Boolean
Private mAutoIncrement As Boolean
Private mAutoIncrementSeed As LongLong
Private mAutoIncrementStep As LongLong
Private mAutoIncrementCurrent As LongLong
Private mDefaultValue As ROneCOne
Private mRowState As Long
Private mProviderObject As Object
Private mProviderCommand As Object
Private mProcessExitCode As Long
Private mProcessOutput As String
Private mProcessError As String
Private mProcessOutputPath As String
Private mProcessErrorPath As String
Private mRegexSuccess As Boolean
Private mRegexIndex As Long
Private mRegexLength As Long
Private mDateTimeMs As Double
Private mDateTimeOffsetMinutes As Long
Private mTimeSpanMs As Double
Private mStringBuilder As JsonTextBuilder
Private mLoggerMinimumLevel As Long
Private mWatcherFilter As String
Private mWatcherChangeType As String
Private mWatcherSnapshot As Collection
Private mWatcherLastTick As LongLong
Private mZipWork() As Byte
Private mZipTablesReady As Boolean
Private mPow2() As Long
Private mCrcTable() As LongLong
Private mZipLengthBase() As Long
Private mZipLengthExtra() As Long
Private mZipDistBase() As Long
Private mZipDistExtra() As Long
Private mFixedLitCounts() As Long
Private mFixedLitSymbols() As Long
Private mFixedDistCounts() As Long
Private mFixedDistSymbols() As Long
Private mZipEntries As Collection
Private mZipArchivePath As String
Private mZipEntryMethod As Long
Private mZipEntryCrc As Double
Private mZipEntryCompressedSize As Long
Private mZipEntrySize As Long
Private mZipEntryOffset As Long
Private mZipEntryIsDirectory As Boolean
Private mSqlQueryConnection As ROneCOne
Private mSqlQueryTable As String
Private mSqlQueryDialect As Long
Private mSqlQueryWheres As Collection
Private mSqlQueryOrderings As Collection
Private mSqlQueryColumns As Collection
Private mSqlQueryTake As Long
Private mSqlQuerySkip As Long
Private mSqlQueryLastValues As Collection
Private mSessionProcess As LongPtr
Private mSessionProcessId As Long
Private mSessionStdIn As LongPtr
Private mSessionStdOut As LongPtr
Private mSessionStdErr As LongPtr
Private mSessionCharset As String
Private mSessionExitCode As Long
Private mSessionExited As Boolean
Private mSessionInputClosed As Boolean
Private mSessionOutBytes() As Byte
Private mSessionOutCount As Long
Private mSessionOutCursor As Long
Private mSessionOutDrained As Boolean
Private mSessionErrBytes() As Byte
Private mSessionErrCount As Long
Private mSessionErrCursor As Long
Private mSessionErrDrained As Boolean
Private mSessionWriteQueue As Collection
Private mSessionWriteDone As Long
Private mSessionWriteNext As Long
Private mSessionWriteBuffer() As Byte
Private mSessionWriteActive As Boolean
Private mSessionWriteOverlapped As OVERLAPPED
Private mSessionReadOutBusy As Boolean
Private mSessionReadErrBusy As Boolean
Private mHttpBaseAddress As String
Private mHttpTimeoutMs As Long
Private mHttpHeaderNames As Collection
Private mHttpHeaderValues As Collection
Private mHttpStatusCode As Long
Private mHttpReasonPhrase As String
Private mHttpContent As String
Private mHttpRawHeaders As String
Private mHttpBytes As Variant
Private mJsonDecimalSeparator As String
Private mJsonSeparatorKnown As Boolean
Private mConnectionString As String
Private mCommandText As String
Private mProviderName As String
Private mProviderType As Long
Private mProviderSize As Long
Private mProviderPrecision As Byte
Private mProviderScale As Byte
Private mSourceColumn As String
Private mSourceVersion As String
Private mCommandTimeout As Long
Private mCommandType As Long
Private mConnectionTimeout As Long
Private mTransactionActive As Boolean
Private mAdapterUseTransaction As Boolean
Private mAdapterContinueOnError As Boolean
Private mProviderErrors As Collection

Public Function Parameter( _
    Optional ByVal expectedType As VbVarType = vbVariant _
) As ROneCOne
Attribute Parameter.VB_Description = "Creates a typed lambda parameter."
    Dim node As ROneCOne

    Set node = New ROneCOne
    node.ConfigureParameter expectedType, vbNullString
    Set Parameter = node
End Function

Public Function ParameterLike(ByVal example As Variant) As ROneCOne
Attribute ParameterLike.VB_Description = "Creates a parameter typed like an example value."
    Dim node As ROneCOne
    Dim exampleObject As Object

    Set node = New ROneCOne
    If IsObject(example) Then
        Set exampleObject = example
        If exampleObject Is Nothing Then
            RaiseContractError ERROR_INVALID_ARGUMENT, "ParameterLike", _
                "The example object cannot be Nothing."
        End If
        node.ConfigureParameter vbObject, TypeName(exampleObject)
    Else
        node.ConfigureParameter VarTypeOf(example), vbNullString
    End If
    Set ParameterLike = node
End Function

Public Function Var( _
    Optional ByVal expectedType As VbVarType = vbVariant _
) As ROneCOne
Attribute Var.VB_Description = "Concise C#-style alias for a typed Parameter."
    Set Var = Parameter(expectedType)
End Function

Public Function VarLike(ByVal example As Variant) As ROneCOne
Attribute VarLike.VB_Description = "Creates a concise parameter typed like an example value."
    Set VarLike = ParameterLike(example)
End Function

Public Function Value(Optional ByVal capturedValue As Variant) As Variant
Attribute Value.VB_Description = "Captures an expression value, or reads a node, entry, or match."
    Dim node As ROneCOne
    Dim capturedObject As Object

    If mRole = ROLE_REGEX_MATCH Then
        If Not IsMissing(capturedValue) Then
            RaiseContractError ERROR_INVALID_ARGUMENT, "Value", _
                "This value does not accept an argument."
        End If
        Value = mName
        Exit Function
    End If
    If mRole = ROLE_XML_NODE Then
        If Not IsMissing(capturedValue) Then
            RaiseContractError ERROR_INVALID_ARGUMENT, "Value", _
                "This value does not accept an argument."
        End If
        Value = mProviderObject.Text
        Exit Function
    End If
    If mRole = ROLE_NODE Or mRole = ROLE_ENTRY Then
        If Not IsMissing(capturedValue) Then
            RaiseContractError ERROR_INVALID_ARGUMENT, "Value", _
                "This value does not accept an argument."
        End If
        AssignPublicResult Value, mQueryValue
        Exit Function
    End If
    If mRole <> ROLE_EMPTY Then
        RaiseContractError ERROR_INVALID_OPERATION, "Value", _
            "Value creation must start from the ROneCOne factory."
    End If
    If IsMissing(capturedValue) Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "Value", _
            "A captured value is required."
    End If

    Set node = New ROneCOne
    If IsObject(capturedValue) Then
        Set capturedObject = capturedValue
        node.ConfigureObjectValue capturedObject
    Else
        node.ConfigureScalarValue capturedValue
    End If
    Set Value = node
End Function

Public Function ListOf( _
    ByVal typeToken As Variant, _
    ParamArray initialItems() As Variant _
) As ROneCOne
Attribute ListOf.VB_Description = "Creates a strict List<T> with optional initial items."
    Dim initialItem As Variant
    Dim listValue As ROneCOne
    Dim prototype As Object
    Dim wrappedItems As Collection

    Set listValue = New ROneCOne
    If IsObject(typeToken) Then
        Set prototype = typeToken
        If prototype Is Nothing Then
            RaiseContractError ERROR_INVALID_ARGUMENT, "ListOf", _
                "The user-defined type prototype cannot be Nothing."
        End If
        listValue.ConfigureList vbObject, TypeName(prototype)
    Else
        ValidateTypeToken typeToken, "ListOf"
        listValue.ConfigureList CLng(typeToken), vbNullString
    End If
    Set wrappedItems = New Collection
    For Each initialItem In initialItems
        wrappedItems.Add listValue.InternalWrapElement(initialItem, "ListOf")
    Next initialItem
    CommitInitialItems listValue, wrappedItems
    Set ListOf = listValue
End Function

Public Function ListLike(ByVal example As Variant) As ROneCOne
Attribute ListLike.VB_Description = "Creates a List<T> typed like an example value."
    Dim listValue As ROneCOne
    Dim prototype As Object

    Set listValue = New ROneCOne
    If IsObject(example) Then
        Set prototype = example
        If prototype Is Nothing Then
            RaiseContractError ERROR_INVALID_ARGUMENT, "ListLike", _
                "The user-defined type prototype cannot be Nothing."
        End If
        listValue.ConfigureList vbObject, TypeName(prototype)
    Else
        listValue.ConfigureList VarTypeOf(example), vbNullString
    End If
    Set ListLike = listValue
End Function

Public Function ListFrom( _
    ByVal firstItem As Variant, _
    ParamArray remainingItems() As Variant _
) As ROneCOne
Attribute ListFrom.VB_Description = "Infers T and creates a populated List<T>."
    Dim remainingItem As Variant
    Dim listValue As ROneCOne
    Dim wrappedItems As Collection

    Set listValue = ListLike(firstItem)
    Set wrappedItems = New Collection
    wrappedItems.Add listValue.InternalWrapElement(firstItem, "ListFrom")
    For Each remainingItem In remainingItems
        wrappedItems.Add listValue.InternalWrapElement(remainingItem, "ListFrom")
    Next remainingItem
    CommitInitialItems listValue, wrappedItems
    Set ListFrom = listValue
End Function

Public Function DictionaryOf( _
    ByVal keyType As Variant, _
    ByVal valueType As Variant, _
    Optional ByVal comparer As Variant _
) As ROneCOne
Attribute DictionaryOf.VB_Description = "Creates a Dictionary<TKey, TValue>."
    Set DictionaryOf = CreateMapCollection( _
        ROLE_DICTIONARY, COLLECTION_DICTIONARY, keyType, valueType, comparer)
End Function

Public Function HashSetOf( _
    ByVal typeToken As Variant, _
    Optional ByVal comparer As Variant _
) As ROneCOne
Attribute HashSetOf.VB_Description = "Creates a HashSet<T>."
    Set HashSetOf = CreateSequenceCollection( _
        ROLE_HASH_SET, COLLECTION_HASH_SET, typeToken, comparer)
End Function

Public Function QueueOf(ByVal typeToken As Variant) As ROneCOne
Attribute QueueOf.VB_Description = "Creates a Queue<T>."
    Set QueueOf = CreateSequenceCollection( _
        ROLE_COLLECTION, COLLECTION_QUEUE, typeToken)
End Function

Public Function StackOf(ByVal typeToken As Variant) As ROneCOne
Attribute StackOf.VB_Description = "Creates a Stack<T>."
    Set StackOf = CreateSequenceCollection( _
        ROLE_COLLECTION, COLLECTION_STACK, typeToken)
End Function

Public Function LinkedListOf(ByVal typeToken As Variant) As ROneCOne
Attribute LinkedListOf.VB_Description = "Creates a LinkedList<T>."
    Set LinkedListOf = CreateSequenceCollection( _
        ROLE_COLLECTION, COLLECTION_LINKED_LIST, typeToken)
End Function

Public Function SortedListOf( _
    ByVal keyType As Variant, _
    ByVal valueType As Variant, _
    Optional ByVal comparer As Variant _
) As ROneCOne
Attribute SortedListOf.VB_Description = "Creates a SortedList<TKey, TValue>."
    Set SortedListOf = CreateMapCollection( _
        ROLE_DICTIONARY, COLLECTION_SORTED_LIST, keyType, valueType, comparer)
End Function

Public Function SortedDictionaryOf( _
    ByVal keyType As Variant, _
    ByVal valueType As Variant, _
    Optional ByVal comparer As Variant _
) As ROneCOne
Attribute SortedDictionaryOf.VB_Description = _
    "Creates a SortedDictionary<TKey, TValue>."
    Set SortedDictionaryOf = CreateMapCollection( _
        ROLE_DICTIONARY, COLLECTION_SORTED_DICTIONARY, _
        keyType, valueType, comparer)
End Function

Public Function SortedSetOf( _
    ByVal typeToken As Variant, _
    Optional ByVal comparer As Variant _
) As ROneCOne
Attribute SortedSetOf.VB_Description = "Creates a SortedSet<T>."
    Set SortedSetOf = CreateSequenceCollection( _
        ROLE_HASH_SET, COLLECTION_SORTED_SET, typeToken, comparer)
End Function

Public Function OrderedDictionaryOf( _
    ByVal keyType As Variant, _
    ByVal valueType As Variant, _
    Optional ByVal comparer As Variant _
) As ROneCOne
Attribute OrderedDictionaryOf.VB_Description = _
    "Creates an OrderedDictionary<TKey, TValue>."
    Set OrderedDictionaryOf = CreateMapCollection( _
        ROLE_DICTIONARY, COLLECTION_ORDERED_DICTIONARY, _
        keyType, valueType, comparer)
End Function

Public Function PriorityQueueOf( _
    ByVal elementType As Variant, _
    ByVal priorityType As Variant, _
    Optional ByVal comparer As Variant _
) As ROneCOne
Attribute PriorityQueueOf.VB_Description = _
    "Creates a PriorityQueue<TElement, TPriority>."
    Set PriorityQueueOf = CreateMapCollection( _
        ROLE_COLLECTION, COLLECTION_PRIORITY_QUEUE, _
        priorityType, elementType, comparer)
End Function

Public Function ObservableCollectionOf(ByVal typeToken As Variant) As ROneCOne
Attribute ObservableCollectionOf.VB_Description = _
    "Creates an ObservableCollection<T>."
    Set ObservableCollectionOf = CreateSequenceCollection( _
        ROLE_COLLECTION, COLLECTION_OBSERVABLE, typeToken)
    ObservableCollectionOf.ConfigureObservableCollection
End Function

Public Function ReadOnlyCollectionOf(ByVal source As ROneCOne) As ROneCOne
Attribute ReadOnlyCollectionOf.VB_Description = _
    "Creates a live ReadOnlyCollection<T> view."
    Dim result As ROneCOne

    RequireSequence source, "ReadOnlyCollectionOf"
    Set result = New ROneCOne
    result.ConfigureGenericCollection ROLE_COLLECTION, COLLECTION_READ_ONLY, _
        source.InternalElementType, source.InternalElementObjectType
    result.ConfigureReadOnlySource source
    Set ReadOnlyCollectionOf = result
End Function

Public Function KeyedCollectionOf( _
    ByVal keyType As Variant, _
    ByVal itemType As Variant, _
    ByVal keySelector As ROneCOne, _
    Optional ByVal comparer As Variant _
) As ROneCOne
Attribute KeyedCollectionOf.VB_Description = _
    "Creates a KeyedCollection<TKey, TItem>."
    Dim result As ROneCOne

    Set result = CreateMapCollection( _
        ROLE_COLLECTION, COLLECTION_KEYED, keyType, itemType, comparer)
    result.ConfigureCollectionKeySelector NormalizeStandaloneUnary( _
        keySelector, "KeyedCollectionOf")
    Set KeyedCollectionOf = result
End Function

Public Function ConcurrentDictionaryOf( _
    ByVal keyType As Variant, _
    ByVal valueType As Variant, _
    Optional ByVal comparer As Variant _
) As ROneCOne
Attribute ConcurrentDictionaryOf.VB_Description = _
    "Creates a ConcurrentDictionary<TKey, TValue>."
    Set ConcurrentDictionaryOf = CreateMapCollection( _
        ROLE_CONCURRENT, COLLECTION_CONCURRENT_DICTIONARY, _
        keyType, valueType, comparer)
End Function

Public Function ConcurrentQueueOf(ByVal typeToken As Variant) As ROneCOne
Attribute ConcurrentQueueOf.VB_Description = "Creates a ConcurrentQueue<T>."
    Set ConcurrentQueueOf = CreateSequenceCollection( _
        ROLE_CONCURRENT, COLLECTION_CONCURRENT_QUEUE, typeToken)
End Function

Public Function ConcurrentStackOf(ByVal typeToken As Variant) As ROneCOne
Attribute ConcurrentStackOf.VB_Description = "Creates a ConcurrentStack<T>."
    Set ConcurrentStackOf = CreateSequenceCollection( _
        ROLE_CONCURRENT, COLLECTION_CONCURRENT_STACK, typeToken)
End Function

Public Function ConcurrentBagOf(ByVal typeToken As Variant) As ROneCOne
Attribute ConcurrentBagOf.VB_Description = "Creates a ConcurrentBag<T>."
    Set ConcurrentBagOf = CreateSequenceCollection( _
        ROLE_CONCURRENT, COLLECTION_CONCURRENT_BAG, typeToken)
End Function

Public Function BlockingCollectionOf( _
    ByVal typeToken As Variant, _
    Optional ByVal boundedCapacity As Long = -1 _
) As ROneCOne
Attribute BlockingCollectionOf.VB_Description = _
    "Creates a cooperative BlockingCollection<T>."
    Set BlockingCollectionOf = CreateSequenceCollection( _
        ROLE_CONCURRENT, COLLECTION_BLOCKING, typeToken)
    BlockingCollectionOf.ConfigureCollectionCapacity boundedCapacity
End Function

Public Function ImmutableArrayOf(ByVal typeToken As Variant) As ROneCOne
Attribute ImmutableArrayOf.VB_Description = "Creates an empty ImmutableArray<T>."
    Set ImmutableArrayOf = CreateSequenceCollection( _
        ROLE_IMMUTABLE, COLLECTION_IMMUTABLE_ARRAY, typeToken)
End Function

Public Function ImmutableListOf(ByVal typeToken As Variant) As ROneCOne
Attribute ImmutableListOf.VB_Description = "Creates an empty ImmutableList<T>."
    Set ImmutableListOf = CreateSequenceCollection( _
        ROLE_IMMUTABLE, COLLECTION_IMMUTABLE_LIST, typeToken)
End Function

Public Function ImmutableDictionaryOf( _
    ByVal keyType As Variant, _
    ByVal valueType As Variant, _
    Optional ByVal comparer As Variant _
) As ROneCOne
Attribute ImmutableDictionaryOf.VB_Description = _
    "Creates an ImmutableDictionary<TKey, TValue>."
    Set ImmutableDictionaryOf = CreateMapCollection( _
        ROLE_IMMUTABLE, COLLECTION_IMMUTABLE_DICTIONARY, _
        keyType, valueType, comparer)
End Function

Public Function ImmutableHashSetOf( _
    ByVal typeToken As Variant, _
    Optional ByVal comparer As Variant _
) As ROneCOne
Attribute ImmutableHashSetOf.VB_Description = _
    "Creates an ImmutableHashSet<T>."
    Set ImmutableHashSetOf = CreateSequenceCollection( _
        ROLE_IMMUTABLE, COLLECTION_IMMUTABLE_HASH_SET, typeToken, comparer)
End Function

Public Function ImmutableQueueOf(ByVal typeToken As Variant) As ROneCOne
Attribute ImmutableQueueOf.VB_Description = "Creates an empty ImmutableQueue<T>."
    Set ImmutableQueueOf = CreateSequenceCollection( _
        ROLE_IMMUTABLE, COLLECTION_IMMUTABLE_QUEUE, typeToken)
End Function

Public Function ImmutableStackOf(ByVal typeToken As Variant) As ROneCOne
Attribute ImmutableStackOf.VB_Description = "Creates an empty ImmutableStack<T>."
    Set ImmutableStackOf = CreateSequenceCollection( _
        ROLE_IMMUTABLE, COLLECTION_IMMUTABLE_STACK, typeToken)
End Function

Public Function ImmutableSortedDictionaryOf( _
    ByVal keyType As Variant, _
    ByVal valueType As Variant, _
    Optional ByVal comparer As Variant _
) As ROneCOne
Attribute ImmutableSortedDictionaryOf.VB_Description = _
    "Creates an ImmutableSortedDictionary<TKey, TValue>."
    Set ImmutableSortedDictionaryOf = CreateMapCollection( _
        ROLE_IMMUTABLE, COLLECTION_IMMUTABLE_SORTED_DICTIONARY, _
        keyType, valueType, comparer)
End Function

Public Function ImmutableSortedSetOf( _
    ByVal typeToken As Variant, _
    Optional ByVal comparer As Variant _
) As ROneCOne
Attribute ImmutableSortedSetOf.VB_Description = _
    "Creates an ImmutableSortedSet<T>."
    Set ImmutableSortedSetOf = CreateSequenceCollection( _
        ROLE_IMMUTABLE, COLLECTION_IMMUTABLE_SORTED_SET, typeToken, comparer)
End Function

Public Function EmptyOf(ByVal typeToken As Variant) As ROneCOne
Attribute EmptyOf.VB_Description = "Creates an empty typed enumerable in legal VBA syntax."
    Set EmptyOf = CreateSequenceCollection( _
        ROLE_IMMUTABLE, COLLECTION_IMMUTABLE_ARRAY, typeToken)
End Function

Public Function Range(ByVal start As Long, ByVal count As Long) As ROneCOne
Attribute Range.VB_Description = "Creates a sequence of consecutive Long values."
    Dim index As Long
    Dim listValue As ROneCOne

    If count < 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "Range", _
            "Count cannot be negative."
    End If
    If count > 0 Then
        If CDbl(start) + CDbl(count) - 1# > CDbl(2147483647) Then
            RaiseContractError ERROR_INVALID_ARGUMENT, "Range", _
                "The requested range exceeds the Long data type."
        End If
    End If

    Set listValue = ROneCOne.ListOf(vbLong)
    For index = 0 To count - 1
        listValue.Add CLng(CDbl(start) + CDbl(index))
    Next index
    Set Range = listValue
End Function

Public Function Repeat(ByVal element As Variant, ByVal count As Long) As ROneCOne
Attribute Repeat.VB_Description = "Creates a sequence containing a repeated element."
    Dim index As Long
    Dim listValue As ROneCOne

    If count < 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "Repeat", _
            "Count cannot be negative."
    End If
    Set listValue = ROneCOne.ListLike(element)
    For index = 1 To count
        listValue.Add element
    Next index
    Set Repeat = listValue
End Function

Public Function Lambda( _
    ByVal body As ROneCOne, _
    ParamArray parameters() As Variant _
) As ROneCOne
Attribute Lambda.VB_Description = "Builds an immutable expression-tree lambda delegate."
    Dim delegateValue As ROneCOne
    Dim parameterList As Collection
    Dim parameterValue As ROneCOne
    Dim rawParameter As Variant

    If body Is Nothing Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "Lambda", _
            "The lambda body cannot be Nothing."
    End If

    Set parameterList = New Collection
    For Each rawParameter In parameters
        Set parameterValue = RequireROneCOne( _
            rawParameter, _
            "Lambda parameter")
        If parameterValue.InternalRole <> ROLE_PARAMETER Then
            RaiseContractError ERROR_INVALID_ARGUMENT, "Lambda", _
                "Every lambda parameter must be created by ROneCOne.Parameter."
        End If
        If CollectionContainsIdentity(parameterList, parameterValue) Then
            RaiseContractError ERROR_INVALID_ARGUMENT, "Lambda", _
                "A lambda parameter cannot be declared more than once."
        End If
        parameterList.Add parameterValue
    Next rawParameter

    If parameterList.Count = 0 Then
        CollectExpressionParameters body, parameterList
    End If

    ValidateBindings body, parameterList

    Set delegateValue = New ROneCOne
    delegateValue.ConfigureLambda body, parameterList
    Set Lambda = delegateValue
End Function

Public Function AsFunc() As ROneCOne
Attribute AsFunc.VB_Description = "Infers parameters and turns this expression into a Func."
    RequireExpression Me, "AsFunc"
    Set AsFunc = ROneCOne.Lambda(Me)
End Function

Public Function Func( _
    ByVal callable As Variant, _
    Optional ByVal memberName As String = vbNullString _
) As ROneCOne
Attribute Func.VB_Description = "Creates a Func from an expression, object, or procedure."
    Set Func = CreateCallableDelegate(callable, memberName, False)
End Function

Public Function Action( _
    ByVal callable As Variant, _
    Optional ByVal memberName As String = vbNullString _
) As ROneCOne
Attribute Action.VB_Description = "Creates an Action from an expression, object, or procedure."
    Set Action = CreateCallableDelegate(callable, memberName, True)
End Function

Public Function EventOf(ParamArray typeTokens() As Variant) As ROneCOne
Attribute EventOf.VB_Description = "Creates a strongly typed mutable event."
    Dim descriptor As ROneCOne
    Dim descriptors As Collection
    Dim eventValue As ROneCOne
    Dim typeToken As Variant

    Set descriptors = New Collection
    For Each typeToken In typeTokens
        Set descriptor = CreateParameterDescriptor(typeToken, "EventOf")
        If descriptor.InternalIsByRef Then
            RaiseContractError ERROR_INVALID_ARGUMENT, "EventOf", _
                "Events support ByVal parameters only."
        End If
        descriptors.Add descriptor
    Next typeToken
    Set eventValue = New ROneCOne
    eventValue.ConfigureEvent descriptors
    Set EventOf = eventValue
End Function

Public Function Try(ByVal work As ROneCOne) As ROneCOne
Attribute Try.VB_Description = "Begins a structured Try/Catch/Finally operation."
    Dim operation As ROneCOne

    If mRole <> ROLE_EMPTY Then
        RaiseContractError ERROR_INVALID_OPERATION, "Try", _
            "Try must be called from the ROneCOne factory."
    End If
    ValidateAction work, "Try", True
    Set operation = New ROneCOne
    operation.ConfigureTry work
    Set Try = operation
End Function

Public Property Get Exception() As ROneCOne
Attribute Exception.VB_Description = "Gets the strongly typed exception prototype."
    Dim errorValue As ROneCOne

    If mRole = ROLE_TASK Then
        If mTaskStatus = TASK_STATUS_FAULTED Then
            Set Exception = AsAggregateException(mTaskException)
        End If
        Exit Property
    End If

    Set errorValue = New ROneCOne
    errorValue.ConfigureError 0, vbNullString, vbNullString, vbNullString, 0
    Set Exception = errorValue
End Property

Public Property Get Task() As ROneCOne
Attribute Task.VB_Description = "Gets the C#-style Task factory surface."
    Dim factory As ROneCOne

    If mRole = ROLE_TASK_COMPLETION_SOURCE Then
        Set Task = mSource
        Exit Property
    End If
    Set factory = New ROneCOne
    factory.ConfigureTaskFactory
    Set Task = factory
End Property

Public Property Get CompletedTask() As ROneCOne
Attribute CompletedTask.VB_Description = "Gets a successfully completed Task."
    Set CompletedTask = FromResult(Empty)
End Property

Public Function FromResult(ByVal value As Variant) As ROneCOne
Attribute FromResult.VB_Description = "Creates a completed Task<T> from a value."
    Dim taskValue As ROneCOne

    Set taskValue = New ROneCOne
    taskValue.ConfigureCompletedTask WrapInvocationArgument(value)
    Set FromResult = taskValue
End Function

Public Function Using(ByVal resource As ROneCOne) As ROneCOne
Attribute Using.VB_Description = "Creates a deterministic resource scope."
    Dim scope As ROneCOne

    If mRole <> ROLE_EMPTY Then
        RaiseContractError ERROR_INVALID_OPERATION, "Using", _
            "Using must be called from the ROneCOne factory."
    End If
    If resource Is Nothing Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "Using", _
            "The resource cannot be Nothing."
    End If
    Set scope = New ROneCOne
    scope.ConfigureUsing resource
    Set Using = scope
End Function

Friend Function ScheduleTaskWork( _
    ByVal work As ROneCOne, _
    Optional ByVal cancellationToken As ROneCOne _
) As ROneCOne
    Dim taskValue As ROneCOne

    ValidateTaskWork work, "Task.Run"
    ValidateCancellationToken cancellationToken, "Task.Run"
    Set taskValue = New ROneCOne
    taskValue.ConfigureTask TASK_WORK, work, cancellationToken
    Set ScheduleTaskWork = taskValue
End Function

Public Function Delay( _
    ByVal milliseconds As Long, _
    Optional ByVal cancellationToken As ROneCOne _
) As ROneCOne
Attribute Delay.VB_Description = "Creates a cooperative delay Task."
    Dim taskValue As ROneCOne

    If milliseconds < 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "Delay", _
            "The delay cannot be negative."
    End If
    ValidateCancellationToken cancellationToken, "Delay"
    Set taskValue = New ROneCOne
    taskValue.ConfigureTask TASK_DELAY, Nothing, cancellationToken
    taskValue.ConfigureCollectionCapacity milliseconds
    taskValue.ConfigureTaskDeadline MonotonicMilliseconds + _
        CLngLng(milliseconds)
    Set Delay = taskValue
End Function

Public Function YieldOnce() As ROneCOne
Attribute YieldOnce.VB_Description = "Creates a Task that yields once to Excel."
    Dim taskValue As ROneCOne

    If mRole <> ROLE_TASK_FACTORY Then
        RaiseContractError ERROR_INVALID_OPERATION, "YieldOnce", _
            "YieldOnce must be called from the Task factory."
    End If
    Set taskValue = New ROneCOne
    taskValue.ConfigureTask TASK_YIELD, Nothing, Nothing
    Set YieldOnce = taskValue
End Function

Public Function WhenAll(ParamArray tasks() As Variant) As ROneCOne
Attribute WhenAll.VB_Description = "Creates a Task that completes after every input Task."
    Set WhenAll = CreateTaskCombinator(tasks, TASK_WHEN_ALL, "WhenAll")
End Function

Public Function WhenAny(ParamArray tasks() As Variant) As ROneCOne
Attribute WhenAny.VB_Description = "Creates a Task that completes after one input Task."
    Set WhenAny = CreateTaskCombinator(tasks, TASK_WHEN_ANY, "WhenAny")
End Function

Public Function Await() As Variant
Attribute Await.VB_Description = "Cooperatively awaits completion and returns the Task result."
    RequireTask Me, "Await"
    EnsureTaskCompleted
    RaiseTaskTerminalState "Await"
    AssignPublicResult Await, mTaskResult
End Function

Public Function Wait(Optional ByVal timeoutMilliseconds As Long = -1) As Boolean
Attribute Wait.VB_Description = "Waits cooperatively for Task completion."
    RequireTask Me, "Wait"
    If timeoutMilliseconds < -1 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "Wait", _
            "The timeout must be -1 or nonnegative."
    End If
    If Not WaitForTaskCompletion(timeoutMilliseconds) Then Exit Function
    RaiseTaskTerminalState "Wait"
    Wait = True
End Function

Public Function WaitAsync( _
    Optional ByVal timeoutMilliseconds As Long = -1, _
    Optional ByVal cancellationToken As ROneCOne _
) As ROneCOne
Attribute WaitAsync.VB_Description = "Creates a Task that applies a timeout or cancellation."
    Dim taskValue As ROneCOne

    RequireTask Me, "WaitAsync"
    If timeoutMilliseconds < -1 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "WaitAsync", _
            "The timeout must be -1 or nonnegative."
    End If
    ValidateCancellationToken cancellationToken, "WaitAsync"
    Set taskValue = New ROneCOne
    taskValue.ConfigureTask TASK_WAIT_ASYNC, Nothing, cancellationToken
    taskValue.ConfigureReadOnlySource Me
    If timeoutMilliseconds >= 0 Then
        taskValue.ConfigureTaskDeadline MonotonicMilliseconds + _
            CLngLng(timeoutMilliseconds)
    Else
        taskValue.ConfigureTaskDeadline -1
    End If
    Set WaitAsync = taskValue
End Function

Public Function ContinueWith(ByVal continuation As ROneCOne) As ROneCOne
Attribute ContinueWith.VB_Description = "Creates a continuation Task."
    Dim taskValue As ROneCOne

    RequireTask Me, "ContinueWith"
    If continuation Is Nothing Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "ContinueWith", _
            "The continuation cannot be Nothing."
    End If
    RequireDelegate continuation, "ContinueWith"
    If continuation.Arity <> 1 And continuation.Arity <> -1 Then
        RaiseContractError ERROR_ARITY_MISMATCH, "ContinueWith", _
            "The continuation must accept the antecedent Task."
    End If
    Set taskValue = New ROneCOne
    taskValue.ConfigureTask TASK_CONTINUATION, continuation, Nothing
    taskValue.ConfigureReadOnlySource Me
    Set ContinueWith = taskValue
End Function

Public Property Get Status() As String
Attribute Status.VB_Description = "Gets the .NET-aligned Task status name."
    RequireTask Me, "Status"
    Select Case mTaskStatus
        Case TASK_STATUS_CREATED
            Status = "Created"
        Case TASK_STATUS_RUNNING
            Status = "Running"
        Case TASK_STATUS_COMPLETED
            Status = "RanToCompletion"
        Case TASK_STATUS_CANCELED
            Status = "Canceled"
        Case TASK_STATUS_FAULTED
            Status = "Faulted"
    End Select
End Property

Public Property Get IsCompleted() As Boolean
Attribute IsCompleted.VB_Description = "Reports terminal Task or blocking-collection state."
    If mRole = ROLE_TASK Then
        IsCompleted = (mTaskStatus >= TASK_STATUS_COMPLETED)
    ElseIf mCollectionKind = COLLECTION_BLOCKING Then
        IsCompleted = (mAddingCompleted And mItemsCount = 0)
    Else
        RaiseContractError ERROR_INVALID_OPERATION, "IsCompleted", _
            "IsCompleted requires a Task or BlockingCollection."
    End If
End Property

Public Property Get IsCanceled() As Boolean
Attribute IsCanceled.VB_Description = "Reports whether a Task was canceled."
    RequireTask Me, "IsCanceled"
    IsCanceled = (mTaskStatus = TASK_STATUS_CANCELED)
End Property

Public Property Get IsFaulted() As Boolean
Attribute IsFaulted.VB_Description = "Reports whether a Task faulted."
    RequireTask Me, "IsFaulted"
    IsFaulted = (mTaskStatus = TASK_STATUS_FAULTED)
End Property

Public Property Get Result() As Variant
Attribute Result.VB_Description = "Waits for and returns a Task<T> result."
    RequireTask Me, "Result"
    EnsureTaskCompleted
    RaiseTaskTerminalState "Result"
    AssignPublicResult Result, mTaskResult
End Property

Public Function CancellationTokenSource() As ROneCOne
Attribute CancellationTokenSource.VB_Description = "Creates a CancellationTokenSource."
    Dim source As ROneCOne

    Set source = New ROneCOne
    source.ConfigureCancellationSource
    Set CancellationTokenSource = source
End Function

Public Function TaskCompletionSourceOf(ByVal typeToken As Variant) As ROneCOne
Attribute TaskCompletionSourceOf.VB_Description = _
    "Creates a typed cooperative TaskCompletionSource<T>."
    Dim objectType As String
    Dim source As ROneCOne
    Dim valueType As VbVarType

    ResolveTypeDescriptor typeToken, valueType, objectType, _
        "TaskCompletionSourceOf"
    Set source = New ROneCOne
    source.ConfigureTaskCompletionSource valueType, objectType
    Set TaskCompletionSourceOf = source
End Function

Public Function ProgressOf( _
    ByVal typeToken As Variant, _
    ByVal handler As ROneCOne _
) As ROneCOne
Attribute ProgressOf.VB_Description = "Creates a typed synchronous progress reporter."
    Dim objectType As String
    Dim progressValue As ROneCOne
    Dim valueType As VbVarType

    ResolveTypeDescriptor typeToken, valueType, objectType, "ProgressOf"
    ValidateProgressHandler handler, valueType, objectType
    Set progressValue = New ROneCOne
    progressValue.ConfigureProgress valueType, objectType, handler
    Set ProgressOf = progressValue
End Function

Public Property Get Token() As ROneCOne
Attribute Token.VB_Description = "Gets the token owned by a CancellationTokenSource."
    Dim tokenValue As ROneCOne

    If mRole <> ROLE_CANCELLATION_SOURCE Then
        RaiseContractError ERROR_INVALID_OPERATION, "Token", _
            "Token requires a CancellationTokenSource."
    End If
    Set tokenValue = New ROneCOne
    tokenValue.ConfigureCancellationToken Me
    Set Token = tokenValue
End Property

Public Sub Cancel()
Attribute Cancel.VB_Description = "Requests cancellation and runs registered callbacks."
    If mRole <> ROLE_CANCELLATION_SOURCE Then
        RaiseContractError ERROR_INVALID_OPERATION, "Cancel", _
            "Cancel requires a CancellationTokenSource."
    End If
    CancelCancellationSource
End Sub

Public Sub CancelAfter(ByVal milliseconds As Long)
Attribute CancelAfter.VB_Description = "Schedules cooperative cancellation after a delay."
    If mRole <> ROLE_CANCELLATION_SOURCE Then
        RaiseContractError ERROR_INVALID_OPERATION, "CancelAfter", _
            "CancelAfter requires a CancellationTokenSource."
    End If
    If milliseconds < -1 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "CancelAfter", _
            "The delay must be -1 or nonnegative."
    End If
    mQueryCount = milliseconds
    If milliseconds >= 0 Then
        mTaskDeadline = MonotonicMilliseconds + CLngLng(milliseconds)
    Else
        mTaskDeadline = -1
    End If
End Sub

Public Property Get IsCancellationRequested() As Boolean
Attribute IsCancellationRequested.VB_Description = "Reports whether cancellation was requested."
    If mRole = ROLE_CANCELLATION_SOURCE Then
        EvaluateScheduledCancellation
        IsCancellationRequested = mCancellationRequested
    ElseIf mRole = ROLE_CANCELLATION_TOKEN Then
        IsCancellationRequested = mCancellationSource.IsCancellationRequested
    Else
        RaiseContractError ERROR_INVALID_OPERATION, "IsCancellationRequested", _
            "This property requires a cancellation source or token."
    End If
End Property

Public Function Register(ByVal callback As ROneCOne) As ROneCOne
Attribute Register.VB_Description = "Registers a callback for cooperative cancellation."
    Dim registration As ROneCOne
    Dim source As ROneCOne

    If mRole = ROLE_CANCELLATION_SOURCE Then
        Set source = Me
    ElseIf mRole = ROLE_CANCELLATION_TOKEN Then
        Set source = mCancellationSource
    Else
        RaiseContractError ERROR_INVALID_OPERATION, "Register", _
            "Register requires a cancellation source or token."
    End If
    ValidateAction callback, "Register", True
    Set registration = New ROneCOne
    registration.ConfigureCancellationRegistration source, callback
    source.RegisterCancellationCallback registration
    Set Register = registration
End Function

Public Sub ThrowIfCancellationRequested()
Attribute ThrowIfCancellationRequested.VB_Description = _
    "Raises when cooperative cancellation was requested."
    If Not IsCancellationRequested Then Exit Sub
    RaiseContractError ERROR_INVALID_OPERATION, _
        "ThrowIfCancellationRequested", "Cancellation was requested."
End Sub

Public Function TrySetResult(ByVal value As Variant) As Boolean
Attribute TrySetResult.VB_Description = "Attempts to complete a TaskCompletionSource<T>."
    RequireTaskCompletionSource "TrySetResult"
    TrySetResult = mSource.TryCompletePromise( _
        WrapTypedCompletionValue(value), TASK_STATUS_COMPLETED, Nothing)
End Function

Public Sub SetResult(ByVal value As Variant)
Attribute SetResult.VB_Description = "Completes a TaskCompletionSource<T> successfully."
    If Not TrySetResult(value) Then
        RaiseContractError ERROR_INVALID_OPERATION, "SetResult", _
            "The TaskCompletionSource is already complete."
    End If
End Sub

Public Function TrySetCanceled() As Boolean
Attribute TrySetCanceled.VB_Description = "Attempts to cancel a TaskCompletionSource<T>."
    RequireTaskCompletionSource "TrySetCanceled"
    TrySetCanceled = mSource.TryCompletePromise( _
        Nothing, TASK_STATUS_CANCELED, Nothing)
End Function

Public Sub SetCanceled()
Attribute SetCanceled.VB_Description = "Cancels a TaskCompletionSource<T>."
    If Not TrySetCanceled Then
        RaiseContractError ERROR_INVALID_OPERATION, "SetCanceled", _
            "The TaskCompletionSource is already complete."
    End If
End Sub

Public Function TrySetException(ByVal errorValue As ROneCOne) As Boolean
Attribute TrySetException.VB_Description = _
    "Attempts to fault a TaskCompletionSource<T>."
    RequireTaskCompletionSource "TrySetException"
    If errorValue Is Nothing Or errorValue.InternalRole <> ROLE_ERROR Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "TrySetException", _
            "TrySetException requires a ROneCOne exception value."
    End If
    TrySetException = mSource.TryCompletePromise( _
        Nothing, TASK_STATUS_FAULTED, errorValue)
End Function

Public Sub SetException(ByVal errorValue As ROneCOne)
Attribute SetException.VB_Description = "Faults a TaskCompletionSource<T>."
    If Not TrySetException(errorValue) Then
        RaiseContractError ERROR_INVALID_OPERATION, "SetException", _
            "The TaskCompletionSource is already complete."
    End If
End Sub

Public Sub Report(ByVal value As Variant)
Attribute Report.VB_Description = "Reports one typed progress value."
    Dim arguments As Collection
    Dim ignored As ROneCOne
    Dim wrapped As ROneCOne

    If mRole <> ROLE_PROGRESS Then
        RaiseContractError ERROR_INVALID_OPERATION, "Report", _
            "Report requires a Progress<T> value."
    End If
    Set wrapped = WrapInvocationArgument(value)
    Set wrapped = WidenScalarToward(wrapped, mExpectedType)
    ValidateWrappedType wrapped, mExpectedType, mExpectedObjectType, _
        "Report", "progress value"
    Set arguments = New Collection
    arguments.Add wrapped
    Set ignored = mTaskBody.InvokeValue(arguments)
End Sub

Public Property Get DBNull() As ROneCOne
Attribute DBNull.VB_Description = "Gets the explicit database-null sentinel."
    Dim nullValue As ROneCOne

    Set nullValue = New ROneCOne
    nullValue.ConfigureDBNull
    Set DBNull = nullValue
End Property

Public Function DataTable(Optional ByVal tableName As String = vbNullString) As ROneCOne
Attribute DataTable.VB_Description = "Creates an in-memory DataTable."
    Dim tableValue As ROneCOne

    Set tableValue = New ROneCOne
    tableValue.ConfigureDataTable tableName
    Set DataTable = tableValue
End Function

Public Function Table(ByVal listObject As Object) As ROneCOne
Attribute Table.VB_Description = _
    "Reads an Excel Table into a DataTable that stays attached to it."
    Dim tableValue As ROneCOne

    If Not IsExcelTable(listObject) Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "Table", _
            "Table needs an Excel Table, for example " & _
            "Sheet1.ListObjects(""Sales"")."
    End If
    Set tableValue = ROneCOne.DataTableFromRange(listObject, True, listObject.Name)
    tableValue.AttachSourceTable listObject
    Set Table = tableValue
End Function

Public Sub Refresh()
Attribute Refresh.VB_Description = "Re-reads this table in place from the Excel Table it came from."
    ' In place rather than returning a copy, so a bare Refresh statement
    ' cannot silently do nothing.
    InternalFillFromRange RequireSourceTable("Refresh"), True, "Refresh"
End Sub

Public Function WriteBack(Optional ByVal source As Variant) As Long
Attribute WriteBack.VB_Description = _
    "Writes rows into the Excel Table this came from, resizing it to fit."
    Dim listObject As Object
    Dim writer As ROneCOne

    Set listObject = RequireSourceTable("WriteBack")
    If IsMissing(source) Then
        Set writer = Me
    ElseIf IsObject(source) Then
        Set writer = source
    Else
        RaiseContractError ERROR_INVALID_ARGUMENT, "WriteBack", _
            "WriteBack needs a DataTable or DataView, or no argument to " & _
            "write this table's own rows."
    End If
    WriteBack = writer.ToRange(listObject)
End Function

Private Function RequireSourceTable(ByVal memberName As String) As Object
    If mSourceTable Is Nothing Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            memberName & " needs a table read with ROneCOne.Table(listObject), " & _
            "which is what remembers where the rows came from."
    End If
    Set RequireSourceTable = mSourceTable
End Function

Public Function DataColumn( _
    ByVal columnName As String, _
    ByVal typeToken As Variant _
) As ROneCOne
Attribute DataColumn.VB_Description = "Creates a typed DataColumn."
    Dim columnValue As ROneCOne
    Dim objectType As String
    Dim valueType As VbVarType

    columnName = Trim$(columnName)
    If Len(columnName) = 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "DataColumn", _
            "The column name cannot be empty."
    End If
    ResolveTypeDescriptor typeToken, valueType, objectType, "DataColumn"
    Set columnValue = New ROneCOne
    columnValue.ConfigureDataColumn columnName, valueType, objectType
    Set DataColumn = columnValue
End Function

Public Function DataSet(Optional ByVal dataSetName As String = vbNullString) As ROneCOne
Attribute DataSet.VB_Description = "Creates an in-memory relational DataSet."
    Dim dataValue As ROneCOne

    Set dataValue = New ROneCOne
    dataValue.ConfigureDataSet dataSetName
    Set DataSet = dataValue
End Function

Public Function DataRelation( _
    ByVal relationName As String, _
    ByVal parentColumns As Variant, _
    ByVal childColumns As Variant, _
    Optional ByVal createConstraints As Boolean = True _
) As ROneCOne
Attribute DataRelation.VB_Description = "Creates a parent-child DataRelation."
    Dim childColumn As ROneCOne
    Dim childValues As Collection
    Dim index As Long
    Dim parentColumn As ROneCOne
    Dim parentValues As Collection
    Dim relationValue As ROneCOne

    Set parentValues = NormalizeRelationColumns( _
        parentColumns, "DataRelation parent")
    Set childValues = NormalizeRelationColumns( _
        childColumns, "DataRelation child")
    If parentValues.Count <> childValues.Count Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "DataRelation", _
            "Parent and child relations require the same column count."
    End If
    For index = 1 To parentValues.Count
        Set parentColumn = parentValues.Item(index)
        Set childColumn = childValues.Item(index)
        If parentColumn.InternalExpectedType <> childColumn.InternalExpectedType Or _
            parentColumn.InternalExpectedObjectType <> _
                childColumn.InternalExpectedObjectType Then
            RaiseContractError ERROR_TYPE_MISMATCH, "DataRelation", _
                "Paired parent and child columns must have the same data type."
        End If
    Next index
    Set relationValue = New ROneCOne
    relationValue.ConfigureDataRelation relationName, parentValues, _
        childValues, createConstraints
    Set DataRelation = relationValue
End Function

Public Function DataView(ByVal table As ROneCOne) As ROneCOne
Attribute DataView.VB_Description = "Creates a live DataView over a DataTable."
    Dim viewValue As ROneCOne

    RequireDataTable table, "DataView"
    Set viewValue = New ROneCOne
    viewValue.ConfigureDataView table
    Set DataView = viewValue
End Function

Public Function AsUnique( _
    Optional ByVal allowNull As Boolean = False _
) As ROneCOne
Attribute AsUnique.VB_Description = "Marks a DataColumn as unique and fluent."
    RequireDataColumn Me, "AsUnique"
    mUnique = True
    If Not allowNull Then mAllowDBNull = False
    If Not mSource Is Nothing Then mSource.InternalMarkDataIndexDirty
    Set AsUnique = Me
End Function

Public Function WithDefault(ByVal value As Variant) As ROneCOne
Attribute WithDefault.VB_Description = "Sets a DataColumn default value fluently."
    Dim wrapped As ROneCOne

    RequireDataColumn Me, "WithDefault"
    Set wrapped = WrapDataValue(value)
    Set wrapped = WidenScalarToward(wrapped, mExpectedType)
    ValidateDataColumnValue Me, wrapped, "WithDefault"
    Set mDefaultValue = wrapped
    Set WithDefault = Me
End Function

Public Function AutoNumber( _
    Optional ByVal seed As LongLong = 0, _
    Optional ByVal stepValue As LongLong = 1 _
) As ROneCOne
Attribute AutoNumber.VB_Description = "Configures an auto-increment integer DataColumn."
    RequireDataColumn Me, "AutoNumber"
    Select Case mExpectedType
        Case vbByte, vbInteger, vbLong, vbLongLong
        Case Else
            RaiseContractError ERROR_TYPE_MISMATCH, "AutoNumber", _
                "AutoNumber requires an integer DataColumn."
    End Select
    If stepValue = 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "AutoNumber", _
            "The auto-increment step cannot be zero."
    End If
    mAutoIncrement = True
    mAutoIncrementSeed = seed
    mAutoIncrementStep = stepValue
    mAutoIncrementCurrent = seed
    mAllowDBNull = False
    Set AutoNumber = Me
End Function

Public Function AsPrimaryKey( _
    Optional ByVal enabled As Boolean = True _
) As ROneCOne
Attribute AsPrimaryKey.VB_Description = "Makes this column the table primary key."
    RequireDataColumn Me, "AsPrimaryKey"
    If Not enabled Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "AsPrimaryKey", _
            "AsPrimaryKey does not accept False."
    End If
    If mSource Is Nothing Or mSource.InternalRole <> ROLE_DATA_TABLE Then
        RaiseContractError ERROR_INVALID_OPERATION, "AsPrimaryKey", _
            "The DataColumn must belong to a DataTable."
    End If
    mSource.ConfigureSinglePrimaryKey Me
    Set AsPrimaryKey = Me
End Function

Public Property Get AutoIncrement() As Boolean
Attribute AutoIncrement.VB_Description = "Reports whether a DataColumn auto-increments."
    RequireDataColumn Me, "AutoIncrement"
    AutoIncrement = mAutoIncrement
End Property

Public Property Get DefaultValue() As Variant
Attribute DefaultValue.VB_Description = "Gets a DataColumn default value."
    RequireDataColumn Me, "DefaultValue"
    If mDefaultValue Is Nothing Then
        DefaultValue = Empty
    Else
        AssignPublicResult DefaultValue, mDefaultValue
    End If
End Property

Public Property Get ColumnName() As String
Attribute ColumnName.VB_Description = "Gets a DataColumn name."
    RequireDataColumn Me, "ColumnName"
    ColumnName = mName
End Property

Public Property Let ColumnName(ByVal value As String)
    RequireDataColumn Me, "ColumnName"
    value = Trim$(value)
    If Len(value) = 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "ColumnName", _
            "The column name cannot be empty."
    End If
    mName = value
End Property

Public Property Get TableName() As String
Attribute TableName.VB_Description = "Gets a DataTable name."
    RequireDataTable Me, "TableName"
    TableName = mName
End Property

Public Property Get DataSetName() As String
Attribute DataSetName.VB_Description = "Gets a DataSet name."
    If mRole <> ROLE_DATA_SET Then
        RaiseContractError ERROR_INVALID_OPERATION, "DataSetName", _
            "DataSetName requires a DataSet."
    End If
    DataSetName = mName
End Property

Public Property Get EnforceConstraints() As Boolean
Attribute EnforceConstraints.VB_Description = "Gets or sets DataSet constraint enforcement."
    If mRole <> ROLE_DATA_SET Then
        RaiseContractError ERROR_INVALID_OPERATION, "EnforceConstraints", _
            "EnforceConstraints requires a DataSet."
    End If
    EnforceConstraints = mUnique
End Property

Public Property Let EnforceConstraints(ByVal value As Boolean)
    If mRole <> ROLE_DATA_SET Then
        RaiseContractError ERROR_INVALID_OPERATION, "EnforceConstraints", _
            "EnforceConstraints requires a DataSet."
    End If
    mUnique = value
End Property

Public Property Get DataTypeName() As String
Attribute DataTypeName.VB_Description = "Gets a DataColumn generic type name."
    RequireDataColumn Me, "DataTypeName"
    If Len(mExpectedObjectType) > 0 Then
        DataTypeName = mExpectedObjectType
    Else
        DataTypeName = TypeNameForVarType(mExpectedType)
    End If
End Property

Public Property Get AllowDBNull() As Boolean
Attribute AllowDBNull.VB_Description = "Gets or sets whether a DataColumn accepts Null."
    RequireDataColumn Me, "AllowDBNull"
    AllowDBNull = mAllowDBNull
End Property

Public Property Let AllowDBNull(ByVal value As Boolean)
    RequireDataColumn Me, "AllowDBNull"
    mAllowDBNull = value
End Property

Public Property Get Unique() As Boolean
Attribute Unique.VB_Description = "Gets or sets DataColumn uniqueness."
    RequireDataColumn Me, "Unique"
    Unique = mUnique
End Property

Public Property Let Unique(ByVal value As Boolean)
    RequireDataColumn Me, "Unique"
    mUnique = value
End Property

Public Sub AddColumn(ByVal column As ROneCOne)
Attribute AddColumn.VB_Description = "Adds a DataColumn to a DataTable schema."
    Dim index As Long

    RequireDataTable Me, "AddColumn"
    RequireDataColumn column, "AddColumn"
    If FindNamedObjectIndex(ArrSnapshot(mKeys, mKeysCount), column.ColumnName) >= 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "AddColumn", _
            "A column with the same name already exists."
    End If
    If Not column.InternalSource Is Nothing Then
        RaiseContractError ERROR_INVALID_OPERATION, "AddColumn", _
            "The DataColumn already belongs to a table."
    End If
    column.ConfigureReadOnlySource Me
    ArrAppend mKeys, mKeysCount, column
    ' Rows must stay exactly as wide as the schema, so a column added after
    ' rows exist backfills every row with the column's default cell.
    For index = 1 To mItemsCount
        mItems(index).InternalAppendRowCell DefaultDataColumnValue(column)
    Next index
    If column.InternalUnique Then mDataIndexDirty = True
    mVersion = mVersion + 1
End Sub

Public Function Column( _
    ByVal columnName As String, _
    ByVal typeToken As Variant _
) As ROneCOne
Attribute Column.VB_Description = "Defines and returns a typed DataColumn."
    Dim columnValue As ROneCOne

    RequireDataTable Me, "Column"
    Set columnValue = ROneCOne.DataColumn(columnName, typeToken)
    AddColumn columnValue
    Set Column = columnValue
End Function

Public Function NewRow() As ROneCOne
Attribute NewRow.VB_Description = "Creates a detached DataRow for a DataTable."
    Dim rowValue As ROneCOne

    RequireDataTable Me, "NewRow"
    Set rowValue = New ROneCOne
    rowValue.ConfigureDataRow Me
    Set NewRow = rowValue
End Function

Public Function Row(ParamArray values() As Variant) As ROneCOne
Attribute Row.VB_Description = "Creates a detached DataRow from positional values."
    Dim index As Long
    Dim column As ROneCOne
    Dim columnIndex As Long
    Dim nonAutoCount As Long
    Dim rowValue As ROneCOne
    Dim valueCount As Long
    Dim valueIndex As Long

    RequireDataTable Me, "Row"
    Set rowValue = NewRow
    On Error Resume Next
    valueCount = UBound(values) - LBound(values) + 1
    If Err.Number <> 0 Then
        Err.Clear
        valueCount = 0
    End If
    On Error GoTo 0
    If valueCount > mKeysCount Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "Row", _
            "The row contains more values than the table has columns."
    End If
    For Each column In ArrSnapshot(mKeys, mKeysCount)
        If Not column.InternalAutoIncrement Then nonAutoCount = nonAutoCount + 1
    Next column
    If valueCount > 0 And nonAutoCount < mKeysCount And _
        valueCount <= nonAutoCount Then
        valueIndex = LBound(values)
        For columnIndex = 1 To mKeysCount
            Set column = mKeys(columnIndex)
            If Not column.InternalAutoIncrement And valueIndex <= UBound(values) Then
                rowValue.Item(columnIndex - 1) = values(valueIndex)
                valueIndex = valueIndex + 1
            End If
        Next columnIndex
    ElseIf valueCount > 0 Then
        For index = LBound(values) To UBound(values)
            rowValue.Item(index - LBound(values)) = values(index)
        Next index
    End If
    Set Row = rowValue
End Function

Public Sub AddRow(ByVal row As ROneCOne)
Attribute AddRow.VB_Description = "Adds a detached DataRow after validating constraints."
    RequireDataTable Me, "AddRow"
    RequireDataRow row, "AddRow"
    If Not row.InternalSource Is Me Then
        RaiseContractError ERROR_INVALID_OPERATION, "AddRow", _
            "The DataRow was created by a different table."
    End If
    If row.RowState <> "Detached" Then
        RaiseContractError ERROR_INVALID_OPERATION, "AddRow", _
            "Only a detached DataRow can be added."
    End If
    ValidateDataRowConstraints row, Nothing
    ValidateForeignKeyConstraints row
    ArrAppend mItems, mItemsCount, row
    row.ConfigureRowState ROW_STATE_ADDED
    ' Validation just proved the tokens absent, so the new row slots into
    ' the constraint indexes incrementally; a bulk load stays linear.
    IndexDataRow row
    mVersion = mVersion + 1
End Sub

Public Function LoadRow(ByVal values As Variant) As ROneCOne
Attribute LoadRow.VB_Description = "Creates and adds a DataRow from an array or sequence."
    Dim index As Long
    Dim row As ROneCOne
    Dim wrappedValues As Collection

    RequireDataTable Me, "LoadRow"
    Set wrappedValues = BuildDynamicArguments(values)
    If wrappedValues.Count > mKeysCount Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "LoadRow", _
            "The row contains more values than the table has columns."
    End If
    Set row = NewRow
    For index = 1 To wrappedValues.Count
        row.Item(index - 1) = Unwrap(wrappedValues.Item(index))
    Next index
    AddRow row
    Set LoadRow = row
End Function

Public Function SelectRows(ByVal predicate As Variant) As ROneCOne
Attribute SelectRows.VB_Description = "Returns DataRows matching a typed predicate."
    RequireDataTable Me, "SelectRows"
    Set SelectRows = Rows.Where(predicate).ToList
End Function

Public Function DataTableFromRange( _
    ByVal source As Object, _
    Optional ByVal headers As Boolean = True, _
    Optional ByVal tableName As String = vbNullString _
) As ROneCOne
Attribute DataTableFromRange.VB_Description = _
    "Reads a worksheet range into a new DataTable in one bulk call."
    Dim table As ROneCOne

    Set table = ROneCOne.DataTable(tableName)
    table.InternalFillFromRange source, headers, "DataTableFromRange"
    Set DataTableFromRange = table
End Function

Public Function ListFromRange(ByVal source As Object) As ROneCOne
Attribute ListFromRange.VB_Description = _
    "Reads a single worksheet row or column into a List in one bulk call."
    Dim colCount As Long
    Dim grid As Variant
    Dim index As Long
    Dim listValue As ROneCOne
    Dim rowCount As Long

    ReadRangeGrid ResolveTableRange(source, False, "ListFromRange"), _
        grid, rowCount, colCount, "ListFromRange"
    If rowCount > 1 And colCount > 1 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "ListFromRange", _
            "ListFromRange needs a single row or column; " & _
            "use DataTableFromRange for a block."
    End If
    Set listValue = ROneCOne.ListOf(vbVariant)
    If colCount = 1 Then
        For index = 1 To rowCount
            listValue.Add grid(index, 1)
        Next index
    Else
        For index = 1 To colCount
            listValue.Add grid(1, index)
        Next index
    End If
    Set ListFromRange = listValue
End Function

Public Function LoadFromRange( _
    ByVal source As Object, _
    Optional ByVal headers As Boolean = True _
) As Long
Attribute LoadFromRange.VB_Description = _
    "Appends worksheet-range rows into this typed DataTable in one bulk call."
    Dim colCount As Long
    Dim colIndex As Long
    Dim firstDataRow As Long
    Dim grid As Variant
    Dim row As ROneCOne
    Dim rowCount As Long
    Dim rowIndex As Long

    RequireDataTable Me, "LoadFromRange"
    ReadRangeGrid ResolveTableRange(source, headers, "LoadFromRange"), _
        grid, rowCount, colCount, "LoadFromRange"
    If colCount <> mKeysCount Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "LoadFromRange", _
            "The range column count does not match the table column count."
    End If
    firstDataRow = IIf(headers, 2, 1)
    For rowIndex = firstDataRow To rowCount
        Set row = NewRow
        For colIndex = 1 To colCount
            row.Item(colIndex - 1) = grid(rowIndex, colIndex)
        Next colIndex
        AddRow row
        LoadFromRange = LoadFromRange + 1
    Next rowIndex
End Function

Public Function ToRange( _
    ByVal target As Object, _
    Optional ByVal writeHeaders As Boolean = True _
) As Long
Attribute ToRange.VB_Description = _
    "Writes a DataTable grid or a scalar sequence to a worksheet range."
    ' Writing into an Excel Table is a different operation from writing at a
    ' cell: the table owns its extent, so it is resized to fit rather than
    ' overwritten in place.
    If IsExcelTable(target) Then
        ToRange = WriteToListObject(target, "ToRange")
        Exit Function
    End If
    If mRole = ROLE_DATA_TABLE Then
        ToRange = WriteDataTableToRange(target, writeHeaders)
    ElseIf mRole = ROLE_DATA_VIEW Then
        ToRange = ToTable.ToRange(target, writeHeaders)
    ElseIf mRole = ROLE_LIST Or mRole = ROLE_QUERY Or _
        IsGenericCollectionRole(mRole) Then
        ToRange = WriteSequenceToRange(target)
    Else
        RaiseContractError ERROR_INVALID_OPERATION, "ToRange", _
            "ToRange requires a DataTable, DataView, or typed sequence."
    End If
End Function

Private Sub ReadRangeGrid( _
    ByVal source As Object, _
    ByRef grid As Variant, _
    ByRef rowCount As Long, _
    ByRef colCount As Long, _
    ByVal memberName As String _
)
    Dim raw As Variant
    Dim scalarGrid(1 To 1, 1 To 1) As Variant

    If source Is Nothing Then
        RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
            "The worksheet range cannot be Nothing."
    End If
    raw = source.Value
    If IsArrayValue(raw) Then
        grid = raw
        rowCount = UBound(raw, 1) - LBound(raw, 1) + 1
        colCount = UBound(raw, 2) - LBound(raw, 2) + 1
    Else
        scalarGrid(1, 1) = raw
        grid = scalarGrid
        rowCount = 1
        colCount = 1
    End If
End Sub

Private Function WriteDataTableToRange( _
    ByVal target As Object, _
    ByVal writeHeaders As Boolean _
) As Long
    Dim colCount As Long
    Dim colIndex As Long
    Dim column As ROneCOne
    Dim liveRows As Collection
    Dim output() As Variant
    Dim outputRows As Long
    Dim row As ROneCOne
    Dim rowOffset As Long
    Dim writtenRow As Long

    RequireRangeTarget target, "ToRange"
    colCount = mKeysCount
    Set liveRows = New Collection
    For Each row In ArrSnapshot(mItems, mItemsCount)
        If row.RowState <> "Deleted" Then liveRows.Add row
    Next row
    outputRows = liveRows.Count
    If writeHeaders Then outputRows = outputRows + 1
    If outputRows = 0 Or colCount = 0 Then Exit Function
    ReDim output(1 To outputRows, 1 To colCount)
    If writeHeaders Then
        For colIndex = 1 To colCount
            Set column = mKeys(colIndex)
            output(1, colIndex) = column.ColumnName
        Next colIndex
        rowOffset = 1
    End If
    For writtenRow = 1 To liveRows.Count
        Set row = liveRows.Item(writtenRow)
        For colIndex = 1 To colCount
            output(writtenRow + rowOffset, colIndex) = _
                RangeCellValue(row.InternalDataRows.Item(colIndex))
        Next colIndex
    Next writtenRow
    target.Resize(outputRows, colCount).Value = output
    WriteDataTableToRange = liveRows.Count
End Function

Private Function WriteSequenceToRange(ByVal target As Object) As Long
    Dim index As Long
    Dim output() As Variant
    Dim values As Collection
    Dim wrapped As ROneCOne

    RequireRangeTarget target, "ToRange"
    Set values = MaterializeWrapped
    If values.Count = 0 Then Exit Function
    ReDim output(1 To values.Count, 1 To 1)
    index = 1
    For Each wrapped In values
        output(index, 1) = RangeCellValue(wrapped)
        index = index + 1
    Next wrapped
    target.Resize(values.Count, 1).Value = output
    WriteSequenceToRange = values.Count
End Function

' The only two places the runtime recognizes an Excel Table. A ListObject is
' not a Range and has no Value, so every bridge raised 438 when handed one,
' which reads as the runtime rejecting the table rather than VBA rejecting
' the type. A Range passes straight through, so recognition costs one
' TypeName test on the way in.
Private Function IsExcelTable(ByVal candidate As Object) As Boolean
    If candidate Is Nothing Then Exit Function
    IsExcelTable = (TypeName(candidate) = "ListObject")
End Function

Private Function ResolveTableRange( _
    ByVal source As Object, _
    ByVal wantHeaders As Boolean, _
    ByVal memberName As String _
) As Object
    Dim owner As Object
    Dim resolved As Object

    If source Is Nothing Then
        RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
            "The worksheet range cannot be Nothing."
    End If
    Select Case TypeName(source)
        Case "ListObject"
            Set owner = source
        Case "ListColumn"
            Set owner = source.Parent
        Case Else
            Set ResolveTableRange = source
            Exit Function
    End Select

    If wantHeaders Then
        ' Range spans the totals row when one is showing, and a total is not
        ' data, so the slice stops after the last body row. An empty table
        ' still yields its header row, which reads back as columns and no
        ' rows rather than as a failure.
        Set resolved = source.Range.Resize(1 + owner.ListRows.Count)
    Else
        Set resolved = source.DataBodyRange
    End If
    If resolved Is Nothing Then
        RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
            "The Excel Table '" & owner.Name & "' has no data rows."
    End If
    Set ResolveTableRange = resolved
End Function

' Writes this table's rows into an Excel Table, resizing it to fit. Excel
' refuses to resize a table to nothing and leaves the vacated cells behind
' when one shrinks, and it places a totals row inconsistently across a grow
' and a shrink, so totals are switched off for the duration and the old tail
' is cleared by hand. All three behaviors were measured rather than assumed.
Private Function WriteToListObject( _
    ByVal listObject As Object, _
    ByVal memberName As String _
) As Long
    Dim body() As Variant
    Dim bodyRow As ROneCOne
    Dim colCount As Long
    Dim colIndex As Long
    Dim firstBodyRow As Long
    Dim firstColumn As Long
    Dim hadTotals As Boolean
    Dim liveRows As Collection
    Dim oldRows As Long
    Dim rowIndex As Long
    Dim sheet As Object
    Dim source As ROneCOne

    If mRole = ROLE_DATA_VIEW Then
        Set source = ToTable
    ElseIf mRole = ROLE_DATA_TABLE Then
        Set source = Me
    Else
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            "Writing into an Excel Table requires a DataTable or DataView."
    End If

    colCount = listObject.ListColumns.Count
    Set liveRows = source.InternalLiveDataRows
    If source.Columns.Count <> colCount Then
        RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
            "This value has " & CStr(source.Columns.Count) & _
            " columns and the Excel Table '" & listObject.Name & "' has " & _
            CStr(colCount) & "; they must match."
    End If

    Set sheet = listObject.Parent
    oldRows = listObject.ListRows.Count
    firstBodyRow = listObject.Range.Row + 1
    firstColumn = listObject.Range.Column
    hadTotals = listObject.ShowTotals
    listObject.ShowTotals = False

    If liveRows.Count = 0 Then
        ' Resize refuses to leave a table with no data rows, so the body is
        ' deleted instead. The header row and the table itself survive.
        If Not listObject.DataBodyRange Is Nothing Then
            listObject.DataBodyRange.Delete
        End If
    Else
        listObject.Resize listObject.Range.Resize(1 + liveRows.Count, colCount)
        ReDim body(1 To liveRows.Count, 1 To colCount)
        For rowIndex = 1 To liveRows.Count
            ' The typed local matters. Collection.Item returns a Variant, so
            ' reading the Friend InternalDataRows straight off it would bind
            ' late and raise 438 at run time, the defect closed in #3.
            Set bodyRow = liveRows.Item(rowIndex)
            For colIndex = 1 To colCount
                body(rowIndex, colIndex) = RangeCellValue( _
                    bodyRow.InternalDataRows.Item(colIndex))
            Next colIndex
        Next rowIndex
        listObject.DataBodyRange.Value = body
    End If

    If oldRows > liveRows.Count Then
        sheet.Cells(firstBodyRow + liveRows.Count, firstColumn) _
            .Resize(oldRows - liveRows.Count, colCount).ClearContents
    End If
    listObject.ShowTotals = hadTotals
    WriteToListObject = liveRows.Count
End Function

Private Function RangeCellValue(ByVal wrapped As ROneCOne) As Variant
    If wrapped.InternalHasObjectValue Then
        RaiseContractError ERROR_TYPE_MISMATCH, "ToRange", _
            "ToRange cannot write an object-valued cell to a worksheet."
    End If
    RangeCellValue = wrapped.InternalScalarValue
End Function

Private Sub RequireRangeTarget( _
    ByVal target As Object, _
    ByVal memberName As String _
)
    If target Is Nothing Then
        RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
            "The worksheet range cannot be Nothing."
    End If
End Sub

Public Function CloneTable() As ROneCOne
Attribute CloneTable.VB_Description = "Copies a DataTable schema without rows."
    RequireDataTable Me, "CloneTable"
    Set CloneTable = CloneDataTableSchema(Me)
End Function

Public Function Copy( _
    Optional ByVal sourcePath As Variant, _
    Optional ByVal destinationPath As Variant, _
    Optional ByVal overwrite As Boolean = False _
) As ROneCOne
Attribute Copy.VB_Description = "Copies a DataTable with its rows, or copies a file."
    Dim result As ROneCOne
    Dim row As ROneCOne

    If mRole = ROLE_FILE_FACTORY Then
        FileCopyCore RequirePathText(sourcePath, "Copy"), _
            RequirePathText(destinationPath, "Copy"), overwrite
        Exit Function
    End If
    RequireDataTable Me, "Copy"
    If Not IsMissing(sourcePath) Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "Copy", _
            "Copy on a DataTable takes no arguments."
    End If
    Set result = CloneDataTableSchema(Me)
    For Each row In ArrSnapshot(mItems, mItemsCount)
        If row.RowState <> "Deleted" Then CopyDataRowInto row, result
    Next row
    result.AcceptChanges
    Set Copy = result
End Function

Public Sub Merge(ByVal sourceTable As ROneCOne)
Attribute Merge.VB_Description = "Merges compatible rows into this DataTable."
    Dim row As ROneCOne

    RequireDataTable Me, "Merge"
    RequireDataTable sourceTable, "Merge"
    ValidateCompatibleDataTable sourceTable, "Merge"
    For Each row In sourceTable.InternalDataRows
        If row.RowState <> "Deleted" Then CopyDataRowInto row, Me
    Next row
End Sub

Public Function GetChanges() As ROneCOne
Attribute GetChanges.VB_Description = "Copies rows that have pending changes."
    Dim result As ROneCOne
    Dim row As ROneCOne
    Dim table As ROneCOne

    If mRole = ROLE_DATA_SET Then
        Set result = ROneCOne.DataSet(mName)
        For Each table In mDataTables
            Set row = table.GetChanges
            If row.Rows.Count > 0 Then result.AddTable row
        Next table
        Set GetChanges = result
        Exit Function
    End If
    RequireDataTable Me, "GetChanges"
    Set result = CloneDataTableSchema(Me)
    For Each row In ArrSnapshot(mItems, mItemsCount)
        If row.RowState <> "Unchanged" Then CopyDataRowInto row, result
    Next row
    Set GetChanges = result
End Function

Public Sub Delete( _
    Optional ByVal targetPath As Variant, _
    Optional ByVal recursive As Boolean = False _
)
Attribute Delete.VB_Description = "Marks a DataRow deleted, or deletes a file or folder."
    If mRole = ROLE_FILE_FACTORY Then
        FileDeleteCore RequirePathText(targetPath, "Delete")
        Exit Sub
    End If
    If mRole = ROLE_DIRECTORY_FACTORY Then
        DirectoryDeleteCore _
            TrimTrailingSeparators(RequirePathText(targetPath, "Delete")), _
            recursive
        Exit Sub
    End If
    RequireDataRow Me, "Delete"
    If Not IsMissing(targetPath) Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "Delete", _
            "Delete on a DataRow takes no arguments."
    End If
    If mRowState = ROW_STATE_DETACHED Then
        RaiseContractError ERROR_INVALID_OPERATION, "Delete", _
            "A detached DataRow cannot be deleted."
    End If
    mRowState = ROW_STATE_DELETED
    mSource.InternalMarkDataIndexDirty
    mSource.TouchDataVersion
End Sub

Public Sub AddTable(ByVal table As ROneCOne)
Attribute AddTable.VB_Description = "Adds a DataTable to a DataSet."
    If mRole <> ROLE_DATA_SET Then
        RaiseContractError ERROR_INVALID_OPERATION, "AddTable", _
            "AddTable requires a DataSet."
    End If
    RequireDataTable table, "AddTable"
    If FindNamedObjectIndex(mDataTables, table.TableName) >= 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "AddTable", _
            "A table with the same name already exists."
    End If
    If Not table.InternalSource Is Nothing Then
        RaiseContractError ERROR_INVALID_OPERATION, "AddTable", _
            "The DataTable already belongs to a DataSet."
    End If
    table.ConfigureReadOnlySource Me
    mDataTables.Add table
    mVersion = mVersion + 1
End Sub

Public Sub AddRelation(ByVal relation As ROneCOne)
Attribute AddRelation.VB_Description = "Adds a DataRelation to a DataSet."
    If mRole <> ROLE_DATA_SET Then
        RaiseContractError ERROR_INVALID_OPERATION, "AddRelation", _
            "AddRelation requires a DataSet."
    End If
    RequireDataRelation relation, "AddRelation"
    If FindNamedObjectIndex(mDataRelations, relation.InternalName) >= 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "AddRelation", _
            "A relation with the same name already exists."
    End If
    ValidateRelationOwnership relation
    relation.ConfigureReadOnlySource Me
    mDataRelations.Add relation
    mVersion = mVersion + 1
End Sub

Public Property Get Columns() As ROneCOne
Attribute Columns.VB_Description = "Gets the named DataColumn collection."
    RequireDataTable Me, "Columns"
    ' Snapshot properties cache the built list against the structural version
    ' so a positional loop pays for one build, not one per access. Field-level
    ' edits advance only the data version and leave these caches valid.
    If mColumnsSnapshot Is Nothing Or mColumnsSnapshotVersion <> mVersion Then
        Set mColumnsSnapshot = CreateNamedDataCollection( _
            ArrSnapshot(mKeys, mKeysCount))
        mColumnsSnapshotVersion = mVersion
    End If
    Set Columns = mColumnsSnapshot
End Property

Public Property Get Rows() As ROneCOne
Attribute Rows.VB_Description = "Gets the DataRow collection."
    RequireDataTable Me, "Rows"
    If mRowsSnapshot Is Nothing Or mRowsSnapshotVersion <> mVersion Then
        Set mRowsSnapshot = CreateObjectListSnapshot( _
            ArrSnapshot(mItems, mItemsCount))
        mRowsSnapshotVersion = mVersion
    End If
    Set Rows = mRowsSnapshot
End Property

Public Property Get Tables() As ROneCOne
Attribute Tables.VB_Description = "Gets the named DataTable collection."
    If mRole <> ROLE_DATA_SET Then
        RaiseContractError ERROR_INVALID_OPERATION, "Tables", _
            "Tables requires a DataSet."
    End If
    If mTablesSnapshot Is Nothing Or mTablesSnapshotVersion <> mVersion Then
        Set mTablesSnapshot = CreateNamedDataCollection(mDataTables)
        mTablesSnapshotVersion = mVersion
    End If
    Set Tables = mTablesSnapshot
End Property

Public Property Get Relations() As ROneCOne
Attribute Relations.VB_Description = "Gets the named DataRelation collection."
    If mRole <> ROLE_DATA_SET Then
        RaiseContractError ERROR_INVALID_OPERATION, "Relations", _
            "Relations requires a DataSet."
    End If
    If mRelationsSnapshot Is Nothing Or _
        mRelationsSnapshotVersion <> mVersion Then
        Set mRelationsSnapshot = CreateNamedDataCollection(mDataRelations)
        mRelationsSnapshotVersion = mVersion
    End If
    Set Relations = mRelationsSnapshot
End Property

Public Property Get PrimaryKey() As ROneCOne
Attribute PrimaryKey.VB_Description = "Gets or sets the DataTable primary-key columns."
    RequireDataTable Me, "PrimaryKey"
    If mPrimaryKeySnapshot Is Nothing Or _
        mPrimaryKeySnapshotVersion <> mVersion Then
        Set mPrimaryKeySnapshot = CreateObjectListSnapshot(mPrimaryKey)
        mPrimaryKeySnapshotVersion = mVersion
    End If
    Set PrimaryKey = mPrimaryKeySnapshot
End Property

Public Property Set PrimaryKey(ByVal columns As ROneCOne)
    Dim column As ROneCOne

    RequireDataTable Me, "PrimaryKey"
    RequireSequence columns, "PrimaryKey"
    Set mPrimaryKey = New Collection
    For Each column In columns
        RequireDataColumn column, "PrimaryKey"
        If Not column.InternalSource Is Me Then
            RaiseContractError ERROR_INVALID_OPERATION, "PrimaryKey", _
                "Every primary-key column must belong to this table."
        End If
        mPrimaryKey.Add column
    Next column
    If mPrimaryKey.Count = 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "PrimaryKey", _
            "A primary key requires at least one column."
    End If
    For Each column In mPrimaryKey
        column.ConfigurePrimaryKeyColumn (mPrimaryKey.Count = 1)
    Next column
    RebuildPrimaryKeyIndex
    mVersion = mVersion + 1
End Property

Public Function Find(ParamArray keys() As Variant) As ROneCOne
Attribute Find.VB_Description = "Finds a DataRow by its single or composite primary key."
    Dim arguments As Collection
    Dim keyToken As String

    RequireDataTable Me, "Find"
    If mPrimaryKey.Count = 0 Then
        RaiseContractError ERROR_INVALID_OPERATION, "Find", _
            "The DataTable does not define a primary key."
    End If
    Set arguments = BuildPrimaryKeyArguments(keys)
    If arguments.Count <> mPrimaryKey.Count Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "Find", _
            "The key value count must match the primary-key column count."
    End If
    EnsureDataIndexesCurrent
    keyToken = PrimaryKeyTokenFromArguments(arguments)
    If mDataIndex.ContainsKey(keyToken) Then Set Find = mDataIndex.Item(keyToken)
End Function

Public Sub AcceptChanges()
Attribute AcceptChanges.VB_Description = "Commits current DataRow values as original values."
    Dim index As Long
    Dim row As ROneCOne

    If mRole = ROLE_DATA_ROW Then
        AcceptRowChanges
        Exit Sub
    End If
    If mRole = ROLE_DATA_TABLE Then
        For index = mItemsCount To 1 Step -1
            Set row = mItems(index)
            If row.RowState = "Deleted" Then
                ArrRemoveAt mItems, mItemsCount, index
            Else
                row.AcceptChanges
            End If
        Next index
        RebuildPrimaryKeyIndex
        mVersion = mVersion + 1
        Exit Sub
    End If
    If mRole = ROLE_DATA_SET Then
        For Each row In mDataTables
            row.AcceptChanges
        Next row
        Exit Sub
    End If
    RaiseContractError ERROR_INVALID_OPERATION, "AcceptChanges", _
        "AcceptChanges requires a DataRow, DataTable, or DataSet."
End Sub

Public Sub RejectChanges()
Attribute RejectChanges.VB_Description = "Rolls back DataRow values to their original state."
    Dim index As Long
    Dim row As ROneCOne

    If mRole = ROLE_DATA_ROW Then
        RejectRowChanges
        Exit Sub
    End If
    If mRole = ROLE_DATA_TABLE Then
        For index = mItemsCount To 1 Step -1
            Set row = mItems(index)
            If row.RowState = "Added" Then
                ArrRemoveAt mItems, mItemsCount, index
            Else
                row.RejectChanges
            End If
        Next index
        RebuildPrimaryKeyIndex
        mVersion = mVersion + 1
        Exit Sub
    End If
    If mRole = ROLE_DATA_SET Then
        For Each row In mDataTables
            row.RejectChanges
        Next row
        Exit Sub
    End If
    RaiseContractError ERROR_INVALID_OPERATION, "RejectChanges", _
        "RejectChanges requires a DataRow, DataTable, or DataSet."
End Sub

Public Property Get RowState() As String
Attribute RowState.VB_Description = "Gets the .NET-aligned DataRow state name."
    RequireDataRow Me, "RowState"
    Select Case mRowState
        Case ROW_STATE_DETACHED
            RowState = "Detached"
        Case ROW_STATE_UNCHANGED
            RowState = "Unchanged"
        Case ROW_STATE_ADDED
            RowState = "Added"
        Case ROW_STATE_DELETED
            RowState = "Deleted"
        Case ROW_STATE_MODIFIED
            RowState = "Modified"
    End Select
End Property

Public Function GetChildRows(ByVal relationName As String) As ROneCOne
Attribute GetChildRows.VB_Description = "Gets child rows through a DataRelation."
    RequireDataRow Me, "GetChildRows"
    Set GetChildRows = NavigateRelation(relationName, True)
End Function

Public Function GetParentRow(ByVal relationName As String) As ROneCOne
Attribute GetParentRow.VB_Description = "Gets the parent row through a DataRelation."
    Dim parents As ROneCOne

    RequireDataRow Me, "GetParentRow"
    Set parents = NavigateRelation(relationName, False)
    If parents.Count > 0 Then Set GetParentRow = parents.Item(0)
End Function

Public Function WithFilter(ByVal predicate As Variant) As ROneCOne
Attribute WithFilter.VB_Description = "Applies a typed filter to a DataView."
    If mRole <> ROLE_DATA_VIEW Then
        RaiseContractError ERROR_INVALID_OPERATION, "WithFilter", _
            "WithFilter requires a DataView."
    End If
    Set mQueryDelegate = NormalizeContextUnary(predicate, "WithFilter", False)
    mVersion = mVersion + 1
    Set WithFilter = Me
End Function

Public Function WithSort( _
    ByVal keySelector As Variant, _
    Optional ByVal descending As Boolean = False _
) As ROneCOne
Attribute WithSort.VB_Description = "Applies one typed sort key to a DataView."
    If mRole <> ROLE_DATA_VIEW Then
        RaiseContractError ERROR_INVALID_OPERATION, "WithSort", _
            "WithSort requires a DataView."
    End If
    Set mKeySelector = NormalizeContextUnary(keySelector, "WithSort", False)
    mExpressionOption = Abs(CLng(descending))
    mVersion = mVersion + 1
    Set WithSort = Me
End Function

Public Function ToTable( _
    Optional ByVal tableName As String = vbNullString _
) As ROneCOne
Attribute ToTable.VB_Description = "Materializes a DataView as an independent DataTable."
    Dim result As ROneCOne
    Dim row As Variant

    If mRole <> ROLE_DATA_VIEW Then
        RaiseContractError ERROR_INVALID_OPERATION, "ToTable", _
            "ToTable requires a DataView."
    End If
    Set result = CloneDataTableSchema(mSource)
    If Len(tableName) > 0 Then result.ConfigureDataName tableName
    For Each row In Me
        CopyDataRowInto row, result
    Next row
    result.AcceptChanges
    Set ToTable = result
End Function

Public Property Get Json() As ROneCOne
Attribute Json.VB_Description = "Gets the System.Text.Json-style serializer surface."
    Dim factory As ROneCOne

    Set factory = New ROneCOne
    factory.ConfigureJsonFactory
    Set Json = factory
End Property

Public Function Serialize( _
    ByVal value As Variant, _
    Optional ByVal writeIndented As Boolean = False _
) As String
Attribute Serialize.VB_Description = "Serializes a value to JSON text, or a table to CSV text."
    Dim builder As JsonTextBuilder

    If mRole = ROLE_CSV_FACTORY Then
        If writeIndented Then
            RaiseContractError ERROR_INVALID_ARGUMENT, "Serialize", _
                "CSV has no indented form."
        End If
        Serialize = CsvSerializeValue(value)
        Exit Function
    End If
    RequireJsonFactory "Serialize"
    JsonBuilderInit builder, 512
    If writeIndented Then
        JsonWriteValuePretty builder, value, 0
    Else
        JsonWriteValue builder, value
    End If
    Serialize = JsonBuilderText(builder)
End Function

Public Function Deserialize(ByVal jsonText As String) As Variant
Attribute Deserialize.VB_Description = "Parses JSON text into dictionaries, lists, and scalars."
    Dim reader As JsonReader

    RequireJsonFactory "Deserialize"
    JsonReaderInit reader, jsonText
    AssignVariant Deserialize, JsonReadValue(reader)
    JsonSkipWhitespace reader
    If reader.pos <= reader.textLen Then
        RaiseJsonError "Unexpected trailing characters at position " & _
            CStr(reader.pos)
    End If
End Function

Public Function DeserializeOnly( _
    ByVal jsonText As String, _
    ByVal paths As Variant _
) As Variant
Attribute DeserializeOnly.VB_Description = "Parses only the named paths, skipping everything else."
    Dim entry As Variant
    Dim reader As JsonReader
    Dim stepSets As Collection

    ' A fat API response can be almost entirely members you will never read.
    ' This walks the document once and materializes only what you asked for,
    ' stepping over the rest, and the result is a true subset of the source
    ' so navigation is identical to Deserialize.
    RequireJsonFactory "DeserializeOnly"
    Set stepSets = New Collection
    If IsArrayValue(paths) Then
        For Each entry In paths
            stepSets.Add JsonSplitPath(CStr(entry))
        Next entry
    Else
        stepSets.Add JsonSplitPath(CStr(paths))
    End If
    If stepSets.Count = 0 Then
        RaiseJsonError "DeserializeOnly requires at least one path."
    End If
    JsonReaderInit reader, jsonText
    AssignVariant DeserializeOnly, JsonReadFiltered(reader, stepSets)
    JsonSkipWhitespace reader
    If reader.pos <= reader.textLen Then
        RaiseJsonError "Unexpected trailing characters at position " & _
            CStr(reader.pos)
    End If
End Function

Public Function DeserializeAt( _
    ByVal jsonText As String, _
    ByVal path As String _
) As Variant
Attribute DeserializeAt.VB_Description = "Parses just the value one path addresses."
    Dim reader As JsonReader

    ' Unlike DeserializeOnly, which omits a missing path silently because an
    ' allowlist over a varying API should, this raises: asking for one value
    ' and getting nothing back is a question that deserves an answer.
    RequireJsonFactory "DeserializeAt"
    JsonReaderInit reader, jsonText
    AssignVariant DeserializeAt, _
        JsonReadAtPath(reader, JsonSplitPath(path), path)
End Function

Public Function DeserializeTable( _
    ByVal text As String, _
    Optional ByVal tableName As String = vbNullString, _
    Optional ByVal arrayPath As String = "$" _
) As ROneCOne
Attribute DeserializeTable.VB_Description = "Parses JSON, CSV, or XML text into a DataTable."
    If mRole = ROLE_XML_FACTORY Then
        If Len(tableName) = 0 Then tableName = "Xml"
        ' The path selects one node per row; the default takes every child
        ' element of the document element.
        If arrayPath = "$" Then arrayPath = "/*/*"
        Set DeserializeTable = XmlDeserializeTable(text, tableName, arrayPath)
        Exit Function
    End If
    If mRole = ROLE_CSV_FACTORY Then
        If arrayPath <> "$" Then
            RaiseContractError ERROR_INVALID_ARGUMENT, "DeserializeTable", _
                "CSV has no array path."
        End If
        If Len(tableName) = 0 Then tableName = "Csv"
        Set DeserializeTable = CsvParseTable(text, tableName)
        Exit Function
    End If
    RequireJsonFactory "DeserializeTable"
    If Len(tableName) = 0 Then tableName = "Json"
    ' arrayPath selects during the scan, so an envelope's other members are
    ' skipped rather than materialized and then discarded.
    Set DeserializeTable = BuildTableFromJsonArray( _
        ResolveJsonArray(DeserializeAt(text, arrayPath), "$"), tableName)
End Function

Public Function DeserializeInto( _
    ByVal jsonText As String, _
    ByVal instance As Object _
) As Object
Attribute DeserializeInto.VB_Description = "Binds JSON members onto an existing instance."
    Dim tree As Variant

    RequireJsonFactory "DeserializeInto"
    If instance Is Nothing Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "DeserializeInto", _
            "The target instance cannot be Nothing."
    End If
    AssignVariant tree, Deserialize(jsonText)
    RequireJsonObjectValue tree, "DeserializeInto"
    BindJsonObjectInto instance, tree
    Set DeserializeInto = instance
End Function

Public Function DeserializeObjects( _
    ByVal jsonText As String, _
    ByVal factory As ROneCOne, _
    Optional ByVal arrayPath As String = "$" _
) As ROneCOne
Attribute DeserializeObjects.VB_Description = "Parses a JSON array into factory-built objects."
    Dim element As Variant
    Dim instance As Object
    Dim result As ROneCOne
    Dim rows As ROneCOne

    RequireJsonFactory "DeserializeObjects"
    ValidateJsonObjectFactory factory, "DeserializeObjects"
    ' The path now selects during the scan instead of after a full parse, so
    ' an envelope's unrelated members are stepped over rather than built.
    Set rows = ResolveJsonArray(DeserializeAt(jsonText, arrayPath), "$")
    Set result = Nothing
    For Each element In rows
        RequireJsonObjectValue element, "DeserializeObjects"
        Set instance = InvokeJsonObjectFactory(factory, "DeserializeObjects")
        BindJsonObjectInto instance, element
        If result Is Nothing Then Set result = ROneCOne.ListLike(instance)
        result.Add instance
    Next element
    If result Is Nothing Then
        Set result = New ROneCOne
        result.ConfigureList vbObject, "Object"
    End If
    Set DeserializeObjects = result
End Function

Public Function DataTableFromObjects( _
    ByVal source As Variant, _
    ByVal propertyNames As Variant, _
    Optional ByVal tableName As String = "Objects" _
) As ROneCOne
Attribute DataTableFromObjects.VB_Description = "Builds a DataTable from typed objects."
    Set DataTableFromObjects = BuildTableFromObjects(source, _
        propertyNames, tableName)
End Function

Public Function ToObjects(ByVal factory As ROneCOne) As ROneCOne
Attribute ToObjects.VB_Description = "Maps DataTable rows onto factory-built typed objects."
    Dim cells As Collection
    Dim columnIndex As Long
    Dim instance As Object
    Dim result As ROneCOne
    Dim rowIndex As Long
    Dim wrapped As ROneCOne

    RequireDataTable Me, "ToObjects"
    ValidateJsonObjectFactory factory, "ToObjects"
    Set result = Nothing
    For rowIndex = 1 To mItemsCount
        If mItems(rowIndex).RowState <> "Deleted" Then
            Set instance = InvokeJsonObjectFactory(factory, "ToObjects")
            Set cells = mItems(rowIndex).InternalDataRows
            For columnIndex = 1 To mKeysCount
                Set wrapped = cells.Item(columnIndex)
                If Not IsNullDataValue(wrapped) Then
                    AssignInstanceProperty instance, _
                        mKeys(columnIndex).ColumnName, wrapped
                End If
            Next columnIndex
            If result Is Nothing Then Set result = ROneCOne.ListLike(instance)
            result.Add instance
        End If
    Next rowIndex
    If result Is Nothing Then
        Set result = New ROneCOne
        result.ConfigureList vbObject, "Object"
    End If
    Set ToObjects = result
End Function

Public Function ToJson( _
    Optional ByVal writeIndented As Boolean = False _
) As String
Attribute ToJson.VB_Description = "Serializes this collection, table, or row to JSON text."
    ToJson = ROneCOne.Json.Serialize(Me, writeIndented)
End Function

Public Function ToCsv() As String
Attribute ToCsv.VB_Description = "Serializes this DataTable or DataView to CSV text."
    ToCsv = ROneCOne.Csv.Serialize(Me)
End Function

Public Property Get File() As ROneCOne
Attribute File.VB_Description = "Gets the System.IO.File-style file surface."
    Dim factory As ROneCOne

    Set factory = New ROneCOne
    factory.ConfigureFileFactory
    Set File = factory
End Property

Public Property Get Directory() As ROneCOne
Attribute Directory.VB_Description = "Gets the System.IO.Directory-style folder surface."
    Dim factory As ROneCOne

    Set factory = New ROneCOne
    factory.ConfigureDirectoryFactory
    Set Directory = factory
End Property

Public Property Get Path() As ROneCOne
Attribute Path.VB_Description = "Gets the System.IO.Path-style path helper surface."
    Dim factory As ROneCOne

    Set factory = New ROneCOne
    factory.ConfigurePathFactory
    Set Path = factory
End Property

Public Property Get Csv() As ROneCOne
Attribute Csv.VB_Description = "Gets the CSV serializer surface for DataTables and views."
    Dim factory As ROneCOne

    Set factory = New ROneCOne
    factory.ConfigureCsvFactory
    Set Csv = factory
End Property

Public Function ReadAllText( _
    Optional ByVal filePath As String = vbNullString, _
    Optional ByVal encoding As String = "utf-8" _
) As String
Attribute ReadAllText.VB_Description = "Reads a file or a zip entry as decoded text."
    Dim dataBytes() As Byte

    If mRole = ROLE_ZIP_ENTRY Then
        If Len(filePath) > 0 Then
            RaiseContractError ERROR_INVALID_ARGUMENT, "ReadAllText", _
                "A zip entry reads itself; no path is accepted."
        End If
        dataBytes = ZipEntryBytes("ReadAllText")
        ReadAllText = ZipEntryText(dataBytes, _
            NormalizeTextEncoding(encoding, "ReadAllText"))
        Exit Function
    End If
    RequireFileFactory "ReadAllText"
    If Len(filePath) = 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "ReadAllText", _
            "ReadAllText requires a file path."
    End If
    ReadAllText = IoReadTextFile(filePath, _
        NormalizeTextEncoding(encoding, "ReadAllText"), "ReadAllText")
End Function

Public Sub WriteAllText( _
    ByVal filePath As String, _
    ByVal contents As String, _
    Optional ByVal encoding As String = "utf-8" _
)
Attribute WriteAllText.VB_Description = "Writes a text file; UTF-8 output has no byte-order mark."
    RequireFileFactory "WriteAllText"
    IoWriteTextFile filePath, contents, _
        NormalizeTextEncoding(encoding, "WriteAllText"), False, "WriteAllText"
End Sub

Public Sub AppendAllText( _
    ByVal filePath As String, _
    ByVal contents As String, _
    Optional ByVal encoding As String = "utf-8" _
)
Attribute AppendAllText.VB_Description = "Appends encoded text, creating the file when missing."
    RequireFileFactory "AppendAllText"
    IoWriteTextFile filePath, contents, _
        NormalizeTextEncoding(encoding, "AppendAllText"), True, "AppendAllText"
End Sub

Public Function ReadAllLines( _
    ByVal filePath As String, _
    Optional ByVal encoding As String = "utf-8" _
) As ROneCOne
Attribute ReadAllLines.VB_Description = "Reads a text file into a List of String lines."
    Dim line As Variant
    Dim lines As ROneCOne
    Dim normalized As String

    RequireFileFactory "ReadAllLines"
    Set lines = ROneCOne.ListOf(vbString)
    normalized = IoReadTextFile(filePath, _
        NormalizeTextEncoding(encoding, "ReadAllLines"), "ReadAllLines")
    If Len(normalized) = 0 Then
        Set ReadAllLines = lines
        Exit Function
    End If
    normalized = VBA.Replace$(normalized, vbCrLf, vbLf)
    normalized = VBA.Replace$(normalized, vbCr, vbLf)
    If Right$(normalized, 1) = vbLf Then
        normalized = Left$(normalized, Len(normalized) - 1)
    End If
    If Len(normalized) = 0 Then
        lines.Add vbNullString
    Else
        For Each line In VBA.Split(normalized, vbLf)
            lines.Add CStr(line)
        Next line
    End If
    Set ReadAllLines = lines
End Function

Public Sub WriteAllLines( _
    ByVal filePath As String, _
    ByVal lines As Variant, _
    Optional ByVal encoding As String = "utf-8" _
)
Attribute WriteAllLines.VB_Description = "Writes lines, terminating every one including the last."
    Dim contents As String
    Dim element As Variant
    Dim elements As Collection
    Dim index As Long
    Dim parts() As String

    RequireFileFactory "WriteAllLines"
    Set elements = CollectSourceElements(lines, "WriteAllLines")
    If elements.Count > 0 Then
        ReDim parts(0 To elements.Count - 1)
        For Each element In elements
            parts(index) = CStr(element)
            index = index + 1
        Next element
        contents = VBA.Join(parts, vbCrLf) & vbCrLf
    End If
    IoWriteTextFile filePath, contents, _
        NormalizeTextEncoding(encoding, "WriteAllLines"), False, "WriteAllLines"
End Sub

Public Function ReadAllBytes( _
    Optional ByVal filePath As String = vbNullString _
) As Variant
Attribute ReadAllBytes.VB_Description = "Reads a file or a zip entry as a Byte array."
    If mRole = ROLE_ZIP_ENTRY Then
        If Len(filePath) > 0 Then
            RaiseContractError ERROR_INVALID_ARGUMENT, "ReadAllBytes", _
                "A zip entry reads itself; no path is accepted."
        End If
        ReadAllBytes = ZipEntryBytes("ReadAllBytes")
        Exit Function
    End If
    RequireFileFactory "ReadAllBytes"
    If Len(filePath) = 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "ReadAllBytes", _
            "ReadAllBytes requires a file path."
    End If
    ReadAllBytes = IoReadAllBytesCore(filePath, "ReadAllBytes")
End Function

Public Sub WriteAllBytes(ByVal filePath As String, ByVal bytes As Variant)
Attribute WriteAllBytes.VB_Description = "Writes a Byte array as a whole file."
    RequireFileFactory "WriteAllBytes"
    If VarTypeOf(bytes) <> (vbArray Or vbByte) Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "WriteAllBytes", _
            "WriteAllBytes requires a Byte array."
    End If
    IoWriteAllBytesCore filePath, bytes, "WriteAllBytes"
End Sub

Public Sub Move(ByVal sourcePath As String, ByVal destinationPath As String)
Attribute Move.VB_Description = "Moves a file; an existing destination is an error."
    RequireFileFactory "Move"
    FileMoveCore sourcePath, destinationPath
End Sub

Public Sub CreateDirectory(ByVal directoryPath As String)
Attribute CreateDirectory.VB_Description = "Creates a folder, including missing parents."
    RequireDirectoryFactory "CreateDirectory"
    DirectoryCreateCore TrimTrailingSeparators(directoryPath), _
        "CreateDirectory"
End Sub

Public Function GetFiles( _
    ByVal directoryPath As String, _
    Optional ByVal searchPattern As String = "*", _
    Optional ByVal allDirectories As Boolean = False _
) As ROneCOne
Attribute GetFiles.VB_Description = "Lists matching file paths, sorted, optionally recursive."
    RequireDirectoryFactory "GetFiles"
    Set GetFiles = DirectoryEntriesCore(directoryPath, searchPattern, _
        allDirectories, True, "GetFiles")
End Function

Public Function GetDirectories( _
    ByVal directoryPath As String, _
    Optional ByVal searchPattern As String = "*", _
    Optional ByVal allDirectories As Boolean = False _
) As ROneCOne
Attribute GetDirectories.VB_Description = "Lists matching folder paths, sorted, maybe recursive."
    RequireDirectoryFactory "GetDirectories"
    Set GetDirectories = DirectoryEntriesCore(directoryPath, searchPattern, _
        allDirectories, False, "GetDirectories")
End Function

Public Function GetFileName(ByVal pathText As String) As String
Attribute GetFileName.VB_Description = "Gets the file name and extension of a path."
    RequirePathFactory "GetFileName"
    GetFileName = Mid$(pathText, PathLastSeparator(pathText) + 1)
End Function

Public Function GetDirectoryName(ByVal pathText As String) As String
Attribute GetDirectoryName.VB_Description = "Gets the directory portion of a path."
    Dim separatorAt As Long

    RequirePathFactory "GetDirectoryName"
    separatorAt = PathLastSeparator(pathText)
    If separatorAt = 0 Then Exit Function
    GetDirectoryName = TrimTrailingSeparators(Left$(pathText, separatorAt))
    If Len(GetDirectoryName) = 2 And Right$(GetDirectoryName, 1) = ":" Then
        GetDirectoryName = GetDirectoryName & "\"
    ElseIf Len(GetDirectoryName) = 0 Then
        GetDirectoryName = Left$(pathText, 1)
    End If
End Function

Public Function GetExtension(ByVal pathText As String) As String
Attribute GetExtension.VB_Description = "Gets the extension of a path, including its dot."
    Dim dotAt As Long
    Dim nameText As String

    RequirePathFactory "GetExtension"
    nameText = Mid$(pathText, PathLastSeparator(pathText) + 1)
    dotAt = InStrRev(nameText, ".")
    If dotAt = 0 Or dotAt = Len(nameText) Then Exit Function
    GetExtension = Mid$(nameText, dotAt)
End Function

Public Function GetFileNameWithoutExtension(ByVal pathText As String) As String
Attribute GetFileNameWithoutExtension.VB_Description = "Gets the file name minus its extension."
    Dim dotAt As Long
    Dim nameText As String

    RequirePathFactory "GetFileNameWithoutExtension"
    nameText = Mid$(pathText, PathLastSeparator(pathText) + 1)
    dotAt = InStrRev(nameText, ".")
    If dotAt > 0 Then nameText = Left$(nameText, dotAt - 1)
    GetFileNameWithoutExtension = nameText
End Function

Public Function ChangeExtension( _
    ByVal pathText As String, _
    ByVal extension As String _
) As String
Attribute ChangeExtension.VB_Description = "Replaces a path's extension; empty text removes it."
    Dim dotAt As Long
    Dim nameText As String
    Dim trimmed As String

    RequirePathFactory "ChangeExtension"
    nameText = Mid$(pathText, PathLastSeparator(pathText) + 1)
    dotAt = InStrRev(nameText, ".")
    If dotAt > 0 Then
        trimmed = Left$(pathText, _
            Len(pathText) - (Len(nameText) - dotAt + 1))
    Else
        trimmed = pathText
    End If
    If Len(extension) = 0 Then
        ChangeExtension = trimmed
    ElseIf Left$(extension, 1) = "." Then
        ChangeExtension = trimmed & extension
    Else
        ChangeExtension = trimmed & "." & extension
    End If
End Function

Public Function GetFullPath(ByVal pathText As String) As String
Attribute GetFullPath.VB_Description = "Resolves a path against the current directory."
    RequirePathFactory "GetFullPath"
    GetFullPath = CreateFileSystemObject().GetAbsolutePathName(pathText)
End Function

Public Function GetTempPath() As String
Attribute GetTempPath.VB_Description = "Gets the temporary folder path with a trailing separator."
    RequirePathFactory "GetTempPath"
    GetTempPath = CreateFileSystemObject().GetSpecialFolder(2).Path & "\"
End Function

Public Property Get Process() As ROneCOne
Attribute Process.VB_Description = "Gets the awaitable process execution surface."
    Dim factory As ROneCOne

    Set factory = New ROneCOne
    factory.ConfigureProcessFactory
    Set Process = factory
End Property

Public Function RunAsync( _
    ByVal command As String, _
    Optional ByVal workingDirectory As String = vbNullString, _
    Optional ByVal cancellationToken As ROneCOne, _
    Optional ByVal standardInput As String = vbNullString _
) As ROneCOne
Attribute RunAsync.VB_Description = "Starts a shell command; the Task result carries its output."
    Dim errorPath As String
    Dim execObject As Object
    Dim fso As Object
    Dim inner As String
    Dim outputPath As String
    Dim pending As ROneCOne
    Dim taskValue As ROneCOne

    RequireProcessFactory "RunAsync"
    ValidateCancellationToken cancellationToken, "RunAsync"
    If Len(Trim$(command)) = 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "RunAsync", _
            "The command cannot be empty."
    End If
    ' The command runs through cmd.exe with both streams redirected to
    ' scratch files, so polling never touches the blocking pipe readers;
    ' the process starts now and the Task polls Status cooperatively.
    Set fso = CreateFileSystemObject()
    outputPath = fso.GetSpecialFolder(2).Path & "\" & fso.GetTempName
    errorPath = fso.GetSpecialFolder(2).Path & "\" & fso.GetTempName
    inner = "(" & command & ")"
    If Len(workingDirectory) > 0 Then
        inner = "cd /d """ & workingDirectory & """ && " & inner
    End If
    On Error GoTo StartFailed
    Set execObject = CreateObject(WSHELL_PROG_ID).Exec( _
        "cmd.exe /d /s /c """ & inner & " 1>""" & outputPath & _
        """ 2>""" & errorPath & """""")
    On Error GoTo 0
    ' Standard input is written now and always closed, so a filter such as
    ' sort sees end-of-file instead of waiting forever. A command that has
    ' already exited simply leaves the input undelivered.
    On Error Resume Next
    If Len(standardInput) > 0 Then execObject.StdIn.Write standardInput
    execObject.StdIn.Close
    On Error GoTo 0
    Set pending = New ROneCOne
    pending.ConfigureProcessResult execObject, outputPath, errorPath
    Set taskValue = New ROneCOne
    taskValue.ConfigureProcessTask pending, cancellationToken
    Set RunAsync = taskValue
    Exit Function

StartFailed:
    Err.Raise ERROR_PROCESS, "ROneCOne.ProcessException", _
        "The process could not start: " & Err.Description
End Function

Public Function StartSession( _
    ByVal command As String, _
    Optional ByVal workingDirectory As String = vbNullString, _
    Optional ByVal encodingName As String = vbNullString _
) As ROneCOne
Attribute StartSession.VB_Description = "Starts a command you can write to and read from."
    Dim session As ROneCOne

    RequireProcessFactory "StartSession"
    If Len(Trim$(command)) = 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "StartSession", _
            "The command cannot be empty."
    End If
    Set session = New ROneCOne
    session.ConfigureProcessSession command, workingDirectory, encodingName
    Set StartSession = session
End Function

Public Function WriteAsync( _
    ByVal text As String, _
    Optional ByVal cancellationToken As ROneCOne _
) As ROneCOne
Attribute WriteAsync.VB_Description = "Queues text to the session's standard input."
    Set WriteAsync = SessionQueueWrite(text, cancellationToken, "WriteAsync")
End Function

Public Function WriteLineAsync( _
    ByVal text As String, _
    Optional ByVal cancellationToken As ROneCOne _
) As ROneCOne
Attribute WriteLineAsync.VB_Description = "Queues one line to the session's standard input."
    Set WriteLineAsync = SessionQueueWrite( _
        text & vbCrLf, cancellationToken, "WriteLineAsync")
End Function

Public Sub CloseInput()
Attribute CloseInput.VB_Description = "Closes standard input so the command sees end of file."
    RequireProcessSession "CloseInput"
    If mSessionWriteQueue.Count > 0 Or mSessionWriteActive Then
        RaiseProcessError "Await every queued write before closing input."
    End If
    SessionCloseInputHandle
End Sub

Public Function ReadLineAsync( _
    Optional ByVal timeoutMilliseconds As Long = -1, _
    Optional ByVal cancellationToken As ROneCOne _
) As ROneCOne
Attribute ReadLineAsync.VB_Description = "Awaits the next standard output line."
    Set ReadLineAsync = SessionCreateTask( _
        "ReadOut", cancellationToken, "ReadLineAsync", timeoutMilliseconds)
End Function

Public Function ReadErrorLineAsync( _
    Optional ByVal timeoutMilliseconds As Long = -1, _
    Optional ByVal cancellationToken As ROneCOne _
) As ROneCOne
Attribute ReadErrorLineAsync.VB_Description = "Awaits the next standard error line."
    Set ReadErrorLineAsync = SessionCreateTask( _
        "ReadErr", cancellationToken, "ReadErrorLineAsync", _
        timeoutMilliseconds)
End Function

Public Function ReadAvailable() As String
Attribute ReadAvailable.VB_Description = "Takes buffered standard output without waiting."
    RequireProcessSession "ReadAvailable"
    ReadAvailable = InternalSessionRemainingText(False)
End Function

Public Function ReadErrorAvailable() As String
Attribute ReadErrorAvailable.VB_Description = "Takes buffered standard error without waiting."
    RequireProcessSession "ReadErrorAvailable"
    ReadErrorAvailable = InternalSessionRemainingText(True)
End Function

Public Function ReadToEndAsync( _
    Optional ByVal cancellationToken As ROneCOne _
) As ROneCOne
Attribute ReadToEndAsync.VB_Description = "Awaits exit and resolves to the rest of standard output."
    Set ReadToEndAsync = SessionCreateTask( _
        "ReadToEnd", cancellationToken, "ReadToEndAsync")
End Function

Public Function WaitForExitAsync( _
    Optional ByVal cancellationToken As ROneCOne _
) As ROneCOne
Attribute WaitForExitAsync.VB_Description = "Awaits exit and resolves to the exit code."
    Set WaitForExitAsync = SessionCreateTask( _
        "Exit", cancellationToken, "WaitForExitAsync")
End Function

' VBA reserves Kill for its file-deletion statement, so the legal member
' name here is KillProcess; System.Diagnostics.Process.Kill is the model.
Public Sub KillProcess()
Attribute KillProcess.VB_Description = "Terminates the session's command immediately."
    RequireProcessSession "KillProcess"
    If mSessionExited Then Exit Sub
    TerminateProcess mSessionProcess, 1
    SessionRefreshExit
End Sub

Public Property Get HasExited() As Boolean
Attribute HasExited.VB_Description = "Tests whether the session's command has finished."
    RequireProcessSession "HasExited"
    SessionRefreshExit
    HasExited = mSessionExited
End Property

Public Property Get ProcessId() As Long
Attribute ProcessId.VB_Description = "Gets the session's operating system process identifier."
    RequireProcessSession "ProcessId"
    ProcessId = mSessionProcessId
End Property

Public Property Get ExitCode() As Long
Attribute ExitCode.VB_Description = "Gets a completed process's exit code."
    If mRole = ROLE_PROCESS_SESSION Then
        SessionRefreshExit
        If Not mSessionExited Then
            RaiseProcessError "The command is still running; await " & _
                "WaitForExitAsync before reading ExitCode."
        End If
        ExitCode = mSessionExitCode
        Exit Property
    End If
    RequireProcessResult "ExitCode"
    ExitCode = mProcessExitCode
End Property

Public Property Get StandardOutput() As String
Attribute StandardOutput.VB_Description = "Gets a completed process's captured standard output."
    RequireProcessResult "StandardOutput"
    StandardOutput = mProcessOutput
End Property

Public Property Get StandardError() As String
Attribute StandardError.VB_Description = "Gets a completed process's captured standard error."
    RequireProcessResult "StandardError"
    StandardError = mProcessError
End Property

Public Function Regex( _
    ByVal pattern As String, _
    Optional ByVal ignoreCase As Boolean = False, _
    Optional ByVal multiLine As Boolean = False _
) As ROneCOne
Attribute Regex.VB_Description = "Compiles a regular expression with optional case and line modes."
    Dim expression As ROneCOne

    Set expression = New ROneCOne
    expression.ConfigureRegex pattern, ignoreCase, multiLine
    Set Regex = expression
End Function

Public Property Get Pattern() As String
Attribute Pattern.VB_Description = "Gets the source pattern of a regular expression."
    RequireRegex "Pattern"
    Pattern = mName
End Property

Public Function IsMatch(ByVal inputText As String) As Boolean
Attribute IsMatch.VB_Description = "Tests whether the pattern matches anywhere in the input."
    RequireRegex "IsMatch"
    IsMatch = mProviderObject.Test(inputText)
End Function

Public Function Matches(ByVal inputText As String) As ROneCOne
Attribute Matches.VB_Description = "Finds every match as a queryable typed list."
    Dim rawMatch As Object
    Dim results As ROneCOne

    RequireRegex "Matches"
    Set results = ROneCOne.ListOf(Me)
    mProviderObject.Global = True
    For Each rawMatch In mProviderObject.Execute(inputText)
        results.Add BuildRegexMatch(rawMatch)
    Next rawMatch
    Set Matches = results
End Function

Public Function Replace( _
    ByVal inputText As String, _
    ByVal replacement As String _
) As String
Attribute Replace.VB_Description = "Replaces every match, expanding $1-style group references."
    RequireRegex "Replace"
    mProviderObject.Global = True
    Replace = mProviderObject.Replace(inputText, replacement)
End Function

Public Function Split(ByVal inputText As String) As ROneCOne
Attribute Split.VB_Description = "Splits the input at every non-empty pattern match."
    Dim cursor As Long
    Dim pieces As ROneCOne
    Dim rawMatch As Object

    RequireRegex "Split"
    Set pieces = ROneCOne.ListOf(vbString)
    cursor = 1
    mProviderObject.Global = True
    For Each rawMatch In mProviderObject.Execute(inputText)
        ' A zero-length match separates nothing; splitting there would
        ' manufacture empty pieces between every character.
        If rawMatch.Length > 0 Then
            pieces.Add Mid$(inputText, cursor, _
                rawMatch.FirstIndex + 1 - cursor)
            cursor = rawMatch.FirstIndex + rawMatch.Length + 1
        End If
    Next rawMatch
    pieces.Add Mid$(inputText, cursor)
    Set Split = pieces
End Function

Public Property Get Success() As Boolean
Attribute Success.VB_Description = "Reports whether a regular expression match succeeded."
    RequireRegexMatch "Success"
    Success = mRegexSuccess
End Property

Public Property Get FirstIndex() As Long
Attribute FirstIndex.VB_Description = "Gets a match's zero-based position in the input."
    RequireRegexMatch "FirstIndex"
    FirstIndex = mRegexIndex
End Property

Public Property Get Length() As Long
Attribute Length.VB_Description = "Gets a match, builder, or zip entry length."
    Select Case mRole
        Case ROLE_REGEX_MATCH
            Length = mRegexLength
        Case ROLE_STRING_BUILDER
            Length = mStringBuilder.used
        Case ROLE_ZIP_ENTRY
            Length = mZipEntrySize
        Case Else
            RaiseContractError ERROR_INVALID_OPERATION, "Length", _
                "Length requires a match, StringBuilder, or zip entry."
    End Select
End Property

Public Property Get Groups() As ROneCOne
Attribute Groups.VB_Description = "Gets the match text and captured groups as a String list."
    RequireRegexMatch "Groups"
    Set Groups = mSource
End Property

Public Property Get Hash() As ROneCOne
Attribute Hash.VB_Description = "Gets the Windows CNG hashing surface."
    Dim factory As ROneCOne

    Set factory = New ROneCOne
    factory.ConfigureHashFactory
    Set Hash = factory
End Property

Public Function Sha256(ByVal value As Variant) As Variant
Attribute Sha256.VB_Description = "Computes the SHA-256 digest of text or bytes."
    RequireHashFactory "Sha256"
    Sha256 = HashValueCore("SHA256", value, "Sha256")
End Function

Public Function Sha512(ByVal value As Variant) As Variant
Attribute Sha512.VB_Description = "Computes the SHA-512 digest of text or bytes."
    RequireHashFactory "Sha512"
    Sha512 = HashValueCore("SHA512", value, "Sha512")
End Function

Public Function Sha1(ByVal value As Variant) As Variant
Attribute Sha1.VB_Description = "Computes the SHA-1 digest of text or bytes."
    RequireHashFactory "Sha1"
    Sha1 = HashValueCore("SHA1", value, "Sha1")
End Function

Public Function Md5(ByVal value As Variant) As Variant
Attribute Md5.VB_Description = "Computes the MD5 digest of text or bytes."
    RequireHashFactory "Md5"
    Md5 = HashValueCore("MD5", value, "Md5")
End Function

Public Function HmacSha256( _
    ByVal key As Variant, _
    ByVal value As Variant _
) As Variant
Attribute HmacSha256.VB_Description = "Computes an HMAC-SHA256 signature of text or bytes."
    Dim dataBytes() As Byte
    Dim dataLength As Long
    Dim keyBytes() As Byte
    Dim keyLength As Long

    RequireHashFactory "HmacSha256"
    CoerceHashInput key, keyBytes, keyLength, "HmacSha256"
    CoerceHashInput value, dataBytes, dataLength, "HmacSha256"
    HmacSha256 = BcryptHashBytes("SHA256", True, keyBytes, keyLength, _
        dataBytes, dataLength, "HmacSha256")
End Function

Public Property Get Convert() As ROneCOne
Attribute Convert.VB_Description = "Gets the System.Convert-style encoding helper surface."
    Dim factory As ROneCOne

    Set factory = New ROneCOne
    factory.ConfigureConvertFactory
    Set Convert = factory
End Property

Public Function ToBase64String(ByVal bytes As Variant) As String
Attribute ToBase64String.VB_Description = "Encodes a Byte array as base64 text."
    RequireConvertFactory "ToBase64String"
    If VarTypeOf(bytes) <> (vbArray Or vbByte) Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "ToBase64String", _
            "ToBase64String requires a Byte array."
    End If
    ToBase64String = Base64EncodeBytes(bytes)
End Function

Public Function FromBase64String(ByVal text As String) As Variant
Attribute FromBase64String.VB_Description = "Decodes base64 text into a Byte array."
    RequireConvertFactory "FromBase64String"
    FromBase64String = Base64DecodeText(text)
End Function

Public Function ToHexString(ByVal bytes As Variant) As String
Attribute ToHexString.VB_Description = "Encodes a Byte array as uppercase hexadecimal text."
    Dim index As Long
    Dim parts() As String

    RequireConvertFactory "ToHexString"
    If VarTypeOf(bytes) <> (vbArray Or vbByte) Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "ToHexString", _
            "ToHexString requires a Byte array."
    End If
    If UBound(bytes) < LBound(bytes) Then Exit Function
    ReDim parts(LBound(bytes) To UBound(bytes))
    For index = LBound(bytes) To UBound(bytes)
        parts(index) = Right$("0" & Hex$(bytes(index)), 2)
    Next index
    ToHexString = VBA.Join(parts, vbNullString)
End Function

Public Function FromHexString(ByVal text As String) As Variant
Attribute FromHexString.VB_Description = "Decodes hexadecimal text into a Byte array."
    Dim bytes() As Byte
    Dim code As Long
    Dim index As Long

    RequireConvertFactory "FromHexString"
    If (Len(text) Mod 2) <> 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "FromHexString", _
            "Hexadecimal text must have an even number of digits."
    End If
    If Len(text) = 0 Then
        FromHexString = EmptyByteArray()
        Exit Function
    End If
    ReDim bytes(0 To Len(text) \ 2 - 1)
    For index = 0 To UBound(bytes)
        code = HexDigitValue(Mid$(text, index * 2 + 1, 1)) * 16 + _
            HexDigitValue(Mid$(text, index * 2 + 2, 1))
        bytes(index) = code
    Next index
    FromHexString = bytes
End Function

Public Property Get ZipFile() As ROneCOne
Attribute ZipFile.VB_Description = "Gets the System.IO.Compression-style zip surface."
    Dim factory As ROneCOne

    Set factory = New ROneCOne
    factory.ConfigureZipFileFactory
    Set ZipFile = factory
End Property

Public Sub CreateFromDirectory( _
    ByVal sourceDirectoryName As String, _
    ByVal destinationArchiveFileName As String _
)
Attribute CreateFromDirectory.VB_Description = "Zips a folder tree into a new stored archive."
    Dim fso As Object

    RequireZipFileFactory "CreateFromDirectory"
    Set fso = CreateFileSystemObject()
    If Not fso.FolderExists(sourceDirectoryName) Then
        RaiseIoError "The folder '" & sourceDirectoryName & _
            "' does not exist."
    End If
    If fso.FileExists(destinationArchiveFileName) Then
        RaiseIoError "The archive '" & destinationArchiveFileName & _
            "' already exists."
    End If
    ZipWriteArchiveFromDirectory sourceDirectoryName, _
        destinationArchiveFileName
End Sub

Public Sub ExtractToDirectory( _
    ByVal sourceArchiveFileName As String, _
    ByVal destinationDirectoryName As String _
)
Attribute ExtractToDirectory.VB_Description = "Extracts an archive, refusing existing files."
    Dim dataBytes() As Byte
    Dim fso As Object
    Dim parentPath As String
    Dim rawEntry As Variant
    Dim rawEntries As Collection
    Dim safePath As String
    Dim targetPath As String

    RequireZipFileFactory "ExtractToDirectory"
    mZipWork = IoReadAllBytesCore(sourceArchiveFileName, _
        "ExtractToDirectory")
    Set rawEntries = ZipReadDirectoryCore()
    Set fso = CreateFileSystemObject()
    DirectoryCreateCore destinationDirectoryName, "ExtractToDirectory"
    For Each rawEntry In rawEntries
        safePath = SafeRelativeEntryPath(CStr(rawEntry(0)))
        targetPath = fso.BuildPath(destinationDirectoryName, safePath)
        If CBool(rawEntry(7)) Then
            DirectoryCreateCore targetPath, "ExtractToDirectory"
        Else
            parentPath = fso.GetParentFolderName(targetPath)
            If Len(parentPath) > 0 Then
                DirectoryCreateCore parentPath, "ExtractToDirectory"
            End If
            If fso.FileExists(targetPath) Then
                Erase mZipWork
                RaiseIoError "The file '" & targetPath & _
                    "' already exists."
            End If
            dataBytes = ZipExtractEntryCore(CStr(rawEntry(0)), _
                CLng(rawEntry(1)), CDbl(rawEntry(3)), CLng(rawEntry(4)), _
                CLng(rawEntry(5)), CLng(rawEntry(6)), False)
            IoWriteAllBytesCore targetPath, dataBytes, "ExtractToDirectory"
        End If
    Next rawEntry
    Erase mZipWork
End Sub

Public Function OpenRead(ByVal archiveFileName As String) As ROneCOne
Attribute OpenRead.VB_Description = "Opens an archive for reading its entries."
    Dim archive As ROneCOne
    Dim entry As ROneCOne
    Dim entryObjects As Collection
    Dim rawEntry As Variant
    Dim rawEntries As Collection

    RequireZipFileFactory "OpenRead"
    mZipWork = IoReadAllBytesCore(archiveFileName, "OpenRead")
    Set rawEntries = ZipReadDirectoryCore()
    Erase mZipWork
    Set entryObjects = New Collection
    For Each rawEntry In rawEntries
        Set entry = New ROneCOne
        entry.ConfigureZipEntry archiveFileName, CStr(rawEntry(0)), _
            CLng(rawEntry(1)), CDbl(rawEntry(3)), CLng(rawEntry(4)), _
            CLng(rawEntry(5)), CLng(rawEntry(6)), CBool(rawEntry(7))
        entryObjects.Add entry
    Next rawEntry
    Set archive = New ROneCOne
    archive.ConfigureZipArchive archiveFileName, entryObjects
    Set OpenRead = archive
End Function

Public Property Get Entries() As ROneCOne
Attribute Entries.VB_Description = "Gets an archive's entries as a queryable typed list."
    Dim entry As Variant
    Dim results As ROneCOne

    RequireZipArchive "Entries"
    Set results = ROneCOne.ListOf(Me)
    For Each entry In mZipEntries
        results.Add entry
    Next entry
    Set Entries = results
End Property

Public Function GetEntry(ByVal entryName As String) As ROneCOne
Attribute GetEntry.VB_Description = "Finds an entry by exact name; Nothing when absent."
    Dim entry As Variant

    RequireZipArchive "GetEntry"
    For Each entry In mZipEntries
        If entry.FullName = entryName Then
            Set GetEntry = entry
            Exit Function
        End If
    Next entry
End Function

Public Property Get FullName() As String
Attribute FullName.VB_Description = "Gets an entry's full name inside the archive."
    RequireZipEntry "FullName"
    FullName = mName
End Property

Public Property Get CompressedLength() As Long
Attribute CompressedLength.VB_Description = "Gets an entry's compressed size in bytes."
    RequireZipEntry "CompressedLength"
    CompressedLength = mZipEntryCompressedSize
End Property

Public Sub ExtractToFile( _
    ByVal destinationFileName As String, _
    Optional ByVal overwrite As Boolean = False _
)
Attribute ExtractToFile.VB_Description = "Writes one entry's bytes to a file on disk."
    Dim dataBytes() As Byte

    RequireZipEntry "ExtractToFile"
    If mZipEntryIsDirectory Then
        RaiseContractError ERROR_INVALID_OPERATION, "ExtractToFile", _
            "A directory entry has no file to extract."
    End If
    If Not overwrite Then
        If CreateFileSystemObject().FileExists(destinationFileName) Then
            RaiseIoError "The file '" & destinationFileName & _
                "' already exists."
        End If
    End If
    dataBytes = ZipEntryBytes("ExtractToFile")
    IoWriteAllBytesCore destinationFileName, dataBytes, "ExtractToFile"
End Sub

Public Function FileWatcher( _
    ByVal folderPath As String, _
    Optional ByVal filterPattern As String = "*" _
) As ROneCOne
Attribute FileWatcher.VB_Description = "Watches a folder for awaitable file changes."
    Dim instance As ROneCOne

    If Not CreateFileSystemObject().FolderExists(folderPath) Then
        RaiseIoError "The folder '" & folderPath & "' does not exist."
    End If
    If Len(Trim$(filterPattern)) = 0 Then filterPattern = "*"
    Set instance = New ROneCOne
    instance.ConfigureFileWatcher folderPath, filterPattern
    Set FileWatcher = instance
End Function

Public Function WaitForChangeAsync( _
    Optional ByVal cancellationToken As ROneCOne _
) As ROneCOne
Attribute WaitForChangeAsync.VB_Description = "Returns a Task resolving on the next file change."
    Dim taskValue As ROneCOne

    RequireFileWatcher "WaitForChangeAsync"
    ValidateCancellationToken cancellationToken, "WaitForChangeAsync"
    ' The baseline is captured now, so a change landing between this call
    ' and the await is still observed on the first poll.
    Set taskValue = New ROneCOne
    taskValue.ConfigureWatchTask Me, _
        WatcherSnapshot(mName, mWatcherFilter), cancellationToken
    Set WaitForChangeAsync = taskValue
End Function

Public Property Get ChangeType() As String
Attribute ChangeType.VB_Description = "Gets Created, Changed, or Deleted for a file change."
    If mRole <> ROLE_FILE_CHANGE Then
        RaiseContractError ERROR_INVALID_OPERATION, "ChangeType", _
            "ChangeType requires a file change value."
    End If
    ChangeType = mWatcherChangeType
End Property

Public Function Logger( _
    ByVal filePath As String, _
    Optional ByVal minimumLevel As String = "Information" _
) As ROneCOne
Attribute Logger.VB_Description = "Creates a file logger shaped after M.E.Logging."
    Dim instance As ROneCOne

    If Len(Trim$(filePath)) = 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "Logger", _
            "The log file path cannot be empty."
    End If
    Set instance = New ROneCOne
    instance.ConfigureLogger filePath, _
        LoggerLevelValue(minimumLevel, "Logger")
    Set Logger = instance
End Function

Public Sub LogTrace(ByVal template As String, ParamArray values() As Variant)
Attribute LogTrace.VB_Description = "Writes a Trace line when the level is enabled."
    Dim argumentValues() As Variant

    RequireLogger "LogTrace"
    argumentValues = values
    LoggerWrite LOG_LEVEL_TRACE, template, argumentValues
End Sub

Public Sub LogDebug(ByVal template As String, ParamArray values() As Variant)
Attribute LogDebug.VB_Description = "Writes a Debug line when the level is enabled."
    Dim argumentValues() As Variant

    RequireLogger "LogDebug"
    argumentValues = values
    LoggerWrite LOG_LEVEL_DEBUG, template, argumentValues
End Sub

Public Sub LogInformation( _
    ByVal template As String, _
    ParamArray values() As Variant _
)
Attribute LogInformation.VB_Description = "Writes an Information line when enabled."
    Dim argumentValues() As Variant

    RequireLogger "LogInformation"
    argumentValues = values
    LoggerWrite LOG_LEVEL_INFORMATION, template, argumentValues
End Sub

Public Sub LogWarning( _
    ByVal template As String, _
    ParamArray values() As Variant _
)
Attribute LogWarning.VB_Description = "Writes a Warning line when the level is enabled."
    Dim argumentValues() As Variant

    RequireLogger "LogWarning"
    argumentValues = values
    LoggerWrite LOG_LEVEL_WARNING, template, argumentValues
End Sub

Public Sub LogError(ByVal template As String, ParamArray values() As Variant)
Attribute LogError.VB_Description = "Writes an Error line when the level is enabled."
    Dim argumentValues() As Variant

    RequireLogger "LogError"
    argumentValues = values
    LoggerWrite LOG_LEVEL_ERROR, template, argumentValues
End Sub

Public Sub LogCritical( _
    ByVal template As String, _
    ParamArray values() As Variant _
)
Attribute LogCritical.VB_Description = "Writes a Critical line; Critical is never filtered."
    Dim argumentValues() As Variant

    RequireLogger "LogCritical"
    argumentValues = values
    LoggerWrite LOG_LEVEL_CRITICAL, template, argumentValues
End Sub

Public Function IsEnabled(ByVal levelName As String) As Boolean
Attribute IsEnabled.VB_Description = "Tests whether a severity level would be written."
    RequireLogger "IsEnabled"
    IsEnabled = LoggerLevelValue(levelName, "IsEnabled") >= _
        mLoggerMinimumLevel
End Function

Public Property Get MinimumLevel() As String
Attribute MinimumLevel.VB_Description = "Gets the logger's minimum severity level name."
    RequireLogger "MinimumLevel"
    Select Case mLoggerMinimumLevel
        Case LOG_LEVEL_TRACE
            MinimumLevel = "Trace"
        Case LOG_LEVEL_DEBUG
            MinimumLevel = "Debug"
        Case LOG_LEVEL_INFORMATION
            MinimumLevel = "Information"
        Case LOG_LEVEL_WARNING
            MinimumLevel = "Warning"
        Case LOG_LEVEL_ERROR
            MinimumLevel = "Error"
        Case Else
            MinimumLevel = "Critical"
    End Select
End Property

Public Property Get Uri() As ROneCOne
Attribute Uri.VB_Description = "Gets the System.Uri-style escaping surface."
    Dim factory As ROneCOne

    Set factory = New ROneCOne
    factory.ConfigureUriFactory
    Set Uri = factory
End Property

Public Function EscapeDataString(ByVal text As String) As String
Attribute EscapeDataString.VB_Description = "Percent-encodes text for a URI query or path piece."
    Dim buffer As JsonTextBuilder
    Dim encoded As Variant
    Dim index As Long

    RequireUriFactory "EscapeDataString"
    If Len(text) = 0 Then Exit Function
    encoded = IoEncodeText(text, "utf-8", 3)
    JsonBuilderInit buffer, Len(text) + 16
    For index = LBound(encoded) To UBound(encoded)
        If UriIsUnreservedByte(encoded(index)) Then
            JsonBuilderAppend buffer, Chr$(encoded(index))
        Else
            JsonBuilderAppend buffer, "%" & HexByteText(CByte(encoded(index)))
        End If
    Next index
    EscapeDataString = JsonBuilderText(buffer)
End Function

Public Function UnescapeDataString(ByVal text As String) As String
Attribute UnescapeDataString.VB_Description = "Decodes percent escapes, leaving bad escapes alone."
    Dim buffer As JsonTextBuilder
    Dim index As Long
    Dim pendingBytes() As Byte
    Dim pendingCount As Long
    Dim textLength As Long

    RequireUriFactory "UnescapeDataString"
    textLength = Len(text)
    If textLength = 0 Then Exit Function
    ReDim pendingBytes(0 To textLength \ 3)
    JsonBuilderInit buffer, textLength
    index = 1
    Do While index <= textLength
        If Mid$(text, index, 1) = "%" And _
            IsHexPairText(Mid$(text, index + 1, 2)) Then
            pendingBytes(pendingCount) = _
                CByte(Val("&H" & Mid$(text, index + 1, 2)))
            pendingCount = pendingCount + 1
            index = index + 3
        Else
            FlushPendingUtf8 buffer, pendingBytes, pendingCount
            JsonBuilderAppend buffer, Mid$(text, index, 1)
            index = index + 1
        End If
    Loop
    FlushPendingUtf8 buffer, pendingBytes, pendingCount
    UnescapeDataString = JsonBuilderText(buffer)
End Function

Public Property Get WebUtility() As ROneCOne
Attribute WebUtility.VB_Description = "Gets the System.Net.WebUtility-style HTML text surface."
    Dim factory As ROneCOne

    Set factory = New ROneCOne
    factory.ConfigureWebUtilityFactory
    Set WebUtility = factory
End Property

Public Function HtmlEncode(ByVal text As String) As String
Attribute HtmlEncode.VB_Description = "Encodes markup characters and non-ASCII text for HTML."
    Dim buffer As JsonTextBuilder
    Dim code As Long
    Dim index As Long
    Dim lowUnit As Long
    Dim textLength As Long

    RequireWebUtilityFactory "HtmlEncode"
    textLength = Len(text)
    If textLength = 0 Then Exit Function
    JsonBuilderInit buffer, textLength + 16
    index = 1
    Do While index <= textLength
        code = AscW(Mid$(text, index, 1)) And &HFFFF&
        If code >= &HD800& And code <= &HDBFF& And index < textLength Then
            lowUnit = AscW(Mid$(text, index + 1, 1)) And &HFFFF&
            If lowUnit >= &HDC00& And lowUnit <= &HDFFF& Then
                code = &H10000 + (code - &HD800&) * &H400& + _
                    (lowUnit - &HDC00&)
                index = index + 1
            End If
        End If
        Select Case code
            Case 38
                JsonBuilderAppend buffer, "&amp;"
            Case 60
                JsonBuilderAppend buffer, "&lt;"
            Case 62
                JsonBuilderAppend buffer, "&gt;"
            Case 34
                JsonBuilderAppend buffer, "&quot;"
            Case 39
                JsonBuilderAppend buffer, "&#39;"
            Case Is >= 160
                JsonBuilderAppend buffer, "&#" & CStr(code) & ";"
            Case Else
                JsonBuilderAppend buffer, ChrW$(code)
        End Select
        index = index + 1
    Loop
    HtmlEncode = JsonBuilderText(buffer)
End Function

Public Function HtmlDecode(ByVal text As String) As String
Attribute HtmlDecode.VB_Description = "Decodes amp lt gt quot apos and numeric HTML entities."
    Dim buffer As JsonTextBuilder
    Dim entity As String
    Dim index As Long
    Dim semicolonAt As Long
    Dim textLength As Long

    RequireWebUtilityFactory "HtmlDecode"
    textLength = Len(text)
    If textLength = 0 Then Exit Function
    JsonBuilderInit buffer, textLength
    index = 1
    Do While index <= textLength
        If Mid$(text, index, 1) = "&" Then
            semicolonAt = InStr(index + 1, text, ";")
            If semicolonAt > index + 1 And semicolonAt - index <= 10 Then
                entity = Mid$(text, index + 1, semicolonAt - index - 1)
                If TryAppendHtmlEntity(buffer, entity) Then
                    index = semicolonAt + 1
                Else
                    JsonBuilderAppend buffer, "&"
                    index = index + 1
                End If
            Else
                JsonBuilderAppend buffer, "&"
                index = index + 1
            End If
        Else
            JsonBuilderAppend buffer, Mid$(text, index, 1)
            index = index + 1
        End If
    Loop
    HtmlDecode = JsonBuilderText(buffer)
End Function

Public Property Get Xml() As ROneCOne
Attribute Xml.VB_Description = "Gets the System.Xml-style document surface over MSXML6."
    Dim factory As ROneCOne

    Set factory = New ROneCOne
    factory.ConfigureXmlFactory
    Set Xml = factory
End Property

Public Function Load( _
    ByVal path As String, _
    Optional ByVal selectionNamespaces As String = vbNullString _
) As ROneCOne
Attribute Load.VB_Description = "Loads an XML file, honoring its declared encoding."
    Dim dom As Object

    RequireXmlFactory "Load"
    Set dom = CreateSecuredXmlDocument(selectionNamespaces)
    If Not dom.Load(path) Then RaiseXmlParseError dom
    If dom.documentElement Is Nothing Then
        RaiseXmlError "The XML has no document element."
    End If
    Set Load = BuildXmlNode(dom.documentElement)
End Function

Public Property Get Name() As String
Attribute Name.VB_Description = "Gets an XML node, file change, or zip entry name."
    Select Case mRole
        Case ROLE_XML_NODE
            Name = mProviderObject.nodeName
        Case ROLE_FILE_CHANGE
            Name = mName
        Case ROLE_ZIP_ENTRY
            Name = Mid$(mName, InStrRev(mName, "/") + 1)
        Case Else
            RaiseContractError ERROR_INVALID_OPERATION, "Name", _
                "Name requires an XML node, file change, or zip entry."
    End Select
End Property

Public Function GetAttribute(ByVal attributeName As String) As String
Attribute GetAttribute.VB_Description = "Gets an attribute's text, raising when it is absent."
    Dim attributeNode As Object

    RequireXmlNode "GetAttribute"
    Set attributeNode = XmlFindAttribute(attributeName)
    If attributeNode Is Nothing Then
        RaiseXmlError "The attribute '" & attributeName & _
            "' is not present on <" & mProviderObject.nodeName & ">."
    End If
    GetAttribute = attributeNode.Text
End Function

Public Function HasAttribute(ByVal attributeName As String) As Boolean
Attribute HasAttribute.VB_Description = "Tests whether an attribute is present on the node."
    RequireXmlNode "HasAttribute"
    HasAttribute = Not XmlFindAttribute(attributeName) Is Nothing
End Function

Public Function Elements( _
    Optional ByVal elementName As String = vbNullString _
) As ROneCOne
Attribute Elements.VB_Description = "Lists child elements, optionally filtered by name."
    Dim child As Object
    Dim results As ROneCOne

    RequireXmlNode "Elements"
    Set results = ROneCOne.ListOf(Me)
    For Each child In mProviderObject.childNodes
        If child.nodeType = 1 Then
            If Len(elementName) = 0 Or child.nodeName = elementName Then
                results.Add BuildXmlNode(child)
            End If
        End If
    Next child
    Set Elements = results
End Function

Public Function SelectNodes(ByVal xpath As String) As ROneCOne
Attribute SelectNodes.VB_Description = "Runs an XPath query and lists every matching node."
    Dim matched As Object
    Dim node As Object
    Dim results As ROneCOne

    RequireXmlNode "SelectNodes"
    Set matched = XmlRunXPath(xpath, "SelectNodes")
    Set results = ROneCOne.ListOf(Me)
    For Each node In matched
        results.Add BuildXmlNode(node)
    Next node
    Set SelectNodes = results
End Function

Public Function SelectSingleNode(ByVal xpath As String) As ROneCOne
Attribute SelectSingleNode.VB_Description = "Runs an XPath query; Nothing when nothing matches."
    Dim matched As Object

    RequireXmlNode "SelectSingleNode"
    Set matched = XmlRunXPath(xpath, "SelectSingleNode")
    If matched.Length > 0 Then
        Set SelectSingleNode = BuildXmlNode(matched.Item(0))
    End If
End Function

Public Property Get OuterXml() As String
Attribute OuterXml.VB_Description = "Gets the node's serialized markup."
    RequireXmlNode "OuterXml"
    OuterXml = mProviderObject.xml
End Property

Public Function ToXml( _
    Optional ByVal rootName As String = "NewDataSet", _
    Optional ByVal rowName As String = vbNullString _
) As String
Attribute ToXml.VB_Description = "Serializes this DataTable to element-per-column XML."
    If mRole <> ROLE_DATA_TABLE Then
        RaiseContractError ERROR_INVALID_OPERATION, "ToXml", _
            "ToXml requires a DataTable."
    End If
    ToXml = XmlSerializeTable(Me, rootName, rowName)
End Function

Public Property Get Strings() As ROneCOne
Attribute Strings.VB_Description = "Gets the String.Format-style text formatting surface."
    Dim factory As ROneCOne

    Set factory = New ROneCOne
    factory.ConfigureStringsFactory
    Set Strings = factory
End Property

Public Function Format( _
    ByVal template As String, _
    ParamArray values() As Variant _
) As String
Attribute Format.VB_Description = "Formats a template with {index,alignment:format} items."
    Dim argumentValues() As Variant

    RequireStringsFactory "Format"
    ' A ParamArray cannot be forwarded directly; copying it into a plain
    ' dynamic array keeps the composite core callable from AppendFormat too.
    argumentValues = values
    Format = FormatCompositeCore(template, argumentValues)
End Function

Public Function StringBuilder() As ROneCOne
Attribute StringBuilder.VB_Description = "Creates a growable text buffer like StringBuilder."
    Dim builder As ROneCOne

    Set builder = New ROneCOne
    builder.ConfigureStringBuilder
    Set StringBuilder = builder
End Function

Public Function AppendLine(Optional ByVal element As Variant) As ROneCOne
Attribute AppendLine.VB_Description = "Appends text and a CRLF line break to the builder."
    RequireStringBuilder "AppendLine"
    If Not IsMissing(element) Then
        JsonBuilderAppend mStringBuilder, FormatArgText(element, vbNullString)
    End If
    JsonBuilderAppend mStringBuilder, vbCrLf
    Set AppendLine = Me
End Function

Public Function AppendFormat( _
    ByVal template As String, _
    ParamArray values() As Variant _
) As ROneCOne
Attribute AppendFormat.VB_Description = "Appends a formatted template to the builder."
    Dim argumentValues() As Variant

    RequireStringBuilder "AppendFormat"
    argumentValues = values
    JsonBuilderAppend mStringBuilder, _
        FormatCompositeCore(template, argumentValues)
    Set AppendFormat = Me
End Function

Public Property Get Guid() As ROneCOne
Attribute Guid.VB_Description = "Gets the System.Guid-style identifier surface."
    Dim factory As ROneCOne

    Set factory = New ROneCOne
    factory.ConfigureGuidFactory
    Set Guid = factory
End Property

Public Function NewGuid() As String
Attribute NewGuid.VB_Description = "Creates a random version 4 GUID in canonical text form."
    Dim guidBytes(0 To 15) As Byte
    Dim state As Long

    RequireGuidFactory "NewGuid"
    state = CoCreateGuid(guidBytes(0))
    If state <> 0 Then
        RaiseContractError ERROR_INVALID_OPERATION, "NewGuid", _
            "Windows could not create a GUID, status " & CStr(state) & "."
    End If
    NewGuid = GuidBytesText(guidBytes)
End Function

Public Property Get EmptyGuid() As String
Attribute EmptyGuid.VB_Description = "Gets the all-zero GUID in canonical text form."
    RequireGuidFactory "EmptyGuid"
    EmptyGuid = "00000000-0000-0000-0000-000000000000"
End Property

Public Property Get RandomNumberGenerator() As ROneCOne
Attribute RandomNumberGenerator.VB_Description = "Gets the crypto-grade random number surface."
    Dim factory As ROneCOne

    Set factory = New ROneCOne
    factory.ConfigureRandomFactory
    Set RandomNumberGenerator = factory
End Property

Public Function GetBytes(ByVal byteCount As Long) As Variant
Attribute GetBytes.VB_Description = "Fills a new Byte array with crypto-grade random bytes."
    Dim buffer() As Byte
    Dim state As Long

    RequireRandomFactory "GetBytes"
    If byteCount < 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "GetBytes", _
            "The byte count cannot be negative."
    End If
    If byteCount = 0 Then
        GetBytes = EmptyByteArray()
        Exit Function
    End If
    ReDim buffer(0 To byteCount - 1)
    state = BCryptGenRandom(0, buffer(0), byteCount, BCRYPT_USE_SYSTEM_RNG)
    If state <> 0 Then
        RaiseContractError ERROR_INVALID_OPERATION, "GetBytes", _
            "Windows could not produce random bytes, status &H" & _
            Hex$(state) & "."
    End If
    GetBytes = buffer
End Function

Public Function GetInt32( _
    ByVal fromInclusive As Long, _
    ByVal toExclusive As Long _
) As Long
Attribute GetInt32.VB_Description = "Gets a uniform random integer at or above fromInclusive."
    Dim drawn As Double
    Dim randomBytes As Variant
    Dim rangeSize As Double
    Dim rejectionLimit As Double

    RequireRandomFactory "GetInt32"
    If toExclusive <= fromInclusive Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "GetInt32", _
            "toExclusive must be greater than fromInclusive."
    End If
    ' Rejection sampling keeps the distribution uniform: draws at or beyond
    ' the largest whole multiple of the range are thrown away and redrawn.
    rangeSize = CDbl(toExclusive) - CDbl(fromInclusive)
    rejectionLimit = Int(4294967296# / rangeSize) * rangeSize
    Do
        randomBytes = GetBytes(4)
        drawn = CDbl(randomBytes(0)) + CDbl(randomBytes(1)) * 256# + _
            CDbl(randomBytes(2)) * 65536# + CDbl(randomBytes(3)) * 16777216#
    Loop While drawn >= rejectionLimit
    GetInt32 = CLng(fromInclusive + (drawn - Int(drawn / rangeSize) * rangeSize))
End Function

Public Property Get DateTime() As ROneCOne
Attribute DateTime.VB_Description = "Gets the DateTimeOffset-style date and time surface."
    Dim factory As ROneCOne

    Set factory = New ROneCOne
    factory.ConfigureDateTimeFactory
    Set DateTime = factory
End Property

Public Property Get TimeSpan() As ROneCOne
Attribute TimeSpan.VB_Description = "Gets the TimeSpan duration surface."
    Dim factory As ROneCOne

    Set factory = New ROneCOne
    factory.ConfigureTimeSpanFactory
    Set TimeSpan = factory
End Property

Public Function Parse( _
    ByVal text As String, _
    Optional ByVal selectionNamespaces As String = vbNullString _
) As ROneCOne
Attribute Parse.VB_Description = "Parses ISO 8601 text, a duration, or an XML document."
    Select Case mRole
        Case ROLE_DATETIME_FACTORY
            RequireNoNamespaces selectionNamespaces, "Parse"
            Set Parse = ParseDateTimeText(text, "Parse")
        Case ROLE_TIMESPAN_FACTORY
            RequireNoNamespaces selectionNamespaces, "Parse"
            Set Parse = ParseTimeSpanText(text)
        Case ROLE_XML_FACTORY
            Set Parse = XmlParseCore(text, selectionNamespaces)
        Case Else
            RaiseContractError ERROR_INVALID_OPERATION, "Parse", _
                "Parse requires the DateTime, TimeSpan, or Xml surface."
    End Select
End Function

Public Function TryParse( _
    ByVal text As String, _
    ByRef result As ROneCOne _
) As Boolean
Attribute TryParse.VB_Description = "Parses like Parse but reports failure instead of raising."
    If mRole <> ROLE_DATETIME_FACTORY And mRole <> ROLE_TIMESPAN_FACTORY Then
        RaiseContractError ERROR_INVALID_OPERATION, "TryParse", _
            "TryParse requires the DateTime or TimeSpan surface."
    End If
    Set result = Nothing
    On Error GoTo Failed
    If mRole = ROLE_DATETIME_FACTORY Then
        Set result = ParseDateTimeText(text, "TryParse")
    Else
        Set result = ParseTimeSpanText(text)
    End If
    TryParse = True
    Exit Function
Failed:
    Set result = Nothing
End Function

Public Property Get UtcNow() As ROneCOne
Attribute UtcNow.VB_Description = "Gets the current instant with a zero offset."
    Dim systemNow As SYSTEMTIME

    RequireDateTimeFactory "UtcNow"
    GetSystemTime systemNow
    Set UtcNow = BuildDateTime(SystemTimeToMs(systemNow), 0)
End Property

Public Property Get Now() As ROneCOne
Attribute Now.VB_Description = "Gets the current instant at the machine's local offset."
    Dim systemNow As SYSTEMTIME
    Dim utcMs As Double

    RequireDateTimeFactory "Now"
    GetSystemTime systemNow
    utcMs = SystemTimeToMs(systemNow)
    Set Now = BuildDateTime(utcMs, LocalOffsetMinutesForUtc(utcMs, "Now"))
End Property

Public Property Get Today() As ROneCOne
Attribute Today.VB_Description = "Gets the local calendar date at midnight."
    Dim clockDate As Date
    Dim msOfDay As Long
    Dim systemNow As SYSTEMTIME

    RequireDateTimeFactory "Today"
    GetSystemTime systemNow
    MsToClockParts UtcMsToLocalMs(SystemTimeToMs(systemNow), "Today"), _
        clockDate, msOfDay
    Set Today = FromLocal(clockDate)
End Property

Public Function FromUnixTimeSeconds(ByVal value As Variant) As ROneCOne
Attribute FromUnixTimeSeconds.VB_Description = "Creates an instant from Unix epoch seconds."
    RequireDateTimeFactory "FromUnixTimeSeconds"
    Set FromUnixTimeSeconds = BuildDateTime( _
        RequireIntegralNumber(value, "FromUnixTimeSeconds") * MS_PER_SECOND, 0)
End Function

Public Function FromUnixTimeMilliseconds(ByVal value As Variant) As ROneCOne
Attribute FromUnixTimeMilliseconds.VB_Description = "Creates an instant from epoch milliseconds."
    RequireDateTimeFactory "FromUnixTimeMilliseconds"
    Set FromUnixTimeMilliseconds = BuildDateTime( _
        RequireIntegralNumber(value, "FromUnixTimeMilliseconds"), 0)
End Function

Public Function FromLocal(ByVal value As Date) As ROneCOne
Attribute FromLocal.VB_Description = "Wraps a local VBA Date with the machine's clock offset."
    Dim clockMs As Double
    Dim utcMs As Double

    RequireDateTimeFactory "FromLocal"
    clockMs = DateToMs(value)
    utcMs = LocalClockMsToUtcMs(clockMs, "FromLocal")
    Set FromLocal = BuildDateTime(utcMs, CLng((clockMs - utcMs) / MS_PER_MINUTE))
End Function

Public Function FromUtc(ByVal value As Date) As ROneCOne
Attribute FromUtc.VB_Description = "Wraps a universal VBA Date with a zero offset."
    RequireDateTimeFactory "FromUtc"
    Set FromUtc = BuildDateTime(DateToMs(value), 0)
End Function

Public Property Get Year() As Long
Attribute Year.VB_Description = "Gets the calendar year of the clock time."
    Dim clockDate As Date
    Dim msOfDay As Long

    RequireDateTime "Year"
    MsToClockParts DateTimeClockMs(), clockDate, msOfDay
    Year = VBA.Year(clockDate)
End Property

Public Property Get Month() As Long
Attribute Month.VB_Description = "Gets the calendar month of the clock time."
    Dim clockDate As Date
    Dim msOfDay As Long

    RequireDateTime "Month"
    MsToClockParts DateTimeClockMs(), clockDate, msOfDay
    Month = VBA.Month(clockDate)
End Property

Public Property Get Day() As Long
Attribute Day.VB_Description = "Gets the calendar day of the clock time."
    Dim clockDate As Date
    Dim msOfDay As Long

    RequireDateTime "Day"
    MsToClockParts DateTimeClockMs(), clockDate, msOfDay
    Day = VBA.Day(clockDate)
End Property

Public Property Get Hour() As Long
Attribute Hour.VB_Description = "Gets the clock hour from 0 through 23."
    Dim clockDate As Date
    Dim msOfDay As Long

    RequireDateTime "Hour"
    MsToClockParts DateTimeClockMs(), clockDate, msOfDay
    Hour = msOfDay \ 3600000
End Property

Public Property Get Minute() As Long
Attribute Minute.VB_Description = "Gets the clock minute from 0 through 59."
    Dim clockDate As Date
    Dim msOfDay As Long

    RequireDateTime "Minute"
    MsToClockParts DateTimeClockMs(), clockDate, msOfDay
    Minute = (msOfDay \ 60000) Mod 60
End Property

Public Property Get Second() As Long
Attribute Second.VB_Description = "Gets the clock second from 0 through 59."
    Dim clockDate As Date
    Dim msOfDay As Long

    RequireDateTime "Second"
    MsToClockParts DateTimeClockMs(), clockDate, msOfDay
    Second = (msOfDay \ 1000) Mod 60
End Property

Public Property Get Millisecond() As Long
Attribute Millisecond.VB_Description = "Gets the clock millisecond from 0 through 999."
    Dim clockDate As Date
    Dim msOfDay As Long

    RequireDateTime "Millisecond"
    MsToClockParts DateTimeClockMs(), clockDate, msOfDay
    Millisecond = msOfDay Mod 1000
End Property

Public Property Get DayOfWeek() As Long
Attribute DayOfWeek.VB_Description = "Gets the week day with Sunday as zero, like System.DayOfWeek."
    Dim clockDate As Date
    Dim msOfDay As Long

    RequireDateTime "DayOfWeek"
    MsToClockParts DateTimeClockMs(), clockDate, msOfDay
    DayOfWeek = VBA.Weekday(clockDate, vbSunday) - 1
End Property

Public Property Get DayOfYear() As Long
Attribute DayOfYear.VB_Description = "Gets the one-based day number within the clock year."
    Dim clockDate As Date
    Dim msOfDay As Long

    RequireDateTime "DayOfYear"
    MsToClockParts DateTimeClockMs(), clockDate, msOfDay
    DayOfYear = DateDiff("d", DateSerial(VBA.Year(clockDate), 1, 1), clockDate) + 1
End Property

Public Property Get UtcDateTime() As Date
Attribute UtcDateTime.VB_Description = "Gets the instant as a VBA Date in universal time."
    RequireDateTime "UtcDateTime"
    UtcDateTime = MsToDate(mDateTimeMs)
End Property

Public Property Get LocalDateTime() As Date
Attribute LocalDateTime.VB_Description = "Gets the instant as a VBA Date in machine local time."
    RequireDateTime "LocalDateTime"
    LocalDateTime = MsToDate(UtcMsToLocalMs(mDateTimeMs, "LocalDateTime"))
End Property

Public Property Get Offset() As ROneCOne
Attribute Offset.VB_Description = "Gets the clock offset from universal time as a TimeSpan."
    RequireDateTime "Offset"
    Set Offset = BuildTimeSpan(CDbl(mDateTimeOffsetMinutes) * MS_PER_MINUTE)
End Property

Public Function ToUniversalTime() As ROneCOne
Attribute ToUniversalTime.VB_Description = "Returns the same instant with a zero offset."
    RequireDateTime "ToUniversalTime"
    Set ToUniversalTime = BuildDateTime(mDateTimeMs, 0)
End Function

Public Function ToLocalTime() As ROneCOne
Attribute ToLocalTime.VB_Description = "Returns the same instant at the machine's local offset."
    RequireDateTime "ToLocalTime"
    Set ToLocalTime = BuildDateTime(mDateTimeMs, _
        LocalOffsetMinutesForUtc(mDateTimeMs, "ToLocalTime"))
End Function

Public Function ToOffset(ByVal offsetValue As Variant) As ROneCOne
Attribute ToOffset.VB_Description = "Returns the same instant viewed at another clock offset."
    RequireDateTime "ToOffset"
    Set ToOffset = BuildDateTime(mDateTimeMs, _
        CoerceOffsetMinutes(offsetValue, "ToOffset"))
End Function

Public Function AddYears(ByVal value As Long) As ROneCOne
Attribute AddYears.VB_Description = "Adds calendar years to the clock time, keeping the offset."
    RequireDateTime "AddYears"
    Set AddYears = AddClockInterval("yyyy", value, "AddYears")
End Function

Public Function AddMonths(ByVal value As Long) As ROneCOne
Attribute AddMonths.VB_Description = "Adds calendar months, clamping to the target month's end."
    RequireDateTime "AddMonths"
    Set AddMonths = AddClockInterval("m", value, "AddMonths")
End Function

Public Function AddDays(ByVal value As Double) As ROneCOne
Attribute AddDays.VB_Description = "Adds days, including fractions, to the instant."
    RequireDateTime "AddDays"
    Set AddDays = BuildDateTime(mDateTimeMs + RoundToMs(value * MS_PER_DAY), _
        mDateTimeOffsetMinutes)
End Function

Public Function AddHours(ByVal value As Double) As ROneCOne
Attribute AddHours.VB_Description = "Adds hours, including fractions, to the instant."
    RequireDateTime "AddHours"
    Set AddHours = BuildDateTime(mDateTimeMs + RoundToMs(value * MS_PER_HOUR), _
        mDateTimeOffsetMinutes)
End Function

Public Function AddMinutes(ByVal value As Double) As ROneCOne
Attribute AddMinutes.VB_Description = "Adds minutes, including fractions, to the instant."
    RequireDateTime "AddMinutes"
    Set AddMinutes = BuildDateTime(mDateTimeMs + RoundToMs(value * MS_PER_MINUTE), _
        mDateTimeOffsetMinutes)
End Function

Public Function AddSeconds(ByVal value As Double) As ROneCOne
Attribute AddSeconds.VB_Description = "Adds seconds, including fractions, to the instant."
    RequireDateTime "AddSeconds"
    Set AddSeconds = BuildDateTime(mDateTimeMs + RoundToMs(value * MS_PER_SECOND), _
        mDateTimeOffsetMinutes)
End Function

Public Function AddMilliseconds(ByVal value As Double) As ROneCOne
Attribute AddMilliseconds.VB_Description = "Adds whole milliseconds to the instant."
    RequireDateTime "AddMilliseconds"
    Set AddMilliseconds = BuildDateTime(mDateTimeMs + RoundToMs(value), _
        mDateTimeOffsetMinutes)
End Function

Public Function ToUnixTimeSeconds() As LongLong
Attribute ToUnixTimeSeconds.VB_Description = "Gets whole seconds since the Unix epoch."
    RequireDateTime "ToUnixTimeSeconds"
    ToUnixTimeSeconds = CLngLng(FloorDivDouble(mDateTimeMs, MS_PER_SECOND))
End Function

Public Function ToUnixTimeMilliseconds() As LongLong
Attribute ToUnixTimeMilliseconds.VB_Description = "Gets milliseconds since the Unix epoch."
    RequireDateTime "ToUnixTimeMilliseconds"
    ToUnixTimeMilliseconds = CLngLng(mDateTimeMs)
End Function

Public Function ToIsoString() As String
Attribute ToIsoString.VB_Description = "Formats the instant as round-trip ISO 8601 text."
    RequireDateTime "ToIsoString"
    ToIsoString = DateTimeIsoText()
End Function

Public Function ToString(Optional ByVal pattern As String = vbNullString) As String
Attribute ToString.VB_Description = "Formats a DateTime with a token pattern, or a TimeSpan."
    Select Case mRole
        Case ROLE_DATETIME
            If Len(pattern) = 0 Then
                ToString = DateTimeIsoText()
            Else
                ToString = FormatDateTimeTokens(pattern)
            End If
        Case ROLE_TIMESPAN
            If Len(pattern) > 0 Then
                RaiseContractError ERROR_INVALID_ARGUMENT, "ToString", _
                    "TimeSpan.ToString accepts no format pattern."
            End If
            ToString = TimeSpanText()
        Case ROLE_STRING_BUILDER
            If Len(pattern) > 0 Then
                RaiseContractError ERROR_INVALID_ARGUMENT, "ToString", _
                    "StringBuilder.ToString accepts no format pattern."
            End If
            ToString = JsonBuilderText(mStringBuilder)
        Case Else
            RaiseContractError ERROR_INVALID_OPERATION, "ToString", _
                "ToString requires a DateTime, TimeSpan, or StringBuilder."
    End Select
End Function

Public Function CompareTo(ByVal other As Variant) As Long
Attribute CompareTo.VB_Description = "Orders two DateTime instants or two TimeSpan durations."
    Dim difference As Double

    Select Case mRole
        Case ROLE_DATETIME
            difference = mDateTimeMs - RequireDateTimeOperand(other, "CompareTo")
        Case ROLE_TIMESPAN
            difference = mTimeSpanMs - RequireTimeSpanOperandMs(other, "CompareTo")
        Case Else
            RaiseContractError ERROR_INVALID_OPERATION, "CompareTo", _
                "CompareTo requires a DateTime or TimeSpan value."
    End Select
    If difference < 0 Then
        CompareTo = -1
    ElseIf difference > 0 Then
        CompareTo = 1
    End If
End Function

Public Function FromDays(ByVal value As Double) As ROneCOne
Attribute FromDays.VB_Description = "Creates a duration from a day count."
    RequireTimeSpanFactory "FromDays"
    Set FromDays = BuildTimeSpan(RoundToMs(value * MS_PER_DAY))
End Function

Public Function FromHours(ByVal value As Double) As ROneCOne
Attribute FromHours.VB_Description = "Creates a duration from an hour count."
    RequireTimeSpanFactory "FromHours"
    Set FromHours = BuildTimeSpan(RoundToMs(value * MS_PER_HOUR))
End Function

Public Function FromMinutes(ByVal value As Double) As ROneCOne
Attribute FromMinutes.VB_Description = "Creates a duration from a minute count."
    RequireTimeSpanFactory "FromMinutes"
    Set FromMinutes = BuildTimeSpan(RoundToMs(value * MS_PER_MINUTE))
End Function

Public Function FromSeconds(ByVal value As Double) As ROneCOne
Attribute FromSeconds.VB_Description = "Creates a duration from a second count."
    RequireTimeSpanFactory "FromSeconds"
    Set FromSeconds = BuildTimeSpan(RoundToMs(value * MS_PER_SECOND))
End Function

Public Function FromMilliseconds(ByVal value As Double) As ROneCOne
Attribute FromMilliseconds.VB_Description = "Creates a duration from a millisecond count."
    RequireTimeSpanFactory "FromMilliseconds"
    Set FromMilliseconds = BuildTimeSpan(RoundToMs(value))
End Function

Public Property Get Zero() As ROneCOne
Attribute Zero.VB_Description = "Gets an empty duration."
    RequireTimeSpanFactory "Zero"
    Set Zero = BuildTimeSpan(0)
End Property

Public Property Get TotalDays() As Double
Attribute TotalDays.VB_Description = "Gets the duration expressed in fractional days."
    RequireTimeSpan "TotalDays"
    TotalDays = mTimeSpanMs / MS_PER_DAY
End Property

Public Property Get TotalHours() As Double
Attribute TotalHours.VB_Description = "Gets the duration expressed in fractional hours."
    RequireTimeSpan "TotalHours"
    TotalHours = mTimeSpanMs / MS_PER_HOUR
End Property

Public Property Get TotalMinutes() As Double
Attribute TotalMinutes.VB_Description = "Gets the duration expressed in fractional minutes."
    RequireTimeSpan "TotalMinutes"
    TotalMinutes = mTimeSpanMs / MS_PER_MINUTE
End Property

Public Property Get TotalSeconds() As Double
Attribute TotalSeconds.VB_Description = "Gets the duration expressed in fractional seconds."
    RequireTimeSpan "TotalSeconds"
    TotalSeconds = mTimeSpanMs / MS_PER_SECOND
End Property

Public Property Get TotalMilliseconds() As Double
Attribute TotalMilliseconds.VB_Description = "Gets the duration expressed in milliseconds."
    RequireTimeSpan "TotalMilliseconds"
    TotalMilliseconds = mTimeSpanMs
End Property

Public Property Get Days() As Long
Attribute Days.VB_Description = "Gets the whole-day component, signed like the duration."
    RequireTimeSpan "Days"
    Days = CLng(Fix(mTimeSpanMs / MS_PER_DAY))
End Property

Public Property Get Hours() As Long
Attribute Hours.VB_Description = "Gets the hour component, signed like the duration."
    RequireTimeSpan "Hours"
    Hours = CLng(Fix(mTimeSpanMs / MS_PER_HOUR) - _
        Fix(mTimeSpanMs / MS_PER_DAY) * 24#)
End Property

Public Property Get Minutes() As Long
Attribute Minutes.VB_Description = "Gets the minute component, signed like the duration."
    RequireTimeSpan "Minutes"
    Minutes = CLng(Fix(mTimeSpanMs / MS_PER_MINUTE) - _
        Fix(mTimeSpanMs / MS_PER_HOUR) * 60#)
End Property

Public Property Get Seconds() As Long
Attribute Seconds.VB_Description = "Gets the second component, signed like the duration."
    RequireTimeSpan "Seconds"
    Seconds = CLng(Fix(mTimeSpanMs / MS_PER_SECOND) - _
        Fix(mTimeSpanMs / MS_PER_MINUTE) * 60#)
End Property

Public Property Get Milliseconds() As Long
Attribute Milliseconds.VB_Description = "Gets the millisecond component, signed like the duration."
    RequireTimeSpan "Milliseconds"
    Milliseconds = CLng(mTimeSpanMs - Fix(mTimeSpanMs / MS_PER_SECOND) * 1000#)
End Property

Public Function Duration() As ROneCOne
Attribute Duration.VB_Description = "Returns the absolute value of the duration."
    RequireTimeSpan "Duration"
    Set Duration = BuildTimeSpan(Abs(mTimeSpanMs))
End Function

Public Function HttpClient() As ROneCOne
Attribute HttpClient.VB_Description = "Creates an HttpClient for awaitable web requests."
    Dim client As ROneCOne

    Set client = New ROneCOne
    client.ConfigureHttpClient
    Set HttpClient = client
End Function

Public Property Get BaseAddress() As String
Attribute BaseAddress.VB_Description = "Gets or sets the base URL for relative requests."
    RequireHttpClient "BaseAddress"
    BaseAddress = mHttpBaseAddress
End Property

Public Property Let BaseAddress(ByVal value As String)
    RequireHttpClient "BaseAddress"
    mHttpBaseAddress = value
End Property

Public Property Get Timeout() As Long
Attribute Timeout.VB_Description = "Gets or sets the request timeout in milliseconds."
    RequireHttpClient "Timeout"
    Timeout = mHttpTimeoutMs
End Property

Public Property Let Timeout(ByVal value As Long)
    RequireHttpClient "Timeout"
    If value <= 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "Timeout", _
            "The timeout must be a positive number of milliseconds."
    End If
    mHttpTimeoutMs = value
End Property

Public Function DefaultRequestHeader( _
    ByVal headerName As String, _
    ByVal headerValue As String _
) As ROneCOne
Attribute DefaultRequestHeader.VB_Description = "Adds a header sent with every request."
    RequireHttpClient "DefaultRequestHeader"
    If Len(Trim$(headerName)) = 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "DefaultRequestHeader", _
            "The header name cannot be empty."
    End If
    mHttpHeaderNames.Add headerName
    mHttpHeaderValues.Add headerValue
    Set DefaultRequestHeader = Me
End Function

Public Function GetAsync( _
    ByVal url As String, _
    Optional ByVal cancellationToken As ROneCOne _
) As ROneCOne
Attribute GetAsync.VB_Description = "Starts a GET request and returns a Task of HttpResponse."
    RequireHttpClient "GetAsync"
    Set GetAsync = StartHttpSend("GET", url, Empty, False, _
        vbNullString, "Response", "GetAsync", cancellationToken)
End Function

Public Function GetStringAsync( _
    ByVal url As String, _
    Optional ByVal cancellationToken As ROneCOne _
) As ROneCOne
Attribute GetStringAsync.VB_Description = "Starts a GET and returns a Task of the body text."
    RequireHttpClient "GetStringAsync"
    Set GetStringAsync = StartHttpSend("GET", url, Empty, False, _
        vbNullString, "String", "GetStringAsync", cancellationToken)
End Function

Public Function GetByteArrayAsync( _
    ByVal url As String, _
    Optional ByVal cancellationToken As ROneCOne _
) As ROneCOne
Attribute GetByteArrayAsync.VB_Description = "Starts a GET and returns a Task of the body bytes."
    RequireHttpClient "GetByteArrayAsync"
    Set GetByteArrayAsync = StartHttpSend("GET", url, Empty, False, _
        vbNullString, "Bytes", "GetByteArrayAsync", cancellationToken)
End Function

Public Function DownloadFileAsync( _
    ByVal url As String, _
    ByVal filePath As String, _
    Optional ByVal cancellationToken As ROneCOne _
) As ROneCOne
Attribute DownloadFileAsync.VB_Description = "Starts a GET whose Task saves the body to a file."
    Dim taskValue As ROneCOne

    RequireHttpClient "DownloadFileAsync"
    If Len(Trim$(filePath)) = 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "DownloadFileAsync", _
            "The target file path cannot be empty."
    End If
    Set taskValue = StartHttpSend("GET", url, Empty, False, _
        vbNullString, "File", "DownloadFileAsync", cancellationToken)
    taskValue.ConfigureDataName filePath
    Set DownloadFileAsync = taskValue
End Function

Public Function PostAsync( _
    ByVal url As String, _
    ByVal body As String, _
    Optional ByVal contentType As String = vbNullString, _
    Optional ByVal cancellationToken As ROneCOne _
) As ROneCOne
Attribute PostAsync.VB_Description = "Starts a POST request and returns a Task of HttpResponse."
    RequireHttpClient "PostAsync"
    Set PostAsync = StartHttpSend("POST", url, body, True, _
        contentType, "Response", "PostAsync", cancellationToken)
End Function

Public Function PutAsync( _
    ByVal url As String, _
    ByVal body As String, _
    Optional ByVal contentType As String = vbNullString, _
    Optional ByVal cancellationToken As ROneCOne _
) As ROneCOne
Attribute PutAsync.VB_Description = "Starts a PUT request and returns a Task of HttpResponse."
    RequireHttpClient "PutAsync"
    Set PutAsync = StartHttpSend("PUT", url, body, True, _
        contentType, "Response", "PutAsync", cancellationToken)
End Function

Public Function PatchAsync( _
    ByVal url As String, _
    ByVal body As String, _
    Optional ByVal contentType As String = vbNullString, _
    Optional ByVal cancellationToken As ROneCOne _
) As ROneCOne
Attribute PatchAsync.VB_Description = "Starts a PATCH request and returns a Task of HttpResponse."
    RequireHttpClient "PatchAsync"
    Set PatchAsync = StartHttpSend("PATCH", url, body, True, _
        contentType, "Response", "PatchAsync", cancellationToken)
End Function

Public Function DeleteAsync( _
    ByVal url As String, _
    Optional ByVal cancellationToken As ROneCOne _
) As ROneCOne
Attribute DeleteAsync.VB_Description = "Starts a DELETE request and returns a Task of HttpResponse."
    RequireHttpClient "DeleteAsync"
    Set DeleteAsync = StartHttpSend("DELETE", url, Empty, False, _
        vbNullString, "Response", "DeleteAsync", cancellationToken)
End Function

Public Function SendAsync( _
    ByVal httpMethod As String, _
    ByVal url As String, _
    Optional ByVal body As Variant, _
    Optional ByVal contentType As String = vbNullString, _
    Optional ByVal cancellationToken As ROneCOne _
) As ROneCOne
Attribute SendAsync.VB_Description = "Starts any HTTP method and returns a Task of HttpResponse."
    Dim hasBody As Boolean
    Dim payload As Variant

    RequireHttpClient "SendAsync"
    If Len(Trim$(httpMethod)) = 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "SendAsync", _
            "The HTTP method cannot be empty."
    End If
    hasBody = Not IsMissing(body)
    If hasBody Then
        payload = CStr(body)
    Else
        payload = Empty
    End If
    Set SendAsync = StartHttpSend(UCase$(httpMethod), url, payload, _
        hasBody, contentType, "Response", "SendAsync", cancellationToken)
End Function

Public Property Get StatusCode() As Long
Attribute StatusCode.VB_Description = "Gets the HTTP response status code."
    RequireHttpResponse "StatusCode"
    StatusCode = mHttpStatusCode
End Property

Public Property Get ReasonPhrase() As String
Attribute ReasonPhrase.VB_Description = "Gets the HTTP response reason phrase."
    RequireHttpResponse "ReasonPhrase"
    ReasonPhrase = mHttpReasonPhrase
End Property

Public Property Get IsSuccessStatusCode() As Boolean
Attribute IsSuccessStatusCode.VB_Description = "Reports whether the status is 2xx."
    RequireHttpResponse "IsSuccessStatusCode"
    IsSuccessStatusCode = (mHttpStatusCode >= 200 And mHttpStatusCode < 300)
End Property

Public Function EnsureSuccessStatusCode() As ROneCOne
Attribute EnsureSuccessStatusCode.VB_Description = "Raises unless the status is 2xx."
    RequireHttpResponse "EnsureSuccessStatusCode"
    If Not IsSuccessStatusCode Then
        Err.Raise ERROR_HTTP_REQUEST, "ROneCOne.HttpRequestException", _
            HttpStatusFailureMessage(Me)
    End If
    Set EnsureSuccessStatusCode = Me
End Function

Public Property Get Content() As String
Attribute Content.VB_Description = "Gets the response body as text."
    RequireHttpResponse "Content"
    Content = mHttpContent
End Property

Public Function Header(ByVal headerName As String) As String
Attribute Header.VB_Description = "Gets one response header value, or an empty string."
    Dim colonAt As Long
    Dim line As Variant

    RequireHttpResponse "Header"
    For Each line In VBA.Split(mHttpRawHeaders, vbCrLf)
        colonAt = InStr(1, CStr(line), ":")
        If colonAt > 1 Then
            If StrComp(Trim$(Left$(CStr(line), colonAt - 1)), _
                Trim$(headerName), vbTextCompare) = 0 Then
                Header = Trim$(Mid$(CStr(line), colonAt + 1))
                Exit Function
            End If
        End If
    Next line
End Function

Public Property Get AllHeaders() As String
Attribute AllHeaders.VB_Description = "Gets the raw response header block."
    RequireHttpResponse "AllHeaders"
    AllHeaders = mHttpRawHeaders
End Property

Public Function DbConnection( _
    ByVal connectionString As String, _
    Optional ByVal providerName As String = "OLEDB" _
) As ROneCOne
Attribute DbConnection.VB_Description = "Creates a lazy OLE DB or ODBC connection."
    Dim connection As ROneCOne

    Set connection = New ROneCOne
    connection.ConfigureDbConnection connectionString, providerName
    Set DbConnection = connection
End Function

Public Function DbCommand( _
    ByVal commandText As String, _
    ByVal connection As ROneCOne _
) As ROneCOne
Attribute DbCommand.VB_Description = "Creates a parameterized database command."
    Dim command As ROneCOne

    RequireDbConnection connection, "DbCommand"
    Set command = New ROneCOne
    command.ConfigureDbCommand commandText, connection
    Set DbCommand = command
End Function

Public Function DbParameter( _
    ByVal parameterName As String, _
    ByVal value As Variant, _
    Optional ByVal direction As Long = 1 _
) As ROneCOne
Attribute DbParameter.VB_Description = "Creates an inferred database parameter."
    Dim parameter As ROneCOne

    Set parameter = New ROneCOne
    parameter.ConfigureDbParameter parameterName, WrapDataValue(value), direction
    Set DbParameter = parameter
End Function

Public Function DbDataAdapter(ByVal selectCommand As ROneCOne) As ROneCOne
Attribute DbDataAdapter.VB_Description = "Creates a DataAdapter over a select command."
    Dim adapter As ROneCOne

    RequireDbCommand selectCommand, "DbDataAdapter"
    Set adapter = New ROneCOne
    adapter.ConfigureDbAdapter selectCommand
    Set DbDataAdapter = adapter
End Function

Public Function Queryable(ByVal tableName As String) As ROneCOne
Attribute Queryable.VB_Description = "Creates a deferred SQL query over one table."
    Dim query As ROneCOne

    RequireDbConnection Me, "Queryable"
    Set query = New ROneCOne
    query.ConfigureQueryable Me, tableName
    Set Queryable = query
End Function

Public Function SelectColumns(ParamArray columnNames() As Variant) As ROneCOne
Attribute SelectColumns.VB_Description = "Projects the query onto the named columns."
    Dim clone As ROneCOne
    Dim columns As Collection
    Dim hasAny As Boolean
    Dim index As Long

    RequireQueryable "SelectColumns"
    On Error Resume Next
    hasAny = (UBound(columnNames) >= LBound(columnNames))
    On Error GoTo 0
    If Not hasAny Then
        RaiseQueryError "SelectColumns requires at least one column name."
    End If
    Set columns = New Collection
    For index = LBound(columnNames) To UBound(columnNames)
        If VarTypeOf(columnNames(index)) <> vbString Then
            RaiseQueryError "SelectColumns takes column names as text."
        End If
        columns.Add ValidateQueryIdentifier( _
            CStr(columnNames(index)), "SelectColumns")
    Next index
    Set clone = QueryableClone()
    clone.InternalQueryableSetColumns columns
    Set SelectColumns = clone
End Function

Public Function ToDataTable() As ROneCOne
Attribute ToDataTable.VB_Description = "Runs the query and lands the rows in a typed DataTable."
    RequireQueryable "ToDataTable"
    Set ToDataTable = QueryableExecuteTable()
End Function

Public Function ToDataTableAsync( _
    Optional ByVal cancellationToken As ROneCOne _
) As ROneCOne
Attribute ToDataTableAsync.VB_Description = "Starts the query async; the Task result is the table."
    Dim adapter As ROneCOne
    Dim arguments As Collection
    Dim command As ROneCOne
    Dim pending As ROneCOne
    Dim table As ROneCOne
    Dim taskValue As ROneCOne

    RequireQueryable "ToDataTableAsync"
    ValidateCancellationToken cancellationToken, "ToDataTableAsync"
    Set command = QueryableBuildCommand(QueryableBuildSql(False))
    Set adapter = ROneCOne.DbDataAdapter(command)
    Set table = ROneCOne.DataTable(mSqlQueryTable)
    Set arguments = New Collection
    arguments.Add WrapInvocationArgument(adapter)
    arguments.Add WrapInvocationArgument(table)
    Set pending = command.InternalBeginAsyncExecute
    Set taskValue = New ROneCOne
    taskValue.ConfigureDbAsyncTask pending, "QueryFill", arguments, _
        cancellationToken
    Set ToDataTableAsync = taskValue
End Function

Public Function CountAsync( _
    Optional ByVal cancellationToken As ROneCOne _
) As ROneCOne
Attribute CountAsync.VB_Description = "Starts a server-side COUNT; the Task result is the number."
    Dim command As ROneCOne

    RequireQueryable "CountAsync"
    Set command = QueryableBuildCommand(QueryableBuildSql(True))
    Set CountAsync = command.ExecuteScalarAsync(cancellationToken)
End Function

Public Function ToSqlString() As String
Attribute ToSqlString.VB_Description = "Renders the exact SQL this query will send."
    RequireQueryable "ToSqlString"
    ToSqlString = QueryableBuildSql(False)
End Function

Public Function SqlParameterValues() As ROneCOne
Attribute SqlParameterValues.VB_Description = "Lists the values bound to the ? markers in order."
    Dim captured As Variant
    Dim list As ROneCOne

    RequireQueryable "SqlParameterValues"
    QueryableBuildSql False
    Set list = ROneCOne.ListOf(vbVariant)
    For Each captured In mSqlQueryLastValues
        list.Add captured
    Next captured
    Set SqlParameterValues = list
End Function

Public Sub AddParameter(ByVal parameter As ROneCOne)
Attribute AddParameter.VB_Description = "Adds a parameter to a database command."
    RequireDbCommand Me, "AddParameter"
    If parameter Is Nothing Or parameter.InternalRole <> ROLE_DB_PARAMETER Then
        RaiseContractError ERROR_INVALID_OPERATION, "AddParameter", _
            "AddParameter requires a DbParameter."
    End If
    ArrAppend mItems, mItemsCount, parameter
End Sub

Public Function WithParameter( _
    ByVal parameterName As String, _
    ByVal value As Variant, _
    Optional ByVal direction As Long = 1 _
) As ROneCOne
Attribute WithParameter.VB_Description = "Adds an inferred parameter and returns the command."
    AddParameter ROneCOne.DbParameter(parameterName, value, direction)
    Set WithParameter = Me
End Function

Public Function WithTimeout(ByVal seconds As Long) As ROneCOne
Attribute WithTimeout.VB_Description = "Sets the command timeout and returns the command."
    RequireDbCommand Me, "WithTimeout"
    If seconds < 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "WithTimeout", _
            "Command timeout cannot be negative."
    End If
    mCommandTimeout = seconds
    Set WithTimeout = Me
End Function

Public Function WithCommandType(ByVal commandType As Long) As ROneCOne
Attribute WithCommandType.VB_Description = "Sets the ADO command type and returns the command."
    RequireDbCommand Me, "WithCommandType"
    mCommandType = commandType
    Set WithCommandType = Me
End Function

Public Function FromColumn( _
    ByVal columnName As String, _
    Optional ByVal sourceVersion As String = "Current" _
) As ROneCOne
Attribute FromColumn.VB_Description = "Binds an adapter parameter to a DataColumn value."
    If mRole <> ROLE_DB_PARAMETER Then
        RaiseContractError ERROR_INVALID_OPERATION, "FromColumn", _
            "FromColumn requires a DbParameter."
    End If
    mSourceColumn = Trim$(columnName)
    If Len(mSourceColumn) = 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "FromColumn", _
            "Source column cannot be empty."
    End If
    mSourceVersion = NormalizeDataRowVersion(sourceVersion, "FromColumn")
    Set FromColumn = Me
End Function

Public Function WithDbType( _
    ByVal providerType As Long, _
    Optional ByVal size As Long = 0 _
) As ROneCOne
Attribute WithDbType.VB_Description = "Sets the ADO data type and optional parameter size."
    If mRole <> ROLE_DB_PARAMETER Then
        RaiseContractError ERROR_INVALID_OPERATION, "WithDbType", _
            "WithDbType requires a DbParameter."
    End If
    If size < 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "WithDbType", _
            "Parameter size cannot be negative."
    End If
    mProviderType = providerType
    mProviderSize = size
    Set WithDbType = Me
End Function

Public Function WithPrecision( _
    ByVal precision As Byte, _
    ByVal numericScale As Byte _
) As ROneCOne
Attribute WithPrecision.VB_Description = "Sets numeric parameter precision and scale."
    If mRole <> ROLE_DB_PARAMETER Then
        RaiseContractError ERROR_INVALID_OPERATION, "WithPrecision", _
            "WithPrecision requires a DbParameter."
    End If
    mProviderPrecision = precision
    mProviderScale = numericScale
    Set WithPrecision = Me
End Function

Public Property Get Parameters() As ROneCOne
Attribute Parameters.VB_Description = "Gets a read-only view of command parameters."
    RequireDbCommand Me, "Parameters"
    Set Parameters = CreateReadOnlyCollectionFromItems(ArrSnapshot(mItems, mItemsCount))
End Property

Public Property Get State() As String
Attribute State.VB_Description = "Gets the .NET-aligned provider resource state."
    Dim providerState As Long

    If mRole = ROLE_DB_TRANSACTION Then
        If mTransactionActive Then State = "Open" Else State = "Closed"
        Exit Property
    End If
    If mRole <> ROLE_DB_CONNECTION Then
        RaiseContractError ERROR_INVALID_OPERATION, "State", _
            "State requires a DbConnection or transaction."
    End If
    If mProviderObject Is Nothing Then
        State = "Closed"
        Exit Property
    End If
    providerState = CLng(CallByName(mProviderObject, "State", VbGet))
    Select Case providerState
        Case 0
            State = "Closed"
        Case 1
            State = "Open"
        Case 2
            State = "Connecting"
        Case 4
            State = "Executing"
        Case 8
            State = "Fetching"
        Case Else
            State = "Open"
    End Select
End Property

Public Property Get AsyncMode() As String
Attribute AsyncMode.VB_Description = "Gets the provider async mode; Native polls inside ADO."
    If mRole <> ROLE_DB_CONNECTION And mRole <> ROLE_DB_COMMAND And _
        mRole <> ROLE_DB_READER And mRole <> ROLE_DB_ADAPTER Then
        RaiseContractError ERROR_INVALID_OPERATION, "AsyncMode", _
            "This property requires a provider resource."
    End If
    ' OpenAsync, ExecuteReaderAsync, ExecuteScalarAsync, and FillAsync start
    ' the operation natively inside ADO and poll its State from the scheduler,
    ' so Excel stays responsive while the provider works. ExecuteNonQueryAsync,
    ' UpdateAsync, and ReadAsync complete their work inside one cooperative
    ' task step because ADO reports no usable async completion for them.
    AsyncMode = "Native"
End Property

Public Sub Connect()
Attribute Connect.VB_Description = "Opens a database connection explicitly."
    RequireDbConnection Me, "Connect"
    If Not mProviderObject Is Nothing Then
        If CLng(CallByName(mProviderObject, "State", VbGet)) <> 0 Then Exit Sub
    End If
    Set mProviderObject = CreateObject("ADODB.Connection")
    If mConnectionTimeout > 0 Then
        CallByName mProviderObject, "ConnectionTimeout", VbLet, mConnectionTimeout
    End If
    CallByName mProviderObject, "Open", VbMethod, mConnectionString
End Sub

Public Function OpenAsync( _
    Optional ByVal cancellationToken As ROneCOne _
) As ROneCOne
Attribute OpenAsync.VB_Description = "Opens the connection async and returns a polled Task."
    Dim taskValue As ROneCOne

    RequireDbConnection Me, "OpenAsync"
    ValidateCancellationToken cancellationToken, "OpenAsync"
    ' The provider starts connecting immediately with adAsyncConnect, so the
    ' handshake proceeds inside ADO while the returned Task polls State
    ' cooperatively on Excel's thread.
    InternalBeginAsyncOpen
    Set taskValue = New ROneCOne
    taskValue.ConfigureDbAsyncTask Me, "Connect", Nothing, cancellationToken
    Set OpenAsync = taskValue
End Function

Friend Sub InternalBeginAsyncOpen()
    If Not mProviderObject Is Nothing Then
        If CLng(CallByName(mProviderObject, "State", VbGet)) <> 0 Then Exit Sub
    End If
    Set mProviderObject = CreateObject("ADODB.Connection")
    If mConnectionTimeout > 0 Then
        CallByName mProviderObject, "ConnectionTimeout", VbLet, mConnectionTimeout
    End If
    mProviderObject.Open mConnectionString, , , ADO_ASYNC_OPTION
End Sub

Friend Function InternalDbConnectPoll() As Boolean
    Dim stateValue As Long

    If mProviderObject Is Nothing Then
        InternalDbConnectPoll = True
        Exit Function
    End If
    stateValue = CLng(CallByName(mProviderObject, "State", VbGet))
    If (stateValue And ADO_STATE_CONNECTING) <> 0 Then Exit Function
    If (stateValue And ADO_STATE_OPEN) <> 0 Then
        InternalDbConnectPoll = True
        Exit Function
    End If
    ' A closed connection after an async open means the handshake failed;
    ' the provider left its explanation in the Errors collection.
    RaiseDbAsyncFailure mProviderObject, "OpenAsync"
End Function

Friend Sub InternalDbAsyncAbort()
    On Error Resume Next
    If mRole = ROLE_DB_CONNECTION Then
        If Not mProviderObject Is Nothing Then
            CallByName mProviderObject, "Cancel", VbMethod
            CallByName mProviderObject, "Close", VbMethod
        End If
        Set mProviderObject = Nothing
    ElseIf mRole = ROLE_DB_READER Then
        If Not mProviderCommand Is Nothing Then
            CallByName mProviderCommand, "Cancel", VbMethod
        End If
        CloseProviderObject mProviderObject
        Set mProviderObject = Nothing
        Set mProviderCommand = Nothing
    End If
    On Error GoTo 0
End Sub

Public Sub Disconnect()
Attribute Disconnect.VB_Description = "Closes a provider resource deterministically."
    If mRole = ROLE_DB_CONNECTION Or mRole = ROLE_DB_READER Then
        If Not mProviderObject Is Nothing Then
            If CLng(CallByName(mProviderObject, "State", VbGet)) <> 0 Then
                CallByName mProviderObject, "Close", VbMethod
            End If
        End If
        Set mProviderObject = Nothing
        Set mProviderCommand = Nothing
        Exit Sub
    End If
    RaiseContractError ERROR_INVALID_OPERATION, "Disconnect", _
        "Disconnect requires a connection or data reader."
End Sub

Public Function ExecuteReader() As ROneCOne
Attribute ExecuteReader.VB_Description = "Executes a command and returns a forward-only reader."
    Dim commandObject As Object
    Dim reader As ROneCOne
    Dim recordset As Object

    RequireDbCommand Me, "ExecuteReader"
    Set commandObject = BuildAdoCommand
    Set recordset = commandObject.Execute
    Set reader = New ROneCOne
    reader.ConfigureDbReader recordset
    Set ExecuteReader = reader
End Function

Public Function ExecuteReaderAsync( _
    Optional ByVal cancellationToken As ROneCOne _
) As ROneCOne
Attribute ExecuteReaderAsync.VB_Description = "Starts async execution; returns a reader Task."
    Dim pending As ROneCOne
    Dim taskValue As ROneCOne

    RequireDbCommand Me, "ExecuteReaderAsync"
    ValidateCancellationToken cancellationToken, "ExecuteReaderAsync"
    Set pending = InternalBeginAsyncExecute
    Set taskValue = New ROneCOne
    taskValue.ConfigureDbAsyncTask pending, "Reader", Nothing, cancellationToken
    Set ExecuteReaderAsync = taskValue
End Function

Friend Function InternalBeginAsyncExecute() As ROneCOne
    Dim commandObject As Object
    Dim pending As ROneCOne
    Dim recordsAffected As Variant
    Dim recordsetObject As Object

    ' The statement starts running inside the provider with adAsyncExecute;
    ' the returned pending reader owns the in-flight recordset and the command
    ' so completion polling and cancellation both have their target.
    Set commandObject = BuildAdoCommand
    Set recordsetObject = commandObject.Execute(recordsAffected, , ADO_ASYNC_OPTION)
    Set pending = New ROneCOne
    pending.ConfigureDbReader recordsetObject
    pending.InternalAttachProviderCommand commandObject
    Set InternalBeginAsyncExecute = pending
End Function

Friend Sub InternalAttachProviderCommand(ByVal commandObject As Object)
    Set mProviderCommand = commandObject
End Sub

Friend Function InternalDbResultPoll() As Boolean
    Dim stateValue As Long

    If mProviderObject Is Nothing Then
        InternalDbResultPoll = True
        Exit Function
    End If
    stateValue = CLng(CallByName(mProviderObject, "State", VbGet))
    If (stateValue And (ADO_STATE_EXECUTING Or ADO_STATE_FETCHING)) <> 0 Then
        Exit Function
    End If
    If (stateValue And ADO_STATE_OPEN) <> 0 Then
        Set mProviderCommand = Nothing
        InternalDbResultPoll = True
        Exit Function
    End If
    ' A closed recordset before any row arrived means the async execution
    ' failed; the connection's Errors collection carries the provider detail.
    RaiseDbAsyncFailure DbAsyncErrorSource, "ExecuteAsync"
End Function

Friend Function InternalScalarHarvest() As Variant
    Dim capturedDescription As String
    Dim capturedNumber As Long
    Dim capturedSource As String

    On Error GoTo HarvestFailed
    If Read Then AssignVariant InternalScalarHarvest, Item(0)
    Disconnect
    On Error GoTo 0
    Exit Function

HarvestFailed:
    capturedNumber = Err.Number
    capturedSource = Err.Source
    capturedDescription = Err.Description
    On Error Resume Next
    Disconnect
    On Error GoTo 0
    Err.Raise capturedNumber, capturedSource, capturedDescription
End Function

Private Function DbAsyncErrorSource() As Object
    Dim connectionObject As Object

    On Error Resume Next
    If Not mProviderCommand Is Nothing Then
        Set connectionObject = CallByName( _
            mProviderCommand, "ActiveConnection", VbGet)
    End If
    On Error GoTo 0
    Set mProviderCommand = Nothing
    Set mProviderObject = Nothing
    Set DbAsyncErrorSource = connectionObject
End Function

Private Sub RaiseDbAsyncFailure( _
    ByVal providerValue As Object, _
    ByVal memberName As String _
)
    Dim errorItem As Object
    Dim errorsObject As Object

    On Error Resume Next
    If Not providerValue Is Nothing Then
        Set errorsObject = CallByName(providerValue, "Errors", VbGet)
        If Not errorsObject Is Nothing Then
            If CLng(CallByName(errorsObject, "Count", VbGet)) > 0 Then
                Set errorItem = CallByName(errorsObject, "Item", VbGet, 0)
            End If
        End If
    End If
    On Error GoTo 0
    If mRole = ROLE_DB_CONNECTION Then Set mProviderObject = Nothing
    If errorItem Is Nothing Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            "The asynchronous provider operation failed without detail."
    End If
    Err.Raise CLng(CallByName(errorItem, "Number", VbGet)), _
        CStr(CallByName(errorItem, "Source", VbGet)), _
        CStr(CallByName(errorItem, "Description", VbGet))
End Sub

Public Function ExecuteNonQuery() As Long
Attribute ExecuteNonQuery.VB_Description = "Executes a command and returns affected rows."
    Dim commandObject As Object
    Dim capturedDescription As String
    Dim capturedNumber As Long
    Dim capturedSource As String
    Dim recordsAffected As Variant
    Dim recordset As Object

    RequireDbCommand Me, "ExecuteNonQuery"
    Set commandObject = BuildAdoCommand
    On Error GoTo ExecuteFailed
    Set recordset = commandObject.Execute(recordsAffected)
    ExecuteNonQuery = CLng(recordsAffected)
    CloseProviderObject recordset
    On Error GoTo 0
    Exit Function

ExecuteFailed:
    capturedNumber = Err.Number
    capturedSource = Err.Source
    capturedDescription = Err.Description
    On Error Resume Next
    CloseProviderObject recordset
    On Error GoTo 0
    Err.Raise capturedNumber, capturedSource, capturedDescription
End Function

Public Function ExecuteNonQueryAsync( _
    Optional ByVal cancellationToken As ROneCOne _
) As ROneCOne
Attribute ExecuteNonQueryAsync.VB_Description = "Returns a Task that executes a non-query."
    Set ExecuteNonQueryAsync = CreateProviderTask( _
        "ExecuteNonQuery", cancellationToken:=cancellationToken)
End Function

Public Function ExecuteScalar() As Variant
Attribute ExecuteScalar.VB_Description = "Executes a command and returns its first field."
    Dim reader As ROneCOne
    Dim capturedDescription As String
    Dim capturedNumber As Long
    Dim capturedSource As String

    On Error GoTo ScalarFailed
    Set reader = ExecuteReader
    If reader.Read Then ExecuteScalar = reader.Item(0)
    reader.Disconnect
    On Error GoTo 0
    Exit Function

ScalarFailed:
    capturedNumber = Err.Number
    capturedSource = Err.Source
    capturedDescription = Err.Description
    On Error Resume Next
    If Not reader Is Nothing Then reader.Disconnect
    On Error GoTo 0
    Err.Raise capturedNumber, capturedSource, capturedDescription
End Function

Public Function ExecuteScalarAsync( _
    Optional ByVal cancellationToken As ROneCOne _
) As ROneCOne
Attribute ExecuteScalarAsync.VB_Description = "Starts async execution; returns a scalar Task."
    Dim pending As ROneCOne
    Dim taskValue As ROneCOne

    RequireDbCommand Me, "ExecuteScalarAsync"
    ValidateCancellationToken cancellationToken, "ExecuteScalarAsync"
    Set pending = InternalBeginAsyncExecute
    Set taskValue = New ROneCOne
    taskValue.ConfigureDbAsyncTask pending, "Scalar", Nothing, cancellationToken
    Set ExecuteScalarAsync = taskValue
End Function

Public Function BeginTransaction() As ROneCOne
Attribute BeginTransaction.VB_Description = "Begins a provider transaction."
    Dim transaction As ROneCOne

    RequireDbConnection Me, "BeginTransaction"
    Connect
    CallByName mProviderObject, "BeginTrans", VbMethod
    Set transaction = New ROneCOne
    transaction.ConfigureDbTransaction Me
    Set BeginTransaction = transaction
End Function

Public Sub Commit()
Attribute Commit.VB_Description = "Commits an active provider transaction."
    If mRole <> ROLE_DB_TRANSACTION Then
        RaiseContractError ERROR_INVALID_OPERATION, "Commit", _
            "Commit requires a transaction."
    End If
    If Not mTransactionActive Then
        RaiseContractError ERROR_INVALID_OPERATION, "Commit", _
            "The transaction has already completed."
    End If
    CallByName mSource.InternalProviderObject, "CommitTrans", VbMethod
    mTransactionActive = False
    mTaskStatus = TASK_STATUS_COMPLETED
End Sub

Public Sub Rollback()
Attribute Rollback.VB_Description = "Rolls back an active provider transaction."
    If mRole <> ROLE_DB_TRANSACTION Then
        RaiseContractError ERROR_INVALID_OPERATION, "Rollback", _
            "Rollback requires a transaction."
    End If
    If Not mTransactionActive Then
        RaiseContractError ERROR_INVALID_OPERATION, "Rollback", _
            "The transaction has already completed."
    End If
    CallByName mSource.InternalProviderObject, "RollbackTrans", VbMethod
    mTransactionActive = False
    mTaskStatus = TASK_STATUS_COMPLETED
End Sub

Public Function Read() As Boolean
Attribute Read.VB_Description = "Advances a database reader to the next record."
    If mRole <> ROLE_DB_READER Then
        RaiseContractError ERROR_INVALID_OPERATION, "Read", _
            "Read requires a data reader."
    End If
    If CBool(CallByName(mProviderObject, "EOF", VbGet)) Then Exit Function
    If mQueryCount > 0 Then CallByName mProviderObject, "MoveNext", VbMethod
    If CBool(CallByName(mProviderObject, "EOF", VbGet)) Then Exit Function
    mQueryCount = mQueryCount + 1
    Read = True
End Function

Public Property Get FieldCount() As Long
Attribute FieldCount.VB_Description = "Gets the number of fields in a data reader."
    Dim fields As Object

    If mRole <> ROLE_DB_READER Then
        RaiseContractError ERROR_INVALID_OPERATION, "FieldCount", _
            "FieldCount requires a data reader."
    End If
    Set fields = CallByName(mProviderObject, "Fields", VbGet)
    FieldCount = CLng(CallByName(fields, "Count", VbGet))
End Property

Public Function GetName(ByVal ordinal As Long) As String
Attribute GetName.VB_Description = "Gets a data-reader field name by ordinal."
    Dim field As Object
    Dim fields As Object

    If mRole <> ROLE_DB_READER Then
        RaiseContractError ERROR_INVALID_OPERATION, "GetName", _
            "GetName requires a data reader."
    End If
    Set fields = CallByName(mProviderObject, "Fields", VbGet)
    Set field = CallByName(fields, "Item", VbGet, ordinal)
    GetName = CStr(CallByName(field, "Name", VbGet))
End Function

Public Function GetOrdinal(ByVal fieldName As String) As Long
Attribute GetOrdinal.VB_Description = "Gets a data-reader field ordinal by name."
    Dim ordinal As Long

    If mRole <> ROLE_DB_READER Then
        RaiseContractError ERROR_INVALID_OPERATION, "GetOrdinal", _
            "GetOrdinal requires a data reader."
    End If
    For ordinal = 0 To FieldCount - 1
        If StrComp(GetName(ordinal), fieldName, vbTextCompare) = 0 Then
            GetOrdinal = ordinal
            Exit Function
        End If
    Next ordinal
    RaiseContractError ERROR_INVALID_ARGUMENT, "GetOrdinal", _
        "The requested field was not found."
End Function

Public Function GetValues() As ROneCOne
Attribute GetValues.VB_Description = "Copies the current reader row into a typed sequence."
    Dim ordinal As Long
    Dim result As ROneCOne

    If mRole <> ROLE_DB_READER Then
        RaiseContractError ERROR_INVALID_OPERATION, "GetValues", _
            "GetValues requires a data reader."
    End If
    Set result = ROneCOne.ListOf(vbVariant)
    For ordinal = 0 To FieldCount - 1
        result.Add ReadProviderField(ordinal)
    Next ordinal
    Set GetValues = result
End Function

Public Function IsDBNull(ByVal ordinal As Variant) As Boolean
Attribute IsDBNull.VB_Description = "Tests whether the current provider field contains Null."
    If mRole <> ROLE_DB_READER Then
        RaiseContractError ERROR_INVALID_OPERATION, "IsDBNull", _
            "IsDBNull requires a data reader."
    End If
    IsDBNull = IsNull(ReadProviderField(ordinal))
End Function

Public Property Get HasRows() As Boolean
Attribute HasRows.VB_Description = "Reports whether a data reader contains at least one row."
    If mRole <> ROLE_DB_READER Then
        RaiseContractError ERROR_INVALID_OPERATION, "HasRows", _
            "HasRows requires a data reader."
    End If
    If mProviderObject Is Nothing Then Exit Property
    HasRows = Not CBool(CallByName(mProviderObject, "EOF", VbGet))
End Property

Public Property Get IsClosed() As Boolean
Attribute IsClosed.VB_Description = "Reports whether a provider reader is closed."
    If mRole <> ROLE_DB_READER Then
        RaiseContractError ERROR_INVALID_OPERATION, "IsClosed", _
            "IsClosed requires a data reader."
    End If
    If mProviderObject Is Nothing Then
        IsClosed = True
    Else
        IsClosed = (CLng(CallByName(mProviderObject, "State", VbGet)) = 0)
    End If
End Property

Public Function ReadAsync( _
    Optional ByVal cancellationToken As ROneCOne _
) As ROneCOne
Attribute ReadAsync.VB_Description = "Returns a Task that advances a provider reader."
    Set ReadAsync = CreateProviderTask( _
        "Read", cancellationToken:=cancellationToken)
End Function

Public Property Get SelectCommand() As ROneCOne
Attribute SelectCommand.VB_Description = "Gets the adapter select command."
    RequireDbAdapter Me, "SelectCommand"
    Set SelectCommand = mSource
End Property

Public Property Set SelectCommand(ByVal value As ROneCOne)
    RequireDbAdapter Me, "SelectCommand"
    RequireDbCommand value, "SelectCommand"
    Set mSource = value
End Property

Public Property Get InsertCommand() As ROneCOne
Attribute InsertCommand.VB_Description = "Gets the adapter insert command."
    RequireDbAdapter Me, "InsertCommand"
    Set InsertCommand = mLeft
End Property

Public Property Set InsertCommand(ByVal value As ROneCOne)
    RequireDbAdapter Me, "InsertCommand"
    RequireDbCommand value, "InsertCommand"
    Set mLeft = value
End Property

Public Property Get UpdateCommand() As ROneCOne
Attribute UpdateCommand.VB_Description = "Gets the adapter update command."
    RequireDbAdapter Me, "UpdateCommand"
    Set UpdateCommand = mRight
End Property

Public Property Set UpdateCommand(ByVal value As ROneCOne)
    RequireDbAdapter Me, "UpdateCommand"
    RequireDbCommand value, "UpdateCommand"
    Set mRight = value
End Property

Public Property Get DeleteCommand() As ROneCOne
Attribute DeleteCommand.VB_Description = "Gets the adapter delete command."
    RequireDbAdapter Me, "DeleteCommand"
    Set DeleteCommand = mNext
End Property

Public Property Set DeleteCommand(ByVal value As ROneCOne)
    RequireDbAdapter Me, "DeleteCommand"
    RequireDbCommand value, "DeleteCommand"
    Set mNext = value
End Property

Public Function UseTransaction( _
    Optional ByVal enabled As Boolean = True _
) As ROneCOne
Attribute UseTransaction.VB_Description = "Makes adapter updates atomic."
    RequireDbAdapter Me, "UseTransaction"
    mAdapterUseTransaction = enabled
    Set UseTransaction = Me
End Function

Public Function ContinueUpdateOnError( _
    Optional ByVal enabled As Boolean = True _
) As ROneCOne
Attribute ContinueUpdateOnError.VB_Description = "Continues adapter updates after row errors."
    RequireDbAdapter Me, "ContinueUpdateOnError"
    mAdapterContinueOnError = enabled
    Set ContinueUpdateOnError = Me
End Function

Public Property Get LastUpdateErrors() As ROneCOne
Attribute LastUpdateErrors.VB_Description = "Gets errors captured by the last adapter update."
    RequireDbAdapter Me, "LastUpdateErrors"
    If mProviderErrors Is Nothing Then Set mProviderErrors = New Collection
    Set LastUpdateErrors = CreateObjectListSnapshot(mProviderErrors)
End Property

Public Function Fill(ByVal dataTarget As ROneCOne) As Long
Attribute Fill.VB_Description = "Fills a DataTable or DataSet from a DataAdapter."
    Dim reader As ROneCOne
    Dim capturedDescription As String
    Dim capturedNumber As Long
    Dim capturedSource As String

    If mRole <> ROLE_DB_ADAPTER Then
        RaiseContractError ERROR_INVALID_OPERATION, "Fill", _
            "Fill requires a DataAdapter."
    End If
    On Error GoTo FillFailed
    Set reader = mSource.ExecuteReader
    Fill = InternalFillFromReader(reader, dataTarget)
    On Error GoTo 0
    Exit Function

FillFailed:
    capturedNumber = Err.Number
    capturedSource = Err.Source
    capturedDescription = Err.Description
    On Error Resume Next
    If Not reader Is Nothing Then reader.Disconnect
    On Error GoTo 0
    Err.Raise capturedNumber, capturedSource, capturedDescription
End Function

Public Function FillAsync( _
    ByVal dataTarget As ROneCOne, _
    Optional ByVal cancellationToken As ROneCOne _
) As ROneCOne
Attribute FillAsync.VB_Description = "Starts a native async execution and returns a fill Task."
    Dim arguments As Collection
    Dim pending As ROneCOne
    Dim taskValue As ROneCOne

    If mRole <> ROLE_DB_ADAPTER Then
        RaiseContractError ERROR_INVALID_OPERATION, "FillAsync", _
            "FillAsync requires a DataAdapter."
    End If
    If dataTarget Is Nothing Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "FillAsync", _
            "The fill target cannot be Nothing."
    End If
    ValidateCancellationToken cancellationToken, "FillAsync"
    Set arguments = New Collection
    arguments.Add WrapInvocationArgument(Me)
    arguments.Add WrapInvocationArgument(dataTarget)
    Set pending = mSource.InternalBeginAsyncExecute
    Set taskValue = New ROneCOne
    taskValue.ConfigureDbAsyncTask pending, "Fill", arguments, cancellationToken
    Set FillAsync = taskValue
End Function

Friend Function InternalFillFromReader( _
    ByVal reader As ROneCOne, _
    ByVal dataTarget As ROneCOne _
) As Long
    Dim table As ROneCOne

    If dataTarget.InternalRole = ROLE_DATA_SET Then
        Set table = ROneCOne.DataTable("Table")
        dataTarget.AddTable table
    Else
        RequireDataTable dataTarget, "Fill"
        Set table = dataTarget
    End If
    InternalFillFromReader = FillTableFromReader(table, reader)
    reader.Disconnect
    table.AcceptChanges
End Function

Public Function Update(ByVal dataTarget As ROneCOne) As Long
Attribute Update.VB_Description = "Persists DataSet changes through configured commands."
    Dim table As ROneCOne

    If mRole <> ROLE_DB_ADAPTER Then
        RaiseContractError ERROR_INVALID_OPERATION, "Update", _
            "Update requires a DataAdapter."
    End If
    If dataTarget.InternalRole = ROLE_DATA_SET Then
        For Each table In dataTarget.InternalDataTables
            Update = Update + UpdateDataTable(table)
        Next table
        Exit Function
    End If
    RequireDataTable dataTarget, "Update"
    Update = UpdateDataTable(dataTarget)
End Function

Public Function UpdateAsync( _
    ByVal dataTarget As ROneCOne, _
    Optional ByVal cancellationToken As ROneCOne _
) As ROneCOne
Attribute UpdateAsync.VB_Description = "Returns a Task that persists DataTable changes."
    Set UpdateAsync = CreateProviderTask( _
        "Update", dataTarget, cancellationToken)
End Function

Public Sub Dispose()
Attribute Dispose.VB_Description = "Releases a provider resource deterministically."
    If mRole = ROLE_DB_CONNECTION Or mRole = ROLE_DB_READER Then
        Disconnect
        Exit Sub
    End If
    If mRole = ROLE_DB_TRANSACTION Then
        If mTransactionActive Then Rollback
        Exit Sub
    End If
    If mRole = ROLE_CANCELLATION_REGISTRATION Then
        If Not mAddingCompleted Then
            mAddingCompleted = True
            mSource.UnregisterCancellationRegistration Me
        End If
        Exit Sub
    End If
    RaiseContractError ERROR_INVALID_OPERATION, "Dispose", _
        "Dispose requires a provider resource or cancellation registration."
End Sub

Public Function Native(ByVal procedureAddress As LongPtr) As ROneCOne
Attribute Native.VB_Description = "Creates a signature-bound native function delegate."
    Dim delegateValue As ROneCOne

    If procedureAddress = 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "Native", _
            "The procedure address cannot be zero."
    End If
    Set delegateValue = New ROneCOne
    delegateValue.ConfigureNative procedureAddress, False
    Set Native = delegateValue
End Function

Public Function NativeAction(ByVal procedureAddress As LongPtr) As ROneCOne
Attribute NativeAction.VB_Description = "Creates a signature-bound native Action delegate."
    Dim delegateValue As ROneCOne

    If procedureAddress = 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "NativeAction", _
            "The procedure address cannot be zero."
    End If
    Set delegateValue = New ROneCOne
    delegateValue.ConfigureNative procedureAddress, True
    Set NativeAction = delegateValue
End Function

Public Function RefOf(ByVal typeToken As Variant) As ROneCOne
Attribute RefOf.VB_Description = "Marks a numeric delegate parameter as ByRef."
    Dim descriptor As ROneCOne
    Dim objectType As String
    Dim valueType As VbVarType

    ResolveTypeDescriptor typeToken, valueType, objectType, "RefOf"
    ValidateReferenceType valueType, "RefOf"
    Set descriptor = New ROneCOne
    descriptor.ConfigureParameter valueType, objectType, True
    Set RefOf = descriptor
End Function

Public Function OutOf(ByVal typeToken As Variant) As ROneCOne
Attribute OutOf.VB_Description = "Creates a typed managed out-value container."
    Dim objectType As String
    Dim referenceValue As ROneCOne
    Dim valueType As VbVarType

    ResolveTypeDescriptor typeToken, valueType, objectType, "OutOf"
    Set referenceValue = New ROneCOne
    referenceValue.ConfigureManagedReference valueType, objectType
    Set OutOf = referenceValue
End Function

Public Property Get Current() As Variant
Attribute Current.VB_Description = "Gets the value held by a managed out-value container."
    RequireReference Me, "Current"
    If mQueryValue Is Nothing Then
        AssignDefaultReferenceResult Current
    Else
        AssignPublicResult Current, mQueryValue
    End If
End Property

Public Function RefLong(ByRef value As Long) As ROneCOne
Attribute RefLong.VB_Description = "Captures a Long variable for true ByRef invocation."
    Set RefLong = CreateReference(VarPtr(value), vbLong)
End Function

Public Function RefByte(ByRef value As Byte) As ROneCOne
Attribute RefByte.VB_Description = "Captures a Byte variable for true ByRef invocation."
    Set RefByte = CreateReference(VarPtr(value), vbByte)
End Function

Public Function RefInteger(ByRef value As Integer) As ROneCOne
Attribute RefInteger.VB_Description = "Captures an Integer for true ByRef invocation."
    Set RefInteger = CreateReference(VarPtr(value), vbInteger)
End Function

Public Function RefLongLong(ByRef value As LongLong) As ROneCOne
Attribute RefLongLong.VB_Description = "Captures a LongLong for true ByRef invocation."
    Set RefLongLong = CreateReference(VarPtr(value), vbLongLong)
End Function

Public Function RefSingle(ByRef value As Single) As ROneCOne
Attribute RefSingle.VB_Description = "Captures a Single variable for true ByRef invocation."
    Set RefSingle = CreateReference(VarPtr(value), vbSingle)
End Function

Public Function RefDouble(ByRef value As Double) As ROneCOne
Attribute RefDouble.VB_Description = "Captures a Double variable for true ByRef invocation."
    Set RefDouble = CreateReference(VarPtr(value), vbDouble)
End Function

Public Function RefCurrency(ByRef value As Currency) As ROneCOne
Attribute RefCurrency.VB_Description = "Captures a Currency value for true ByRef invocation."
    Set RefCurrency = CreateReference(VarPtr(value), vbCurrency)
End Function

Public Function RefDate(ByRef value As Date) As ROneCOne
Attribute RefDate.VB_Description = "Captures a Date variable for true ByRef invocation."
    Set RefDate = CreateReference(VarPtr(value), vbDate)
End Function

Public Function RefBoolean(ByRef value As Boolean) As ROneCOne
Attribute RefBoolean.VB_Description = "Captures a Boolean for true ByRef invocation."
    Set RefBoolean = CreateReference(VarPtr(value), vbBoolean)
End Function

Public Function Takes(ParamArray typeTokens() As Variant) As ROneCOne
Attribute Takes.VB_Description = "Returns a delegate with an immutable parameter signature."
    Dim clone As ROneCOne
    Dim descriptor As ROneCOne
    Dim descriptors As Collection
    Dim rawToken As Variant

    RequireDelegate Me, "Takes"
    Set descriptors = New Collection
    For Each rawToken In typeTokens
        Set descriptor = CreateParameterDescriptor(rawToken, "Takes")
        descriptors.Add descriptor
    Next rawToken

    Set clone = CloneDelegate(Me)
    clone.ConfigureParameterSignature descriptors
    Set Takes = clone
End Function

Public Function Returns(ByVal typeToken As Variant) As ROneCOne
Attribute Returns.VB_Description = "Returns a Func with an immutable result type."
    Dim clone As ROneCOne
    Dim objectType As String
    Dim valueType As VbVarType

    RequireDelegate Me, "Returns"
    If mIsActionDelegate Then
        RaiseContractError ERROR_INVALID_OPERATION, "Returns", _
            "An Action has no result; create a Func for a return value."
    End If
    ResolveTypeDescriptor typeToken, valueType, objectType, "Returns"
    Set clone = CloneDelegate(Me)
    clone.ConfigureReturnSignature valueType, objectType
    Set Returns = clone
End Function

Public Function DynamicInvoke(ByVal arguments As Variant) As Variant
Attribute DynamicInvoke.VB_Description = "Invokes from an array or sequence of arguments."
    Dim argumentValues As Collection
    Dim result As ROneCOne

    RequireDelegate Me, "DynamicInvoke"
    Set argumentValues = BuildDynamicArguments(arguments)
    Set result = InvokeValue(argumentValues)
    AssignPublicResult DynamicInvoke, result
End Function

Public Function Combine(ParamArray delegates() As Variant) As Variant
Attribute Combine.VB_Description = "Combines delegates into a multicast, or joins path parts."
    Dim combined As ROneCOne
    Dim invocationList As Collection
    Dim rawDelegate As Variant
    Dim delegateValue As ROneCOne

    If mRole = ROLE_PATH_FACTORY Then
        Combine = CombinePathParts(delegates)
        Exit Function
    End If
    Set invocationList = New Collection
    If IsDelegateRole(mRole) Then
        AppendInvocationList invocationList, Me
    ElseIf mRole <> ROLE_EMPTY Then
        RaiseContractError ERROR_INVALID_OPERATION, "Combine", _
            "Combine must be called on a delegate or the ROneCOne factory."
    End If

    For Each rawDelegate In delegates
        Set delegateValue = RequireROneCOne(rawDelegate, "Combine")
        RequireDelegate delegateValue, "Combine"
        AppendInvocationList invocationList, delegateValue
    Next rawDelegate
    If invocationList.Count = 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "Combine", _
            "At least one delegate is required."
    End If
    ValidateInvocationCompatibility invocationList

    If invocationList.Count = 1 Then
        Set Combine = invocationList.Item(1)
        Exit Function
    End If
    Set combined = New ROneCOne
    combined.ConfigureMulticast invocationList
    Set Combine = combined
End Function

Public Function Run(ParamArray arguments() As Variant) As Variant
Attribute Run.VB_Description = "Invokes a delegate, schedules Task work, or indexes a sequence."
Attribute Run.VB_UserMemId = 0
    Dim argumentValues As Collection
    Dim rawArgument As Variant

    Set argumentValues = New Collection
    For Each rawArgument In arguments
        argumentValues.Add WrapInvocationArgument(rawArgument)
    Next rawArgument

    AssignInvocationResult Run, argumentValues
End Function

Public Sub Execute(ParamArray arguments() As Variant)
Attribute Execute.VB_Description = "Executes an Action or structured operation."
    Dim argumentValues As Collection
    Dim ignored As ROneCOne
    Dim rawArgument As Variant

    Set argumentValues = New Collection
    For Each rawArgument In arguments
        argumentValues.Add WrapInvocationArgument(rawArgument)
    Next rawArgument

    If IsDelegateRole(mRole) Then
        If Not mIsActionDelegate Then
            RaiseContractError ERROR_INVALID_OPERATION, "Execute", _
                "Execute requires an Action delegate."
        End If
        Set ignored = InvokeValue(argumentValues)
        Exit Sub
    End If
    If mRole = ROLE_TRY Then
        If argumentValues.Count <> 0 Then
            RaiseContractError ERROR_ARITY_MISMATCH, "Execute", _
                "A Try operation does not accept arguments."
        End If
        ExecuteTryOperation
        Exit Sub
    End If
    RaiseContractError ERROR_INVALID_OPERATION, "Execute", _
        "Execute requires an Action or Try operation."
End Sub

Public Function Subscribe(ByVal handler As ROneCOne) As ROneCOne
Attribute Subscribe.VB_Description = "Adds an Action handler to this event."
    RequireEvent Me, "Subscribe"
    ValidateEventHandler handler, "Subscribe"
    mInvocationList.Add handler
    Set Subscribe = Me
End Function

Public Function Unsubscribe(ByVal handler As ROneCOne) As Boolean
Attribute Unsubscribe.VB_Description = "Removes the last matching event handler."
    Dim index As Long

    RequireEvent Me, "Unsubscribe"
    ValidateEventHandler handler, "Unsubscribe"
    For index = mInvocationList.Count To 1 Step -1
        If DelegateIdentityEquals(mInvocationList.Item(index), handler) Then
            mInvocationList.Remove index
            Unsubscribe = True
            Exit Function
        End If
    Next index
End Function

Public Sub Emit(ParamArray arguments() As Variant)
Attribute Emit.VB_Description = "Raises this event for a snapshot of its handlers."
    Dim argumentValues As Collection
    Dim handler As ROneCOne
    Dim handlers As Collection
    Dim ignored As ROneCOne
    Dim rawArgument As Variant

    RequireEvent Me, "Emit"
    Set argumentValues = New Collection
    For Each rawArgument In arguments
        argumentValues.Add WrapInvocationArgument(rawArgument)
    Next rawArgument
    ValidateDelegateArguments argumentValues
    Set handlers = CopyWrappedCollection(mInvocationList)
    For Each handler In handlers
        Set ignored = handler.InvokeValue(argumentValues)
    Next handler
End Sub

Public Property Get HandlerCount() As Long
Attribute HandlerCount.VB_Description = "Gets the number of subscribed event handlers."
    RequireEvent Me, "HandlerCount"
    HandlerCount = mInvocationList.Count
End Property

Public Function Catch( _
    ByVal handlerOrErrorNumber As Variant, _
    Optional ByVal handler As Variant _
) As ROneCOne
Attribute Catch.VB_Description = "Adds a catch-all or error-number Catch handler."
    Dim catchAction As ROneCOne
    Dim errorNumber As Long
    Dim operation As ROneCOne

    RequireTry Me, "Catch"
    If IsMissing(handler) Then
        Set catchAction = RequireROneCOne(handlerOrErrorNumber, "Catch")
    Else
        If Not IsNumeric(handlerOrErrorNumber) Then
            RaiseContractError ERROR_INVALID_ARGUMENT, "Catch", _
                "A filtered Catch requires an error number."
        End If
        errorNumber = CLng(handlerOrErrorNumber)
        Set catchAction = RequireROneCOne(handler, "Catch")
    End If
    ValidateCatchAction catchAction
    Set operation = CloneTryOperation(Me)
    operation.ConfigureCatch errorNumber, catchAction
    Set Catch = operation
End Function

Public Function Finally(ByVal cleanup As ROneCOne) As ROneCOne
Attribute Finally.VB_Description = "Adds the Action that always runs after Try and Catch."
    Dim operation As ROneCOne

    RequireTry Me, "Finally"
    ValidateAction cleanup, "Finally", True
    If Not mFinallyAction Is Nothing Then
        RaiseContractError ERROR_INVALID_OPERATION, "Finally", _
            "A Try operation can have only one Finally block."
    End If
    Set operation = CloneTryOperation(Me)
    operation.ConfigureFinally cleanup
    Set Finally = operation
End Function

Private Sub AssignInvocationResult( _
    ByRef destination As Variant, _
    ByVal argumentValues As Collection _
)
    Dim argumentValue As ROneCOne
    Dim itemIndex As Long
    Dim keyWrapped As ROneCOne
    Dim result As ROneCOne

    If mRole = ROLE_USING Then
        Set result = ExecuteUsingScope(argumentValues)
        AssignPublicResult destination, result
        Exit Sub
    End If

    If mRole = ROLE_TASK_FACTORY Then
        If argumentValues.Count < 1 Or argumentValues.Count > 2 Then
            RaiseContractError ERROR_ARITY_MISMATCH, "Task.Run", _
                "Task.Run requires work and an optional cancellation token."
        End If
        Set argumentValue = argumentValues.Item(1)
        If argumentValues.Count = 1 Then
            Set result = ScheduleTaskWork( _
                RequireROneCOne(Unwrap(argumentValue), "Task.Run"))
        Else
            Set result = ScheduleTaskWork( _
                RequireROneCOne(Unwrap(argumentValue), "Task.Run"), _
                RequireROneCOne(Unwrap(argumentValues.Item(2)), "Task.Run"))
        End If
        Set destination = result
        Exit Sub
    End If

    If mRole = ROLE_VALUE Or mRole = ROLE_PARAMETER Or _
        mRole = ROLE_EXPRESSION Then
        If argumentValues.Count = 1 Then
            Set argumentValue = argumentValues.Item(1)
            If Not argumentValue.InternalHasObjectValue Then
                If VarTypeOf(argumentValue.InternalScalarValue) = vbString Then
                    Set result = Member(CStr(argumentValue.InternalScalarValue))
                    Set destination = result
                    Exit Sub
                End If
            End If
        End If
    End If

    If IsGenericCollectionRole(mRole) And _
        (IsDictionaryCollection Or mCollectionKind = COLLECTION_KEYED) Then
        If argumentValues.Count <> 1 Then
            RaiseContractError ERROR_ARITY_MISMATCH, "Run", _
                "A keyed indexer requires one argument."
        End If
        Set argumentValue = argumentValues.Item(1)
        Set keyWrapped = WrapKey(Unwrap(argumentValue), "Run")
        If HashIndexSupported Then
            Set result = FindHashValue(keyWrapped, "Run")
        Else
            itemIndex = FindKeyIndex(keyWrapped, "Run")
            If itemIndex >= 0 Then Set result = mItems(itemIndex + 1)
        End If
        If result Is Nothing Then
            RaiseContractError ERROR_INVALID_ARGUMENT, "Run", _
                "The requested key was not found."
        End If
        AssignPublicResult destination, result
        Exit Sub
    End If

    If mRole = ROLE_LIST Or mRole = ROLE_QUERY Or _
        IsGenericCollectionRole(mRole) Then
        If argumentValues.Count <> 1 Then
            RaiseContractError ERROR_ARITY_MISMATCH, "Run", _
                "A sequence indexer or member selector requires one argument."
        End If
        Set argumentValue = argumentValues.Item(1)
        If argumentValue.InternalHasObjectValue Then
            RaiseContractError ERROR_TYPE_MISMATCH, "Run", _
                "A sequence index must be numeric or a member name."
        End If
        If VarTypeOf(argumentValue.InternalScalarValue) = vbString Then
            Set result = Condition(CStr(argumentValue.InternalScalarValue))
            Set destination = result
            Exit Sub
        End If
        If IsGenericCollectionRole(mRole) Then
            Set result = MaterializeWrapped.Item( _
                CLng(argumentValue.InternalScalarValue) + 1)
        Else
            Set result = WrappedItem(CLng(argumentValue.InternalScalarValue))
        End If
        If result.InternalHasObjectValue Then
            Set destination = result.InternalObjectValue
        Else
            destination = result.InternalScalarValue
        End If
        Exit Sub
    End If

    Set result = InvokeValue(argumentValues)
    AssignPublicResult destination, result
End Sub

Public Function PipeTo(ByVal nextDelegate As ROneCOne) As ROneCOne
Attribute PipeTo.VB_Description = "Pipes this result into a one-parameter delegate."
    Dim composition As ROneCOne

    RequireDelegate Me, "PipeTo source"
    RequireDelegate nextDelegate, "PipeTo continuation"
    If nextDelegate.Arity <> 1 And nextDelegate.Arity <> -1 Then
        RaiseContractError ERROR_ARITY_MISMATCH, "PipeTo", _
            "The continuation must accept one value."
    End If

    Set composition = New ROneCOne
    composition.ConfigureComposition Me, nextDelegate
    Set PipeTo = composition
End Function

Public Function Add( _
    Optional ByVal operand As Variant, _
    Optional ByVal value As Variant _
) As ROneCOne
Attribute Add.VB_Description = "Adds a list item or creates an addition expression."
    Dim clone As ROneCOne
    Dim wrapped As ROneCOne

    If mRole = ROLE_DATA_ROW Then
        If Not IsMissing(operand) Or Not IsMissing(value) Then
            RaiseContractError ERROR_INVALID_ARGUMENT, "DataRow.Add", _
                "A DataRow.Add call accepts no arguments."
        End If
        mSource.AddRow Me
        Set Add = Me
        Exit Function
    End If

    If mRole = ROLE_DATETIME Then
        If Not IsMissing(value) Then
            RaiseContractError ERROR_INVALID_ARGUMENT, "Add", _
                "DateTime.Add accepts one TimeSpan."
        End If
        Set Add = BuildDateTime(mDateTimeMs + _
            RequireTimeSpanOperandMs(operand, "Add"), mDateTimeOffsetMinutes)
        Exit Function
    End If

    If mRole = ROLE_TIMESPAN Then
        If Not IsMissing(value) Then
            RaiseContractError ERROR_INVALID_ARGUMENT, "Add", _
                "TimeSpan.Add accepts one TimeSpan."
        End If
        Set Add = BuildTimeSpan(mTimeSpanMs + _
            RequireTimeSpanOperandMs(operand, "Add"))
        Exit Function
    End If

    If mRole = ROLE_LIST Then
        If Not IsMissing(value) Then
            RaiseContractError ERROR_INVALID_ARGUMENT, "Add", _
                "List<T>.Add accepts one item."
        End If
        Set wrapped = WrapElement(operand, "Add")
        ArrAppend mItems, mItemsCount, wrapped
        mVersion = mVersion + 1
        Set Add = Me
    ElseIf mRole = ROLE_IMMUTABLE Then
        Set clone = CloneGenericCollection(Me)
        clone.ConfigureCollectionRole MutableRoleForKind(mCollectionKind)
        clone.AddMutable operand, value, "Add"
        clone.ConfigureCollectionRole ROLE_IMMUTABLE
        Set Add = clone
    ElseIf IsGenericCollectionRole(mRole) Then
        AddMutable operand, value, "Add"
        Set Add = Me
    Else
        If Not IsMissing(value) Then
            RaiseContractError ERROR_INVALID_ARGUMENT, "Add", _
                "An addition expression accepts one operand."
        End If
        Set Add = CreateBinaryExpression(OP_ADD, operand)
    End If
End Function

Public Function Subtract(ByVal operand As Variant) As ROneCOne
Attribute Subtract.VB_Description = "Subtracts times or durations, or creates an expression."
    If mRole = ROLE_DATETIME Then
        If IsDateTimeInstance(operand) Then
            Set Subtract = BuildTimeSpan(mDateTimeMs - _
                RequireDateTimeOperand(operand, "Subtract"))
        Else
            Set Subtract = BuildDateTime(mDateTimeMs - _
                RequireTimeSpanOperandMs(operand, "Subtract"), _
                mDateTimeOffsetMinutes)
        End If
        Exit Function
    End If
    If mRole = ROLE_TIMESPAN Then
        Set Subtract = BuildTimeSpan(mTimeSpanMs - _
            RequireTimeSpanOperandMs(operand, "Subtract"))
        Exit Function
    End If
    Set Subtract = CreateBinaryExpression(OP_SUBTRACT, operand)
End Function

Public Function Multiply(ByVal operand As Variant) As ROneCOne
Attribute Multiply.VB_Description = "Creates a multiplication expression."
    Set Multiply = CreateBinaryExpression(OP_MULTIPLY, operand)
End Function

Public Function Divide(ByVal operand As Variant) As ROneCOne
Attribute Divide.VB_Description = "Creates a division expression."
    Set Divide = CreateBinaryExpression(OP_DIVIDE, operand)
End Function

Public Function Modulo(ByVal operand As Variant) As ROneCOne
Attribute Modulo.VB_Description = "Creates a modulo expression."
    Set Modulo = CreateBinaryExpression(OP_MODULO, operand)
End Function

Public Function Negate() As ROneCOne
Attribute Negate.VB_Description = "Negates a duration or creates a negation expression."
    If mRole = ROLE_TIMESPAN Then
        Set Negate = BuildTimeSpan(-mTimeSpanMs)
        Exit Function
    End If
    Set Negate = CreateUnaryExpression(OP_NEGATE)
End Function

Public Function Concat(ByVal operand As Variant) As ROneCOne
Attribute Concat.VB_Description = "Concatenates strings or compatible sequences."
    If mRole = ROLE_LIST Or mRole = ROLE_QUERY Or _
        IsGenericCollectionRole(mRole) Then
        Set Concat = ConcatSequence(operand)
    Else
        Set Concat = CreateBinaryExpression(OP_CONCAT, operand)
    End If
End Function

Public Function EqualTo(ByVal operand As Variant) As ROneCOne
Attribute EqualTo.VB_Description = "Creates an equality expression."
    If mRole = ROLE_WHERE_BUILDER Then
        Set EqualTo = CompleteWhereBinary(OP_EQUAL, operand)
    Else
        Set EqualTo = CreateBinaryExpression(OP_EQUAL, operand)
    End If
End Function

Public Function NotEqualTo(ByVal operand As Variant) As ROneCOne
Attribute NotEqualTo.VB_Description = "Creates an inequality expression."
    If mRole = ROLE_WHERE_BUILDER Then
        Set NotEqualTo = CompleteWhereBinary(OP_NOT_EQUAL, operand)
    Else
        Set NotEqualTo = CreateBinaryExpression(OP_NOT_EQUAL, operand)
    End If
End Function

Public Function EqualToIgnoreCase(ByVal operand As String) As ROneCOne
Attribute EqualToIgnoreCase.VB_Description = "Tests string equality without case sensitivity."
    Set EqualToIgnoreCase = CompleteStringCondition( _
        OP_STRING_EQUAL, operand, vbTextCompare)
End Function

Public Function NotEqualToIgnoreCase(ByVal operand As String) As ROneCOne
Attribute NotEqualToIgnoreCase.VB_Description = "Tests string inequality without case sensitivity."
    Set NotEqualToIgnoreCase = CompleteStringCondition( _
        OP_STRING_NOT_EQUAL, operand, vbTextCompare)
End Function

Public Function LessThan(ByVal operand As Variant) As ROneCOne
Attribute LessThan.VB_Description = "Creates a less-than expression."
    If mRole = ROLE_WHERE_BUILDER Then
        Set LessThan = CompleteWhereBinary(OP_LESS_THAN, operand)
    Else
        Set LessThan = CreateBinaryExpression(OP_LESS_THAN, operand)
    End If
End Function

Public Function LessThanOrEqual(ByVal operand As Variant) As ROneCOne
Attribute LessThanOrEqual.VB_Description = "Creates a less-than-or-equal expression."
    If mRole = ROLE_WHERE_BUILDER Then
        Set LessThanOrEqual = CompleteWhereBinary( _
            OP_LESS_THAN_OR_EQUAL, operand)
    Else
        Set LessThanOrEqual = CreateBinaryExpression( _
            OP_LESS_THAN_OR_EQUAL, operand)
    End If
End Function

Public Function GreaterThan(ByVal operand As Variant) As ROneCOne
Attribute GreaterThan.VB_Description = "Creates a greater-than expression."
    If mRole = ROLE_WHERE_BUILDER Then
        Set GreaterThan = CompleteWhereBinary(OP_GREATER_THAN, operand)
    Else
        Set GreaterThan = CreateBinaryExpression(OP_GREATER_THAN, operand)
    End If
End Function

Public Function GreaterThanOrEqual(ByVal operand As Variant) As ROneCOne
Attribute GreaterThanOrEqual.VB_Description = "Creates a greater-than-or-equal expression."
    If mRole = ROLE_WHERE_BUILDER Then
        Set GreaterThanOrEqual = CompleteWhereBinary( _
            OP_GREATER_THAN_OR_EQUAL, operand)
    Else
        Set GreaterThanOrEqual = CreateBinaryExpression( _
            OP_GREATER_THAN_OR_EQUAL, operand)
    End If
End Function

Public Function AtLeast(ByVal operand As Variant) As ROneCOne
Attribute AtLeast.VB_Description = "Sugar for a greater-than-or-equal expression."
    Set AtLeast = GreaterThanOrEqual(operand)
End Function

Public Function AtMost(ByVal operand As Variant) As ROneCOne
Attribute AtMost.VB_Description = "Sugar for a less-than-or-equal expression."
    Set AtMost = LessThanOrEqual(operand)
End Function

Public Function Between( _
    ByVal lowerBound As Variant, _
    ByVal upperBound As Variant, _
    Optional ByVal inclusive As Boolean = True _
) As ROneCOne
Attribute Between.VB_Description = "Tests whether a value lies between two bounds."
    Dim receiver As ROneCOne
    Dim expression As ROneCOne

    If mRole = ROLE_WHERE_BUILDER Then
        Set receiver = ContextMemberExpression(False)
    Else
        RequireExpression Me, "Between receiver"
        Set receiver = Me
    End If
    If inclusive Then
        Set expression = receiver.GreaterThanOrEqual(lowerBound) _
            .AndAlso(receiver.LessThanOrEqual(upperBound))
    Else
        Set expression = receiver.GreaterThan(lowerBound) _
            .AndAlso(receiver.LessThan(upperBound))
    End If
    Set Between = CompleteWhereOrExpression(expression)
End Function

Public Function OneOf(ParamArray values() As Variant) As ROneCOne
Attribute OneOf.VB_Description = "Tests whether a value equals any supplied candidate."
    Dim candidate As Variant
    Dim expression As ROneCOne
    Dim hasSingleValue As Boolean
    Dim receiver As ROneCOne

    On Error Resume Next
    hasSingleValue = (UBound(values) = LBound(values))
    On Error GoTo 0
    If hasSingleValue Then
        If IsSequenceContainer(values(LBound(values))) Then
            Set OneOf = CreateMembershipCondition( _
                values(LBound(values)), False)
            Exit Function
        End If
    End If
    If mRole = ROLE_WHERE_BUILDER Then
        Set receiver = ContextMemberExpression(False)
    Else
        RequireExpression Me, "OneOf receiver"
        Set receiver = Me
    End If
    For Each candidate In values
        If expression Is Nothing Then
            Set expression = receiver.EqualTo(candidate)
        Else
            Set expression = expression.OrElse(receiver.EqualTo(candidate))
        End If
    Next candidate
    If expression Is Nothing Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "OneOf", _
            "At least one candidate value is required."
    End If
    Set OneOf = CompleteWhereOrExpression(expression)
End Function

Public Function IsIn( _
    ByVal values As Variant, _
    Optional ByVal comparer As Variant _
) As ROneCOne
Attribute IsIn.VB_Description = "Tests membership in a sequence, array, or Collection."
    If IsMissing(comparer) Then
        Set IsIn = CreateMembershipCondition(values, False)
    Else
        Set IsIn = CreateMembershipCondition(values, False, comparer)
    End If
End Function

Public Function NotIn( _
    ByVal values As Variant, _
    Optional ByVal comparer As Variant _
) As ROneCOne
Attribute NotIn.VB_Description = "Tests non-membership in a sequence, array, or Collection."
    If IsMissing(comparer) Then
        Set NotIn = CreateMembershipCondition(values, True)
    Else
        Set NotIn = CreateMembershipCondition(values, True, comparer)
    End If
End Function

Public Function StartsWith( _
    ByVal prefix As String, _
    Optional ByVal compareMode As VbCompareMethod = vbBinaryCompare _
) As ROneCOne
Attribute StartsWith.VB_Description = "Tests a string prefix using an explicit comparison mode."
    Set StartsWith = CompleteStringCondition( _
        OP_STARTS_WITH, prefix, compareMode)
End Function

Public Function EndsWith( _
    ByVal suffix As String, _
    Optional ByVal compareMode As VbCompareMethod = vbBinaryCompare _
) As ROneCOne
Attribute EndsWith.VB_Description = "Tests a string suffix using an explicit comparison mode."
    Set EndsWith = CompleteStringCondition( _
        OP_ENDS_WITH, suffix, compareMode)
End Function

Public Function ContainsText( _
    ByVal fragment As String, _
    Optional ByVal compareMode As VbCompareMethod = vbBinaryCompare _
) As ROneCOne
Attribute ContainsText.VB_Description = "Tests whether a string contains a fragment."
    Set ContainsText = CompleteStringCondition( _
        OP_CONTAINS_TEXT, fragment, compareMode)
End Function

Public Function StartsWithIgnoreCase(ByVal prefix As String) As ROneCOne
Attribute StartsWithIgnoreCase.VB_Description = "Tests a string prefix without case sensitivity."
    Set StartsWithIgnoreCase = StartsWith(prefix, vbTextCompare)
End Function

Public Function EndsWithIgnoreCase(ByVal suffix As String) As ROneCOne
Attribute EndsWithIgnoreCase.VB_Description = "Tests a string suffix without case sensitivity."
    Set EndsWithIgnoreCase = EndsWith(suffix, vbTextCompare)
End Function

Public Function ContainsIgnoreCase(ByVal fragment As String) As ROneCOne
Attribute ContainsIgnoreCase.VB_Description = "Tests string containment without case sensitivity."
    Set ContainsIgnoreCase = ContainsText(fragment, vbTextCompare)
End Function

Public Function MatchesPattern( _
    ByVal pattern As String, _
    Optional ByVal compareMode As VbCompareMethod = vbBinaryCompare _
) As ROneCOne
Attribute MatchesPattern.VB_Description = "Tests a value with VBA's Like pattern syntax."
    Dim expression As ROneCOne
    Dim receiver As ROneCOne
    Dim rightValue As ROneCOne

    Set receiver = ConditionReceiver(False, "MatchesPattern")
    Set rightValue = CoerceExpression(pattern)
    Set expression = New ROneCOne
    ValidateStringCompareMode compareMode, "MatchesPattern"
    expression.ConfigureExpression OP_MATCHES_PATTERN, receiver, _
        rightValue, compareMode
    Set MatchesPattern = CompleteWhereOrExpression(expression)
End Function

Public Function MatchesPatternIgnoreCase(ByVal pattern As String) As ROneCOne
Attribute MatchesPatternIgnoreCase.VB_Description = "Tests a Like pattern without case sensitivity."
    Set MatchesPatternIgnoreCase = MatchesPattern(pattern, vbTextCompare)
End Function

Public Function IsNothing() As ROneCOne
Attribute IsNothing.VB_Description = "Tests whether an object reference is Nothing."
    Dim expression As ROneCOne
    Dim receiver As ROneCOne

    Set receiver = ConditionReceiver(True, "IsNothing")
    Set expression = New ROneCOne
    expression.ConfigureExpression OP_IS_NOTHING, receiver, Nothing
    Set IsNothing = CompleteWhereOrExpression(expression)
End Function

Public Function IsNotNothing() As ROneCOne
Attribute IsNotNothing.VB_Description = "Tests whether an object reference is not Nothing."
    Dim expression As ROneCOne
    Dim receiver As ROneCOne

    Set receiver = ConditionReceiver(True, "IsNotNothing")
    Set expression = New ROneCOne
    expression.ConfigureExpression OP_IS_NOT_NOTHING, receiver, Nothing
    Set IsNotNothing = CompleteWhereOrExpression(expression)
End Function

Public Function IsNullOrEmpty() As ROneCOne
Attribute IsNullOrEmpty.VB_Description = "Tests whether a scalar is Null or an empty string."
    Dim expression As ROneCOne
    Dim receiver As ROneCOne

    Set receiver = ConditionReceiver(False, "IsNullOrEmpty")
    Set expression = New ROneCOne
    expression.ConfigureExpression OP_IS_NULL_OR_EMPTY, receiver, Nothing
    Set IsNullOrEmpty = CompleteWhereOrExpression(expression)
End Function

Public Function IsTrue() As ROneCOne
Attribute IsTrue.VB_Description = "Tests whether a Boolean value is True."
    If mRole = ROLE_WHERE_BUILDER Then
        Set IsTrue = CompleteWhereBinary(OP_EQUAL, True)
    Else
        Set IsTrue = EqualTo(True)
    End If
End Function

Public Function IsFalse() As ROneCOne
Attribute IsFalse.VB_Description = "Tests whether a Boolean value is False."
    If mRole = ROLE_WHERE_BUILDER Then
        Set IsFalse = CompleteWhereBinary(OP_EQUAL, False)
    Else
        Set IsFalse = EqualTo(False)
    End If
End Function

Public Function AndAlso(ByVal operand As Variant) As ROneCOne
Attribute AndAlso.VB_Description = "Creates a short-circuit Boolean AND expression."
    Set AndAlso = CreateBinaryExpression(OP_AND_ALSO, operand)
End Function

Public Function OrElse(ByVal operand As Variant) As ROneCOne
Attribute OrElse.VB_Description = "Creates a short-circuit Boolean OR expression."
    Set OrElse = CreateBinaryExpression(OP_OR_ELSE, operand)
End Function

Public Function NotExpression() As ROneCOne
Attribute NotExpression.VB_Description = "Creates a Boolean NOT expression."
    Set NotExpression = CreateUnaryExpression(OP_NOT)
End Function

Public Function Member( _
    ByVal memberName As String, _
    Optional ByVal returnsObject As Boolean = False, _
    Optional ByVal nullSafe As Boolean = False _
) As ROneCOne
Attribute Member.VB_Description = "Creates an object property-access expression."
    Dim expression As ROneCOne
    Dim operatorCode As Long

    RequireExpression Me, "Member receiver"
    memberName = Trim$(memberName)
    If Len(memberName) = 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "Member", _
            "The member name cannot be empty."
    End If

    If nullSafe Then
        operatorCode = OP_SAFE_MEMBER_VALUE
        If returnsObject Then operatorCode = OP_SAFE_MEMBER_OBJECT
    Else
        operatorCode = OP_MEMBER_VALUE
        If returnsObject Then operatorCode = OP_MEMBER_OBJECT
    End If
    Set expression = New ROneCOne
    expression.ConfigureMemberExpression operatorCode, Me, memberName
    Set Member = expression
End Function

Public Property Get Item(ByVal index As Variant) As Variant
Attribute Item.VB_Description = "Gets or sets a zero-based typed list item."
    Dim itemIndex As Long
    Dim keyWrapped As ROneCOne
    Dim wrapped As ROneCOne

    If mRole = ROLE_DB_READER Then
        Set wrapped = WrapDataValue(ReadProviderField(index))
    ElseIf mRole = ROLE_DATA_ROW Then
        Set wrapped = DataRowItem(index)
    ElseIf IsDictionaryCollection Then
        If HashIndexSupported Then
            If DirectKeyLookupSupported Then
                ValidateKeyValue index, "Item"
                Set wrapped = FindHashValueByKey(index, "Item")
            Else
                Set keyWrapped = WrapKey(index, "Item")
                Set wrapped = FindHashValue(keyWrapped, "Item")
            End If
        Else
            Set keyWrapped = WrapKey(index, "Item")
            itemIndex = FindKeyIndex(keyWrapped, "Item")
            If itemIndex >= 0 Then Set wrapped = mItems(itemIndex + 1)
        End If
        If wrapped Is Nothing Then
            RaiseContractError ERROR_INVALID_ARGUMENT, "Item", _
                "The requested key was not found."
        End If
    ElseIf mCollectionKind = COLLECTION_KEYED Or _
        mCollectionKind = COLLECTION_LOOKUP Then
        If HashIndexSupported Then
            If DirectKeyLookupSupported Then
                ValidateKeyValue index, "Item"
                Set wrapped = FindHashValueByKey(index, "Item")
            Else
                Set keyWrapped = WrapKey(index, "Item")
                Set wrapped = FindHashValue(keyWrapped, "Item")
            End If
        Else
            Set keyWrapped = WrapKey(index, "Item")
            itemIndex = FindKeyIndex(keyWrapped, "Item")
            If itemIndex >= 0 Then Set wrapped = mItems(itemIndex + 1)
        End If
        If wrapped Is Nothing Then
            RaiseContractError ERROR_INVALID_ARGUMENT, "Item", _
                "The requested key was not found."
        End If
    ElseIf mCollectionKind = COLLECTION_READ_ONLY Then
        Set wrapped = mSource.MaterializeWrapped.Item(CLng(index) + 1)
    ElseIf IsGenericCollectionRole(mRole) Then
        ValidateIndex CLng(index), mItemsCount, "Item"
        Set wrapped = mItems(CLng(index) + 1)
    Else
        Set wrapped = WrappedItem(CLng(index))
    End If
    AssignPublicResult Item, wrapped
End Property

Public Property Let Item(ByVal index As Variant, ByVal value As Variant)
    SetItemValue index, value
End Property

Public Property Set Item(ByVal index As Variant, ByVal value As Object)
    SetItemValue index, value
End Property

Public Property Get Count(Optional ByVal predicate As Variant) As Long
Attribute Count.VB_Description = "Counts every element or only predicate matches."
    Dim predicateValue As ROneCOne
    Dim values As Collection
    Dim wrapped As ROneCOne

    If mRole = ROLE_QUERYABLE Then
        If IsMissing(predicate) Then
            Count = QueryableExecuteCount()
        Else
            Count = QueryableWithPredicate( _
                RequireROneCOne(predicate, "Count")).Count
        End If
        Exit Property
    End If
    If mRole = ROLE_LIST And IsMissing(predicate) Then
        Count = mItemsCount
        Exit Property
    End If
    If IsGenericCollectionRole(mRole) And IsMissing(predicate) Then
        If mCollectionKind = COLLECTION_READ_ONLY Then
            Count = mSource.Count
        Else
            Count = mItemsCount
        End If
        Exit Property
    End If
    Set values = MaterializeWrapped
    If IsMissing(predicate) Then
        Count = values.Count
        Exit Property
    End If
    Set predicateValue = NormalizeContextUnary(predicate, "Count", False)
    For Each wrapped In values
        If InvokePredicate(predicateValue, wrapped, "Count") Then Count = Count + 1
    Next wrapped
End Property

Public Property Get Capacity() As Long
Attribute Capacity.VB_Description = "Gets the reserved hash capacity."
    RequireHashCapacityCollection "Capacity"
    Capacity = mHashCapacity
End Property

Public Function EnsureCapacity(ByVal capacityValue As Long) As Long
Attribute EnsureCapacity.VB_Description = "Reserves hash capacity without changing Count."
    RequireHashCapacityCollection "EnsureCapacity"
    If capacityValue < mItemsCount Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "EnsureCapacity", _
            "Capacity cannot be smaller than Count."
    End If
    If capacityValue > mHashCapacity Then
        ResizeHashIndex NextHashCapacity( _
            CLng(((CLngLng(capacityValue) * 10) + 6) \ 7))
    End If
    EnsureCapacity = mHashCapacity
End Function

Public Sub TrimExcess()
Attribute TrimExcess.VB_Description = "Rebuilds the hash index at the smallest safe capacity."
    RequireHashCapacityCollection "TrimExcess"
    If mItemsCount = 0 Then
        ResetHashIndex
    Else
        ResizeHashIndex NextHashCapacity(mItemsCount * 2)
    End If
End Sub

Public Property Get LongCount() As LongLong
Attribute LongCount.VB_Description = "Counts sequence elements as a 64-bit integer."
    Dim values As Collection

    Set values = MaterializeWrapped
    LongCount = CLngLng(values.Count)
End Property

Public Function TryGetNonEnumeratedCount( _
    ByVal countReference As ROneCOne _
) As Boolean
Attribute TryGetNonEnumeratedCount.VB_Description = _
    "Gets a known count without executing a deferred query."
    RequireReference countReference, "TryGetNonEnumeratedCount"
    If mRole = ROLE_QUERY Then Exit Function
    If mRole <> ROLE_LIST And Not IsGenericCollectionRole(mRole) Then
        RaiseContractError ERROR_INVALID_OPERATION, _
            "TryGetNonEnumeratedCount", "This operation requires a sequence."
    End If
    AssignReferenceWrapped countReference, CreateScalarResult(CLng(Count)), _
        "TryGetNonEnumeratedCount"
    TryGetNonEnumeratedCount = True
End Function

Public Property Get GenericTypeName() As String
Attribute GenericTypeName.VB_Description = "Gets the C#-style generic sequence type name."
    RequireSequence Me, "GenericTypeName"
    If IsGenericCollectionRole(mRole) Then
        GenericTypeName = BuildCollectionTypeName
        Exit Property
    End If
    If Len(mElementObjectType) <> 0 Then
        GenericTypeName = "List<" & mElementObjectType & ">"
    Else
        GenericTypeName = "List<" & TypeNameForVarType(mElementType) & ">"
    End If
End Property

Public Sub AddRange(ByVal values As Variant)
Attribute AddRange.VB_Description = "Atomically appends a sequence, array, or Collection."
    Dim incoming As Collection
    Dim wrapped As ROneCOne

    RequireList Me, "AddRange"
    Set incoming = BuildIncomingItems(values, "AddRange")
    For Each wrapped In incoming
        ValidateElementWrapped wrapped, "AddRange"
    Next wrapped
    For Each wrapped In incoming
        ArrAppend mItems, mItemsCount, wrapped
    Next wrapped
    If incoming.Count > 0 Then mVersion = mVersion + 1
End Sub

Public Sub Clear()
Attribute Clear.VB_Description = "Empties a typed collection or a StringBuilder."
    If mRole = ROLE_STRING_BUILDER Then
        mStringBuilder.used = 0
        Exit Sub
    End If
    If IsGenericCollectionRole(mRole) Then
        RequireMutableCollection "Clear"
        ArrReset mItems, mItemsCount
        ArrReset mKeys, mKeysCount
        ArrReset mPriorities, mPrioritiesCount
        If HashIndexSupported Then
            RebuildHashIndex 8
        Else
            ResetHashIndex
        End If
        mVersion = mVersion + 1
        NotifyCollectionChanged "Reset"
        Exit Sub
    End If
    RequireList Me, "Clear"
    ArrReset mItems, mItemsCount
    Set mEnumerationValues = Nothing
    mVersion = mVersion + 1
End Sub

Public Function Contains( _
    ByVal value As Variant, _
    Optional ByVal comparer As Variant _
) As Variant
Attribute Contains.VB_Description = "Tests sequence membership or string containment."
    Dim expression As ROneCOne

    If mRole = ROLE_LIST Or mRole = ROLE_QUERY Or _
        IsGenericCollectionRole(mRole) Then
        If IsExpressionVariant(value) Then
            Set expression = RequireROneCOne(value, "Contains")
            If IsMissing(comparer) Then
                Set Contains = expression.IsIn(Me)
            Else
                Set Contains = expression.IsIn(Me, comparer)
            End If
        ElseIf IsMissing(comparer) Then
            Contains = (IndexOf(value) >= 0)
        Else
            Contains = (IndexOf(value, comparer) >= 0)
        End If
        Exit Function
    End If
    Set Contains = ContainsText(CStr(value))
End Function

Public Function Both(ByVal operand As Variant) As ROneCOne
Attribute Both.VB_Description = "Concise alias for short-circuit Boolean AndAlso."
    Set Both = AndAlso(operand)
End Function

Public Function Either(ByVal operand As Variant) As ROneCOne
Attribute Either.VB_Description = "Concise alias for short-circuit Boolean OrElse."
    Set Either = OrElse(operand)
End Function

Public Function Negated() As ROneCOne
Attribute Negated.VB_Description = "Concise alias for a Boolean NOT expression."
    Set Negated = NotExpression
End Function

Public Function IndexOf( _
    ByVal value As Variant, _
    Optional ByVal comparer As Variant _
) As Long
Attribute IndexOf.VB_Description = "Gets the first zero-based index of a value, or -1."
    Dim candidate As ROneCOne
    Dim comparerValue As ROneCOne
    Dim index As Long
    Dim values As Collection

    RequireSequence Me, "IndexOf"
    Set candidate = WrapElement(value, "IndexOf")
    If Not IsMissing(comparer) Then
        Set comparerValue = NormalizeComparer(comparer, True, "IndexOf")
    End If
    Set values = MaterializeWrapped
    IndexOf = -1
    For index = 1 To values.Count
        If ValuesEqualWithComparer( _
            values.Item(index), candidate, comparerValue, "IndexOf") Then
            IndexOf = index - 1
            Exit Function
        End If
    Next index
End Function

Public Function ContainsMember( _
    ByVal source As ROneCOne, _
    ByVal memberPath As String _
) As ROneCOne
Attribute ContainsMember.VB_Description = "Builds a reusable collection-membership expression."
    RequireSequence Me, "ContainsMember"
    RequireSequence source, "ContainsMember"
    Set ContainsMember = source.Condition(memberPath).IsIn(Me)
End Function

Public Function SequenceEqual( _
    ByVal other As Variant, _
    Optional ByVal comparer As Variant _
) As Boolean
Attribute SequenceEqual.VB_Description = "Tests ordered equality with an optional comparer."
    Dim comparerValue As ROneCOne
    Dim index As Long
    Dim leftValues As Collection
    Dim rightSequence As ROneCOne
    Dim rightValues As Collection

    Set rightSequence = NormalizeMembershipSequence(other, "SequenceEqual")
    Set leftValues = MaterializeWrapped
    Set rightValues = rightSequence.MaterializeWrapped
    If leftValues.Count <> rightValues.Count Then Exit Function
    If Not IsMissing(comparer) Then
        Set comparerValue = NormalizeComparer(comparer, True, "SequenceEqual")
    End If
    For index = 1 To leftValues.Count
        If Not ValuesEqualWithComparer( _
            leftValues.Item(index), rightValues.Item(index), _
            comparerValue, "SequenceEqual") Then Exit Function
    Next index
    SequenceEqual = True
End Function

Public Sub Insert(ByVal index As Long, ByVal value As Variant)
Attribute Insert.VB_Description = "Inserts a typed item at a zero-based list index."
    Dim wrapped As ROneCOne

    RequireList Me, "Insert"
    If index < 0 Or index > mItemsCount Then
        RaiseIndexError "Insert", index
    End If
    Set wrapped = WrapElement(value, "Insert")
    If index = mItemsCount Then
        ArrAppend mItems, mItemsCount, wrapped
    Else
        ArrInsert mItems, mItemsCount, index + 1, wrapped
    End If
    mVersion = mVersion + 1
End Sub

Public Function Remove( _
    ByVal value As Variant, _
    Optional ByVal removal As Variant _
) As Variant
Attribute Remove.VB_Description = "Removes a list item or delegate invocation."
    Dim delegateValue As ROneCOne
    Dim immutableResult As ROneCOne
    Dim result As ROneCOne

    If mRole = ROLE_LIST Then
        If Not IsMissing(removal) Then
            RaiseContractError ERROR_INVALID_ARGUMENT, "Remove", _
                "List removal accepts exactly one value."
        End If
        Remove = RemoveListItem(value)
        Exit Function
    End If

    If mRole = ROLE_IMMUTABLE Then
        If Not IsMissing(removal) Then
            RaiseContractError ERROR_INVALID_ARGUMENT, "Remove", _
                "Immutable collection removal accepts one value or key."
        End If
        Set immutableResult = CloneGenericCollection(Me)
        immutableResult.ConfigureCollectionRole MutableRoleForKind(mCollectionKind)
        RemoveCollectionValue immutableResult, value, "Remove"
        immutableResult.ConfigureCollectionRole ROLE_IMMUTABLE
        Set Remove = immutableResult
        Exit Function
    End If
    If IsGenericCollectionRole(mRole) Then
        If Not IsMissing(removal) Then
            RaiseContractError ERROR_INVALID_ARGUMENT, "Remove", _
                "Collection removal accepts one value or key."
        End If
        Remove = RemoveCollectionValue(Me, value, "Remove")
        Exit Function
    End If

    If IsDelegateRole(mRole) Then
        If Not IsMissing(removal) Then
            RaiseContractError ERROR_INVALID_ARGUMENT, "Remove", _
                "Instance delegate removal accepts one delegate."
        End If
        Set delegateValue = RequireROneCOne(value, "Remove")
        Set result = RemoveDelegateValue(Me, delegateValue)
    ElseIf mRole = ROLE_EMPTY Then
        If IsMissing(removal) Then
            RaiseContractError ERROR_INVALID_ARGUMENT, "Remove", _
                "Static delegate removal requires source and value delegates."
        End If
        Set delegateValue = RequireROneCOne(removal, "Remove")
        Set result = RemoveDelegateValue(RequireROneCOne(value, "Remove"), delegateValue)
    Else
        RaiseContractError ERROR_INVALID_OPERATION, "Remove", _
            "Remove requires a list or delegate."
    End If

    Set Remove = result
End Function

Public Sub RemoveAt(ByVal index As Long)
Attribute RemoveAt.VB_Description = "Removes the item at a zero-based list index."
    If IsGenericCollectionRole(mRole) Then
        RequireMutableCollection "RemoveAt"
        ValidateIndex index, mItemsCount, "RemoveAt"
        RemoveCollectionAt index
        Exit Sub
    End If
    RequireList Me, "RemoveAt"
    ValidateIndex index, mItemsCount, "RemoveAt"
    ArrRemoveAt mItems, mItemsCount, index + 1
    mVersion = mVersion + 1
End Sub

Public Function ContainsKey(ByVal key As Variant) As Boolean
Attribute ContainsKey.VB_Description = "Tests whether a keyed collection contains a key."
    RequireKeyedCollection "ContainsKey"
    ContainsKey = (FindKeyIndex(WrapKey(key, "ContainsKey"), "ContainsKey") >= 0)
End Function

Public Function TryAdd( _
    ByVal keyOrValue As Variant, _
    Optional ByVal value As Variant _
) As Boolean
Attribute TryAdd.VB_Description = "Adds only when a key or set value is absent."
    Dim keyWrapped As ROneCOne

    RequireMutableCollection "TryAdd"
    If IsDictionaryCollection Then
        If IsMissing(value) Then
            RaiseContractError ERROR_INVALID_ARGUMENT, "TryAdd", _
                "A dictionary value is required."
        End If
        Set keyWrapped = WrapKey(keyOrValue, "TryAdd")
        If FindKeyIndex(keyWrapped, "TryAdd") >= 0 Then Exit Function
        AddDictionaryPair keyWrapped, WrapElement(value, "TryAdd"), "TryAdd"
        TryAdd = True
        Exit Function
    End If
    If IsSetCollection Then
        If Not IsMissing(value) Then
            RaiseContractError ERROR_INVALID_ARGUMENT, "TryAdd", _
                "A set accepts one value."
        End If
        Set keyWrapped = WrapElement(keyOrValue, "TryAdd")
        If FindElementIndex(keyWrapped, "TryAdd") >= 0 Then Exit Function
        AddSetValue keyWrapped, "TryAdd"
        TryAdd = True
        Exit Function
    End If
    RaiseContractError ERROR_INVALID_OPERATION, "TryAdd", _
        "TryAdd requires a dictionary or set."
End Function

Public Function TryGetValue( _
    ByVal key As Variant, _
    ByVal valueReference As ROneCOne _
) As Boolean
Attribute TryGetValue.VB_Description = "Copies a keyed value into a typed reference."
    Dim itemIndex As Long
    Dim keyWrapped As ROneCOne
    Dim selected As ROneCOne

    RequireKeyedCollection "TryGetValue"
    RequireReference valueReference, "TryGetValue"
    Set keyWrapped = WrapKey(key, "TryGetValue")
    If HashIndexSupported Then
        Set selected = FindHashValue(keyWrapped, "TryGetValue")
    Else
        itemIndex = FindKeyIndex(keyWrapped, "TryGetValue")
        If itemIndex >= 0 Then Set selected = mItems(itemIndex + 1)
    End If
    If selected Is Nothing Then Exit Function
    AssignReferenceWrapped valueReference, selected, "TryGetValue"
    TryGetValue = True
End Function

Public Property Get Keys() As ROneCOne
Attribute Keys.VB_Description = "Gets a typed snapshot of the collection keys."
    Dim result As ROneCOne
    Dim wrapped As ROneCOne

    RequireKeyedCollection "Keys"
    Set result = New ROneCOne
    result.ConfigureList mKeyType, mKeyObjectType
    For Each wrapped In ArrSnapshot(mKeys, mKeysCount)
        result.InternalAddWrapped wrapped
    Next wrapped
    Set Keys = result
End Property

Public Property Get Values() As ROneCOne
Attribute Values.VB_Description = "Gets a typed snapshot of the collection values."
    Dim result As ROneCOne
    Dim wrapped As ROneCOne

    RequireKeyedCollection "Values"
    Set result = New ROneCOne
    result.ConfigureList mElementType, mElementObjectType
    For Each wrapped In ArrSnapshot(mItems, mItemsCount)
        result.InternalAddWrapped wrapped
    Next wrapped
    Set Values = result
End Property

Public Property Get Key() As Variant
Attribute Key.VB_Description = "Gets the key of a KeyValuePair value."
    If mRole <> ROLE_ENTRY And mCollectionKind <> COLLECTION_GROUPING Then
        RaiseContractError ERROR_INVALID_OPERATION, "Key", _
            "Key requires a KeyValuePair value."
    End If
    AssignPublicResult Key, mLeft
End Property

Public Function Enqueue( _
    ByVal element As Variant, _
    Optional ByVal priority As Variant _
) As ROneCOne
Attribute Enqueue.VB_Description = "Adds an element to a queue or priority queue."
    Dim target As ROneCOne

    If mRole = ROLE_IMMUTABLE Then
        Set target = CloneGenericCollection(Me)
        target.ConfigureCollectionRole MutableRoleForKind(mCollectionKind)
    Else
        RequireMutableCollection "Enqueue"
        Set target = Me
    End If
    If IsPriorityQueueKind(target.InternalCollectionKind) Then
        If IsMissing(priority) Then
            RaiseContractError ERROR_INVALID_ARGUMENT, "Enqueue", _
                "A priority value is required."
        End If
        target.AddPriorityValue target.WrapElement(element, "Enqueue"), _
            target.WrapKey(priority, "Enqueue")
    ElseIf IsQueueKind(target.InternalCollectionKind) Then
        If Not IsMissing(priority) Then
            RaiseContractError ERROR_INVALID_ARGUMENT, "Enqueue", _
                "This queue does not accept a priority."
        End If
        target.InternalQueueValue target.WrapElement(element, "Enqueue")
    Else
        RaiseContractError ERROR_INVALID_OPERATION, "Enqueue", _
            "Enqueue requires a queue."
    End If
    If mRole = ROLE_IMMUTABLE Then
        target.ConfigureCollectionRole ROLE_IMMUTABLE
    End If
    Set Enqueue = target
End Function

Public Function Dequeue() As Variant
Attribute Dequeue.VB_Description = "Removes and returns the next queue element."
    Dim wrapped As ROneCOne

    RequireMutableCollection "Dequeue"
    If Not IsQueueKind(mCollectionKind) And _
        Not IsPriorityQueueKind(mCollectionKind) Then
        RaiseContractError ERROR_INVALID_OPERATION, "Dequeue", _
            "Dequeue requires a queue."
    End If
    If mItemsCount = 0 Then RaiseEmptySequence "Dequeue"
    Set wrapped = mItems(1)
    RemoveCollectionAt 0
    AssignPublicResult Dequeue, wrapped
End Function

Public Function TryDequeue(ByVal valueReference As ROneCOne) As Boolean
Attribute TryDequeue.VB_Description = "Attempts to remove the next queue element."
    RequireMutableCollection "TryDequeue"
    RequireReference valueReference, "TryDequeue"
    If Not IsQueueKind(mCollectionKind) And _
        Not IsPriorityQueueKind(mCollectionKind) Then
        RaiseContractError ERROR_INVALID_OPERATION, "TryDequeue", _
            "TryDequeue requires a queue."
    End If
    If mItemsCount = 0 Then Exit Function
    AssignReferenceWrapped valueReference, mItems(1), "TryDequeue"
    RemoveCollectionAt 0
    TryDequeue = True
End Function

Public Function Push(ByVal element As Variant) As ROneCOne
Attribute Push.VB_Description = "Pushes an element onto a stack."
    Dim target As ROneCOne

    If mRole = ROLE_IMMUTABLE Then
        Set target = CloneGenericCollection(Me)
        target.ConfigureCollectionRole MutableRoleForKind(mCollectionKind)
    Else
        RequireMutableCollection "Push"
        Set target = Me
    End If
    If Not IsStackKind(target.InternalCollectionKind) Then
        RaiseContractError ERROR_INVALID_OPERATION, "Push", _
            "Push requires a stack."
    End If
    target.InternalPushValue target.WrapElement(element, "Push")
    If mRole = ROLE_IMMUTABLE Then
        target.ConfigureCollectionRole ROLE_IMMUTABLE
    End If
    Set Push = target
End Function

Public Function Pop() As Variant
Attribute Pop.VB_Description = "Removes and returns the top stack element."
    Dim wrapped As ROneCOne

    RequireMutableCollection "Pop"
    If Not IsStackKind(mCollectionKind) Then
        RaiseContractError ERROR_INVALID_OPERATION, "Pop", _
            "Pop requires a stack."
    End If
    If mItemsCount = 0 Then RaiseEmptySequence "Pop"
    Set wrapped = mItems(1)
    RemoveCollectionAt 0
    AssignPublicResult Pop, wrapped
End Function

Public Function TryPop(ByVal valueReference As ROneCOne) As Boolean
Attribute TryPop.VB_Description = "Attempts to remove the top stack element."
    RequireMutableCollection "TryPop"
    RequireReference valueReference, "TryPop"
    If Not IsStackKind(mCollectionKind) Then
        RaiseContractError ERROR_INVALID_OPERATION, "TryPop", _
            "TryPop requires a stack."
    End If
    If mItemsCount = 0 Then Exit Function
    AssignReferenceWrapped valueReference, mItems(1), "TryPop"
    RemoveCollectionAt 0
    TryPop = True
End Function

Public Function Peek() As Variant
Attribute Peek.VB_Description = "Returns the next queue or stack element without removal."
    If Not IsQueueKind(mCollectionKind) And _
        Not IsPriorityQueueKind(mCollectionKind) And _
        Not IsStackKind(mCollectionKind) Then
        RaiseContractError ERROR_INVALID_OPERATION, "Peek", _
            "Peek requires a queue or stack."
    End If
    If mItemsCount = 0 Then RaiseEmptySequence "Peek"
    AssignPublicResult Peek, mItems(1)
End Function

Public Function TryPeek(ByVal valueReference As ROneCOne) As Boolean
Attribute TryPeek.VB_Description = "Attempts to read the next queue or stack element."
    RequireReference valueReference, "TryPeek"
    If Not IsQueueKind(mCollectionKind) And _
        Not IsPriorityQueueKind(mCollectionKind) And _
        Not IsStackKind(mCollectionKind) Then
        RaiseContractError ERROR_INVALID_OPERATION, "TryPeek", _
            "TryPeek requires a queue or stack."
    End If
    If mItemsCount = 0 Then Exit Function
    AssignReferenceWrapped valueReference, mItems(1), "TryPeek"
    TryPeek = True
End Function

Public Function AddFirst(ByVal value As Variant) As ROneCOne
Attribute AddFirst.VB_Description = "Adds an item at the beginning of a linked list."
    RequireLinkedList "AddFirst"
    RequireMutableCollection "AddFirst"
    If mItemsCount = 0 Then
        ArrAppend mItems, mItemsCount, WrapElement(value, "AddFirst")
    Else
        ArrInsert mItems, mItemsCount, 1, WrapElement(value, "AddFirst")
    End If
    mVersion = mVersion + 1
    Set AddFirst = FirstNode
End Function

Public Function AddLast(ByVal value As Variant) As ROneCOne
Attribute AddLast.VB_Description = "Adds an item at the end of a linked list."
    RequireLinkedList "AddLast"
    RequireMutableCollection "AddLast"
    ArrAppend mItems, mItemsCount, WrapElement(value, "AddLast")
    mVersion = mVersion + 1
    Set AddLast = LastNode
End Function

Public Sub RemoveFirst()
Attribute RemoveFirst.VB_Description = "Removes the first linked-list item."
    RequireLinkedList "RemoveFirst"
    RequireMutableCollection "RemoveFirst"
    If mItemsCount = 0 Then RaiseEmptySequence "RemoveFirst"
    RemoveCollectionAt 0
End Sub

Public Sub RemoveLast()
Attribute RemoveLast.VB_Description = "Removes the last linked-list item."
    RequireLinkedList "RemoveLast"
    RequireMutableCollection "RemoveLast"
    If mItemsCount = 0 Then RaiseEmptySequence "RemoveLast"
    RemoveCollectionAt mItemsCount - 1
End Sub

Public Property Get FirstNode() As ROneCOne
Attribute FirstNode.VB_Description = "Gets the first linked-list node."
    RequireLinkedList "FirstNode"
    If mItemsCount = 0 Then Exit Property
    Set FirstNode = CreateCollectionNode(0)
End Property

Public Property Get LastNode() As ROneCOne
Attribute LastNode.VB_Description = "Gets the last linked-list node."
    RequireLinkedList "LastNode"
    If mItemsCount = 0 Then Exit Property
    Set LastNode = CreateCollectionNode(mItemsCount - 1)
End Property

Public Property Get MinValue() As Variant
Attribute MinValue.VB_Description = "Gets the minimum element of a sorted set."
    RequireSortedSet "MinValue"
    If mItemsCount = 0 Then RaiseEmptySequence "MinValue"
    AssignPublicResult MinValue, mItems(1)
End Property

Public Property Get MaxValue() As Variant
Attribute MaxValue.VB_Description = "Gets the maximum element of a sorted set."
    RequireSortedSet "MaxValue"
    If mItemsCount = 0 Then RaiseEmptySequence "MaxValue"
    AssignPublicResult MaxValue, mItems(mItemsCount)
End Property

Public Function GetViewBetween( _
    ByVal lowerValue As Variant, _
    ByVal upperValue As Variant _
) As ROneCOne
Attribute GetViewBetween.VB_Description = "Gets a sorted-set range view."
    Dim lowerWrapped As ROneCOne
    Dim result As ROneCOne
    Dim upperWrapped As ROneCOne
    Dim wrapped As ROneCOne

    RequireSortedSet "GetViewBetween"
    Set lowerWrapped = WrapElement(lowerValue, "GetViewBetween")
    Set upperWrapped = WrapElement(upperValue, "GetViewBetween")
    Set result = New ROneCOne
    result.ConfigureGenericCollection ROLE_HASH_SET, COLLECTION_SORTED_SET, _
        mElementType, mElementObjectType, vbEmpty, vbNullString, mQueryComparer
    For Each wrapped In ArrSnapshot(mItems, mItemsCount)
        If CompareWithComparer(wrapped, lowerWrapped, mQueryComparer, _
            "GetViewBetween") >= 0 And _
            CompareWithComparer(wrapped, upperWrapped, mQueryComparer, _
            "GetViewBetween") <= 0 Then
            result.AddSetValue wrapped, "GetViewBetween"
        End If
    Next wrapped
    Set GetViewBetween = result
End Function

Public Property Get CollectionChanged() As ROneCOne
Attribute CollectionChanged.VB_Description = "Gets the observable collection change event."
    If mCollectionKind <> COLLECTION_OBSERVABLE Then
        RaiseContractError ERROR_INVALID_OPERATION, "CollectionChanged", _
            "CollectionChanged requires an ObservableCollection<T>."
    End If
    Set CollectionChanged = mCollectionChanged
End Property

Public Property Get IsReadOnly() As Boolean
Attribute IsReadOnly.VB_Description = "Reports whether mutating operations are rejected."
    IsReadOnly = (mRole = ROLE_IMMUTABLE Or _
        mCollectionKind = COLLECTION_READ_ONLY Or _
        mCollectionKind = COLLECTION_LOOKUP Or _
        mCollectionKind = COLLECTION_GROUPING)
End Property

Public Function ToBuilder() As ROneCOne
Attribute ToBuilder.VB_Description = "Creates a mutable builder for an immutable collection."
    Dim result As ROneCOne

    If mRole <> ROLE_IMMUTABLE Then
        RaiseContractError ERROR_INVALID_OPERATION, "ToBuilder", _
            "ToBuilder requires an immutable collection."
    End If
    Set result = CloneGenericCollection(Me)
    result.ConfigureCollectionRole MutableRoleForKind(mCollectionKind)
    Set ToBuilder = result
End Function

Public Function ToImmutable() As ROneCOne
Attribute ToImmutable.VB_Description = "Creates an immutable collection snapshot."
    Dim result As ROneCOne

    If Not IsGenericCollectionRole(mRole) Then
        RaiseContractError ERROR_INVALID_OPERATION, "ToImmutable", _
            "ToImmutable requires a generic collection."
    End If
    Set result = CloneGenericCollection(Me)
    result.ConfigureCollectionRole ROLE_IMMUTABLE
    Set ToImmutable = result
End Function

Public Sub UnionWith(ByVal other As Variant)
Attribute UnionWith.VB_Description = "Modifies a set to contain values from both sets."
    Dim incoming As Collection
    Dim wrapped As ROneCOne

    RequireSetCollection "UnionWith"
    RequireMutableCollection "UnionWith"
    Set incoming = BuildIncomingItems(other, "UnionWith")
    For Each wrapped In incoming
        If FindElementIndex(wrapped, "UnionWith") < 0 Then
            AddSetValue wrapped, "UnionWith"
        End If
    Next wrapped
End Sub

Public Sub IntersectWith(ByVal other As Variant)
Attribute IntersectWith.VB_Description = "Keeps only values found in both sets."
    Dim incoming As Collection
    Dim index As Long

    RequireSetCollection "IntersectWith"
    RequireMutableCollection "IntersectWith"
    Set incoming = BuildIncomingItems(other, "IntersectWith")
    ' This loop never probes the set between removals, so one deferred
    ' rebuild beats maintaining the index through every removal.
    MarkHashIndexDirty
    For index = mItemsCount To 1 Step -1
        If FindWrappedInCollection(incoming, mItems(index), _
            "IntersectWith") < 0 Then RemoveCollectionAt index - 1
    Next index
End Sub

Public Function Overlaps(ByVal other As Variant) As Boolean
Attribute Overlaps.VB_Description = "Tests whether two sets share at least one value."
    Dim incoming As Collection
    Dim wrapped As ROneCOne

    RequireSetCollection "Overlaps"
    Set incoming = BuildIncomingItems(other, "Overlaps")
    For Each wrapped In incoming
        If FindElementIndex(wrapped, "Overlaps") >= 0 Then
            Overlaps = True
            Exit Function
        End If
    Next wrapped
End Function

Public Function SetEquals(ByVal other As Variant) As Boolean
Attribute SetEquals.VB_Description = "Tests whether two sets contain the same values."
    Dim incoming As Collection
    Dim wrapped As ROneCOne

    RequireSetCollection "SetEquals"
    Set incoming = BuildIncomingItems(other, "SetEquals")
    If incoming.Count <> mItemsCount Then Exit Function
    For Each wrapped In incoming
        If FindElementIndex(wrapped, "SetEquals") < 0 Then Exit Function
    Next wrapped
    SetEquals = True
End Function

Public Function GetOrAdd( _
    ByVal key As Variant, _
    ByVal value As Variant _
) As Variant
Attribute GetOrAdd.VB_Description = "Gets an existing value or atomically adds one."
    Dim itemIndex As Long
    Dim keyWrapped As ROneCOne
    Dim selected As ROneCOne

    RequireMutableCollection "GetOrAdd"
    RequireKeyedCollection "GetOrAdd"
    Set keyWrapped = WrapKey(key, "GetOrAdd")
    If HashIndexSupported Then
        Set selected = FindHashValue(keyWrapped, "GetOrAdd")
    Else
        itemIndex = FindKeyIndex(keyWrapped, "GetOrAdd")
        If itemIndex >= 0 Then Set selected = mItems(itemIndex + 1)
    End If
    If selected Is Nothing Then
        Set selected = WrapElement(value, "GetOrAdd")
        AddDictionaryPair keyWrapped, selected, "GetOrAdd"
    End If
    AssignPublicResult GetOrAdd, selected
End Function

Public Function TryRemove( _
    ByVal key As Variant, _
    ByVal valueReference As ROneCOne _
) As Boolean
Attribute TryRemove.VB_Description = "Atomically removes a keyed value when present."
    Dim itemIndex As Long
    Dim keyWrapped As ROneCOne
    Dim selected As ROneCOne
    Dim slot As Long

    RequireMutableCollection "TryRemove"
    RequireKeyedCollection "TryRemove"
    RequireReference valueReference, "TryRemove"
    Set keyWrapped = WrapKey(key, "TryRemove")
    If HashIndexSupported Then
        slot = FindHashSlot(keyWrapped, "TryRemove")
        If slot < 0 Then Exit Function
        itemIndex = mHashSlots(slot) - 1
        Set selected = mHashValues(slot)
    Else
        itemIndex = FindKeyIndex(keyWrapped, "TryRemove")
        If itemIndex < 0 Then Exit Function
        Set selected = mItems(itemIndex + 1)
    End If
    AssignReferenceWrapped valueReference, selected, "TryRemove"
    RemoveCollectionAt itemIndex
    TryRemove = True
End Function

Public Function TryUpdate( _
    ByVal key As Variant, _
    ByVal newValue As Variant, _
    ByVal comparisonValue As Variant _
) As Boolean
Attribute TryUpdate.VB_Description = "Updates a concurrent value after comparison succeeds."
    Dim itemIndex As Long
    Dim expected As ROneCOne
    Dim keyWrapped As ROneCOne
    Dim selected As ROneCOne
    Dim slot As Long

    RequireConcurrentDictionary "TryUpdate"
    Set keyWrapped = WrapKey(key, "TryUpdate")
    If HashIndexSupported Then
        slot = FindHashSlot(keyWrapped, "TryUpdate")
        If slot < 0 Then Exit Function
        itemIndex = mHashSlots(slot) - 1
        Set selected = mHashValues(slot)
    Else
        itemIndex = FindKeyIndex(keyWrapped, "TryUpdate")
        If itemIndex < 0 Then Exit Function
        Set selected = mItems(itemIndex + 1)
    End If
    Set expected = WrapElement(comparisonValue, "TryUpdate")
    If Not ValuesEqual(selected, expected) Then Exit Function
    If HashIndexSupported Then
        ReplaceKeyedValueAtSlot itemIndex, slot, _
            WrapElement(newValue, "TryUpdate")
    Else
        ReplaceCollectionItem itemIndex, WrapElement(newValue, "TryUpdate")
    End If
    TryUpdate = True
End Function

Public Function AddOrUpdate( _
    ByVal key As Variant, _
    ByVal addValue As Variant, _
    ByVal updateValue As Variant _
) As Variant
Attribute AddOrUpdate.VB_Description = "Adds a concurrent value or replaces the existing value."
    Dim itemIndex As Long
    Dim keyWrapped As ROneCOne
    Dim selected As ROneCOne

    Dim slot As Long

    RequireConcurrentDictionary "AddOrUpdate"
    Set keyWrapped = WrapKey(key, "AddOrUpdate")
    If HashIndexSupported Then
        slot = FindHashSlot(keyWrapped, "AddOrUpdate")
        If slot < 0 Then
            Set selected = WrapElement(addValue, "AddOrUpdate")
            AddDictionaryPair keyWrapped, selected, "AddOrUpdate"
        Else
            Set selected = WrapElement(updateValue, "AddOrUpdate")
            ReplaceKeyedValueAtSlot mHashSlots(slot) - 1, slot, selected
        End If
        AssignPublicResult AddOrUpdate, selected
        Exit Function
    End If
    itemIndex = FindKeyIndex(keyWrapped, "AddOrUpdate")
    If itemIndex < 0 Then
        Set selected = WrapElement(addValue, "AddOrUpdate")
        AddDictionaryPair keyWrapped, selected, "AddOrUpdate"
    Else
        Set selected = WrapElement(updateValue, "AddOrUpdate")
        ReplaceCollectionItem itemIndex, selected
    End If
    AssignPublicResult AddOrUpdate, selected
End Function

Public Sub ExceptWith(ByVal other As Variant)
Attribute ExceptWith.VB_Description = "Removes every value found in another sequence."
    Dim incoming As Collection
    Dim index As Long

    RequireSetCollection "ExceptWith"
    RequireMutableCollection "ExceptWith"
    Set incoming = BuildIncomingItems(other, "ExceptWith")
    ' This loop never probes the set between removals, so one deferred
    ' rebuild beats maintaining the index through every removal.
    MarkHashIndexDirty
    For index = mItemsCount To 1 Step -1
        If FindWrappedInCollection(incoming, mItems(index), _
            "ExceptWith") >= 0 Then RemoveCollectionAt index - 1
    Next index
End Sub

Public Sub SymmetricExceptWith(ByVal other As Variant)
Attribute SymmetricExceptWith.VB_Description = _
    "Keeps values present in exactly one of two sets."
    Dim incoming As Collection
    Dim itemIndex As Long
    Dim uniqueValues As Collection
    Dim wrapped As ROneCOne

    RequireSetCollection "SymmetricExceptWith"
    RequireMutableCollection "SymmetricExceptWith"
    Set incoming = BuildIncomingItems(other, "SymmetricExceptWith")
    Set uniqueValues = DistinctWrappedValues(incoming, "SymmetricExceptWith")
    For Each wrapped In uniqueValues
        itemIndex = FindElementIndex(wrapped, "SymmetricExceptWith")
        If itemIndex < 0 Then
            AddSetValue wrapped, "SymmetricExceptWith"
        Else
            RemoveCollectionAt itemIndex
        End If
    Next wrapped
End Sub

Public Function IsSubsetOf(ByVal other As Variant) As Boolean
Attribute IsSubsetOf.VB_Description = "Tests whether every set value occurs in another sequence."
    IsSubsetOf = IsSetSubset(other, False, "IsSubsetOf")
End Function

Public Function IsProperSubsetOf(ByVal other As Variant) As Boolean
Attribute IsProperSubsetOf.VB_Description = "Tests whether this set is a strict subset."
    IsProperSubsetOf = IsSetSubset(other, True, "IsProperSubsetOf")
End Function

Public Function IsSupersetOf(ByVal other As Variant) As Boolean
Attribute IsSupersetOf.VB_Description = "Tests whether this set contains another sequence."
    IsSupersetOf = IsSetSuperset(other, False, "IsSupersetOf")
End Function

Public Function IsProperSupersetOf(ByVal other As Variant) As Boolean
Attribute IsProperSupersetOf.VB_Description = "Tests whether this set is a strict superset."
    IsProperSupersetOf = IsSetSuperset(other, True, "IsProperSupersetOf")
End Function

Public Function RemoveWhere(ByVal predicate As Variant) As Long
Attribute RemoveWhere.VB_Description = "Removes set values matching a predicate."
    Dim index As Long
    Dim predicateValue As ROneCOne

    RequireSetCollection "RemoveWhere"
    RequireMutableCollection "RemoveWhere"
    Set predicateValue = NormalizeContextUnary(predicate, "RemoveWhere", False)
    ' The predicate reads the element array, never the hash index, so one
    ' deferred rebuild beats maintaining the index through every removal.
    MarkHashIndexDirty
    For index = mItemsCount To 1 Step -1
        If InvokePredicate(predicateValue, mItems(index), "RemoveWhere") Then
            RemoveCollectionAt index - 1
            RemoveWhere = RemoveWhere + 1
        End If
    Next index
End Function

Public Function TryAddItem(ByVal value As Variant) As Boolean
Attribute TryAddItem.VB_Description = "Attempts to add to a blocking collection."
    RequireBlockingCollection "TryAddItem"
    If mAddingCompleted Then Exit Function
    If mQueryCount >= 0 And mItemsCount >= mQueryCount Then Exit Function
    AddMutable value, , "TryAddItem"
    TryAddItem = True
End Function

Public Sub CompleteAdding()
Attribute CompleteAdding.VB_Description = "Prevents further blocking-collection additions."
    RequireBlockingCollection "CompleteAdding"
    mAddingCompleted = True
End Sub

Public Property Get IsAddingCompleted() As Boolean
Attribute IsAddingCompleted.VB_Description = "Reports whether adding has completed."
    RequireBlockingCollection "IsAddingCompleted"
    IsAddingCompleted = mAddingCompleted
End Property

Public Function TakeItem() As Variant
Attribute TakeItem.VB_Description = "Removes the next blocking-collection value."
    RequireBlockingCollection "TakeItem"
    If mItemsCount = 0 Then RaiseEmptySequence "TakeItem"
    AssignPublicResult TakeItem, mItems(1)
    RemoveCollectionAt 0
End Function

Public Function TryTake(ByVal valueReference As ROneCOne) As Boolean
Attribute TryTake.VB_Description = "Attempts to remove a blocking-collection value."
    RequireBlockingCollection "TryTake"
    RequireReference valueReference, "TryTake"
    If mItemsCount = 0 Then Exit Function
    AssignReferenceWrapped valueReference, mItems(1), "TryTake"
    RemoveCollectionAt 0
    TryTake = True
End Function

Public Function TryPeekPair( _
    ByVal elementReference As ROneCOne, _
    ByVal priorityReference As ROneCOne _
) As Boolean
Attribute TryPeekPair.VB_Description = "Reads the next priority-queue element and priority."
    RequirePriorityQueue "TryPeekPair"
    RequireReference elementReference, "TryPeekPair"
    RequireReference priorityReference, "TryPeekPair"
    If mItemsCount = 0 Then Exit Function
    AssignReferenceWrapped elementReference, mItems(1), "TryPeekPair"
    AssignReferenceWrapped priorityReference, mPriorities(1), "TryPeekPair"
    TryPeekPair = True
End Function

Public Function TryDequeuePair( _
    ByVal elementReference As ROneCOne, _
    ByVal priorityReference As ROneCOne _
) As Boolean
Attribute TryDequeuePair.VB_Description = _
    "Removes the next priority-queue element and priority."
    If Not TryPeekPair(elementReference, priorityReference) Then Exit Function
    RequireMutableCollection "TryDequeuePair"
    RemoveCollectionAt 0
    TryDequeuePair = True
End Function

Public Function DequeueEnqueue( _
    ByVal element As Variant, _
    ByVal priority As Variant _
) As Variant
Attribute DequeueEnqueue.VB_Description = _
    "Dequeues the minimum and then enqueues a replacement."
    Dim removed As ROneCOne

    RequirePriorityQueue "DequeueEnqueue"
    RequireMutableCollection "DequeueEnqueue"
    If mItemsCount = 0 Then RaiseEmptySequence "DequeueEnqueue"
    Set removed = mItems(1)
    RemoveCollectionAt 0
    AddPriorityValue WrapElement(element, "DequeueEnqueue"), _
        WrapKey(priority, "DequeueEnqueue")
    AssignPublicResult DequeueEnqueue, removed
End Function

Public Function EnqueueDequeue( _
    ByVal element As Variant, _
    ByVal priority As Variant _
) As Variant
Attribute EnqueueDequeue.VB_Description = _
    "Enqueues a value and immediately dequeues the minimum."
    Dim removed As ROneCOne

    RequirePriorityQueue "EnqueueDequeue"
    RequireMutableCollection "EnqueueDequeue"
    AddPriorityValue WrapElement(element, "EnqueueDequeue"), _
        WrapKey(priority, "EnqueueDequeue")
    Set removed = mItems(1)
    RemoveCollectionAt 0
    AssignPublicResult EnqueueDequeue, removed
End Function

Public Property Get UnorderedItems() As ROneCOne
Attribute UnorderedItems.VB_Description = "Gets priority-queue element-priority pairs."
    Dim index As Long
    Dim result As ROneCOne

    RequirePriorityQueue "UnorderedItems"
    Set result = ROneCOne.ListOf(vbObject)
    For index = 1 To mItemsCount
        result.Add CreateDictionaryEntry( _
            mItems(index), mPriorities(index))
    Next index
    Set UnorderedItems = result
End Property

Public Function AddBefore( _
    ByVal node As ROneCOne, _
    ByVal value As Variant _
) As ROneCOne
Attribute AddBefore.VB_Description = "Adds a value before a linked-list node."
    Set AddBefore = InsertRelativeNode(node, value, False, "AddBefore")
End Function

Public Function AddAfter( _
    ByVal node As ROneCOne, _
    ByVal value As Variant _
) As ROneCOne
Attribute AddAfter.VB_Description = "Adds a value after a linked-list node."
    Set AddAfter = InsertRelativeNode(node, value, True, "AddAfter")
End Function

Public Property Get Previous() As ROneCOne
Attribute Previous.VB_Description = "Gets the previous linked-list node."
    Dim index As Long

    RequireCollectionNode "Previous"
    index = mSource.FindNodeWrappedIndex(mQueryValue)
    If index > 0 Then Set Previous = mSource.CreateNodeAt(index - 1)
End Property

Public Property Get NextNode() As ROneCOne
Attribute NextNode.VB_Description = "Gets the next linked-list node."
    Dim index As Long

    RequireCollectionNode "NextNode"
    index = mSource.FindNodeWrappedIndex(mQueryValue)
    If index >= 0 And index + 1 < mSource.Count Then
        Set NextNode = mSource.CreateNodeAt(index + 1)
    End If
End Property

Public Property Get List() As ROneCOne
Attribute List.VB_Description = "Gets the linked list that owns this node."
    RequireCollectionNode "List"
    Set List = mSource
End Property

Public Function FindNode(ByVal value As Variant) As ROneCOne
Attribute FindNode.VB_Description = "Finds the first linked-list node with a value."
    Dim index As Long

    RequireLinkedList "FindNode"
    index = FindElementIndex(WrapElement(value, "FindNode"), "FindNode")
    If index >= 0 Then Set FindNode = CreateCollectionNode(index)
End Function

Public Function FindLastNode(ByVal value As Variant) As ROneCOne
Attribute FindLastNode.VB_Description = "Finds the last linked-list node with a value."
    Dim candidate As ROneCOne
    Dim index As Long

    RequireLinkedList "FindLastNode"
    Set candidate = WrapElement(value, "FindLastNode")
    For index = mItemsCount To 1 Step -1
        If ValuesEqualWithComparer(mItems(index), candidate, _
            mQueryComparer, "FindLastNode") Then
            Set FindLastNode = CreateCollectionNode(index - 1)
            Exit Function
        End If
    Next index
End Function

Public Sub RemoveNode(ByVal node As ROneCOne)
Attribute RemoveNode.VB_Description = "Removes an owned linked-list node."
    Dim index As Long

    RequireLinkedList "RemoveNode"
    RequireMutableCollection "RemoveNode"
    RequireOwnedNode node, "RemoveNode"
    index = FindNodeWrappedIndex(node.InternalQueryValue)
    If index < 0 Then
        RaiseContractError ERROR_INVALID_OPERATION, "RemoveNode", _
            "The node no longer belongs to this list."
    End If
    RemoveCollectionAt index
End Sub

Public Function AddRangeImmutable(ByVal values As Variant) As ROneCOne
Attribute AddRangeImmutable.VB_Description = "Returns an immutable collection with added values."
    Dim incoming As Collection
    Dim result As ROneCOne
    Dim wrapped As ROneCOne

    RequireImmutableCollection "AddRangeImmutable"
    Set result = CloneGenericCollection(Me)
    result.ConfigureCollectionRole MutableRoleForKind(mCollectionKind)
    Set incoming = BuildIncomingItems(values, "AddRangeImmutable")
    For Each wrapped In incoming
        result.AddMutable Unwrap(wrapped), memberName:="AddRangeImmutable"
    Next wrapped
    result.ConfigureCollectionRole ROLE_IMMUTABLE
    Set AddRangeImmutable = result
End Function

Public Function SetItem( _
    ByVal keyOrIndex As Variant, _
    ByVal value As Variant _
) As ROneCOne
Attribute SetItem.VB_Description = "Returns an immutable collection with one value replaced."
    Dim itemIndex As Long
    Dim keyWrapped As ROneCOne
    Dim result As ROneCOne

    RequireImmutableCollection "SetItem"
    Set result = CloneGenericCollection(Me)
    result.ConfigureCollectionRole MutableRoleForKind(mCollectionKind)
    If result.IsDictionaryCollection Then
        Set keyWrapped = result.WrapKey(keyOrIndex, "SetItem")
        itemIndex = result.FindKeyIndex(keyWrapped, "SetItem")
        If itemIndex < 0 Then
            result.AddDictionaryPair keyWrapped, _
                result.WrapElement(value, "SetItem"), "SetItem"
        Else
            result.ReplaceCollectionItem itemIndex, _
                result.WrapElement(value, "SetItem")
        End If
    Else
        itemIndex = CLng(keyOrIndex)
        ValidateIndex itemIndex, result.Count, "SetItem"
        result.ReplaceCollectionItem itemIndex, _
            result.WrapElement(value, "SetItem")
    End If
    result.ConfigureCollectionRole ROLE_IMMUTABLE
    Set SetItem = result
End Function

Public Function RemoveRange(ByVal values As Variant) As ROneCOne
Attribute RemoveRange.VB_Description = "Returns an immutable collection without selected values."
    Dim incoming As Collection
    Dim result As ROneCOne
    Dim wrapped As ROneCOne

    RequireImmutableCollection "RemoveRange"
    Set result = CloneGenericCollection(Me)
    result.ConfigureCollectionRole MutableRoleForKind(mCollectionKind)
    Set incoming = BuildIncomingItems(values, "RemoveRange")
    For Each wrapped In incoming
        RemoveCollectionValue result, Unwrap(wrapped), "RemoveRange"
    Next wrapped
    result.ConfigureCollectionRole ROLE_IMMUTABLE
    Set RemoveRange = result
End Function

Public Function DequeueImmutable() As ROneCOne
Attribute DequeueImmutable.VB_Description = "Returns an immutable queue without its first value."
    Dim result As ROneCOne

    RequireImmutableCollection "DequeueImmutable"
    If Not IsQueueKind(mCollectionKind) Then
        RaiseContractError ERROR_INVALID_OPERATION, "DequeueImmutable", _
            "DequeueImmutable requires an immutable queue."
    End If
    If mItemsCount = 0 Then RaiseEmptySequence "DequeueImmutable"
    Set result = CloneGenericCollection(Me)
    result.ConfigureCollectionRole MutableRoleForKind(mCollectionKind)
    result.RemoveCollectionAt 0
    result.ConfigureCollectionRole ROLE_IMMUTABLE
    Set DequeueImmutable = result
End Function

Public Function PopImmutable() As ROneCOne
Attribute PopImmutable.VB_Description = "Returns an immutable stack without its top value."
    Dim result As ROneCOne

    RequireImmutableCollection "PopImmutable"
    If Not IsStackKind(mCollectionKind) Then
        RaiseContractError ERROR_INVALID_OPERATION, "PopImmutable", _
            "PopImmutable requires an immutable stack."
    End If
    If mItemsCount = 0 Then RaiseEmptySequence "PopImmutable"
    Set result = CloneGenericCollection(Me)
    result.ConfigureCollectionRole MutableRoleForKind(mCollectionKind)
    result.RemoveCollectionAt 0
    result.ConfigureCollectionRole ROLE_IMMUTABLE
    Set PopImmutable = result
End Function

Public Property Get Element() As ROneCOne
Attribute Element.VB_Description = "Gets a typed parameter representing each sequence element."
    If mRole = ROLE_QUERYABLE Then
        If mElementParameter Is Nothing Then
            Set mElementParameter = New ROneCOne
            mElementParameter.ConfigureParameter vbVariant, vbNullString
        End If
        Set Element = mElementParameter
        Exit Property
    End If
    RequireSequence Me, "Element"
    If mElementParameter Is Nothing Then
        Set mElementParameter = New ROneCOne
        mElementParameter.ConfigureParameter mElementType, mElementObjectType
    End If
    Set Element = mElementParameter
End Property

Public Function AsEnumerable() As ROneCOne
Attribute AsEnumerable.VB_Description = "Returns this sequence through its enumerable surface."
    RequireSequence Me, "AsEnumerable"
    Set AsEnumerable = Me
End Function

Public Function Condition( _
    ByVal memberPath As String, _
    Optional ByVal returnsObject As Boolean = False _
) As ROneCOne
Attribute Condition.VB_Description = "Creates a reusable sequence-member expression."
    If mRole = ROLE_QUERYABLE Then
        Set Condition = BuildMemberPath(memberPath, returnsObject)
        Exit Function
    End If
    RequireSequence Me, "Condition"
    Set Condition = BuildMemberPath(memberPath, returnsObject)
End Function

Public Function Match( _
    ByVal memberPath As String, _
    Optional ByVal value As Variant _
) As ROneCOne
Attribute Match.VB_Description = "Builds an equality predicate, or finds the first regex match."
    Dim firstMatches As Object

    If mRole = ROLE_REGEX Then
        If Not IsMissing(value) Then
            RaiseContractError ERROR_INVALID_ARGUMENT, "Match", _
                "A regular expression Match takes only the input text."
        End If
        mProviderObject.Global = False
        Set firstMatches = mProviderObject.Execute(memberPath)
        If firstMatches.Count = 0 Then
            Set Match = BuildFailedRegexMatch()
        Else
            Set Match = BuildRegexMatch(firstMatches.Item(0))
        End If
        Exit Function
    End If
    If IsMissing(value) Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "Match", _
            "Match requires a comparison value."
    End If
    Set Match = Condition(memberPath).EqualTo(value)
End Function

Public Function NotMatch( _
    ByVal memberPath As String, _
    ByVal value As Variant _
) As ROneCOne
Attribute NotMatch.VB_Description = "Creates a reusable member-inequality predicate expression."
    Set NotMatch = Condition(memberPath).NotEqualTo(value)
End Function

Public Function Always() As ROneCOne
Attribute Always.VB_Description = "Creates a reusable predicate that always returns True."
    Set Always = ROneCOne.Lambda(ROneCOne.Value(True), Element)
End Function

Public Function Never() As ROneCOne
Attribute Never.VB_Description = "Creates a reusable predicate that always returns False."
    Set Never = ROneCOne.Lambda(ROneCOne.Value(False), Element)
End Function

Public Function EqualityComparer( _
    ByVal callable As Variant, _
    Optional ByVal memberName As String = vbNullString _
) As ROneCOne
Attribute EqualityComparer.VB_Description = "Creates a typed Func<T, T, Boolean> equality comparer."
    Set EqualityComparer = CreateSequenceComparer( _
        callable, memberName, True, "EqualityComparer")
End Function

Public Function Comparer( _
    ByVal callable As Variant, _
    Optional ByVal memberName As String = vbNullString _
) As ROneCOne
Attribute Comparer.VB_Description = "Creates a typed Func<T, T, Long> ordering comparer."
    Set Comparer = CreateSequenceComparer( _
        callable, memberName, False, "Comparer")
End Function

Public Function Predicate( _
    ByVal callable As Variant, _
    Optional ByVal memberName As String = vbNullString _
) As ROneCOne
Attribute Predicate.VB_Description = "Infers a typed Func<T, Boolean> for this sequence."
    Dim predicateDelegate As ROneCOne

    RequireSequence Me, "Predicate"
    Set predicateDelegate = ROneCOne.Func(callable, memberName)
    Set Predicate = BindContextPredicate(predicateDelegate)
End Function

Public Function WhereMethod( _
    ByVal callable As Variant, _
    Optional ByVal memberName As String = vbNullString _
) As ROneCOne
Attribute WhereMethod.VB_Description = "Filters with an inferred method or procedure predicate."
    Set WhereMethod = Where(Predicate(callable, memberName))
End Function

Public Function Where(Optional ByVal predicateOrMember As Variant) As ROneCOne
Attribute Where.VB_Description = "Defers filtering or begins a fluent member condition."
    Dim builder As ROneCOne
    Dim predicateDelegate As ROneCOne

    If mRole = ROLE_QUERYABLE Then
        If IsMissing(predicateOrMember) Then
            Set builder = New ROneCOne
            builder.ConfigureWhereBuilder Me, vbNullString
            Set Where = builder
            Exit Function
        End If
        If Not IsObject(predicateOrMember) Then
            If VarTypeOf(predicateOrMember) = vbString Then
                Set builder = New ROneCOne
                builder.ConfigureWhereBuilder Me, CStr(predicateOrMember)
                Set Where = builder
                Exit Function
            End If
        End If
        Set Where = QueryableWithPredicate( _
            RequireROneCOne(predicateOrMember, "Where"))
        Exit Function
    End If
    RequireSequence Me, "Where"
    If IsMissing(predicateOrMember) Then
        Set builder = New ROneCOne
        builder.ConfigureWhereBuilder Me, vbNullString
        Set Where = builder
        Exit Function
    End If
    If Not IsObject(predicateOrMember) Then
        If VarTypeOf(predicateOrMember) = vbString Then
            Set builder = New ROneCOne
            builder.ConfigureWhereBuilder Me, CStr(predicateOrMember)
            Set Where = builder
            Exit Function
        End If
    End If
    Set predicateDelegate = NormalizeUnary( _
        RequireROneCOne(predicateOrMember, "Where"), "Where")
    Set Where = CreateQuery(QUERY_WHERE, predicateDelegate, Nothing, 0, _
        mElementType, mElementObjectType)
End Function

Public Function WhereNot(ByVal predicate As Variant) As ROneCOne
Attribute WhereNot.VB_Description = "Filters out elements matching a predicate."
    Dim predicateValue As ROneCOne

    Set predicateValue = NormalizeContextUnary(predicate, "WhereNot", False)
    Set WhereNot = CreateQuery( _
        QUERY_WHERE_NOT, predicateValue, Nothing, 0, _
        mElementType, mElementObjectType)
End Function

Public Function WhereAny( _
    ByVal collectionPath As String, _
    Optional ByVal predicate As Variant _
) As ROneCOne
Attribute WhereAny.VB_Description = "Filters when a nested sequence has any matching element."
    Dim expression As ROneCOne

    Set expression = Condition(collectionPath, True)
    If IsMissing(predicate) Then
        Set WhereAny = Where(expression.AnyMatch)
    Else
        Set WhereAny = Where(expression.AnyMatch(predicate))
    End If
End Function

Public Function WhereAll( _
    ByVal collectionPath As String, _
    ByVal predicate As Variant _
) As ROneCOne
Attribute WhereAll.VB_Description = "Filters when all nested-sequence elements match."
    Set WhereAll = Where( _
        Condition(collectionPath, True).AllMatch(predicate))
End Function

Public Function WhereNone( _
    ByVal collectionPath As String, _
    Optional ByVal predicate As Variant _
) As ROneCOne
Attribute WhereNone.VB_Description = "Filters when no nested-sequence element matches."
    Dim expression As ROneCOne

    Set expression = Condition(collectionPath, True)
    If IsMissing(predicate) Then
        Set WhereNone = Where(expression.NoneMatch)
    Else
        Set WhereNone = Where(expression.NoneMatch(predicate))
    End If
End Function

Public Function SelectItems( _
    ByVal selector As Variant, _
    ByVal resultTypeToken As Variant _
) As ROneCOne
Attribute SelectItems.VB_Description = "Defers a strongly typed LINQ Select projection."
    Dim objectType As String
    Dim selectorDelegate As ROneCOne
    Dim valueType As VbVarType

    ResolveTypeDescriptor resultTypeToken, valueType, objectType, "SelectItems"
    Set selectorDelegate = NormalizeContextUnary( _
        selector, "SelectItems", valueType = vbObject)
    Set SelectItems = CreateQuery(QUERY_SELECT, selectorDelegate, Nothing, 0, _
        valueType, objectType)
End Function

Public Function Map( _
    ByVal selector As Variant, _
    Optional ByVal resultTypeToken As Variant _
) As ROneCOne
Attribute Map.VB_Description = "Concise sugar for a strongly typed SelectItems projection."
    If IsMissing(resultTypeToken) Then
        Set Map = SelectItems(selector, vbVariant)
    Else
        Set Map = SelectItems(selector, resultTypeToken)
    End If
End Function

Public Function Take(ByVal count As Long) As ROneCOne
Attribute Take.VB_Description = "Defers taking up to a number of leading elements."
    Dim clone As ROneCOne

    If mRole = ROLE_QUERYABLE Then
        If count < 1 Then
            RaiseQueryError "Take requires a count of at least one."
        End If
        Set clone = QueryableClone()
        clone.InternalQueryableSetTake count
        Set Take = clone
        Exit Function
    End If
    If count < 0 Then count = 0
    Set Take = CreateQuery(QUERY_TAKE, Nothing, Nothing, count, _
        mElementType, mElementObjectType)
End Function

Public Function Skip(ByVal count As Long) As ROneCOne
Attribute Skip.VB_Description = "Defers bypassing a number of leading elements."
    Dim clone As ROneCOne

    If mRole = ROLE_QUERYABLE Then
        If mSqlQueryDialect = QUERY_DIALECT_ACE Then
            RaiseQueryError "The ACE dialect has no OFFSET clause, " & _
                "so Skip cannot translate for this connection."
        End If
        If count < 0 Then
            RaiseQueryError "Skip requires a count of zero or more."
        End If
        Set clone = QueryableClone()
        clone.InternalQueryableSetSkip count
        Set Skip = clone
        Exit Function
    End If
    If count < 0 Then count = 0
    Set Skip = CreateQuery(QUERY_SKIP, Nothing, Nothing, count, _
        mElementType, mElementObjectType)
End Function

Public Function TakeWhile(ByVal predicate As Variant) As ROneCOne
Attribute TakeWhile.VB_Description = "Takes elements while a predicate remains true."
    Set TakeWhile = CreateQuery( _
        QUERY_TAKE_WHILE, NormalizeContextUnary(predicate, "TakeWhile", False), _
        Nothing, 0, InternalElementType, InternalElementObjectType)
End Function

Public Function SkipWhile(ByVal predicate As Variant) As ROneCOne
Attribute SkipWhile.VB_Description = "Skips elements while a predicate remains true."
    Set SkipWhile = CreateQuery( _
        QUERY_SKIP_WHILE, NormalizeContextUnary(predicate, "SkipWhile", False), _
        Nothing, 0, InternalElementType, InternalElementObjectType)
End Function

Public Function TakeLast(ByVal count As Long) As ROneCOne
Attribute TakeLast.VB_Description = "Takes the final count elements."
    If count < 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "TakeLast", _
            "Count cannot be negative."
    End If
    Set TakeLast = CreateQuery(QUERY_TAKE_LAST, Nothing, Nothing, count, _
        InternalElementType, InternalElementObjectType)
End Function

Public Function SkipLast(ByVal count As Long) As ROneCOne
Attribute SkipLast.VB_Description = "Skips the final count elements."
    If count < 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "SkipLast", _
            "Count cannot be negative."
    End If
    Set SkipLast = CreateQuery(QUERY_SKIP_LAST, Nothing, Nothing, count, _
        InternalElementType, InternalElementObjectType)
End Function

Public Function ConcatSequence(ByVal other As Variant) As ROneCOne
Attribute ConcatSequence.VB_Description = "Concatenates two compatible sequences."
    Dim otherSequence As ROneCOne

    Set otherSequence = NormalizeMembershipSequence(other, "Concat")
    RequireCompatibleSequence otherSequence, "Concat"
    Set ConcatSequence = CreateQuery(QUERY_CONCAT, Nothing, otherSequence, 0, _
        InternalElementType, InternalElementObjectType)
End Function

Public Function Union( _
    ByVal other As Variant, _
    Optional ByVal comparer As Variant _
) As ROneCOne
Attribute Union.VB_Description = "Produces the set union of two sequences."
    Set Union = CreateSetQuery(QUERY_UNION, other, comparer, "Union")
End Function

Public Function Intersect( _
    ByVal other As Variant, _
    Optional ByVal comparer As Variant _
) As ROneCOne
Attribute Intersect.VB_Description = "Produces the set intersection of two sequences."
    Set Intersect = CreateSetQuery(QUERY_INTERSECT, other, comparer, "Intersect")
End Function

Public Function Except( _
    ByVal other As Variant, _
    Optional ByVal comparer As Variant _
) As ROneCOne
Attribute Except.VB_Description = "Produces values absent from the second sequence."
    Set Except = CreateSetQuery(QUERY_EXCEPT, other, comparer, "Except")
End Function

Public Function DefaultIfEmpty(Optional ByVal defaultValue As Variant) As ROneCOne
Attribute DefaultIfEmpty.VB_Description = "Supplies one default when a sequence is empty."
    Dim wrappedDefault As ROneCOne

    If Not IsMissing(defaultValue) Then
        Set wrappedDefault = WrapElement(defaultValue, "DefaultIfEmpty")
    End If
    Set DefaultIfEmpty = CreateQuery( _
        QUERY_DEFAULT_IF_EMPTY, Nothing, wrappedDefault, 0, _
        InternalElementType, InternalElementObjectType)
End Function

Public Function Chunk(ByVal size As Long) As ROneCOne
Attribute Chunk.VB_Description = "Splits a sequence into typed List<T> chunks."
    If size <= 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "Chunk", _
            "Chunk size must be positive."
    End If
    Set Chunk = CreateQuery( _
        QUERY_CHUNK, Nothing, Nothing, size, vbObject, "ROneCOne")
End Function

Public Function SelectMany( _
    ByVal selector As Variant, _
    Optional ByVal resultType As Variant _
) As ROneCOne
Attribute SelectMany.VB_Description = "Flattens nested sequences through a selector."
    Dim objectType As String
    Dim valueType As VbVarType

    If IsMissing(resultType) Then
        valueType = vbVariant
    Else
        ResolveTypeDescriptor resultType, valueType, objectType, "SelectMany"
    End If
    Set SelectMany = CreateQuery( _
        QUERY_SELECT_MANY, NormalizeContextUnary(selector, "SelectMany", False), _
        Nothing, 0, valueType, objectType)
End Function

Public Function Cast(ByVal typeToken As Variant) As ROneCOne
Attribute Cast.VB_Description = "Casts every sequence value to the requested runtime type."
    Dim objectType As String
    Dim valueType As VbVarType

    ResolveTypeDescriptor typeToken, valueType, objectType, "Cast"
    Set Cast = CreateQuery( _
        QUERY_CAST, Nothing, Nothing, 0, valueType, objectType)
End Function

Public Function OfType(ByVal typeToken As Variant) As ROneCOne
Attribute OfType.VB_Description = "Filters values assignable to the requested runtime type."
    Dim objectType As String
    Dim valueType As VbVarType

    ResolveTypeDescriptor typeToken, valueType, objectType, "OfType"
    Set OfType = CreateQuery( _
        QUERY_OF_TYPE, Nothing, Nothing, 0, valueType, objectType)
End Function

Public Function UnionBy( _
    ByVal other As Variant, _
    ByVal keySelector As Variant, _
    Optional ByVal comparer As Variant _
) As ROneCOne
Attribute UnionBy.VB_Description = "Produces a union using selected distinct keys."
    Set UnionBy = CreateSetByQuery( _
        QUERY_UNION_BY, other, keySelector, comparer, "UnionBy", True)
End Function

Public Function IntersectBy( _
    ByVal keys As Variant, _
    ByVal keySelector As Variant, _
    Optional ByVal comparer As Variant _
) As ROneCOne
Attribute IntersectBy.VB_Description = "Keeps source values whose selected key is present."
    Set IntersectBy = CreateSetByQuery( _
        QUERY_INTERSECT_BY, keys, keySelector, comparer, _
        "IntersectBy", False)
End Function

Public Function ExceptBy( _
    ByVal keys As Variant, _
    ByVal keySelector As Variant, _
    Optional ByVal comparer As Variant _
) As ROneCOne
Attribute ExceptBy.VB_Description = "Keeps source values whose selected key is absent."
    Set ExceptBy = CreateSetByQuery( _
        QUERY_EXCEPT_BY, keys, keySelector, comparer, "ExceptBy", False)
End Function

Public Function Distinct(Optional ByVal comparer As Variant) As ROneCOne
Attribute Distinct.VB_Description = "Defers removal of duplicate sequence values."
    Dim comparerValue As ROneCOne

    If Not IsMissing(comparer) Then
        Set comparerValue = NormalizeComparer(comparer, True, "Distinct")
    End If
    Set Distinct = CreateQuery(QUERY_DISTINCT, Nothing, Nothing, 0, _
        mElementType, mElementObjectType, comparerValue)
End Function

Public Function DistinctBy( _
    ByVal keySelector As Variant, _
    Optional ByVal comparer As Variant _
) As ROneCOne
Attribute DistinctBy.VB_Description = "Defers removal of elements with duplicate selected keys."
    Dim comparerValue As ROneCOne
    Dim selectorDelegate As ROneCOne

    Set selectorDelegate = NormalizeContextUnary( _
        keySelector, "DistinctBy", False)
    If Not IsMissing(comparer) Then
        Set comparerValue = NormalizeComparer(comparer, True, "DistinctBy")
    End If
    Set DistinctBy = CreateQuery(QUERY_DISTINCT_BY, selectorDelegate, _
        Nothing, 0, mElementType, mElementObjectType, comparerValue)
End Function

Public Function OrderBy( _
    ByVal keySelector As Variant, _
    Optional ByVal comparer As Variant _
) As ROneCOne
Attribute OrderBy.VB_Description = "Defers stable ascending ordering by a selected key."
    Dim comparerValue As ROneCOne
    Dim selectorDelegate As ROneCOne

    If mRole = ROLE_QUERYABLE Then
        QueryableRefuseComparer comparer, "OrderBy"
        Set OrderBy = QueryableOrdered(keySelector, False, False, "OrderBy")
        Exit Function
    End If
    Set selectorDelegate = NormalizeContextUnary( _
        keySelector, "OrderBy", False)
    If Not IsMissing(comparer) Then
        Set comparerValue = NormalizeComparer(comparer, False, "OrderBy")
    End If
    Set OrderBy = CreatePrimaryOrderQuery( _
        QUERY_ORDER_BY, selectorDelegate, _
        mElementType, mElementObjectType, comparerValue)
End Function

Public Function OrderByDescending( _
    ByVal keySelector As Variant, _
    Optional ByVal comparer As Variant _
) As ROneCOne
Attribute OrderByDescending.VB_Description = "Defers stable descending ordering by a selected key."
    Dim comparerValue As ROneCOne
    Dim selectorDelegate As ROneCOne

    If mRole = ROLE_QUERYABLE Then
        QueryableRefuseComparer comparer, "OrderByDescending"
        Set OrderByDescending = QueryableOrdered( _
            keySelector, True, False, "OrderByDescending")
        Exit Function
    End If
    Set selectorDelegate = NormalizeContextUnary( _
        keySelector, "OrderByDescending", False)
    If Not IsMissing(comparer) Then
        Set comparerValue = NormalizeComparer( _
            comparer, False, "OrderByDescending")
    End If
    Set OrderByDescending = CreatePrimaryOrderQuery( _
        QUERY_ORDER_BY_DESCENDING, selectorDelegate, _
        mElementType, mElementObjectType, comparerValue)
End Function

Public Function Order(Optional ByVal comparer As Variant) As ROneCOne
Attribute Order.VB_Description = "Defers stable ascending ordering by each element."
    Dim comparerValue As ROneCOne
    Dim selectorDelegate As ROneCOne

    Set selectorDelegate = NormalizeContextUnary(Element, "Order", False)
    If IsMissing(comparer) Then
        Set comparerValue = Nothing
    Else
        Set comparerValue = NormalizeComparer(comparer, False, "Order")
    End If
    Set Order = CreatePrimaryOrderQuery( _
        QUERY_ORDER, selectorDelegate, _
        mElementType, mElementObjectType, comparerValue)
End Function

Public Function OrderDescending(Optional ByVal comparer As Variant) As ROneCOne
Attribute OrderDescending.VB_Description = "Defers stable descending ordering by each element."
    Dim comparerValue As ROneCOne
    Dim selectorDelegate As ROneCOne

    Set selectorDelegate = NormalizeContextUnary( _
        Element, "OrderDescending", False)
    If IsMissing(comparer) Then
        Set comparerValue = Nothing
    Else
        Set comparerValue = NormalizeComparer( _
            comparer, False, "OrderDescending")
    End If
    Set OrderDescending = CreatePrimaryOrderQuery( _
        QUERY_ORDER_DESCENDING, selectorDelegate, _
        mElementType, mElementObjectType, comparerValue)
End Function

Public Function ThenBy( _
    ByVal keySelector As Variant, _
    Optional ByVal comparer As Variant _
) As ROneCOne
Attribute ThenBy.VB_Description = "Adds a stable ascending key to an active ordered query."
    Dim comparerValue As ROneCOne
    Dim selectorDelegate As ROneCOne

    If mRole = ROLE_QUERYABLE Then
        QueryableRefuseComparer comparer, "ThenBy"
        Set ThenBy = QueryableOrdered(keySelector, False, True, "ThenBy")
        Exit Function
    End If
    RequireOrderedQuery Me, "ThenBy"
    Set selectorDelegate = NormalizeContextUnary( _
        keySelector, "ThenBy", False)
    If Not IsMissing(comparer) Then
        Set comparerValue = NormalizeComparer(comparer, False, "ThenBy")
    End If
    Set ThenBy = CreateQuery(QUERY_THEN_BY, selectorDelegate, Nothing, 0, _
        mElementType, mElementObjectType, comparerValue)
End Function

Public Function ThenByDescending( _
    ByVal keySelector As Variant, _
    Optional ByVal comparer As Variant _
) As ROneCOne
Attribute ThenByDescending.VB_Description = "Adds a stable descending key to an ordered query."
    Dim comparerValue As ROneCOne
    Dim selectorDelegate As ROneCOne

    If mRole = ROLE_QUERYABLE Then
        QueryableRefuseComparer comparer, "ThenByDescending"
        Set ThenByDescending = QueryableOrdered( _
            keySelector, True, True, "ThenByDescending")
        Exit Function
    End If
    RequireOrderedQuery Me, "ThenByDescending"
    Set selectorDelegate = NormalizeContextUnary( _
        keySelector, "ThenByDescending", False)
    If Not IsMissing(comparer) Then
        Set comparerValue = NormalizeComparer( _
            comparer, False, "ThenByDescending")
    End If
    Set ThenByDescending = CreateQuery( _
        QUERY_THEN_BY_DESCENDING, selectorDelegate, Nothing, 0, _
        mElementType, mElementObjectType, comparerValue)
End Function

Public Function Append(ByVal element As Variant) As ROneCOne
Attribute Append.VB_Description = "Appends builder text or defers appending to a sequence."
    Dim wrapped As ROneCOne

    If mRole = ROLE_STRING_BUILDER Then
        JsonBuilderAppend mStringBuilder, FormatArgText(element, vbNullString)
        Set Append = Me
        Exit Function
    End If
    RequireSequence Me, "Append"
    Set wrapped = WrapElement(element, "Append")
    Set Append = CreateQuery(QUERY_APPEND, Nothing, wrapped, 0, _
        mElementType, mElementObjectType)
End Function

Public Function Prepend(ByVal element As Variant) As ROneCOne
Attribute Prepend.VB_Description = "Defers prepending one typed element to a sequence."
    Dim wrapped As ROneCOne

    RequireSequence Me, "Prepend"
    Set wrapped = WrapElement(element, "Prepend")
    Set Prepend = CreateQuery(QUERY_PREPEND, Nothing, wrapped, 0, _
        mElementType, mElementObjectType)
End Function

Public Function Reverse() As ROneCOne
Attribute Reverse.VB_Description = "Defers reversing a sequence."
    Set Reverse = CreateQuery(QUERY_REVERSE, Nothing, Nothing, 0, _
        mElementType, mElementObjectType)
End Function

Public Function AnyMatch(Optional ByVal predicate As Variant) As ROneCOne
Attribute AnyMatch.VB_Description = "Tests whether a nested sequence has any matching element."
    If IsMissing(predicate) Then
        Set AnyMatch = CreateQuantifierCondition(OP_ANY_MATCH, Empty, False)
    Else
        Set AnyMatch = CreateQuantifierCondition( _
            OP_ANY_MATCH, predicate, True)
    End If
End Function

Public Function AllMatch(ByVal predicate As Variant) As ROneCOne
Attribute AllMatch.VB_Description = "Tests whether all nested-sequence elements match."
    Set AllMatch = CreateQuantifierCondition(OP_ALL_MATCH, predicate, True)
End Function

Public Function NoneMatch(Optional ByVal predicate As Variant) As ROneCOne
Attribute NoneMatch.VB_Description = "Tests whether no nested-sequence element matches."
    If IsMissing(predicate) Then
        Set NoneMatch = CreateQuantifierCondition(OP_NONE_MATCH, Empty, False)
    Else
        Set NoneMatch = CreateQuantifierCondition( _
            OP_NONE_MATCH, predicate, True)
    End If
End Function

Public Function AnyItem(Optional ByVal predicate As Variant) As Boolean
Attribute AnyItem.VB_Description = "Tests whether any element exists or matches a predicate."
    Dim filtered As ROneCOne
    Dim predicateValue As ROneCOne
    Dim values As Collection
    Dim wrapped As ROneCOne

    If mRole = ROLE_QUERYABLE Then
        If IsMissing(predicate) Then
            Set filtered = Me
        Else
            Set filtered = QueryableWithPredicate( _
                RequireROneCOne(predicate, "Any"))
        End If
        AnyItem = (filtered.Take(1).Count > 0)
        Exit Function
    End If
    Set values = MaterializeWrapped
    If IsMissing(predicate) Then
        AnyItem = (values.Count <> 0)
        Exit Function
    End If
    Set predicateValue = NormalizeContextUnary( _
        predicate, "Any", False)
    For Each wrapped In values
        If InvokePredicate(predicateValue, wrapped, "Any") Then
            AnyItem = True
            Exit Function
        End If
    Next wrapped
End Function

Public Function Exists(Optional ByVal predicate As Variant) As Boolean
Attribute Exists.VB_Description = "Tests whether an element matches or a file system path exists."
    If mRole = ROLE_FILE_FACTORY Then
        Exists = CreateFileSystemObject().FileExists( _
            RequirePathText(predicate, "Exists"))
        Exit Function
    End If
    If mRole = ROLE_DIRECTORY_FACTORY Then
        Exists = CreateFileSystemObject().FolderExists( _
            TrimTrailingSeparators(RequirePathText(predicate, "Exists")))
        Exit Function
    End If
    If IsMissing(predicate) Then
        Exists = AnyItem
    Else
        Exists = AnyItem(predicate)
    End If
End Function

Public Function None(Optional ByVal predicate As Variant) As Boolean
Attribute None.VB_Description = "Tests whether no element exists or matches a predicate."
    If IsMissing(predicate) Then
        None = Not AnyItem
    Else
        None = Not AnyItem(predicate)
    End If
End Function

Public Function All(ByVal predicate As Variant) As Boolean
Attribute All.VB_Description = "Tests whether every element matches a predicate."
    Dim predicateDelegate As ROneCOne
    Dim values As Collection
    Dim wrapped As ROneCOne

    Set predicateDelegate = NormalizeContextUnary(predicate, "All", False)
    Set values = MaterializeWrapped
    All = True
    For Each wrapped In values
        If Not InvokePredicate(predicateDelegate, wrapped, "All") Then
            All = False
            Exit Function
        End If
    Next wrapped
End Function

Public Function First(Optional ByVal predicate As Variant) As Variant
Attribute First.VB_Description = "Returns the first element, optionally matching a predicate."
    Dim predicateValue As ROneCOne
    Dim wrapped As ROneCOne

    If Not IsMissing(predicate) Then
        Set predicateValue = NormalizeContextUnary( _
            predicate, "First", False)
    End If
    Set wrapped = FindTerminalElement(predicateValue, False, "First")
    If wrapped.InternalHasObjectValue Then
        Set First = wrapped.InternalObjectValue
    Else
        First = wrapped.InternalScalarValue
    End If
End Function

Public Function Last(Optional ByVal predicate As Variant) As Variant
Attribute Last.VB_Description = "Returns the last element, optionally matching a predicate."
    Dim predicateValue As ROneCOne
    Dim wrapped As ROneCOne

    If Not IsMissing(predicate) Then
        Set predicateValue = NormalizeContextUnary( _
            predicate, "Last", False)
    End If
    Set wrapped = FindTerminalElement(predicateValue, True, "Last")
    If wrapped.InternalHasObjectValue Then
        Set Last = wrapped.InternalObjectValue
    Else
        Last = wrapped.InternalScalarValue
    End If
End Function

Public Function Sum(Optional ByVal selector As Variant) As Variant
Attribute Sum.VB_Description = "Sums numeric elements or selected numeric values."
    Dim values As Collection
    Dim wrapped As ROneCOne
    Dim selected As ROneCOne
    Dim selectorValue As ROneCOne
    Dim total As Variant

    total = 0
    If Not IsMissing(selector) Then
        Set selectorValue = NormalizeContextUnary( _
            selector, "Sum", False)
    End If
    Set values = MaterializeWrapped
    For Each wrapped In values
        Set selected = SelectTerminalValue(wrapped, selectorValue)
        total = total + RequireScalar(selected, "Sum")
    Next wrapped
    Sum = total
End Function

Public Function Average(Optional ByVal selector As Variant) As Double
Attribute Average.VB_Description = "Averages numeric elements or selected numeric values."
    Dim selected As ROneCOne
    Dim selectorValue As ROneCOne
    Dim total As Double
    Dim valueCount As Long
    Dim values As Collection
    Dim wrapped As ROneCOne

    If Not IsMissing(selector) Then
        Set selectorValue = NormalizeContextUnary( _
            selector, "Average", False)
    End If
    Set values = MaterializeWrapped
    If values.Count = 0 Then RaiseEmptySequence "Average"
    For Each wrapped In values
        Set selected = SelectTerminalValue(wrapped, selectorValue)
        total = total + CDbl(RequireScalar(selected, "Average"))
    Next wrapped
    Average = total / CDbl(values.Count)
End Function

Public Function Min( _
    Optional ByVal selector As Variant, _
    Optional ByVal comparer As Variant _
) As Variant
Attribute Min.VB_Description = "Returns the minimum element or selected value."
    Dim comparerValue As ROneCOne
    Dim selectorValue As ROneCOne
    Dim wrapped As ROneCOne

    If Not IsMissing(selector) Then
        Set selectorValue = NormalizeContextUnary( _
            selector, "Min", False)
    End If
    If Not IsMissing(comparer) Then
        Set comparerValue = NormalizeComparer(comparer, False, "Min")
    End If
    Set wrapped = ExtremeValue(selectorValue, comparerValue, False, "Min")
    If wrapped.InternalHasObjectValue Then
        Set Min = wrapped.InternalObjectValue
    Else
        Min = wrapped.InternalScalarValue
    End If
End Function

Public Function Max( _
    Optional ByVal selector As Variant, _
    Optional ByVal comparer As Variant _
) As Variant
Attribute Max.VB_Description = "Returns the maximum element or selected value."
    Dim comparerValue As ROneCOne
    Dim selectorValue As ROneCOne
    Dim wrapped As ROneCOne

    If Not IsMissing(selector) Then
        Set selectorValue = NormalizeContextUnary( _
            selector, "Max", False)
    End If
    If Not IsMissing(comparer) Then
        Set comparerValue = NormalizeComparer(comparer, False, "Max")
    End If
    Set wrapped = ExtremeValue(selectorValue, comparerValue, True, "Max")
    If wrapped.InternalHasObjectValue Then
        Set Max = wrapped.InternalObjectValue
    Else
        Max = wrapped.InternalScalarValue
    End If
End Function

Public Function FirstOrDefault(Optional ByVal predicate As Variant) As Variant
Attribute FirstOrDefault.VB_Description = "Returns the first match or default(T)."
    Dim filtered As ROneCOne
    Dim firstRows As ROneCOne
    Dim found As Boolean
    Dim predicateValue As ROneCOne
    Dim wrapped As ROneCOne

    If mRole = ROLE_QUERYABLE Then
        If IsMissing(predicate) Then
            Set filtered = Me
        Else
            Set filtered = QueryableWithPredicate( _
                RequireROneCOne(predicate, "FirstOrDefault"))
        End If
        Set firstRows = filtered.Take(1).ToDataTable
        If firstRows.Rows.Count = 0 Then
            Set FirstOrDefault = Nothing
        Else
            Set FirstOrDefault = firstRows.Rows.Item(0)
        End If
        Exit Function
    End If
    If Not IsMissing(predicate) Then
        Set predicateValue = NormalizeContextUnary( _
            predicate, "FirstOrDefault", False)
    End If
    Set wrapped = TryFindTerminalElement( _
        predicateValue, False, "FirstOrDefault", found)
    If found Then
        AssignPublicResult FirstOrDefault, wrapped
    Else
        AssignDefaultPublicResult FirstOrDefault
    End If
End Function

Public Function LastOrDefault(Optional ByVal predicate As Variant) As Variant
Attribute LastOrDefault.VB_Description = "Returns the last match or default(T)."
    Dim found As Boolean
    Dim predicateValue As ROneCOne
    Dim wrapped As ROneCOne

    If Not IsMissing(predicate) Then
        Set predicateValue = NormalizeContextUnary( _
            predicate, "LastOrDefault", False)
    End If
    Set wrapped = TryFindTerminalElement( _
        predicateValue, True, "LastOrDefault", found)
    If found Then
        AssignPublicResult LastOrDefault, wrapped
    Else
        AssignDefaultPublicResult LastOrDefault
    End If
End Function

Public Function SingleItem(Optional ByVal predicate As Variant) As Variant
Attribute SingleItem.VB_Description = "Returns the sole matching element or raises."
    Dim predicateValue As ROneCOne
    Dim wrapped As ROneCOne

    If Not IsMissing(predicate) Then
        Set predicateValue = NormalizeContextUnary(predicate, "Single", False)
    End If
    Set wrapped = FindSingleElement(predicateValue, False, "Single")
    AssignPublicResult SingleItem, wrapped
End Function

Public Function SingleOrDefault(Optional ByVal predicate As Variant) As Variant
Attribute SingleOrDefault.VB_Description = "Returns one match, default(T), or raises for many."
    Dim predicateValue As ROneCOne
    Dim wrapped As ROneCOne

    If Not IsMissing(predicate) Then
        Set predicateValue = NormalizeContextUnary( _
            predicate, "SingleOrDefault", False)
    End If
    Set wrapped = FindSingleElement( _
        predicateValue, True, "SingleOrDefault")
    If wrapped Is Nothing Then
        AssignDefaultPublicResult SingleOrDefault
    Else
        AssignPublicResult SingleOrDefault, wrapped
    End If
End Function

Public Function ElementAt(ByVal index As Long) As Variant
Attribute ElementAt.VB_Description = "Gets an element by zero-based index."
    Dim values As Collection

    Set values = MaterializeWrapped
    ValidateIndex index, values.Count, "ElementAt"
    AssignPublicResult ElementAt, values.Item(index + 1)
End Function

Public Function ElementAtOrDefault(ByVal index As Long) As Variant
Attribute ElementAtOrDefault.VB_Description = "Gets an element or the type default."
    Dim values As Collection

    Set values = MaterializeWrapped
    If index < 0 Or index >= values.Count Then
        AssignDefaultPublicResult ElementAtOrDefault
    Else
        AssignPublicResult ElementAtOrDefault, values.Item(index + 1)
    End If
End Function

Public Function Index() As ROneCOne
Attribute Index.VB_Description = "Pairs every element with its zero-based index."
    Dim indexValue As Long
    Dim result As ROneCOne
    Dim values As Collection

    Set result = ROneCOne.ListOf(vbObject)
    Set values = MaterializeWrapped
    For indexValue = 1 To values.Count
        result.Add CreateDictionaryEntry( _
            CreateScalarResult(CLng(indexValue - 1)), values.Item(indexValue))
    Next indexValue
    Set Index = result
End Function

Public Function Aggregate( _
    ByVal seedOrAccumulator As Variant, _
    Optional ByVal accumulator As Variant, _
    Optional ByVal resultSelector As Variant _
) As Variant
Attribute Aggregate.VB_Description = "Applies an accumulator across a sequence."
    Dim accumulatorValue As ROneCOne
    Dim current As ROneCOne
    Dim index As Long
    Dim resultSelectorValue As ROneCOne
    Dim values As Collection

    Set values = MaterializeWrapped
    If IsMissing(accumulator) Then
        If values.Count = 0 Then RaiseEmptySequence "Aggregate"
        Set accumulatorValue = NormalizeBinaryDelegate( _
            seedOrAccumulator, "Aggregate")
        Set current = values.Item(1)
        index = 2
    Else
        Set current = WrapInvocationArgument(seedOrAccumulator)
        Set accumulatorValue = NormalizeBinaryDelegate(accumulator, "Aggregate")
        index = 1
    End If
    For index = index To values.Count
        Set current = InvokeBinary(accumulatorValue, current, values.Item(index))
    Next index
    If Not IsMissing(resultSelector) Then
        Set resultSelectorValue = NormalizeStandaloneUnary( _
            RequireROneCOne(resultSelector, "Aggregate"), "Aggregate")
        Set current = InvokeUnary(resultSelectorValue, current)
    End If
    AssignPublicResult Aggregate, current
End Function

Public Function CountBy( _
    ByVal keySelector As Variant, _
    Optional ByVal comparer As Variant _
) As ROneCOne
Attribute CountBy.VB_Description = "Counts elements grouped by a selected key."
    Dim comparerValue As ROneCOne
    Dim itemIndex As Long
    Dim key As ROneCOne
    Dim result As ROneCOne
    Dim selectorValue As ROneCOne
    Dim values As Collection
    Dim wrapped As ROneCOne

    Set selectorValue = NormalizeContextUnary(keySelector, "CountBy", False)
    If Not IsMissing(comparer) Then
        Set comparerValue = NormalizeComparer(comparer, True, "CountBy")
    End If
    Set result = New ROneCOne
    result.ConfigureGenericCollection ROLE_DICTIONARY, COLLECTION_DICTIONARY, _
        vbLong, vbNullString, vbVariant, vbNullString, comparerValue
    Set values = MaterializeWrapped
    For Each wrapped In values
        Set key = InvokeUnary(selectorValue, wrapped)
        itemIndex = result.FindKeyIndex(key, "CountBy")
        If itemIndex < 0 Then
            result.AddDictionaryPair key, CreateScalarResult(CLng(1)), "CountBy"
        Else
            result.ReplaceCollectionItem itemIndex, CreateScalarResult( _
                CLng(result.InternalCollectionItem(itemIndex).InternalScalarValue) + 1)
        End If
    Next wrapped
    Set CountBy = result
End Function

Public Function AggregateBy( _
    ByVal keySelector As Variant, _
    ByVal seed As Variant, _
    ByVal accumulator As Variant, _
    Optional ByVal comparer As Variant _
) As ROneCOne
Attribute AggregateBy.VB_Description = "Aggregates elements independently for each key."
    Dim accumulatorValue As ROneCOne
    Dim comparerValue As ROneCOne
    Dim current As ROneCOne
    Dim itemIndex As Long
    Dim key As ROneCOne
    Dim result As ROneCOne
    Dim selectorValue As ROneCOne
    Dim values As Collection
    Dim wrapped As ROneCOne

    Set selectorValue = NormalizeContextUnary(keySelector, "AggregateBy", False)
    Set accumulatorValue = NormalizeBinaryDelegate(accumulator, "AggregateBy")
    If Not IsMissing(comparer) Then
        Set comparerValue = NormalizeComparer(comparer, True, "AggregateBy")
    End If
    Set result = New ROneCOne
    result.ConfigureGenericCollection ROLE_DICTIONARY, COLLECTION_DICTIONARY, _
        vbVariant, vbNullString, vbVariant, vbNullString, comparerValue
    Set values = MaterializeWrapped
    For Each wrapped In values
        Set key = InvokeUnary(selectorValue, wrapped)
        itemIndex = result.FindKeyIndex(key, "AggregateBy")
        If itemIndex < 0 Then
            Set current = WrapInvocationArgument(seed)
        Else
            Set current = result.InternalCollectionItem(itemIndex)
        End If
        Set current = InvokeBinary(accumulatorValue, current, wrapped)
        If itemIndex < 0 Then
            result.AddDictionaryPair key, current, "AggregateBy"
        Else
            result.ReplaceCollectionItem itemIndex, current
        End If
    Next wrapped
    Set AggregateBy = result
End Function

Public Function MinBy( _
    ByVal keySelector As Variant, _
    Optional ByVal comparer As Variant _
) As Variant
Attribute MinBy.VB_Description = "Returns the element with the minimum selected key."
    Dim comparerValue As ROneCOne
    Dim selectorValue As ROneCOne
    Dim wrapped As ROneCOne

    Set selectorValue = NormalizeContextUnary(keySelector, "MinBy", False)
    If Not IsMissing(comparer) Then
        Set comparerValue = NormalizeComparer(comparer, False, "MinBy")
    End If
    Set wrapped = ExtremeElement( _
        selectorValue, comparerValue, False, "MinBy")
    AssignPublicResult MinBy, wrapped
End Function

Public Function MaxBy( _
    ByVal keySelector As Variant, _
    Optional ByVal comparer As Variant _
) As Variant
Attribute MaxBy.VB_Description = "Returns the element with the maximum selected key."
    Dim comparerValue As ROneCOne
    Dim selectorValue As ROneCOne
    Dim wrapped As ROneCOne

    Set selectorValue = NormalizeContextUnary(keySelector, "MaxBy", False)
    If Not IsMissing(comparer) Then
        Set comparerValue = NormalizeComparer(comparer, False, "MaxBy")
    End If
    Set wrapped = ExtremeElement( _
        selectorValue, comparerValue, True, "MaxBy")
    AssignPublicResult MaxBy, wrapped
End Function

Public Function ToList() As ROneCOne
Attribute ToList.VB_Description = "Immediately materializes a deferred sequence as List<T>."
    Dim listValue As ROneCOne
    Dim values As Collection
    Dim wrapped As ROneCOne

    RequireSequence Me, "ToList"
    Set listValue = New ROneCOne
    listValue.ConfigureList mElementType, mElementObjectType
    Set values = MaterializeWrapped
    For Each wrapped In values
        listValue.InternalAddWrapped wrapped
    Next wrapped
    Set ToList = listValue
End Function

Public Function ToArray() As Variant
Attribute ToArray.VB_Description = "Immediately materializes a sequence as a zero-based array."
    Dim index As Long
    Dim output() As Variant
    Dim values As Collection
    Dim wrapped As ROneCOne

    Set values = MaterializeWrapped
    If values.Count = 0 Then
        ToArray = Array()
        Exit Function
    End If
    ReDim output(0 To values.Count - 1)
    For index = 1 To values.Count
        Set wrapped = values.Item(index)
        If wrapped.InternalHasObjectValue Then
            Set output(index - 1) = wrapped.InternalObjectValue
        Else
            output(index - 1) = wrapped.InternalScalarValue
        End If
    Next index
    ToArray = output
End Function

Public Function ToDictionary( _
    ByVal keySelector As Variant, _
    Optional ByVal elementSelector As Variant, _
    Optional ByVal comparer As Variant _
) As ROneCOne
Attribute ToDictionary.VB_Description = "Materializes a Dictionary from selected keys."
    Dim comparerValue As ROneCOne
    Dim elementSelectorValue As ROneCOne
    Dim elementValue As ROneCOne
    Dim key As ROneCOne
    Dim keySelectorValue As ROneCOne
    Dim result As ROneCOne
    Dim resultObjectType As String
    Dim resultType As VbVarType
    Dim values As Collection
    Dim wrapped As ROneCOne

    Set keySelectorValue = NormalizeContextUnary( _
        keySelector, "ToDictionary", False)
    resultType = InternalElementType
    resultObjectType = InternalElementObjectType
    If Not IsMissing(elementSelector) Then
        Set elementSelectorValue = NormalizeContextUnary( _
            elementSelector, "ToDictionary", False)
        resultType = vbVariant
        resultObjectType = vbNullString
    End If
    If Not IsMissing(comparer) Then
        Set comparerValue = NormalizeComparer(comparer, True, "ToDictionary")
    End If
    Set result = New ROneCOne
    result.ConfigureGenericCollection ROLE_DICTIONARY, COLLECTION_DICTIONARY, _
        resultType, resultObjectType, vbVariant, _
        vbNullString, comparerValue
    Set values = MaterializeWrapped
    For Each wrapped In values
        Set key = InvokeUnary(keySelectorValue, wrapped)
        If result.FindKeyIndex(key, "ToDictionary") >= 0 Then
            RaiseContractError ERROR_INVALID_ARGUMENT, "ToDictionary", _
                "The selected key is duplicated."
        End If
        If elementSelectorValue Is Nothing Then
            Set elementValue = wrapped
        Else
            Set elementValue = InvokeUnary(elementSelectorValue, wrapped)
        End If
        result.AddDictionaryPair key, elementValue, "ToDictionary"
    Next wrapped
    Set ToDictionary = result
End Function

Public Function ToHashSet(Optional ByVal comparer As Variant) As ROneCOne
Attribute ToHashSet.VB_Description = "Materializes a HashSet<T>."
    Dim comparerValue As ROneCOne
    Dim result As ROneCOne
    Dim values As Collection
    Dim wrapped As ROneCOne

    If Not IsMissing(comparer) Then
        Set comparerValue = NormalizeComparer(comparer, True, "ToHashSet")
    End If
    Set result = New ROneCOne
    result.ConfigureGenericCollection ROLE_HASH_SET, COLLECTION_HASH_SET, _
        InternalElementType, InternalElementObjectType, vbEmpty, _
        vbNullString, comparerValue
    Set values = MaterializeWrapped
    For Each wrapped In values
        If result.FindElementIndex(wrapped, "ToHashSet") < 0 Then
            result.AddSetValue wrapped, "ToHashSet"
        End If
    Next wrapped
    Set ToHashSet = result
End Function

Public Function ToLookup( _
    ByVal keySelector As Variant, _
    Optional ByVal elementSelector As Variant, _
    Optional ByVal comparer As Variant _
) As ROneCOne
Attribute ToLookup.VB_Description = "Materializes a one-to-many Lookup<TKey, TElement>."
    Dim comparerValue As ROneCOne
    Dim elementSelectorValue As ROneCOne
    Dim elementValue As ROneCOne
    Dim group As ROneCOne
    Dim groupIndex As Long
    Dim key As ROneCOne
    Dim keySelectorValue As ROneCOne
    Dim lookup As ROneCOne
    Dim values As Collection
    Dim wrapped As ROneCOne

    Set keySelectorValue = NormalizeContextUnary( _
        keySelector, "ToLookup", False)
    If Not IsMissing(elementSelector) Then
        Set elementSelectorValue = NormalizeContextUnary( _
            elementSelector, "ToLookup", False)
    End If
    If Not IsMissing(comparer) Then
        Set comparerValue = NormalizeComparer(comparer, True, "ToLookup")
    End If
    Set lookup = New ROneCOne
    lookup.ConfigureGenericCollection ROLE_COLLECTION, COLLECTION_LOOKUP, _
        vbObject, "ROneCOne", vbVariant, vbNullString, comparerValue
    Set values = MaterializeWrapped
    For Each wrapped In values
        Set key = InvokeUnary(keySelectorValue, wrapped)
        If elementSelectorValue Is Nothing Then
            Set elementValue = wrapped
        Else
            Set elementValue = InvokeUnary(elementSelectorValue, wrapped)
        End If
        groupIndex = lookup.FindKeyIndex(key, "ToLookup")
        If groupIndex < 0 Then
            Set group = New ROneCOne
            group.ConfigureGrouping key, elementValue
            lookup.InternalAddLookupGroup key, group
        Else
            Set group = lookup.InternalCollectionItem(groupIndex)
            group.InternalAppendCollectionItem elementValue
        End If
    Next wrapped
    Set ToLookup = lookup
End Function

Public Function GroupBy( _
    ByVal keySelector As Variant, _
    Optional ByVal elementSelector As Variant, _
    Optional ByVal comparer As Variant _
) As ROneCOne
Attribute GroupBy.VB_Description = "Groups elements by a selected key."
    Set GroupBy = ToLookup(keySelector, elementSelector, comparer).ToList
End Function

Public Function Join( _
    ByVal inner As Variant, _
    ByVal outerKeySelector As Variant, _
    ByVal innerKeySelector As Variant, _
    ByVal resultSelector As Variant _
) As ROneCOne
Attribute Join.VB_Description = "Correlates two sequences by matching keys."
    Set Join = ExecuteJoin( _
        inner, outerKeySelector, innerKeySelector, resultSelector, False)
End Function

Public Function GroupJoin( _
    ByVal inner As Variant, _
    ByVal outerKeySelector As Variant, _
    ByVal innerKeySelector As Variant, _
    ByVal resultSelector As Variant _
) As ROneCOne
Attribute GroupJoin.VB_Description = "Correlates outer elements with grouped inner matches."
    Set GroupJoin = ExecuteJoin( _
        inner, outerKeySelector, innerKeySelector, resultSelector, True)
End Function

Public Function Zip( _
    ByVal second As Variant, _
    Optional ByVal resultSelector As Variant _
) As ROneCOne
Attribute Zip.VB_Description = "Combines elements at corresponding positions."
    Dim firstValues As Collection
    Dim index As Long
    Dim itemCount As Long
    Dim result As ROneCOne
    Dim resultSelectorValue As ROneCOne
    Dim secondSequence As ROneCOne
    Dim secondValues As Collection
    Dim selected As ROneCOne

    Set secondSequence = NormalizeMembershipSequence(second, "Zip")
    Set firstValues = MaterializeWrapped
    Set secondValues = secondSequence.MaterializeWrapped
    itemCount = firstValues.Count
    If secondValues.Count < itemCount Then itemCount = secondValues.Count
    Set result = ROneCOne.ListOf(vbVariant)
    If Not IsMissing(resultSelector) Then
        Set resultSelectorValue = NormalizeBinaryDelegate(resultSelector, "Zip")
    End If
    For index = 1 To itemCount
        If resultSelectorValue Is Nothing Then
            Set selected = CreateDictionaryEntry( _
                firstValues.Item(index), secondValues.Item(index))
            result.Add selected
        Else
            Set selected = InvokeBinary(resultSelectorValue, _
                firstValues.Item(index), secondValues.Item(index))
            result.InternalAddWrapped selected
        End If
    Next index
    Set Zip = result
End Function

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_Description = "Gets the VBA For Each enumerator for this sequence."
Attribute NewEnum.VB_UserMemId = -4
    If mRole = ROLE_LIST Or mRole = ROLE_QUERY Or _
        mRole = ROLE_DATA_VIEW Or IsGenericCollectionRole(mRole) Then
        If mEnumerationValues Is Nothing Or _
            mEnumerationVersion <> InternalSequenceVersion Then
            Set mEnumerationValues = MaterializeValues
            mEnumerationVersion = InternalSequenceVersion
        End If
    End If
    Set NewEnum = mEnumerationValues.[_NewEnum]
End Property

Public Property Get Arity() As Long
Attribute Arity.VB_Description = "Gets the number of delegate parameters, or -1 when dynamic."
    Select Case mRole
        Case ROLE_LAMBDA
            Arity = mParameters.Count
        Case ROLE_METHOD, ROLE_PROCEDURE, ROLE_NATIVE
            Arity = mMethodArity
        Case ROLE_COMPOSITION
            Arity = mLeft.Arity
        Case ROLE_MULTICAST
            Arity = mInvocationList.Item(1).Arity
        Case Else
            Arity = -1
    End Select
End Property

Public Property Get InvocationCount() As Long
Attribute InvocationCount.VB_Description = "Gets the immutable delegate invocation count."
    RequireDelegate Me, "InvocationCount"
    If mRole = ROLE_MULTICAST Then
        InvocationCount = mInvocationList.Count
    Else
        InvocationCount = 1
    End If
End Property

Public Property Get Target() As Object
Attribute Target.VB_Description = "Gets the bound object target, or Nothing."
    RequireDelegate Me, "Target"
    Set Target = mTarget
End Property

Public Property Get MethodName() As String
Attribute MethodName.VB_Description = "Gets the bound method or workbook procedure name."
    RequireDelegate Me, "MethodName"
    MethodName = mMethodName
End Property

Public Property Get IsAction() As Boolean
Attribute IsAction.VB_Description = "Gets whether the delegate has no result value."
    RequireDelegate Me, "IsAction"
    IsAction = mIsActionDelegate
End Property

Public Property Get Signature() As String
Attribute Signature.VB_Description = "Gets the C#-style delegate signature description."
    RequireDelegate Me, "Signature"
    Signature = BuildSignature(Me)
End Property

Public Function GetInvocationList() As ROneCOne
Attribute GetInvocationList.VB_Description = "Returns a typed copy of the invocation list."
    Dim delegateValue As ROneCOne
    Dim result As ROneCOne

    RequireDelegate Me, "GetInvocationList"
    Set result = ROneCOne.ListOf(Me)
    If mRole = ROLE_MULTICAST Then
        For Each delegateValue In mInvocationList
            result.Add delegateValue
        Next delegateValue
    Else
        result.Add Me
    End If
    Set GetInvocationList = result
End Function

Public Function ToDisplayString() As String
Attribute ToDisplayString.VB_Description = _
    "Renders an expression, lambda, or delegate as readable text for debugging."
    Dim parameters As Collection

    If mRole = ROLE_LAMBDA Then
        ToDisplayString = RenderExpressionNode(mLeft, mParameters)
    ElseIf mRole = ROLE_VALUE Or mRole = ROLE_PARAMETER Or _
        mRole = ROLE_EXPRESSION Then
        Set parameters = New Collection
        ToDisplayString = RenderExpressionNode(Me, parameters)
    ElseIf IsDelegateRole(mRole) Then
        ToDisplayString = BuildSignature(Me)
    Else
        RaiseContractError ERROR_INVALID_OPERATION, "ToDisplayString", _
            "ToDisplayString requires an expression or delegate."
    End If
End Function

Private Function RenderExpressionNode( _
    ByVal node As ROneCOne, _
    ByVal parameters As Collection _
) As String
    Select Case node.InternalRole
        Case ROLE_VALUE
            RenderExpressionNode = RenderExpressionLiteral(node)
        Case ROLE_PARAMETER
            RenderExpressionNode = ParameterDisplayName(node, parameters)
        Case ROLE_EXPRESSION
            RenderExpressionNode = RenderExpressionOperator(node, parameters)
        Case Else
            RenderExpressionNode = "?"
    End Select
End Function

Private Function RenderExpressionLiteral(ByVal node As ROneCOne) As String
    Dim scalarValue As Variant

    If node.InternalHasObjectValue Then
        If node.InternalObjectValue Is Nothing Then
            RenderExpressionLiteral = "Nothing"
        Else
            RenderExpressionLiteral = "{" & TypeName(node.InternalObjectValue) & "}"
        End If
        Exit Function
    End If
    scalarValue = node.InternalScalarValue
    If IsNull(scalarValue) Then
        RenderExpressionLiteral = "Null"
    ElseIf VarTypeOf(scalarValue) = vbString Then
        RenderExpressionLiteral = """" & CStr(scalarValue) & """"
    ElseIf VarTypeOf(scalarValue) = vbBoolean Then
        RenderExpressionLiteral = IIf(CBool(scalarValue), "True", "False")
    Else
        RenderExpressionLiteral = CStr(scalarValue)
    End If
End Function

Private Function ParameterDisplayName( _
    ByVal node As ROneCOne, _
    ByVal parameters As Collection _
) As String
    Dim index As Long

    For index = 1 To parameters.Count
        If parameters.Item(index) Is node Then
            If index <= 3 Then
                ParameterDisplayName = Mid$("xyz", index, 1)
            Else
                ParameterDisplayName = "p" & CStr(index)
            End If
            Exit Function
        End If
    Next index
    ParameterDisplayName = "x"
End Function

Private Function RenderExpressionOperator( _
    ByVal node As ROneCOne, _
    ByVal parameters As Collection _
) As String
    Dim leftText As String
    Dim rightText As String

    leftText = RenderExpressionNode(node.InternalLeft, parameters)
    If Not node.InternalRight Is Nothing Then
        rightText = RenderExpressionNode(node.InternalRight, parameters)
    End If

    Select Case node.InternalOperator
        Case OP_ADD
            RenderExpressionOperator = "(" & leftText & " + " & rightText & ")"
        Case OP_SUBTRACT
            RenderExpressionOperator = "(" & leftText & " - " & rightText & ")"
        Case OP_MULTIPLY
            RenderExpressionOperator = "(" & leftText & " * " & rightText & ")"
        Case OP_DIVIDE
            RenderExpressionOperator = "(" & leftText & " / " & rightText & ")"
        Case OP_MODULO
            RenderExpressionOperator = "(" & leftText & " Mod " & rightText & ")"
        Case OP_CONCAT
            RenderExpressionOperator = "(" & leftText & " & " & rightText & ")"
        Case OP_EQUAL, OP_STRING_EQUAL
            RenderExpressionOperator = "(" & leftText & " = " & rightText & ")"
        Case OP_NOT_EQUAL, OP_STRING_NOT_EQUAL
            RenderExpressionOperator = "(" & leftText & " <> " & rightText & ")"
        Case OP_LESS_THAN
            RenderExpressionOperator = "(" & leftText & " < " & rightText & ")"
        Case OP_LESS_THAN_OR_EQUAL
            RenderExpressionOperator = "(" & leftText & " <= " & rightText & ")"
        Case OP_GREATER_THAN
            RenderExpressionOperator = "(" & leftText & " > " & rightText & ")"
        Case OP_GREATER_THAN_OR_EQUAL
            RenderExpressionOperator = "(" & leftText & " >= " & rightText & ")"
        Case OP_AND_ALSO
            RenderExpressionOperator = "(" & leftText & " And " & rightText & ")"
        Case OP_OR_ELSE
            RenderExpressionOperator = "(" & leftText & " Or " & rightText & ")"
        Case OP_NEGATE
            RenderExpressionOperator = "-" & leftText
        Case OP_NOT
            RenderExpressionOperator = "Not " & leftText
        Case OP_MEMBER_VALUE, OP_MEMBER_OBJECT
            RenderExpressionOperator = leftText & "." & node.InternalMethodName
        Case OP_SAFE_MEMBER_VALUE, OP_SAFE_MEMBER_OBJECT
            RenderExpressionOperator = leftText & "?." & node.InternalMethodName
        Case OP_STARTS_WITH
            RenderExpressionOperator = leftText & ".StartsWith(" & rightText & ")"
        Case OP_ENDS_WITH
            RenderExpressionOperator = leftText & ".EndsWith(" & rightText & ")"
        Case OP_CONTAINS_TEXT
            RenderExpressionOperator = leftText & ".Contains(" & rightText & ")"
        Case OP_MATCHES_PATTERN
            RenderExpressionOperator = leftText & ".Like(" & rightText & ")"
        Case OP_IS_NOTHING
            RenderExpressionOperator = leftText & " Is Nothing"
        Case OP_IS_NOT_NOTHING
            RenderExpressionOperator = leftText & " Is Not Nothing"
        Case OP_IS_NULL_OR_EMPTY
            RenderExpressionOperator = leftText & " Is Null Or Empty"
        Case OP_IN_SEQUENCE
            RenderExpressionOperator = leftText & " In [...]"
        Case OP_NOT_IN_SEQUENCE
            RenderExpressionOperator = leftText & " Not In [...]"
        Case OP_ANY_MATCH
            RenderExpressionOperator = leftText & ".Any(...)"
        Case OP_ALL_MATCH
            RenderExpressionOperator = leftText & ".All(...)"
        Case OP_NONE_MATCH
            RenderExpressionOperator = leftText & ".None(...)"
        Case Else
            RenderExpressionOperator = leftText
    End Select
End Function

Private Sub RequireQueryable(ByVal memberName As String)
    If mRole <> ROLE_QUERYABLE Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            memberName & " requires a Queryable."
    End If
End Sub

Private Sub RaiseQueryError(ByVal message As String)
    Err.Raise ERROR_QUERY, "ROneCOne.QueryTranslationException", message
End Sub

' The parameter below is deliberately not called "name". VBA normalizes
' identifier casing across the whole project, so declaring a lowercase
' "name" rewrites every bare Name identifier, including the member that
' bang syntax turns into a string (rows!Name becomes rows!name). Column
' lookup compares case-sensitively, so that silently breaks it.
Private Function ValidateQueryIdentifier( _
    ByVal identifierText As String, _
    ByVal memberName As String _
) As String
    Dim code As Long
    Dim index As Long

    identifierText = Trim$(identifierText)
    If Len(identifierText) = 0 Or Len(identifierText) > 128 Then
        RaiseQueryError memberName & _
            " requires an identifier of 1 to 128 characters."
    End If
    For index = 1 To Len(identifierText)
        code = AscW(Mid$(identifierText, index, 1))
        If code < 32 Or code = 91 Or code = 93 Then
            RaiseQueryError memberName & " cannot use the identifier [" & _
                identifierText & "]; brackets and control characters " & _
                "are refused."
        End If
    Next index
    ValidateQueryIdentifier = identifierText
End Function

Private Function QueryQuoteIdentifier( _
    ByVal identifierText As String _
) As String
    QueryQuoteIdentifier = "[" & identifierText & "]"
End Function

Private Function QueryableClone() As ROneCOne
    Dim clone As ROneCOne
    Dim columns As Collection
    Dim item As Variant
    Dim orderings As Collection
    Dim wheres As Collection

    Set wheres = New Collection
    For Each item In mSqlQueryWheres
        wheres.Add item
    Next item
    Set orderings = New Collection
    For Each item In mSqlQueryOrderings
        orderings.Add item
    Next item
    Set columns = New Collection
    For Each item In mSqlQueryColumns
        columns.Add item
    Next item
    Set clone = New ROneCOne
    clone.ConfigureQueryableState mSqlQueryConnection, mSqlQueryTable, _
        mSqlQueryDialect, wheres, orderings, columns, mSqlQueryTake, _
        mSqlQuerySkip
    Set QueryableClone = clone
End Function

Private Function QueryableWithPredicate( _
    ByVal predicate As ROneCOne _
) As ROneCOne
    Dim clone As ROneCOne

    Set clone = QueryableClone()
    clone.InternalQueryableAddWhere QueryablePredicateBody(predicate)
    Set QueryableWithPredicate = clone
End Function

Private Function QueryablePredicateBody( _
    ByVal predicate As ROneCOne _
) As ROneCOne
    If predicate Is Nothing Then
        RaiseQueryError "The predicate cannot be Nothing."
    End If
    Select Case predicate.InternalRole
        Case ROLE_EXPRESSION
            Set QueryablePredicateBody = predicate
        Case ROLE_LAMBDA
            Set QueryablePredicateBody = predicate.InternalLeft
        Case ROLE_WHERE_BUILDER
            RaiseQueryError "Finish the fluent condition, for example " & _
                "Where(""Total"").AtLeast(100), before it can translate."
        Case Else
            RaiseQueryError "Only expression predicates translate to " & _
                "SQL; method and native delegates carry no expression tree."
    End Select
End Function

Private Sub QueryableRefuseComparer( _
    Optional ByVal comparer As Variant, _
    Optional ByVal memberName As String = "OrderBy" _
)
    If Not IsMissing(comparer) Then
        RaiseQueryError memberName & " cannot send a comparer to SQL; " & _
            "order by a column instead."
    End If
End Sub

Private Function QueryableOrdered( _
    ByVal keySelector As Variant, _
    ByVal descending As Boolean, _
    ByVal thenBy As Boolean, _
    ByVal memberName As String _
) As ROneCOne
    Dim clone As ROneCOne
    Dim columnName As String

    columnName = QueryableOrderColumn(keySelector, memberName)
    If thenBy And mSqlQueryOrderings.Count = 0 Then
        RaiseQueryError memberName & " requires OrderBy first."
    End If
    Set clone = QueryableClone()
    clone.InternalQueryableOrder columnName, descending, Not thenBy
    Set QueryableOrdered = clone
End Function

Private Function QueryableOrderColumn( _
    ByVal keySelector As Variant, _
    ByVal memberName As String _
) As String
    Dim node As ROneCOne

    If Not IsObject(keySelector) Then
        If VarTypeOf(keySelector) = vbString Then
            QueryableOrderColumn = ValidateQueryIdentifier( _
                CStr(keySelector), memberName)
            Exit Function
        End If
    End If
    Set node = RequireROneCOne(keySelector, memberName)
    If node.InternalRole = ROLE_EXPRESSION Then
        If node.InternalOperator = OP_MEMBER_VALUE Then
            If node.InternalLeft.InternalRole = ROLE_PARAMETER Then
                QueryableOrderColumn = ValidateQueryIdentifier( _
                    node.InternalMethodName, memberName)
                Exit Function
            End If
        End If
    End If
    RaiseQueryError memberName & " requires a column name or a single " & _
        "member expression."
End Function

Private Function QueryableExecuteTable() As ROneCOne
    Dim command As ROneCOne
    Dim table As ROneCOne

    Set command = QueryableBuildCommand(QueryableBuildSql(False))
    Set table = ROneCOne.DataTable(mSqlQueryTable)
    ROneCOne.DbDataAdapter(command).Fill table
    Set QueryableExecuteTable = table
End Function

Private Function QueryableExecuteCount() As Long
    Dim command As ROneCOne

    Set command = QueryableBuildCommand(QueryableBuildSql(True))
    QueryableExecuteCount = CLng(command.ExecuteScalar)
End Function

Private Function QueryableBuildCommand(ByVal sql As String) As ROneCOne
    Dim command As ROneCOne
    Dim index As Long
    Dim value As Variant

    Set command = ROneCOne.DbCommand(sql, mSqlQueryConnection)
    For Each value In mSqlQueryLastValues
        index = index + 1
        command.WithParameter "p" & CStr(index), value
    Next value
    Set QueryableBuildCommand = command
End Function

Private Function QueryableBuildSql(ByVal forCount As Boolean) As String
    Dim body As ROneCOne
    Dim columnName As Variant
    Dim orderText As String
    Dim ordering As Variant
    Dim rowsSql As String
    Dim selectList As String
    Dim topClause As String
    Dim values As Collection
    Dim whereText As String

    Set values = New Collection
    For Each body In mSqlQueryWheres
        If Len(whereText) > 0 Then whereText = whereText & " AND "
        whereText = whereText & QueryableRenderNode(body, values)
    Next body

    If mSqlQueryColumns.Count = 0 Then
        selectList = "*"
    Else
        For Each columnName In mSqlQueryColumns
            If Len(selectList) > 0 Then selectList = selectList & ", "
            selectList = selectList & QueryQuoteIdentifier(CStr(columnName))
        Next columnName
    End If

    For Each ordering In mSqlQueryOrderings
        If Len(orderText) > 0 Then orderText = orderText & ", "
        orderText = orderText & QueryQuoteIdentifier(CStr(ordering(0)))
        If ordering(1) Then orderText = orderText & " DESC"
    Next ordering

    If mSqlQuerySkip >= 0 And Len(orderText) = 0 Then
        RaiseQueryError "Skip translates to OFFSET, which SQL only " & _
            "accepts under an ORDER BY; call OrderBy first."
    End If

    If mSqlQueryTake >= 0 And mSqlQuerySkip < 0 Then
        If mSqlQueryDialect = QUERY_DIALECT_TSQL Then
            topClause = "TOP (" & CStr(mSqlQueryTake) & ") "
        Else
            topClause = "TOP " & CStr(mSqlQueryTake) & " "
        End If
    End If

    rowsSql = "SELECT " & topClause & selectList & " FROM " & _
        QueryQuoteIdentifier(mSqlQueryTable)
    If Len(whereText) > 0 Then rowsSql = rowsSql & " WHERE " & whereText
    If Len(orderText) > 0 Then
        If Not forCount Or mSqlQueryTake >= 0 Or mSqlQuerySkip >= 0 Then
            rowsSql = rowsSql & " ORDER BY " & orderText
        End If
    End If
    If mSqlQuerySkip >= 0 Then
        rowsSql = rowsSql & " OFFSET " & CStr(mSqlQuerySkip) & " ROWS"
        If mSqlQueryTake >= 0 Then
            rowsSql = rowsSql & " FETCH NEXT " & CStr(mSqlQueryTake) & _
                " ROWS ONLY"
        End If
    End If

    If forCount Then
        If mSqlQueryTake >= 0 Or mSqlQuerySkip >= 0 Then
            QueryableBuildSql = "SELECT COUNT(*) FROM (" & rowsSql & _
                ") AS [q]"
        ElseIf Len(whereText) > 0 Then
            QueryableBuildSql = "SELECT COUNT(*) FROM " & _
                QueryQuoteIdentifier(mSqlQueryTable) & " WHERE " & whereText
        Else
            QueryableBuildSql = "SELECT COUNT(*) FROM " & _
                QueryQuoteIdentifier(mSqlQueryTable)
        End If
    Else
        QueryableBuildSql = rowsSql
    End If
    Set mSqlQueryLastValues = values
End Function

Private Function QueryableRenderNode( _
    ByVal node As ROneCOne, _
    ByVal values As Collection _
) As String
    Select Case node.InternalRole
        Case ROLE_VALUE
            QueryableRenderNode = QueryableRenderValue(node, values)
        Case ROLE_PARAMETER
            RaiseQueryError "The row itself cannot appear in SQL; " & _
                "select a column with a member path."
        Case ROLE_EXPRESSION
            QueryableRenderNode = QueryableRenderOperator(node, values)
        Case Else
            RaiseQueryError "This value cannot translate to SQL."
    End Select
End Function

Private Function QueryableIsNullValueNode(ByVal node As ROneCOne) As Boolean
    Dim wrapped As ROneCOne

    If node Is Nothing Then Exit Function
    If node.InternalRole <> ROLE_VALUE Then Exit Function
    If node.InternalHasObjectValue Then
        If node.InternalObjectValue Is Nothing Then
            QueryableIsNullValueNode = True
            Exit Function
        End If
        If TypeOf node.InternalObjectValue Is ROneCOne Then
            Set wrapped = node.InternalObjectValue
            QueryableIsNullValueNode = _
                (wrapped.InternalRole = ROLE_DB_NULL)
        End If
        Exit Function
    End If
    QueryableIsNullValueNode = IsNull(node.InternalScalarValue)
End Function

Private Function QueryableRenderValue( _
    ByVal node As ROneCOne, _
    ByVal values As Collection _
) As String
    Dim scalarValue As Variant

    If QueryableIsNullValueNode(node) Then
        RaiseQueryError "A null compares only through EqualTo or " & _
            "NotEqualTo, which translate to IS NULL and IS NOT NULL."
    End If
    If node.InternalHasObjectValue Then
        RaiseQueryError "An object value cannot become a SQL parameter."
    End If
    scalarValue = node.InternalScalarValue
    If IsEmpty(scalarValue) Then
        RaiseQueryError "An Empty value cannot become a SQL parameter."
    End If
    values.Add scalarValue
    QueryableRenderValue = "?"
End Function

Private Function QueryableRenderOperator( _
    ByVal node As ROneCOne, _
    ByVal values As Collection _
) As String
    Dim beforeCount As Long
    Dim caseFunction As String
    Dim leftNull As Boolean
    Dim leftText As String
    Dim rightNull As Boolean
    Dim rightText As String

    Select Case node.InternalOperator
        Case OP_MEMBER_VALUE
            If node.InternalLeft.InternalRole <> ROLE_PARAMETER Then
                RaiseQueryError "Only direct columns of the row " & _
                    "translate; nested member paths would need a join."
            End If
            QueryableRenderOperator = QueryQuoteIdentifier( _
                ValidateQueryIdentifier(node.InternalMethodName, "Member"))
            Exit Function
        Case OP_EQUAL, OP_NOT_EQUAL
            leftNull = QueryableIsNullValueNode(node.InternalLeft)
            rightNull = QueryableIsNullValueNode(node.InternalRight)
            If leftNull And rightNull Then
                QueryableRenderOperator = IIf( _
                    node.InternalOperator = OP_EQUAL, "(1 = 1)", "(1 = 0)")
                Exit Function
            End If
            If rightNull Then
                leftText = QueryableRenderNode(node.InternalLeft, values)
                QueryableRenderOperator = "(" & leftText & " IS " & IIf( _
                    node.InternalOperator = OP_EQUAL, "", "NOT ") & "NULL)"
                Exit Function
            End If
            If leftNull Then
                rightText = QueryableRenderNode(node.InternalRight, values)
                QueryableRenderOperator = "(" & rightText & " IS " & IIf( _
                    node.InternalOperator = OP_EQUAL, "", "NOT ") & "NULL)"
                Exit Function
            End If
            leftText = QueryableRenderNode(node.InternalLeft, values)
            rightText = QueryableRenderNode(node.InternalRight, values)
            QueryableRenderOperator = "(" & leftText & IIf( _
                node.InternalOperator = OP_EQUAL, " = ", " <> ") & _
                rightText & ")"
            Exit Function
        Case OP_IN_SEQUENCE, OP_NOT_IN_SEQUENCE
            QueryableRenderOperator = QueryableRenderIn(node, values)
            Exit Function
        Case OP_IS_NOTHING, OP_IS_NOT_NOTHING
            leftText = QueryableRenderNode(node.InternalLeft, values)
            QueryableRenderOperator = "(" & leftText & " IS " & IIf( _
                node.InternalOperator = OP_IS_NOTHING, "", "NOT ") & "NULL)"
            Exit Function
        Case OP_IS_NULL_OR_EMPTY
            beforeCount = values.Count
            leftText = QueryableRenderNode(node.InternalLeft, values)
            If values.Count <> beforeCount Then
                RaiseQueryError "IsNullOrEmpty translates only over a " & _
                    "plain column."
            End If
            QueryableRenderOperator = "(" & leftText & " IS NULL OR " & _
                leftText & " = '')"
            Exit Function
        Case OP_STARTS_WITH, OP_ENDS_WITH, OP_CONTAINS_TEXT
            QueryableRenderOperator = QueryableRenderLike(node, values)
            Exit Function
        Case OP_STRING_EQUAL, OP_STRING_NOT_EQUAL
            caseFunction = IIf(mSqlQueryDialect = QUERY_DIALECT_TSQL, _
                "UPPER", "UCASE")
            leftText = QueryableRenderNode(node.InternalLeft, values)
            rightText = QueryableRenderNode(node.InternalRight, values)
            QueryableRenderOperator = "(" & caseFunction & "(" & _
                leftText & ")" & IIf( _
                node.InternalOperator = OP_STRING_EQUAL, " = ", " <> ") & _
                caseFunction & "(" & rightText & "))"
            Exit Function
    End Select

    leftText = QueryableRenderNode(node.InternalLeft, values)
    If node.InternalOperator = OP_NOT Then
        QueryableRenderOperator = "(NOT " & leftText & ")"
        Exit Function
    End If
    If node.InternalOperator = OP_NEGATE Then
        QueryableRenderOperator = "(-" & leftText & ")"
        Exit Function
    End If

    rightText = QueryableRenderNode(node.InternalRight, values)
    Select Case node.InternalOperator
        Case OP_ADD
            QueryableRenderOperator = "(" & leftText & " + " & rightText & ")"
        Case OP_SUBTRACT
            QueryableRenderOperator = "(" & leftText & " - " & rightText & ")"
        Case OP_MULTIPLY
            QueryableRenderOperator = "(" & leftText & " * " & rightText & ")"
        Case OP_DIVIDE
            QueryableRenderOperator = "(" & leftText & " / " & rightText & ")"
        Case OP_MODULO
            QueryableRenderOperator = "(" & leftText & IIf( _
                mSqlQueryDialect = QUERY_DIALECT_TSQL, " % ", " MOD ") & _
                rightText & ")"
        Case OP_CONCAT
            QueryableRenderOperator = "(" & leftText & IIf( _
                mSqlQueryDialect = QUERY_DIALECT_TSQL, " + ", " & ") & _
                rightText & ")"
        Case OP_LESS_THAN
            QueryableRenderOperator = "(" & leftText & " < " & rightText & ")"
        Case OP_LESS_THAN_OR_EQUAL
            QueryableRenderOperator = "(" & leftText & " <= " & _
                rightText & ")"
        Case OP_GREATER_THAN
            QueryableRenderOperator = "(" & leftText & " > " & rightText & ")"
        Case OP_GREATER_THAN_OR_EQUAL
            QueryableRenderOperator = "(" & leftText & " >= " & _
                rightText & ")"
        Case OP_AND_ALSO
            QueryableRenderOperator = "(" & leftText & " AND " & _
                rightText & ")"
        Case OP_OR_ELSE
            QueryableRenderOperator = "(" & leftText & " OR " & _
                rightText & ")"
        Case Else
            RaiseQueryError "SQL cannot express " & _
                QueryableOperatorLabel(node.InternalOperator) & "."
    End Select
End Function

Private Function QueryableRenderLike( _
    ByVal node As ROneCOne, _
    ByVal values As Collection _
) As String
    Dim leftText As String
    Dim patternText As String
    Dim rightNode As ROneCOne

    Set rightNode = node.InternalRight
    If rightNode.InternalRole <> ROLE_VALUE Or _
        rightNode.InternalHasObjectValue Or _
        QueryableIsNullValueNode(rightNode) Then
        RaiseQueryError "The text argument of StartsWith, EndsWith, and " & _
            "Contains must be a captured text value."
    End If
    patternText = QueryableEscapeLike(CStr(rightNode.InternalScalarValue))
    Select Case node.InternalOperator
        Case OP_STARTS_WITH
            patternText = patternText & "%"
        Case OP_ENDS_WITH
            patternText = "%" & patternText
        Case Else
            patternText = "%" & patternText & "%"
    End Select
    leftText = QueryableRenderNode(node.InternalLeft, values)
    values.Add patternText
    QueryableRenderLike = "(" & leftText & " LIKE ?" & IIf( _
        mSqlQueryDialect = QUERY_DIALECT_TSQL, " ESCAPE '\'", "") & ")"
End Function

Private Function QueryableEscapeLike(ByVal text As String) As String
    If mSqlQueryDialect = QUERY_DIALECT_TSQL Then
        text = VBA.Replace(text, "\", "\\")
        text = VBA.Replace(text, "%", "\%")
        text = VBA.Replace(text, "_", "\_")
        text = VBA.Replace(text, "[", "\[")
    Else
        text = VBA.Replace(text, "[", "[[]")
        text = VBA.Replace(text, "%", "[%]")
        text = VBA.Replace(text, "_", "[_]")
    End If
    QueryableEscapeLike = text
End Function

Private Function QueryableRenderIn( _
    ByVal node As ROneCOne, _
    ByVal values As Collection _
) As String
    Dim element As Variant
    Dim leftText As String
    Dim markers As String
    Dim negated As Boolean
    Dim total As Long

    negated = (node.InternalOperator = OP_NOT_IN_SEQUENCE)
    If Not node.InternalQueryComparer Is Nothing Then
        RaiseQueryError "IsIn cannot send a custom comparer to SQL."
    End If
    leftText = QueryableRenderNode(node.InternalLeft, values)
    For Each element In node.InternalQueryValue
        If IsObject(element) Or IsNull(element) Or IsEmpty(element) Then
            RaiseQueryError "IsIn candidates must be captured scalar " & _
                "values; SQL IN never matches NULL."
        End If
        values.Add element
        If Len(markers) > 0 Then markers = markers & ", "
        markers = markers & "?"
        total = total + 1
    Next element
    If total = 0 Then
        QueryableRenderIn = IIf(negated, "(1 = 1)", "(1 = 0)")
        Exit Function
    End If
    QueryableRenderIn = "(" & leftText & IIf(negated, " NOT IN (", _
        " IN (") & markers & "))"
End Function

Private Function QueryableOperatorLabel(ByVal operatorCode As Long) As String
    Select Case operatorCode
        Case OP_MATCHES_PATTERN
            QueryableOperatorLabel = "a Like pattern"
        Case OP_ANY_MATCH, OP_ALL_MATCH, OP_NONE_MATCH
            QueryableOperatorLabel = "a nested quantifier"
        Case OP_SAFE_MEMBER_VALUE, OP_SAFE_MEMBER_OBJECT
            QueryableOperatorLabel = "a null-safe member path"
        Case OP_MEMBER_OBJECT
            QueryableOperatorLabel = "an object member path"
        Case Else
            QueryableOperatorLabel = "this expression"
    End Select
End Function

Public Sub ForEach(ByVal action As ROneCOne)
Attribute ForEach.VB_Description = "Executes an Action once for every sequence element."
    Dim actionDelegate As ROneCOne
    Dim ignored As ROneCOne
    Dim values As Collection
    Dim wrapped As ROneCOne

    Set actionDelegate = NormalizeUnaryAction(action, "ForEach")
    Set values = MaterializeWrapped
    For Each wrapped In values
        Set ignored = InvokeUnary(actionDelegate, wrapped)
    Next wrapped
End Sub

Public Function JoinText( _
    ByVal separator As String, _
    Optional ByVal selector As Variant _
) As String
Attribute JoinText.VB_Description = "Joins scalar sequence values into one String."
    Dim index As Long
    Dim parts() As String
    Dim selected As ROneCOne
    Dim selectorDelegate As ROneCOne
    Dim values As Collection
    Dim wrapped As ROneCOne

    RequireSequence Me, "JoinText"
    If Not IsMissing(selector) Then
        Set selectorDelegate = NormalizeContextUnary( _
            selector, "JoinText", False)
    End If
    Set values = MaterializeWrapped
    If values.Count = 0 Then Exit Function

    ReDim parts(0 To values.Count - 1)
    For Each wrapped In values
        If selectorDelegate Is Nothing Then
            Set selected = wrapped
        Else
            Set selected = InvokeUnary(selectorDelegate, wrapped)
        End If
        parts(index) = CStr(RequireScalar(selected, "JoinText"))
        index = index + 1
    Next wrapped
    JoinText = VBA.Join(parts, separator)
End Function

Public Property Get TypeMismatchError() As Long
Attribute TypeMismatchError.VB_Description = "Gets the ROneCOne type-mismatch error number."
    TypeMismatchError = ERROR_TYPE_MISMATCH
End Property

Public Property Get HttpRequestError() As Long
Attribute HttpRequestError.VB_Description = "Gets the ROneCOne HTTP request-failure error number."
    HttpRequestError = ERROR_HTTP_REQUEST
End Property

Public Property Get JsonError() As Long
Attribute JsonError.VB_Description = "Gets the ROneCOne JSON parse and write error number."
    JsonError = ERROR_JSON_PARSE
End Property

Public Property Get IOError() As Long
Attribute IOError.VB_Description = "Gets the ROneCOne file system error number."
    IOError = ERROR_IO
End Property

Public Property Get CsvError() As Long
Attribute CsvError.VB_Description = "Gets the ROneCOne CSV parse and write error number."
    CsvError = ERROR_CSV_PARSE
End Property

Public Property Get ProcessError() As Long
Attribute ProcessError.VB_Description = "Gets the ROneCOne process start error number."
    ProcessError = ERROR_PROCESS
End Property

Public Property Get RegexError() As Long
Attribute RegexError.VB_Description = "Gets the ROneCOne regular expression error number."
    RegexError = ERROR_REGEX
End Property

Public Property Get FormatError() As Long
Attribute FormatError.VB_Description = "Gets the ROneCOne text format error number."
    FormatError = ERROR_FORMAT
End Property

Public Property Get XmlError() As Long
Attribute XmlError.VB_Description = "Gets the ROneCOne XML parse and query error number."
    XmlError = ERROR_XML
End Property

Public Property Get ZipError() As Long
Attribute ZipError.VB_Description = "Gets the ROneCOne zip archive error number."
    ZipError = ERROR_ZIP
End Property

Public Property Get QueryError() As Long
Attribute QueryError.VB_Description = "Gets the ROneCOne SQL translation error number."
    QueryError = ERROR_QUERY
End Property

Public Property Get UnboundParameterError() As Long
Attribute UnboundParameterError.VB_Description = "Gets the unbound-parameter error number."
    UnboundParameterError = ERROR_UNBOUND_PARAMETER
End Property

Public Property Get MemberAccessError() As Long
Attribute MemberAccessError.VB_Description = "Gets the object-member access error number."
    MemberAccessError = ERROR_MEMBER_ACCESS
End Property

Public Property Get NativeInvocationError() As Long
Attribute NativeInvocationError.VB_Description = "Gets the native invocation error number."
    NativeInvocationError = ERROR_NATIVE_INVOCATION
End Property

Public Property Get ByRefInvocationError() As Long
Attribute ByRefInvocationError.VB_Description = "Gets the unsupported ByRef boundary error."
    ByRefInvocationError = ERROR_BYREF_INVOCATION
End Property

Public Property Get InvalidArgumentError() As Long
Attribute InvalidArgumentError.VB_Description = "Gets the ROneCOne invalid-argument error number."
    InvalidArgumentError = ERROR_INVALID_ARGUMENT
End Property

Public Property Get InvalidOperationError() As Long
Attribute InvalidOperationError.VB_Description = "Gets the invalid delegate operation error."
    InvalidOperationError = ERROR_INVALID_OPERATION
End Property

Public Property Get ErrorNumber() As Long
Attribute ErrorNumber.VB_Description = "Gets the captured VBA error number."
    RequireError Me, "ErrorNumber"
    ErrorNumber = mErrorNumber
End Property

Public Property Get ErrorSource() As String
Attribute ErrorSource.VB_Description = "Gets the captured VBA error source."
    RequireError Me, "ErrorSource"
    ErrorSource = mErrorSource
End Property

Public Property Get Message() As String
Attribute Message.VB_Description = "Gets the captured VBA error description."
    RequireError Me, "Message"
    Message = mErrorDescription
End Property

Public Property Get HelpFile() As String
Attribute HelpFile.VB_Description = "Gets the captured VBA error help file."
    RequireError Me, "HelpFile"
    HelpFile = mErrorHelpFile
End Property

Public Property Get HelpContext() As Long
Attribute HelpContext.VB_Description = "Gets the captured VBA error help context."
    RequireError Me, "HelpContext"
    HelpContext = mErrorHelpContext
End Property

Friend Property Get InternalRole() As Long
Attribute InternalRole.VB_MemberFlags = "40"
    InternalRole = mRole
End Property

Friend Property Get InternalDateTimeMs() As Double
Attribute InternalDateTimeMs.VB_MemberFlags = "40"
    InternalDateTimeMs = mDateTimeMs
End Property

Friend Property Get InternalDateTimeOffsetMinutes() As Long
Attribute InternalDateTimeOffsetMinutes.VB_MemberFlags = "40"
    InternalDateTimeOffsetMinutes = mDateTimeOffsetMinutes
End Property

Friend Property Get InternalTimeSpanMs() As Double
Attribute InternalTimeSpanMs.VB_MemberFlags = "40"
    InternalTimeSpanMs = mTimeSpanMs
End Property

Friend Property Get InternalOperator() As Long
Attribute InternalOperator.VB_MemberFlags = "40"
    InternalOperator = mOperator
End Property

Public Property Get TimeoutError() As Long
Attribute TimeoutError.VB_Description = "Gets the asynchronous timeout error number."
    TimeoutError = ERROR_TIMEOUT
End Property

Public Property Get OperationCanceledError() As Long
Attribute OperationCanceledError.VB_Description = "Gets the cancellation error number."
    OperationCanceledError = ERROR_CANCELED
End Property

Public Property Get ExceptionType() As String
Attribute ExceptionType.VB_Description = "Gets the structured exception type name."
    RequireError Me, "ExceptionType"
    If Len(mName) = 0 Then
        ExceptionType = "Exception"
    Else
        ExceptionType = mName
    End If
End Property

Public Property Get InnerExceptions() As ROneCOne
Attribute InnerExceptions.VB_Description = "Gets AggregateException inner exceptions."
    RequireError Me, "InnerExceptions"
    If mItemsCount = 0 Then
        Set InnerExceptions = ROneCOne.ListOf(ROneCOne.Exception)
    Else
        Set InnerExceptions = CreateObjectListSnapshot(ArrSnapshot(mItems, mItemsCount))
    End If
End Property

Public Function Flatten() As ROneCOne
Attribute Flatten.VB_Description = "Flattens nested AggregateException values."
    Dim flattened As Collection

    RequireError Me, "Flatten"
    Set flattened = New Collection
    AppendFlattenedExceptions Me, flattened
    Set Flatten = CreateAggregateException(flattened)
End Function

Public Sub Handle(ByVal predicate As ROneCOne)
Attribute Handle.VB_Description = "Handles matching inner exceptions or raises the remainder."
    Dim arguments As Collection
    Dim errorValue As ROneCOne
    Dim flattened As ROneCOne
    Dim result As ROneCOne
    Dim unhandled As Collection

    RequireError Me, "Handle"
    RequireDelegate predicate, "Handle"
    If predicate.IsAction Or (predicate.Arity <> 1 And predicate.Arity <> -1) Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "Handle", _
            "Handle requires a Func<Exception, Boolean>."
    End If
    Set unhandled = New Collection
    Set flattened = Flatten
    For Each errorValue In flattened.InternalInnerExceptions
        Set arguments = New Collection
        arguments.Add CreateObjectResult(errorValue)
        Set result = predicate.InvokeValue(arguments)
        If Not CBool(RequireScalar(result, "Handle")) Then
            unhandled.Add errorValue
        End If
    Next errorValue
    If unhandled.Count > 0 Then RaiseCapturedError _
        CreateAggregateException(unhandled)
End Sub

Friend Property Get InternalSource() As ROneCOne
Attribute InternalSource.VB_MemberFlags = "40"
    Set InternalSource = mSource
End Property

Friend Property Get InternalExpressionOption() As Long
Attribute InternalExpressionOption.VB_MemberFlags = "40"
    InternalExpressionOption = mExpressionOption
End Property

Friend Property Get InternalLeft() As ROneCOne
Attribute InternalLeft.VB_MemberFlags = "40"
    Set InternalLeft = mLeft
End Property

Friend Property Get InternalRight() As ROneCOne
Attribute InternalRight.VB_MemberFlags = "40"
    Set InternalRight = mRight
End Property

Friend Property Get InternalHasObjectValue() As Boolean
Attribute InternalHasObjectValue.VB_MemberFlags = "40"
    InternalHasObjectValue = mHasObjectValue
End Property

Friend Property Get InternalObjectValue() As Object
Attribute InternalObjectValue.VB_MemberFlags = "40"
    Set InternalObjectValue = mObjectValue
End Property

Friend Property Get InternalScalarValue() As Variant
Attribute InternalScalarValue.VB_MemberFlags = "40"
    InternalScalarValue = mScalarValue
End Property

Friend Property Get InternalExpectedType() As VbVarType
Attribute InternalExpectedType.VB_MemberFlags = "40"
    InternalExpectedType = mExpectedType
End Property

Friend Property Get InternalExpectedObjectType() As String
Attribute InternalExpectedObjectType.VB_MemberFlags = "40"
    InternalExpectedObjectType = mExpectedObjectType
End Property

Friend Property Get InternalIsByRef() As Boolean
Attribute InternalIsByRef.VB_MemberFlags = "40"
    InternalIsByRef = mIsByRef
End Property

Friend Property Get InternalNext() As ROneCOne
Attribute InternalNext.VB_MemberFlags = "40"
    Set InternalNext = mNext
End Property

Friend Property Get InternalParameters() As Collection
Attribute InternalParameters.VB_MemberFlags = "40"
    Set InternalParameters = mParameters
End Property

Friend Property Get InternalTarget() As Object
Attribute InternalTarget.VB_MemberFlags = "40"
    Set InternalTarget = mTarget
End Property

Friend Property Get InternalMethodName() As String
Attribute InternalMethodName.VB_MemberFlags = "40"
    InternalMethodName = mMethodName
End Property

Friend Property Get InternalQueryDelegate() As ROneCOne
Attribute InternalQueryDelegate.VB_MemberFlags = "40"
    Set InternalQueryDelegate = mQueryDelegate
End Property

Friend Property Get InternalQueryValue() As ROneCOne
Attribute InternalQueryValue.VB_MemberFlags = "40"
    Set InternalQueryValue = mQueryValue
End Property

Friend Property Get InternalQueryComparer() As ROneCOne
Attribute InternalQueryComparer.VB_MemberFlags = "40"
    Set InternalQueryComparer = mQueryComparer
End Property

Friend Property Get InternalCollectionKind() As Long
Attribute InternalCollectionKind.VB_MemberFlags = "40"
    InternalCollectionKind = mCollectionKind
End Property

Friend Property Get InternalCollectionItems() As Collection
Attribute InternalCollectionItems.VB_MemberFlags = "40"
    Set InternalCollectionItems = ArrSnapshot(mItems, mItemsCount)
End Property

Friend Property Get InternalCollectionKeys() As Collection
Attribute InternalCollectionKeys.VB_MemberFlags = "40"
    Set InternalCollectionKeys = ArrSnapshot(mKeys, mKeysCount)
End Property

Friend Property Get InternalCollectionPriorities() As Collection
Attribute InternalCollectionPriorities.VB_MemberFlags = "40"
    Set InternalCollectionPriorities = ArrSnapshot(mPriorities, mPrioritiesCount)
End Property

Friend Property Get InternalKeyType() As VbVarType
Attribute InternalKeyType.VB_MemberFlags = "40"
    InternalKeyType = mKeyType
End Property

Friend Property Get InternalKeyObjectType() As String
Attribute InternalKeyObjectType.VB_MemberFlags = "40"
    InternalKeyObjectType = mKeyObjectType
End Property

Friend Property Get InternalKeySelector() As ROneCOne
Attribute InternalKeySelector.VB_MemberFlags = "40"
    Set InternalKeySelector = mKeySelector
End Property

Friend Property Get InternalCollectionChanged() As ROneCOne
Attribute InternalCollectionChanged.VB_MemberFlags = "40"
    Set InternalCollectionChanged = mCollectionChanged
End Property

Friend Property Get InternalCollectionCapacity() As Long
Attribute InternalCollectionCapacity.VB_MemberFlags = "40"
    InternalCollectionCapacity = mQueryCount
End Property

Friend Property Get InternalCollectionVersion() As Long
Attribute InternalCollectionVersion.VB_MemberFlags = "40"
    InternalCollectionVersion = mVersion
End Property

Friend Property Get InternalName() As String
Attribute InternalName.VB_MemberFlags = "40"
    InternalName = mName
End Property

Friend Property Get InternalDataColumns() As Collection
Attribute InternalDataColumns.VB_MemberFlags = "40"
    Set InternalDataColumns = ArrSnapshot(mKeys, mKeysCount)
End Property

Friend Property Get InternalDataRows() As Collection
Attribute InternalDataRows.VB_MemberFlags = "40"
    Set InternalDataRows = ArrSnapshot(mItems, mItemsCount)
End Property

Friend Property Get InternalDataTables() As Collection
Attribute InternalDataTables.VB_MemberFlags = "40"
    Set InternalDataTables = mDataTables
End Property

Friend Property Get InternalDataRelations() As Collection
Attribute InternalDataRelations.VB_MemberFlags = "40"
    Set InternalDataRelations = mDataRelations
End Property

Friend Property Get InternalEnforceConstraints() As Boolean
Attribute InternalEnforceConstraints.VB_MemberFlags = "40"
    InternalEnforceConstraints = mUnique
End Property

Friend Property Get InternalUnique() As Boolean
Attribute InternalUnique.VB_MemberFlags = "40"
    InternalUnique = mUnique
End Property

Friend Property Get InternalAllowDBNull() As Boolean
Attribute InternalAllowDBNull.VB_MemberFlags = "40"
    InternalAllowDBNull = mAllowDBNull
End Property

Friend Property Get InternalRelationParentColumn() As ROneCOne
Attribute InternalRelationParentColumn.VB_MemberFlags = "40"
    Set InternalRelationParentColumn = mLeft
End Property

Friend Property Get InternalRelationParentColumns() As Collection
Attribute InternalRelationParentColumns.VB_MemberFlags = "40"
    Set InternalRelationParentColumns = mRelationParentColumns
End Property

Friend Property Get InternalRelationChildColumn() As ROneCOne
Attribute InternalRelationChildColumn.VB_MemberFlags = "40"
    Set InternalRelationChildColumn = mRight
End Property

Friend Property Get InternalRelationChildColumns() As Collection
Attribute InternalRelationChildColumns.VB_MemberFlags = "40"
    Set InternalRelationChildColumns = mRelationChildColumns
End Property

Friend Property Get InternalProviderObject() As Object
Attribute InternalProviderObject.VB_MemberFlags = "40"
    Set InternalProviderObject = mProviderObject
End Property

Friend Property Get InternalProviderType() As Long
Attribute InternalProviderType.VB_MemberFlags = "40"
    InternalProviderType = mProviderType
End Property

Friend Property Get InternalProviderSize() As Long
Attribute InternalProviderSize.VB_MemberFlags = "40"
    InternalProviderSize = mProviderSize
End Property

Friend Property Get InternalProviderPrecision() As Byte
Attribute InternalProviderPrecision.VB_MemberFlags = "40"
    InternalProviderPrecision = mProviderPrecision
End Property

Friend Property Get InternalProviderScale() As Byte
Attribute InternalProviderScale.VB_MemberFlags = "40"
    InternalProviderScale = mProviderScale
End Property

Friend Property Get InternalSourceColumn() As String
Attribute InternalSourceColumn.VB_MemberFlags = "40"
    InternalSourceColumn = mSourceColumn
End Property

Friend Property Get InternalSourceVersion() As String
Attribute InternalSourceVersion.VB_MemberFlags = "40"
    InternalSourceVersion = mSourceVersion
End Property

Friend Property Get InternalElementType() As VbVarType
Attribute InternalElementType.VB_MemberFlags = "40"
    If IsDictionaryCollection Then
        InternalElementType = vbObject
    Else
        InternalElementType = mElementType
    End If
End Property

Friend Property Get InternalElementObjectType() As String
Attribute InternalElementObjectType.VB_MemberFlags = "40"
    If IsDictionaryCollection Then
        InternalElementObjectType = "ROneCOne"
    Else
        InternalElementObjectType = mElementObjectType
    End If
End Property

Friend Property Get InternalValueType() As VbVarType
Attribute InternalValueType.VB_MemberFlags = "40"
    InternalValueType = mElementType
End Property

Friend Property Get InternalValueObjectType() As String
Attribute InternalValueObjectType.VB_MemberFlags = "40"
    InternalValueObjectType = mElementObjectType
End Property

Friend Property Get InternalMethodArity() As Long
Attribute InternalMethodArity.VB_MemberFlags = "40"
    InternalMethodArity = mMethodArity
End Property

Friend Property Get InternalIsActionDelegate() As Boolean
Attribute InternalIsActionDelegate.VB_MemberFlags = "40"
    InternalIsActionDelegate = mIsActionDelegate
End Property

Friend Property Get InternalHasParameterSignature() As Boolean
Attribute InternalHasParameterSignature.VB_MemberFlags = "40"
    InternalHasParameterSignature = mHasParameterSignature
End Property

Friend Property Get InternalHasReturnSignature() As Boolean
Attribute InternalHasReturnSignature.VB_MemberFlags = "40"
    InternalHasReturnSignature = mHasReturnSignature
End Property

Friend Property Get InternalReturnType() As VbVarType
Attribute InternalReturnType.VB_MemberFlags = "40"
    InternalReturnType = mReturnType
End Property

Friend Property Get InternalReturnObjectType() As String
Attribute InternalReturnObjectType.VB_MemberFlags = "40"
    InternalReturnObjectType = mReturnObjectType
End Property

Friend Property Get InternalInvocationList() As Collection
Attribute InternalInvocationList.VB_MemberFlags = "40"
    Set InternalInvocationList = mInvocationList
End Property

Friend Property Get InternalNativeAddress() As LongPtr
Attribute InternalNativeAddress.VB_MemberFlags = "40"
    InternalNativeAddress = mNativeAddress
End Property

Friend Property Get InternalReferenceAddress() As LongPtr
Attribute InternalReferenceAddress.VB_MemberFlags = "40"
    InternalReferenceAddress = mReferenceAddress
End Property

Friend Property Get InternalTryBody() As ROneCOne
Attribute InternalTryBody.VB_MemberFlags = "40"
    Set InternalTryBody = mTryBody
End Property

Friend Property Get InternalCatchActions() As Collection
Attribute InternalCatchActions.VB_MemberFlags = "40"
    Set InternalCatchActions = mCatchActions
End Property

Friend Property Get InternalCatchNumbers() As Collection
Attribute InternalCatchNumbers.VB_MemberFlags = "40"
    Set InternalCatchNumbers = mCatchNumbers
End Property

Friend Property Get InternalFinallyAction() As ROneCOne
Attribute InternalFinallyAction.VB_MemberFlags = "40"
    Set InternalFinallyAction = mFinallyAction
End Property

Friend Sub ConfigureParameter( _
    ByVal expectedType As VbVarType, _
    ByVal expectedObjectType As String, _
    Optional ByVal isByRef As Boolean = False _
)
    mRole = ROLE_PARAMETER
    mExpectedType = expectedType
    mExpectedObjectType = expectedObjectType
    mIsByRef = isByRef
End Sub

Friend Sub ConfigureScalarValue(ByVal scalarValue As Variant)
    mRole = ROLE_VALUE
    mScalarValue = scalarValue
    mHasObjectValue = False
End Sub

Friend Sub ConfigureObjectValue(ByVal objectValue As Object)
    mRole = ROLE_VALUE
    Set mObjectValue = objectValue
    mHasObjectValue = True
End Sub

Friend Sub ConfigureExpression( _
    ByVal expressionOperator As Long, _
    ByVal leftValue As ROneCOne, _
    ByVal rightValue As ROneCOne, _
    Optional ByVal expressionOption As Long = 0 _
)
    mRole = ROLE_EXPRESSION
    mOperator = expressionOperator
    mExpressionOption = expressionOption
    Set mLeft = leftValue
    Set mRight = rightValue
End Sub

Friend Sub ConfigureLambda( _
    ByVal body As ROneCOne, _
    ByVal parameters As Collection _
)
    mRole = ROLE_LAMBDA
    Set mLeft = body
    Set mParameters = parameters
    mMethodArity = parameters.Count
    mHasParameterSignature = True
    mIsActionDelegate = False
End Sub

Friend Sub ConfigureMemberExpression( _
    ByVal expressionOperator As Long, _
    ByVal leftValue As ROneCOne, _
    ByVal memberName As String _
)
    mRole = ROLE_EXPRESSION
    mOperator = expressionOperator
    Set mLeft = leftValue
    mMethodName = memberName
End Sub

Friend Sub ConfigureMembershipExpression( _
    ByVal expressionOperator As Long, _
    ByVal leftValue As ROneCOne, _
    ByVal sequenceValue As ROneCOne, _
    Optional ByVal comparer As ROneCOne _
)
    mRole = ROLE_EXPRESSION
    mOperator = expressionOperator
    Set mLeft = leftValue
    Set mQueryValue = sequenceValue
    Set mQueryComparer = comparer
End Sub

Friend Sub ConfigureQuantifierExpression( _
    ByVal expressionOperator As Long, _
    ByVal leftValue As ROneCOne, _
    ByVal predicate As ROneCOne _
)
    mRole = ROLE_EXPRESSION
    mOperator = expressionOperator
    Set mLeft = leftValue
    Set mQueryDelegate = predicate
End Sub

Friend Sub ConfigureMethod( _
    ByVal target As Object, _
    ByVal methodName As String, _
    ByVal isAction As Boolean _
)
    mRole = ROLE_METHOD
    Set mTarget = target
    mMethodName = methodName
    mMethodArity = -1
    mIsActionDelegate = isAction
    If isAction Then
        mHasReturnSignature = True
        mReturnType = vbEmpty
    End If
End Sub

Friend Sub ConfigureProcedure(ByVal procedureName As String, ByVal isAction As Boolean)
    mRole = ROLE_PROCEDURE
    mMethodName = procedureName
    mMethodArity = -1
    mIsActionDelegate = isAction
    If isAction Then
        mHasReturnSignature = True
        mReturnType = vbEmpty
    End If
End Sub

Friend Sub ConfigureNative(ByVal procedureAddress As LongPtr, ByVal isAction As Boolean)
    mRole = ROLE_NATIVE
    mNativeAddress = procedureAddress
    mMethodArity = -1
    mIsActionDelegate = isAction
    If isAction Then
        mHasReturnSignature = True
        mReturnType = vbEmpty
    End If
End Sub

Friend Sub ConfigureReference(ByVal variableAddress As LongPtr, ByVal valueType As VbVarType)
    mRole = ROLE_REFERENCE
    mReferenceAddress = variableAddress
    mExpectedType = valueType
End Sub

Friend Sub ConfigureManagedReference( _
    ByVal valueType As VbVarType, _
    ByVal objectType As String _
)
    mRole = ROLE_REFERENCE
    mExpectedType = valueType
    mExpectedObjectType = objectType
    mReferenceAddress = 0
End Sub

Friend Sub ConfigureManagedReferenceValue(ByVal wrapped As ROneCOne)
    Set mQueryValue = wrapped
End Sub

Friend Sub ConfigureEvent(ByVal descriptors As Collection)
    mRole = ROLE_EVENT
    Set mParameters = descriptors
    mMethodArity = descriptors.Count
    mHasParameterSignature = True
    mHasReturnSignature = True
    mIsActionDelegate = True
    mReturnType = vbEmpty
    Set mInvocationList = New Collection
End Sub

Friend Sub ConfigureTry(ByVal work As ROneCOne)
    mRole = ROLE_TRY
    Set mTryBody = work
    Set mCatchActions = New Collection
    Set mCatchNumbers = New Collection
End Sub

Friend Sub ConfigureCatch(ByVal errorNumber As Long, ByVal handler As ROneCOne)
    mCatchNumbers.Add errorNumber
    mCatchActions.Add handler
End Sub

Friend Sub ConfigureFinally(ByVal cleanup As ROneCOne)
    Set mFinallyAction = cleanup
End Sub

Friend Sub ConfigureError( _
    ByVal errorNumber As Long, _
    ByVal errorSource As String, _
    ByVal errorDescription As String, _
    ByVal errorHelpFile As String, _
    ByVal errorHelpContext As Long _
)
    mRole = ROLE_ERROR
    mName = "Exception"
    mErrorNumber = errorNumber
    mErrorSource = errorSource
    mErrorDescription = errorDescription
    mErrorHelpFile = errorHelpFile
    mErrorHelpContext = errorHelpContext
End Sub

Friend Sub ConfigureParameterSignature(ByVal descriptors As Collection)
    Set mParameters = descriptors
    mMethodArity = descriptors.Count
    mHasParameterSignature = True
End Sub

Friend Sub ConfigureReturnSignature( _
    ByVal valueType As VbVarType, _
    ByVal objectType As String _
)
    mReturnType = valueType
    mReturnObjectType = objectType
    mHasReturnSignature = True
    mIsActionDelegate = False
End Sub

Friend Sub ConfigureActionSignature(ByVal isAction As Boolean)
    mIsActionDelegate = isAction
    If isAction Then
        mHasReturnSignature = True
        mReturnType = vbEmpty
        mReturnObjectType = vbNullString
    End If
End Sub

Friend Sub ConfigureMulticast(ByVal invocationList As Collection)
    Dim firstDelegate As ROneCOne

    mRole = ROLE_MULTICAST
    Set mInvocationList = invocationList
    Set firstDelegate = invocationList.Item(1)
    CopySignatureFrom firstDelegate
End Sub

Friend Sub ConfigureComposition( _
    ByVal firstDelegate As ROneCOne, _
    ByVal nextDelegate As ROneCOne _
)
    mRole = ROLE_COMPOSITION
    Set mLeft = firstDelegate
    Set mNext = nextDelegate
    mMethodArity = firstDelegate.Arity
    mIsActionDelegate = nextDelegate.IsAction
End Sub

Friend Sub ConfigureList( _
    ByVal elementType As VbVarType, _
    ByVal elementObjectType As String _
)
    mRole = ROLE_LIST
    mElementType = elementType
    mElementObjectType = elementObjectType
    ArrReset mItems, mItemsCount
    Set mEnumerationValues = Nothing
    mEnumerationVersion = 0
End Sub

Friend Sub ConfigureAggregateError(ByVal errors As Collection)
    Dim errorValue As ROneCOne

    mRole = ROLE_ERROR
    mName = "AggregateException"
    mErrorNumber = ERROR_INVALID_OPERATION
    mErrorSource = "ROneCOne.AggregateException"
    mErrorDescription = CStr(errors.Count) & _
        " exception(s) occurred. Inspect InnerExceptions for details."
    ArrReset mItems, mItemsCount
    For Each errorValue In errors
        ArrAppend mItems, mItemsCount, errorValue
    Next errorValue
End Sub

Friend Sub ConfigureExceptionType(ByVal exceptionTypeName As String)
    mName = exceptionTypeName
End Sub

Friend Property Get InternalInnerExceptions() As Collection
Attribute InternalInnerExceptions.VB_MemberFlags = "40"
    Set InternalInnerExceptions = ArrSnapshot(mItems, mItemsCount)
End Property

Friend Sub ConfigureGenericCollection( _
    ByVal roleValue As Long, _
    ByVal collectionKind As Long, _
    ByVal elementType As VbVarType, _
    ByVal elementObjectType As String, _
    Optional ByVal keyType As VbVarType = vbEmpty, _
    Optional ByVal keyObjectType As String = vbNullString, _
    Optional ByVal comparer As ROneCOne _
)
    mRole = roleValue
    mCollectionKind = collectionKind
    mElementType = elementType
    mElementObjectType = elementObjectType
    mKeyType = keyType
    mKeyObjectType = keyObjectType
    Set mQueryComparer = comparer
    ArrReset mItems, mItemsCount
    ArrReset mKeys, mKeysCount
    ArrReset mPriorities, mPrioritiesCount
    Set mEnumerationValues = Nothing
    If HashIndexSupported Then RebuildHashIndex 8
End Sub

Friend Sub ConfigureReadOnlySource(ByVal source As ROneCOne)
    Set mSource = source
End Sub

Friend Sub ConfigureCollectionKeySelector(ByVal selector As ROneCOne)
    Set mKeySelector = selector
End Sub

Friend Sub ConfigureObservableCollection()
    Set mCollectionChanged = ROneCOne.EventOf(vbString)
End Sub

Friend Sub ConfigureCollectionCapacity(ByVal boundedCapacity As Long)
    mQueryCount = boundedCapacity
End Sub

Friend Sub ConfigureCollectionRole(ByVal roleValue As Long)
    mRole = roleValue
End Sub

Friend Sub ConfigureCollectionVersion(ByVal version As Long)
    mVersion = version
End Sub

Friend Sub ConfigureCollectionNode( _
    ByVal wrapped As ROneCOne, _
    ByVal owner As ROneCOne _
)
    mRole = ROLE_NODE
    Set mQueryValue = wrapped
    Set mSource = owner
End Sub

Friend Sub ConfigureDictionaryEntry( _
    ByVal keyWrapped As ROneCOne, _
    ByVal valueWrapped As ROneCOne _
)
    mRole = ROLE_ENTRY
    Set mLeft = keyWrapped
    Set mQueryValue = valueWrapped
End Sub

Friend Sub ConfigureTaskFactory()
    mRole = ROLE_TASK_FACTORY
End Sub

Friend Sub ConfigureTask( _
    ByVal taskKind As Long, _
    ByVal taskBody As ROneCOne, _
    ByVal cancellationToken As ROneCOne _
)
    mRole = ROLE_TASK
    mOperator = taskKind
    Set mTaskBody = taskBody
    Set mCancellationSource = cancellationToken
    Set mTaskChildren = New Collection
    mTaskStatus = TASK_STATUS_CREATED
End Sub

Friend Sub ConfigureTaskArguments(ByVal arguments As Collection)
    Set mTaskChildren = arguments
End Sub

Friend Sub ConfigureTaskDeadline(ByVal deadline As LongLong)
    mTaskDeadline = deadline
End Sub

Friend Sub ConfigureProviderTask( _
    ByVal providerTarget As ROneCOne, _
    ByVal methodName As String, _
    ByVal arguments As Collection, _
    ByVal cancellationToken As ROneCOne _
)
    ConfigureTask TASK_PROVIDER_CALL, Nothing, cancellationToken
    Set mSource = providerTarget
    mMethodName = methodName
    Set mTaskChildren = arguments
End Sub

Friend Sub ConfigureDbAsyncTask( _
    ByVal pendingSource As ROneCOne, _
    ByVal resultMode As String, _
    ByVal arguments As Collection, _
    ByVal cancellationToken As ROneCOne _
)
    ConfigureTask TASK_DB_ASYNC, Nothing, cancellationToken
    Set mSource = pendingSource
    mMethodName = resultMode
    If arguments Is Nothing Then Set arguments = New Collection
    Set mTaskChildren = arguments
End Sub

Friend Sub ConfigureHttpClient()
    mRole = ROLE_HTTP_CLIENT
    mHttpTimeoutMs = HTTP_DEFAULT_TIMEOUT_MS
    Set mHttpHeaderNames = New Collection
    Set mHttpHeaderValues = New Collection
End Sub

Friend Sub ConfigureJsonFactory()
    mRole = ROLE_JSON_FACTORY
End Sub

Friend Sub ConfigureFileFactory()
    mRole = ROLE_FILE_FACTORY
End Sub

Friend Sub ConfigureDirectoryFactory()
    mRole = ROLE_DIRECTORY_FACTORY
End Sub

Friend Sub ConfigurePathFactory()
    mRole = ROLE_PATH_FACTORY
End Sub

Friend Sub ConfigureCsvFactory()
    mRole = ROLE_CSV_FACTORY
End Sub

Friend Sub ConfigureProcessFactory()
    mRole = ROLE_PROCESS_FACTORY
End Sub

Friend Sub ConfigureRegex( _
    ByVal pattern As String, _
    ByVal ignoreCase As Boolean, _
    ByVal multiLine As Boolean _
)
    mRole = ROLE_REGEX
    mName = pattern
    Set mProviderObject = CreateObject(REGEX_PROG_ID)
    mProviderObject.Pattern = pattern
    mProviderObject.IgnoreCase = ignoreCase
    mProviderObject.MultiLine = multiLine
    ' The script engine compiles lazily, so a throwaway test surfaces an
    ' invalid pattern here, at creation, instead of at first use.
    On Error GoTo BadPattern
    mProviderObject.Test vbNullString
    On Error GoTo 0
    Exit Sub

BadPattern:
    Set mProviderObject = Nothing
    Err.Raise ERROR_REGEX, "ROneCOne.RegexException", _
        "The regular expression pattern '" & pattern & "' is not valid."
End Sub

Friend Sub ConfigureRegexMatch( _
    ByVal succeeded As Boolean, _
    ByVal valueText As String, _
    ByVal matchIndex As Long, _
    ByVal matchLength As Long, _
    ByVal groups As ROneCOne _
)
    mRole = ROLE_REGEX_MATCH
    mRegexSuccess = succeeded
    mName = valueText
    mRegexIndex = matchIndex
    mRegexLength = matchLength
    Set mSource = groups
End Sub

Friend Sub ConfigureHashFactory()
    mRole = ROLE_HASH_FACTORY
End Sub

Friend Sub ConfigureConvertFactory()
    mRole = ROLE_CONVERT_FACTORY
End Sub

Friend Sub ConfigureDateTimeFactory()
    mRole = ROLE_DATETIME_FACTORY
End Sub

Friend Sub ConfigureDateTime(ByVal msValue As Double, ByVal offsetMinutes As Long)
    mRole = ROLE_DATETIME
    mDateTimeMs = msValue
    mDateTimeOffsetMinutes = offsetMinutes
End Sub

Friend Sub ConfigureTimeSpanFactory()
    mRole = ROLE_TIMESPAN_FACTORY
End Sub

Friend Sub ConfigureTimeSpan(ByVal msValue As Double)
    mRole = ROLE_TIMESPAN
    mTimeSpanMs = msValue
End Sub

Friend Sub ConfigureStringsFactory()
    mRole = ROLE_STRINGS_FACTORY
End Sub

Friend Sub ConfigureStringBuilder()
    mRole = ROLE_STRING_BUILDER
    JsonBuilderInit mStringBuilder, 256
End Sub

Friend Sub ConfigureGuidFactory()
    mRole = ROLE_GUID_FACTORY
End Sub

Friend Sub ConfigureRandomFactory()
    mRole = ROLE_RANDOM_FACTORY
End Sub

Friend Sub ConfigureXmlFactory()
    mRole = ROLE_XML_FACTORY
End Sub

Friend Sub ConfigureXmlNode(ByVal nodeObject As Object)
    mRole = ROLE_XML_NODE
    Set mProviderObject = nodeObject
End Sub

Friend Sub ConfigureUriFactory()
    mRole = ROLE_URI_FACTORY
End Sub

Friend Sub ConfigureLogger(ByVal filePath As String, ByVal levelValue As Long)
    mRole = ROLE_LOGGER
    mName = filePath
    mLoggerMinimumLevel = levelValue
End Sub

Friend Sub ConfigureFileWatcher( _
    ByVal folderPath As String, _
    ByVal filterPattern As String _
)
    mRole = ROLE_FILE_WATCHER
    mName = folderPath
    mWatcherFilter = filterPattern
End Sub

Friend Sub ConfigureFileChange( _
    ByVal changeType As String, _
    ByVal fileName As String _
)
    mRole = ROLE_FILE_CHANGE
    mWatcherChangeType = changeType
    mName = fileName
End Sub

Friend Sub ConfigureWatchTask( _
    ByVal watcher As ROneCOne, _
    ByVal baseline As Collection, _
    ByVal cancellationToken As ROneCOne _
)
    ConfigureTask TASK_FILE_WATCH, Nothing, cancellationToken
    Set mSource = watcher
    Set mWatcherSnapshot = baseline
End Sub

Friend Property Get InternalWatcherFolder() As String
Attribute InternalWatcherFolder.VB_MemberFlags = "40"
    InternalWatcherFolder = mName
End Property

Friend Property Get InternalWatcherFilter() As String
Attribute InternalWatcherFilter.VB_MemberFlags = "40"
    InternalWatcherFilter = mWatcherFilter
End Property

Friend Sub ConfigureZipFileFactory()
    mRole = ROLE_ZIPFILE_FACTORY
End Sub

Friend Sub ConfigureZipArchive( _
    ByVal archivePath As String, _
    ByVal entryObjects As Collection _
)
    mRole = ROLE_ZIP_ARCHIVE
    mName = archivePath
    Set mZipEntries = entryObjects
End Sub

Friend Sub ConfigureZipEntry( _
    ByVal archivePath As String, _
    ByVal entryName As String, _
    ByVal method As Long, _
    ByVal crcValue As Double, _
    ByVal compressedSize As Long, _
    ByVal size As Long, _
    ByVal localOffset As Long, _
    ByVal isDirectory As Boolean _
)
    mRole = ROLE_ZIP_ENTRY
    mZipArchivePath = archivePath
    mName = entryName
    mZipEntryMethod = method
    mZipEntryCrc = crcValue
    mZipEntryCompressedSize = compressedSize
    mZipEntrySize = size
    mZipEntryOffset = localOffset
    mZipEntryIsDirectory = isDirectory
End Sub

Friend Sub ConfigureWebUtilityFactory()
    mRole = ROLE_WEB_UTILITY_FACTORY
End Sub

Friend Sub ConfigureProcessResult( _
    ByVal execObject As Object, _
    ByVal outputPath As String, _
    ByVal errorPath As String _
)
    mRole = ROLE_PROCESS_RESULT
    Set mProviderObject = execObject
    mProcessOutputPath = outputPath
    mProcessErrorPath = errorPath
End Sub

Friend Sub ConfigureProcessTask( _
    ByVal pendingResult As ROneCOne, _
    ByVal cancellationToken As ROneCOne _
)
    ConfigureTask TASK_PROCESS_RUN, Nothing, cancellationToken
    Set mSource = pendingResult
End Sub

Friend Sub ConfigureHttpResponse(ByVal request As Object)
    mRole = ROLE_HTTP_RESPONSE
    Set mProviderObject = request
End Sub

Friend Sub ConfigureHttpTask( _
    ByVal response As ROneCOne, _
    ByVal resultMode As String, _
    ByVal cancellationToken As ROneCOne _
)
    ConfigureTask TASK_HTTP_SEND, Nothing, cancellationToken
    Set mSource = response
    mMethodName = resultMode
End Sub

Friend Function InternalHttpPoll(ByVal captureBytes As Boolean) As Boolean
    Dim capturedDescription As String
    Dim capturedNumber As Long
    Dim capturedSource As String

    ' The response was already captured, so a completed task can keep asking.
    If mProviderObject Is Nothing Then
        InternalHttpPoll = True
        Exit Function
    End If
    ' WaitForResponse(0) is a non-blocking poll: False while the transfer is
    ' still in flight, True once the response arrived, and a raised error for
    ' hard failures such as name resolution or connection loss.
    On Error GoTo PollFailed
    If Not mProviderObject.WaitForResponse(0) Then Exit Function
    mHttpStatusCode = CLng(mProviderObject.Status)
    mHttpReasonPhrase = CStr(mProviderObject.StatusText)
    mHttpRawHeaders = CStr(mProviderObject.GetAllResponseHeaders)
    mHttpContent = CStr(mProviderObject.ResponseText)
    If captureBytes Then mHttpBytes = mProviderObject.ResponseBody
    Set mProviderObject = Nothing
    On Error GoTo 0
    InternalHttpPoll = True
    Exit Function

PollFailed:
    capturedNumber = Err.Number
    capturedSource = Err.Source
    capturedDescription = Err.Description
    On Error GoTo 0
    InternalHttpAbort
    Err.Raise capturedNumber, "ROneCOne.HttpRequestException", _
        "The HTTP request failed: " & capturedDescription
End Function

Friend Function InternalProcessPoll() As Boolean
    Dim errorText As String
    Dim outputText As String

    ' A completed result keeps answering True so a finished task can keep
    ' asking. While the process runs, Status stays 0; after it exits, the
    ' redirect files can stay locked for a beat, so a failed read simply
    ' polls again instead of blocking or faulting.
    If mProviderObject Is Nothing Then
        InternalProcessPoll = True
        Exit Function
    End If
    If CLng(mProviderObject.Status) = 0 Then Exit Function
    If Not TryReadProcessFile(mProcessOutputPath, outputText) Then
        Exit Function
    End If
    If Not TryReadProcessFile(mProcessErrorPath, errorText) Then
        Exit Function
    End If
    mProcessExitCode = CLng(mProviderObject.ExitCode)
    mProcessOutput = outputText
    mProcessError = errorText
    Set mProviderObject = Nothing
    DeleteProcessScratchFiles
    InternalProcessPoll = True
End Function

Friend Sub InternalProcessAbort()
    If mProviderObject Is Nothing Then Exit Sub
    On Error Resume Next
    mProviderObject.Terminate
    On Error GoTo 0
    Set mProviderObject = Nothing
    DeleteProcessScratchFiles
End Sub

Private Sub DeleteProcessScratchFiles()
    On Error Resume Next
    If Len(mProcessOutputPath) > 0 Then Kill mProcessOutputPath
    If Len(mProcessErrorPath) > 0 Then Kill mProcessErrorPath
    On Error GoTo 0
    mProcessOutputPath = vbNullString
    mProcessErrorPath = vbNullString
End Sub

Private Function TryReadProcessFile( _
    ByVal filePath As String, _
    ByRef content As String _
) As Boolean
    Dim stream As Object
    Dim streamText As Variant

    ' Console redirection writes in the OEM/ANSI code page, so the capture
    ' decodes as windows-1252; ASCII output is always exact.
    On Error GoTo NotReadable
    Set stream = CreateStreamObject(STREAM_TYPE_TEXT)
    stream.Charset = "windows-1252"
    stream.LoadFromFile filePath
    streamText = stream.ReadText(-1)
    stream.Close
    On Error GoTo 0
    If IsNull(streamText) Then
        content = vbNullString
    Else
        content = CStr(streamText)
    End If
    TryReadProcessFile = True
    Exit Function

NotReadable:
    TryReadProcessFile = False
End Function

Friend Sub InternalHttpAbort()
    If mProviderObject Is Nothing Then Exit Sub
    On Error Resume Next
    mProviderObject.Abort
    On Error GoTo 0
    Set mProviderObject = Nothing
End Sub

Friend Property Get InternalHttpBytes() As Variant
Attribute InternalHttpBytes.VB_MemberFlags = "40"
    InternalHttpBytes = mHttpBytes
    ' The task's result now owns the only needed copy; releasing this one
    ' keeps a large download from being held twice until the response dies.
    mHttpBytes = Empty
End Property

Friend Sub ConfigureCompletedTask(ByVal result As ROneCOne)
    mRole = ROLE_TASK
    mOperator = TASK_COMPLETED
    Set mTaskResult = result
    Set mTaskChildren = New Collection
    mTaskStatus = TASK_STATUS_COMPLETED
End Sub

Friend Sub ConfigureTaskCompletionSource( _
    ByVal valueType As VbVarType, _
    ByVal objectType As String _
)
    Dim promise As ROneCOne

    mRole = ROLE_TASK_COMPLETION_SOURCE
    mExpectedType = valueType
    mExpectedObjectType = objectType
    Set promise = New ROneCOne
    promise.ConfigureTask TASK_PROMISE, Nothing, Nothing
    Set mSource = promise
End Sub

Friend Sub ConfigureProgress( _
    ByVal valueType As VbVarType, _
    ByVal objectType As String, _
    ByVal handler As ROneCOne _
)
    mRole = ROLE_PROGRESS
    mExpectedType = valueType
    mExpectedObjectType = objectType
    Set mTaskBody = handler
End Sub

Friend Sub ConfigureCancellationSource()
    mRole = ROLE_CANCELLATION_SOURCE
    mCancellationRequested = False
    mQueryCount = -1
    mTaskDeadline = -1
    Set mInvocationList = New Collection
End Sub

Friend Sub ConfigureCancellationToken(ByVal source As ROneCOne)
    mRole = ROLE_CANCELLATION_TOKEN
    Set mCancellationSource = source
End Sub

Friend Sub ConfigureCancellationRegistration( _
    ByVal source As ROneCOne, _
    ByVal callback As ROneCOne _
)
    mRole = ROLE_CANCELLATION_REGISTRATION
    Set mSource = source
    Set mQueryDelegate = callback
End Sub

Friend Property Get InternalTaskResult() As ROneCOne
Attribute InternalTaskResult.VB_MemberFlags = "40"
    Set InternalTaskResult = mTaskResult
End Property

Friend Property Get InternalTaskException() As ROneCOne
Attribute InternalTaskException.VB_MemberFlags = "40"
    Set InternalTaskException = mTaskException
End Property

Friend Property Get InternalTaskStatus() As Long
Attribute InternalTaskStatus.VB_MemberFlags = "40"
    InternalTaskStatus = mTaskStatus
End Property

Friend Property Get InternalTaskChildCount() As Long
Attribute InternalTaskChildCount.VB_MemberFlags = "40"
    InternalTaskChildCount = mTaskChildren.Count
End Property

Friend Property Get InternalRegistrationCallback() As ROneCOne
Attribute InternalRegistrationCallback.VB_MemberFlags = "40"
    Set InternalRegistrationCallback = mQueryDelegate
End Property

Friend Sub RegisterCancellationCallback(ByVal registration As ROneCOne)
    Dim arguments As Collection
    Dim ignored As ROneCOne

    If mCancellationRequested Then
        Set arguments = New Collection
        Set ignored = registration.InternalRegistrationCallback.InvokeValue(arguments)
    Else
        mInvocationList.Add registration.InternalRegistrationCallback
    End If
End Sub

Friend Sub UnregisterCancellationRegistration(ByVal registration As ROneCOne)
    Dim callback As ROneCOne
    Dim candidate As ROneCOne
    Dim index As Long

    Set callback = registration.InternalRegistrationCallback
    For index = mInvocationList.Count To 1 Step -1
        Set candidate = mInvocationList.Item(index)
        If candidate Is callback Then
            mInvocationList.Remove index
            Exit Sub
        End If
    Next index
End Sub

Friend Function TryCompletePromise( _
    ByVal result As ROneCOne, _
    ByVal terminalStatus As Long, _
    ByVal errorValue As ROneCOne _
) As Boolean
    If mRole <> ROLE_TASK Or mOperator <> TASK_PROMISE Then Exit Function
    If mTaskStatus >= TASK_STATUS_COMPLETED Then Exit Function
    Set mTaskResult = result
    Set mTaskException = errorValue
    mTaskStatus = terminalStatus
    TryCompletePromise = True
End Function

' The live rows of a table or a view, in the order a reader should see
' them. A table hides rows marked Deleted; a view has already applied its
' filter and sort, and hands back wrappers that still need unwrapping.
' Friend rather than Private because the JSON writer and the Excel Table
' writer both ask another instance for its rows, and a Friend member read
' through a typed local is the one binding VBA honors.
Friend Function InternalLiveDataRows() As Collection
    Dim liveRows As Collection
    Dim row As ROneCOne
    Dim wrapped As ROneCOne

    Set liveRows = New Collection
    If mRole = ROLE_DATA_VIEW Then
        For Each wrapped In MaterializeWrapped
            If wrapped.InternalHasObjectValue Then
                liveRows.Add wrapped.InternalObjectValue
            Else
                liveRows.Add wrapped
            End If
        Next wrapped
    Else
        RequireDataTable Me, "InternalLiveDataRows"
        For Each row In ArrSnapshot(mItems, mItemsCount)
            If row.RowState <> "Deleted" Then liveRows.Add row
        Next row
    End If
    Set InternalLiveDataRows = liveRows
End Function

Friend Sub InternalFillFromRange( _
    ByVal source As Object, _
    ByVal headers As Boolean, _
    ByVal memberName As String _
)
    Dim colCount As Long
    Dim colIndex As Long
    Dim columnName As String
    Dim firstDataRow As Long
    Dim grid As Variant
    Dim rowCount As Long
    Dim rowIndex As Long
    Dim values() As Variant

    RequireDataTable Me, memberName
    ReadRangeGrid ResolveTableRange(source, headers, memberName), _
        grid, rowCount, colCount, memberName
    ConfigureDataTable mName
    For colIndex = 1 To colCount
        columnName = vbNullString
        If headers Then columnName = Trim$(CStr(grid(1, colIndex)))
        If Len(columnName) = 0 Then columnName = "Column" & CStr(colIndex)
        Column columnName, vbVariant
    Next colIndex
    firstDataRow = IIf(headers, 2, 1)
    For rowIndex = firstDataRow To rowCount
        ReDim values(0 To colCount - 1)
        For colIndex = 1 To colCount
            values(colIndex - 1) = grid(rowIndex, colIndex)
        Next colIndex
        LoadRow values
    Next rowIndex
End Sub

Friend Sub AttachSourceTable(ByVal listObject As Object)
    Set mSourceTable = listObject
End Sub

Friend Sub ConfigureDataTable(ByVal tableName As String)
    mRole = ROLE_DATA_TABLE
    mName = tableName
    ArrReset mItems, mItemsCount
    ArrReset mKeys, mKeysCount
    Set mPrimaryKey = New Collection
    Set mDataIndex = ROneCOne.DictionaryOf(vbString, ROneCOne.Exception)
    Set mUniqueIndexes = New Collection
End Sub

Friend Sub ConfigureDBNull()
    mRole = ROLE_DB_NULL
    mName = "DBNull"
End Sub

Friend Sub ConfigureUsing(ByVal resource As ROneCOne)
    mRole = ROLE_USING
    Set mSource = resource
End Sub

Friend Sub ConfigureDataColumn( _
    ByVal columnName As String, _
    ByVal valueType As VbVarType, _
    ByVal objectType As String _
)
    mRole = ROLE_DATA_COLUMN
    mName = columnName
    mExpectedType = valueType
    mExpectedObjectType = objectType
    mAllowDBNull = True
End Sub

Friend Sub ConfigureDataSet(ByVal dataSetName As String)
    mRole = ROLE_DATA_SET
    mName = dataSetName
    Set mDataTables = New Collection
    Set mDataRelations = New Collection
    mUnique = True
End Sub

Friend Sub ConfigureDataRelation( _
    ByVal relationName As String, _
    ByVal parentColumns As Collection, _
    ByVal childColumns As Collection, _
    ByVal createConstraints As Boolean _
)
    mRole = ROLE_DATA_RELATION
    mName = relationName
    Set mRelationParentColumns = parentColumns
    Set mRelationChildColumns = childColumns
    Set mLeft = parentColumns.Item(1)
    Set mRight = childColumns.Item(1)
    mExpressionOption = Abs(CLng(createConstraints))
    mVersion = -1
    mEnumerationVersion = -1
End Sub

Friend Function FindRelatedRows( _
    ByVal sourceRow As ROneCOne, _
    ByVal findChildren As Boolean _
) As ROneCOne
    Dim group As ROneCOne
    Dim indexValue As ROneCOne
    Dim keyToken As String
    Dim sourceColumns As Collection

    EnsureRelationIndexes
    If findChildren Then
        Set sourceColumns = mRelationParentColumns
        Set indexValue = mRelationChildIndex
    Else
        Set sourceColumns = mRelationChildColumns
        Set indexValue = mRelationParentIndex
    End If
    keyToken = RelationKeyToken(sourceRow, sourceColumns)
    If indexValue.ContainsKey(keyToken) Then
        Set group = indexValue.Item(keyToken)
        Set FindRelatedRows = group.ToList
    Else
        Set FindRelatedRows = ROneCOne.ListOf(sourceRow)
    End If
End Function

Private Sub EnsureRelationIndexes()
    Dim childColumn As ROneCOne
    Dim childTable As ROneCOne
    Dim parentColumn As ROneCOne
    Dim parentTable As ROneCOne

    Set parentColumn = mRelationParentColumns.Item(1)
    Set childColumn = mRelationChildColumns.Item(1)
    Set parentTable = parentColumn.InternalSource
    Set childTable = childColumn.InternalSource
    If mVersion = parentTable.InternalSequenceVersion And _
        mEnumerationVersion = childTable.InternalSequenceVersion Then Exit Sub
    Set mRelationParentIndex = BuildRelationIndex( _
        parentTable, mRelationParentColumns)
    Set mRelationChildIndex = BuildRelationIndex( _
        childTable, mRelationChildColumns)
    mVersion = parentTable.InternalSequenceVersion
    mEnumerationVersion = childTable.InternalSequenceVersion
End Sub

Private Function BuildRelationIndex( _
    ByVal table As ROneCOne, _
    ByVal columns As Collection _
) As ROneCOne
    Dim group As ROneCOne
    Dim indexValue As ROneCOne
    Dim keyToken As String
    Dim row As ROneCOne

    Set indexValue = ROneCOne.DictionaryOf(vbString, ROneCOne.Exception)
    For Each row In table.InternalDataRows
        If row.RowState <> "Deleted" Then
            keyToken = RelationKeyToken(row, columns)
            If indexValue.ContainsKey(keyToken) Then
                Set group = indexValue.Item(keyToken)
            Else
                Set group = ROneCOne.ListOf(row)
                indexValue.Add keyToken, group
            End If
            group.Add row
        End If
    Next row
    Set BuildRelationIndex = indexValue
End Function

Private Function RelationKeyToken( _
    ByVal row As ROneCOne, _
    ByVal columns As Collection _
) As String
    Dim column As ROneCOne
    Dim token As String

    For Each column In columns
        token = DataIndexToken(row.DataFieldWrapped(column.ColumnName))
        RelationKeyToken = RelationKeyToken & CStr(Len(token)) & ":" & token
    Next column
End Function

Friend Sub ConfigureDataView(ByVal table As ROneCOne)
    mRole = ROLE_DATA_VIEW
    Set mSource = table
    mElementType = vbObject
    mElementObjectType = "ROneCOne"
End Sub

Friend Sub ConfigureDataName(ByVal value As String)
    mName = value
End Sub

Friend Sub ConfigureDataRow(ByVal table As ROneCOne)
    Dim column As ROneCOne

    mRole = ROLE_DATA_ROW
    Set mSource = table
    ArrReset mItems, mItemsCount
    ArrReset mOriginalItems, mOriginalItemsCount
    For Each column In table.InternalDataColumns
        ArrAppend mItems, mItemsCount, DefaultDataColumnValue(column)
    Next column
    mRowState = ROW_STATE_DETACHED
End Sub

Friend Sub ConfigureRowState(ByVal rowState As Long)
    mRowState = rowState
End Sub

Friend Sub InternalAppendRowCell(ByVal wrapped As ROneCOne)
    ArrAppend mItems, mItemsCount, wrapped
    If mOriginalItemsCount > 0 Then
        ArrAppend mOriginalItems, mOriginalItemsCount, wrapped
    End If
End Sub

Friend Sub TouchDataVersion()
    ' Field edits and row-state changes advance the data version only, so
    ' views and enumerations refresh while structural snapshot caches (whose
    ' membership did not change) survive. Membership changes advance mVersion.
    mDataVersion = mDataVersion + 1
End Sub

Friend Sub ConfigureUniqueColumn()
    mUnique = True
    mAllowDBNull = False
End Sub

Friend Sub ConfigurePrimaryKeyColumn(ByVal singleColumn As Boolean)
    mAllowDBNull = False
    If singleColumn Then mUnique = True
End Sub

Friend Sub ConfigureSinglePrimaryKey(ByVal column As ROneCOne)
    Set mPrimaryKey = New Collection
    column.ConfigurePrimaryKeyColumn True
    mPrimaryKey.Add column
    RebuildPrimaryKeyIndex
End Sub

Friend Property Get InternalAutoIncrement() As Boolean
Attribute InternalAutoIncrement.VB_MemberFlags = "40"
    InternalAutoIncrement = mAutoIncrement
End Property

Friend Property Get InternalDefaultValue() As ROneCOne
Attribute InternalDefaultValue.VB_MemberFlags = "40"
    Set InternalDefaultValue = mDefaultValue
End Property

Friend Property Get InternalAutoIncrementSeed() As LongLong
Attribute InternalAutoIncrementSeed.VB_MemberFlags = "40"
    InternalAutoIncrementSeed = mAutoIncrementSeed
End Property

Friend Property Get InternalAutoIncrementStep() As LongLong
Attribute InternalAutoIncrementStep.VB_MemberFlags = "40"
    InternalAutoIncrementStep = mAutoIncrementStep
End Property

Friend Property Get InternalAutoIncrementCurrent() As LongLong
Attribute InternalAutoIncrementCurrent.VB_MemberFlags = "40"
    InternalAutoIncrementCurrent = mAutoIncrementCurrent
End Property

Friend Property Get InternalPrimaryKey() As Collection
Attribute InternalPrimaryKey.VB_MemberFlags = "40"
    Set InternalPrimaryKey = mPrimaryKey
End Property

Friend Sub CopyDataColumnConfiguration(ByVal source As ROneCOne)
    mAllowDBNull = source.InternalAllowDBNull
    mUnique = source.InternalUnique
    mAutoIncrement = source.InternalAutoIncrement
    mAutoIncrementSeed = source.InternalAutoIncrementSeed
    mAutoIncrementStep = source.InternalAutoIncrementStep
    mAutoIncrementCurrent = source.InternalAutoIncrementCurrent
    Set mDefaultValue = source.InternalDefaultValue
End Sub

Friend Function NextAutoIncrementValue() As ROneCOne
    Dim currentValue As LongLong

    currentValue = mAutoIncrementCurrent
    mAutoIncrementCurrent = mAutoIncrementCurrent + mAutoIncrementStep
    Select Case mExpectedType
        Case vbByte
            Set NextAutoIncrementValue = CreateScalarResult(CByte(currentValue))
        Case vbInteger
            Set NextAutoIncrementValue = CreateScalarResult(CInt(currentValue))
        Case vbLong
            Set NextAutoIncrementValue = CreateScalarResult(CLng(currentValue))
        Case Else
            Set NextAutoIncrementValue = CreateScalarResult(CLngLng(currentValue))
    End Select
End Function

Friend Sub ConfigureDbConnection( _
    ByVal connectionString As String, _
    ByVal providerName As String _
)
    mRole = ROLE_DB_CONNECTION
    mConnectionString = connectionString
    mProviderName = UCase$(Trim$(providerName))
    If mProviderName <> "OLEDB" And mProviderName <> "ODBC" Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "DbConnection", _
            "Provider name must be OLEDB or ODBC."
    End If
End Sub

Friend Function InternalQueryDialect() As Long
    Dim probe As String

    probe = UCase$(mConnectionString & ";" & mProviderName)
    If InStr(1, probe, "ACE.OLEDB") > 0 Or _
        InStr(1, probe, "JET.OLEDB") > 0 Or _
        InStr(1, probe, "MICROSOFT ACCESS DRIVER") > 0 Or _
        InStr(1, probe, "MICROSOFT EXCEL DRIVER") > 0 Then
        InternalQueryDialect = QUERY_DIALECT_ACE
    ElseIf InStr(1, probe, "MSOLEDBSQL") > 0 Or _
        InStr(1, probe, "SQLOLEDB") > 0 Or _
        InStr(1, probe, "SQLNCLI") > 0 Or _
        InStr(1, probe, "SQL SERVER") > 0 Then
        InternalQueryDialect = QUERY_DIALECT_TSQL
    Else
        RaiseQueryError "Queryable supports the SQL Server and ACE " & _
            "dialects; this connection string names neither."
    End If
End Function

Friend Sub ConfigureSessionTask( _
    ByVal session As ROneCOne, _
    ByVal operationName As String, _
    ByVal writeSequence As Long, _
    ByVal cancellationToken As ROneCOne _
)
    ConfigureTask TASK_SESSION, Nothing, cancellationToken
    mMethodName = operationName
    mExpressionOption = writeSequence
    Set mSource = session
End Sub

Friend Sub ConfigureQueryable( _
    ByVal connection As ROneCOne, _
    ByVal tableName As String _
)
    mRole = ROLE_QUERYABLE
    Set mSqlQueryConnection = connection
    mSqlQueryTable = ValidateQueryIdentifier(tableName, "Queryable")
    mSqlQueryDialect = connection.InternalQueryDialect()
    Set mSqlQueryWheres = New Collection
    Set mSqlQueryOrderings = New Collection
    Set mSqlQueryColumns = New Collection
    mSqlQueryTake = -1
    mSqlQuerySkip = -1
End Sub

Friend Sub ConfigureQueryableState( _
    ByVal connection As ROneCOne, _
    ByVal tableName As String, _
    ByVal dialect As Long, _
    ByVal wheres As Collection, _
    ByVal orderings As Collection, _
    ByVal columns As Collection, _
    ByVal takeCount As Long, _
    ByVal skipCount As Long _
)
    mRole = ROLE_QUERYABLE
    Set mSqlQueryConnection = connection
    mSqlQueryTable = tableName
    mSqlQueryDialect = dialect
    Set mSqlQueryWheres = wheres
    Set mSqlQueryOrderings = orderings
    Set mSqlQueryColumns = columns
    mSqlQueryTake = takeCount
    mSqlQuerySkip = skipCount
End Sub

Friend Sub InternalQueryableAddWhere(ByVal body As ROneCOne)
    mSqlQueryWheres.Add body
End Sub

Friend Sub InternalQueryableSetTake(ByVal count As Long)
    mSqlQueryTake = count
End Sub

Friend Sub InternalQueryableSetSkip(ByVal count As Long)
    mSqlQuerySkip = count
End Sub

Friend Sub InternalQueryableSetColumns(ByVal columns As Collection)
    Set mSqlQueryColumns = columns
End Sub

Friend Sub InternalQueryableOrder( _
    ByVal columnName As String, _
    ByVal descending As Boolean, _
    ByVal replaceExisting As Boolean _
)
    If replaceExisting Then Set mSqlQueryOrderings = New Collection
    mSqlQueryOrderings.Add Array(columnName, descending)
End Sub

Friend Sub ConfigureDbCommand( _
    ByVal commandText As String, _
    ByVal connection As ROneCOne _
)
    mRole = ROLE_DB_COMMAND
    mCommandText = commandText
    Set mSource = connection
    ArrReset mItems, mItemsCount
    mCommandType = 1
End Sub

Friend Sub ConfigureDbParameter( _
    ByVal parameterName As String, _
    ByVal valueWrapped As ROneCOne, _
    ByVal direction As Long _
)
    mRole = ROLE_DB_PARAMETER
    mName = parameterName
    Set mQueryValue = valueWrapped
    mExpressionOption = direction
    mSourceVersion = "Current"
End Sub

Friend Sub ConfigureDbAdapter(ByVal selectCommand As ROneCOne)
    mRole = ROLE_DB_ADAPTER
    Set mSource = selectCommand
    Set mProviderErrors = New Collection
End Sub

Friend Sub ConfigureDbReader(ByVal providerReader As Object)
    mRole = ROLE_DB_READER
    Set mProviderObject = providerReader
    mQueryCount = 0
End Sub

Friend Sub ConfigureDbTransaction(ByVal connection As ROneCOne)
    mRole = ROLE_DB_TRANSACTION
    Set mSource = connection
    mTaskStatus = TASK_STATUS_RUNNING
    mTransactionActive = True
End Sub

Friend Sub ConfigureGrouping( _
    ByVal keyWrapped As ROneCOne, _
    ByVal firstElement As ROneCOne _
)
    ConfigureGenericCollection ROLE_COLLECTION, COLLECTION_GROUPING, _
        vbVariant, vbNullString, vbVariant
    Set mLeft = keyWrapped
    ArrAppend mItems, mItemsCount, firstElement
End Sub

Friend Sub ConfigureEmptyGrouping(ByVal keyWrapped As ROneCOne)
    ConfigureGenericCollection ROLE_COLLECTION, COLLECTION_GROUPING, _
        vbVariant, vbNullString, vbVariant
    Set mLeft = keyWrapped
End Sub

Friend Sub InternalAddLookupGroup( _
    ByVal keyWrapped As ROneCOne, _
    ByVal group As ROneCOne _
)
    Dim groupWrapped As ROneCOne

    PrepareHashInsert
    Set groupWrapped = CreateObjectResult(group)
    ArrAppend mKeys, mKeysCount, keyWrapped
    ArrAppend mItems, mItemsCount, groupWrapped
    AddHashIndexEntry keyWrapped, mKeysCount - 1, groupWrapped
End Sub

Friend Function InternalCollectionItem(ByVal index As Long) As ROneCOne
    Dim wrapped As ROneCOne

    Set wrapped = mItems(index + 1)
    If wrapped.InternalHasObjectValue Then
        Set InternalCollectionItem = wrapped.InternalObjectValue
    Else
        Set InternalCollectionItem = wrapped
    End If
End Function

Friend Sub InternalAddTaskChild(ByVal taskValue As ROneCOne)
    mTaskChildren.Add taskValue
End Sub

Friend Property Get InternalPairCount() As Long
Attribute InternalPairCount.VB_MemberFlags = "40"
    InternalPairCount = mKeysCount
End Property

Friend Function InternalPairKeyText(ByVal position As Long) As String
    InternalPairKeyText = CStr(mKeys(position).InternalScalarValue)
End Function

Friend Function InternalPairValue(ByVal position As Long) As Variant
    Dim wrapped As ROneCOne

    Set wrapped = mItems(position)
    If wrapped.InternalHasObjectValue Then
        Set InternalPairValue = wrapped.InternalObjectValue
    Else
        InternalPairValue = wrapped.InternalScalarValue
    End If
End Function

Friend Sub InternalAppendCollectionItem(ByVal wrapped As ROneCOne)
    ArrAppend mItems, mItemsCount, wrapped
End Sub

Friend Sub InternalAppendCollectionKey(ByVal wrapped As ROneCOne)
    ArrAppend mKeys, mKeysCount, wrapped
End Sub

Friend Sub InternalAppendCollectionPriority(ByVal wrapped As ROneCOne)
    ArrAppend mPriorities, mPrioritiesCount, wrapped
End Sub

Friend Sub InternalQueueValue(ByVal wrapped As ROneCOne)
    ArrAppend mItems, mItemsCount, wrapped
    mVersion = mVersion + 1
End Sub

Friend Sub InternalPushValue(ByVal wrapped As ROneCOne)
    If mItemsCount = 0 Then
        ArrAppend mItems, mItemsCount, wrapped
    Else
        ArrInsert mItems, mItemsCount, 1, wrapped
    End If
    mVersion = mVersion + 1
End Sub

Friend Sub ConfigureQuery( _
    ByVal source As ROneCOne, _
    ByVal queryOperator As Long, _
    ByVal queryDelegate As ROneCOne, _
    ByVal queryValue As ROneCOne, _
    ByVal queryCount As Long, _
    ByVal elementType As VbVarType, _
    ByVal elementObjectType As String, _
    Optional ByVal queryComparer As ROneCOne _
)
    mRole = ROLE_QUERY
    Set mSource = source
    mOperator = queryOperator
    Set mQueryDelegate = queryDelegate
    Set mQueryComparer = queryComparer
    Set mQueryValue = queryValue
    mQueryCount = queryCount
    mElementType = elementType
    mElementObjectType = elementObjectType
    mEnumerationVersion = -1
End Sub

Friend Sub ConfigureWhereBuilder( _
    ByVal source As ROneCOne, _
    ByVal memberPath As String _
)
    mRole = ROLE_WHERE_BUILDER
    Set mSource = source
    If Len(Trim$(memberPath)) = 0 Then
        mContextMemberPath = vbNullString
    Else
        mContextMemberPath = NormalizeMemberPath(memberPath, "Where")
    End If
End Sub

Friend Sub InternalAddWrapped(ByVal wrapped As ROneCOne)
    ValidateElementWrapped wrapped, "ToList"
    ArrAppend mItems, mItemsCount, wrapped
    mVersion = mVersion + 1
End Sub

Friend Property Get InternalSequenceVersion() As Long
Attribute InternalSequenceVersion.VB_MemberFlags = "40"
    If mRole = ROLE_LIST Then
        InternalSequenceVersion = mVersion
    ElseIf mRole = ROLE_QUERY Then
        InternalSequenceVersion = mSource.InternalSequenceVersion
    ElseIf IsGenericCollectionRole(mRole) Then
        If mCollectionKind = COLLECTION_READ_ONLY Then
            InternalSequenceVersion = mSource.InternalSequenceVersion
        Else
            InternalSequenceVersion = mVersion
        End If
    ElseIf mRole = ROLE_DATA_VIEW Then
        ' A view changes when its table changes or when its own filter or
        ' sort is replaced, so both counters feed the staleness key.
        InternalSequenceVersion = mSource.InternalSequenceVersion + mVersion
    ElseIf mRole = ROLE_DATA_TABLE Then
        InternalSequenceVersion = mVersion + mDataVersion
    End If
End Property

Friend Function MaterializeWrapped() As Collection
    Dim entryIndex As Long
    Dim sourceValues As Collection
    Dim values As Collection
    Dim wrapped As ROneCOne

    RequireSequence Me, "Materialize"
    If mRole = ROLE_LIST Then
        Set MaterializeWrapped = ArrSnapshot(mItems, mItemsCount)
        Exit Function
    End If
    If mRole = ROLE_DATA_VIEW Then
        Set MaterializeWrapped = MaterializeDataView
        Exit Function
    End If
    If IsGenericCollectionRole(mRole) Then
        If mCollectionKind = COLLECTION_READ_ONLY Then
            Set MaterializeWrapped = mSource.MaterializeWrapped
            Exit Function
        End If
        If IsDictionaryCollection Then
            Set values = New Collection
            For entryIndex = 1 To mItemsCount
                values.Add CreateDictionaryEntry( _
                    mKeys(entryIndex), mItems(entryIndex))
            Next entryIndex
            Set MaterializeWrapped = values
        Else
            Set MaterializeWrapped = ArrSnapshot(mItems, mItemsCount)
        End If
        Exit Function
    End If

    If IsOrderingOperator(mOperator) Then
        Set MaterializeWrapped = ApplyOrderChain
        Exit Function
    End If

    Set sourceValues = mSource.MaterializeWrapped
    Select Case mOperator
        Case QUERY_WHERE
            Set values = ApplyWhere(sourceValues)
        Case QUERY_WHERE_NOT
            Set values = ApplyWhere(sourceValues, True)
        Case QUERY_SELECT
            Set values = ApplySelect(sourceValues)
        Case QUERY_TAKE
            Set values = ApplyTake(sourceValues)
        Case QUERY_SKIP
            Set values = ApplySkip(sourceValues)
        Case QUERY_TAKE_WHILE
            Set values = ApplyTakeWhile(sourceValues)
        Case QUERY_SKIP_WHILE
            Set values = ApplySkipWhile(sourceValues)
        Case QUERY_TAKE_LAST
            Set values = ApplyTakeLast(sourceValues)
        Case QUERY_SKIP_LAST
            Set values = ApplySkipLast(sourceValues)
        Case QUERY_DISTINCT
            Set values = ApplyDistinct(sourceValues)
        Case QUERY_DISTINCT_BY
            Set values = ApplyDistinctBy(sourceValues)
        Case QUERY_APPEND
            Set values = CopyWrappedCollection(sourceValues)
            values.Add mQueryValue
        Case QUERY_PREPEND
            Set values = New Collection
            values.Add mQueryValue
            AddWrappedRange values, sourceValues
        Case QUERY_CONCAT
            Set values = CopyWrappedCollection(sourceValues)
            AddWrappedRange values, mQueryValue.MaterializeWrapped
        Case QUERY_UNION
            Set values = ApplyUnion(sourceValues)
        Case QUERY_INTERSECT
            Set values = ApplyIntersect(sourceValues)
        Case QUERY_EXCEPT
            Set values = ApplyExcept(sourceValues)
        Case QUERY_DEFAULT_IF_EMPTY
            Set values = ApplyDefaultIfEmpty(sourceValues)
        Case QUERY_CHUNK
            Set values = ApplyChunk(sourceValues)
        Case QUERY_SELECT_MANY
            Set values = ApplySelectMany(sourceValues)
        Case QUERY_CAST
            Set values = ApplyCast(sourceValues, False)
        Case QUERY_OF_TYPE
            Set values = ApplyCast(sourceValues, True)
        Case QUERY_UNION_BY, QUERY_INTERSECT_BY, QUERY_EXCEPT_BY
            Set values = ApplySetBy(sourceValues)
        Case QUERY_REVERSE
            Set values = ApplyReverse(sourceValues)
        Case Else
            RaiseContractError ERROR_INVALID_OPERATION, "Materialize", _
                "Unknown deferred query operator."
    End Select
    Set MaterializeWrapped = values
End Function

Private Function CreateCallableDelegate( _
    ByVal callable As Variant, _
    ByVal memberName As String, _
    ByVal isAction As Boolean _
) As ROneCOne
    Dim candidate As Object
    Dim created As ROneCOne
    Dim runtimeValue As ROneCOne
    Dim resolvedMember As String
    Dim procedureName As String

    If IsObject(callable) Then
        Set candidate = callable
        If candidate Is Nothing Then
            RaiseContractError ERROR_INVALID_ARGUMENT, "Func", _
                "The callable object cannot be Nothing."
        End If
        If TypeOf candidate Is ROneCOne Then
            Set runtimeValue = candidate
            resolvedMember = Trim$(memberName)
            If Len(resolvedMember) > 0 Then
                Set created = New ROneCOne
                created.ConfigureMethod candidate, resolvedMember, isAction
                Set CreateCallableDelegate = created
                Exit Function
            End If
            If IsDelegateRole(runtimeValue.InternalRole) Then
                If Not isAction And runtimeValue.InternalIsActionDelegate Then
                    RaiseContractError ERROR_INVALID_OPERATION, "Func", _
                        "An Action cannot be converted to a Func without a result."
                End If
                Set created = CloneDelegate(runtimeValue)
                created.ConfigureActionSignature isAction
                Set CreateCallableDelegate = created
                Exit Function
            End If
            RequireExpression runtimeValue, "Func"
            Set created = ROneCOne.Lambda(runtimeValue)
            created.ConfigureActionSignature isAction
            Set CreateCallableDelegate = created
            Exit Function
        End If

        resolvedMember = Trim$(memberName)
        If Len(resolvedMember) = 0 Then resolvedMember = "Run"
        Set created = New ROneCOne
        created.ConfigureMethod candidate, resolvedMember, isAction
        Set CreateCallableDelegate = created
        Exit Function
    End If

    If VarTypeOf(callable) <> vbString Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "Func", _
            "A callable must be an expression, object, or workbook procedure name."
    End If
    If Len(Trim$(memberName)) <> 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "Func", _
            "A workbook procedure does not accept a separate member name."
    End If
    procedureName = Trim$(CStr(callable))
    If Len(procedureName) = 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "Func", _
            "The workbook procedure name cannot be empty."
    End If
    Set created = New ROneCOne
    created.ConfigureProcedure procedureName, isAction
    Set CreateCallableDelegate = created
End Function

Private Function CloneDelegate(ByVal source As ROneCOne) As ROneCOne
    Dim clone As ROneCOne

    RequireDelegate source, "CloneDelegate"
    Set clone = New ROneCOne
    Select Case source.InternalRole
        Case ROLE_LAMBDA
            clone.ConfigureLambda source.InternalLeft, source.InternalParameters
        Case ROLE_METHOD
            clone.ConfigureMethod source.InternalTarget, source.InternalMethodName, _
                source.InternalIsActionDelegate
        Case ROLE_PROCEDURE
            clone.ConfigureProcedure source.InternalMethodName, _
                source.InternalIsActionDelegate
        Case ROLE_NATIVE
            clone.ConfigureNative source.InternalNativeAddress, _
                source.InternalIsActionDelegate
        Case ROLE_COMPOSITION
            clone.ConfigureComposition source.InternalLeft, source.InternalNext
        Case ROLE_MULTICAST
            clone.ConfigureMulticast source.InternalInvocationList
    End Select
    If source.InternalHasParameterSignature Then
        clone.ConfigureParameterSignature source.InternalParameters
    End If
    If source.InternalHasReturnSignature Then
        If source.InternalIsActionDelegate Then
            clone.ConfigureActionSignature True
        Else
            clone.ConfigureReturnSignature source.InternalReturnType, _
                source.InternalReturnObjectType
        End If
    End If
    Set CloneDelegate = clone
End Function

Private Function CloneTryOperation(ByVal source As ROneCOne) As ROneCOne
    Dim clone As ROneCOne
    Dim index As Long

    RequireTry source, "CloneTryOperation"
    Set clone = New ROneCOne
    clone.ConfigureTry source.InternalTryBody
    For index = 1 To source.InternalCatchActions.Count
        clone.ConfigureCatch source.InternalCatchNumbers.Item(index), _
            source.InternalCatchActions.Item(index)
    Next index
    If Not source.InternalFinallyAction Is Nothing Then
        clone.ConfigureFinally source.InternalFinallyAction
    End If
    Set CloneTryOperation = clone
End Function

Private Sub ValidateAction( _
    ByVal actionValue As ROneCOne, _
    ByVal memberName As String, _
    ByVal requiresNoArguments As Boolean _
)
    RequireDelegate actionValue, memberName
    If Not actionValue.IsAction Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            "This operation requires an Action delegate."
    End If
    If requiresNoArguments Then
        If actionValue.Arity <> 0 And actionValue.Arity <> -1 Then
            RaiseContractError ERROR_ARITY_MISMATCH, memberName, _
                "The Action must accept no arguments."
        End If
    End If
End Sub

Friend Function InternalWrapElement( _
    ByVal candidate As Variant, _
    ByVal memberName As String _
) As ROneCOne
    Set InternalWrapElement = WrapElement(candidate, memberName)
End Function

Private Sub ValidateCatchAction(ByVal actionValue As ROneCOne)
    Dim descriptor As ROneCOne

    ValidateAction actionValue, "Catch", False
    If Not actionValue.InternalHasParameterSignature Then
        RaiseContractError ERROR_INVALID_OPERATION, "Catch", _
            "A Catch Action requires a complete Takes signature."
    End If
    If actionValue.Arity <> 0 And actionValue.Arity <> 1 Then
        RaiseContractError ERROR_ARITY_MISMATCH, "Catch", _
            "A Catch Action must accept zero arguments or one exception."
    End If
    If actionValue.Arity = 1 And actionValue.InternalHasParameterSignature Then
        Set descriptor = actionValue.InternalParameters.Item(1)
        If descriptor.InternalExpectedType <> vbObject Or _
            StrComp(descriptor.InternalExpectedObjectType, "ROneCOne", _
                vbTextCompare) <> 0 Then
            RaiseContractError ERROR_TYPE_MISMATCH, "Catch", _
                "The Catch argument must use ROneCOne.Exception as its type."
        End If
    End If
End Sub

Private Sub ValidateEventHandler( _
    ByVal handler As ROneCOne, _
    ByVal memberName As String _
)
    Dim index As Long

    ValidateAction handler, memberName, False
    If Not handler.InternalHasParameterSignature Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            "An event handler requires a complete Takes signature."
    End If
    If handler.Arity <> mMethodArity Then
        RaiseContractError ERROR_ARITY_MISMATCH, memberName, _
            "The event handler arity does not match the event."
    End If
    For index = 1 To mParameters.Count
        If Not ParameterDescriptorsMatch( _
            mParameters.Item(index), handler.InternalParameters.Item(index) _
        ) Then
            RaiseContractError ERROR_TYPE_MISMATCH, memberName, _
                "The event handler signature does not match the event."
        End If
    Next index
End Sub

Private Sub CopySignatureFrom(ByVal source As ROneCOne)
    Set mParameters = source.InternalParameters
    mMethodArity = source.InternalMethodArity
    mIsActionDelegate = source.InternalIsActionDelegate
    mHasParameterSignature = source.InternalHasParameterSignature
    mHasReturnSignature = source.InternalHasReturnSignature
    mReturnType = source.InternalReturnType
    mReturnObjectType = source.InternalReturnObjectType
End Sub

Private Function CreateParameterDescriptor( _
    ByVal typeToken As Variant, _
    ByVal memberName As String _
) As ROneCOne
    Dim candidate As Object
    Dim descriptor As ROneCOne
    Dim objectType As String
    Dim valueType As VbVarType

    If IsObject(typeToken) Then
        Set candidate = typeToken
        If Not candidate Is Nothing Then
            If TypeOf candidate Is ROneCOne Then
                Set descriptor = candidate
                If descriptor.InternalRole = ROLE_PARAMETER And _
                    descriptor.InternalIsByRef Then
                    Set CreateParameterDescriptor = descriptor
                    Exit Function
                End If
            End If
        End If
    End If

    ResolveTypeDescriptor typeToken, valueType, objectType, memberName
    Set descriptor = New ROneCOne
    descriptor.ConfigureParameter valueType, objectType
    Set CreateParameterDescriptor = descriptor
End Function

Private Sub ValidateReferenceType( _
    ByVal valueType As VbVarType, _
    ByVal memberName As String _
)
    Select Case valueType
        Case vbByte, vbInteger, vbLong, vbLongLong
        Case vbSingle, vbDouble, vbCurrency, vbDate, vbBoolean
        Case Else
            RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
                "ByRef currently supports numeric, Date, and Boolean variables."
    End Select
End Sub

Private Function CreateReference( _
    ByVal variableAddress As LongPtr, _
    ByVal valueType As VbVarType _
) As ROneCOne
    Dim reference As ROneCOne

    Set reference = New ROneCOne
    reference.ConfigureReference variableAddress, valueType
    Set CreateReference = reference
End Function

Private Function BuildDynamicArguments(ByVal arguments As Variant) As Collection
    Dim candidate As Object
    Dim rawArgument As Variant
    Dim result As Collection
    Dim values As Variant

    Set result = New Collection
    If IsObject(arguments) Then
        Set candidate = arguments
        If candidate Is Nothing Then
            RaiseContractError ERROR_INVALID_ARGUMENT, "DynamicInvoke", _
                "The argument collection cannot be Nothing."
        End If
        If TypeOf candidate Is ROneCOne Then
            values = candidate.ToArray
            For Each rawArgument In values
                result.Add WrapInvocationArgument(rawArgument)
            Next rawArgument
            Set BuildDynamicArguments = result
            Exit Function
        End If
        If TypeName(candidate) = "Collection" Then
            For Each rawArgument In candidate
                result.Add WrapInvocationArgument(rawArgument)
            Next rawArgument
            Set BuildDynamicArguments = result
            Exit Function
        End If
    ElseIf IsArrayValue(arguments) Then
        For Each rawArgument In arguments
            result.Add WrapInvocationArgument(rawArgument)
        Next rawArgument
        Set BuildDynamicArguments = result
        Exit Function
    ElseIf IsEmpty(arguments) Then
        Set BuildDynamicArguments = result
        Exit Function
    End If

    RaiseContractError ERROR_INVALID_ARGUMENT, "DynamicInvoke", _
        "Arguments must be a VBA array, Collection, or ROneCOne sequence."
End Function

Private Function WrapInvocationArgument(ByVal rawArgument As Variant) As ROneCOne
    Dim argumentObject As Object
    Dim wrapped As ROneCOne

    If IsObject(rawArgument) Then
        Set argumentObject = rawArgument
        If Not argumentObject Is Nothing Then
            If TypeOf argumentObject Is ROneCOne Then
                Set wrapped = argumentObject
                If wrapped.InternalRole = ROLE_REFERENCE Then
                    Set WrapInvocationArgument = wrapped
                    Exit Function
                End If
            End If
        End If
    End If

    Set wrapped = New ROneCOne
    If IsObject(rawArgument) Then
        wrapped.ConfigureObjectValue argumentObject
    Else
        wrapped.ConfigureScalarValue rawArgument
    End If
    Set WrapInvocationArgument = wrapped
End Function

Private Sub AssignPublicResult( _
    ByRef destination As Variant, _
    ByVal result As ROneCOne _
)
    If result.InternalHasObjectValue Then
        Set destination = result.InternalObjectValue
    Else
        destination = result.InternalScalarValue
    End If
End Sub

Private Sub AssignDefaultPublicResult(ByRef destination As Variant)
    If mElementType = vbObject Then
        Set destination = Nothing
        Exit Sub
    End If
    Select Case mElementType
        Case vbString
            destination = vbNullString
        Case vbBoolean
            destination = False
        Case vbDate
            destination = CDate(0)
        Case vbByte
            destination = CByte(0)
        Case vbInteger
            destination = CInt(0)
        Case vbLong
            destination = CLng(0)
        Case vbLongLong
            destination = CLngLng(0)
        Case vbSingle
            destination = CSng(0)
        Case vbDouble
            destination = CDbl(0)
        Case vbCurrency
            destination = CCur(0)
        Case Else
            destination = Empty
    End Select
End Sub

Private Function IsDelegateRole(ByVal roleValue As Long) As Boolean
    Select Case roleValue
        Case ROLE_LAMBDA, ROLE_METHOD, ROLE_COMPOSITION
            IsDelegateRole = True
        Case ROLE_PROCEDURE, ROLE_MULTICAST, ROLE_NATIVE
            IsDelegateRole = True
    End Select
End Function

Private Sub AppendInvocationList( _
    ByVal destination As Collection, _
    ByVal delegateValue As ROneCOne _
)
    Dim item As ROneCOne

    If delegateValue.InternalRole = ROLE_MULTICAST Then
        For Each item In delegateValue.InternalInvocationList
            destination.Add item
        Next item
    Else
        destination.Add delegateValue
    End If
End Sub

Private Sub ValidateInvocationCompatibility(ByVal invocationList As Collection)
    Dim firstDelegate As ROneCOne
    Dim index As Long

    Set firstDelegate = invocationList.Item(1)
    For index = 2 To invocationList.Count
        If Not DelegateSignaturesMatch(firstDelegate, invocationList.Item(index)) Then
            RaiseContractError ERROR_TYPE_MISMATCH, "Combine", _
                "Multicast delegates must have matching signatures."
        End If
    Next index
End Sub

Private Function DelegateSignaturesMatch( _
    ByVal leftDelegate As ROneCOne, _
    ByVal rightDelegate As ROneCOne _
) As Boolean
    Dim index As Long

    If leftDelegate.IsAction <> rightDelegate.IsAction Then Exit Function
    If leftDelegate.Arity <> rightDelegate.Arity Then Exit Function
    If leftDelegate.InternalHasParameterSignature <> _
        rightDelegate.InternalHasParameterSignature Then Exit Function
    If leftDelegate.InternalHasReturnSignature <> _
        rightDelegate.InternalHasReturnSignature Then Exit Function
    If leftDelegate.InternalHasReturnSignature Then
        If leftDelegate.InternalReturnType <> _
            rightDelegate.InternalReturnType Then Exit Function
        If StrComp( _
            leftDelegate.InternalReturnObjectType, _
            rightDelegate.InternalReturnObjectType, _
            vbTextCompare _
        ) <> 0 Then Exit Function
    End If
    If leftDelegate.InternalHasParameterSignature Then
        For index = 1 To leftDelegate.InternalParameters.Count
            If Not ParameterDescriptorsMatch( _
                leftDelegate.InternalParameters.Item(index), _
                rightDelegate.InternalParameters.Item(index) _
            ) Then Exit Function
        Next index
    End If
    DelegateSignaturesMatch = True
End Function

Private Function ParameterDescriptorsMatch( _
    ByVal leftParameter As ROneCOne, _
    ByVal rightParameter As ROneCOne _
) As Boolean
    If leftParameter.InternalExpectedType <> _
        rightParameter.InternalExpectedType Then Exit Function
    If leftParameter.InternalIsByRef <> _
        rightParameter.InternalIsByRef Then Exit Function
    ParameterDescriptorsMatch = (StrComp( _
        leftParameter.InternalExpectedObjectType, _
        rightParameter.InternalExpectedObjectType, _
        vbTextCompare) = 0)
End Function

Private Function RemoveListItem(ByVal value As Variant) As Boolean
    Dim index As Long

    RequireList Me, "Remove"
    index = IndexOf(value)
    If index >= 0 Then
        ArrRemoveAt mItems, mItemsCount, index + 1
        mVersion = mVersion + 1
        RemoveListItem = True
    End If
End Function

Private Function RemoveDelegateValue( _
    ByVal source As ROneCOne, _
    ByVal value As ROneCOne _
) As ROneCOne
    Dim candidate As ROneCOne
    Dim index As Long
    Dim matchIndex As Long
    Dim offset As Long
    Dim remaining As Collection
    Dim sourceList As Collection
    Dim valueList As Collection

    RequireDelegate source, "Remove"
    RequireDelegate value, "Remove"
    Set sourceList = New Collection
    Set valueList = New Collection
    AppendInvocationList sourceList, source
    AppendInvocationList valueList, value
    matchIndex = 0
    For index = sourceList.Count - valueList.Count + 1 To 1 Step -1
        For offset = 1 To valueList.Count
            If Not DelegateIdentityEquals( _
                sourceList.Item(index + offset - 1), _
                valueList.Item(offset) _
            ) Then Exit For
        Next offset
        If offset > valueList.Count Then
            matchIndex = index
            Exit For
        End If
    Next index
    If matchIndex = 0 Then
        Set RemoveDelegateValue = source
        Exit Function
    End If

    Set remaining = New Collection
    For index = 1 To sourceList.Count
        If index < matchIndex Or index >= matchIndex + valueList.Count Then
            remaining.Add sourceList.Item(index)
        End If
    Next index
    If remaining.Count = 0 Then Exit Function
    If remaining.Count = 1 Then
        Set RemoveDelegateValue = remaining.Item(1)
        Exit Function
    End If
    Set candidate = New ROneCOne
    candidate.ConfigureMulticast remaining
    Set RemoveDelegateValue = candidate
End Function

Private Function DelegateIdentityEquals( _
    ByVal leftDelegate As ROneCOne, _
    ByVal rightDelegate As ROneCOne _
) As Boolean
    If leftDelegate Is rightDelegate Then
        DelegateIdentityEquals = True
        Exit Function
    End If
    If leftDelegate.InternalRole <> rightDelegate.InternalRole Then Exit Function
    If Not DelegateSignaturesMatch(leftDelegate, rightDelegate) Then Exit Function

    Select Case leftDelegate.InternalRole
        Case ROLE_METHOD
            DelegateIdentityEquals = _
                (leftDelegate.InternalTarget Is rightDelegate.InternalTarget) And _
                (StrComp(leftDelegate.InternalMethodName, _
                    rightDelegate.InternalMethodName, _
                    vbTextCompare) = 0)
        Case ROLE_PROCEDURE
            DelegateIdentityEquals = (StrComp( _
                leftDelegate.InternalMethodName, _
                rightDelegate.InternalMethodName, _
                vbTextCompare) = 0)
        Case ROLE_NATIVE
            DelegateIdentityEquals = _
                (leftDelegate.InternalNativeAddress = _
                    rightDelegate.InternalNativeAddress)
        Case ROLE_COMPOSITION
            DelegateIdentityEquals = _
                (leftDelegate.InternalLeft Is rightDelegate.InternalLeft) And _
                (leftDelegate.InternalNext Is rightDelegate.InternalNext)
    End Select
End Function

Private Function BuildSignature(ByVal delegateValue As ROneCOne) As String
    Dim descriptor As ROneCOne
    Dim parts As String

    If delegateValue.InternalHasParameterSignature Then
        For Each descriptor In delegateValue.InternalParameters
            If Len(parts) > 0 Then parts = parts & ", "
            If descriptor.InternalIsByRef Then parts = parts & "ref "
            If Len(descriptor.InternalExpectedObjectType) > 0 Then
                parts = parts & descriptor.InternalExpectedObjectType
            Else
                parts = parts & TypeNameForVarType(descriptor.InternalExpectedType)
            End If
        Next descriptor
    Else
        parts = "dynamic"
    End If

    If delegateValue.InternalIsActionDelegate Then
        If parts = vbNullString Then
            BuildSignature = "Action"
        Else
            BuildSignature = "Action<" & parts & ">"
        End If
        Exit Function
    End If
    If Len(parts) > 0 Then parts = parts & ", "
    If delegateValue.InternalHasReturnSignature Then
        If Len(delegateValue.InternalReturnObjectType) > 0 Then
            parts = parts & delegateValue.InternalReturnObjectType
        Else
            parts = parts & TypeNameForVarType(delegateValue.InternalReturnType)
        End If
    Else
        parts = parts & "dynamic"
    End If
    BuildSignature = "Func<" & parts & ">"
End Function

Private Sub RequireDbConnection( _
    ByVal connection As ROneCOne, _
    ByVal memberName As String _
)
    If connection Is Nothing Or connection.InternalRole <> ROLE_DB_CONNECTION Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            "This operation requires a DbConnection."
    End If
End Sub

Private Sub RequireDbCommand( _
    ByVal command As ROneCOne, _
    ByVal memberName As String _
)
    If command Is Nothing Or command.InternalRole <> ROLE_DB_COMMAND Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            "This operation requires a DbCommand."
    End If
End Sub

Private Sub RequireDbAdapter( _
    ByVal adapter As ROneCOne, _
    ByVal memberName As String _
)
    If adapter Is Nothing Or adapter.InternalRole <> ROLE_DB_ADAPTER Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            "This operation requires a DbDataAdapter."
    End If
End Sub

' ---------------------------------------------------------------------------
' JSON engine. The reader/writer mechanics follow ModernJsonInVBA (used with
' its author's permission): byte-snapshot character scanning, bounded-chunk
' string reading with an escape-free fast path, inline accumulation of short
' integers, culture-invariant number writing, and RFC 8259 strictness with
' position-carrying errors. The model is runtime-native: a JSON object is an
' ordered String-to-Variant dictionary, a JSON array is a Variant list.
' ---------------------------------------------------------------------------

Private Sub RequireJsonFactory(ByVal memberName As String)
    If mRole <> ROLE_JSON_FACTORY Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            memberName & " is called through ROneCOne.Json."
    End If
End Sub

Private Sub RaiseJsonError(ByVal message As String)
    Err.Raise ERROR_JSON_PARSE, "ROneCOne.JsonException", message
End Sub

Private Sub AssignVariant(ByRef target As Variant, ByVal value As Variant)
    If IsObject(value) Then
        Set target = value
    Else
        target = value
    End If
End Sub

Private Sub JsonReaderInit(ByRef reader As JsonReader, ByVal jsonText As String)
    reader.text = jsonText
    reader.pos = 1
    reader.textLen = Len(jsonText)
    If reader.textLen > 0 Then reader.bytes = jsonText
End Sub

Private Function JsonCodeAt(ByRef reader As JsonReader) As Long
    Dim offset As Long

    If reader.pos > reader.textLen Then
        JsonCodeAt = -1
    Else
        offset = (reader.pos - 1) * 2
        JsonCodeAt = reader.bytes(offset) + reader.bytes(offset + 1) * 256&
    End If
End Function

Private Sub JsonSkipWhitespace(ByRef reader As JsonReader)
    Dim offset As Long

    Do While reader.pos <= reader.textLen
        offset = (reader.pos - 1) * 2
        Select Case reader.bytes(offset) + reader.bytes(offset + 1) * 256&
            Case 32, 9, 13, 10
                reader.pos = reader.pos + 1
            Case Else
                Exit Do
        End Select
    Loop
End Sub

Private Sub JsonExpectChar( _
    ByRef reader As JsonReader, _
    ByVal expectedCode As Long, _
    ByVal expectedChar As String _
)
    JsonSkipWhitespace reader
    If JsonCodeAt(reader) = expectedCode Then
        reader.pos = reader.pos + 1
        Exit Sub
    End If
    RaiseJsonError "Expected '" & expectedChar & "' at position " & _
        CStr(reader.pos)
End Sub

Private Sub JsonExpectLiteral(ByRef reader As JsonReader, ByVal literal As String)
    If Mid$(reader.text, reader.pos, Len(literal)) <> literal Then
        RaiseJsonError "Expected literal '" & literal & _
            "' near position " & CStr(reader.pos)
    End If
    reader.pos = reader.pos + Len(literal)
End Sub

Private Function JsonReadValue(ByRef reader As JsonReader) As Variant
    JsonSkipWhitespace reader
    Select Case JsonCodeAt(reader)
        Case 34
            JsonReadValue = JsonReadString(reader)
        Case 116
            JsonExpectLiteral reader, "true"
            JsonReadValue = True
        Case 102
            JsonExpectLiteral reader, "false"
            JsonReadValue = False
        Case 110
            JsonExpectLiteral reader, "null"
            JsonReadValue = Null
        Case 45, 48 To 57
            JsonReadValue = JsonReadNumber(reader)
        Case 91
            Set JsonReadValue = JsonReadArray(reader)
        Case 123
            Set JsonReadValue = JsonReadObject(reader)
        Case Else
            RaiseJsonError "Unexpected token at position " & CStr(reader.pos)
    End Select
End Function

Private Function JsonReadArray(ByRef reader As JsonReader) As ROneCOne
    Dim element As Variant
    Dim result As ROneCOne

    reader.pos = reader.pos + 1
    Set result = ROneCOne.ListOf(vbVariant)
    JsonSkipWhitespace reader
    If JsonCodeAt(reader) = 93 Then
        reader.pos = reader.pos + 1
        Set JsonReadArray = result
        Exit Function
    End If
    Do
        AssignVariant element, JsonReadValue(reader)
        result.Add element
        JsonSkipWhitespace reader
        Select Case JsonCodeAt(reader)
            Case 44
                reader.pos = reader.pos + 1
            Case 93
                reader.pos = reader.pos + 1
                Exit Do
            Case Else
                RaiseJsonError "Expected ',' or ']' at position " & _
                    CStr(reader.pos)
        End Select
    Loop
    Set JsonReadArray = result
End Function

' Advances the reader past one complete value without building anything.
' This is what makes a partial read cheap: the members a caller did not ask
' for are stepped over instead of materialized.
Private Sub JsonSkipValue(ByRef reader As JsonReader)
    JsonSkipWhitespace reader
    Select Case JsonCodeAt(reader)
        Case 34
            JsonSkipString reader
        Case 116
            JsonExpectLiteral reader, "true"
        Case 102
            JsonExpectLiteral reader, "false"
        Case 110
            JsonExpectLiteral reader, "null"
        Case 45, 48 To 57
            JsonSkipNumber reader
        Case 91
            JsonSkipArray reader
        Case 123
            JsonSkipObject reader
        Case Else
            RaiseJsonError "Unexpected token at position " & CStr(reader.pos)
    End Select
End Sub

Private Sub JsonSkipString(ByRef reader As JsonReader)
    Dim code As Long

    reader.pos = reader.pos + 1
    Do
        code = JsonCodeAt(reader)
        If code = -1 Then RaiseJsonError "Unterminated string"
        reader.pos = reader.pos + 1
        If code = 92 Then
            If JsonCodeAt(reader) = -1 Then RaiseJsonError "Unterminated string"
            reader.pos = reader.pos + 1
        ElseIf code = 34 Then
            Exit Do
        End If
    Loop
End Sub

Private Sub JsonSkipNumber(ByRef reader As JsonReader)
    Dim code As Long

    Do
        code = JsonCodeAt(reader)
        Select Case code
            Case 43, 45, 46, 69, 101, 48 To 57
                reader.pos = reader.pos + 1
            Case Else
                Exit Do
        End Select
    Loop
End Sub

Private Sub JsonSkipArray(ByRef reader As JsonReader)
    reader.pos = reader.pos + 1
    JsonSkipWhitespace reader
    If JsonCodeAt(reader) = 93 Then
        reader.pos = reader.pos + 1
        Exit Sub
    End If
    Do
        JsonSkipValue reader
        JsonSkipWhitespace reader
        Select Case JsonCodeAt(reader)
            Case 44
                reader.pos = reader.pos + 1
            Case 93
                reader.pos = reader.pos + 1
                Exit Do
            Case Else
                RaiseJsonError "Expected ',' or ']' at position " & _
                    CStr(reader.pos)
        End Select
    Loop
End Sub

Private Sub JsonSkipObject(ByRef reader As JsonReader)
    reader.pos = reader.pos + 1
    JsonSkipWhitespace reader
    If JsonCodeAt(reader) = 125 Then
        reader.pos = reader.pos + 1
        Exit Sub
    End If
    Do
        JsonSkipWhitespace reader
        If JsonCodeAt(reader) <> 34 Then
            RaiseJsonError "Expected a quoted member name at position " & _
                CStr(reader.pos)
        End If
        JsonSkipString reader
        JsonExpectChar reader, 58, ":"
        JsonSkipValue reader
        JsonSkipWhitespace reader
        Select Case JsonCodeAt(reader)
            Case 44
                reader.pos = reader.pos + 1
            Case 125
                reader.pos = reader.pos + 1
                Exit Do
            Case Else
                RaiseJsonError "Expected ',' or '}' at position " & _
                    CStr(reader.pos)
        End Select
    Loop
End Sub

Private Function JsonReadObject(ByRef reader As JsonReader) As ROneCOne
    Dim key As String
    Dim result As ROneCOne
    Dim value As Variant

    reader.pos = reader.pos + 1
    Set result = ROneCOne.OrderedDictionaryOf(vbString, vbVariant)
    JsonSkipWhitespace reader
    If JsonCodeAt(reader) = 125 Then
        reader.pos = reader.pos + 1
        Set JsonReadObject = result
        Exit Function
    End If
    Do
        JsonSkipWhitespace reader
        If JsonCodeAt(reader) <> 34 Then
            RaiseJsonError "Expected a quoted member name at position " & _
                CStr(reader.pos)
        End If
        key = JsonReadString(reader)
        JsonExpectChar reader, 58, ":"
        AssignVariant value, JsonReadValue(reader)
        ' A duplicate member name keeps the last value, matching how
        ' System.Text.Json deserializes into a dictionary.
        result.Item(key) = value
        JsonSkipWhitespace reader
        Select Case JsonCodeAt(reader)
            Case 44
                reader.pos = reader.pos + 1
            Case 125
                reader.pos = reader.pos + 1
                Exit Do
            Case Else
                RaiseJsonError "Expected ',' or '}' at position " & _
                    CStr(reader.pos)
        End Select
    Loop
    Set JsonReadObject = result
End Function

Private Function JsonReadNumber(ByRef reader As JsonReader) As Variant
    Dim accumulated As Long
    Dim code As Long
    Dim digitCount As Long
    Dim integral As Boolean
    Dim negative As Boolean
    Dim numberText As String
    Dim startPos As Long

    startPos = reader.pos
    code = JsonCodeAt(reader)
    If code = 45 Then
        negative = True
        reader.pos = reader.pos + 1
        code = JsonCodeAt(reader)
    End If
    If code = 48 Then
        digitCount = 1
        reader.pos = reader.pos + 1
        code = JsonCodeAt(reader)
        If code >= 48 And code <= 57 Then
            RaiseJsonError "Leading zeros are not allowed at position " & _
                CStr(reader.pos)
        End If
    ElseIf code >= 49 And code <= 57 Then
        Do
            If digitCount < 9 Then accumulated = accumulated * 10 + (code - 48)
            digitCount = digitCount + 1
            reader.pos = reader.pos + 1
            code = JsonCodeAt(reader)
        Loop While code >= 48 And code <= 57
    Else
        RaiseJsonError "Invalid number at position " & CStr(reader.pos)
    End If
    integral = True
    If code = 46 Then
        integral = False
        reader.pos = reader.pos + 1
        code = JsonCodeAt(reader)
        If Not (code >= 48 And code <= 57) Then
            RaiseJsonError "Invalid fraction at position " & CStr(reader.pos)
        End If
        Do
            reader.pos = reader.pos + 1
            code = JsonCodeAt(reader)
        Loop While code >= 48 And code <= 57
    End If
    If code = 101 Or code = 69 Then
        integral = False
        reader.pos = reader.pos + 1
        code = JsonCodeAt(reader)
        If code = 43 Or code = 45 Then
            reader.pos = reader.pos + 1
            code = JsonCodeAt(reader)
        End If
        If Not (code >= 48 And code <= 57) Then
            RaiseJsonError "Invalid exponent at position " & CStr(reader.pos)
        End If
        Do
            reader.pos = reader.pos + 1
            code = JsonCodeAt(reader)
        Loop While code >= 48 And code <= 57
    End If
    If integral And digitCount <= 9 Then
        ' Short integers were accumulated during the scan: no substring,
        ' no conversion call.
        If negative Then
            JsonReadNumber = -accumulated
        Else
            JsonReadNumber = accumulated
        End If
        Exit Function
    End If
    numberText = Mid$(reader.text, startPos, reader.pos - startPos)
    If integral Then
        On Error Resume Next
        JsonReadNumber = CLng(numberText)
        If Err.Number = 0 Then
            On Error GoTo 0
            Exit Function
        End If
        Err.Clear
        JsonReadNumber = CLngLng(numberText)
        If Err.Number = 0 Then
            On Error GoTo 0
            Exit Function
        End If
        Err.Clear
        On Error GoTo 0
    End If
    ' Val always parses a dot decimal separator, so a comma-decimal locale
    ' cannot misread the already-validated invariant number text.
    JsonReadNumber = Val(numberText)
End Function

Private Function JsonReadString(ByRef reader As JsonReader) As String
    Dim builder As JsonTextBuilder
    Dim chunk As String
    Dim chunkStart As Long
    Dim cursor As Long
    Dim escapeAt As Long
    Dim escapeChar As String
    Dim quoteAt As Long

    ' The candidate closing quote is found with InStr and escapes are probed
    ' only within that bounded chunk, so an escape-free string (the common
    ' case) is returned with a single Mid$ and never touches the builder.
    reader.pos = reader.pos + 1
    quoteAt = InStr(reader.pos, reader.text, """", vbBinaryCompare)
    If quoteAt = 0 Then RaiseJsonError "Unterminated string"
    chunkStart = reader.pos
    chunk = Mid$(reader.text, chunkStart, quoteAt - chunkStart)
    escapeAt = InStr(1, chunk, "\", vbBinaryCompare)
    If escapeAt = 0 Then
        JsonValidateStringRange reader, chunkStart, quoteAt - 1
        reader.pos = quoteAt + 1
        JsonReadString = chunk
        Exit Function
    End If
    JsonBuilderInit builder, Len(chunk) + 16
    cursor = 1
    Do
        If escapeAt > cursor Then
            JsonValidateStringRange reader, chunkStart + cursor - 1, _
                chunkStart + escapeAt - 2
            JsonBuilderAppend builder, Mid$(chunk, cursor, escapeAt - cursor)
        End If
        If escapeAt = Len(chunk) Then
            ' The escape is the quote candidate itself: an escaped quote.
            JsonBuilderAppend builder, """"
            chunkStart = quoteAt + 1
            quoteAt = InStr(chunkStart, reader.text, """", vbBinaryCompare)
            If quoteAt = 0 Then RaiseJsonError "Unterminated string"
            chunk = Mid$(reader.text, chunkStart, quoteAt - chunkStart)
            cursor = 1
        Else
            escapeChar = Mid$(chunk, escapeAt + 1, 1)
            Select Case escapeChar
                Case "\"
                    JsonBuilderAppend builder, "\"
                    cursor = escapeAt + 2
                Case "/"
                    JsonBuilderAppend builder, "/"
                    cursor = escapeAt + 2
                Case """"
                    JsonBuilderAppend builder, """"
                    cursor = escapeAt + 2
                Case "b"
                    JsonBuilderAppend builder, Chr$(8)
                    cursor = escapeAt + 2
                Case "f"
                    JsonBuilderAppend builder, Chr$(12)
                    cursor = escapeAt + 2
                Case "n"
                    JsonBuilderAppend builder, vbLf
                    cursor = escapeAt + 2
                Case "r"
                    JsonBuilderAppend builder, vbCr
                    cursor = escapeAt + 2
                Case "t"
                    JsonBuilderAppend builder, vbTab
                    cursor = escapeAt + 2
                Case "u"
                    reader.pos = chunkStart + escapeAt + 1
                    JsonBuilderAppend builder, JsonReadUnicodeEscape(reader)
                    cursor = reader.pos - chunkStart + 1
                Case Else
                    RaiseJsonError "Invalid escape '\" & escapeChar & _
                        "' at position " & CStr(chunkStart + escapeAt)
            End Select
        End If
        escapeAt = InStr(cursor, chunk, "\", vbBinaryCompare)
        If escapeAt = 0 Then
            If cursor <= Len(chunk) Then
                JsonValidateStringRange reader, chunkStart + cursor - 1, _
                    quoteAt - 1
                JsonBuilderAppend builder, Mid$(chunk, cursor)
            End If
            reader.pos = quoteAt + 1
            JsonReadString = JsonBuilderText(builder)
            Exit Function
        End If
    Loop
End Function

Private Sub JsonValidateStringRange( _
    ByRef reader As JsonReader, _
    ByVal fromPos As Long, _
    ByVal toPos As Long _
)
    Dim index As Long

    For index = fromPos To toPos
        If reader.bytes((index - 1) * 2) < 32 Then
            If reader.bytes((index - 1) * 2 + 1) = 0 Then
                RaiseJsonError "Unescaped control character in string at " & _
                    "position " & CStr(index)
            End If
        End If
    Next index
End Sub

Private Function JsonReadUnicodeEscape(ByRef reader As JsonReader) As String
    Dim high As Long
    Dim low As Long

    high = JsonReadHex4(reader)
    If high >= &HD800& And high <= &HDBFF& Then
        If Mid$(reader.text, reader.pos, 2) <> "\u" Then
            RaiseJsonError "Invalid surrogate pair"
        End If
        reader.pos = reader.pos + 2
        low = JsonReadHex4(reader)
        If low < &HDC00& Or low > &HDFFF& Then
            RaiseJsonError "Invalid surrogate pair"
        End If
        JsonReadUnicodeEscape = ChrW$(high) & ChrW$(low)
        Exit Function
    End If
    If high >= &HDC00& And high <= &HDFFF& Then
        RaiseJsonError "Invalid surrogate pair"
    End If
    JsonReadUnicodeEscape = ChrW$(high)
End Function

Private Function JsonReadHex4(ByRef reader As JsonReader) As Long
    Dim code As Long
    Dim index As Long
    Dim value As Long

    If reader.pos + 3 > reader.textLen Then
        RaiseJsonError "Incomplete \uXXXX escape"
    End If
    For index = 0 To 3
        code = reader.bytes((reader.pos + index - 1) * 2) + _
            reader.bytes((reader.pos + index - 1) * 2 + 1) * 256&
        Select Case code
            Case 48 To 57
                value = value * 16 + (code - 48)
            Case 65 To 70
                value = value * 16 + (code - 55)
            Case 97 To 102
                value = value * 16 + (code - 87)
            Case Else
                RaiseJsonError "Invalid \uXXXX escape"
        End Select
    Next index
    reader.pos = reader.pos + 4
    JsonReadHex4 = value
End Function

Private Sub JsonBuilderInit( _
    ByRef builder As JsonTextBuilder, _
    Optional ByVal initialCapacity As Long = 256 _
)
    If initialCapacity < 16 Then initialCapacity = 16
    builder.buffer = Space$(initialCapacity)
    builder.used = 0
    builder.capacity = initialCapacity
End Sub

Private Sub JsonBuilderAppend(ByRef builder As JsonTextBuilder, ByVal text As String)
    Dim addLength As Long
    Dim newCapacity As Long

    addLength = Len(text)
    If addLength = 0 Then Exit Sub
    If builder.used + addLength > builder.capacity Then
        newCapacity = builder.capacity
        Do While builder.used + addLength > newCapacity
            newCapacity = newCapacity * 2
        Loop
        builder.buffer = builder.buffer & Space$(newCapacity - builder.capacity)
        builder.capacity = newCapacity
    End If
    Mid$(builder.buffer, builder.used + 1, addLength) = text
    builder.used = builder.used + addLength
End Sub

Private Function JsonBuilderText(ByRef builder As JsonTextBuilder) As String
    If builder.used > 0 Then
        JsonBuilderText = Left$(builder.buffer, builder.used)
    End If
End Function

Private Sub JsonWriteValue(ByRef builder As JsonTextBuilder, ByVal value As Variant)
    Select Case VarTypeOf(value)
        Case vbString
            JsonBuilderAppend builder, """"
            JsonWriteEscaped builder, CStr(value)
            JsonBuilderAppend builder, """"
        Case vbBoolean
            If value Then
                JsonBuilderAppend builder, "true"
            Else
                JsonBuilderAppend builder, "false"
            End If
        Case vbByte, vbInteger, vbLong, vbLongLong
            JsonBuilderAppend builder, CStr(value)
        Case vbSingle, vbDouble, vbCurrency, vbDecimal
            JsonBuilderAppend builder, JsonNumberText(CDbl(value))
        Case vbDate
            ' Dates travel as ISO 8601 text, matching System.Text.Json.
            JsonBuilderAppend builder, """" & _
                VBA.Format$(value, "yyyy-mm-dd\Thh:nn:ss") & """"
        Case vbNull, vbEmpty
            JsonBuilderAppend builder, "null"
        Case vbObject, vbDataObject
            If value Is Nothing Then
                JsonBuilderAppend builder, "null"
            ElseIf TypeOf value Is ROneCOne Then
                JsonWriteRuntimeValue builder, value
            ElseIf TypeName(value) = "Collection" Then
                JsonWriteVbaCollection builder, value
            Else
                RaiseJsonError "Objects of type '" & TypeName(value) & _
                    "' are not serializable; map them with " & _
                    "DataTableFromObjects first."
            End If
        Case Else
            If IsArrayValue(value) Then
                JsonWriteVariantArray builder, value
            Else
                JsonBuilderAppend builder, """"
                JsonWriteEscaped builder, CStr(value)
                JsonBuilderAppend builder, """"
            End If
    End Select
End Sub

Private Sub JsonWriteRuntimeValue( _
    ByRef builder As JsonTextBuilder, _
    ByVal value As ROneCOne _
)
    Dim pairIndex As Long
    Dim row As ROneCOne
    Dim wrapped As ROneCOne

    Select Case value.InternalRole
        Case ROLE_DATA_TABLE, ROLE_DATA_VIEW
            JsonBuilderAppend builder, "["
            Dim firstRow As Boolean
            firstRow = True
            For Each row In value.InternalLiveDataRows
                If Not firstRow Then JsonBuilderAppend builder, ","
                firstRow = False
                JsonWriteDataRow builder, row
            Next row
            JsonBuilderAppend builder, "]"
        Case ROLE_DATA_ROW
            JsonWriteDataRow builder, value
        Case Else
            If value.IsDictionaryCollection Then
                JsonBuilderAppend builder, "{"
                For pairIndex = 1 To value.InternalPairCount
                    If pairIndex > 1 Then JsonBuilderAppend builder, ","
                    JsonBuilderAppend builder, """"
                    JsonWriteEscaped builder, value.InternalPairKeyText(pairIndex)
                    JsonBuilderAppend builder, """:"
                    JsonWriteValue builder, value.InternalPairValue(pairIndex)
                Next pairIndex
                JsonBuilderAppend builder, "}"
            ElseIf JsonIsSequenceRole(value) Then
                JsonBuilderAppend builder, "["
                Dim firstElement As Boolean
                firstElement = True
                For Each wrapped In value.MaterializeWrapped
                    If Not firstElement Then JsonBuilderAppend builder, ","
                    firstElement = False
                    JsonWriteValue builder, JsonUnwrapValue(wrapped)
                Next wrapped
                JsonBuilderAppend builder, "]"
            Else
                RaiseJsonError "This ROneCOne value is not serializable " & _
                    "to JSON; serialize collections, tables, rows, or views."
            End If
    End Select
End Sub

Private Function JsonIsSequenceRole(ByVal value As ROneCOne) As Boolean
    Select Case value.InternalRole
        Case ROLE_LIST, ROLE_QUERY
            JsonIsSequenceRole = True
        Case Else
            JsonIsSequenceRole = IsGenericCollectionRole(value.InternalRole)
    End Select
End Function

Private Function JsonUnwrapValue(ByVal wrapped As ROneCOne) As Variant
    If wrapped.InternalHasObjectValue Then
        Set JsonUnwrapValue = wrapped.InternalObjectValue
    Else
        JsonUnwrapValue = wrapped.InternalScalarValue
    End If
End Function

Private Function JsonUnwrapObject(ByVal wrapped As ROneCOne) As ROneCOne
    If wrapped.InternalHasObjectValue Then
        Set JsonUnwrapObject = wrapped.InternalObjectValue
    Else
        Set JsonUnwrapObject = wrapped
    End If
End Function

Private Sub JsonWriteDataRow(ByRef builder As JsonTextBuilder, ByVal row As ROneCOne)
    Dim cells As Collection
    Dim columnIndex As Long
    Dim columns As Collection
    Dim table As ROneCOne
    Dim wrapped As ROneCOne

    Set table = row.InternalSource
    Set columns = table.InternalDataColumns
    Set cells = row.InternalDataRows
    JsonBuilderAppend builder, "{"
    For columnIndex = 1 To columns.Count
        If columnIndex > 1 Then JsonBuilderAppend builder, ","
        JsonBuilderAppend builder, """"
        JsonWriteEscaped builder, columns.Item(columnIndex).ColumnName
        JsonBuilderAppend builder, """:"
        Set wrapped = cells.Item(columnIndex)
        If IsNullDataValue(wrapped) Then
            JsonBuilderAppend builder, "null"
        Else
            JsonWriteValue builder, JsonUnwrapValue(wrapped)
        End If
    Next columnIndex
    JsonBuilderAppend builder, "}"
End Sub

Private Sub JsonWriteVbaCollection( _
    ByRef builder As JsonTextBuilder, _
    ByVal values As Collection _
)
    Dim element As Variant
    Dim firstElement As Boolean

    JsonBuilderAppend builder, "["
    firstElement = True
    For Each element In values
        If Not firstElement Then JsonBuilderAppend builder, ","
        firstElement = False
        JsonWriteValue builder, element
    Next element
    JsonBuilderAppend builder, "]"
End Sub

Private Sub JsonWriteVariantArray( _
    ByRef builder As JsonTextBuilder, _
    ByVal values As Variant _
)
    Dim index As Long

    On Error Resume Next
    index = LBound(values, 2)
    If Err.Number = 0 Then
        On Error GoTo 0
        RaiseJsonError "Two-dimensional arrays are not serializable; " & _
            "load them into a DataTable first."
    End If
    Err.Clear
    On Error GoTo 0
    JsonBuilderAppend builder, "["
    For index = LBound(values) To UBound(values)
        If index > LBound(values) Then JsonBuilderAppend builder, ","
        JsonWriteValue builder, values(index)
    Next index
    JsonBuilderAppend builder, "]"
End Sub

Private Sub JsonWriteEscaped(ByRef builder As JsonTextBuilder, ByVal text As String)
    Dim bytes() As Byte
    Dim escapeText As String
    Dim index As Long
    Dim length As Long
    Dim runStart As Long

    length = Len(text)
    If length = 0 Then Exit Sub
    bytes = text
    runStart = 1
    For index = 1 To length
        If bytes((index - 1) * 2 + 1) = 0 Then
            escapeText = vbNullString
            Select Case bytes((index - 1) * 2)
                Case 34
                    escapeText = "\"""
                Case 92
                    escapeText = "\\"
                Case 8
                    escapeText = "\b"
                Case 12
                    escapeText = "\f"
                Case 13
                    escapeText = "\r"
                Case 10
                    escapeText = "\n"
                Case 9
                    escapeText = "\t"
                Case 0 To 7, 11, 14 To 31
                    escapeText = "\u" & Right$("0000" & _
                        Hex$(bytes((index - 1) * 2)), 4)
            End Select
            If Len(escapeText) > 0 Then
                If index > runStart Then
                    JsonBuilderAppend builder, _
                        Mid$(text, runStart, index - runStart)
                End If
                JsonBuilderAppend builder, escapeText
                runStart = index + 1
            End If
        End If
    Next index
    If runStart = 1 Then
        JsonBuilderAppend builder, text
    ElseIf runStart <= length Then
        JsonBuilderAppend builder, Mid$(text, runStart)
    End If
End Sub

Private Function JsonNumberText(ByVal value As Double) As String
    Dim result As String

    ' CStr writes the locale decimal separator, which JSON forbids; the
    ' separator is discovered once and swapped for a dot when needed.
    If value >= -2147483648# And value <= 2147483647# Then
        If Fix(value) = value Then
            JsonNumberText = CStr(CLng(value))
            Exit Function
        End If
    End If
    result = CStr(value)
    If Not mJsonSeparatorKnown Then
        mJsonDecimalSeparator = Mid$(CStr(1.1), 2, 1)
        mJsonSeparatorKnown = True
    End If
    If mJsonDecimalSeparator <> "." Then
        result = VBA.Replace$(result, mJsonDecimalSeparator, ".")
    End If
    JsonNumberText = result
End Function

Private Sub JsonWriteValuePretty( _
    ByRef builder As JsonTextBuilder, _
    ByVal value As Variant, _
    ByVal depth As Long _
)
    Dim element As Variant
    Dim firstEntry As Boolean
    Dim pairIndex As Long
    Dim runtimeValue As ROneCOne
    Dim wrapped As ROneCOne

    ' Only containers carry line structure; scalars and every error path
    ' stay with the compact writer.
    If IsObject(value) Then
        If Not value Is Nothing Then
            If TypeOf value Is ROneCOne Then
                Set runtimeValue = value
                Select Case runtimeValue.InternalRole
                    Case ROLE_DATA_TABLE, ROLE_DATA_VIEW
                        JsonWriteDataRowsPretty builder, _
                            runtimeValue.InternalLiveDataRows, depth
                        Exit Sub
                    Case ROLE_DATA_ROW
                        JsonWriteDataRowPretty builder, runtimeValue, depth
                        Exit Sub
                End Select
                If runtimeValue.IsDictionaryCollection Then
                    If runtimeValue.InternalPairCount = 0 Then
                        JsonBuilderAppend builder, "{}"
                        Exit Sub
                    End If
                    JsonBuilderAppend builder, "{" & vbLf
                    firstEntry = True
                    For pairIndex = 1 To runtimeValue.InternalPairCount
                        If Not firstEntry Then
                            JsonBuilderAppend builder, "," & vbLf
                        End If
                        firstEntry = False
                        JsonBuilderAppend builder, JsonIndent(depth + 1) & """"
                        JsonWriteEscaped builder, _
                            runtimeValue.InternalPairKeyText(pairIndex)
                        JsonBuilderAppend builder, """: "
                        JsonWriteValuePretty builder, _
                            runtimeValue.InternalPairValue(pairIndex), depth + 1
                    Next pairIndex
                    JsonBuilderAppend builder, vbLf & JsonIndent(depth) & "}"
                    Exit Sub
                End If
                If JsonIsSequenceRole(runtimeValue) Then
                    firstEntry = True
                    For Each wrapped In runtimeValue.MaterializeWrapped
                        If firstEntry Then JsonBuilderAppend builder, "[" & vbLf
                        If Not firstEntry Then
                            JsonBuilderAppend builder, "," & vbLf
                        End If
                        firstEntry = False
                        JsonBuilderAppend builder, JsonIndent(depth + 1)
                        JsonWriteValuePretty builder, _
                            JsonUnwrapValue(wrapped), depth + 1
                    Next wrapped
                    If firstEntry Then
                        JsonBuilderAppend builder, "[]"
                    Else
                        JsonBuilderAppend builder, vbLf & JsonIndent(depth) & "]"
                    End If
                    Exit Sub
                End If
            End If
        End If
    End If
    JsonWriteValue builder, value
End Sub

Private Sub JsonWriteDataRowsPretty( _
    ByRef builder As JsonTextBuilder, _
    ByVal liveRows As Collection, _
    ByVal depth As Long _
)
    Dim firstRow As Boolean
    Dim row As ROneCOne

    If liveRows.Count = 0 Then
        JsonBuilderAppend builder, "[]"
        Exit Sub
    End If
    JsonBuilderAppend builder, "[" & vbLf
    firstRow = True
    For Each row In liveRows
        If Not firstRow Then JsonBuilderAppend builder, "," & vbLf
        firstRow = False
        JsonBuilderAppend builder, JsonIndent(depth + 1)
        JsonWriteDataRowPretty builder, row, depth + 1
    Next row
    JsonBuilderAppend builder, vbLf & JsonIndent(depth) & "]"
End Sub

Private Sub JsonWriteDataRowPretty( _
    ByRef builder As JsonTextBuilder, _
    ByVal row As ROneCOne, _
    ByVal depth As Long _
)
    Dim cells As Collection
    Dim columnIndex As Long
    Dim columns As Collection
    Dim table As ROneCOne
    Dim wrapped As ROneCOne

    Set table = row.InternalSource
    Set columns = table.InternalDataColumns
    Set cells = row.InternalDataRows
    If columns.Count = 0 Then
        JsonBuilderAppend builder, "{}"
        Exit Sub
    End If
    JsonBuilderAppend builder, "{" & vbLf
    For columnIndex = 1 To columns.Count
        If columnIndex > 1 Then JsonBuilderAppend builder, "," & vbLf
        JsonBuilderAppend builder, JsonIndent(depth + 1) & """"
        JsonWriteEscaped builder, columns.Item(columnIndex).ColumnName
        JsonBuilderAppend builder, """: "
        Set wrapped = cells.Item(columnIndex)
        If IsNullDataValue(wrapped) Then
            JsonBuilderAppend builder, "null"
        Else
            JsonWriteValuePretty builder, JsonUnwrapValue(wrapped), depth + 1
        End If
    Next columnIndex
    JsonBuilderAppend builder, vbLf & JsonIndent(depth) & "}"
End Sub

Private Function JsonIndent(ByVal depth As Long) As String
    JsonIndent = String$(depth * 2, " ")
End Function

' Splits "$.a.b[0]" into a Collection of step strings. An index step keeps
' its brackets so the walker can tell it from a member of the same name.
Private Function JsonSplitPath(ByVal path As String) As Collection
    Dim cursor As Long
    Dim segment As String
    Dim steps As Collection
    Dim stepChar As String
    Dim trimmed As String

    Set steps = New Collection
    trimmed = Trim$(path)
    If Len(trimmed) = 0 Then RaiseJsonError "The path cannot be empty."
    If Left$(trimmed, 1) <> "$" Then
        RaiseJsonError "The path must begin with '$'."
    End If
    cursor = 2
    Do While cursor <= Len(trimmed)
        stepChar = Mid$(trimmed, cursor, 1)
        If stepChar = "." Then
            cursor = cursor + 1
            segment = vbNullString
            Do While cursor <= Len(trimmed)
                stepChar = Mid$(trimmed, cursor, 1)
                If stepChar = "." Or stepChar = "[" Then Exit Do
                segment = segment & stepChar
                cursor = cursor + 1
            Loop
            If Len(segment) = 0 Then
                RaiseJsonError "The path '" & path & _
                    "' contains an empty member step."
            End If
            steps.Add segment
        ElseIf stepChar = "[" Then
            cursor = cursor + 1
            segment = vbNullString
            Do While cursor <= Len(trimmed)
                stepChar = Mid$(trimmed, cursor, 1)
                If stepChar = "]" Then Exit Do
                segment = segment & stepChar
                cursor = cursor + 1
            Loop
            If cursor > Len(trimmed) Then
                RaiseJsonError "The path '" & path & "' is missing ']'."
            End If
            cursor = cursor + 1
            If Not IsNumeric(segment) Then
                RaiseJsonError "The path index '" & segment & _
                    "' must be a number."
            End If
            steps.Add "[" & CStr(CLng(segment)) & "]"
        Else
            RaiseJsonError "The path '" & path & _
                "' has an unexpected character at " & CStr(cursor) & "."
        End If
    Loop
    Set JsonSplitPath = steps
End Function

' Each entry of stepSets is a Collection of remaining steps for one path.
' Returns the entries whose next step matches, with that step consumed;
' matchedWhole reports that some path ends here, meaning keep everything.
Private Function JsonMatchStep( _
    ByVal stepSets As Collection, _
    ByVal stepText As String, _
    ByRef matchedWhole As Boolean _
) As Collection
    Dim remaining As Collection
    Dim result As Collection
    Dim stepSet As Collection
    Dim index As Long

    matchedWhole = False
    Set result = New Collection
    For Each stepSet In stepSets
        If stepSet.Count > 0 Then
            If StrComp(CStr(stepSet.Item(1)), stepText, vbBinaryCompare) = 0 Then
                If stepSet.Count = 1 Then
                    matchedWhole = True
                Else
                    Set remaining = New Collection
                    For index = 2 To stepSet.Count
                        remaining.Add stepSet.Item(index)
                    Next index
                    result.Add remaining
                End If
            End If
        End If
    Next stepSet
    Set JsonMatchStep = result
End Function

' Reads one value keeping only what the remaining steps ask for. Everything
' else is skipped, so the shape that comes back is a true subset of the
' source document and navigation is identical to a full Deserialize.
Private Function JsonReadFiltered( _
    ByRef reader As JsonReader, _
    ByVal stepSets As Collection _
) As Variant
    Dim elementIndex As Long
    Dim keptList As ROneCOne
    Dim keptObject As ROneCOne
    Dim key As String
    Dim matched As Collection
    Dim matchedWhole As Boolean
    Dim value As Variant

    JsonSkipWhitespace reader
    Select Case JsonCodeAt(reader)
        Case 123
            reader.pos = reader.pos + 1
            Set keptObject = ROneCOne.OrderedDictionaryOf(vbString, vbVariant)
            JsonSkipWhitespace reader
            If JsonCodeAt(reader) = 125 Then
                reader.pos = reader.pos + 1
                Set JsonReadFiltered = keptObject
                Exit Function
            End If
            Do
                JsonSkipWhitespace reader
                If JsonCodeAt(reader) <> 34 Then
                    RaiseJsonError "Expected a quoted member name at " & _
                        "position " & CStr(reader.pos)
                End If
                key = JsonReadString(reader)
                JsonExpectChar reader, 58, ":"
                Set matched = JsonMatchStep(stepSets, key, matchedWhole)
                If matchedWhole Then
                    AssignVariant value, JsonReadValue(reader)
                    keptObject.Item(key) = value
                ElseIf matched.Count > 0 Then
                    AssignVariant value, JsonReadFiltered(reader, matched)
                    keptObject.Item(key) = value
                Else
                    JsonSkipValue reader
                End If
                JsonSkipWhitespace reader
                Select Case JsonCodeAt(reader)
                    Case 44
                        reader.pos = reader.pos + 1
                    Case 125
                        reader.pos = reader.pos + 1
                        Exit Do
                    Case Else
                        RaiseJsonError "Expected ',' or '}' at position " & _
                            CStr(reader.pos)
                End Select
            Loop
            Set JsonReadFiltered = keptObject
        Case 91
            reader.pos = reader.pos + 1
            Set keptList = ROneCOne.ListOf(vbVariant)
            JsonSkipWhitespace reader
            If JsonCodeAt(reader) = 93 Then
                reader.pos = reader.pos + 1
                Set JsonReadFiltered = keptList
                Exit Function
            End If
            Do
                Set matched = JsonMatchStep( _
                    stepSets, "[" & CStr(elementIndex) & "]", matchedWhole)
                If matchedWhole Then
                    AssignVariant value, JsonReadValue(reader)
                    keptList.Add value
                ElseIf matched.Count > 0 Then
                    AssignVariant value, JsonReadFiltered(reader, matched)
                    keptList.Add value
                Else
                    JsonSkipValue reader
                End If
                elementIndex = elementIndex + 1
                JsonSkipWhitespace reader
                Select Case JsonCodeAt(reader)
                    Case 44
                        reader.pos = reader.pos + 1
                    Case 93
                        reader.pos = reader.pos + 1
                        Exit Do
                    Case Else
                        RaiseJsonError "Expected ',' or ']' at position " & _
                            CStr(reader.pos)
                End Select
            Loop
            Set JsonReadFiltered = keptList
        Case Else
            ' A path continues past a scalar, so nothing below it exists.
            JsonSkipValue reader
            JsonReadFiltered = Null
    End Select
End Function

' Reads the single value a path addresses, skipping every sibling on the
' way down. Raises when the path is absent, mirroring GetAttribute.
Private Function JsonReadAtPath( _
    ByRef reader As JsonReader, _
    ByVal steps As Collection, _
    ByVal path As String _
) As Variant
    Dim elementIndex As Long
    Dim found As Boolean
    Dim key As String
    Dim remaining As Collection
    Dim index As Long
    Dim stepText As String
    Dim value As Variant

    If steps.Count = 0 Then
        AssignVariant JsonReadAtPath, JsonReadValue(reader)
        Exit Function
    End If
    stepText = CStr(steps.Item(1))
    Set remaining = New Collection
    For index = 2 To steps.Count
        remaining.Add steps.Item(index)
    Next index

    JsonSkipWhitespace reader
    If Left$(stepText, 1) = "[" Then
        If JsonCodeAt(reader) <> 91 Then
            RaiseJsonError "The path step '" & stepText & _
                "' requires a JSON array."
        End If
        reader.pos = reader.pos + 1
        JsonSkipWhitespace reader
        If JsonCodeAt(reader) = 93 Then
            RaiseJsonError "The path '" & path & "' was not found."
        End If
        Do
            If stepText = "[" & CStr(elementIndex) & "]" Then
                AssignVariant value, JsonReadAtPath(reader, remaining, path)
                AssignVariant JsonReadAtPath, value
                Exit Function
            End If
            JsonSkipValue reader
            elementIndex = elementIndex + 1
            JsonSkipWhitespace reader
            Select Case JsonCodeAt(reader)
                Case 44
                    reader.pos = reader.pos + 1
                Case 93
                    RaiseJsonError "The path '" & path & "' was not found."
                Case Else
                    RaiseJsonError "Expected ',' or ']' at position " & _
                        CStr(reader.pos)
            End Select
        Loop
    End If

    If JsonCodeAt(reader) <> 123 Then
        RaiseJsonError "The path step '" & stepText & _
            "' requires a JSON object."
    End If
    reader.pos = reader.pos + 1
    JsonSkipWhitespace reader
    If JsonCodeAt(reader) = 125 Then
        RaiseJsonError "The path '" & path & "' was not found."
    End If
    Do
        JsonSkipWhitespace reader
        If JsonCodeAt(reader) <> 34 Then
            RaiseJsonError "Expected a quoted member name at position " & _
                CStr(reader.pos)
        End If
        key = JsonReadString(reader)
        JsonExpectChar reader, 58, ":"
        If StrComp(key, stepText, vbBinaryCompare) = 0 Then
            AssignVariant value, JsonReadAtPath(reader, remaining, path)
            AssignVariant JsonReadAtPath, value
            found = True
            Exit Function
        End If
        JsonSkipValue reader
        JsonSkipWhitespace reader
        Select Case JsonCodeAt(reader)
            Case 44
                reader.pos = reader.pos + 1
            Case 125
                Exit Do
            Case Else
                RaiseJsonError "Expected ',' or '}' at position " & _
                    CStr(reader.pos)
        End Select
    Loop
    If Not found Then
        RaiseJsonError "The path '" & path & "' was not found."
    End If
End Function

Private Function ResolveJsonArray( _
    ByVal root As Variant, _
    ByVal arrayPath As String _
) As ROneCOne
    Dim current As Variant
    Dim resolved As ROneCOne

    AssignVariant current, ResolveJsonPath(root, arrayPath)
    If IsObject(current) Then
        If TypeOf current Is ROneCOne Then
            Set resolved = current
            If resolved.InternalRole = ROLE_LIST Then
                Set ResolveJsonArray = resolved
                Exit Function
            End If
        End If
    End If
    RaiseJsonError "The path '" & arrayPath & _
        "' did not resolve to a JSON array."
End Function

Private Function ResolveJsonPath( _
    ByVal root As Variant, _
    ByVal path As String _
) As Variant
    Dim current As Variant
    Dim cursor As Long
    Dim elementIndex As Long
    Dim segment As String
    Dim stepChar As String
    Dim target As ROneCOne
    Dim trimmed As String

    ' Supports "$", member steps ("$.data.items"), and zero-based index
    ' steps ("$.rows[0].values"). Member names with dots are not addressable.
    trimmed = Trim$(path)
    If Len(trimmed) = 0 Then RaiseJsonError "The path cannot be empty."
    If Left$(trimmed, 1) <> "$" Then
        RaiseJsonError "The path must begin with '$'."
    End If
    AssignVariant current, root
    cursor = 2
    Do While cursor <= Len(trimmed)
        stepChar = Mid$(trimmed, cursor, 1)
        If stepChar = "." Then
            cursor = cursor + 1
            segment = vbNullString
            Do While cursor <= Len(trimmed)
                stepChar = Mid$(trimmed, cursor, 1)
                If stepChar = "." Or stepChar = "[" Then Exit Do
                segment = segment & stepChar
                cursor = cursor + 1
            Loop
            If Len(segment) = 0 Then
                RaiseJsonError "The path '" & path & _
                    "' contains an empty member step."
            End If
            If Not JsonValueIsDictionary(current) Then
                RaiseJsonError "The path step '" & segment & _
                    "' requires a JSON object."
            End If
            Set target = current
            If Not target.ContainsKey(segment) Then
                RaiseJsonError "The path member '" & segment & _
                    "' was not found."
            End If
            AssignVariant current, target.Item(segment)
        ElseIf stepChar = "[" Then
            cursor = cursor + 1
            segment = vbNullString
            Do While cursor <= Len(trimmed)
                stepChar = Mid$(trimmed, cursor, 1)
                If stepChar = "]" Then Exit Do
                segment = segment & stepChar
                cursor = cursor + 1
            Loop
            If stepChar <> "]" Or Len(segment) = 0 Then
                RaiseJsonError "The path '" & path & _
                    "' contains an invalid index step."
            End If
            cursor = cursor + 1
            elementIndex = CLng(segment)
            If Not JsonValueIsList(current) Then
                RaiseJsonError "The path index [" & segment & _
                    "] requires a JSON array."
            End If
            Set target = current
            AssignVariant current, target.Item(elementIndex)
        Else
            RaiseJsonError "The path '" & path & _
                "' contains an unexpected character '" & stepChar & "'."
        End If
    Loop
    AssignVariant ResolveJsonPath, current
End Function

Private Function JsonValueIsDictionary(ByVal value As Variant) As Boolean
    Dim candidate As ROneCOne

    If Not IsObject(value) Then Exit Function
    If Not TypeOf value Is ROneCOne Then Exit Function
    Set candidate = value
    JsonValueIsDictionary = candidate.IsDictionaryCollection
End Function

Private Function JsonValueIsList(ByVal value As Variant) As Boolean
    Dim candidate As ROneCOne

    If Not IsObject(value) Then Exit Function
    If Not TypeOf value Is ROneCOne Then Exit Function
    Set candidate = value
    JsonValueIsList = (candidate.InternalRole = ROLE_LIST)
End Function

Private Sub RequireJsonObjectValue(ByVal value As Variant, ByVal memberName As String)
    If Not JsonValueIsDictionary(value) Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            "The JSON value must be an object at this position."
    End If
End Sub

Private Sub ValidateJsonObjectFactory( _
    ByVal factory As ROneCOne, _
    ByVal memberName As String _
)
    If factory Is Nothing Then
        RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
            "A zero-argument factory delegate is required."
    End If
    If Not IsDelegateRole(factory.InternalRole) Then
        RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
            "The factory must be a Func delegate returning a new instance."
    End If
    If factory.Arity <> 0 And factory.Arity <> -1 Then
        RaiseContractError ERROR_ARITY_MISMATCH, memberName, _
            "The factory delegate accepts no arguments."
    End If
    ' The factory returns an instance, so it must declare an object result;
    ' otherwise the scalar invocation path would coerce the object and fail.
    If Not factory.InternalHasReturnSignature Or _
        factory.InternalReturnType <> vbObject Then
        RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
            "The factory must declare an object result, for example " & _
            "ROneCOne.Func(""Module.NewItem"").Takes().Returns(vbObject)."
    End If
End Sub

Private Function InvokeJsonObjectFactory( _
    ByVal factory As ROneCOne, _
    ByVal memberName As String _
) As Object
    Dim wrapped As ROneCOne

    Set wrapped = factory.InvokeValue(New Collection)
    If Not wrapped.InternalHasObjectValue Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            "The factory delegate must return an object instance."
    End If
    Set InvokeJsonObjectFactory = wrapped.InternalObjectValue
End Function

Private Sub BindJsonObjectInto(ByVal instance As Object, ByVal source As Variant)
    Dim pairIndex As Long
    Dim tree As ROneCOne
    Dim value As Variant

    ' Scalar members bind by name; null, object, and array members are left
    ' for the caller, mirroring a flat DTO bind. A member the target does
    ' not expose is skipped rather than raising, matching how
    ' System.Text.Json ignores unknown members: a real API response carries
    ' far more than any one class models.
    Set tree = source
    For pairIndex = 1 To tree.InternalPairCount
        AssignVariant value, tree.InternalPairValue(pairIndex)
        If Not IsObject(value) Then
            If Not IsNull(value) Then
                On Error Resume Next
                CallByName instance, tree.InternalPairKeyText(pairIndex), _
                    VbLet, value
                If Err.Number <> 0 Then Err.Clear
                On Error GoTo 0
            End If
        End If
    Next pairIndex
End Sub

Private Sub AssignInstanceProperty( _
    ByVal instance As Object, _
    ByVal propertyName As String, _
    ByVal wrapped As ROneCOne _
)
    If wrapped.InternalHasObjectValue Then
        CallByName instance, propertyName, VbSet, wrapped.InternalObjectValue
    Else
        CallByName instance, propertyName, VbLet, wrapped.InternalScalarValue
    End If
End Sub

Private Function BuildTableFromJsonArray( _
    ByVal rows As ROneCOne, _
    ByVal tableName As String _
) As ROneCOne
    Dim columnIndex As Long
    Dim columnNames As ROneCOne
    Dim columnTypes As ROneCOne
    Dim element As Variant
    Dim flatRows As Collection
    Dim flat As ROneCOne
    Dim table As ROneCOne
    Dim values() As Variant

    ' Pass one flattens each object (nested objects become dotted columns,
    ' nested arrays become JSON text) and unifies column types; pass two
    ' loads the rows in discovered column order.
    Set columnNames = ROneCOne.OrderedDictionaryOf(vbString, vbLong)
    Set flatRows = New Collection
    For Each element In rows
        RequireJsonObjectValue element, "DeserializeTable"
        Set flat = ROneCOne.OrderedDictionaryOf(vbString, vbVariant)
        FlattenJsonObjectInto element, vbNullString, flat, columnNames
        flatRows.Add flat
    Next element
    Set table = ROneCOne.DataTable(tableName)
    For columnIndex = 1 To columnNames.InternalPairCount
        table.Column columnNames.InternalPairKeyText(columnIndex), _
            JsonFinalColumnType(CLng(columnNames.InternalPairValue(columnIndex)))
    Next columnIndex
    For Each flat In flatRows
        If columnNames.InternalPairCount > 0 Then
            ReDim values(0 To columnNames.InternalPairCount - 1)
            For columnIndex = 1 To columnNames.InternalPairCount
                If flat.ContainsKey( _
                    columnNames.InternalPairKeyText(columnIndex)) Then
                    values(columnIndex - 1) = flat.Item( _
                        columnNames.InternalPairKeyText(columnIndex))
                Else
                    values(columnIndex - 1) = Null
                End If
            Next columnIndex
            table.LoadRow values
        Else
            table.LoadRow Array()
        End If
    Next flat
    Set BuildTableFromJsonArray = table
End Function

Private Sub FlattenJsonObjectInto( _
    ByVal source As ROneCOne, _
    ByVal prefix As String, _
    ByVal flat As ROneCOne, _
    ByVal columnNames As ROneCOne _
)
    Dim columnName As String
    Dim pairIndex As Long
    Dim value As Variant

    For pairIndex = 1 To source.InternalPairCount
        If Len(prefix) = 0 Then
            columnName = source.InternalPairKeyText(pairIndex)
        Else
            columnName = prefix & "." & source.InternalPairKeyText(pairIndex)
        End If
        AssignVariant value, source.InternalPairValue(pairIndex)
        If JsonValueIsDictionary(value) Then
            FlattenJsonObjectInto value, columnName, flat, columnNames
        Else
            If JsonValueIsList(value) Then
                value = ROneCOne.Json.Serialize(value)
            End If
            flat.Item(columnName) = value
            MergeJsonColumnType columnNames, columnName, value
        End If
    Next pairIndex
End Sub

Private Sub MergeJsonColumnType( _
    ByVal columnNames As ROneCOne, _
    ByVal columnName As String, _
    ByVal value As Variant _
)
    Dim candidate As Long
    Dim current As Long

    If IsNull(value) Then
        candidate = 0
    Else
        candidate = VarTypeOf(value)
    End If
    If Not columnNames.ContainsKey(columnName) Then
        columnNames.Add columnName, candidate
        Exit Sub
    End If
    current = CLng(columnNames.Item(columnName))
    If current = candidate Or candidate = 0 Then Exit Sub
    If current = 0 Then
        columnNames.Item(columnName) = candidate
    ElseIf JsonIsNumericType(current) And JsonIsNumericType(candidate) Then
        columnNames.Item(columnName) = _
            JsonWiderNumericType(current, candidate)
    Else
        columnNames.Item(columnName) = vbVariant
    End If
End Sub

Private Function JsonIsNumericType(ByVal typeCode As Long) As Boolean
    Select Case typeCode
        Case vbLong, vbLongLong, vbDouble
            JsonIsNumericType = True
    End Select
End Function

Private Function JsonWiderNumericType( _
    ByVal leftType As Long, _
    ByVal rightType As Long _
) As Long
    If leftType = vbDouble Or rightType = vbDouble Then
        JsonWiderNumericType = vbDouble
    ElseIf leftType = vbLongLong Or rightType = vbLongLong Then
        JsonWiderNumericType = vbLongLong
    Else
        JsonWiderNumericType = vbLong
    End If
End Function

Private Function JsonFinalColumnType(ByVal typeCode As Long) As Long
    If typeCode = 0 Then
        JsonFinalColumnType = vbVariant
    Else
        JsonFinalColumnType = typeCode
    End If
End Function

Private Function BuildTableFromObjects( _
    ByVal source As Variant, _
    ByVal propertyNames As Variant, _
    Optional ByVal tableName As String = "Objects" _
) As ROneCOne
    Dim columnIndex As Long
    Dim columnTyped() As Boolean
    Dim columnTypes() As Long
    Dim instance As Object
    Dim instances As Collection
    Dim names As Collection
    Dim table As ROneCOne
    Dim value As Variant
    Dim values() As Variant

    ' Explicit property names stand in for the reflection VBA does not have;
    ' column types come from the first non-null value each property yields.
    Set names = CollectJsonPropertyNames(propertyNames, _
        "DataTableFromObjects")
    Set instances = CollectJsonInstances(source, "DataTableFromObjects")
    ReDim columnTypes(1 To names.Count)
    ReDim columnTyped(1 To names.Count)
    For Each instance In instances
        For columnIndex = 1 To names.Count
            If Not columnTyped(columnIndex) Then
                value = CallByName(instance, names.Item(columnIndex), VbGet)
                If Not IsNull(value) And Not IsEmpty(value) And _
                    Not IsObject(value) Then
                    columnTypes(columnIndex) = VarTypeOf(value)
                    columnTyped(columnIndex) = True
                End If
            End If
        Next columnIndex
    Next instance
    Set table = ROneCOne.DataTable(tableName)
    For columnIndex = 1 To names.Count
        If Not columnTyped(columnIndex) Then
            columnTypes(columnIndex) = vbVariant
        End If
        table.Column CStr(names.Item(columnIndex)), columnTypes(columnIndex)
    Next columnIndex
    For Each instance In instances
        ReDim values(0 To names.Count - 1)
        For columnIndex = 1 To names.Count
            value = CallByName(instance, names.Item(columnIndex), VbGet)
            If IsEmpty(value) Then value = Null
            values(columnIndex - 1) = value
        Next columnIndex
        table.LoadRow values
    Next instance
    Set BuildTableFromObjects = table
End Function

Private Function CollectJsonPropertyNames( _
    ByVal propertyNames As Variant, _
    ByVal memberName As String _
) As Collection
    Dim element As Variant
    Dim names As Collection

    Set names = New Collection
    For Each element In CollectSourceElements(propertyNames, memberName)
        If IsObject(element) Or VarTypeOf(element) <> vbString Then
            RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
                "Every property name must be text."
        End If
        names.Add CStr(element)
    Next element
    If names.Count = 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
            "At least one property name is required."
    End If
    Set CollectJsonPropertyNames = names
End Function

Private Function CollectJsonInstances( _
    ByVal source As Variant, _
    ByVal memberName As String _
) As Collection
    Dim element As Variant
    Dim instances As Collection

    Set instances = New Collection
    For Each element In CollectSourceElements(source, memberName)
        If Not IsObject(element) Then
            RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
                "Every source element must be an object instance."
        End If
        instances.Add element
    Next element
    Set CollectJsonInstances = instances
End Function

Private Function CollectSourceElements( _
    ByVal values As Variant, _
    ByVal memberName As String _
) As Collection
    Dim element As Variant
    Dim elements As Collection
    Dim sequence As ROneCOne

    ' Element-typed helpers such as WrapElement belong to configured
    ' collections; this JSON entry point runs on the root instance, so the
    ' incoming values are normalized here without element validation.
    Set elements = New Collection
    If IsObject(values) Then
        If values Is Nothing Then
            RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
                "The incoming values cannot be Nothing."
        End If
        If TypeOf values Is ROneCOne Then
            Set sequence = values
            RequireSequence sequence, memberName
            For Each element In sequence
                elements.Add element
            Next element
            Set CollectSourceElements = elements
            Exit Function
        End If
        If TypeName(values) = "Collection" Then
            For Each element In values
                elements.Add element
            Next element
            Set CollectSourceElements = elements
            Exit Function
        End If
    ElseIf IsArrayValue(values) Then
        For Each element In values
            elements.Add element
        Next element
        Set CollectSourceElements = elements
        Exit Function
    End If
    RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
        "Expected a ROneCOne sequence, VBA array, or Collection."
End Function

Private Sub RequireFileFactory(ByVal memberName As String)
    If mRole <> ROLE_FILE_FACTORY Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            memberName & " requires the File surface."
    End If
End Sub

Private Sub RequireDirectoryFactory(ByVal memberName As String)
    If mRole <> ROLE_DIRECTORY_FACTORY Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            memberName & " requires the Directory surface."
    End If
End Sub

Private Sub RequirePathFactory(ByVal memberName As String)
    If mRole <> ROLE_PATH_FACTORY Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            memberName & " requires the Path surface."
    End If
End Sub

Private Sub RaiseIoError(ByVal message As String)
    Err.Raise ERROR_IO, "ROneCOne.IOException", message
End Sub

Private Function RequirePathText( _
    ByVal value As Variant, _
    ByVal memberName As String _
) As String
    If Not IsMissing(value) Then
        If VarTypeOf(value) = vbString Then
            If Len(Trim$(CStr(value))) > 0 Then
                RequirePathText = CStr(value)
                Exit Function
            End If
        End If
    End If
    RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
        "A file system path is required."
End Function

Private Function CreateFileSystemObject() As Object
    Set CreateFileSystemObject = CreateObject("Scripting.FileSystemObject")
End Function

Private Function CreateStreamObject(ByVal streamType As Long) As Object
    Set CreateStreamObject = CreateObject(STREAM_PROG_ID)
    CreateStreamObject.Type = streamType
    CreateStreamObject.Open
End Function

Private Function NormalizeTextEncoding( _
    ByVal encoding As String, _
    ByVal memberName As String _
) As String
    Select Case LCase$(Trim$(encoding))
        Case "utf-8"
            NormalizeTextEncoding = "utf-8"
        Case "utf-16"
            NormalizeTextEncoding = "unicode"
        Case "windows-1252"
            NormalizeTextEncoding = "windows-1252"
        Case Else
            RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
                "Encoding must be utf-8, utf-16, or windows-1252."
    End Select
End Function

Private Function TextBomSkipBytes(ByVal charset As String) As Long
    If charset = "utf-8" Then
        TextBomSkipBytes = 3
    ElseIf charset = "unicode" Then
        TextBomSkipBytes = 2
    End If
End Function

Private Function SniffTextCharset( _
    ByVal filePath As String, _
    ByVal requestedCharset As String _
) As String
    Dim fileNumber As Integer
    Dim head(0 To 2) As Byte
    Dim headLength As Long

    ' A byte-order mark decides the decoder; without one, the caller's
    ' requested encoding stands. The native read touches at most 3 bytes.
    SniffTextCharset = requestedCharset
    fileNumber = FreeFile
    Open filePath For Binary Access Read As #fileNumber
    headLength = LOF(fileNumber)
    If headLength > 3 Then headLength = 3
    If headLength >= 2 Then Get #fileNumber, 1, head
    Close #fileNumber
    If headLength < 2 Then Exit Function
    If head(0) = &HFF And head(1) = &HFE Then
        SniffTextCharset = "unicode"
    ElseIf head(0) = &HFE And head(1) = &HFF Then
        SniffTextCharset = "bigendianunicode"
    ElseIf headLength >= 3 Then
        If head(0) = &HEF And head(1) = &HBB And head(2) = &HBF Then
            SniffTextCharset = "utf-8"
        End If
    End If
End Function

Private Function IoReadTextFile( _
    ByVal filePath As String, _
    ByVal charset As String, _
    ByVal memberName As String _
) As String
    Dim content As Variant
    Dim stream As Object

    IoRequireExistingFile filePath, memberName
    On Error GoTo ReadFailed
    Set stream = CreateStreamObject(STREAM_TYPE_TEXT)
    stream.Charset = SniffTextCharset(filePath, charset)
    stream.LoadFromFile filePath
    content = stream.ReadText(-1)
    stream.Close
    On Error GoTo 0
    If Not IsNull(content) Then IoReadTextFile = CStr(content)
    Exit Function

ReadFailed:
    RaiseIoError "The file '" & filePath & "' could not be read: " & _
        Err.Description
End Function

Private Sub IoWriteTextFile( _
    ByVal filePath As String, _
    ByVal contents As String, _
    ByVal charset As String, _
    ByVal append As Boolean, _
    ByVal memberName As String _
)
    Dim encoded As Variant
    Dim skipBytes As Long
    Dim target As Object

    ' UTF-8 output never carries a byte-order mark; UTF-16 keeps its mark at
    ' the start of a file, and appended text never embeds one mid-file.
    If append And CreateFileSystemObject().FileExists(filePath) Then
        encoded = IoEncodeText(contents, charset, TextBomSkipBytes(charset))
        On Error GoTo WriteFailed
        Set target = CreateStreamObject(STREAM_TYPE_BINARY)
        target.LoadFromFile filePath
        target.Position = target.Size
    Else
        If charset = "unicode" Then
            skipBytes = 0
        Else
            skipBytes = TextBomSkipBytes(charset)
        End If
        encoded = IoEncodeText(contents, charset, skipBytes)
        On Error GoTo WriteFailed
        Set target = CreateStreamObject(STREAM_TYPE_BINARY)
    End If
    If UBound(encoded) >= LBound(encoded) Then target.Write encoded
    target.SaveToFile filePath, STREAM_SAVE_OVERWRITE
    target.Close
    On Error GoTo 0
    Exit Sub

WriteFailed:
    RaiseIoError "The file '" & filePath & "' could not be written: " & _
        Err.Description
End Sub

Private Function IoEncodeText( _
    ByVal contents As String, _
    ByVal charset As String, _
    ByVal skipBytes As Long _
) As Variant
    Dim binaryStream As Object
    Dim encoded As Variant
    Dim textStream As Object

    ' An empty document has no bytes to encode in any charset, and passing
    ' VBA's null BSTR (vbNullString) to WriteText makes ADODB.Stream raise
    ' error 3001, so an empty write is skipped rather than attempted.
    If Len(contents) = 0 Then
        IoEncodeText = EmptyByteArray()
        Exit Function
    End If
    Set textStream = CreateStreamObject(STREAM_TYPE_TEXT)
    textStream.Charset = charset
    textStream.WriteText contents
    If skipBytes > textStream.Size Then skipBytes = textStream.Size
    textStream.Position = skipBytes
    Set binaryStream = CreateStreamObject(STREAM_TYPE_BINARY)
    textStream.CopyTo binaryStream
    textStream.Close
    If binaryStream.Size > 0 Then
        binaryStream.Position = 0
        encoded = binaryStream.Read(-1)
    Else
        encoded = EmptyByteArray()
    End If
    binaryStream.Close
    IoEncodeText = encoded
End Function

Private Function IoDecodeText( _
    ByRef bytes() As Byte, _
    ByVal charset As String _
) As String
    Dim binaryStream As Object

    ' The mirror of IoEncodeText: raw bytes become text through one stream
    ' whose type flips from binary to text at position zero. Callers pass an
    ' allocated, non-empty array.
    Set binaryStream = CreateStreamObject(STREAM_TYPE_BINARY)
    binaryStream.Write bytes
    binaryStream.Position = 0
    binaryStream.Type = STREAM_TYPE_TEXT
    binaryStream.Charset = charset
    IoDecodeText = binaryStream.ReadText(-1)
    binaryStream.Close
End Function

Private Function EmptyByteArray() As Variant
    Dim bytes() As Byte

    bytes = vbNullString
    EmptyByteArray = bytes
End Function

Private Function IoReadAllBytesCore( _
    ByVal filePath As String, _
    ByVal memberName As String _
) As Variant
    Dim stream As Object

    IoRequireExistingFile filePath, memberName
    On Error GoTo ReadFailed
    Set stream = CreateStreamObject(STREAM_TYPE_BINARY)
    stream.LoadFromFile filePath
    If stream.Size > 0 Then
        IoReadAllBytesCore = stream.Read(-1)
    Else
        IoReadAllBytesCore = EmptyByteArray()
    End If
    stream.Close
    On Error GoTo 0
    Exit Function

ReadFailed:
    RaiseIoError "The file '" & filePath & "' could not be read: " & _
        Err.Description
End Function

Private Sub IoWriteAllBytesCore( _
    ByVal filePath As String, _
    ByVal bytes As Variant, _
    ByVal memberName As String _
)
    Dim stream As Object

    On Error GoTo WriteFailed
    Set stream = CreateStreamObject(STREAM_TYPE_BINARY)
    If UBound(bytes) >= LBound(bytes) Then stream.Write bytes
    stream.SaveToFile filePath, STREAM_SAVE_OVERWRITE
    stream.Close
    On Error GoTo 0
    Exit Sub

WriteFailed:
    RaiseIoError "The file '" & filePath & "' could not be written: " & _
        Err.Description
End Sub

Private Sub IoRequireExistingFile( _
    ByVal filePath As String, _
    ByVal memberName As String _
)
    If Not CreateFileSystemObject().FileExists(filePath) Then
        RaiseIoError "The file '" & filePath & "' was not found."
    End If
End Sub

Private Sub FileDeleteCore(ByVal filePath As String)
    Dim fso As Object

    Set fso = CreateFileSystemObject()
    If Not fso.FileExists(filePath) Then Exit Sub
    On Error GoTo DeleteFailed
    fso.DeleteFile filePath, True
    On Error GoTo 0
    Exit Sub

DeleteFailed:
    RaiseIoError "The file '" & filePath & "' could not be deleted: " & _
        Err.Description
End Sub

Private Sub FileCopyCore( _
    ByVal sourcePath As String, _
    ByVal destinationPath As String, _
    ByVal overwrite As Boolean _
)
    Dim fso As Object

    Set fso = CreateFileSystemObject()
    If Not fso.FileExists(sourcePath) Then
        RaiseIoError "The file '" & sourcePath & "' was not found."
    End If
    If Not overwrite Then
        If fso.FileExists(destinationPath) Then
            RaiseIoError "The file '" & destinationPath & _
                "' already exists."
        End If
    End If
    On Error GoTo CopyFailed
    fso.CopyFile sourcePath, destinationPath, overwrite
    On Error GoTo 0
    Exit Sub

CopyFailed:
    RaiseIoError "The file '" & sourcePath & "' could not be copied: " & _
        Err.Description
End Sub

Private Sub FileMoveCore( _
    ByVal sourcePath As String, _
    ByVal destinationPath As String _
)
    Dim fso As Object

    Set fso = CreateFileSystemObject()
    If Not fso.FileExists(sourcePath) Then
        RaiseIoError "The file '" & sourcePath & "' was not found."
    End If
    If fso.FileExists(destinationPath) Then
        RaiseIoError "The file '" & destinationPath & "' already exists."
    End If
    On Error GoTo MoveFailed
    fso.MoveFile sourcePath, destinationPath
    On Error GoTo 0
    Exit Sub

MoveFailed:
    RaiseIoError "The file '" & sourcePath & "' could not be moved: " & _
        Err.Description
End Sub

Private Sub DirectoryCreateCore( _
    ByVal directoryPath As String, _
    ByVal memberName As String _
)
    Dim fso As Object
    Dim parentPath As String

    Set fso = CreateFileSystemObject()
    If fso.FolderExists(directoryPath) Then Exit Sub
    parentPath = fso.GetParentFolderName(directoryPath)
    If Len(parentPath) > 0 Then
        If Not fso.FolderExists(parentPath) Then
            DirectoryCreateCore parentPath, memberName
        End If
    End If
    On Error GoTo CreateFailed
    fso.CreateFolder directoryPath
    On Error GoTo 0
    Exit Sub

CreateFailed:
    RaiseIoError "The folder '" & directoryPath & _
        "' could not be created: " & Err.Description
End Sub

Private Sub DirectoryDeleteCore( _
    ByVal directoryPath As String, _
    ByVal recursive As Boolean _
)
    Dim folderObject As Object
    Dim fso As Object

    Set fso = CreateFileSystemObject()
    If Not fso.FolderExists(directoryPath) Then
        RaiseIoError "The folder '" & directoryPath & "' was not found."
    End If
    If Not recursive Then
        Set folderObject = fso.GetFolder(directoryPath)
        If folderObject.Files.Count > 0 Or _
            folderObject.SubFolders.Count > 0 Then
            RaiseIoError "The folder '" & directoryPath & _
                "' is not empty; pass recursive to delete it."
        End If
    End If
    On Error GoTo DeleteFailed
    fso.DeleteFolder directoryPath, True
    On Error GoTo 0
    Exit Sub

DeleteFailed:
    RaiseIoError "The folder '" & directoryPath & _
        "' could not be deleted: " & Err.Description
End Sub

Private Function DirectoryEntriesCore( _
    ByVal directoryPath As String, _
    ByVal searchPattern As String, _
    ByVal allDirectories As Boolean, _
    ByVal wantFiles As Boolean, _
    ByVal memberName As String _
) As ROneCOne
    Dim entry As Variant
    Dim fso As Object
    Dim index As Long
    Dim patternLower As String
    Dim results As Collection
    Dim sorted() As String
    Dim values As ROneCOne

    Set fso = CreateFileSystemObject()
    If Not fso.FolderExists(directoryPath) Then
        RaiseIoError "The folder '" & directoryPath & "' was not found."
    End If
    patternLower = LCase$(VBA.Replace$(searchPattern, "[", "[[]"))
    patternLower = VBA.Replace$(patternLower, "#", "[#]")
    Set results = New Collection
    CollectFolderEntries fso.GetFolder(directoryPath), patternLower, _
        allDirectories, wantFiles, results
    Set values = ROneCOne.ListOf(vbString)
    If results.Count > 0 Then
        ReDim sorted(1 To results.Count)
        index = 0
        For Each entry In results
            index = index + 1
            sorted(index) = CStr(entry)
        Next entry
        SortTextArray sorted, 1, results.Count
        For index = 1 To results.Count
            values.Add sorted(index)
        Next index
    End If
    Set DirectoryEntriesCore = values
End Function

Private Sub CollectFolderEntries( _
    ByVal folderObject As Object, _
    ByVal patternLower As String, _
    ByVal allDirectories As Boolean, _
    ByVal wantFiles As Boolean, _
    ByVal results As Collection _
)
    Dim entry As Object

    If wantFiles Then
        For Each entry In folderObject.Files
            If LCase$(entry.Name) Like patternLower Then
                results.Add CStr(entry.Path)
            End If
        Next entry
    Else
        For Each entry In folderObject.SubFolders
            If LCase$(entry.Name) Like patternLower Then
                results.Add CStr(entry.Path)
            End If
        Next entry
    End If
    If allDirectories Then
        For Each entry In folderObject.SubFolders
            CollectFolderEntries entry, patternLower, True, wantFiles, _
                results
        Next entry
    End If
End Sub

Private Sub SortTextArray( _
    ByRef values() As String, _
    ByVal low As Long, _
    ByVal high As Long _
)
    Dim lowerIndex As Long
    Dim pivot As String
    Dim swapValue As String
    Dim upperIndex As Long

    If low >= high Then Exit Sub
    lowerIndex = low
    upperIndex = high
    pivot = values((low + high) \ 2)
    Do While lowerIndex <= upperIndex
        Do While StrComp(values(lowerIndex), pivot, vbTextCompare) < 0
            lowerIndex = lowerIndex + 1
        Loop
        Do While StrComp(values(upperIndex), pivot, vbTextCompare) > 0
            upperIndex = upperIndex - 1
        Loop
        If lowerIndex <= upperIndex Then
            swapValue = values(lowerIndex)
            values(lowerIndex) = values(upperIndex)
            values(upperIndex) = swapValue
            lowerIndex = lowerIndex + 1
            upperIndex = upperIndex - 1
        End If
    Loop
    SortTextArray values, low, upperIndex
    SortTextArray values, lowerIndex, high
End Sub

Private Function CombinePathParts(ByVal parts As Variant) As String
    Dim part As Variant
    Dim partText As String
    Dim result As String

    For Each part In parts
        partText = CStr(part)
        If Len(partText) > 0 Then
            If PathIsRooted(partText) Or Len(result) = 0 Then
                result = partText
            ElseIf Right$(result, 1) = "\" Or Right$(result, 1) = "/" Then
                result = result & partText
            Else
                result = result & "\" & partText
            End If
        End If
    Next part
    CombinePathParts = result
End Function

Private Function PathIsRooted(ByVal pathText As String) As Boolean
    If Len(pathText) = 0 Then Exit Function
    If Left$(pathText, 1) = "\" Or Left$(pathText, 1) = "/" Then
        PathIsRooted = True
    ElseIf Len(pathText) >= 2 Then
        PathIsRooted = (Mid$(pathText, 2, 1) = ":")
    End If
End Function

Private Function PathLastSeparator(ByVal pathText As String) As Long
    Dim backslashAt As Long
    Dim slashAt As Long

    backslashAt = InStrRev(pathText, "\")
    slashAt = InStrRev(pathText, "/")
    If backslashAt > slashAt Then
        PathLastSeparator = backslashAt
    Else
        PathLastSeparator = slashAt
    End If
End Function

Private Function TrimTrailingSeparators(ByVal pathText As String) As String
    TrimTrailingSeparators = pathText
    Do While Len(TrimTrailingSeparators) > 1
        If Right$(TrimTrailingSeparators, 1) = "\" Or _
            Right$(TrimTrailingSeparators, 1) = "/" Then
            TrimTrailingSeparators = _
                Left$(TrimTrailingSeparators, Len(TrimTrailingSeparators) - 1)
        Else
            Exit Do
        End If
    Loop
End Function

Private Sub RaiseCsvError(ByVal message As String)
    Err.Raise ERROR_CSV_PARSE, "ROneCOne.CsvException", message
End Sub

Friend Sub ConfigureProcessSession( _
    ByVal command As String, _
    ByVal workingDirectory As String, _
    ByVal encodingName As String _
)
    Dim clientInput As LongPtr
    Dim commandLine As String
    Dim directoryPointer As LongPtr
    Dim errWrite As LongPtr
    Dim outWrite As LongPtr
    Dim pipeName As String
    Dim processInformation As PROCESS_INFORMATION
    Dim security As SECURITY_ATTRIBUTES
    Dim startupInformation As STARTUPINFOW

    mRole = ROLE_PROCESS_SESSION
    mName = command
    mSessionCharset = encodingName
    If Len(Trim$(mSessionCharset)) = 0 Then
        mSessionCharset = "IBM" & CStr(GetOEMCP())
    End If
    Set mSessionWriteQueue = New Collection

    security.nLength = LenB(security)
    security.bInheritHandle = 1

    ' Standard input is an overlapped named pipe so a write larger than the
    ' buffer pends instead of blocking Excel. The child end is inheritable;
    ' the end this class keeps is not.
    pipeName = "\\.\pipe\ronecone_" & VBA.Replace(ROneCOne.Guid.NewGuid, "-", "")
    mSessionStdIn = CreateNamedPipeW(StrPtr(pipeName), _
        PIPE_ACCESS_OUTBOUND Or FILE_FLAG_OVERLAPPED, 0, 1, _
        SESSION_PIPE_BYTES, SESSION_PIPE_BYTES, 0, security)
    If mSessionStdIn = -1 Then
        mSessionStdIn = 0
        RaiseProcessError "The input pipe could not be created."
    End If
    clientInput = CreateFileW(StrPtr(pipeName), GENERIC_READ, 0, _
        security, OPEN_EXISTING, 0, 0)
    If clientInput = -1 Then
        SessionCloseHandles
        RaiseProcessError "The input pipe could not be opened."
    End If
    SetHandleInformation mSessionStdIn, HANDLE_FLAG_INHERIT, 0
    ' A client that connected before this call reports ERROR_PIPE_CONNECTED,
    ' which is success for our purposes.
    ConnectNamedPipe mSessionStdIn, 0

    If CreatePipe(mSessionStdOut, outWrite, security, SESSION_PIPE_BYTES) = 0 _
        Then
        CloseHandle clientInput
        SessionCloseHandles
        RaiseProcessError "The output pipe could not be created."
    End If
    If CreatePipe(mSessionStdErr, errWrite, security, SESSION_PIPE_BYTES) = 0 _
        Then
        CloseHandle clientInput
        CloseHandle outWrite
        SessionCloseHandles
        RaiseProcessError "The error pipe could not be created."
    End If
    SetHandleInformation mSessionStdOut, HANDLE_FLAG_INHERIT, 0
    SetHandleInformation mSessionStdErr, HANDLE_FLAG_INHERIT, 0

    startupInformation.cb = LenB(startupInformation)
    startupInformation.dwFlags = STARTF_USESTDHANDLES
    startupInformation.hStdInput = clientInput
    startupInformation.hStdOutput = outWrite
    startupInformation.hStdError = errWrite

    ' cmd.exe /q /k keeps the interpreter alive between writes, which is what
    ' makes the session a conversation instead of a single command.
    commandLine = "cmd.exe /q /k " & command
    If Len(workingDirectory) > 0 Then directoryPointer = StrPtr(workingDirectory)
    If CreateProcessW(0, StrPtr(commandLine), 0, 0, 1, CREATE_NO_WINDOW, _
        0, directoryPointer, startupInformation, processInformation) = 0 Then
        CloseHandle clientInput
        CloseHandle outWrite
        CloseHandle errWrite
        SessionCloseHandles
        RaiseProcessError "The process could not start: " & command
    End If

    mSessionProcess = processInformation.hProcess
    mSessionProcessId = processInformation.dwProcessId
    CloseHandle processInformation.hThread
    ' The parent must drop the child's ends, or the pipes never report end of
    ' file once the command exits.
    CloseHandle clientInput
    CloseHandle outWrite
    CloseHandle errWrite
End Sub

Private Sub RequireProcessSession(ByVal memberName As String)
    If mRole <> ROLE_PROCESS_SESSION Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            memberName & " requires a process session."
    End If
End Sub

Private Sub RaiseProcessError(ByVal message As String)
    Err.Raise ERROR_PROCESS, "ROneCOne.ProcessException", message
End Sub

Private Function SessionCreateTask( _
    ByVal operationName As String, _
    ByVal cancellationToken As ROneCOne, _
    ByVal memberName As String, _
    Optional ByVal timeoutMilliseconds As Long = -1 _
) As ROneCOne
    Dim taskValue As ROneCOne

    RequireProcessSession memberName
    ValidateCancellationToken cancellationToken, memberName
    If operationName = "ReadOut" Then
        If mSessionReadOutBusy Then
            RaiseProcessError "One standard output read may be pending " & _
                "at a time; await the previous ReadLineAsync."
        End If
        mSessionReadOutBusy = True
    ElseIf operationName = "ReadErr" Then
        If mSessionReadErrBusy Then
            RaiseProcessError "One standard error read may be pending " & _
                "at a time; await the previous ReadErrorLineAsync."
        End If
        mSessionReadErrBusy = True
    End If
    Set taskValue = New ROneCOne
    taskValue.ConfigureSessionTask Me, operationName, 0, cancellationToken
    If timeoutMilliseconds >= 0 Then
        taskValue.InternalSetSessionDeadline _
            MonotonicMilliseconds + CLngLng(timeoutMilliseconds)
    End If
    Set SessionCreateTask = taskValue
End Function

Friend Sub InternalSetSessionDeadline(ByVal deadline As LongLong)
    mTaskDeadline = deadline
End Sub

Private Function SessionQueueWrite( _
    ByVal text As String, _
    ByVal cancellationToken As ROneCOne, _
    ByVal memberName As String _
) As ROneCOne
    Dim payload() As Byte
    Dim taskValue As ROneCOne

    RequireProcessSession memberName
    ValidateCancellationToken cancellationToken, memberName
    If mSessionInputClosed Then
        RaiseProcessError "Standard input is closed for this session."
    End If
    ' Zero skip bytes: a console reads raw bytes, so no byte-order mark.
    payload = IoEncodeText(text, mSessionCharset, 0)
    mSessionWriteNext = mSessionWriteNext + 1
    mSessionWriteQueue.Add payload
    Set taskValue = New ROneCOne
    taskValue.ConfigureSessionTask Me, "Write", mSessionWriteNext, _
        cancellationToken
    Set SessionQueueWrite = taskValue
End Function

' Returns True when the write identified by targetSequence has landed in the
' pipe. Only one overlapped write is in flight at a time; the queue drains in
' order, so a completed counter is enough to resolve every waiting Task.
Friend Function InternalSessionWritePoll( _
    ByVal targetSequence As Long _
) As Boolean
    Dim transferred As Long

    If mSessionWriteActive Then
        If GetOverlappedResult(mSessionStdIn, mSessionWriteOverlapped, _
            transferred, 0) = 0 Then
            If Err.LastDllError = ERROR_IO_INCOMPLETE Then
                InternalSessionWritePoll = (mSessionWriteDone >= targetSequence)
                Exit Function
            End If
            ' The pipe broke, which means the command stopped reading. Treat
            ' the write as finished rather than hanging the Task forever.
        End If
        mSessionWriteActive = False
        If mSessionWriteOverlapped.hEvent <> 0 Then
            CloseHandle mSessionWriteOverlapped.hEvent
            mSessionWriteOverlapped.hEvent = 0
        End If
        mSessionWriteDone = mSessionWriteDone + 1
    End If

    Do While Not mSessionWriteActive And mSessionWriteQueue.Count > 0
        SessionBeginNextWrite
        If mSessionWriteActive Then
            If GetOverlappedResult(mSessionStdIn, mSessionWriteOverlapped, _
                transferred, 0) <> 0 Then
                mSessionWriteActive = False
                If mSessionWriteOverlapped.hEvent <> 0 Then
                    CloseHandle mSessionWriteOverlapped.hEvent
                    mSessionWriteOverlapped.hEvent = 0
                End If
                mSessionWriteDone = mSessionWriteDone + 1
            End If
        End If
    Loop
    InternalSessionWritePoll = (mSessionWriteDone >= targetSequence)
End Function

Private Sub SessionBeginNextWrite()
    Dim payload() As Byte

    payload = mSessionWriteQueue.Item(1)
    mSessionWriteQueue.Remove 1
    mSessionWriteBuffer = payload
    mSessionWriteOverlapped.Internal = 0
    mSessionWriteOverlapped.InternalHigh = 0
    mSessionWriteOverlapped.OffsetLow = 0
    mSessionWriteOverlapped.OffsetHigh = 0
    mSessionWriteOverlapped.hEvent = CreateEventW(0, 1, 0, 0)
    If ByteArrayLength(mSessionWriteBuffer) = 0 Then
        mSessionWriteDone = mSessionWriteDone + 1
        If mSessionWriteOverlapped.hEvent <> 0 Then
            CloseHandle mSessionWriteOverlapped.hEvent
            mSessionWriteOverlapped.hEvent = 0
        End If
        Exit Sub
    End If
    WriteFileOverlapped mSessionStdIn, VarPtr(mSessionWriteBuffer(0)), _
        ByteArrayLength(mSessionWriteBuffer), 0, mSessionWriteOverlapped
    mSessionWriteActive = True
End Sub

Private Sub SessionCloseInputHandle()
    If mSessionStdIn <> 0 Then
        CloseHandle mSessionStdIn
        mSessionStdIn = 0
    End If
    mSessionInputClosed = True
End Sub

Private Sub SessionRefreshExit()
    Dim code As Long

    If mSessionExited Or mSessionProcess = 0 Then Exit Sub
    If GetExitCodeProcess(mSessionProcess, code) = 0 Then Exit Sub
    If code = STILL_ACTIVE Then Exit Sub
    mSessionExitCode = code
    mSessionExited = True
End Sub

' Drains whatever the pipe holds right now. PeekNamedPipe reports the byte
' count first, so ReadFile only ever asks for bytes that are already there
' and never blocks. A broken pipe means the writer is gone for good.
Private Sub SessionDrain(ByVal useError As Boolean)
    Dim available As Long
    Dim buffer() As Byte
    Dim bytesRead As Long
    Dim handleValue As LongPtr

    If useError Then
        handleValue = mSessionStdErr
        If mSessionErrDrained Then Exit Sub
    Else
        handleValue = mSessionStdOut
        If mSessionOutDrained Then Exit Sub
    End If
    If handleValue = 0 Then Exit Sub

    Do
        available = 0
        If PeekNamedPipe(handleValue, 0, 0, 0, available, 0) = 0 Then
            If useError Then
                mSessionErrDrained = True
            Else
                mSessionOutDrained = True
            End If
            Exit Sub
        End If
        If available <= 0 Then Exit Do
        ReDim buffer(0 To available - 1)
        If ReadFile(handleValue, VarPtr(buffer(0)), available, _
            bytesRead, 0) = 0 Then
            If useError Then
                mSessionErrDrained = True
            Else
                mSessionOutDrained = True
            End If
            Exit Sub
        End If
        If bytesRead <= 0 Then Exit Do
        If useError Then
            SessionAppendBytes mSessionErrBytes, mSessionErrCount, _
                buffer, bytesRead
        Else
            SessionAppendBytes mSessionOutBytes, mSessionOutCount, _
                buffer, bytesRead
        End If
    Loop
End Sub

Private Sub SessionAppendBytes( _
    ByRef target() As Byte, _
    ByRef targetCount As Long, _
    ByRef source() As Byte, _
    ByVal sourceCount As Long _
)
    Dim capacity As Long
    Dim index As Long

    On Error Resume Next
    capacity = UBound(target) + 1
    If Err.Number <> 0 Then
        Err.Clear
        capacity = 0
    End If
    On Error GoTo 0
    If targetCount + sourceCount > capacity Then
        capacity = capacity * 2
        If capacity < targetCount + sourceCount Then
            capacity = targetCount + sourceCount
        End If
        If capacity < 4096 Then capacity = 4096
        ReDim Preserve target(0 To capacity - 1)
    End If
    For index = 0 To sourceCount - 1
        target(targetCount + index) = source(index)
    Next index
    targetCount = targetCount + sourceCount
End Sub

Private Function SessionText(ByVal useError As Boolean) As String
    Dim bytes() As Byte
    Dim count As Long
    Dim index As Long

    If useError Then
        count = mSessionErrCount
    Else
        count = mSessionOutCount
    End If
    If count = 0 Then Exit Function
    ReDim bytes(0 To count - 1)
    For index = 0 To count - 1
        If useError Then
            bytes(index) = mSessionErrBytes(index)
        Else
            bytes(index) = mSessionOutBytes(index)
        End If
    Next index
    SessionText = IoDecodeText(bytes, mSessionCharset)
End Function

' Resolves the next line, or reports that the stream is finished. Decoding
' the whole accumulated buffer and tracking a character cursor keeps this
' correct for multibyte encodings, where a chunk can split one character.
Friend Function InternalSessionReadPoll( _
    ByVal useError As Boolean, _
    ByRef lineText As String, _
    ByRef finished As Boolean _
) As Boolean
    Dim breakIndex As Long
    Dim cursor As Long
    Dim fullText As String
    Dim remaining As String

    SessionDrain useError
    SessionRefreshExit
    fullText = SessionText(useError)
    If useError Then
        cursor = mSessionErrCursor
    Else
        cursor = mSessionOutCursor
    End If
    remaining = Mid$(fullText, cursor + 1)
    breakIndex = InStr(1, remaining, vbLf)
    If breakIndex > 0 Then
        lineText = Left$(remaining, breakIndex - 1)
        If Len(lineText) > 0 Then
            If Right$(lineText, 1) = vbCr Then
                lineText = Left$(lineText, Len(lineText) - 1)
            End If
        End If
        If useError Then
            mSessionErrCursor = cursor + breakIndex
        Else
            mSessionOutCursor = cursor + breakIndex
        End If
        InternalSessionReadPoll = True
        Exit Function
    End If

    If Not mSessionExited Then Exit Function
    ' After exit, one more drain settles anything still buffered.
    SessionDrain useError
    fullText = SessionText(useError)
    remaining = Mid$(fullText, cursor + 1)
    breakIndex = InStr(1, remaining, vbLf)
    If breakIndex > 0 Then
        lineText = Left$(remaining, breakIndex - 1)
        If Len(lineText) > 0 Then
            If Right$(lineText, 1) = vbCr Then
                lineText = Left$(lineText, Len(lineText) - 1)
            End If
        End If
        If useError Then
            mSessionErrCursor = cursor + breakIndex
        Else
            mSessionOutCursor = cursor + breakIndex
        End If
        InternalSessionReadPoll = True
        Exit Function
    End If
    If Len(remaining) > 0 Then
        lineText = remaining
        If useError Then
            mSessionErrCursor = cursor + Len(remaining)
        Else
            mSessionOutCursor = cursor + Len(remaining)
        End If
        InternalSessionReadPoll = True
        Exit Function
    End If
    finished = True
    InternalSessionReadPoll = True
End Function

Friend Function InternalSessionExitPoll() As Boolean
    SessionDrain False
    SessionDrain True
    SessionRefreshExit
    InternalSessionExitPoll = mSessionExited
End Function

Friend Function InternalSessionExitCode() As Long
    InternalSessionExitCode = mSessionExitCode
End Function

Friend Function InternalSessionRemainingText( _
    ByVal useError As Boolean _
) As String
    Dim cursor As Long
    Dim fullText As String

    SessionDrain useError
    fullText = SessionText(useError)
    If useError Then
        cursor = mSessionErrCursor
        mSessionErrCursor = Len(fullText)
    Else
        cursor = mSessionOutCursor
        mSessionOutCursor = Len(fullText)
    End If
    InternalSessionRemainingText = Mid$(fullText, cursor + 1)
End Function

Friend Sub InternalSessionReleaseRead(ByVal useError As Boolean)
    If useError Then
        mSessionReadErrBusy = False
    Else
        mSessionReadOutBusy = False
    End If
End Sub

Private Sub SessionCloseHandles()
    If mSessionStdIn <> 0 Then
        CloseHandle mSessionStdIn
        mSessionStdIn = 0
    End If
    If mSessionStdOut <> 0 Then
        CloseHandle mSessionStdOut
        mSessionStdOut = 0
    End If
    If mSessionStdErr <> 0 Then
        CloseHandle mSessionStdErr
        mSessionStdErr = 0
    End If
    If mSessionWriteOverlapped.hEvent <> 0 Then
        CloseHandle mSessionWriteOverlapped.hEvent
        mSessionWriteOverlapped.hEvent = 0
    End If
    If mSessionProcess <> 0 Then
        CloseHandle mSessionProcess
        mSessionProcess = 0
    End If
End Sub

Private Sub RequireProcessFactory(ByVal memberName As String)
    If mRole <> ROLE_PROCESS_FACTORY Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            memberName & " requires the Process surface."
    End If
End Sub

Private Sub RequireRegex(ByVal memberName As String)
    If mRole <> ROLE_REGEX Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            memberName & " requires a regular expression."
    End If
End Sub

Private Sub RequireRegexMatch(ByVal memberName As String)
    If mRole <> ROLE_REGEX_MATCH Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            memberName & " requires a regular expression match."
    End If
End Sub

Private Sub RequireHashFactory(ByVal memberName As String)
    If mRole <> ROLE_HASH_FACTORY Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            memberName & " requires the Hash surface."
    End If
End Sub

Private Sub RequireConvertFactory(ByVal memberName As String)
    If mRole <> ROLE_CONVERT_FACTORY Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            memberName & " requires the Convert surface."
    End If
End Sub

Private Sub RequireDateTimeFactory(ByVal memberName As String)
    If mRole <> ROLE_DATETIME_FACTORY Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            memberName & " requires the DateTime surface."
    End If
End Sub

Private Sub RequireDateTime(ByVal memberName As String)
    If mRole <> ROLE_DATETIME Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            memberName & " requires a DateTime value."
    End If
End Sub

Private Sub RequireTimeSpanFactory(ByVal memberName As String)
    If mRole <> ROLE_TIMESPAN_FACTORY Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            memberName & " requires the TimeSpan surface."
    End If
End Sub

Private Sub RequireTimeSpan(ByVal memberName As String)
    If mRole <> ROLE_TIMESPAN Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            memberName & " requires a TimeSpan value."
    End If
End Sub

Private Sub RequireStringsFactory(ByVal memberName As String)
    If mRole <> ROLE_STRINGS_FACTORY Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            memberName & " requires the Strings surface."
    End If
End Sub

Private Sub RequireStringBuilder(ByVal memberName As String)
    If mRole <> ROLE_STRING_BUILDER Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            memberName & " requires a StringBuilder."
    End If
End Sub

Private Sub RequireGuidFactory(ByVal memberName As String)
    If mRole <> ROLE_GUID_FACTORY Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            memberName & " requires the Guid surface."
    End If
End Sub

Private Sub RequireRandomFactory(ByVal memberName As String)
    If mRole <> ROLE_RANDOM_FACTORY Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            memberName & " requires the Random surface."
    End If
End Sub

Private Sub RequireXmlFactory(ByVal memberName As String)
    If mRole <> ROLE_XML_FACTORY Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            memberName & " requires the Xml surface."
    End If
End Sub

Private Sub RequireXmlNode(ByVal memberName As String)
    If mRole <> ROLE_XML_NODE Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            memberName & " requires an XML node."
    End If
End Sub

Private Sub RequireUriFactory(ByVal memberName As String)
    If mRole <> ROLE_URI_FACTORY Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            memberName & " requires the Uri surface."
    End If
End Sub

Private Sub RequireLogger(ByVal memberName As String)
    If mRole <> ROLE_LOGGER Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            memberName & " requires a Logger."
    End If
End Sub

Private Sub RequireFileWatcher(ByVal memberName As String)
    If mRole <> ROLE_FILE_WATCHER Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            memberName & " requires a FileWatcher."
    End If
End Sub

Private Sub RequireZipFileFactory(ByVal memberName As String)
    If mRole <> ROLE_ZIPFILE_FACTORY Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            memberName & " requires the ZipFile surface."
    End If
End Sub

Private Sub RequireZipArchive(ByVal memberName As String)
    If mRole <> ROLE_ZIP_ARCHIVE Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            memberName & " requires an open zip archive."
    End If
End Sub

Private Sub RequireZipEntry(ByVal memberName As String)
    If mRole <> ROLE_ZIP_ENTRY Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            memberName & " requires a zip entry."
    End If
End Sub

Private Sub RequireWebUtilityFactory(ByVal memberName As String)
    If mRole <> ROLE_WEB_UTILITY_FACTORY Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            memberName & " requires the WebUtility surface."
    End If
End Sub

Private Sub RequireNoNamespaces( _
    ByVal selectionNamespaces As String, _
    ByVal memberName As String _
)
    If Len(selectionNamespaces) > 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
            "Selection namespaces apply only to the Xml surface."
    End If
End Sub

Private Function BuildRegexMatch(ByVal rawMatch As Object) As ROneCOne
    Dim groupIndex As Long
    Dim groups As ROneCOne
    Dim matchValue As ROneCOne
    Dim subValue As Variant

    ' Group 0 is the whole match, then each capture in order, matching the
    ' Groups shape C# programmers expect; an unmatched group is empty text.
    Set groups = ROneCOne.ListOf(vbString)
    groups.Add CStr(rawMatch.Value)
    For groupIndex = 0 To rawMatch.SubMatches.Count - 1
        subValue = rawMatch.SubMatches(groupIndex)
        If IsEmpty(subValue) Then
            groups.Add vbNullString
        Else
            groups.Add CStr(subValue)
        End If
    Next groupIndex
    Set matchValue = New ROneCOne
    matchValue.ConfigureRegexMatch True, CStr(rawMatch.Value), _
        CLng(rawMatch.FirstIndex), CLng(rawMatch.Length), groups
    Set BuildRegexMatch = matchValue
End Function

Private Function BuildFailedRegexMatch() As ROneCOne
    Dim matchValue As ROneCOne

    Set matchValue = New ROneCOne
    matchValue.ConfigureRegexMatch False, vbNullString, -1, 0, _
        ROneCOne.ListOf(vbString)
    Set BuildFailedRegexMatch = matchValue
End Function

Private Function HashValueCore( _
    ByVal algorithmId As String, _
    ByVal value As Variant, _
    ByVal memberName As String _
) As Variant
    Dim dataBytes() As Byte
    Dim dataLength As Long
    Dim keyBytes() As Byte

    CoerceHashInput value, dataBytes, dataLength, memberName
    HashValueCore = BcryptHashBytes(algorithmId, False, keyBytes, 0, _
        dataBytes, dataLength, memberName)
End Function

Private Sub CoerceHashInput( _
    ByVal value As Variant, _
    ByRef bytes() As Byte, _
    ByRef byteLength As Long, _
    ByVal memberName As String _
)
    ' Text hashes as its UTF-8 bytes, matching what every other platform
    ' produces for the same string; byte arrays pass through untouched.
    If VarTypeOf(value) = vbString Then
        bytes = IoEncodeText(CStr(value), "utf-8", 3)
    ElseIf VarTypeOf(value) = (vbArray Or vbByte) Then
        bytes = value
    Else
        RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
            memberName & " requires a String or Byte array."
    End If
    If UBound(bytes) < LBound(bytes) Then
        byteLength = 0
    Else
        byteLength = UBound(bytes) - LBound(bytes) + 1
    End If
End Sub

Private Function BcryptHashBytes( _
    ByVal algorithmId As String, _
    ByVal useHmac As Boolean, _
    ByRef keyBytes() As Byte, _
    ByVal keyLength As Long, _
    ByRef dataBytes() As Byte, _
    ByVal dataLength As Long, _
    ByVal memberName As String _
) As Variant
    Dim algorithmHandle As LongPtr
    Dim digest() As Byte
    Dim digestLength As Long
    Dim flags As Long
    Dim hashHandle As LongPtr
    Dim keyPointer As LongPtr
    Dim resultLength As Long
    Dim status As Long

    If useHmac Then flags = BCRYPT_HMAC_FLAG
    status = BCryptOpenAlgorithmProvider(algorithmHandle, _
        StrPtr(algorithmId), 0, flags)
    If status = 0 Then
        status = BCryptGetProperty(algorithmHandle, _
            StrPtr("HashDigestLength"), digestLength, 4, resultLength, 0)
    End If
    If status = 0 Then
        If useHmac And keyLength > 0 Then
            keyPointer = VarPtr(keyBytes(LBound(keyBytes)))
        End If
        status = BCryptCreateHash(algorithmHandle, hashHandle, 0, 0, _
            keyPointer, keyLength, 0)
    End If
    If status = 0 Then
        If dataLength > 0 Then
            status = BCryptHashData(hashHandle, _
                VarPtr(dataBytes(LBound(dataBytes))), dataLength, 0)
        Else
            status = BCryptHashData(hashHandle, 0, 0, 0)
        End If
    End If
    If status = 0 Then
        ReDim digest(0 To digestLength - 1)
        status = BCryptFinishHash(hashHandle, VarPtr(digest(0)), _
            digestLength, 0)
    End If
    If hashHandle <> 0 Then BCryptDestroyHash hashHandle
    If algorithmHandle <> 0 Then
        BCryptCloseAlgorithmProvider algorithmHandle, 0
    End If
    If status <> 0 Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            "The Windows cryptography provider failed (status &H" & _
            Hex$(status) & ")."
    End If
    BcryptHashBytes = digest
End Function

Private Function Base64EncodeBytes(ByVal bytes As Variant) As String
    Dim alphabet As String
    Dim block As Long
    Dim blockLength As Long
    Dim byteLength As Long
    Dim chunkIndex As Long
    Dim cursor As Long
    Dim lowerBound As Long
    Dim parts() As String

    alphabet = _
        "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
    lowerBound = LBound(bytes)
    If UBound(bytes) < lowerBound Then Exit Function
    byteLength = UBound(bytes) - lowerBound + 1
    ReDim parts(0 To (byteLength + 2) \ 3 - 1)
    For cursor = 0 To byteLength - 1 Step 3
        blockLength = byteLength - cursor
        If blockLength > 3 Then blockLength = 3
        block = CLng(bytes(lowerBound + cursor)) * 65536
        If blockLength >= 2 Then
            block = block + CLng(bytes(lowerBound + cursor + 1)) * 256
        End If
        If blockLength = 3 Then
            block = block + CLng(bytes(lowerBound + cursor + 2))
        End If
        parts(chunkIndex) = _
            Mid$(alphabet, (block \ 262144) + 1, 1) & _
            Mid$(alphabet, ((block \ 4096) And 63) + 1, 1) & _
            IIf(blockLength >= 2, _
                Mid$(alphabet, ((block \ 64) And 63) + 1, 1), "=") & _
            IIf(blockLength = 3, _
                Mid$(alphabet, (block And 63) + 1, 1), "=")
        chunkIndex = chunkIndex + 1
    Next cursor
    Base64EncodeBytes = VBA.Join(parts, vbNullString)
End Function

Private Function Base64DecodeText(ByVal text As String) As Variant
    Dim accumulated As Long
    Dim bytes() As Byte
    Dim character As String
    Dim charIndex As Long
    Dim cleaned As String
    Dim digit As Long
    Dim padCount As Long
    Dim pendingBits As Long
    Dim scaleValue As Long
    Dim writeIndex As Long

    ' Whitespace is tolerated the way System.Convert tolerates it; anything
    ' else outside the alphabet, or misplaced padding, is an error.
    For charIndex = 1 To Len(text)
        character = Mid$(text, charIndex, 1)
        Select Case character
            Case " ", vbTab, vbCr, vbLf
                ' Skip formatting whitespace.
            Case Else
                cleaned = cleaned & character
        End Select
    Next charIndex
    If Len(cleaned) = 0 Then
        Base64DecodeText = EmptyByteArray()
        Exit Function
    End If
    If (Len(cleaned) Mod 4) <> 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "FromBase64String", _
            "Base64 text length must be a multiple of four."
    End If
    ReDim bytes(0 To (Len(cleaned) \ 4) * 3 - 1)
    For charIndex = 1 To Len(cleaned)
        character = Mid$(cleaned, charIndex, 1)
        If character = "=" Then
            If charIndex <= Len(cleaned) - 2 Then
                RaiseContractError ERROR_INVALID_ARGUMENT, _
                    "FromBase64String", "Base64 padding is misplaced."
            End If
            padCount = padCount + 1
        Else
            If padCount > 0 Then
                RaiseContractError ERROR_INVALID_ARGUMENT, _
                    "FromBase64String", "Base64 padding is misplaced."
            End If
            digit = Base64DigitValue(character)
            accumulated = accumulated * 64 + digit
            pendingBits = pendingBits + 6
            If pendingBits >= 8 Then
                pendingBits = pendingBits - 8
                scaleValue = CLng(2 ^ pendingBits)
                bytes(writeIndex) = accumulated \ scaleValue
                accumulated = accumulated Mod scaleValue
                writeIndex = writeIndex + 1
            End If
        End If
    Next charIndex
    If writeIndex = 0 Then
        Base64DecodeText = EmptyByteArray()
        Exit Function
    End If
    ReDim Preserve bytes(0 To writeIndex - 1)
    Base64DecodeText = bytes
End Function

Private Function Base64DigitValue(ByVal character As String) As Long
    Dim code As Long

    code = AscW(character)
    Select Case code
        Case 65 To 90
            Base64DigitValue = code - 65
        Case 97 To 122
            Base64DigitValue = code - 71
        Case 48 To 57
            Base64DigitValue = code + 4
        Case 43
            Base64DigitValue = 62
        Case 47
            Base64DigitValue = 63
        Case Else
            RaiseContractError ERROR_INVALID_ARGUMENT, "FromBase64String", _
                "The character '" & character & "' is not base64."
    End Select
End Function

Private Function HexDigitValue(ByVal character As String) As Long
    Dim code As Long

    code = AscW(character)
    Select Case code
        Case 48 To 57
            HexDigitValue = code - 48
        Case 65 To 70
            HexDigitValue = code - 55
        Case 97 To 102
            HexDigitValue = code - 87
        Case Else
            RaiseContractError ERROR_INVALID_ARGUMENT, "FromHexString", _
                "The character '" & character & "' is not hexadecimal."
    End Select
End Function

Private Sub RequireProcessResult(ByVal memberName As String)
    If mRole <> ROLE_PROCESS_RESULT Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            memberName & " requires a process result."
    End If
End Sub

Private Function CsvSerializeValue(ByVal value As Variant) As String
    Dim builder As JsonTextBuilder
    Dim runtimeValue As ROneCOne

    If IsObject(value) Then
        If Not value Is Nothing Then
            If TypeOf value Is ROneCOne Then Set runtimeValue = value
        End If
    End If
    If runtimeValue Is Nothing Then
        RaiseCsvError "CSV serializes a DataTable or DataView."
    End If
    JsonBuilderInit builder, 512
    Select Case runtimeValue.InternalRole
        Case ROLE_DATA_TABLE
            CsvWriteHeader builder, runtimeValue
            CsvWriteTableRows builder, runtimeValue
        Case ROLE_DATA_VIEW
            CsvWriteHeader builder, runtimeValue.InternalSource
            CsvWriteViewRows builder, runtimeValue
        Case Else
            RaiseCsvError "CSV serializes a DataTable or DataView."
    End Select
    CsvSerializeValue = JsonBuilderText(builder)
End Function

Private Sub CsvWriteHeader( _
    ByRef builder As JsonTextBuilder, _
    ByVal table As ROneCOne _
)
    Dim columnIndex As Long
    Dim columns As Collection

    Set columns = table.InternalDataColumns
    If columns.Count = 0 Then
        RaiseCsvError "CSV requires at least one column."
    End If
    For columnIndex = 1 To columns.Count
        If columnIndex > 1 Then JsonBuilderAppend builder, ","
        JsonBuilderAppend builder, _
            CsvEscapedField(columns.Item(columnIndex).ColumnName)
    Next columnIndex
    JsonBuilderAppend builder, vbCrLf
End Sub

Private Sub CsvWriteTableRows( _
    ByRef builder As JsonTextBuilder, _
    ByVal table As ROneCOne _
)
    Dim row As ROneCOne

    For Each row In table.InternalDataRows
        If row.RowState <> "Deleted" Then CsvWriteDataRow builder, row
    Next row
End Sub

Private Sub CsvWriteViewRows( _
    ByRef builder As JsonTextBuilder, _
    ByVal view As ROneCOne _
)
    Dim wrapped As ROneCOne

    For Each wrapped In view.MaterializeWrapped
        CsvWriteDataRow builder, JsonUnwrapObject(wrapped)
    Next wrapped
End Sub

Private Sub CsvWriteDataRow( _
    ByRef builder As JsonTextBuilder, _
    ByVal row As ROneCOne _
)
    Dim cells As Collection
    Dim columnIndex As Long
    Dim wrapped As ROneCOne

    Set cells = row.InternalDataRows
    For columnIndex = 1 To cells.Count
        If columnIndex > 1 Then JsonBuilderAppend builder, ","
        Set wrapped = cells.Item(columnIndex)
        ' A database null writes an empty unquoted field; an empty string
        ' writes a quoted pair, so the two survive a round trip distinctly.
        If Not IsNullDataValue(wrapped) Then
            CsvWriteCell builder, JsonUnwrapValue(wrapped)
        End If
    Next columnIndex
    JsonBuilderAppend builder, vbCrLf
End Sub

Private Sub CsvWriteCell( _
    ByRef builder As JsonTextBuilder, _
    ByVal value As Variant _
)
    Select Case VarTypeOf(value)
        Case vbString
            JsonBuilderAppend builder, CsvEscapedField(CStr(value))
        Case vbBoolean
            If value Then
                JsonBuilderAppend builder, "true"
            Else
                JsonBuilderAppend builder, "false"
            End If
        Case vbByte, vbInteger, vbLong, vbLongLong
            JsonBuilderAppend builder, CStr(value)
        Case vbSingle, vbDouble, vbCurrency, vbDecimal
            JsonBuilderAppend builder, JsonNumberText(CDbl(value))
        Case vbDate
            JsonBuilderAppend builder, VBA.Format$(value, "yyyy-mm-dd\Thh:nn:ss")
        Case vbNull, vbEmpty
            ' An absent value writes an empty unquoted field.
        Case Else
            RaiseCsvError "CSV cells must be scalar values; '" & _
                TypeName(value) & "' is not writable."
    End Select
End Sub

Private Function CsvEscapedField(ByVal fieldText As String) As String
    Dim needsQuotes As Boolean

    If Len(fieldText) = 0 Then
        needsQuotes = True
    ElseIf InStr(1, fieldText, """", vbBinaryCompare) > 0 Then
        needsQuotes = True
    ElseIf InStr(1, fieldText, ",", vbBinaryCompare) > 0 Then
        needsQuotes = True
    ElseIf InStr(1, fieldText, vbCr, vbBinaryCompare) > 0 Then
        needsQuotes = True
    ElseIf InStr(1, fieldText, vbLf, vbBinaryCompare) > 0 Then
        needsQuotes = True
    ElseIf Left$(fieldText, 1) = " " Or Right$(fieldText, 1) = " " Then
        needsQuotes = True
    End If
    If needsQuotes Then
        CsvEscapedField = """" & _
            VBA.Replace$(fieldText, """", """""") & """"
    Else
        CsvEscapedField = fieldText
    End If
End Function

Private Function CsvParseTable( _
    ByVal text As String, _
    ByVal tableName As String _
) As ROneCOne
    Dim columnCount As Long
    Dim columnIndex As Long
    Dim columnTypes() As Long
    Dim dateValue As Date
    Dim fieldsValue As Variant
    Dim headerFields As Variant
    Dim nameKeys As Collection
    Dim quotedRecords As Collection
    Dim quotedValue As Variant
    Dim records As Collection
    Dim recordIndex As Long
    Dim rowValues() As Variant
    Dim table As ROneCOne

    Set records = New Collection
    Set quotedRecords = New Collection
    CsvSplitRecords text, records, quotedRecords
    If records.Count = 0 Then RaiseCsvError "CSV text is empty."
    headerFields = records.Item(1)
    columnCount = UBound(headerFields) + 1
    Set nameKeys = New Collection
    For columnIndex = 0 To columnCount - 1
        If Len(Trim$(CStr(headerFields(columnIndex)))) = 0 Then
            RaiseCsvError "Column " & CStr(columnIndex + 1) & _
                " has an empty name."
        End If
        On Error Resume Next
        nameKeys.Add columnIndex, "k" & UCase$(CStr(headerFields(columnIndex)))
        If Err.Number <> 0 Then
            On Error GoTo 0
            RaiseCsvError "Duplicate column name '" & _
                CStr(headerFields(columnIndex)) & "'."
        End If
        On Error GoTo 0
    Next columnIndex

    ' Pass one classifies every non-null cell and unifies each column to one
    ' deterministic type; a column that mixes shapes falls back to text with
    ' its original characters. Pass two loads typed rows.
    ReDim columnTypes(0 To columnCount - 1)
    For recordIndex = 2 To records.Count
        fieldsValue = records.Item(recordIndex)
        If UBound(fieldsValue) + 1 <> columnCount Then
            RaiseCsvError "Row " & CStr(recordIndex - 1) & " has " & _
                CStr(UBound(fieldsValue) + 1) & " fields; the header has " & _
                CStr(columnCount) & "."
        End If
        quotedValue = quotedRecords.Item(recordIndex)
        For columnIndex = 0 To columnCount - 1
            If quotedValue(columnIndex) Or _
                Len(fieldsValue(columnIndex)) > 0 Then
                MergeCsvColumnType columnTypes(columnIndex), _
                    CsvClassifyCell(CStr(fieldsValue(columnIndex)), _
                        CBool(quotedValue(columnIndex)))
            End If
        Next columnIndex
    Next recordIndex

    Set table = ROneCOne.DataTable(tableName)
    For columnIndex = 0 To columnCount - 1
        If columnTypes(columnIndex) = 0 Then
            columnTypes(columnIndex) = vbString
        End If
        table.Column CStr(headerFields(columnIndex)), _
            columnTypes(columnIndex)
    Next columnIndex
    For recordIndex = 2 To records.Count
        fieldsValue = records.Item(recordIndex)
        quotedValue = quotedRecords.Item(recordIndex)
        ReDim rowValues(0 To columnCount - 1)
        For columnIndex = 0 To columnCount - 1
            If Not quotedValue(columnIndex) And _
                Len(fieldsValue(columnIndex)) = 0 Then
                rowValues(columnIndex) = Null
            ElseIf columnTypes(columnIndex) = vbString Then
                rowValues(columnIndex) = CStr(fieldsValue(columnIndex))
            ElseIf columnTypes(columnIndex) = vbDate Then
                If CsvTryIsoDate(CStr(fieldsValue(columnIndex)), _
                    dateValue) Then
                    rowValues(columnIndex) = dateValue
                End If
            Else
                rowValues(columnIndex) = CsvConvertCell( _
                    CStr(fieldsValue(columnIndex)), columnTypes(columnIndex))
            End If
        Next columnIndex
        table.LoadRow rowValues
    Next recordIndex
    Set CsvParseTable = table
End Function

Private Sub CsvSplitRecords( _
    ByVal text As String, _
    ByVal records As Collection, _
    ByVal quotedRecords As Collection _
)
    Dim fieldCount As Long
    Dim fields() As String
    Dim position As Long
    Dim quoted() As Boolean
    Dim textLength As Long

    textLength = Len(text)
    position = 1
    Do While position <= textLength
        fieldCount = 0
        ReDim fields(0 To 7)
        ReDim quoted(0 To 7)
        Do
            If fieldCount > UBound(fields) Then
                ReDim Preserve fields(0 To UBound(fields) * 2 + 1)
                ReDim Preserve quoted(0 To UBound(quoted) * 2 + 1)
            End If
            If position <= textLength Then
                If Mid$(text, position, 1) = """" Then
                    fields(fieldCount) = CsvReadQuotedField(text, position)
                    quoted(fieldCount) = True
                Else
                    fields(fieldCount) = CsvReadPlainField(text, position)
                End If
            End If
            fieldCount = fieldCount + 1
            If position > textLength Then Exit Do
            Select Case Mid$(text, position, 1)
                Case ","
                    position = position + 1
                Case vbCr
                    position = position + 1
                    If position <= textLength Then
                        If Mid$(text, position, 1) = vbLf Then
                            position = position + 1
                        End If
                    End If
                    Exit Do
                Case vbLf
                    position = position + 1
                    Exit Do
            End Select
        Loop
        ReDim Preserve fields(0 To fieldCount - 1)
        ReDim Preserve quoted(0 To fieldCount - 1)
        records.Add fields
        quotedRecords.Add quoted
    Loop
End Sub

Private Function CsvReadQuotedField( _
    ByVal text As String, _
    ByRef position As Long _
) As String
    Dim closeAt As Long
    Dim contentStart As Long
    Dim nextChar As String
    Dim textLength As Long

    textLength = Len(text)
    contentStart = position + 1
    closeAt = contentStart
    Do
        closeAt = InStr(closeAt, text, """", vbBinaryCompare)
        If closeAt = 0 Then
            RaiseCsvError "Unterminated quoted field starting at position " & _
                CStr(position) & "."
        End If
        If closeAt < textLength Then
            If Mid$(text, closeAt + 1, 1) = """" Then
                closeAt = closeAt + 2
            Else
                Exit Do
            End If
        Else
            Exit Do
        End If
    Loop
    CsvReadQuotedField = VBA.Replace$( _
        Mid$(text, contentStart, closeAt - contentStart), """""", """")
    position = closeAt + 1
    If position <= textLength Then
        nextChar = Mid$(text, position, 1)
        If nextChar <> "," And nextChar <> vbCr And nextChar <> vbLf Then
            RaiseCsvError "Unexpected character after a closing quote " & _
                "at position " & CStr(position) & "."
        End If
    End If
End Function

Private Function CsvReadPlainField( _
    ByVal text As String, _
    ByRef position As Long _
) As String
    Dim delimiterAt As Long
    Dim endAt As Long
    Dim textLength As Long

    textLength = Len(text)
    endAt = textLength + 1
    delimiterAt = InStr(position, text, ",", vbBinaryCompare)
    If delimiterAt > 0 And delimiterAt < endAt Then endAt = delimiterAt
    delimiterAt = InStr(position, text, vbCr, vbBinaryCompare)
    If delimiterAt > 0 And delimiterAt < endAt Then endAt = delimiterAt
    delimiterAt = InStr(position, text, vbLf, vbBinaryCompare)
    If delimiterAt > 0 And delimiterAt < endAt Then endAt = delimiterAt
    CsvReadPlainField = Mid$(text, position, endAt - position)
    If InStr(1, CsvReadPlainField, """", vbBinaryCompare) > 0 Then
        RaiseCsvError "A quote is not allowed inside an unquoted field " & _
            "at position " & CStr(position) & "."
    End If
    position = endAt
End Function

Private Sub MergeCsvColumnType( _
    ByRef current As Long, _
    ByVal candidate As Long _
)
    If current = 0 Then
        current = candidate
    ElseIf current = candidate Or current = vbString Then
        ' Settled already.
    ElseIf JsonIsNumericType(current) And JsonIsNumericType(candidate) Then
        current = JsonWiderNumericType(current, candidate)
    Else
        current = vbString
    End If
End Sub

Private Function CsvClassifyCell( _
    ByVal fieldText As String, _
    ByVal quoted As Boolean _
) As Long
    Dim dateValue As Date

    ' A quoted cell is always text, which keeps deliberately quoted values
    ' such as "00042" or "true" intact through a round trip.
    If quoted Then
        CsvClassifyCell = vbString
        Exit Function
    End If
    Select Case LCase$(fieldText)
        Case "true", "false"
            CsvClassifyCell = vbBoolean
            Exit Function
    End Select
    CsvClassifyCell = CsvClassifyNumber(fieldText)
    If CsvClassifyCell <> 0 Then Exit Function
    If CsvTryIsoDate(fieldText, dateValue) Then
        CsvClassifyCell = vbDate
    Else
        CsvClassifyCell = vbString
    End If
End Function

Private Function CsvClassifyNumber(ByVal fieldText As String) As Long
    Dim code As Long
    Dim digitCount As Long
    Dim integral As Boolean
    Dim overflowProbe As Variant
    Dim partDigits As Long
    Dim position As Long
    Dim textLength As Long

    ' The grammar is JSON's: optional minus, no leading zeros, optional
    ' fraction and exponent, nothing trailing. Returns 0 for non-numbers.
    textLength = Len(fieldText)
    position = 1
    integral = True
    If textLength = 0 Then Exit Function
    If Mid$(fieldText, 1, 1) = "-" Then position = 2
    If position > textLength Then Exit Function
    code = AscW(Mid$(fieldText, position, 1))
    If code = 48 Then
        digitCount = 1
        position = position + 1
        If position <= textLength Then
            code = AscW(Mid$(fieldText, position, 1))
            If code >= 48 And code <= 57 Then Exit Function
        End If
    ElseIf code >= 49 And code <= 57 Then
        Do While position <= textLength
            code = AscW(Mid$(fieldText, position, 1))
            If code < 48 Or code > 57 Then Exit Do
            digitCount = digitCount + 1
            position = position + 1
        Loop
    Else
        Exit Function
    End If
    If position <= textLength Then
        If Mid$(fieldText, position, 1) = "." Then
            integral = False
            position = position + 1
            partDigits = 0
            Do While position <= textLength
                code = AscW(Mid$(fieldText, position, 1))
                If code < 48 Or code > 57 Then Exit Do
                partDigits = partDigits + 1
                position = position + 1
            Loop
            If partDigits = 0 Then Exit Function
        End If
    End If
    If position <= textLength Then
        code = AscW(Mid$(fieldText, position, 1))
        If code = 101 Or code = 69 Then
            integral = False
            position = position + 1
            If position <= textLength Then
                code = AscW(Mid$(fieldText, position, 1))
                If code = 43 Or code = 45 Then position = position + 1
            End If
            partDigits = 0
            Do While position <= textLength
                code = AscW(Mid$(fieldText, position, 1))
                If code < 48 Or code > 57 Then Exit Do
                partDigits = partDigits + 1
                position = position + 1
            Loop
            If partDigits = 0 Then Exit Function
        End If
    End If
    If position <= textLength Then Exit Function
    If Not integral Then
        CsvClassifyNumber = vbDouble
        Exit Function
    End If
    If digitCount <= 9 Then
        CsvClassifyNumber = vbLong
        Exit Function
    End If
    On Error Resume Next
    Err.Clear
    overflowProbe = CLng(fieldText)
    If Err.Number = 0 Then
        CsvClassifyNumber = vbLong
    Else
        Err.Clear
        overflowProbe = CLngLng(fieldText)
        If Err.Number = 0 Then
            CsvClassifyNumber = vbLongLong
        Else
            CsvClassifyNumber = vbDouble
        End If
    End If
    On Error GoTo 0
End Function

Private Function CsvTryIsoDate( _
    ByVal fieldText As String, _
    ByRef result As Date _
) As Boolean
    Dim dayPart As Long
    Dim hourPart As Long
    Dim index As Long
    Dim minutePart As Long
    Dim monthPart As Long
    Dim probe As String
    Dim secondPart As Long
    Dim textLength As Long
    Dim yearPart As Long

    textLength = Len(fieldText)
    If textLength <> 10 And textLength <> 19 Then Exit Function
    For index = 1 To textLength
        probe = Mid$(fieldText, index, 1)
        Select Case index
            Case 5, 8
                If probe <> "-" Then Exit Function
            Case 11
                If probe <> "T" Then Exit Function
            Case 14, 17
                If probe <> ":" Then Exit Function
            Case Else
                If probe < "0" Or probe > "9" Then Exit Function
        End Select
    Next index
    yearPart = CLng(Mid$(fieldText, 1, 4))
    monthPart = CLng(Mid$(fieldText, 6, 2))
    dayPart = CLng(Mid$(fieldText, 9, 2))
    If monthPart < 1 Or monthPart > 12 Then Exit Function
    If dayPart < 1 Or dayPart > 31 Then Exit Function
    result = DateSerial(yearPart, monthPart, dayPart)
    If VBA.Year(result) <> yearPart Or VBA.Month(result) <> monthPart Or _
        VBA.Day(result) <> dayPart Then Exit Function
    If textLength = 19 Then
        hourPart = CLng(Mid$(fieldText, 12, 2))
        minutePart = CLng(Mid$(fieldText, 15, 2))
        secondPart = CLng(Mid$(fieldText, 18, 2))
        If hourPart > 23 Or minutePart > 59 Or secondPart > 59 Then
            Exit Function
        End If
        result = result + TimeSerial(hourPart, minutePart, secondPart)
    End If
    CsvTryIsoDate = True
End Function

Private Function CsvConvertCell( _
    ByVal fieldText As String, _
    ByVal columnType As Long _
) As Variant
    Select Case columnType
        Case vbLong
            CsvConvertCell = CLng(fieldText)
        Case vbLongLong
            CsvConvertCell = CLngLng(fieldText)
        Case vbDouble
            ' Val always parses the invariant dot separator.
            CsvConvertCell = Val(fieldText)
        Case vbBoolean
            CsvConvertCell = (LCase$(fieldText) = "true")
    End Select
End Function

Private Sub RequireHttpClient(ByVal memberName As String)
    If mRole <> ROLE_HTTP_CLIENT Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            memberName & " requires an HttpClient."
    End If
End Sub

Private Sub RequireHttpResponse(ByVal memberName As String)
    If mRole <> ROLE_HTTP_RESPONSE Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            memberName & " requires an HttpResponse."
    End If
End Sub

Private Function StartHttpSend( _
    ByVal httpMethod As String, _
    ByVal url As String, _
    ByVal body As Variant, _
    ByVal hasBody As Boolean, _
    ByVal contentType As String, _
    ByVal resultMode As String, _
    ByVal memberName As String, _
    ByVal cancellationToken As ROneCOne _
) As ROneCOne
    Dim index As Long
    Dim request As Object
    Dim response As ROneCOne
    Dim taskValue As ROneCOne

    ValidateCancellationToken cancellationToken, memberName
    ' The request is opened asynchronously and sent immediately, so the
    ' transfer proceeds inside WinHTTP while the returned Task is awaited
    ' cooperatively on Excel's thread.
    Set request = CreateObject(HTTP_PROG_ID)
    request.SetTimeouts mHttpTimeoutMs, mHttpTimeoutMs, _
        mHttpTimeoutMs, mHttpTimeoutMs
    request.Open httpMethod, ResolveHttpUrl(url, memberName), True
    For index = 1 To mHttpHeaderNames.Count
        request.SetRequestHeader CStr(mHttpHeaderNames.Item(index)), _
            CStr(mHttpHeaderValues.Item(index))
    Next index
    If Len(contentType) > 0 Then
        request.SetRequestHeader "Content-Type", contentType
    End If
    If hasBody Then
        request.Send CStr(body)
    Else
        request.Send
    End If
    Set response = New ROneCOne
    response.ConfigureHttpResponse request
    Set taskValue = New ROneCOne
    taskValue.ConfigureHttpTask response, resultMode, cancellationToken
    Set StartHttpSend = taskValue
End Function

Private Function ResolveHttpUrl( _
    ByVal url As String, _
    ByVal memberName As String _
) As String
    If InStr(1, url, "://", vbTextCompare) > 0 Then
        ResolveHttpUrl = url
    ElseIf Len(mHttpBaseAddress) > 0 Then
        ResolveHttpUrl = mHttpBaseAddress & url
    Else
        RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
            "A relative URL requires the client's BaseAddress."
    End If
End Function

Private Function HttpStatusFailureMessage( _
    ByVal response As ROneCOne _
) As String
    HttpStatusFailureMessage = "The HTTP request failed with status " & _
        CStr(response.StatusCode) & " (" & response.ReasonPhrase & ")."
End Function

Private Function NormalizeHttpBodyBytes(ByVal value As Variant) As Variant
    If VarTypeOf(value) = (vbArray Or vbByte) Then
        NormalizeHttpBodyBytes = value
    Else
        NormalizeHttpBodyBytes = EmptyByteArray()
    End If
End Function

Private Function CreateHttpStatusException( _
    ByVal response As ROneCOne _
) As ROneCOne
    Dim errorValue As ROneCOne

    Set errorValue = New ROneCOne
    errorValue.ConfigureError ERROR_HTTP_REQUEST, _
        "ROneCOne.HttpRequestException", _
        HttpStatusFailureMessage(response), vbNullString, 0
    errorValue.ConfigureExceptionType "HttpRequestException"
    Set CreateHttpStatusException = errorValue
End Function

Private Function CreateProviderTask( _
    ByVal methodName As String, _
    Optional ByVal argument As Variant, _
    Optional ByVal cancellationToken As ROneCOne _
) As ROneCOne
    Dim arguments As Collection
    Dim taskValue As ROneCOne

    ValidateCancellationToken cancellationToken, methodName & "Async"
    Set arguments = New Collection
    If Not IsMissing(argument) Then
        arguments.Add WrapInvocationArgument(argument)
    End If
    Set taskValue = New ROneCOne
    taskValue.ConfigureProviderTask Me, methodName, arguments, cancellationToken
    Set CreateProviderTask = taskValue
End Function

Private Function CreateReadOnlyCollectionFromItems( _
    ByVal items As Collection _
) As ROneCOne
    Dim item As ROneCOne
    Dim values As ROneCOne

    Set values = ROneCOne.ListOf(Me)
    For Each item In items
        values.Add item
    Next item
    Set CreateReadOnlyCollectionFromItems = _
        ROneCOne.ReadOnlyCollectionOf(values)
End Function

Private Function DbAsyncTaskChild(ByVal index As Long) As ROneCOne
    Dim wrapped As ROneCOne

    Set wrapped = mTaskChildren.Item(index)
    Set DbAsyncTaskChild = wrapped.InternalObjectValue
End Function

Private Function RequireProviderTaskTarget( _
    ByVal arguments As Collection _
) As ROneCOne
    Dim candidate As Object
    Dim wrapped As ROneCOne

    If arguments Is Nothing Or arguments.Count <> 1 Then
        RaiseContractError ERROR_INVALID_OPERATION, "ProviderTask", _
            "The provider Task requires one data target."
    End If
    Set wrapped = arguments.Item(1)
    If Not wrapped.InternalHasObjectValue Then
        RaiseContractError ERROR_TYPE_MISMATCH, "ProviderTask", _
            "The provider Task target must be a DataTable or DataSet."
    End If
    Set candidate = wrapped.InternalObjectValue
    If candidate Is Nothing Or Not TypeOf candidate Is ROneCOne Then
        RaiseContractError ERROR_TYPE_MISMATCH, "ProviderTask", _
            "The provider Task target must be a DataTable or DataSet."
    End If
    Set RequireProviderTaskTarget = candidate
End Function

Private Function BuildAdoCommand( _
    Optional ByVal dataRow As ROneCOne _
) As Object
    Dim commandObject As Object
    Dim connectionObject As Object
    Dim parameter As ROneCOne
    Dim parametersObject As Object
    Dim providerParameter As Object
    Dim rawValue As Variant
    Dim valueSize As Long
    Dim valueWrapped As ROneCOne

    RequireDbCommand Me, "Execute"
    mSource.Connect
    Set connectionObject = mSource.InternalProviderObject
    Set commandObject = CreateObject("ADODB.Command")
    CallByName commandObject, "ActiveConnection", VbSet, connectionObject
    CallByName commandObject, "CommandText", VbLet, mCommandText
    If mCommandTimeout > 0 Then
        CallByName commandObject, "CommandTimeout", VbLet, mCommandTimeout
    End If
    CallByName commandObject, "CommandType", VbLet, mCommandType
    Set parametersObject = CallByName(commandObject, "Parameters", VbGet)
    For Each parameter In ArrSnapshot(mItems, mItemsCount)
        If dataRow Is Nothing Or Len(parameter.InternalSourceColumn) = 0 Then
            Set valueWrapped = parameter.InternalQueryValue
        Else
            Set valueWrapped = dataRow.DataRowVersionValue( _
                parameter.InternalSourceColumn, parameter.InternalSourceVersion)
        End If
        rawValue = Unwrap(valueWrapped)
        If VarTypeOf(rawValue) = vbString Then valueSize = Len(CStr(rawValue))
        If parameter.InternalProviderSize > 0 Then
            valueSize = parameter.InternalProviderSize
        ElseIf valueSize < 1 Then
            valueSize = 1
        End If
        Set providerParameter = CallByName( _
            commandObject, "CreateParameter", VbMethod, _
            parameter.InternalName, AdoTypeForParameter(parameter, valueWrapped), _
            parameter.InternalExpressionOption, valueSize, rawValue)
        If parameter.InternalProviderPrecision > 0 Then
            CallByName providerParameter, "Precision", VbLet, _
                parameter.InternalProviderPrecision
        End If
        If parameter.InternalProviderScale > 0 Then
            CallByName providerParameter, "NumericScale", VbLet, _
                parameter.InternalProviderScale
        End If
        CallByName parametersObject, "Append", VbMethod, providerParameter
    Next parameter
    Set BuildAdoCommand = commandObject
End Function

Private Function AdoTypeForParameter( _
    ByVal parameter As ROneCOne, _
    ByVal wrapped As ROneCOne _
) As Long
    If parameter.InternalProviderType <> 0 Then
        AdoTypeForParameter = parameter.InternalProviderType
    Else
        AdoTypeForParameter = AdoTypeForValue(wrapped)
    End If
End Function

Private Function NormalizeDataRowVersion( _
    ByVal sourceVersion As String, _
    ByVal memberName As String _
) As String
    Select Case UCase$(Trim$(sourceVersion))
        Case "CURRENT"
            NormalizeDataRowVersion = "Current"
        Case "ORIGINAL"
            NormalizeDataRowVersion = "Original"
        Case Else
            RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
                "Source version must be Current or Original."
    End Select
End Function

Friend Function ExecuteNonQueryForDataRow( _
    ByVal dataRow As ROneCOne _
) As Long
    Dim commandObject As Object
    Dim capturedDescription As String
    Dim capturedNumber As Long
    Dim capturedSource As String
    Dim recordsAffected As Variant
    Dim recordset As Object

    RequireDbCommand Me, "Update"
    Set commandObject = BuildAdoCommand(dataRow)
    On Error GoTo ExecuteFailed
    Set recordset = commandObject.Execute(recordsAffected)
    CloseProviderObject recordset
    ExecuteNonQueryForDataRow = CLng(recordsAffected)
    On Error GoTo 0
    Exit Function

ExecuteFailed:
    capturedNumber = Err.Number
    capturedSource = Err.Source
    capturedDescription = Err.Description
    On Error Resume Next
    CloseProviderObject recordset
    On Error GoTo 0
    Err.Raise capturedNumber, capturedSource, capturedDescription
End Function

Private Function UpdateDataTable(ByVal table As ROneCOne) As Long
    Dim capturedDescription As String
    Dim capturedNumber As Long
    Dim capturedSource As String
    Dim command As ROneCOne
    Dim connection As ROneCOne
    Dim errorValue As ROneCOne
    Dim row As ROneCOne
    Dim rowState As String
    Dim transaction As ROneCOne
    Dim transactionState As String

    RequireDataTable table, "Update"
    Set mProviderErrors = New Collection
    If mAdapterUseTransaction Then
        Set connection = mSource.InternalSource
        RequireDbConnection connection, "Update transaction"
        Set transaction = connection.BeginTransaction
    End If
    On Error GoTo UpdateFailed
    For Each row In table.InternalDataRows
        rowState = row.RowState
        Set command = Nothing
        Select Case rowState
            Case "Added"
                Set command = mLeft
            Case "Modified"
                Set command = mRight
            Case "Deleted"
                Set command = mNext
        End Select
        If rowState <> "Unchanged" And rowState <> "Detached" Then
            If command Is Nothing Then
                RaiseContractError ERROR_INVALID_OPERATION, "Update", _
                    "The adapter has no command for " & rowState & " rows."
            End If
            If mAdapterContinueOnError Then
                Set errorValue = Nothing
                If ExecuteUpdateRowCaptured(command, row, errorValue) Then
                    UpdateDataTable = UpdateDataTable + 1
                Else
                    mProviderErrors.Add errorValue
                End If
            Else
                command.ExecuteNonQueryForDataRow row
                UpdateDataTable = UpdateDataTable + 1
            End If
        End If
    Next row
    If mProviderErrors.Count > 0 Then
        If Not transaction Is Nothing Then
            transaction.Rollback
            UpdateDataTable = 0
        End If
        Exit Function
    End If
    If Not transaction Is Nothing Then transaction.Commit
    If UpdateDataTable > 0 Then table.AcceptChanges
    On Error GoTo 0
    Exit Function

UpdateFailed:
    capturedNumber = Err.Number
    capturedSource = Err.Source
    capturedDescription = Err.Description
    On Error Resume Next
    If Not transaction Is Nothing Then
        transactionState = transaction.State
        If transactionState = "Open" Then transaction.Rollback
    End If
    On Error GoTo 0
    Err.Raise capturedNumber, capturedSource, capturedDescription
End Function

Private Function ExecuteUpdateRowCaptured( _
    ByVal command As ROneCOne, _
    ByVal row As ROneCOne, _
    ByRef errorValue As ROneCOne _
) As Boolean
    Dim ignored As Long

    On Error GoTo UpdateFailed
    ignored = command.ExecuteNonQueryForDataRow(row)
    ExecuteUpdateRowCaptured = True
    Exit Function

UpdateFailed:
    Set errorValue = New ROneCOne
    errorValue.ConfigureError Err.Number, Err.Source, Err.Description, _
        Err.HelpFile, Err.HelpContext
    Err.Clear
    On Error GoTo 0
End Function

Private Function AdoTypeForValue(ByVal wrapped As ROneCOne) As Long
    If wrapped.InternalHasObjectValue Then
        AdoTypeForValue = 12
        Exit Function
    End If
    Select Case VarTypeOf(wrapped.InternalScalarValue)
        Case vbByte
            AdoTypeForValue = 16
        Case vbInteger
            AdoTypeForValue = 2
        Case vbLong
            AdoTypeForValue = 3
        Case vbLongLong
            AdoTypeForValue = 20
        Case vbSingle
            AdoTypeForValue = 4
        Case vbDouble
            AdoTypeForValue = 5
        Case vbCurrency
            AdoTypeForValue = 6
        Case vbDate
            AdoTypeForValue = 7
        Case vbBoolean
            AdoTypeForValue = 11
        Case vbString
            AdoTypeForValue = 202
        Case Else
            AdoTypeForValue = 12
    End Select
End Function

Private Function ReadProviderField(ByVal index As Variant) As Variant
    Dim field As Object
    Dim fields As Object

    Set fields = CallByName(mProviderObject, "Fields", VbGet)
    Set field = CallByName(fields, "Item", VbGet, index)
    ReadProviderField = CallByName(field, "Value", VbGet)
End Function

Private Sub CloseProviderObject(ByVal providerValue As Object)
    If providerValue Is Nothing Then Exit Sub
    If CLng(CallByName(providerValue, "State", VbGet)) <> 0 Then
        CallByName providerValue, "Close", VbMethod
    End If
End Sub

Private Function FillTableFromReader( _
    ByVal table As ROneCOne, _
    ByVal reader As ROneCOne _
) As Long
    Dim columnIndex As Long
    Dim row As ROneCOne

    If table.Columns.Count = 0 Then
        For columnIndex = 0 To reader.FieldCount - 1
            table.AddColumn ROneCOne.DataColumn(reader.GetName(columnIndex), vbVariant)
        Next columnIndex
    End If
    Do While reader.Read
        Set row = table.NewRow
        For columnIndex = 0 To reader.FieldCount - 1
            row.Item(columnIndex) = reader.Item(columnIndex)
        Next columnIndex
        table.AddRow row
        FillTableFromReader = FillTableFromReader + 1
    Loop
End Function

Private Sub RequireDataTable( _
    ByVal table As ROneCOne, _
    ByVal memberName As String _
)
    If table Is Nothing Then
        RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
            "The DataTable cannot be Nothing."
    End If
    If table.InternalRole <> ROLE_DATA_TABLE Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            "This operation requires a DataTable."
    End If
End Sub

Private Sub RequireDataColumn( _
    ByVal column As ROneCOne, _
    ByVal memberName As String _
)
    If column Is Nothing Then
        RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
            "The DataColumn cannot be Nothing."
    End If
    If column.InternalRole <> ROLE_DATA_COLUMN Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            "This operation requires a DataColumn."
    End If
End Sub

Private Sub RequireDataRow( _
    ByVal row As ROneCOne, _
    ByVal memberName As String _
)
    If row Is Nothing Then
        RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
            "The DataRow cannot be Nothing."
    End If
    If row.InternalRole <> ROLE_DATA_ROW Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            "This operation requires a DataRow."
    End If
End Sub

Private Sub RequireDataRelation( _
    ByVal relation As ROneCOne, _
    ByVal memberName As String _
)
    If relation Is Nothing Then
        RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
            "The DataRelation cannot be Nothing."
    End If
    If relation.InternalRole <> ROLE_DATA_RELATION Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            "This operation requires a DataRelation."
    End If
End Sub

Private Function NormalizeRelationColumns( _
    ByVal candidate As Variant, _
    ByVal memberName As String _
) As Collection
    Dim candidateObject As Object
    Dim column As ROneCOne
    Dim result As Collection
    Dim sequenceValue As ROneCOne

    Set result = New Collection
    If Not IsObject(candidate) Then
        RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
            "Relation columns must be DataColumns or a typed sequence."
    End If
    Set candidateObject = candidate
    If candidateObject Is Nothing Or Not TypeOf candidateObject Is ROneCOne Then
        RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
            "Relation columns must be ROneCOne values."
    End If
    Set sequenceValue = candidateObject
    If sequenceValue.InternalRole = ROLE_DATA_COLUMN Then
        result.Add sequenceValue
    Else
        RequireSequence sequenceValue, memberName
        For Each column In sequenceValue
            RequireDataColumn column, memberName
            result.Add column
        Next column
    End If
    If result.Count = 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
            "A relation requires at least one column on each side."
    End If
    Set NormalizeRelationColumns = result
End Function

Private Function FindNamedObjectIndex( _
    ByVal values As Collection, _
    ByVal objectName As String _
) As Long
    Dim index As Long
    Dim runtimeValue As ROneCOne

    FindNamedObjectIndex = -1
    For index = 1 To values.Count
        Set runtimeValue = values.Item(index)
        If StrComp(runtimeValue.InternalName, objectName, vbBinaryCompare) = 0 Then
            FindNamedObjectIndex = index - 1
            Exit Function
        End If
    Next index
End Function

Private Function CreateNamedDataCollection( _
    ByVal values As Collection _
) As ROneCOne
    Dim result As ROneCOne
    Dim runtimeValue As ROneCOne

    Set result = New ROneCOne
    result.ConfigureGenericCollection ROLE_COLLECTION, COLLECTION_KEYED, _
        vbObject, "ROneCOne", vbString
    For Each runtimeValue In values
        result.InternalAppendCollectionKey _
            CreateScalarResult(runtimeValue.InternalName)
        result.InternalAppendCollectionItem CreateObjectResult(runtimeValue)
    Next runtimeValue
    result.RebuildHashIndex 8
    Set CreateNamedDataCollection = result
End Function

Private Function CreateObjectListSnapshot( _
    ByVal values As Collection _
) As ROneCOne
    Dim result As ROneCOne
    Dim runtimeValue As ROneCOne

    Set result = New ROneCOne
    result.ConfigureList vbObject, "ROneCOne"
    For Each runtimeValue In values
        result.InternalAddWrapped CreateObjectResult(runtimeValue)
    Next runtimeValue
    Set CreateObjectListSnapshot = result
End Function

Private Function MaterializeDataView() As Collection
    Dim query As ROneCOne
    Dim row As ROneCOne
    Dim rows As ROneCOne

    Set rows = ROneCOne.ListOf(Me)
    For Each row In mSource.InternalDataRows
        If row.RowState <> "Deleted" Then rows.Add row
    Next row
    Set query = rows
    If Not mQueryDelegate Is Nothing Then
        Set query = query.Where(mQueryDelegate)
    End If
    If Not mKeySelector Is Nothing Then
        If mExpressionOption = 0 Then
            Set query = query.OrderBy(mKeySelector)
        Else
            Set query = query.OrderByDescending(mKeySelector)
        End If
    End If
    Set MaterializeDataView = query.MaterializeWrapped
End Function

Private Function CloneDataTableSchema( _
    ByVal sourceTable As ROneCOne _
) As ROneCOne
    Dim clonedColumn As ROneCOne
    Dim column As ROneCOne
    Dim primaryColumns As ROneCOne
    Dim result As ROneCOne

    Set result = ROneCOne.DataTable(sourceTable.TableName)
    For Each column In sourceTable.InternalDataColumns
        Set clonedColumn = New ROneCOne
        clonedColumn.ConfigureDataColumn column.ColumnName, _
            column.InternalExpectedType, column.InternalExpectedObjectType
        clonedColumn.CopyDataColumnConfiguration column
        result.AddColumn clonedColumn
    Next column
    If sourceTable.InternalPrimaryKey.Count > 0 Then
        Set primaryColumns = ROneCOne.ListOf(vbObject)
        For Each column In sourceTable.InternalPrimaryKey
            primaryColumns.Add result.Columns(column.ColumnName)
        Next column
        Set result.PrimaryKey = primaryColumns
    End If
    Set CloneDataTableSchema = result
End Function

Private Sub CopyDataRowInto( _
    ByVal sourceRow As ROneCOne, _
    ByVal targetTable As ROneCOne _
)
    Dim index As Long
    Dim row As ROneCOne

    Set row = targetTable.NewRow
    For index = 0 To targetTable.Columns.Count - 1
        row.Item(index) = Unwrap(sourceRow.InternalDataRows.Item(index + 1))
    Next index
    targetTable.AddRow row
End Sub

Private Sub ValidateCompatibleDataTable( _
    ByVal sourceTable As ROneCOne, _
    ByVal memberName As String _
)
    Dim index As Long
    Dim sourceColumn As ROneCOne
    Dim targetColumn As ROneCOne

    If mKeysCount <> sourceTable.InternalDataColumns.Count Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            "The table schemas have different column counts."
    End If
    For index = 1 To mKeysCount
        Set targetColumn = mKeys(index)
        Set sourceColumn = sourceTable.InternalDataColumns.Item(index)
        If targetColumn.ColumnName <> sourceColumn.ColumnName Or _
            targetColumn.InternalExpectedType <> sourceColumn.InternalExpectedType Or _
            targetColumn.InternalExpectedObjectType <> _
                sourceColumn.InternalExpectedObjectType Then
            RaiseContractError ERROR_INVALID_OPERATION, memberName, _
                "The table schemas are not compatible."
        End If
    Next index
End Sub

Friend Function DataFieldWrapped(ByVal columnName As String) As ROneCOne
    Dim columnIndex As Long

    RequireDataRow Me, "DataRow.Item"
    columnIndex = FindDataColumnIndex(mSource, columnName)
    If columnIndex < 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "DataRow.Item", _
            "The requested column was not found."
    End If
    Set DataFieldWrapped = mItems(columnIndex + 1)
End Function

Private Function DataRowItem(ByVal index As Variant) As ROneCOne
    Dim columnIndex As Long

    If VarTypeOf(index) = vbString Then
        Set DataRowItem = DataFieldWrapped(CStr(index))
        Exit Function
    End If
    columnIndex = CLng(index)
    ValidateIndex columnIndex, mItemsCount, "DataRow.Item"
    Set DataRowItem = mItems(columnIndex + 1)
End Function

Friend Function DataRowVersionValue( _
    ByVal columnName As String, _
    ByVal sourceVersion As String _
) As ROneCOne
    Dim columnIndex As Long

    RequireDataRow Me, "DataRowVersionValue"
    columnIndex = FindNamedObjectIndex(mSource.InternalDataColumns, columnName)
    If columnIndex < 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "DataRowVersionValue", _
            "The requested source column was not found."
    End If
    If NormalizeDataRowVersion(sourceVersion, "DataRowVersionValue") = _
        "Original" And mOriginalItemsCount > 0 Then
        Set DataRowVersionValue = mOriginalItems(columnIndex + 1)
    Else
        Set DataRowVersionValue = mItems(columnIndex + 1)
    End If
End Function

Private Sub SetDataRowItem(ByVal index As Variant, ByVal value As Variant)
    Dim column As ROneCOne
    Dim columnIndex As Long
    Dim previous As ROneCOne
    Dim wrapped As ROneCOne

    Dim capturedDescription As String
    Dim capturedNumber As Long
    Dim capturedSource As String

    RequireDataRow Me, "DataRow.Item"
    If VarTypeOf(index) = vbString Then
        columnIndex = FindDataColumnIndex(mSource, CStr(index))
    Else
        columnIndex = CLng(index)
    End If
    ValidateIndex columnIndex, mItemsCount, "DataRow.Item"
    Set column = mSource.InternalDataColumns.Item(columnIndex + 1)
    Set wrapped = WrapDataValue(value)
    Set wrapped = WidenScalarToward(wrapped, column.InternalExpectedType)
    ValidateDataColumnValue column, wrapped, "DataRow.Item"
    Set previous = mItems(columnIndex + 1)
    ReplaceDataValue columnIndex, wrapped
    On Error GoTo RestorePrevious
    If mRowState <> ROW_STATE_DETACHED Then
        mSource.ValidateDataRowConstraints Me, column
        mSource.ValidateForeignKeyConstraints Me
    End If
    If mRowState = ROW_STATE_UNCHANGED Then mRowState = ROW_STATE_MODIFIED
    If mRowState <> ROW_STATE_DETACHED Then
        mSource.InternalDataRowKeyEdited column
    End If
    mSource.TouchDataVersion
    On Error GoTo 0
    Exit Sub

RestorePrevious:
    capturedNumber = Err.Number
    capturedSource = Err.Source
    capturedDescription = Err.Description
    On Error GoTo 0
    ReplaceDataValue columnIndex, previous
    Err.Raise capturedNumber, capturedSource, capturedDescription
End Sub

Private Function FindDataColumnIndex( _
    ByVal table As ROneCOne, _
    ByVal columnName As String _
) As Long
    FindDataColumnIndex = FindNamedObjectIndex( _
        table.InternalDataColumns, columnName)
End Function

Private Function WrapDataValue(ByVal value As Variant) As ROneCOne
    Dim valueObject As Object
    Dim runtimeValue As ROneCOne

    If IsObject(value) Then
        Set valueObject = value
        If TypeOf valueObject Is ROneCOne Then
            Set runtimeValue = valueObject
            If runtimeValue.InternalRole = ROLE_DB_NULL Then
                Set WrapDataValue = CreateScalarResult(Null)
                Exit Function
            End If
        End If
        Set WrapDataValue = CreateObjectResult(valueObject)
    Else
        Set WrapDataValue = CreateScalarResult(value)
    End If
End Function

Private Function DefaultDataColumnValue( _
    ByVal column As ROneCOne _
) As ROneCOne
    Dim defaultValue As ROneCOne

    If column.InternalAutoIncrement Then
        Set DefaultDataColumnValue = column.NextAutoIncrementValue
        Exit Function
    End If
    Set defaultValue = column.InternalDefaultValue
    If Not defaultValue Is Nothing Then
        Set DefaultDataColumnValue = defaultValue
        Exit Function
    End If
    If column.InternalAllowDBNull Then
        Set DefaultDataColumnValue = CreateScalarResult(Null)
        Exit Function
    End If
    Select Case column.InternalExpectedType
        Case vbString
            Set DefaultDataColumnValue = CreateScalarResult(vbNullString)
        Case vbBoolean
            Set DefaultDataColumnValue = CreateScalarResult(False)
        Case vbDate
            Set DefaultDataColumnValue = CreateScalarResult(CDate(0))
        Case vbByte
            Set DefaultDataColumnValue = CreateScalarResult(CByte(0))
        Case vbInteger
            Set DefaultDataColumnValue = CreateScalarResult(CInt(0))
        Case vbLong
            Set DefaultDataColumnValue = CreateScalarResult(CLng(0))
        Case vbLongLong
            Set DefaultDataColumnValue = CreateScalarResult(CLngLng(0))
        Case vbSingle
            Set DefaultDataColumnValue = CreateScalarResult(CSng(0))
        Case vbDouble
            Set DefaultDataColumnValue = CreateScalarResult(CDbl(0))
        Case vbCurrency
            Set DefaultDataColumnValue = CreateScalarResult(CCur(0))
        Case vbObject
            Set DefaultDataColumnValue = CreateObjectResult(Nothing)
        Case Else
            Set DefaultDataColumnValue = CreateScalarResult(Empty)
    End Select
End Function

Private Sub ValidateDataColumnValue( _
    ByVal column As ROneCOne, _
    ByVal wrapped As ROneCOne, _
    ByVal memberName As String _
)
    If Not wrapped.InternalHasObjectValue Then
        If IsNull(wrapped.InternalScalarValue) Then
            If column.InternalAllowDBNull Then Exit Sub
            RaiseContractError ERROR_TYPE_MISMATCH, memberName, _
                "The DataColumn does not allow Null."
        End If
    End If
    ValidateWrappedType wrapped, column.InternalExpectedType, _
        column.InternalExpectedObjectType, memberName, "value"
End Sub

Private Function BuildPrimaryKeyArguments( _
    ByVal keys As Variant _
) As Collection
    Dim candidate As Object
    Dim hasSingle As Boolean
    Dim rawValue As Variant
    Dim result As Collection
    Dim sequenceValue As ROneCOne

    Set result = New Collection
    On Error Resume Next
    hasSingle = (UBound(keys) = LBound(keys))
    On Error GoTo 0
    If hasSingle Then
        If IsArrayValue(keys(LBound(keys))) Then
            For Each rawValue In keys(LBound(keys))
                result.Add WrapDataValue(rawValue)
            Next rawValue
            Set BuildPrimaryKeyArguments = result
            Exit Function
        End If
        If IsObject(keys(LBound(keys))) Then
            Set candidate = keys(LBound(keys))
            If TypeOf candidate Is ROneCOne Then
                Set sequenceValue = candidate
                If sequenceValue.InternalRole = ROLE_LIST Or _
                    sequenceValue.InternalRole = ROLE_QUERY Or _
                    IsGenericCollectionRole(sequenceValue.InternalRole) Then
                    For Each rawValue In sequenceValue
                        result.Add WrapDataValue(rawValue)
                    Next rawValue
                    Set BuildPrimaryKeyArguments = result
                    Exit Function
                End If
            End If
        End If
    End If
    For Each rawValue In keys
        result.Add WrapDataValue(rawValue)
    Next rawValue
    Set BuildPrimaryKeyArguments = result
End Function

Friend Sub RebuildPrimaryKeyIndex()
    Dim columnIndex As Long
    Dim rowIndex As Long

    ' One pass builds the primary-key index and one token index per Unique
    ' column, so both constraint checks are probes rather than row scans.
    ' Mutations mark these indexes dirty instead of rebuilding per operation;
    ' EnsureDataIndexesCurrent flushes the flag before any probe.
    If mRole <> ROLE_DATA_TABLE Then Exit Sub
    mDataIndexDirty = False
    Set mDataIndex = ROneCOne.DictionaryOf(vbString, ROneCOne.Exception)
    Set mUniqueIndexes = New Collection
    For columnIndex = 1 To mKeysCount
        If mKeys(columnIndex).InternalUnique Then
            mUniqueIndexes.Add ROneCOne.DictionaryOf(vbString, _
                ROneCOne.Exception), CStr(columnIndex)
        End If
    Next columnIndex
    If mPrimaryKey.Count = 0 And mUniqueIndexes.Count = 0 Then Exit Sub
    For rowIndex = 1 To mItemsCount
        If mItems(rowIndex).RowState <> "Deleted" Then
            IndexDataRow mItems(rowIndex)
        End If
    Next rowIndex
End Sub

Private Sub EnsureDataIndexesCurrent()
    If mDataIndexDirty Or mDataIndex Is Nothing Then RebuildPrimaryKeyIndex
End Sub

Friend Sub InternalMarkDataIndexDirty()
    mDataIndexDirty = True
End Sub

Friend Sub InternalDataRowKeyEdited(ByVal column As ROneCOne)
    ' Only an edit to a primary-key or Unique column can change the token
    ' indexes; every other field write leaves them untouched.
    If ColumnInPrimaryKey(column) Or column.InternalUnique Then
        mDataIndexDirty = True
    End If
End Sub

Private Function ColumnInPrimaryKey(ByVal column As ROneCOne) As Boolean
    Dim keyColumn As ROneCOne

    For Each keyColumn In mPrimaryKey
        If keyColumn Is column Then
            ColumnInPrimaryKey = True
            Exit Function
        End If
    Next keyColumn
End Function

Private Sub IndexDataRow(ByVal row As ROneCOne)
    Dim cells As Collection
    Dim columnIndex As Long
    Dim uniqueIndex As ROneCOne

    If mDataIndex Is Nothing Then Exit Sub
    If mPrimaryKey.Count > 0 Then
        If Not mDataIndex.TryAdd(PrimaryKeyTokenForRow(row), row) Then
            RaiseContractError ERROR_INVALID_OPERATION, "PrimaryKey", _
                "The composite primary key contains a duplicate value."
        End If
    End If
    If mUniqueIndexes.Count = 0 Then Exit Sub
    Set cells = row.InternalDataRows
    For columnIndex = 1 To mKeysCount
        If mKeys(columnIndex).InternalUnique Then
            Set uniqueIndex = mUniqueIndexes.Item(CStr(columnIndex))
            uniqueIndex.TryAdd DataIndexToken(cells.Item(columnIndex)), row
        End If
    Next columnIndex
End Sub

Private Function PrimaryKeyTokenForRow( _
    ByVal row As ROneCOne _
) As String
    Dim column As ROneCOne
    Dim token As String

    For Each column In mPrimaryKey
        token = DataIndexToken(row.DataFieldWrapped(column.ColumnName))
        PrimaryKeyTokenForRow = PrimaryKeyTokenForRow & _
            CStr(Len(token)) & ":" & token
    Next column
End Function

Private Function PrimaryKeyTokenFromArguments( _
    ByVal arguments As Collection _
) As String
    Dim column As ROneCOne
    Dim index As Long
    Dim token As String
    Dim wrapped As ROneCOne

    For index = 1 To arguments.Count
        Set column = mPrimaryKey.Item(index)
        Set wrapped = WidenScalarToward(arguments.Item(index), _
            column.InternalExpectedType)
        ValidateDataColumnValue column, wrapped, "Find"
        token = DataIndexToken(wrapped)
        PrimaryKeyTokenFromArguments = PrimaryKeyTokenFromArguments & _
            CStr(Len(token)) & ":" & token
    Next index
End Function

Private Function DataIndexToken(ByVal wrapped As ROneCOne) As String
    Dim scalarValue As Variant

    If wrapped.InternalHasObjectValue Then
        DataIndexToken = "O:" & CStr(ObjPtr(wrapped.InternalObjectValue))
        Exit Function
    End If
    scalarValue = wrapped.InternalScalarValue
    If IsNull(scalarValue) Then
        DataIndexToken = "NULL"
    ElseIf IsEmpty(scalarValue) Then
        DataIndexToken = "N:0"
    ElseIf IsNumeric(scalarValue) Or VarTypeOf(scalarValue) = vbDate Or _
        VarTypeOf(scalarValue) = vbBoolean Then
        DataIndexToken = "N:" & CStr(CDec(scalarValue))
    Else
        DataIndexToken = CStr(VarTypeOf(scalarValue)) & ":" & CStr(scalarValue)
    End If
End Function

Private Sub ReplaceDataValue( _
    ByVal columnIndex As Long, _
    ByVal wrapped As ROneCOne _
)
    ArrReplaceAt mItems, columnIndex + 1, wrapped
End Sub

Friend Sub ValidateDataRowConstraints( _
    ByVal row As ROneCOne, _
    ByVal changedColumn As ROneCOne _
)
    Dim cells As Collection
    Dim column As ROneCOne
    Dim columnIndex As Long
    Dim existing As ROneCOne
    Dim keyToken As String
    Dim uniqueIndex As ROneCOne
    Dim wrapped As ROneCOne

    ' A single-field edit validates only its own column; a new row validates
    ' them all. Uniqueness is a token-index probe, never a row scan, and the
    ' probe tolerates finding this row itself under its unchanged value.
    EnsureDataIndexesCurrent
    Set cells = row.InternalDataRows
    For columnIndex = 1 To mKeysCount
        Set column = mKeys(columnIndex)
        If changedColumn Is Nothing Or column Is changedColumn Then
            Set wrapped = cells.Item(columnIndex)
            ValidateDataColumnValue column, wrapped, "AddRow"
            If column.InternalUnique Then
                Set uniqueIndex = mUniqueIndexes.Item(CStr(columnIndex))
                keyToken = DataIndexToken(wrapped)
                If uniqueIndex.ContainsKey(keyToken) Then
                    Set existing = uniqueIndex.Item(keyToken)
                    If Not existing Is row Then
                        RaiseContractError ERROR_INVALID_OPERATION, "AddRow", _
                            "A unique DataColumn contains a duplicate value."
                    End If
                End If
            End If
        End If
    Next columnIndex
    If mPrimaryKey.Count > 0 Then
        If changedColumn Is Nothing Or ColumnInPrimaryKey(changedColumn) Then
            keyToken = PrimaryKeyTokenForRow(row)
            If mDataIndex.ContainsKey(keyToken) Then
                Set existing = mDataIndex.Item(keyToken)
                If Not existing Is row Then
                    RaiseContractError ERROR_INVALID_OPERATION, "PrimaryKey", _
                        "The composite primary key contains a duplicate value."
                End If
            End If
        End If
    End If
End Sub

Friend Sub ValidateForeignKeyConstraints(ByVal row As ROneCOne)
    Dim dataSetValue As ROneCOne
    Dim relation As ROneCOne

    If mSource Is Nothing Then Exit Sub
    If mSource.InternalRole <> ROLE_DATA_SET Then Exit Sub
    Set dataSetValue = mSource
    If Not dataSetValue.InternalEnforceConstraints Then Exit Sub
    For Each relation In dataSetValue.InternalDataRelations
        If relation.InternalExpressionOption <> 0 Then
            If relation.InternalRelationChildColumn.InternalSource Is Me Then
                ValidateChildRowAgainstRelation row, relation
            End If
        End If
    Next relation
End Sub

Private Function IsNullDataValue(ByVal wrapped As ROneCOne) As Boolean
    If wrapped.InternalHasObjectValue Then
        IsNullDataValue = (wrapped.InternalObjectValue Is Nothing)
    Else
        IsNullDataValue = IsNull(wrapped.InternalScalarValue)
    End If
End Function

Private Function DataValuesEqual( _
    ByVal leftValue As ROneCOne, _
    ByVal rightValue As ROneCOne, _
    ByVal memberName As String _
) As Boolean
    DataValuesEqual = ValuesEqualWithComparer( _
        leftValue, rightValue, Nothing, memberName)
End Function

Private Sub AcceptRowChanges()
    Dim wrapped As ROneCOne

    ArrReset mOriginalItems, mOriginalItemsCount
    For Each wrapped In ArrSnapshot(mItems, mItemsCount)
        ArrAppend mOriginalItems, mOriginalItemsCount, wrapped
    Next wrapped
    mRowState = ROW_STATE_UNCHANGED
End Sub

Private Sub RejectRowChanges()
    Dim wrapped As ROneCOne

    If mRowState = ROW_STATE_ADDED Then
        mRowState = ROW_STATE_DETACHED
        Exit Sub
    End If
    If mRowState = ROW_STATE_UNCHANGED Then Exit Sub
    ArrReset mItems, mItemsCount
    For Each wrapped In ArrSnapshot(mOriginalItems, mOriginalItemsCount)
        ArrAppend mItems, mItemsCount, wrapped
    Next wrapped
    mRowState = ROW_STATE_UNCHANGED
End Sub

Private Sub ValidateRelationOwnership(ByVal relation As ROneCOne)
    Dim childColumn As ROneCOne
    Dim childColumns As Collection
    Dim childTable As ROneCOne
    Dim childRow As ROneCOne
    Dim parentColumn As ROneCOne
    Dim parentColumns As Collection
    Dim parentTable As ROneCOne

    Set parentColumns = relation.InternalRelationParentColumns
    Set childColumns = relation.InternalRelationChildColumns
    Set parentColumn = parentColumns.Item(1)
    Set childColumn = childColumns.Item(1)
    Set parentTable = parentColumn.InternalSource
    Set childTable = childColumn.InternalSource
    If parentTable Is Nothing Or childTable Is Nothing Then
        RaiseContractError ERROR_INVALID_OPERATION, "AddRelation", _
            "Relation columns must belong to DataTables."
    End If
    If Not parentTable.InternalSource Is Me Or _
        Not childTable.InternalSource Is Me Then
        RaiseContractError ERROR_INVALID_OPERATION, "AddRelation", _
            "Both relation tables must belong to this DataSet."
    End If
    For Each parentColumn In parentColumns
        If Not parentColumn.InternalSource Is parentTable Then
            RaiseContractError ERROR_INVALID_OPERATION, "AddRelation", _
                "All parent columns must belong to one DataTable."
        End If
    Next parentColumn
    For Each childColumn In childColumns
        If Not childColumn.InternalSource Is childTable Then
            RaiseContractError ERROR_INVALID_OPERATION, "AddRelation", _
                "All child columns must belong to one DataTable."
        End If
    Next childColumn
    If relation.InternalExpressionOption <> 0 Then
        ValidateExistingUniqueRelation parentColumns
        For Each childRow In childTable.InternalDataRows
            ValidateChildRowAgainstRelation childRow, relation
        Next childRow
        For Each parentColumn In parentColumns
            parentColumn.ConfigurePrimaryKeyColumn (parentColumns.Count = 1)
        Next parentColumn
    End If
End Sub

Private Sub ValidateExistingUniqueRelation(ByVal columns As Collection)
    Dim firstColumn As ROneCOne
    Dim firstIndex As Long
    Dim firstRow As ROneCOne
    Dim rows As Collection
    Dim secondIndex As Long
    Dim secondRow As ROneCOne

    Set firstColumn = columns.Item(1)
    Set rows = firstColumn.InternalSource.InternalDataRows
    For firstIndex = 1 To rows.Count
        Set firstRow = rows.Item(firstIndex)
        For secondIndex = firstIndex + 1 To rows.Count
            Set secondRow = rows.Item(secondIndex)
            If RelationRowsMatch(firstRow, columns, secondRow, columns) Then
                RaiseContractError ERROR_INVALID_OPERATION, "AddRelation", _
                    "The parent relation key contains duplicate values."
            End If
        Next secondIndex
    Next firstIndex
End Sub

Private Sub ValidateChildRowAgainstRelation( _
    ByVal childRow As ROneCOne, _
    ByVal relation As ROneCOne _
)
    Dim childColumn As ROneCOne
    Dim childColumns As Collection
    Dim foundParent As Boolean
    Dim parentColumns As Collection
    Dim parentRow As ROneCOne
    Dim parentTable As ROneCOne

    Set childColumns = relation.InternalRelationChildColumns
    For Each childColumn In childColumns
        If IsNullDataValue( _
            childRow.DataFieldWrapped(childColumn.ColumnName)) Then Exit Sub
    Next childColumn
    Set parentColumns = relation.InternalRelationParentColumns
    Set parentTable = relation.InternalRelationParentColumn.InternalSource
    For Each parentRow In parentTable.InternalDataRows
        If parentRow.RowState <> "Deleted" Then
            If RelationRowsMatch( _
                childRow, childColumns, parentRow, parentColumns) Then
                foundParent = True
                Exit For
            End If
        End If
    Next parentRow
    If Not foundParent Then
        RaiseContractError ERROR_INVALID_OPERATION, "AddRelation", _
            "A DataRelation foreign-key constraint has no parent row."
    End If
End Sub

Private Function NavigateRelation( _
    ByVal relationName As String, _
    ByVal findChildren As Boolean _
) As ROneCOne
    Dim dataSetValue As ROneCOne
    Dim relation As ROneCOne
    Dim relationIndex As Long
    Dim sourceColumn As ROneCOne
    Dim sourceColumns As Collection

    Set dataSetValue = mSource.InternalSource
    If dataSetValue Is Nothing Then
        RaiseContractError ERROR_INVALID_OPERATION, "Relation navigation", _
            "The DataTable does not belong to a DataSet."
    End If
    relationIndex = FindNamedObjectIndex( _
        dataSetValue.InternalDataRelations, relationName)
    If relationIndex < 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "Relation navigation", _
            "The requested DataRelation was not found."
    End If
    Set relation = dataSetValue.InternalDataRelations.Item(relationIndex + 1)
    If findChildren Then
        Set sourceColumns = relation.InternalRelationParentColumns
    Else
        Set sourceColumns = relation.InternalRelationChildColumns
    End If
    Set sourceColumn = sourceColumns.Item(1)
    If Not sourceColumn.InternalSource Is mSource Then
        RaiseContractError ERROR_INVALID_OPERATION, "Relation navigation", _
            "The DataRow is on the wrong side of this relation."
    End If
    Set NavigateRelation = relation.FindRelatedRows(Me, findChildren)
End Function

Private Function RelationRowsMatch( _
    ByVal leftRow As ROneCOne, _
    ByVal leftColumns As Collection, _
    ByVal rightRow As ROneCOne, _
    ByVal rightColumns As Collection _
) As Boolean
    Dim index As Long
    Dim leftColumn As ROneCOne
    Dim rightColumn As ROneCOne

    For index = 1 To leftColumns.Count
        Set leftColumn = leftColumns.Item(index)
        Set rightColumn = rightColumns.Item(index)
        If Not DataValuesEqual( _
            leftRow.DataFieldWrapped(leftColumn.ColumnName), _
            rightRow.DataFieldWrapped(rightColumn.ColumnName), _
            "DataRelation") Then Exit Function
    Next index
    RelationRowsMatch = True
End Function

Private Sub ValidateTaskWork( _
    ByVal work As ROneCOne, _
    ByVal memberName As String _
)
    If work Is Nothing Then
        RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
            "The Task body cannot be Nothing."
    End If
    RequireDelegate work, memberName
    If work.Arity <> 0 And work.Arity <> -1 Then
        RaiseContractError ERROR_ARITY_MISMATCH, memberName, _
            "The Task body must accept no arguments."
    End If
End Sub

Private Sub ValidateProgressHandler( _
    ByVal handler As ROneCOne, _
    ByVal valueType As VbVarType, _
    ByVal objectType As String _
)
    Dim descriptor As ROneCOne

    ValidateAction handler, "ProgressOf", False
    If Not handler.InternalHasParameterSignature Or handler.Arity <> 1 Then
        RaiseContractError ERROR_ARITY_MISMATCH, "ProgressOf", _
            "The progress handler must declare exactly one parameter."
    End If
    Set descriptor = handler.InternalParameters.Item(1)
    If descriptor.InternalExpectedType <> valueType Or _
        StrComp(descriptor.InternalExpectedObjectType, objectType, _
            vbTextCompare) <> 0 Then
        RaiseContractError ERROR_TYPE_MISMATCH, "ProgressOf", _
            "The progress handler parameter does not match Progress<T>."
    End If
End Sub

Private Sub RequireTaskCompletionSource(ByVal memberName As String)
    If mRole = ROLE_TASK_COMPLETION_SOURCE Then Exit Sub
    RaiseContractError ERROR_INVALID_OPERATION, memberName, _
        "This operation requires a TaskCompletionSource<T>."
End Sub

Private Function WrapTypedCompletionValue(ByVal value As Variant) As ROneCOne
    Dim wrapped As ROneCOne

    Set wrapped = WrapInvocationArgument(value)
    Set wrapped = WidenScalarToward(wrapped, mExpectedType)
    ValidateWrappedType wrapped, mExpectedType, mExpectedObjectType, _
        "TaskCompletionSource", "result"
    Set WrapTypedCompletionValue = wrapped
End Function

Private Sub ValidateCancellationToken( _
    ByVal tokenValue As ROneCOne, _
    ByVal memberName As String _
)
    If tokenValue Is Nothing Then Exit Sub
    If tokenValue.InternalRole <> ROLE_CANCELLATION_TOKEN Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            "The cancellation value must be a CancellationToken."
    End If
End Sub

Private Sub RequireTask(ByVal taskValue As ROneCOne, ByVal memberName As String)
    If taskValue Is Nothing Then
        RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
            "The Task cannot be Nothing."
    End If
    If taskValue.InternalRole <> ROLE_TASK Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            "This operation requires a Task."
    End If
End Sub

Private Function CreateTaskCombinator( _
    ByVal tasks As Variant, _
    ByVal taskKind As Long, _
    ByVal memberName As String _
) As ROneCOne
    Dim candidate As Object
    Dim child As ROneCOne
    Dim hasAny As Boolean
    Dim hasSingle As Boolean
    Dim rawTask As Variant
    Dim sequenceValue As ROneCOne
    Dim taskValue As ROneCOne

    Set taskValue = New ROneCOne
    taskValue.ConfigureTask taskKind, Nothing, Nothing
    If Not IsArrayValue(tasks) Then
        RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
            "Tasks must be supplied directly or as a typed sequence or array."
    End If
    On Error Resume Next
    hasAny = (UBound(tasks) >= LBound(tasks))
    hasSingle = (UBound(tasks) = LBound(tasks))
    On Error GoTo 0
    If Not hasAny Then
        If taskKind = TASK_WHEN_ANY Then
            RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
                "WhenAny requires at least one Task."
        End If
        Set CreateTaskCombinator = taskValue
        Exit Function
    End If
    If hasSingle And IsObject(tasks(LBound(tasks))) Then
        Set candidate = tasks(LBound(tasks))
        If Not candidate Is Nothing Then
            If TypeOf candidate Is ROneCOne Then
                Set sequenceValue = candidate
                If sequenceValue.InternalRole <> ROLE_TASK Then
                    RequireSequence sequenceValue, memberName
                    For Each rawTask In sequenceValue
                        Set child = RequireROneCOne(rawTask, memberName)
                        RequireTask child, memberName
                        taskValue.InternalAddTaskChild child
                    Next rawTask
                    If taskKind = TASK_WHEN_ANY And _
                        taskValue.InternalTaskChildCount = 0 Then
                        RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
                            "WhenAny requires at least one Task."
                    End If
                    Set CreateTaskCombinator = taskValue
                    Exit Function
                End If
            End If
        End If
    End If
    For Each rawTask In tasks
        Set child = RequireROneCOne(rawTask, memberName)
        RequireTask child, memberName
        taskValue.InternalAddTaskChild child
    Next rawTask
    Set CreateTaskCombinator = taskValue
End Function

Friend Sub EnsureTaskCompleted()
    Call WaitForTaskCompletion(-1)
End Sub

Friend Sub InternalRaiseTaskTerminalState(ByVal memberName As String)
    RaiseTaskTerminalState memberName
End Sub

Friend Sub AdvanceTask()
Attribute AdvanceTask.VB_MemberFlags = "40"
    Dim capturedDescription As String
    Dim capturedHelpContext As Long
    Dim capturedHelpFile As String
    Dim capturedNumber As Long
    Dim capturedSource As String

    If mTaskStatus >= TASK_STATUS_COMPLETED Then Exit Sub
    If mTaskAdvancing Then
        RaiseContractError ERROR_INVALID_OPERATION, "Await", _
            "A Task cannot await itself or contain a dependency cycle."
    End If
    mTaskAdvancing = True
    On Error GoTo TaskFailed
    AdvanceTaskCore
    mTaskAdvancing = False
    On Error GoTo 0
    Exit Sub

TaskFailed:
    capturedNumber = Err.Number
    capturedSource = Err.Source
    capturedDescription = Err.Description
    capturedHelpFile = Err.HelpFile
    capturedHelpContext = Err.HelpContext
    On Error GoTo 0
    Set mTaskException = New ROneCOne
    mTaskException.ConfigureError capturedNumber, capturedSource, _
        capturedDescription, capturedHelpFile, capturedHelpContext
    mTaskStatus = TASK_STATUS_FAULTED
    mTaskAdvancing = False
End Sub

Private Sub AdvanceTaskCore()
    Dim arguments As Collection
    Dim allCompleted As Boolean
    Dim child As ROneCOne
    Dim dataTarget As ROneCOne
    Dim errors As Collection
    Dim results As ROneCOne
    Dim sawCanceled As Boolean

    If TaskCancellationRequested Then
        ' An in-flight web request must release its transport when canceled,
        ' and an in-flight provider operation must cancel inside ADO.
        If mOperator = TASK_HTTP_SEND Then mSource.InternalHttpAbort
        If mOperator = TASK_DB_ASYNC Then mSource.InternalDbAsyncAbort
        If mOperator = TASK_PROCESS_RUN Then mSource.InternalProcessAbort
        ' A canceled read must hand its slot back, or the session can never
        ' be read again.
        If mOperator = TASK_SESSION Then
            If mMethodName = "ReadOut" Or mMethodName = "ReadErr" Then
                mSource.InternalSessionReleaseRead mMethodName = "ReadErr"
            End If
        End If
        mTaskStatus = TASK_STATUS_CANCELED
        Exit Sub
    End If
    If mOperator = TASK_PROMISE Then Exit Sub
    mTaskStatus = TASK_STATUS_RUNNING
    Select Case mOperator
        Case TASK_WORK
            Set arguments = New Collection
            Set mTaskResult = mTaskBody.InvokeValue(arguments)
        Case TASK_DELAY
            AdvanceDelayTask
            Exit Sub
        Case TASK_HTTP_SEND
            If Not mSource.InternalHttpPoll( _
                mMethodName = "Bytes" Or mMethodName = "File") Then
                Exit Sub
            End If
            If mMethodName = "String" Or mMethodName = "Bytes" Or _
                mMethodName = "File" Then
                ' The body-consuming verbs fault on non-success, mirroring
                ' System.Net.Http; GetAsync hands back the response for the
                ' caller to inspect.
                If Not mSource.IsSuccessStatusCode Then
                    Set mTaskException = CreateHttpStatusException(mSource)
                    mTaskStatus = TASK_STATUS_FAULTED
                    Exit Sub
                End If
            End If
            Select Case mMethodName
                Case "String"
                    Set mTaskResult = CreateScalarResult(mSource.Content)
                Case "Bytes"
                    Set mTaskResult = CreateScalarResult( _
                        mSource.InternalHttpBytes)
                Case "File"
                    IoWriteAllBytesCore mName, NormalizeHttpBodyBytes( _
                        mSource.InternalHttpBytes), "DownloadFileAsync"
                    Set mTaskResult = CreateScalarResult(mName)
                Case Else
                    Set mTaskResult = CreateObjectResult(mSource)
            End Select
        Case TASK_YIELD
            If mExpressionOption = 0 Then
                mExpressionOption = 1
                Exit Sub
            End If
            Set mTaskResult = CreateScalarResult(Empty)
        Case TASK_WHEN_ALL
            allCompleted = True
            For Each child In mTaskChildren
                child.AdvanceTask
                If Not child.IsCompleted Then allCompleted = False
            Next child
            If Not allCompleted Then Exit Sub
            Set results = ROneCOne.ListOf(vbVariant)
            Set errors = New Collection
            For Each child In mTaskChildren
                If child.IsFaulted Then
                    AppendFlattenedExceptions child.Exception, errors
                End If
                If child.IsCanceled Then sawCanceled = True
                If Not child.IsFaulted And Not child.IsCanceled Then
                    results.Add child.Result
                End If
            Next child
            If errors.Count > 0 Then
                Set mTaskException = CreateAggregateException(errors)
                mTaskStatus = TASK_STATUS_FAULTED
                Exit Sub
            End If
            If sawCanceled Then
                mTaskStatus = TASK_STATUS_CANCELED
                Exit Sub
            End If
            Set mTaskResult = CreateObjectResult(results)
        Case TASK_WHEN_ANY
            If mTaskChildren.Count = 0 Then
                RaiseContractError ERROR_INVALID_ARGUMENT, "WhenAny", _
                    "WhenAny requires at least one Task."
            End If
            For Each child In mTaskChildren
                child.AdvanceTask
                If child.IsCompleted Then
                    Set mTaskResult = CreateObjectResult(child)
                    mTaskStatus = TASK_STATUS_COMPLETED
                    Exit Sub
                End If
            Next child
            Exit Sub
        Case TASK_WAIT_ASYNC
            mSource.AdvanceTask
            If mSource.IsCompleted Then
                mTaskStatus = mSource.InternalTaskStatus
                Set mTaskResult = mSource.InternalTaskResult
                Set mTaskException = mSource.InternalTaskException
                Exit Sub
            End If
            If mTaskDeadline >= 0 And _
                MonotonicMilliseconds >= mTaskDeadline Then
                Set mTaskException = New ROneCOne
                mTaskException.ConfigureError ERROR_TIMEOUT, _
                    "ROneCOne.TimeoutException", _
                    "The asynchronous wait timed out.", vbNullString, 0
                mTaskException.ConfigureExceptionType "TimeoutException"
                mTaskStatus = TASK_STATUS_FAULTED
            End If
            Exit Sub
        Case TASK_CONTINUATION
            mSource.AdvanceTask
            If Not mSource.IsCompleted Then Exit Sub
            Set arguments = New Collection
            arguments.Add CreateObjectResult(mSource)
            Set mTaskResult = mTaskBody.InvokeValue(arguments)
        Case TASK_PROVIDER_CALL
            Select Case mMethodName
                Case "ExecuteNonQuery"
                    Set mTaskResult = CreateScalarResult(mSource.ExecuteNonQuery)
                Case "Read"
                    Set mTaskResult = CreateScalarResult(mSource.Read)
                Case "Update"
                    Set dataTarget = RequireProviderTaskTarget(mTaskChildren)
                    Set mTaskResult = CreateScalarResult(mSource.Update(dataTarget))
                Case Else
                    RaiseContractError ERROR_INVALID_OPERATION, "ProviderTask", _
                        "Unknown provider Task operation."
            End Select
        Case TASK_DB_ASYNC
            Select Case mMethodName
                Case "Connect"
                    If Not mSource.InternalDbConnectPoll Then Exit Sub
                    Set mTaskResult = CreateScalarResult(Empty)
                Case "Reader"
                    If Not mSource.InternalDbResultPoll Then Exit Sub
                    Set mTaskResult = CreateObjectResult(mSource)
                Case "Scalar"
                    If Not mSource.InternalDbResultPoll Then Exit Sub
                    Set mTaskResult = WrapDataValue(mSource.InternalScalarHarvest)
                Case "Fill"
                    If Not mSource.InternalDbResultPoll Then Exit Sub
                    Set mTaskResult = CreateScalarResult( _
                        DbAsyncTaskChild(1).InternalFillFromReader( _
                            mSource, DbAsyncTaskChild(2)))
                Case "QueryFill"
                    If Not mSource.InternalDbResultPoll Then Exit Sub
                    DbAsyncTaskChild(1).InternalFillFromReader _
                        mSource, DbAsyncTaskChild(2)
                    Set mTaskResult = CreateObjectResult(DbAsyncTaskChild(2))
                Case Else
                    RaiseContractError ERROR_INVALID_OPERATION, "DbAsyncTask", _
                        "Unknown provider async operation."
            End Select
        Case TASK_PROCESS_RUN
            If Not mSource.InternalProcessPoll Then Exit Sub
            Set mTaskResult = CreateObjectResult(mSource)
        Case TASK_FILE_WATCH
            If Not InternalWatchPoll() Then Exit Sub
        Case TASK_SESSION
            If Not AdvanceSessionTask() Then Exit Sub
    End Select
    mTaskStatus = TASK_STATUS_COMPLETED
End Sub

Private Function AdvanceSessionTask() As Boolean
    Dim finished As Boolean
    Dim lineText As String

    Select Case mMethodName
        Case "Write"
            If Not mSource.InternalSessionWritePoll(mExpressionOption) Then
                Exit Function
            End If
            Set mTaskResult = CreateScalarResult(Empty)
        Case "ReadOut", "ReadErr"
            If Not mSource.InternalSessionReadPoll( _
                mMethodName = "ReadErr", lineText, finished) Then
                ' A prompt arrives without a trailing newline, so a line
                ' read can legitimately have nothing to return while the
                ' command sits waiting. An optional deadline resolves to
                ' empty text instead of waiting forever, and it releases
                ' the read slot so the session stays usable.
                If mTaskDeadline < 0 Then Exit Function
                If MonotonicMilliseconds < mTaskDeadline Then Exit Function
                mSource.InternalSessionReleaseRead mMethodName = "ReadErr"
                Set mTaskResult = CreateScalarResult(vbNullString)
                AdvanceSessionTask = True
                Exit Function
            End If
            mSource.InternalSessionReleaseRead mMethodName = "ReadErr"
            If finished Then
                Set mTaskResult = CreateScalarResult(vbNullString)
            Else
                Set mTaskResult = CreateScalarResult(lineText)
            End If
        Case "ReadToEnd"
            If Not mSource.InternalSessionExitPoll() Then Exit Function
            Set mTaskResult = CreateScalarResult( _
                mSource.InternalSessionRemainingText(False))
        Case "Exit"
            If Not mSource.InternalSessionExitPoll() Then Exit Function
            Set mTaskResult = CreateScalarResult( _
                mSource.InternalSessionExitCode())
        Case Else
            RaiseContractError ERROR_INVALID_OPERATION, "SessionTask", _
                "Unknown process session operation."
    End Select
    AdvanceSessionTask = True
End Function

Private Function WaitForTaskCompletion( _
    ByVal timeoutMilliseconds As Long _
) As Boolean
    Dim capturedDescription As String
    Dim capturedHelpContext As Long
    Dim capturedHelpFile As String
    Dim capturedNumber As Long
    Dim capturedSource As String
    Dim startedAt As LongLong

    If mTaskWaitActive Then
        RaiseContractError ERROR_INVALID_OPERATION, "Await", _
            "Reentrant waits on the same Task are not supported."
    End If
    mTaskWaitActive = True
    startedAt = MonotonicMilliseconds
    On Error GoTo WaitFailed
    Do While mTaskStatus < TASK_STATUS_COMPLETED
        AdvanceTask
        If mTaskStatus >= TASK_STATUS_COMPLETED Then Exit Do
        DoEvents
        Sleep 1
        If timeoutMilliseconds >= 0 Then
            If MonotonicMilliseconds - startedAt >= _
                CLngLng(timeoutMilliseconds) Then
                mTaskWaitActive = False
                Exit Function
            End If
        End If
    Loop
    mTaskWaitActive = False
    WaitForTaskCompletion = True
    On Error GoTo 0
    Exit Function

WaitFailed:
    capturedNumber = Err.Number
    capturedSource = Err.Source
    capturedDescription = Err.Description
    capturedHelpFile = Err.HelpFile
    capturedHelpContext = Err.HelpContext
    mTaskWaitActive = False
    On Error GoTo 0
    Err.Raise capturedNumber, capturedSource, capturedDescription, _
        capturedHelpFile, capturedHelpContext
End Function

Private Sub AdvanceDelayTask()
    If TaskCancellationRequested Then
        mTaskStatus = TASK_STATUS_CANCELED
        Exit Sub
    End If
    If MonotonicMilliseconds >= mTaskDeadline Then
        Set mTaskResult = CreateScalarResult(Empty)
        mTaskStatus = TASK_STATUS_COMPLETED
    End If
End Sub

Private Function MonotonicMilliseconds() As LongLong
    MonotonicMilliseconds = GetTickCount64()
End Function

Private Function TaskCancellationRequested() As Boolean
    If mCancellationSource Is Nothing Then Exit Function
    TaskCancellationRequested = mCancellationSource.IsCancellationRequested
End Function

Private Sub EvaluateScheduledCancellation()
    If mCancellationRequested Or mQueryCount < 0 Then Exit Sub
    If MonotonicMilliseconds >= mTaskDeadline Then
        CancelCancellationSource
    End If
End Sub

Private Sub CancelCancellationSource()
    Dim callback As ROneCOne
    Dim callbacks As Collection
    Dim callbackError As ROneCOne
    Dim errors As Collection

    If mCancellationRequested Then Exit Sub
    mCancellationRequested = True
    Set callbacks = CopyWrappedCollection(mInvocationList)
    Set errors = New Collection
    For Each callback In callbacks
        Set callbackError = InvokeCancellationCallback(callback)
        If Not callbackError Is Nothing Then errors.Add callbackError
    Next callback
    Set mInvocationList = New Collection
    If errors.Count > 0 Then RaiseCapturedError _
        CreateAggregateException(errors)
End Sub

Private Function InvokeCancellationCallback( _
    ByVal callback As ROneCOne _
) As ROneCOne
    Dim arguments As Collection
    Dim capturedDescription As String
    Dim capturedHelpContext As Long
    Dim capturedHelpFile As String
    Dim capturedNumber As Long
    Dim capturedSource As String
    Dim ignored As ROneCOne

    On Error GoTo CallbackFailed
    Set arguments = New Collection
    Set ignored = callback.InvokeValue(arguments)
    On Error GoTo 0
    Exit Function

CallbackFailed:
    capturedNumber = Err.Number
    capturedSource = Err.Source
    capturedDescription = Err.Description
    capturedHelpFile = Err.HelpFile
    capturedHelpContext = Err.HelpContext
    On Error GoTo 0
    Set InvokeCancellationCallback = New ROneCOne
    InvokeCancellationCallback.ConfigureError capturedNumber, _
        capturedSource, capturedDescription, capturedHelpFile, _
        capturedHelpContext
End Function

Private Function CreateAggregateException( _
    ByVal errors As Collection _
) As ROneCOne
    Dim aggregate As ROneCOne

    Set aggregate = New ROneCOne
    aggregate.ConfigureAggregateError errors
    Set CreateAggregateException = aggregate
End Function

Private Function AsAggregateException( _
    ByVal errorValue As ROneCOne _
) As ROneCOne
    Dim errors As Collection

    If errorValue Is Nothing Then Exit Function
    If errorValue.ExceptionType = "AggregateException" Then
        Set AsAggregateException = errorValue
        Exit Function
    End If
    Set errors = New Collection
    errors.Add errorValue
    Set AsAggregateException = CreateAggregateException(errors)
End Function

Private Sub AppendFlattenedExceptions( _
    ByVal errorValue As ROneCOne, _
    ByVal output As Collection _
)
    Dim inner As ROneCOne

    If errorValue Is Nothing Then Exit Sub
    If errorValue.ExceptionType <> "AggregateException" Then
        output.Add errorValue
        Exit Sub
    End If
    For Each inner In errorValue.InternalInnerExceptions
        AppendFlattenedExceptions inner, output
    Next inner
End Sub

Private Function PrimaryException( _
    ByVal errorValue As ROneCOne _
) As ROneCOne
    Dim flattened As Collection

    If errorValue Is Nothing Then Exit Function
    If errorValue.ExceptionType <> "AggregateException" Then
        Set PrimaryException = errorValue
        Exit Function
    End If
    Set flattened = New Collection
    AppendFlattenedExceptions errorValue, flattened
    If flattened.Count > 0 Then Set PrimaryException = flattened.Item(1)
End Function

Private Sub RaiseTaskTerminalState(ByVal memberName As String)
    If mTaskStatus = TASK_STATUS_CANCELED Then
        Err.Raise ERROR_CANCELED, _
            "ROneCOne.OperationCanceledException", _
            "The Task was canceled."
    End If
    If mTaskStatus = TASK_STATUS_FAULTED Then
        RaiseCapturedError PrimaryException(mTaskException)
    End If
End Sub

Private Function IsGenericCollectionRole(ByVal roleValue As Long) As Boolean
    Select Case roleValue
        Case ROLE_DICTIONARY, ROLE_HASH_SET, ROLE_COLLECTION
            IsGenericCollectionRole = True
        Case ROLE_IMMUTABLE, ROLE_CONCURRENT
            IsGenericCollectionRole = True
    End Select
End Function

Friend Function IsDictionaryCollection() As Boolean
    Select Case mCollectionKind
        Case COLLECTION_DICTIONARY, COLLECTION_SORTED_LIST
            IsDictionaryCollection = True
        Case COLLECTION_SORTED_DICTIONARY, COLLECTION_ORDERED_DICTIONARY
            IsDictionaryCollection = True
        Case COLLECTION_CONCURRENT_DICTIONARY
            IsDictionaryCollection = True
        Case COLLECTION_IMMUTABLE_DICTIONARY
            IsDictionaryCollection = True
        Case COLLECTION_IMMUTABLE_SORTED_DICTIONARY
            IsDictionaryCollection = True
    End Select
End Function

Private Function IsSetCollection() As Boolean
    Select Case mCollectionKind
        Case COLLECTION_HASH_SET, COLLECTION_SORTED_SET
            IsSetCollection = True
        Case COLLECTION_IMMUTABLE_HASH_SET, COLLECTION_IMMUTABLE_SORTED_SET
            IsSetCollection = True
    End Select
End Function

Private Function IsQueueKind(ByVal collectionKind As Long) As Boolean
    Select Case collectionKind
        Case COLLECTION_QUEUE, COLLECTION_CONCURRENT_QUEUE
            IsQueueKind = True
        Case COLLECTION_BLOCKING, COLLECTION_IMMUTABLE_QUEUE
            IsQueueKind = True
    End Select
End Function

Private Function IsStackKind(ByVal collectionKind As Long) As Boolean
    Select Case collectionKind
        Case COLLECTION_STACK, COLLECTION_CONCURRENT_STACK
            IsStackKind = True
        Case COLLECTION_IMMUTABLE_STACK
            IsStackKind = True
    End Select
End Function

Private Function IsPriorityQueueKind(ByVal collectionKind As Long) As Boolean
    IsPriorityQueueKind = (collectionKind = COLLECTION_PRIORITY_QUEUE)
End Function

Private Function MutableRoleForKind(ByVal collectionKind As Long) As Long
    Select Case collectionKind
        Case COLLECTION_IMMUTABLE_DICTIONARY
            MutableRoleForKind = ROLE_DICTIONARY
        Case COLLECTION_IMMUTABLE_SORTED_DICTIONARY
            MutableRoleForKind = ROLE_DICTIONARY
        Case COLLECTION_IMMUTABLE_HASH_SET, COLLECTION_IMMUTABLE_SORTED_SET
            MutableRoleForKind = ROLE_HASH_SET
        Case Else
            MutableRoleForKind = ROLE_COLLECTION
    End Select
End Function

Private Sub RequireMutableCollection(ByVal memberName As String)
    If Not IsGenericCollectionRole(mRole) Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            "This operation requires a generic collection."
    End If
    If mRole = ROLE_IMMUTABLE Or mCollectionKind = COLLECTION_READ_ONLY Or _
        mCollectionKind = COLLECTION_LOOKUP Or _
        mCollectionKind = COLLECTION_GROUPING Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            "This collection is read-only."
    End If
End Sub

Private Sub RequireKeyedCollection(ByVal memberName As String)
    If IsDictionaryCollection Or mCollectionKind = COLLECTION_KEYED Or _
        mCollectionKind = COLLECTION_LOOKUP Then Exit Sub
    RaiseContractError ERROR_INVALID_OPERATION, memberName, _
        "This operation requires a keyed collection."
End Sub

Private Sub RequireSetCollection(ByVal memberName As String)
    If IsSetCollection Then Exit Sub
    RaiseContractError ERROR_INVALID_OPERATION, memberName, _
        "This operation requires a set."
End Sub

Private Sub RequireConcurrentDictionary(ByVal memberName As String)
    If mRole = ROLE_CONCURRENT And _
        mCollectionKind = COLLECTION_CONCURRENT_DICTIONARY Then Exit Sub
    RaiseContractError ERROR_INVALID_OPERATION, memberName, _
        "This operation requires a ConcurrentDictionary."
End Sub

Private Sub RequireBlockingCollection(ByVal memberName As String)
    If mRole = ROLE_CONCURRENT And _
        mCollectionKind = COLLECTION_BLOCKING Then Exit Sub
    RaiseContractError ERROR_INVALID_OPERATION, memberName, _
        "This operation requires a BlockingCollection."
End Sub

Private Sub RequirePriorityQueue(ByVal memberName As String)
    If IsPriorityQueueKind(mCollectionKind) Then Exit Sub
    RaiseContractError ERROR_INVALID_OPERATION, memberName, _
        "This operation requires a PriorityQueue."
End Sub

Private Sub RequireImmutableCollection(ByVal memberName As String)
    If mRole = ROLE_IMMUTABLE Then Exit Sub
    RaiseContractError ERROR_INVALID_OPERATION, memberName, _
        "This operation requires an immutable collection."
End Sub

Private Sub RequireCollectionNode(ByVal memberName As String)
    If mRole = ROLE_NODE And Not mSource Is Nothing Then Exit Sub
    RaiseContractError ERROR_INVALID_OPERATION, memberName, _
        "This operation requires an attached LinkedListNode."
End Sub

Private Sub RequireOwnedNode( _
    ByVal node As ROneCOne, _
    ByVal memberName As String _
)
    If node Is Nothing Or node.InternalRole <> ROLE_NODE Then
        RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
            "The node cannot be Nothing and must be a LinkedListNode."
    End If
    If Not node.InternalSource Is Me Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            "The node belongs to a different linked list."
    End If
End Sub

Private Sub RequireLinkedList(ByVal memberName As String)
    If mCollectionKind = COLLECTION_LINKED_LIST Then Exit Sub
    RaiseContractError ERROR_INVALID_OPERATION, memberName, _
        "This operation requires a LinkedList<T>."
End Sub

Private Sub RequireSortedSet(ByVal memberName As String)
    If mCollectionKind = COLLECTION_SORTED_SET Or _
        mCollectionKind = COLLECTION_IMMUTABLE_SORTED_SET Then Exit Sub
    RaiseContractError ERROR_INVALID_OPERATION, memberName, _
        "This operation requires a sorted set."
End Sub

Private Sub RequireReference( _
    ByVal referenceValue As ROneCOne, _
    ByVal memberName As String _
)
    If referenceValue Is Nothing Then
        RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
            "The output reference cannot be Nothing."
    End If
    If referenceValue.InternalRole <> ROLE_REFERENCE Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            "The output must be created by a ROneCOne.Ref method."
    End If
End Sub

Friend Function WrapKey( _
    ByVal value As Variant, _
    ByVal memberName As String _
) As ROneCOne
    Dim valueObject As Object
    Dim wrapped As ROneCOne

    If IsObject(value) Then
        Set valueObject = value
        Set wrapped = CreateObjectResult(valueObject)
    Else
        If IsNull(value) Then
            RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
                "A collection key cannot be Null."
        End If
        Set wrapped = CreateScalarResult(value)
    End If
    Set wrapped = WidenScalarToward(wrapped, mKeyType)
    ValidateWrappedType wrapped, mKeyType, mKeyObjectType, memberName, "key"
    Set WrapKey = wrapped
End Function

Private Function IsLosslessWidening( _
    ByVal fromType As VbVarType, _
    ByVal toType As VbVarType _
) As Boolean
    ' True when every value of fromType is representable in toType with no loss
    ' of magnitude or precision. Narrowing, cross-family, and Boolean/Date/String
    ' conversions are deliberately excluded so admission stays strict where it
    ' protects and only relaxes where the promotion is provably safe.
    If fromType = toType Then Exit Function
    Select Case toType
        Case vbInteger
            IsLosslessWidening = (fromType = vbByte)
        Case vbLong
            IsLosslessWidening = (fromType = vbByte Or fromType = vbInteger)
        Case vbLongLong
            IsLosslessWidening = (fromType = vbByte Or fromType = vbInteger Or _
                fromType = vbLong)
        Case vbCurrency
            IsLosslessWidening = (fromType = vbByte Or fromType = vbInteger Or _
                fromType = vbLong)
        Case vbSingle
            IsLosslessWidening = (fromType = vbByte Or fromType = vbInteger)
        Case vbDouble
            IsLosslessWidening = (fromType = vbByte Or fromType = vbInteger Or _
                fromType = vbLong Or fromType = vbSingle)
    End Select
End Function

Private Function CoerceScalarToType( _
    ByVal value As Variant, _
    ByVal toType As VbVarType _
) As Variant
    Select Case toType
        Case vbInteger
            CoerceScalarToType = CInt(value)
        Case vbLong
            CoerceScalarToType = CLng(value)
        Case vbLongLong
            CoerceScalarToType = CLngLng(value)
        Case vbCurrency
            CoerceScalarToType = CCur(value)
        Case vbSingle
            CoerceScalarToType = CSng(value)
        Case vbDouble
            CoerceScalarToType = CDbl(value)
        Case Else
            CoerceScalarToType = value
    End Select
End Function

Private Function WidenScalarToward( _
    ByVal wrapped As ROneCOne, _
    ByVal expectedType As VbVarType _
) As ROneCOne
    ' Returns wrapped unchanged unless its scalar is losslessly widenable to
    ' expectedType, in which case a fresh wrapped holding the coerced value is
    ' returned so the stored value's runtime type matches the declared type
    ' exactly. Never raises; the strict validator that follows is the gate.
    Dim coerced As ROneCOne

    Set WidenScalarToward = wrapped
    If expectedType = vbVariant Or expectedType = vbObject Then Exit Function
    If wrapped.InternalHasObjectValue Then Exit Function
    If IsNull(wrapped.InternalScalarValue) Then Exit Function
    If IsEmpty(wrapped.InternalScalarValue) Then Exit Function
    If Not IsLosslessWidening( _
        VarTypeOf(wrapped.InternalScalarValue), expectedType) Then Exit Function
    Set coerced = New ROneCOne
    coerced.ConfigureScalarValue CoerceScalarToType( _
        wrapped.InternalScalarValue, expectedType)
    Set WidenScalarToward = coerced
End Function

Private Sub ReplaceCollectionItemAt( _
    ByVal target As Collection, _
    ByVal position As Long, _
    ByVal item As ROneCOne _
)
    If position < target.Count Then
        target.Add item, Before:=position
        target.Remove position + 1
    Else
        target.Remove position
        target.Add item
    End If
End Sub

Private Sub ValidateWrappedType( _
    ByVal wrapped As ROneCOne, _
    ByVal expectedType As VbVarType, _
    ByVal expectedObjectType As String, _
    ByVal memberName As String, _
    ByVal valueName As String _
)
    Dim actualObject As Object

    If expectedType = vbVariant Then Exit Sub
    If expectedType = vbObject Then
        If Not wrapped.InternalHasObjectValue Then GoTo TypeFailed
        Set actualObject = wrapped.InternalObjectValue
        If actualObject Is Nothing Then Exit Sub
        If TypeName(actualObject) = expectedObjectType Then Exit Sub
        GoTo TypeFailed
    End If
    If wrapped.InternalHasObjectValue Then GoTo TypeFailed
    If VarTypeOf(wrapped.InternalScalarValue) = expectedType Then Exit Sub
TypeFailed:
    RaiseContractError ERROR_TYPE_MISMATCH, memberName, _
        "The collection " & valueName & " does not match its generic type."
End Sub

Private Sub ValidateKeyValue( _
    ByVal value As Variant, _
    ByVal memberName As String _
)
    Dim valueObject As Object

    If IsNull(value) Then
        RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
            "A collection key cannot be Null."
    End If
    If mKeyType = vbVariant Then Exit Sub
    If mKeyType = vbObject Then
        If Not IsObject(value) Then GoTo TypeFailed
        Set valueObject = value
        If valueObject Is Nothing Then Exit Sub
        If TypeName(valueObject) = mKeyObjectType Then Exit Sub
        GoTo TypeFailed
    End If
    If IsObject(value) Then GoTo TypeFailed
    If VarTypeOf(value) = mKeyType Then Exit Sub
    ' A losslessly widenable key hashes and compares identically to the stored
    ' key of the declared type, so accept it for lookup without ceremony.
    If IsLosslessWidening(VarTypeOf(value), mKeyType) Then Exit Sub
TypeFailed:
    RaiseContractError ERROR_TYPE_MISMATCH, memberName, _
        "The collection key does not match its generic type."
End Sub

Private Function BuildCollectionTypeName() As String
    Dim keyName As String
    Dim valueName As String

    If Len(mElementObjectType) > 0 Then
        valueName = mElementObjectType
    Else
        valueName = TypeNameForVarType(mElementType)
    End If
    If Len(mKeyObjectType) > 0 Then
        keyName = mKeyObjectType
    Else
        keyName = TypeNameForVarType(mKeyType)
    End If
    Select Case mCollectionKind
        Case COLLECTION_DICTIONARY
            BuildCollectionTypeName = "Dictionary<" & keyName & ", " & valueName & ">"
        Case COLLECTION_HASH_SET
            BuildCollectionTypeName = "HashSet<" & valueName & ">"
        Case COLLECTION_QUEUE
            BuildCollectionTypeName = "Queue<" & valueName & ">"
        Case COLLECTION_STACK
            BuildCollectionTypeName = "Stack<" & valueName & ">"
        Case COLLECTION_LINKED_LIST
            BuildCollectionTypeName = "LinkedList<" & valueName & ">"
        Case COLLECTION_SORTED_LIST
            BuildCollectionTypeName = "SortedList<" & keyName & ", " & valueName & ">"
        Case COLLECTION_SORTED_DICTIONARY
            BuildCollectionTypeName = _
                "SortedDictionary<" & keyName & ", " & valueName & ">"
        Case COLLECTION_SORTED_SET
            BuildCollectionTypeName = "SortedSet<" & valueName & ">"
        Case COLLECTION_ORDERED_DICTIONARY
            BuildCollectionTypeName = _
                "OrderedDictionary<" & keyName & ", " & valueName & ">"
        Case COLLECTION_PRIORITY_QUEUE
            BuildCollectionTypeName = _
                "PriorityQueue<" & valueName & ", " & keyName & ">"
        Case COLLECTION_OBSERVABLE
            BuildCollectionTypeName = "ObservableCollection<" & valueName & ">"
        Case COLLECTION_READ_ONLY
            BuildCollectionTypeName = "ReadOnlyCollection<" & valueName & ">"
        Case COLLECTION_KEYED
            BuildCollectionTypeName = _
                "KeyedCollection<" & keyName & ", " & valueName & ">"
        Case COLLECTION_LOOKUP
            BuildCollectionTypeName = _
                "Lookup<" & keyName & ", " & valueName & ">"
        Case COLLECTION_GROUPING
            BuildCollectionTypeName = _
                "IGrouping<" & keyName & ", " & valueName & ">"
        Case COLLECTION_CONCURRENT_DICTIONARY
            BuildCollectionTypeName = _
                "ConcurrentDictionary<" & keyName & ", " & valueName & ">"
        Case COLLECTION_CONCURRENT_QUEUE
            BuildCollectionTypeName = "ConcurrentQueue<" & valueName & ">"
        Case COLLECTION_CONCURRENT_STACK
            BuildCollectionTypeName = "ConcurrentStack<" & valueName & ">"
        Case COLLECTION_CONCURRENT_BAG
            BuildCollectionTypeName = "ConcurrentBag<" & valueName & ">"
        Case COLLECTION_BLOCKING
            BuildCollectionTypeName = "BlockingCollection<" & valueName & ">"
        Case Else
            BuildCollectionTypeName = BuildImmutableTypeName(keyName, valueName)
    End Select
End Function

Private Function BuildImmutableTypeName( _
    ByVal keyName As String, _
    ByVal valueName As String _
) As String
    Select Case mCollectionKind
        Case COLLECTION_IMMUTABLE_ARRAY
            BuildImmutableTypeName = "ImmutableArray<" & valueName & ">"
        Case COLLECTION_IMMUTABLE_LIST
            BuildImmutableTypeName = "ImmutableList<" & valueName & ">"
        Case COLLECTION_IMMUTABLE_DICTIONARY
            BuildImmutableTypeName = _
                "ImmutableDictionary<" & keyName & ", " & valueName & ">"
        Case COLLECTION_IMMUTABLE_HASH_SET
            BuildImmutableTypeName = "ImmutableHashSet<" & valueName & ">"
        Case COLLECTION_IMMUTABLE_QUEUE
            BuildImmutableTypeName = "ImmutableQueue<" & valueName & ">"
        Case COLLECTION_IMMUTABLE_STACK
            BuildImmutableTypeName = "ImmutableStack<" & valueName & ">"
        Case COLLECTION_IMMUTABLE_SORTED_DICTIONARY
            BuildImmutableTypeName = _
                "ImmutableSortedDictionary<" & keyName & ", " & valueName & ">"
        Case COLLECTION_IMMUTABLE_SORTED_SET
            BuildImmutableTypeName = "ImmutableSortedSet<" & valueName & ">"
    End Select
End Function

Friend Sub AddMutable( _
    ByVal keyOrValue As Variant, _
    Optional ByVal value As Variant, _
    Optional ByVal memberName As String = "Add" _
)
    Dim keyWrapped As ROneCOne
    Dim valueWrapped As ROneCOne

    RequireMutableCollection memberName
    If IsDictionaryCollection Then
        If IsMissing(value) Then
            RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
                "A dictionary value is required."
        End If
        Set keyWrapped = WrapKey(keyOrValue, memberName)
        If FindKeyIndex(keyWrapped, memberName) >= 0 Then
            RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
                "An item with the same key already exists."
        End If
        AddDictionaryPair keyWrapped, WrapElement(value, memberName), memberName
        Exit Sub
    End If
    If IsSetCollection Then
        If Not IsMissing(value) Then
            RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
                "A set accepts one value."
        End If
        Set valueWrapped = WrapElement(keyOrValue, memberName)
        If FindElementIndex(valueWrapped, memberName) < 0 Then
            AddSetValue valueWrapped, memberName
        End If
        Exit Sub
    End If
    If mCollectionKind = COLLECTION_KEYED Then
        If Not IsMissing(value) Then
            RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
                "A keyed collection accepts one item."
        End If
        Set valueWrapped = WrapElement(keyOrValue, memberName)
        Set keyWrapped = InvokeUnary(mKeySelector, valueWrapped)
        Set keyWrapped = WidenScalarToward(keyWrapped, mKeyType)
        ValidateWrappedType keyWrapped, mKeyType, mKeyObjectType, _
            memberName, "key"
        If FindKeyIndex(keyWrapped, memberName) >= 0 Then
            RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
                "An item with the same key already exists."
        End If
        PrepareHashInsert
        ArrAppend mKeys, mKeysCount, keyWrapped
        ArrAppend mItems, mItemsCount, valueWrapped
        AddHashIndexEntry keyWrapped, mKeysCount - 1, valueWrapped
    ElseIf IsPriorityQueueKind(mCollectionKind) Then
        If IsMissing(value) Then
            RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
                "A priority value is required."
        End If
        AddPriorityValue WrapElement(keyOrValue, memberName), _
            WrapKey(value, memberName)
        Exit Sub
    ElseIf IsStackKind(mCollectionKind) Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            "Use Push for a stack."
    Else
        If Not IsMissing(value) Then
            RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
                "This collection accepts one item."
        End If
        If mCollectionKind = COLLECTION_BLOCKING And mAddingCompleted Then
            RaiseContractError ERROR_INVALID_OPERATION, memberName, _
                "Adding has already been completed."
        End If
        If mCollectionKind = COLLECTION_BLOCKING And mQueryCount >= 0 Then
            If mItemsCount >= mQueryCount Then
                RaiseContractError ERROR_INVALID_OPERATION, memberName, _
                    "The bounded collection is full."
            End If
        End If
        ArrAppend mItems, mItemsCount, WrapElement(keyOrValue, memberName)
    End If
    mVersion = mVersion + 1
    NotifyCollectionChanged "Add"
End Sub

Friend Sub AddDictionaryPair( _
    ByVal keyWrapped As ROneCOne, _
    ByVal valueWrapped As ROneCOne, _
    ByVal memberName As String _
)
    Dim insertAt As Long

    If IsSortedMapKind(mCollectionKind) Then
        insertAt = FindSortedInsertIndex(mKeys, mKeysCount, keyWrapped, memberName)
        If insertAt <= mKeysCount Then
            ArrInsert mKeys, mKeysCount, insertAt, keyWrapped
            ArrInsert mItems, mItemsCount, insertAt, valueWrapped
        Else
            ArrAppend mKeys, mKeysCount, keyWrapped
            ArrAppend mItems, mItemsCount, valueWrapped
        End If
    Else
        PrepareHashInsert
        ArrAppend mKeys, mKeysCount, keyWrapped
        ArrAppend mItems, mItemsCount, valueWrapped
        AddHashIndexEntry keyWrapped, mKeysCount - 1, valueWrapped
    End If
    mVersion = mVersion + 1
End Sub

Friend Sub AddSetValue( _
    ByVal wrapped As ROneCOne, _
    ByVal memberName As String _
)
    Dim insertAt As Long

    If IsSortedSetKind(mCollectionKind) Then
        insertAt = FindSortedInsertIndex(mItems, mItemsCount, wrapped, memberName)
        If insertAt <= mItemsCount Then
            ArrInsert mItems, mItemsCount, insertAt, wrapped
        Else
            ArrAppend mItems, mItemsCount, wrapped
        End If
    Else
        PrepareHashInsert
        ArrAppend mItems, mItemsCount, wrapped
        AddHashIndexEntry wrapped, mItemsCount - 1
    End If
    mVersion = mVersion + 1
End Sub

Friend Sub AddPriorityValue( _
    ByVal wrapped As ROneCOne, _
    ByVal priorityWrapped As ROneCOne _
)
    Dim insertAt As Long

    insertAt = FindSortedInsertIndex(mPriorities, mPrioritiesCount, priorityWrapped, "Enqueue")
    If insertAt <= mPrioritiesCount Then
        ArrInsert mPriorities, mPrioritiesCount, insertAt, priorityWrapped
        ArrInsert mItems, mItemsCount, insertAt, wrapped
    Else
        ArrAppend mPriorities, mPrioritiesCount, priorityWrapped
        ArrAppend mItems, mItemsCount, wrapped
    End If
    mVersion = mVersion + 1
End Sub

Private Function IsSortedMapKind(ByVal collectionKind As Long) As Boolean
    Select Case collectionKind
        Case COLLECTION_SORTED_LIST, COLLECTION_SORTED_DICTIONARY
            IsSortedMapKind = True
        Case COLLECTION_IMMUTABLE_SORTED_DICTIONARY
            IsSortedMapKind = True
    End Select
End Function

Private Function IsSortedSetKind(ByVal collectionKind As Long) As Boolean
    IsSortedSetKind = (collectionKind = COLLECTION_SORTED_SET Or _
        collectionKind = COLLECTION_IMMUTABLE_SORTED_SET)
End Function

Friend Function HashIndexSupported() As Boolean
    If Not mQueryComparer Is Nothing Then Exit Function
    Select Case mCollectionKind
        Case COLLECTION_DICTIONARY, COLLECTION_HASH_SET
            HashIndexSupported = True
        Case COLLECTION_ORDERED_DICTIONARY, COLLECTION_KEYED
            HashIndexSupported = True
        Case COLLECTION_CONCURRENT_DICTIONARY
            HashIndexSupported = True
        Case COLLECTION_IMMUTABLE_DICTIONARY
            HashIndexSupported = True
        Case COLLECTION_IMMUTABLE_HASH_SET, COLLECTION_LOOKUP
            HashIndexSupported = True
    End Select
End Function

Private Function DirectKeyLookupSupported() As Boolean
    Select Case mKeyType
        Case vbByte, vbInteger, vbLong, vbLongLong, vbBoolean
            DirectKeyLookupSupported = True
    End Select
End Function

Private Sub RequireHashCapacityCollection(ByVal memberName As String)
    If HashIndexSupported Then Exit Sub
    RaiseContractError ERROR_INVALID_OPERATION, memberName, _
        "This collection does not use the default hash index."
End Sub

Private Function NextHashCapacity(ByVal minimum As Long) As Long
    Dim result As Long

    result = 8
    Do While result < minimum
        If result > 536870912 Then
            RaiseContractError ERROR_INVALID_ARGUMENT, "EnsureCapacity", _
                "The requested capacity is too large."
        End If
        result = result * 2
    Loop
    NextHashCapacity = result
End Function

Private Sub ResetHashIndex()
    Erase mHashSlots
    Erase mHashCodes
    Erase mHashKeys
    Erase mHashValues
    mHashCapacity = 0
    mHashCount = 0
    mHashDirty = False
End Sub

' Element and key storage uses doubling dynamic arrays rather than a linked
' Collection, so a numeric indexer is O(1) and a bulk build stays linear. The
' helpers below keep one growable array (slots 1..count) with amortized append.

Private Sub ArrReset(ByRef arr() As ROneCOne, ByRef count As Long)
    ReDim arr(1 To 8)
    count = 0
End Sub

Private Sub ArrAppend( _
    ByRef arr() As ROneCOne, _
    ByRef count As Long, _
    ByVal item As ROneCOne _
)
    If count >= UBound(arr) Then ReDim Preserve arr(1 To UBound(arr) * 2)
    count = count + 1
    Set arr(count) = item
End Sub

Private Sub ArrInsert( _
    ByRef arr() As ROneCOne, _
    ByRef count As Long, _
    ByVal position As Long, _
    ByVal item As ROneCOne _
)
    Dim shiftIndex As Long

    If count >= UBound(arr) Then ReDim Preserve arr(1 To UBound(arr) * 2)
    For shiftIndex = count To position Step -1
        Set arr(shiftIndex + 1) = arr(shiftIndex)
    Next shiftIndex
    Set arr(position) = item
    count = count + 1
End Sub

Private Sub ArrRemoveAt( _
    ByRef arr() As ROneCOne, _
    ByRef count As Long, _
    ByVal position As Long _
)
    Dim shiftIndex As Long

    For shiftIndex = position To count - 1
        Set arr(shiftIndex) = arr(shiftIndex + 1)
    Next shiftIndex
    Set arr(count) = Nothing
    count = count - 1
End Sub

Private Sub ArrReplaceAt( _
    ByRef arr() As ROneCOne, _
    ByVal position As Long, _
    ByVal item As ROneCOne _
)
    Set arr(position) = item
End Sub

Private Function ArrSnapshot( _
    ByRef arr() As ROneCOne, _
    ByVal count As Long _
) As Collection
    Dim index As Long
    Dim result As Collection

    Set result = New Collection
    For index = 1 To count
        result.Add arr(index)
    Next index
    Set ArrSnapshot = result
End Function

Friend Sub RebuildHashIndex(ByVal requestedCapacity As Long)
    Dim index As Long
    Dim keyed As Boolean
    Dim minimum As Long

    If Not HashIndexSupported Then
        ResetHashIndex
        Exit Sub
    End If
    mHashDirty = False
    If requestedCapacity < 8 Then requestedCapacity = 8
    minimum = mItemsCount * 2
    If requestedCapacity < minimum Then requestedCapacity = minimum
    mHashCapacity = NextHashCapacity(requestedCapacity)
    ReDim mHashSlots(0 To mHashCapacity - 1)
    ReDim mHashCodes(0 To mHashCapacity - 1)
    ReDim mHashKeys(0 To mHashCapacity - 1)
    ReDim mHashValues(0 To mHashCapacity - 1)
    mHashCount = 0
    keyed = (IsDictionaryCollection Or mCollectionKind = COLLECTION_KEYED Or _
        mCollectionKind = COLLECTION_LOOKUP)
    If keyed Then
        For index = 1 To mKeysCount
            InsertHashSlot mKeys(index), index - 1, mItems(index)
        Next index
    Else
        For index = 1 To mItemsCount
            InsertHashSlot mItems(index), index - 1, mItems(index)
        Next index
    End If
End Sub

' Mutations mark the hash index dirty instead of rebuilding it per operation;
' the rebuild happens once, at the next probe or insert. Every path that reads
' or writes a slot must therefore pass through EnsureHashIndexCurrent first.

Private Sub MarkHashIndexDirty()
    If HashIndexSupported Then mHashDirty = True
End Sub

Private Sub EnsureHashIndexCurrent()
    If Not mHashDirty Then Exit Sub
    mHashDirty = False
    RebuildHashIndex mHashCapacity
End Sub

Private Sub RefreshHashSlotValue( _
    ByVal keyWrapped As ROneCOne, _
    ByVal valueWrapped As ROneCOne, _
    ByVal memberName As String _
)
    Dim slot As Long

    slot = FindHashSlot(keyWrapped, memberName)
    If slot < 0 Then
        MarkHashIndexDirty
        Exit Sub
    End If
    Set mHashValues(slot) = valueWrapped
End Sub

Private Sub ReplaceKeyedValueAtSlot( _
    ByVal itemIndex As Long, _
    ByVal slot As Long, _
    ByVal wrapped As ROneCOne _
)
    ' The caller's probe already located the slot, so a keyed value update is
    ' one array write and one slot write with no second probe.
    ArrReplaceAt mItems, itemIndex + 1, wrapped
    Set mHashValues(slot) = wrapped
    mVersion = mVersion + 1
    NotifyCollectionChanged "Replace"
End Sub

Private Sub RemoveHashIndexEntry(ByVal index As Long)
    Dim slot As Long
    Dim wrapped As ROneCOne

    If Not HashIndexSupported Then Exit Sub
    If mHashDirty Then Exit Sub
    If IsDictionaryCollection Or mCollectionKind = COLLECTION_KEYED Or _
        mCollectionKind = COLLECTION_LOOKUP Then
        If index + 1 > mKeysCount Then
            MarkHashIndexDirty
            Exit Sub
        End If
        Set wrapped = mKeys(index + 1)
    Else
        Set wrapped = mItems(index + 1)
    End If
    slot = FindHashSlot(wrapped, "Remove")
    If slot < 0 Then
        MarkHashIndexDirty
        Exit Sub
    End If
    DeleteHashSlot slot
    If index + 1 < mItemsCount Then ShiftHashIndexesAbove index
End Sub

Private Sub DeleteHashSlot(ByVal slot As Long)
    Dim home As Long
    Dim mask As Long
    Dim nextSlot As Long

    mask = mHashCapacity - 1
    mHashSlots(slot) = 0
    mHashCodes(slot) = 0
    Set mHashKeys(slot) = Nothing
    Set mHashValues(slot) = Nothing
    mHashCount = mHashCount - 1
    ' Linear probing cannot leave a gap inside a probe cluster. Walk the
    ' cluster and slide back every entry whose own probe path crosses the
    ' freed slot, then continue from the slot that entry vacated.
    nextSlot = (slot + 1) And mask
    Do While mHashSlots(nextSlot) <> 0
        home = mHashCodes(nextSlot) And mask
        If ((nextSlot - home) And mask) >= ((nextSlot - slot) And mask) Then
            mHashSlots(slot) = mHashSlots(nextSlot)
            mHashCodes(slot) = mHashCodes(nextSlot)
            Set mHashKeys(slot) = mHashKeys(nextSlot)
            Set mHashValues(slot) = mHashValues(nextSlot)
            mHashSlots(nextSlot) = 0
            mHashCodes(nextSlot) = 0
            Set mHashKeys(nextSlot) = Nothing
            Set mHashValues(nextSlot) = Nothing
            slot = nextSlot
        End If
        nextSlot = (nextSlot + 1) And mask
    Loop
End Sub

Private Sub ShiftHashIndexesAbove(ByVal index As Long)
    Dim slot As Long

    For slot = 0 To mHashCapacity - 1
        If mHashSlots(slot) > index + 1 Then
            mHashSlots(slot) = mHashSlots(slot) - 1
        End If
    Next slot
End Sub

Private Sub ResizeHashIndex(ByVal requestedCapacity As Long)
    Dim index As Long
    Dim savedCount As Long
    Dim savedIndices() As Long
    Dim savedKeys() As ROneCOne
    Dim savedValues() As ROneCOne
    Dim slot As Long

    If Not HashIndexSupported Then Exit Sub
    EnsureHashIndexCurrent
    If mHashCount = 0 Then
        RebuildHashIndex requestedCapacity
        Exit Sub
    End If
    ReDim savedIndices(1 To mHashCount)
    ReDim savedKeys(1 To mHashCount)
    ReDim savedValues(1 To mHashCount)
    For slot = 0 To mHashCapacity - 1
        If mHashSlots(slot) <> 0 Then
            savedCount = savedCount + 1
            savedIndices(savedCount) = mHashSlots(slot) - 1
            Set savedKeys(savedCount) = mHashKeys(slot)
            Set savedValues(savedCount) = mHashValues(slot)
        End If
    Next slot
    mHashCapacity = NextHashCapacity(requestedCapacity)
    ReDim mHashSlots(0 To mHashCapacity - 1)
    ReDim mHashCodes(0 To mHashCapacity - 1)
    ReDim mHashKeys(0 To mHashCapacity - 1)
    ReDim mHashValues(0 To mHashCapacity - 1)
    mHashCount = 0
    For index = 1 To savedCount
        InsertHashSlot savedKeys(index), savedIndices(index), savedValues(index)
    Next index
End Sub

Private Sub PrepareHashInsert()
    If Not HashIndexSupported Then Exit Sub
    EnsureHashIndexCurrent
    If mHashCapacity = 0 Then
        RebuildHashIndex 8
    ElseIf (mHashCount + 1) * 10 >= mHashCapacity * 7 Then
        ResizeHashIndex mHashCapacity * 2
    End If
End Sub

Private Sub AddHashIndexEntry( _
    ByVal wrapped As ROneCOne, _
    ByVal canonicalIndex As Long, _
    Optional ByVal valueWrapped As ROneCOne _
)
    If Not HashIndexSupported Then Exit Sub
    If valueWrapped Is Nothing Then Set valueWrapped = wrapped
    InsertHashSlot wrapped, canonicalIndex, valueWrapped
End Sub

Private Sub InsertHashSlot( _
    ByVal wrapped As ROneCOne, _
    ByVal canonicalIndex As Long, _
    ByVal valueWrapped As ROneCOne _
)
    Dim hashCode As Long
    Dim slot As Long

    hashCode = WrappedHashCode(wrapped)
    slot = hashCode And (mHashCapacity - 1)
    Do While mHashSlots(slot) <> 0
        slot = (slot + 1) And (mHashCapacity - 1)
    Loop
    mHashCodes(slot) = hashCode
    mHashSlots(slot) = canonicalIndex + 1
    Set mHashKeys(slot) = wrapped
    Set mHashValues(slot) = valueWrapped
    mHashCount = mHashCount + 1
End Sub

Private Function FindHashSlot( _
    ByVal candidate As ROneCOne, _
    ByVal memberName As String _
) As Long
    Dim hashCode As Long
    Dim probes As Long
    Dim slot As Long

    EnsureHashIndexCurrent
    FindHashSlot = -1
    If mHashCapacity = 0 Then Exit Function
    hashCode = WrappedHashCode(candidate)
    slot = hashCode And (mHashCapacity - 1)
    Do While mHashSlots(slot) <> 0 And probes < mHashCapacity
        If mHashCodes(slot) = hashCode Then
            If ValuesEqualWithComparer(mHashKeys(slot), _
                candidate, Nothing, memberName) Then
                FindHashSlot = slot
                Exit Function
            End If
        End If
        slot = (slot + 1) And (mHashCapacity - 1)
        probes = probes + 1
    Loop
End Function

Private Function FindHashIndex( _
    ByVal candidate As ROneCOne, _
    ByVal memberName As String _
) As Long
    Dim slot As Long

    slot = FindHashSlot(candidate, memberName)
    If slot < 0 Then
        FindHashIndex = -1
    Else
        FindHashIndex = mHashSlots(slot) - 1
    End If
End Function

Private Function FindHashValue( _
    ByVal candidate As ROneCOne, _
    ByVal memberName As String _
) As ROneCOne
    Dim slot As Long

    slot = FindHashSlot(candidate, memberName)
    If slot >= 0 Then Set FindHashValue = mHashValues(slot)
End Function

Private Function FindHashSlotByKey( _
    ByVal candidate As Variant, _
    ByVal memberName As String _
) As Long
    Dim hashCode As Long
    Dim probes As Long
    Dim slot As Long

    EnsureHashIndexCurrent
    FindHashSlotByKey = -1
    If mHashCapacity = 0 Then Exit Function
    hashCode = ValueHashCode(candidate)
    slot = hashCode And (mHashCapacity - 1)
    Do While mHashSlots(slot) <> 0 And probes < mHashCapacity
        If mHashCodes(slot) = hashCode Then
            If WrappedEqualsValue(mHashKeys(slot), candidate) Then
                FindHashSlotByKey = slot
                Exit Function
            End If
        End If
        slot = (slot + 1) And (mHashCapacity - 1)
        probes = probes + 1
    Loop
End Function

Private Function FindHashValueByKey( _
    ByVal candidate As Variant, _
    ByVal memberName As String _
) As ROneCOne
    Dim slot As Long

    slot = FindHashSlotByKey(candidate, memberName)
    If slot >= 0 Then Set FindHashValueByKey = mHashValues(slot)
End Function

Private Function WrappedHashCode(ByVal wrapped As ROneCOne) As Long
    If wrapped.InternalHasObjectValue Then
        WrappedHashCode = ObjectHashCode(wrapped.InternalObjectValue)
    Else
        WrappedHashCode = ValueHashCode(wrapped.InternalScalarValue)
    End If
End Function

Private Function ValueHashCode(ByVal scalarValue As Variant) As Long
    Dim scalarType As VbVarType

    If IsObject(scalarValue) Then
        ValueHashCode = ObjectHashCode(scalarValue)
        Exit Function
    End If
    scalarType = VarTypeOf(scalarValue)
    If IsNull(scalarValue) Then
        ValueHashCode = HashText("NULL")
    ElseIf IsEmpty(scalarValue) Then
        ValueHashCode = 0
    ElseIf scalarType = vbByte Or scalarType = vbInteger Or _
        scalarType = vbLong Or scalarType = vbBoolean Then
        ValueHashCode = CLng(scalarValue) And &H7FFFFFFF
    ElseIf IsNumeric(scalarValue) Or scalarType = vbDate Then
        ValueHashCode = NumericHashCode(scalarValue)
    Else
        ValueHashCode = HashText(CStr(VarTypeOf(scalarValue)) & _
            ":" & CStr(scalarValue))
    End If
End Function

Private Function ObjectHashCode(ByVal value As Object) As Long
    Dim pointerValue As LongLong

    pointerValue = ObjPtr(value)
    ObjectHashCode = CLng(pointerValue And &H7FFFFFFF)
End Function

Private Function WrappedEqualsValue( _
    ByVal wrapped As ROneCOne, _
    ByVal candidate As Variant _
) As Boolean
    Dim candidateObject As Object

    If wrapped.InternalHasObjectValue Or IsObject(candidate) Then
        If Not wrapped.InternalHasObjectValue Or Not IsObject(candidate) Then Exit Function
        Set candidateObject = candidate
        WrappedEqualsValue = wrapped.InternalObjectValue Is candidateObject
    ElseIf IsNull(wrapped.InternalScalarValue) Or IsNull(candidate) Then
        WrappedEqualsValue = IsNull(wrapped.InternalScalarValue) And IsNull(candidate)
    Else
        WrappedEqualsValue = (wrapped.InternalScalarValue = candidate)
    End If
End Function

Private Function NumericHashCode(ByVal scalarValue As Variant) As Long
    Dim decimalValue As Variant

    If VarTypeOf(scalarValue) = vbDate Then
        decimalValue = CDec(CDbl(scalarValue))
    Else
        decimalValue = CDec(scalarValue)
    End If
    If decimalValue = Fix(decimalValue) Then
        If decimalValue >= CDec(-2147483648#) And _
            decimalValue <= CDec(2147483647#) Then
            NumericHashCode = CLng(decimalValue) And &H7FFFFFFF
            Exit Function
        End If
    End If
    NumericHashCode = HashText("N:" & CStr(decimalValue))
End Function

Private Function HashText(ByVal value As String) As Long
    Dim hashValue As LongLong
    Dim index As Long

    hashValue = 2166136261^
    For index = 1 To Len(value)
        hashValue = ((hashValue Xor AscW(Mid$(value, index, 1))) * 16777619) _
            Mod 2147483647
    Next index
    HashText = CLng(hashValue And &H7FFFFFFF)
End Function

Private Function FindSortedValueIndex( _
    ByRef values() As ROneCOne, _
    ByVal count As Long, _
    ByVal candidate As ROneCOne, _
    ByVal memberName As String _
) As Long
    Dim position As Long

    FindSortedValueIndex = -1
    position = FindSortedInsertIndex(values, count, candidate, memberName)
    If position > count Then Exit Function
    If CompareWithComparer(values(position), candidate, _
        mQueryComparer, memberName) = 0 Then
        FindSortedValueIndex = position - 1
    End If
End Function

Private Function FindSortedInsertIndex( _
    ByRef values() As ROneCOne, _
    ByVal count As Long, _
    ByVal candidate As ROneCOne, _
    ByVal memberName As String _
) As Long
    Dim high As Long
    Dim low As Long
    Dim middle As Long

    low = 1
    high = count + 1
    Do While low < high
        middle = low + ((high - low) \ 2)
        If middle <= count And _
            CompareWithComparer(values(middle), candidate, _
                mQueryComparer, memberName) < 0 Then
            low = middle + 1
        Else
            high = middle
        End If
    Loop
    FindSortedInsertIndex = low
End Function

Friend Function FindKeyIndex( _
    ByVal candidate As ROneCOne, _
    ByVal memberName As String _
) As Long
    Dim index As Long

    If IsSortedMapKind(mCollectionKind) Then
        FindKeyIndex = FindSortedValueIndex(mKeys, mKeysCount, candidate, memberName)
        Exit Function
    End If
    If HashIndexSupported Then
        FindKeyIndex = FindHashIndex(candidate, memberName)
        Exit Function
    End If

    FindKeyIndex = -1
    For index = 1 To mKeysCount
        If CollectionUsesEquality(mCollectionKind) Then
            If ValuesEqualWithComparer(mKeys(index), candidate, _
                mQueryComparer, memberName) Then
                FindKeyIndex = index - 1
                Exit Function
            End If
        ElseIf CompareWithComparer(mKeys(index), candidate, _
            mQueryComparer, memberName) = 0 Then
            FindKeyIndex = index - 1
            Exit Function
        End If
    Next index
End Function

Friend Function FindElementIndex( _
    ByVal candidate As ROneCOne, _
    ByVal memberName As String _
) As Long
    Dim index As Long

    If IsSortedSetKind(mCollectionKind) Then
        FindElementIndex = FindSortedValueIndex(mItems, mItemsCount, candidate, memberName)
        Exit Function
    End If
    If HashIndexSupported Then
        FindElementIndex = FindHashIndex(candidate, memberName)
        Exit Function
    End If
    ' Custom-comparer collections scan the element array directly instead of
    ' allocating a snapshot Collection for every membership test.
    FindElementIndex = -1
    For index = 1 To mItemsCount
        If CollectionUsesEquality(mCollectionKind) Then
            If ValuesEqualWithComparer(mItems(index), candidate, _
                mQueryComparer, memberName) Then
                FindElementIndex = index - 1
                Exit Function
            End If
        ElseIf CompareWithComparer(mItems(index), candidate, _
            mQueryComparer, memberName) = 0 Then
            FindElementIndex = index - 1
            Exit Function
        End If
    Next index
End Function

Private Function FindWrappedInCollection( _
    ByVal values As Collection, _
    ByVal candidate As ROneCOne, _
    ByVal memberName As String _
) As Long
    Dim index As Long

    FindWrappedInCollection = -1
    For index = 1 To values.Count
        If CollectionUsesEquality(mCollectionKind) Then
            If ValuesEqualWithComparer(values.Item(index), candidate, _
                mQueryComparer, memberName) Then
                FindWrappedInCollection = index - 1
                Exit Function
            End If
        ElseIf CompareWithComparer(values.Item(index), candidate, _
            mQueryComparer, memberName) = 0 Then
            FindWrappedInCollection = index - 1
            Exit Function
        End If
    Next index
End Function

Private Function DistinctWrappedValues( _
    ByVal values As Collection, _
    ByVal memberName As String _
) As Collection
    Dim output As Collection
    Dim wrapped As ROneCOne

    Set output = New Collection
    For Each wrapped In values
        If FindWrappedInCollection(output, wrapped, memberName) < 0 Then
            output.Add wrapped
        End If
    Next wrapped
    Set DistinctWrappedValues = output
End Function

Private Function IsSetSubset( _
    ByVal other As Variant, _
    ByVal proper As Boolean, _
    ByVal memberName As String _
) As Boolean
    Dim incoming As Collection
    Dim uniqueValues As Collection
    Dim wrapped As ROneCOne

    RequireSetCollection memberName
    Set incoming = BuildIncomingItems(other, memberName)
    Set uniqueValues = DistinctWrappedValues(incoming, memberName)
    If proper And mItemsCount >= uniqueValues.Count Then Exit Function
    For Each wrapped In ArrSnapshot(mItems, mItemsCount)
        If FindWrappedInCollection(uniqueValues, wrapped, memberName) < 0 Then
            Exit Function
        End If
    Next wrapped
    IsSetSubset = True
End Function

Private Function IsSetSuperset( _
    ByVal other As Variant, _
    ByVal proper As Boolean, _
    ByVal memberName As String _
) As Boolean
    Dim incoming As Collection
    Dim uniqueValues As Collection
    Dim wrapped As ROneCOne

    RequireSetCollection memberName
    Set incoming = BuildIncomingItems(other, memberName)
    Set uniqueValues = DistinctWrappedValues(incoming, memberName)
    If proper And mItemsCount <= uniqueValues.Count Then Exit Function
    For Each wrapped In uniqueValues
        If FindElementIndex(wrapped, memberName) < 0 Then Exit Function
    Next wrapped
    IsSetSuperset = True
End Function

Private Sub SetItemValue(ByVal index As Variant, ByVal value As Variant)
    Dim itemIndex As Long
    Dim keyWrapped As ROneCOne
    Dim replacement As ROneCOne
    Dim slot As Long

    If mRole = ROLE_DATA_ROW Then
        SetDataRowItem index, value
        Exit Sub
    End If
    If IsDictionaryCollection Then
        RequireMutableCollection "Item"
        Set keyWrapped = WrapKey(index, "Item")
        Set replacement = WrapElement(value, "Item")
        If HashIndexSupported Then
            slot = FindHashSlot(keyWrapped, "Item")
            If slot < 0 Then
                AddDictionaryPair keyWrapped, replacement, "Item"
            Else
                ReplaceKeyedValueAtSlot mHashSlots(slot) - 1, slot, replacement
            End If
            Exit Sub
        End If
        itemIndex = FindKeyIndex(keyWrapped, "Item")
        If itemIndex < 0 Then
            AddDictionaryPair keyWrapped, replacement, "Item"
        Else
            ReplaceCollectionItem itemIndex, replacement
        End If
        Exit Sub
    End If
    If IsGenericCollectionRole(mRole) Then
        RequireMutableCollection "Item"
        itemIndex = CLng(index)
        ValidateIndex itemIndex, mItemsCount, "Item"
        ReplaceCollectionItem itemIndex, WrapElement(value, "Item")
        Exit Sub
    End If
    ReplaceItem CLng(index), WrapElement(value, "Item")
End Sub

Friend Sub ReplaceCollectionItem( _
    ByVal index As Long, _
    ByVal wrapped As ROneCOne _
)
    ArrReplaceAt mItems, index + 1, wrapped
    If HashIndexSupported Then
        If IsDictionaryCollection Or mCollectionKind = COLLECTION_KEYED Or _
            mCollectionKind = COLLECTION_LOOKUP Then
            ' The key is unchanged, so only the stored value pointer for its
            ' slot needs to move; a rebuild would rescan every entry.
            RefreshHashSlotValue mKeys(index + 1), wrapped, "Item"
        Else
            ' The element is its own hash key, so the slot layout changed.
            MarkHashIndexDirty
        End If
    End If
    mVersion = mVersion + 1
    NotifyCollectionChanged "Replace"
End Sub

Private Function RemoveCollectionValue( _
    ByVal target As ROneCOne, _
    ByVal value As Variant, _
    ByVal memberName As String _
) As Boolean
    Dim itemIndex As Long

    If target.IsDictionaryCollection Or _
        target.InternalCollectionKind = COLLECTION_KEYED Then
        itemIndex = target.FindKeyIndex(target.WrapKey(value, memberName), memberName)
    Else
        itemIndex = target.FindElementIndex( _
            target.WrapElement(value, memberName), memberName)
    End If
    If itemIndex < 0 Then Exit Function
    target.RemoveCollectionAt itemIndex
    RemoveCollectionValue = True
End Function

Friend Sub RemoveCollectionAt(ByVal index As Long)
    RemoveHashIndexEntry index
    ArrRemoveAt mItems, mItemsCount, index + 1
    If IsDictionaryCollection Or mCollectionKind = COLLECTION_KEYED Then
        ArrRemoveAt mKeys, mKeysCount, index + 1
    ElseIf IsPriorityQueueKind(mCollectionKind) Then
        ArrRemoveAt mPriorities, mPrioritiesCount, index + 1
    End If
    mVersion = mVersion + 1
    NotifyCollectionChanged "Remove"
End Sub

Private Function CloneGenericCollection( _
    ByVal source As ROneCOne _
) As ROneCOne
    Dim clone As ROneCOne
    Dim wrapped As ROneCOne

    Set clone = New ROneCOne
    clone.ConfigureGenericCollection source.InternalRole, _
        source.InternalCollectionKind, source.InternalValueType, _
        source.InternalValueObjectType, source.InternalKeyType, _
        source.InternalKeyObjectType, source.InternalQueryComparer
    For Each wrapped In source.InternalCollectionItems
        clone.InternalAppendCollectionItem wrapped
    Next wrapped
    For Each wrapped In source.InternalCollectionKeys
        clone.InternalAppendCollectionKey wrapped
    Next wrapped
    For Each wrapped In source.InternalCollectionPriorities
        clone.InternalAppendCollectionPriority wrapped
    Next wrapped
    clone.ConfigureCollectionKeySelector source.InternalKeySelector
    If Not source.InternalCollectionChanged Is Nothing Then
        clone.ConfigureObservableCollection
    End If
    clone.ConfigureReadOnlySource source.InternalSource
    clone.ConfigureCollectionCapacity source.InternalCollectionCapacity
    clone.ConfigureCollectionVersion source.InternalCollectionVersion
    If clone.HashIndexSupported Then clone.RebuildHashIndex 8
    Set CloneGenericCollection = clone
End Function

Private Function CreateCollectionNode(ByVal index As Long) As ROneCOne
    Dim node As ROneCOne

    ValidateIndex index, mItemsCount, "LinkedListNode"
    Set node = New ROneCOne
    node.ConfigureCollectionNode mItems(index + 1), Me
    Set CreateCollectionNode = node
End Function

Friend Function CreateNodeAt(ByVal index As Long) As ROneCOne
    Set CreateNodeAt = CreateCollectionNode(index)
End Function

Friend Function FindNodeWrappedIndex(ByVal wrapped As ROneCOne) As Long
    Dim index As Long

    FindNodeWrappedIndex = -1
    For index = 1 To mItemsCount
        If mItems(index) Is wrapped Then
            FindNodeWrappedIndex = index - 1
            Exit Function
        End If
    Next index
End Function

Private Function InsertRelativeNode( _
    ByVal node As ROneCOne, _
    ByVal value As Variant, _
    ByVal insertAfter As Boolean, _
    ByVal memberName As String _
) As ROneCOne
    Dim index As Long
    Dim insertAt As Long
    Dim wrapped As ROneCOne

    RequireLinkedList memberName
    RequireMutableCollection memberName
    RequireOwnedNode node, memberName
    index = FindNodeWrappedIndex(node.InternalQueryValue)
    If index < 0 Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            "The node no longer belongs to this list."
    End If
    insertAt = index
    If insertAfter Then insertAt = insertAt + 1
    Set wrapped = WrapElement(value, memberName)
    If insertAt >= mItemsCount Then
        ArrAppend mItems, mItemsCount, wrapped
    Else
        ArrInsert mItems, mItemsCount, insertAt + 1, wrapped
    End If
    mVersion = mVersion + 1
    Set InsertRelativeNode = CreateCollectionNode(insertAt)
End Function

Private Function CreateDictionaryEntry( _
    ByVal keyWrapped As ROneCOne, _
    ByVal valueWrapped As ROneCOne _
) As ROneCOne
    Dim entry As ROneCOne

    Set entry = New ROneCOne
    entry.ConfigureDictionaryEntry keyWrapped, valueWrapped
    Set CreateDictionaryEntry = entry
End Function

Private Sub NotifyCollectionChanged(ByVal actionName As String)
    If mCollectionKind <> COLLECTION_OBSERVABLE Then Exit Sub
    If mCollectionChanged Is Nothing Then Exit Sub
    mCollectionChanged.Emit actionName
End Sub

Private Sub AssignReferenceWrapped( _
    ByVal referenceValue As ROneCOne, _
    ByVal wrapped As ROneCOne, _
    ByVal memberName As String _
)
    Dim byteValue As Byte
    Dim booleanValue As Boolean
    Dim currencyValue As Currency
    Dim dateValue As Date
    Dim doubleValue As Double
    Dim integerValue As Integer
    Dim longLongValue As LongLong
    Dim longValue As Long
    Dim singleValue As Single
    Dim scalarValue As Variant

    Set wrapped = WidenScalarToward(wrapped, referenceValue.InternalExpectedType)
    If referenceValue.InternalReferenceAddress = 0 Then
        ValidateWrappedType wrapped, referenceValue.InternalExpectedType, _
            referenceValue.InternalExpectedObjectType, memberName, "output"
        referenceValue.ConfigureManagedReferenceValue wrapped
        Exit Sub
    End If

    If wrapped.InternalHasObjectValue Then
        RaiseContractError ERROR_TYPE_MISMATCH, memberName, _
            "The output reference cannot receive an object."
    End If
    scalarValue = wrapped.InternalScalarValue
    If VarTypeOf(scalarValue) <> referenceValue.InternalExpectedType Then
        RaiseContractError ERROR_TYPE_MISMATCH, memberName, _
            "The output reference type does not match the value type."
    End If
    Select Case referenceValue.InternalExpectedType
        Case vbByte
            byteValue = CByte(scalarValue)
            CopyMemory ByVal referenceValue.InternalReferenceAddress, byteValue, 1
        Case vbInteger
            integerValue = CInt(scalarValue)
            CopyMemory ByVal referenceValue.InternalReferenceAddress, integerValue, 2
        Case vbLong
            longValue = CLng(scalarValue)
            CopyMemory ByVal referenceValue.InternalReferenceAddress, longValue, 4
        Case vbLongLong
            longLongValue = CLngLng(scalarValue)
            CopyMemory ByVal referenceValue.InternalReferenceAddress, longLongValue, 8
        Case vbSingle
            singleValue = CSng(scalarValue)
            CopyMemory ByVal referenceValue.InternalReferenceAddress, singleValue, 4
        Case vbDouble
            doubleValue = CDbl(scalarValue)
            CopyMemory ByVal referenceValue.InternalReferenceAddress, doubleValue, 8
        Case vbCurrency
            currencyValue = CCur(scalarValue)
            CopyMemory ByVal referenceValue.InternalReferenceAddress, currencyValue, 8
        Case vbDate
            dateValue = CDate(scalarValue)
            CopyMemory ByVal referenceValue.InternalReferenceAddress, dateValue, 8
        Case vbBoolean
            booleanValue = CBool(scalarValue)
            CopyMemory ByVal referenceValue.InternalReferenceAddress, booleanValue, 2
    End Select
End Sub

Private Sub AssignDefaultReferenceResult(ByRef destination As Variant)
    If mExpectedType = vbObject Then
        Set destination = Nothing
        Exit Sub
    End If
    Select Case mExpectedType
        Case vbString
            destination = vbNullString
        Case vbBoolean
            destination = False
        Case vbDate
            destination = CDate(0)
        Case vbByte
            destination = CByte(0)
        Case vbInteger
            destination = CInt(0)
        Case vbLong
            destination = CLng(0)
        Case vbLongLong
            destination = CLngLng(0)
        Case vbSingle
            destination = CSng(0)
        Case vbDouble
            destination = CDbl(0)
        Case vbCurrency
            destination = CCur(0)
        Case Else
            destination = Empty
    End Select
End Sub

Private Function CreateSequenceCollection( _
    ByVal roleValue As Long, _
    ByVal collectionKind As Long, _
    ByVal typeToken As Variant, _
    Optional ByVal comparer As Variant _
) As ROneCOne
    Dim comparerValue As ROneCOne
    Dim elementObjectType As String
    Dim elementType As VbVarType
    Dim result As ROneCOne

    ResolveTypeDescriptor typeToken, elementType, elementObjectType, _
        "Collection factory"
    If Not IsMissing(comparer) Then
        Set comparerValue = NormalizeComparer( _
            comparer, CollectionUsesEquality(collectionKind), _
            "Collection comparer")
    End If
    Set result = New ROneCOne
    result.ConfigureGenericCollection roleValue, collectionKind, _
        elementType, elementObjectType, vbEmpty, vbNullString, comparerValue
    Set CreateSequenceCollection = result
End Function

Private Function CreateMapCollection( _
    ByVal roleValue As Long, _
    ByVal collectionKind As Long, _
    ByVal keyTypeToken As Variant, _
    ByVal valueTypeToken As Variant, _
    Optional ByVal comparer As Variant _
) As ROneCOne
    Dim comparerValue As ROneCOne
    Dim keyObjectType As String
    Dim keyType As VbVarType
    Dim result As ROneCOne
    Dim valueObjectType As String
    Dim valueType As VbVarType

    ResolveTypeDescriptor keyTypeToken, keyType, keyObjectType, _
        "Collection key type"
    ResolveTypeDescriptor valueTypeToken, valueType, valueObjectType, _
        "Collection value type"
    If Not IsMissing(comparer) Then
        Set comparerValue = NormalizeComparer( _
            comparer, CollectionUsesEquality(collectionKind), _
            "Collection comparer")
    End If
    Set result = New ROneCOne
    result.ConfigureGenericCollection roleValue, collectionKind, _
        valueType, valueObjectType, keyType, keyObjectType, comparerValue
    Set CreateMapCollection = result
End Function

Private Function CollectionUsesEquality(ByVal collectionKind As Long) As Boolean
    Select Case collectionKind
        Case COLLECTION_SORTED_LIST, COLLECTION_SORTED_DICTIONARY
        Case COLLECTION_SORTED_SET, COLLECTION_PRIORITY_QUEUE
        Case COLLECTION_IMMUTABLE_SORTED_DICTIONARY
        Case COLLECTION_IMMUTABLE_SORTED_SET
        Case Else
            CollectionUsesEquality = True
    End Select
End Function

Private Sub RequireCompatibleSequence( _
    ByVal other As ROneCOne, _
    ByVal memberName As String _
)
    If InternalElementType <> other.InternalElementType Or _
        InternalElementObjectType <> other.InternalElementObjectType Then
        RaiseContractError ERROR_TYPE_MISMATCH, memberName, _
            "Both sequences must have the same generic element type."
    End If
End Sub

Private Function CreateSetQuery( _
    ByVal queryOperator As Long, _
    ByVal other As Variant, _
    Optional ByVal comparer As Variant, _
    Optional ByVal memberName As String = "Set operator" _
) As ROneCOne
    Dim comparerValue As ROneCOne
    Dim otherSequence As ROneCOne

    Set otherSequence = NormalizeMembershipSequence(other, memberName)
    RequireCompatibleSequence otherSequence, memberName
    If Not IsMissing(comparer) Then
        Set comparerValue = NormalizeComparer(comparer, True, memberName)
    End If
    Set CreateSetQuery = CreateQuery( _
        queryOperator, Nothing, otherSequence, 0, _
        InternalElementType, InternalElementObjectType, comparerValue)
End Function

Private Function CreateSetByQuery( _
    ByVal queryOperator As Long, _
    ByVal other As Variant, _
    ByVal keySelector As Variant, _
    Optional ByVal comparer As Variant, _
    Optional ByVal memberName As String = "SetBy", _
    Optional ByVal requireCompatible As Boolean = False _
) As ROneCOne
    Dim comparerValue As ROneCOne
    Dim otherSequence As ROneCOne
    Dim selectorValue As ROneCOne

    Set selectorValue = NormalizeContextUnary(keySelector, memberName, False)
    Set otherSequence = NormalizeMembershipSequence(other, memberName)
    If requireCompatible Then RequireCompatibleSequence otherSequence, memberName
    If Not IsMissing(comparer) Then
        Set comparerValue = NormalizeComparer(comparer, True, memberName)
    End If
    Set CreateSetByQuery = CreateQuery( _
        queryOperator, selectorValue, otherSequence, 0, _
        InternalElementType, InternalElementObjectType, comparerValue)
End Function

Private Function ExecuteJoin( _
    ByVal inner As Variant, _
    ByVal outerKeySelector As Variant, _
    ByVal innerKeySelector As Variant, _
    ByVal resultSelector As Variant, _
    ByVal grouped As Boolean _
) As ROneCOne
    Dim group As ROneCOne
    Dim groupIndex As Long
    Dim innerSequence As ROneCOne
    Dim innerValue As ROneCOne
    Dim key As ROneCOne
    Dim lookup As ROneCOne
    Dim outerKey As ROneCOne
    Dim outerValue As ROneCOne
    Dim outerValues As Collection
    Dim result As ROneCOne
    Dim resultSelectorValue As ROneCOne
    Dim selected As ROneCOne

    Set innerSequence = NormalizeMembershipSequence(inner, "Join")
    Set outerKey = NormalizeContextUnary(outerKeySelector, "Join", False)
    Set lookup = innerSequence.ToLookup(innerKeySelector)
    Set resultSelectorValue = NormalizeBinaryDelegate(resultSelector, "Join")
    Set result = ROneCOne.ListOf(vbVariant)
    Set outerValues = MaterializeWrapped
    For Each outerValue In outerValues
        Set key = InvokeUnary(outerKey, outerValue)
        groupIndex = lookup.FindKeyIndex(key, "Join")
        If grouped Then
            If groupIndex < 0 Then
                Set group = New ROneCOne
                group.ConfigureEmptyGrouping key
            Else
                Set group = lookup.InternalCollectionItem(groupIndex)
            End If
            Set selected = InvokeBinary(resultSelectorValue, outerValue, _
                CreateObjectResult(group))
            result.InternalAddWrapped selected
        ElseIf groupIndex >= 0 Then
            Set group = lookup.InternalCollectionItem(groupIndex)
            For Each innerValue In group.MaterializeWrapped
                Set selected = InvokeBinary( _
                    resultSelectorValue, outerValue, innerValue)
                result.InternalAddWrapped selected
            Next innerValue
        End If
    Next outerValue
    Set ExecuteJoin = result
End Function

Private Function CreateQuery( _
    ByVal queryOperator As Long, _
    ByVal queryDelegate As ROneCOne, _
    ByVal queryValue As ROneCOne, _
    ByVal queryCount As Long, _
    ByVal elementType As VbVarType, _
    ByVal elementObjectType As String, _
    Optional ByVal queryComparer As ROneCOne _
) As ROneCOne
    Dim query As ROneCOne

    RequireSequence Me, "LINQ"
    Set query = New ROneCOne
    query.ConfigureQuery Me, queryOperator, queryDelegate, queryValue, _
        queryCount, elementType, elementObjectType, queryComparer
    Set CreateQuery = query
End Function

Private Function CreatePrimaryOrderQuery( _
    ByVal queryOperator As Long, _
    ByVal queryDelegate As ROneCOne, _
    ByVal elementType As VbVarType, _
    ByVal elementObjectType As String, _
    Optional ByVal queryComparer As ROneCOne _
) As ROneCOne
    Dim baseSource As ROneCOne
    Dim query As ROneCOne

    RequireSequence Me, "Order"
    Set baseSource = OrderingBaseSource
    Set query = New ROneCOne
    query.ConfigureQuery baseSource, queryOperator, queryDelegate, Nothing, 0, _
        elementType, elementObjectType, queryComparer
    Set CreatePrimaryOrderQuery = query
End Function

Private Function OrderingBaseSource() As ROneCOne
    Dim sourceValue As ROneCOne

    Set sourceValue = Me
    Do While sourceValue.InternalRole = ROLE_QUERY And _
        IsOrderingOperator(sourceValue.InternalOperator)
        Set sourceValue = sourceValue.InternalSource
    Loop
    Set OrderingBaseSource = sourceValue
End Function

Private Sub ResolveTypeDescriptor( _
    ByVal typeToken As Variant, _
    ByRef elementType As VbVarType, _
    ByRef elementObjectType As String, _
    ByVal memberName As String _
)
    Dim prototype As Object

    If IsObject(typeToken) Then
        Set prototype = typeToken
        If prototype Is Nothing Then
            RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
                "The user-defined type prototype cannot be Nothing."
        End If
        elementType = vbObject
        elementObjectType = TypeName(prototype)
    Else
        ValidateTypeToken typeToken, memberName
        elementType = CLng(typeToken)
        elementObjectType = vbNullString
    End If
End Sub

Private Sub ValidateTypeToken(ByVal typeToken As Variant, ByVal memberName As String)
    Dim token As Long

    If Not IsNumeric(typeToken) Then
        RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
            "The type token must be a VbVarType value or an object prototype."
    End If
    token = CLng(typeToken)
    Select Case token
        Case vbEmpty, vbNull, vbInteger, vbLong, vbSingle, vbDouble
        Case vbCurrency, vbDate, vbString, vbObject, vbError, vbBoolean
        Case vbVariant, vbDecimal, vbByte, vbLongLong
        Case Else
            RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
                "The type token is not a supported VbVarType value."
    End Select
End Sub

Friend Function WrapElement( _
    ByVal candidate As Variant, _
    ByVal memberName As String _
) As ROneCOne
    Dim candidateObject As Object
    Dim wrapped As ROneCOne

    RequireSequence Me, memberName
    Set wrapped = New ROneCOne
    If IsObject(candidate) Then
        Set candidateObject = candidate
        wrapped.ConfigureObjectValue candidateObject
    Else
        wrapped.ConfigureScalarValue candidate
    End If
    Set wrapped = WidenScalarToward(wrapped, mElementType)
    ValidateElementWrapped wrapped, memberName
    Set WrapElement = wrapped
End Function

Private Sub CommitInitialItems( _
    ByVal listValue As ROneCOne, _
    ByVal wrappedItems As Collection _
)
    Dim wrapped As ROneCOne

    For Each wrapped In wrappedItems
        listValue.InternalAddWrapped wrapped
    Next wrapped
End Sub

Private Function BuildIncomingItems( _
    ByVal values As Variant, _
    ByVal memberName As String _
) As Collection
    Dim candidate As Object
    Dim incoming As Collection
    Dim rawValue As Variant
    Dim sequence As ROneCOne

    Set incoming = New Collection
    If IsObject(values) Then
        Set candidate = values
        If candidate Is Nothing Then
            RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
                "The incoming values cannot be Nothing."
        End If
        If TypeOf candidate Is ROneCOne Then
            Set sequence = candidate
            RequireSequence sequence, memberName
            Set BuildIncomingItems = sequence.MaterializeWrapped
            Exit Function
        End If
        If TypeName(candidate) = "Collection" Then
            For Each rawValue In candidate
                incoming.Add WrapElement(rawValue, memberName)
            Next rawValue
            Set BuildIncomingItems = incoming
            Exit Function
        End If
    ElseIf IsArrayValue(values) Then
        For Each rawValue In values
            incoming.Add WrapElement(rawValue, memberName)
        Next rawValue
        Set BuildIncomingItems = incoming
        Exit Function
    End If
    RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
        "Expected a ROneCOne sequence, VBA array, or Collection."
End Function

Private Sub ValidateElementWrapped( _
    ByVal wrapped As ROneCOne, _
    ByVal memberName As String _
)
    Dim actualObject As Object

    If mElementType = vbVariant Then Exit Sub
    If mElementType = vbObject Then
        If Not wrapped.InternalHasObjectValue Then
            RaiseElementTypeMismatch wrapped, memberName
        End If
        If Len(mElementObjectType) = 0 Then Exit Sub
        Set actualObject = wrapped.InternalObjectValue
        If actualObject Is Nothing Then Exit Sub
        If StrComp(TypeName(actualObject), mElementObjectType, vbBinaryCompare) <> 0 Then
            RaiseElementTypeMismatch wrapped, memberName
        End If
        Exit Sub
    End If

    If wrapped.InternalHasObjectValue Then
        RaiseElementTypeMismatch wrapped, memberName
    End If
    If VarTypeOf(wrapped.InternalScalarValue) <> mElementType Then
        RaiseElementTypeMismatch wrapped, memberName
    End If
End Sub

Private Sub RaiseElementTypeMismatch( _
    ByVal wrapped As ROneCOne, _
    ByVal memberName As String _
)
    Dim actualName As String
    Dim actualObject As Object
    Dim expectedName As String

    If Len(mElementObjectType) <> 0 Then
        expectedName = mElementObjectType
    Else
        expectedName = TypeNameForVarType(mElementType)
    End If
    If wrapped.InternalHasObjectValue Then
        Set actualObject = wrapped.InternalObjectValue
        If actualObject Is Nothing Then
            actualName = "Nothing"
        Else
            actualName = TypeName(actualObject)
        End If
    Else
        actualName = TypeName(wrapped.InternalScalarValue)
    End If
    RaiseContractError ERROR_TYPE_MISMATCH, memberName, _
        "List<" & expectedName & "> cannot contain " & actualName & "."
End Sub

Private Sub RequireSequence(ByVal candidate As ROneCOne, ByVal context As String)
    If candidate Is Nothing Then
        RaiseContractError ERROR_INVALID_ARGUMENT, context, _
            "The sequence cannot be Nothing."
    End If
    If candidate.InternalRole <> ROLE_LIST And _
        candidate.InternalRole <> ROLE_QUERY And _
        candidate.InternalRole <> ROLE_DATA_VIEW And _
        Not IsGenericCollectionRole(candidate.InternalRole) Then
        RaiseContractError ERROR_INVALID_OPERATION, context, _
            "This operation requires a typed sequence."
    End If
End Sub

Private Sub RequireOrderedQuery( _
    ByVal candidate As ROneCOne, _
    ByVal context As String _
)
    If candidate Is Nothing Then
        RaiseContractError ERROR_INVALID_ARGUMENT, context, _
            "The ordered query cannot be Nothing."
    End If
    If candidate.InternalRole <> ROLE_QUERY Or _
        Not IsOrderingOperator(candidate.InternalOperator) Then
        RaiseContractError ERROR_INVALID_OPERATION, context, _
            context & " is available only after an Order, OrderBy, or ThenBy operation."
    End If
End Sub

Private Sub RequireList(ByVal candidate As ROneCOne, ByVal context As String)
    If candidate Is Nothing Then
        RaiseContractError ERROR_INVALID_ARGUMENT, context, _
            "The list cannot be Nothing."
    End If
    If candidate.InternalRole <> ROLE_LIST Then
        RaiseContractError ERROR_INVALID_OPERATION, context, _
            "This operation requires a materialized List<T>."
    End If
End Sub

Private Sub ValidateIndex( _
    ByVal index As Long, _
    ByVal itemCount As Long, _
    ByVal memberName As String _
)
    If index < 0 Or index >= itemCount Then RaiseIndexError memberName, index
End Sub

Private Sub RaiseIndexError(ByVal memberName As String, ByVal index As Long)
    RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
        "Index " & CStr(index) & " is outside the sequence."
End Sub

Friend Function WrappedItem(ByVal index As Long) As ROneCOne
    Dim values As Collection

    ' Materialized stores index their backing arrays directly in O(1). Only a
    ' deferred query or live view materializes per read, which is the price
    ' of always reflecting the latest data.
    If mRole = ROLE_LIST Then
        ValidateIndex index, mItemsCount, "Item"
        Set WrappedItem = mItems(index + 1)
        Exit Function
    End If
    If IsGenericCollectionRole(mRole) Then
        If mCollectionKind = COLLECTION_READ_ONLY Then
            Set WrappedItem = mSource.WrappedItem(index)
            Exit Function
        End If
        ValidateIndex index, mItemsCount, "Item"
        If IsDictionaryCollection Then
            Set WrappedItem = CreateDictionaryEntry( _
                mKeys(index + 1), mItems(index + 1))
        Else
            Set WrappedItem = mItems(index + 1)
        End If
        Exit Function
    End If
    Set values = MaterializeWrapped
    ValidateIndex index, values.Count, "Item"
    Set WrappedItem = values.Item(index + 1)
End Function

Private Sub ReplaceItem(ByVal index As Long, ByVal wrapped As ROneCOne)
    RequireList Me, "Item"
    ValidateIndex index, mItemsCount, "Item"
    ArrReplaceAt mItems, index + 1, wrapped
    mVersion = mVersion + 1
End Sub

Private Function NormalizeUnary( _
    ByVal candidate As ROneCOne, _
    ByVal memberName As String _
) As ROneCOne
    Dim normalized As ROneCOne
    Dim parameters As Collection

    RequireSequence Me, memberName
    If candidate Is Nothing Then
        RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
            "The expression or delegate cannot be Nothing."
    End If

    If IsDelegateRole(candidate.InternalRole) Then
        If candidate.IsAction Then
            RaiseContractError ERROR_INVALID_OPERATION, memberName, _
                "A query selector or predicate requires a Func."
        End If
        If candidate.Arity <> 1 And candidate.Arity <> -1 Then
            RaiseContractError ERROR_ARITY_MISMATCH, memberName, _
                "The delegate must accept exactly one value."
        End If
        Set NormalizeUnary = candidate
        Exit Function
    End If

    Select Case candidate.InternalRole
        Case ROLE_VALUE, ROLE_PARAMETER, ROLE_EXPRESSION
            Set parameters = New Collection
            CollectExpressionParameters candidate, parameters
            If parameters.Count <> 1 Then
                RaiseContractError ERROR_ARITY_MISMATCH, memberName, _
                    "The expression must reference exactly one parameter."
            End If
            Set normalized = New ROneCOne
            normalized.ConfigureLambda candidate, parameters
            Set NormalizeUnary = normalized

        Case Else
            RaiseContractError ERROR_INVALID_OPERATION, memberName, _
                "Expected a unary expression or delegate."
    End Select
End Function

Private Sub CollectExpressionParameters( _
    ByVal node As ROneCOne, _
    ByVal parameters As Collection _
)
    Select Case node.InternalRole
        Case ROLE_VALUE
            Exit Sub

        Case ROLE_PARAMETER
            If Not CollectionContainsIdentity(parameters, node) Then
                parameters.Add node
            End If

        Case ROLE_EXPRESSION
            CollectExpressionParameters node.InternalLeft, parameters
            If Not node.InternalRight Is Nothing Then
                CollectExpressionParameters node.InternalRight, parameters
            End If

        Case Else
            RaiseContractError ERROR_INVALID_OPERATION, _
                "CollectExpressionParameters", _
                "The expression contains an unsupported node."
    End Select
End Sub

Private Function InvokeUnary( _
    ByVal delegateValue As ROneCOne, _
    ByVal wrapped As ROneCOne _
) As ROneCOne
    Dim arguments As Collection

    Set arguments = New Collection
    arguments.Add wrapped
    Set InvokeUnary = delegateValue.InvokeValue(arguments)
End Function

Private Function InvokePredicate( _
    ByVal predicate As ROneCOne, _
    ByVal wrapped As ROneCOne, _
    ByVal memberName As String _
) As Boolean
    Dim result As ROneCOne

    Set result = InvokeUnary(predicate, wrapped)
    InvokePredicate = CBool(RequireScalar(result, memberName))
End Function

Private Function FindTerminalElement( _
    ByVal predicate As ROneCOne, _
    ByVal takeLast As Boolean, _
    ByVal memberName As String _
) As ROneCOne
    Dim found As ROneCOne
    Dim matches As Boolean
    Dim values As Collection
    Dim wrapped As ROneCOne

    Set values = MaterializeWrapped
    For Each wrapped In values
        If predicate Is Nothing Then
            matches = True
        Else
            matches = InvokePredicate(predicate, wrapped, memberName)
        End If
        If matches Then
            Set found = wrapped
            If Not takeLast Then
                Set FindTerminalElement = found
                Exit Function
            End If
        End If
    Next wrapped
    If found Is Nothing Then RaiseEmptySequence memberName
    Set FindTerminalElement = found
End Function

Private Function SelectTerminalValue( _
    ByVal wrapped As ROneCOne, _
    ByVal selector As ROneCOne _
) As ROneCOne
    If selector Is Nothing Then
        Set SelectTerminalValue = wrapped
    Else
        Set SelectTerminalValue = InvokeUnary(selector, wrapped)
    End If
End Function

Private Function ExtremeValue( _
    ByVal selector As ROneCOne, _
    ByVal comparer As ROneCOne, _
    ByVal findMaximum As Boolean, _
    ByVal memberName As String _
) As ROneCOne
    Dim best As ROneCOne
    Dim candidate As ROneCOne
    Dim values As Collection
    Dim wrapped As ROneCOne

    Set values = MaterializeWrapped
    If values.Count = 0 Then RaiseEmptySequence memberName
    For Each wrapped In values
        Set candidate = SelectTerminalValue(wrapped, selector)
        If best Is Nothing Then
            Set best = candidate
        ElseIf findMaximum And _
            CompareWithComparer(candidate, best, comparer, memberName) > 0 Then
            Set best = candidate
        ElseIf Not findMaximum And _
            CompareWithComparer(candidate, best, comparer, memberName) < 0 Then
            Set best = candidate
        End If
    Next wrapped
    Set ExtremeValue = best
End Function

Private Function TryFindTerminalElement( _
    ByVal predicate As ROneCOne, _
    ByVal takeLast As Boolean, _
    ByVal memberName As String, _
    ByRef found As Boolean _
) As ROneCOne
    Dim matches As Boolean
    Dim values As Collection
    Dim wrapped As ROneCOne

    Set values = MaterializeWrapped
    For Each wrapped In values
        matches = predicate Is Nothing
        If Not predicate Is Nothing Then
            matches = InvokePredicate(predicate, wrapped, memberName)
        End If
        If matches Then
            found = True
            Set TryFindTerminalElement = wrapped
            If Not takeLast Then Exit Function
        End If
    Next wrapped
End Function

Private Function FindSingleElement( _
    ByVal predicate As ROneCOne, _
    ByVal allowDefault As Boolean, _
    ByVal memberName As String _
) As ROneCOne
    Dim found As ROneCOne
    Dim matches As Boolean
    Dim values As Collection
    Dim wrapped As ROneCOne

    Set values = MaterializeWrapped
    For Each wrapped In values
        matches = predicate Is Nothing
        If Not predicate Is Nothing Then
            matches = InvokePredicate(predicate, wrapped, memberName)
        End If
        If matches Then
            If Not found Is Nothing Then
                RaiseContractError ERROR_INVALID_OPERATION, memberName, _
                    "The sequence contains more than one matching element."
            End If
            Set found = wrapped
        End If
    Next wrapped
    If found Is Nothing And Not allowDefault Then RaiseEmptySequence memberName
    Set FindSingleElement = found
End Function

Private Function InvokeBinary( _
    ByVal delegateValue As ROneCOne, _
    ByVal leftValue As ROneCOne, _
    ByVal rightValue As ROneCOne _
) As ROneCOne
    Dim arguments As Collection

    Set arguments = New Collection
    arguments.Add leftValue
    arguments.Add rightValue
    Set InvokeBinary = delegateValue.InvokeValue(arguments)
End Function

Private Function NormalizeStandaloneUnary( _
    ByVal candidate As ROneCOne, _
    ByVal memberName As String _
) As ROneCOne
    Dim normalized As ROneCOne
    Dim parameters As Collection

    If candidate Is Nothing Then
        RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
            "The expression or delegate cannot be Nothing."
    End If
    If IsDelegateRole(candidate.InternalRole) Then
        If candidate.IsAction Then
            RaiseContractError ERROR_INVALID_OPERATION, memberName, _
                "A predicate requires a Func."
        End If
        If candidate.Arity <> 1 And candidate.Arity <> -1 Then
            RaiseContractError ERROR_ARITY_MISMATCH, memberName, _
                "The predicate must accept exactly one value."
        End If
        Set NormalizeStandaloneUnary = candidate
        Exit Function
    End If
    RequireExpression candidate, memberName
    Set parameters = New Collection
    CollectExpressionParameters candidate, parameters
    If parameters.Count <> 1 Then
        RaiseContractError ERROR_ARITY_MISMATCH, memberName, _
            "The expression must reference exactly one parameter."
    End If
    Set normalized = New ROneCOne
    normalized.ConfigureLambda candidate, parameters
    Set NormalizeStandaloneUnary = normalized
End Function

Private Function NormalizeBinaryDelegate( _
    ByVal candidate As Variant, _
    ByVal memberName As String _
) As ROneCOne
    Dim candidateValue As ROneCOne
    Dim parameters As Collection
    Dim normalized As ROneCOne

    Set candidateValue = RequireROneCOne(candidate, memberName)
    If IsDelegateRole(candidateValue.InternalRole) Then
        If candidateValue.IsAction Then
            RaiseContractError ERROR_INVALID_OPERATION, memberName, _
                "The operation requires a Func."
        End If
        If candidateValue.Arity <> 2 And candidateValue.Arity <> -1 Then
            RaiseContractError ERROR_ARITY_MISMATCH, memberName, _
                "The Func must accept exactly two values."
        End If
        Set NormalizeBinaryDelegate = candidateValue
        Exit Function
    End If
    RequireExpression candidateValue, memberName
    Set parameters = New Collection
    CollectExpressionParameters candidateValue, parameters
    If parameters.Count <> 2 Then
        RaiseContractError ERROR_ARITY_MISMATCH, memberName, _
            "The expression must reference exactly two parameters."
    End If
    Set normalized = New ROneCOne
    normalized.ConfigureLambda candidateValue, parameters
    Set NormalizeBinaryDelegate = normalized
End Function

Private Function CreateSequenceComparer( _
    ByVal callable As Variant, _
    ByVal memberName As String, _
    ByVal equality As Boolean, _
    ByVal context As String _
) As ROneCOne
    Dim comparerValue As ROneCOne
    Dim descriptor As ROneCOne
    Dim descriptors As Collection
    Dim expectedType As VbVarType
    Dim objectType As String

    Set comparerValue = CreateCallableDelegate(callable, memberName, False)
    If mRole = ROLE_LIST Or mRole = ROLE_QUERY Then
        expectedType = mElementType
        objectType = mElementObjectType
    Else
        expectedType = vbVariant
    End If
    If Not comparerValue.InternalHasParameterSignature Then
        Set descriptors = New Collection
        Set descriptor = New ROneCOne
        descriptor.ConfigureParameter expectedType, objectType
        descriptors.Add descriptor
        Set descriptor = New ROneCOne
        descriptor.ConfigureParameter expectedType, objectType
        descriptors.Add descriptor
        comparerValue.ConfigureParameterSignature descriptors
    End If
    If Not comparerValue.InternalHasReturnSignature Then
        If equality Then
            comparerValue.ConfigureReturnSignature vbBoolean, vbNullString
        Else
            comparerValue.ConfigureReturnSignature vbLong, vbNullString
        End If
    End If
    Set CreateSequenceComparer = NormalizeComparer( _
        comparerValue, equality, context)
End Function

Private Function NormalizeComparer( _
    ByVal candidate As Variant, _
    ByVal equality As Boolean, _
    ByVal memberName As String _
) As ROneCOne
    Dim comparerValue As ROneCOne

    Set comparerValue = RequireROneCOne(candidate, memberName)
    RequireDelegate comparerValue, memberName
    If comparerValue.IsAction Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            "A comparer requires a Func result."
    End If
    If comparerValue.Arity <> 2 And comparerValue.Arity <> -1 Then
        RaiseContractError ERROR_ARITY_MISMATCH, memberName, _
            "A comparer must accept exactly two values."
    End If
    If comparerValue.InternalHasReturnSignature Then
        If equality And comparerValue.InternalReturnType <> vbBoolean Then
            RaiseContractError ERROR_TYPE_MISMATCH, memberName, _
                "An equality comparer must return Boolean."
        ElseIf Not equality And comparerValue.InternalReturnType <> vbLong Then
            RaiseContractError ERROR_TYPE_MISMATCH, memberName, _
                "An ordering comparer must return Long."
        End If
    End If
    Set NormalizeComparer = comparerValue
End Function

Private Function ExtremeElement( _
    ByVal selector As ROneCOne, _
    ByVal comparer As ROneCOne, _
    ByVal findMaximum As Boolean, _
    ByVal memberName As String _
) As ROneCOne
    Dim bestElement As ROneCOne
    Dim bestKey As ROneCOne
    Dim candidateKey As ROneCOne
    Dim values As Collection
    Dim wrapped As ROneCOne

    Set values = MaterializeWrapped
    If values.Count = 0 Then RaiseEmptySequence memberName
    For Each wrapped In values
        Set candidateKey = InvokeUnary(selector, wrapped)
        If bestElement Is Nothing Then
            Set bestElement = wrapped
            Set bestKey = candidateKey
        ElseIf findMaximum And _
            CompareWithComparer( _
                candidateKey, bestKey, comparer, memberName) > 0 Then
            Set bestElement = wrapped
            Set bestKey = candidateKey
        ElseIf Not findMaximum And _
            CompareWithComparer( _
                candidateKey, bestKey, comparer, memberName) < 0 Then
            Set bestElement = wrapped
            Set bestKey = candidateKey
        End If
    Next wrapped
    Set ExtremeElement = bestElement
End Function

Private Sub RaiseEmptySequence(ByVal memberName As String)
    RaiseContractError ERROR_INVALID_OPERATION, memberName, _
        "The sequence contains no matching elements."
End Sub

Private Function MaterializeValues() As Collection
    Dim values As Collection
    Dim wrapped As ROneCOne
    Dim wrappedValues As Collection

    Set values = New Collection
    Set wrappedValues = MaterializeWrapped
    For Each wrapped In wrappedValues
        AddUnwrappedValue values, wrapped
    Next wrapped
    Set MaterializeValues = values
End Function

Private Sub AddUnwrappedValue( _
    ByVal values As Collection, _
    ByVal wrapped As ROneCOne _
)
    Dim actualObject As Object
    Dim actualValue As Variant

    If wrapped.InternalHasObjectValue Then
        Set actualObject = wrapped.InternalObjectValue
        Set actualValue = actualObject
    Else
        actualValue = wrapped.InternalScalarValue
    End If
    values.Add actualValue
End Sub

Private Function ApplyWhere( _
    ByVal sourceValues As Collection, _
    Optional ByVal negate As Boolean = False _
) As Collection
    Dim output As Collection
    Dim wrapped As ROneCOne

    Set output = New Collection
    For Each wrapped In sourceValues
        If InvokePredicate(mQueryDelegate, wrapped, "Where") Xor negate Then
            output.Add wrapped
        End If
    Next wrapped
    Set ApplyWhere = output
End Function

Private Function ApplySelect(ByVal sourceValues As Collection) As Collection
    Dim output As Collection
    Dim selected As ROneCOne
    Dim wrapped As ROneCOne

    Set output = New Collection
    For Each wrapped In sourceValues
        Set selected = InvokeUnary(mQueryDelegate, wrapped)
        ValidateElementWrapped selected, "SelectItems"
        output.Add selected
    Next wrapped
    Set ApplySelect = output
End Function

Private Function ApplyTake(ByVal sourceValues As Collection) As Collection
    Dim index As Long
    Dim output As Collection
    Dim takeCount As Long

    Set output = New Collection
    takeCount = mQueryCount
    If takeCount > sourceValues.Count Then takeCount = sourceValues.Count
    For index = 1 To takeCount
        output.Add sourceValues.Item(index)
    Next index
    Set ApplyTake = output
End Function

Private Function ApplySkip(ByVal sourceValues As Collection) As Collection
    Dim index As Long
    Dim output As Collection

    Set output = New Collection
    For index = mQueryCount + 1 To sourceValues.Count
        output.Add sourceValues.Item(index)
    Next index
    Set ApplySkip = output
End Function

Private Function ApplyTakeWhile( _
    ByVal sourceValues As Collection _
) As Collection
    Dim output As Collection
    Dim wrapped As ROneCOne

    Set output = New Collection
    For Each wrapped In sourceValues
        If Not InvokePredicate(mQueryDelegate, wrapped, "TakeWhile") Then Exit For
        output.Add wrapped
    Next wrapped
    Set ApplyTakeWhile = output
End Function

Private Function ApplySkipWhile( _
    ByVal sourceValues As Collection _
) As Collection
    Dim output As Collection
    Dim skipping As Boolean
    Dim wrapped As ROneCOne

    skipping = True
    Set output = New Collection
    For Each wrapped In sourceValues
        If skipping Then
            skipping = InvokePredicate(mQueryDelegate, wrapped, "SkipWhile")
        End If
        If Not skipping Then output.Add wrapped
    Next wrapped
    Set ApplySkipWhile = output
End Function

Private Function ApplyTakeLast( _
    ByVal sourceValues As Collection _
) As Collection
    Dim firstIndex As Long
    Dim index As Long
    Dim output As Collection

    Set output = New Collection
    firstIndex = sourceValues.Count - mQueryCount + 1
    If firstIndex < 1 Then firstIndex = 1
    For index = firstIndex To sourceValues.Count
        output.Add sourceValues.Item(index)
    Next index
    Set ApplyTakeLast = output
End Function

Private Function ApplySkipLast( _
    ByVal sourceValues As Collection _
) As Collection
    Dim index As Long
    Dim lastIndex As Long
    Dim output As Collection

    Set output = New Collection
    lastIndex = sourceValues.Count - mQueryCount
    If lastIndex < 0 Then lastIndex = 0
    For index = 1 To lastIndex
        output.Add sourceValues.Item(index)
    Next index
    Set ApplySkipLast = output
End Function

Private Function ApplyDistinct(ByVal sourceValues As Collection) As Collection
    Dim output As Collection
    Dim wrapped As ROneCOne

    Set output = New Collection
    For Each wrapped In sourceValues
        If Not WrappedCollectionContains( _
            output, wrapped, mQueryComparer) Then output.Add wrapped
    Next wrapped
    Set ApplyDistinct = output
End Function

Private Function ApplyDistinctBy( _
    ByVal sourceValues As Collection _
) As Collection
    Dim key As ROneCOne
    Dim keys As Collection
    Dim output As Collection
    Dim wrapped As ROneCOne

    Set keys = New Collection
    Set output = New Collection
    For Each wrapped In sourceValues
        Set key = InvokeUnary(mQueryDelegate, wrapped)
        If Not WrappedCollectionContains(keys, key, mQueryComparer) Then
            keys.Add key
            output.Add wrapped
        End If
    Next wrapped
    Set ApplyDistinctBy = output
End Function

Private Function ApplyUnion(ByVal sourceValues As Collection) As Collection
    Dim otherValues As Collection
    Dim output As Collection
    Dim wrapped As ROneCOne

    Set output = New Collection
    For Each wrapped In sourceValues
        If Not WrappedCollectionContains(output, wrapped, mQueryComparer) Then
            output.Add wrapped
        End If
    Next wrapped
    Set otherValues = mQueryValue.MaterializeWrapped
    For Each wrapped In otherValues
        If Not WrappedCollectionContains(output, wrapped, mQueryComparer) Then
            output.Add wrapped
        End If
    Next wrapped
    Set ApplyUnion = output
End Function

Private Function ApplyIntersect( _
    ByVal sourceValues As Collection _
) As Collection
    Dim otherValues As Collection
    Dim output As Collection
    Dim wrapped As ROneCOne

    Set otherValues = mQueryValue.MaterializeWrapped
    Set output = New Collection
    For Each wrapped In sourceValues
        If WrappedCollectionContains(otherValues, wrapped, mQueryComparer) And _
            Not WrappedCollectionContains(output, wrapped, mQueryComparer) Then
            output.Add wrapped
        End If
    Next wrapped
    Set ApplyIntersect = output
End Function

Private Function ApplyExcept(ByVal sourceValues As Collection) As Collection
    Dim otherValues As Collection
    Dim output As Collection
    Dim wrapped As ROneCOne

    Set otherValues = mQueryValue.MaterializeWrapped
    Set output = New Collection
    For Each wrapped In sourceValues
        If Not WrappedCollectionContains(otherValues, wrapped, mQueryComparer) And _
            Not WrappedCollectionContains(output, wrapped, mQueryComparer) Then
            output.Add wrapped
        End If
    Next wrapped
    Set ApplyExcept = output
End Function

Private Function ApplyDefaultIfEmpty( _
    ByVal sourceValues As Collection _
) As Collection
    Dim output As Collection

    If sourceValues.Count > 0 Then
        Set ApplyDefaultIfEmpty = CopyWrappedCollection(sourceValues)
        Exit Function
    End If
    Set output = New Collection
    If mQueryValue Is Nothing Then
        output.Add DefaultWrappedElement
    Else
        output.Add mQueryValue
    End If
    Set ApplyDefaultIfEmpty = output
End Function

Private Function ApplyChunk(ByVal sourceValues As Collection) As Collection
    Dim chunk As ROneCOne
    Dim index As Long
    Dim output As Collection

    Set output = New Collection
    For index = 1 To sourceValues.Count
        If (index - 1) Mod mQueryCount = 0 Then
            Set chunk = New ROneCOne
            chunk.ConfigureList mSource.InternalElementType, _
                mSource.InternalElementObjectType
            output.Add CreateObjectResult(chunk)
        End If
        chunk.InternalAddWrapped sourceValues.Item(index)
    Next index
    Set ApplyChunk = output
End Function

Private Function ApplySelectMany( _
    ByVal sourceValues As Collection _
) As Collection
    Dim nested As ROneCOne
    Dim nestedObject As Object
    Dim nestedValues As Collection
    Dim output As Collection
    Dim selected As ROneCOne
    Dim wrapped As ROneCOne

    Set output = New Collection
    For Each wrapped In sourceValues
        Set selected = InvokeUnary(mQueryDelegate, wrapped)
        If Not selected.InternalHasObjectValue Then
            RaiseContractError ERROR_TYPE_MISMATCH, "SelectMany", _
                "The selector must return a ROneCOne sequence."
        End If
        Set nestedObject = selected.InternalObjectValue
        If nestedObject Is Nothing Or Not TypeOf nestedObject Is ROneCOne Then
            RaiseContractError ERROR_TYPE_MISMATCH, "SelectMany", _
                "The selector must return a ROneCOne sequence."
        End If
        Set nested = nestedObject
        RequireSequence nested, "SelectMany"
        Set nestedValues = nested.MaterializeWrapped
        For Each selected In nestedValues
            ValidateElementWrapped selected, "SelectMany"
            output.Add selected
        Next selected
    Next wrapped
    Set ApplySelectMany = output
End Function

Private Function ApplyCast( _
    ByVal sourceValues As Collection, _
    ByVal filterOnly As Boolean _
) As Collection
    Dim output As Collection
    Dim wrapped As ROneCOne

    Set output = New Collection
    For Each wrapped In sourceValues
        If WrappedMatchesType(wrapped, mElementType, mElementObjectType) Then
            output.Add wrapped
        ElseIf Not filterOnly Then
            RaiseContractError ERROR_TYPE_MISMATCH, "Cast", _
                "A sequence value cannot be cast to the requested type."
        End If
    Next wrapped
    Set ApplyCast = output
End Function

Private Function WrappedMatchesType( _
    ByVal wrapped As ROneCOne, _
    ByVal expectedType As VbVarType, _
    ByVal expectedObjectType As String _
) As Boolean
    If expectedType = vbVariant Then
        WrappedMatchesType = True
        Exit Function
    End If
    If expectedType = vbObject Then
        If Not wrapped.InternalHasObjectValue Then Exit Function
        If wrapped.InternalObjectValue Is Nothing Then
            WrappedMatchesType = True
        ElseIf Len(expectedObjectType) = 0 Then
            WrappedMatchesType = True
        Else
            WrappedMatchesType = (StrComp( _
                TypeName(wrapped.InternalObjectValue), expectedObjectType, _
                vbTextCompare) = 0)
        End If
        Exit Function
    End If
    If wrapped.InternalHasObjectValue Then Exit Function
    WrappedMatchesType = (VarTypeOf(wrapped.InternalScalarValue) = expectedType)
End Function

Private Function ApplySetBy(ByVal sourceValues As Collection) As Collection
    Dim candidateKeys As Collection
    Dim key As ROneCOne
    Dim otherValues As Collection
    Dim output As Collection
    Dim seenKeys As Collection
    Dim wrapped As ROneCOne

    Set output = New Collection
    Set seenKeys = New Collection
    Set otherValues = mQueryValue.MaterializeWrapped
    If mOperator = QUERY_UNION_BY Then
        AddDistinctSelectedValues sourceValues, output, seenKeys
        AddDistinctSelectedValues otherValues, output, seenKeys
        Set ApplySetBy = output
        Exit Function
    End If

    Set candidateKeys = DistinctQueryKeys(otherValues)
    For Each wrapped In sourceValues
        Set key = InvokeUnary(mQueryDelegate, wrapped)
        If Not WrappedCollectionContains( _
            seenKeys, key, mQueryComparer) Then
            If mOperator = QUERY_INTERSECT_BY Then
                If WrappedCollectionContains( _
                    candidateKeys, key, mQueryComparer) Then
                    output.Add wrapped
                    seenKeys.Add key
                End If
            ElseIf Not WrappedCollectionContains( _
                candidateKeys, key, mQueryComparer) Then
                output.Add wrapped
                seenKeys.Add key
            End If
        End If
    Next wrapped
    Set ApplySetBy = output
End Function

Private Sub AddDistinctSelectedValues( _
    ByVal sourceValues As Collection, _
    ByVal output As Collection, _
    ByVal seenKeys As Collection _
)
    Dim key As ROneCOne
    Dim wrapped As ROneCOne

    For Each wrapped In sourceValues
        Set key = InvokeUnary(mQueryDelegate, wrapped)
        If Not WrappedCollectionContains(seenKeys, key, mQueryComparer) Then
            seenKeys.Add key
            output.Add wrapped
        End If
    Next wrapped
End Sub

Private Function DistinctQueryKeys( _
    ByVal values As Collection _
) As Collection
    Dim output As Collection
    Dim wrapped As ROneCOne

    Set output = New Collection
    For Each wrapped In values
        If Not WrappedCollectionContains(output, wrapped, mQueryComparer) Then
            output.Add wrapped
        End If
    Next wrapped
    Set DistinctQueryKeys = output
End Function

Private Function DefaultWrappedElement() As ROneCOne
    If mElementType = vbObject Then
        Set DefaultWrappedElement = CreateObjectResult(Nothing)
        Exit Function
    End If
    Select Case mElementType
        Case vbString
            Set DefaultWrappedElement = CreateScalarResult(vbNullString)
        Case vbBoolean
            Set DefaultWrappedElement = CreateScalarResult(False)
        Case vbDate
            Set DefaultWrappedElement = CreateScalarResult(CDate(0))
        Case vbByte
            Set DefaultWrappedElement = CreateScalarResult(CByte(0))
        Case vbInteger
            Set DefaultWrappedElement = CreateScalarResult(CInt(0))
        Case vbLong
            Set DefaultWrappedElement = CreateScalarResult(CLng(0))
        Case vbLongLong
            Set DefaultWrappedElement = CreateScalarResult(CLngLng(0))
        Case vbSingle
            Set DefaultWrappedElement = CreateScalarResult(CSng(0))
        Case vbDouble
            Set DefaultWrappedElement = CreateScalarResult(CDbl(0))
        Case vbCurrency
            Set DefaultWrappedElement = CreateScalarResult(CCur(0))
        Case Else
            Set DefaultWrappedElement = CreateScalarResult(Empty)
    End Select
End Function

Private Function ApplyOrderChain() As Collection
    Dim comparers() As ROneCOne
    Dim current As ROneCOne
    Dim descending() As Boolean
    Dim indexes() As Long
    Dim itemCount As Long
    Dim itemIndex As Long
    Dim items() As ROneCOne
    Dim key As ROneCOne
    Dim keyScalars() As Variant
    Dim keyTypes() As VbVarType
    Dim keys() As ROneCOne
    Dim level As ROneCOne
    Dim levelCount As Long
    Dim levelIndex As Long
    Dim levels As Collection
    Dim output As Collection
    Dim sourceValues As Collection
    Dim workspace() As Long

    Set levels = New Collection
    Set current = Me
    Do While current.InternalRole = ROLE_QUERY And _
        IsOrderingOperator(current.InternalOperator)
        If levels.Count = 0 Then
            levels.Add current
        Else
            levels.Add current, Before:=1
        End If
        Set current = current.InternalSource
    Loop

    Set sourceValues = current.MaterializeWrapped
    Set output = New Collection
    itemCount = sourceValues.Count
    If itemCount = 0 Then
        Set ApplyOrderChain = output
        Exit Function
    End If

    levelCount = levels.Count
    ReDim items(1 To itemCount)
    ReDim keys(1 To levelCount, 1 To itemCount)
    ReDim keyScalars(1 To levelCount, 1 To itemCount)
    ReDim keyTypes(1 To levelCount, 1 To itemCount)
    ReDim comparers(1 To levelCount)
    ReDim descending(1 To levelCount)
    ReDim indexes(1 To itemCount)
    ReDim workspace(1 To itemCount)

    For itemIndex = 1 To itemCount
        Set items(itemIndex) = sourceValues.Item(itemIndex)
        indexes(itemIndex) = itemIndex
    Next itemIndex
    For levelIndex = 1 To levelCount
        Set level = levels.Item(levelIndex)
        Set comparers(levelIndex) = level.InternalQueryComparer
        descending(levelIndex) = _
            IsDescendingOrderingOperator(level.InternalOperator)
        For itemIndex = 1 To itemCount
            Set key = InvokeUnary(level.InternalQueryDelegate, items(itemIndex))
            Set keys(levelIndex, itemIndex) = key
            If comparers(levelIndex) Is Nothing Then
                If key.InternalHasObjectValue Then
                    RaiseContractError ERROR_TYPE_MISMATCH, "Order", _
                        "Object values require an explicit ordering comparer."
                End If
                keyScalars(levelIndex, itemIndex) = key.InternalScalarValue
                keyTypes(levelIndex, itemIndex) = _
                    VarTypeOf(keyScalars(levelIndex, itemIndex))
            End If
        Next itemIndex
    Next levelIndex

    StableMergeOrdering indexes, workspace, keys, keyScalars, keyTypes, _
        comparers, descending, levelCount, itemCount
    For itemIndex = 1 To itemCount
        output.Add items(indexes(itemIndex))
    Next itemIndex
    Set ApplyOrderChain = output
End Function

Private Sub StableMergeOrdering( _
    ByRef indexes() As Long, _
    ByRef workspace() As Long, _
    ByRef keys() As ROneCOne, _
    ByRef keyScalars() As Variant, _
    ByRef keyTypes() As VbVarType, _
    ByRef comparers() As ROneCOne, _
    ByRef descending() As Boolean, _
    ByVal levelCount As Long, _
    ByVal itemCount As Long _
)
    Dim comparison As Long
    Dim leftEnd As Long
    Dim leftPosition As Long
    Dim leftStart As Long
    Dim outputPosition As Long
    Dim rightEnd As Long
    Dim rightPosition As Long
    Dim width As Long

    width = 1
    Do While width < itemCount
        leftStart = 1
        Do While leftStart <= itemCount
            leftEnd = leftStart + width - 1
            If leftEnd > itemCount Then leftEnd = itemCount
            rightEnd = leftStart + (width * 2) - 1
            If rightEnd > itemCount Then rightEnd = itemCount
            leftPosition = leftStart
            rightPosition = leftEnd + 1
            outputPosition = leftStart

            Do While leftPosition <= leftEnd And rightPosition <= rightEnd
                comparison = CompareOrderingIndexes( _
                    indexes(leftPosition), indexes(rightPosition), _
                    keys, keyScalars, keyTypes, comparers, descending, _
                    levelCount)
                If comparison <= 0 Then
                    workspace(outputPosition) = indexes(leftPosition)
                    leftPosition = leftPosition + 1
                Else
                    workspace(outputPosition) = indexes(rightPosition)
                    rightPosition = rightPosition + 1
                End If
                outputPosition = outputPosition + 1
            Loop
            Do While leftPosition <= leftEnd
                workspace(outputPosition) = indexes(leftPosition)
                leftPosition = leftPosition + 1
                outputPosition = outputPosition + 1
            Loop
            Do While rightPosition <= rightEnd
                workspace(outputPosition) = indexes(rightPosition)
                rightPosition = rightPosition + 1
                outputPosition = outputPosition + 1
            Loop
            For outputPosition = leftStart To rightEnd
                indexes(outputPosition) = workspace(outputPosition)
            Next outputPosition
            leftStart = leftStart + (width * 2)
        Loop
        width = width * 2
    Loop
End Sub

Private Function CompareOrderingIndexes( _
    ByVal leftIndex As Long, _
    ByVal rightIndex As Long, _
    ByRef keys() As ROneCOne, _
    ByRef keyScalars() As Variant, _
    ByRef keyTypes() As VbVarType, _
    ByRef comparers() As ROneCOne, _
    ByRef descending() As Boolean, _
    ByVal levelCount As Long _
) As Long
    Dim comparison As Long
    Dim levelIndex As Long

    For levelIndex = 1 To levelCount
        If comparers(levelIndex) Is Nothing Then
            comparison = CompareCachedOrderingKeys( _
                keyScalars(levelIndex, leftIndex), _
                keyTypes(levelIndex, leftIndex), _
                keyScalars(levelIndex, rightIndex), _
                keyTypes(levelIndex, rightIndex))
        Else
            comparison = CompareWithComparer( _
                keys(levelIndex, leftIndex), _
                keys(levelIndex, rightIndex), _
                comparers(levelIndex), "Order")
        End If
        If comparison < 0 Then
            comparison = -1
        ElseIf comparison > 0 Then
            comparison = 1
        End If
        If descending(levelIndex) Then
            comparison = -comparison
        End If
        If comparison <> 0 Then
            CompareOrderingIndexes = comparison
            Exit Function
        End If
    Next levelIndex

    If leftIndex < rightIndex Then
        CompareOrderingIndexes = -1
    ElseIf leftIndex > rightIndex Then
        CompareOrderingIndexes = 1
    End If
End Function

Private Function CompareCachedOrderingKeys( _
    ByVal leftValue As Variant, _
    ByVal leftType As VbVarType, _
    ByVal rightValue As Variant, _
    ByVal rightType As VbVarType _
) As Long
    If leftType = vbNull Or rightType = vbNull Then
        If leftType = rightType Then
            CompareCachedOrderingKeys = 0
        ElseIf leftType = vbNull Then
            CompareCachedOrderingKeys = -1
        Else
            CompareCachedOrderingKeys = 1
        End If
        Exit Function
    End If
    If IsOrderingNumericType(leftType) And _
        IsOrderingNumericType(rightType) Then
        CompareCachedOrderingKeys = CompareScalarValues( _
            leftValue, rightValue, "Order")
        Exit Function
    End If
    If leftType <> rightType Then
        RaiseIncompatibleOrderingTypes leftType, rightType, "Order"
    End If
    Select Case leftType
        Case vbEmpty
            CompareCachedOrderingKeys = 0
        Case vbBoolean
            If CBool(leftValue) = CBool(rightValue) Then
                CompareCachedOrderingKeys = 0
            ElseIf Not CBool(leftValue) Then
                CompareCachedOrderingKeys = -1
            Else
                CompareCachedOrderingKeys = 1
            End If
        Case vbString
            CompareCachedOrderingKeys = CLng(Sgn(StrComp( _
                CStr(leftValue), CStr(rightValue), vbBinaryCompare)))
        Case vbDate
            CompareCachedOrderingKeys = CompareScalarValues( _
                leftValue, rightValue, "Order")
        Case Else
            RaiseIncompatibleOrderingTypes leftType, rightType, "Order"
    End Select
End Function

Private Function ApplyReverse(ByVal sourceValues As Collection) As Collection
    Dim index As Long
    Dim output As Collection

    Set output = New Collection
    For index = sourceValues.Count To 1 Step -1
        output.Add sourceValues.Item(index)
    Next index
    Set ApplyReverse = output
End Function

Private Function CopyWrappedCollection( _
    ByVal sourceValues As Collection _
) As Collection
    Dim output As Collection

    Set output = New Collection
    AddWrappedRange output, sourceValues
    Set CopyWrappedCollection = output
End Function

Private Sub AddWrappedRange( _
    ByVal target As Collection, _
    ByVal sourceValues As Collection _
)
    Dim wrapped As ROneCOne

    For Each wrapped In sourceValues
        target.Add wrapped
    Next wrapped
End Sub

Private Function WrappedCollectionContains( _
    ByVal values As Collection, _
    ByVal candidate As ROneCOne, _
    Optional ByVal comparer As ROneCOne _
) As Boolean
    Dim wrapped As ROneCOne

    For Each wrapped In values
        If ValuesEqualWithComparer( _
            wrapped, candidate, comparer, "Contains") Then
            WrappedCollectionContains = True
            Exit Function
        End If
    Next wrapped
End Function

Private Function CompareWrapped( _
    ByVal leftValue As ROneCOne, _
    ByVal rightValue As ROneCOne, _
    ByVal memberName As String _
) As Long
    Dim leftScalar As Variant
    Dim leftType As VbVarType
    Dim rightScalar As Variant
    Dim rightType As VbVarType

    If leftValue.InternalHasObjectValue Or rightValue.InternalHasObjectValue Then
        RaiseContractError ERROR_TYPE_MISMATCH, memberName, _
            "Object values require an explicit ordering comparer."
    End If
    leftScalar = leftValue.InternalScalarValue
    rightScalar = rightValue.InternalScalarValue
    If IsNull(leftScalar) Or IsNull(rightScalar) Then
        If IsNull(leftScalar) And IsNull(rightScalar) Then
            CompareWrapped = 0
        ElseIf IsNull(leftScalar) Then
            CompareWrapped = -1
        Else
            CompareWrapped = 1
        End If
        Exit Function
    End If

    leftType = VarTypeOf(leftScalar)
    rightType = VarTypeOf(rightScalar)
    If IsOrderingNumericType(leftType) And _
        IsOrderingNumericType(rightType) Then
        CompareWrapped = CompareScalarValues( _
            leftScalar, rightScalar, memberName)
        Exit Function
    End If
    If leftType <> rightType Then
        RaiseIncompatibleOrderingTypes leftType, rightType, memberName
    End If

    Select Case leftType
        Case vbEmpty
            CompareWrapped = 0
        Case vbBoolean
            If CBool(leftScalar) = CBool(rightScalar) Then
                CompareWrapped = 0
            ElseIf Not CBool(leftScalar) Then
                CompareWrapped = -1
            Else
                CompareWrapped = 1
            End If
        Case vbString
            CompareWrapped = CLng(Sgn(StrComp( _
                CStr(leftScalar), CStr(rightScalar), vbBinaryCompare)))
        Case vbDate
            CompareWrapped = CompareScalarValues( _
                leftScalar, rightScalar, memberName)
        Case Else
            RaiseIncompatibleOrderingTypes leftType, rightType, memberName
    End Select
End Function

Private Function CompareScalarValues( _
    ByVal leftScalar As Variant, _
    ByVal rightScalar As Variant, _
    ByVal memberName As String _
) As Long
    On Error GoTo ComparisonFailed
    If leftScalar = rightScalar Then
        CompareScalarValues = 0
    ElseIf leftScalar < rightScalar Then
        CompareScalarValues = -1
    Else
        CompareScalarValues = 1
    End If
    On Error GoTo 0
    Exit Function

ComparisonFailed:
    On Error GoTo 0
    RaiseContractError ERROR_TYPE_MISMATCH, memberName, _
        "The selected keys cannot be compared safely."
End Function

Private Function IsOrderingNumericType( _
    ByVal valueType As VbVarType _
) As Boolean
    Select Case valueType
        Case vbByte, vbInteger, vbLong, vbLongLong
            IsOrderingNumericType = True
        Case vbSingle, vbDouble, vbCurrency, vbDecimal
            IsOrderingNumericType = True
    End Select
End Function

Private Sub RaiseIncompatibleOrderingTypes( _
    ByVal leftType As VbVarType, _
    ByVal rightType As VbVarType, _
    ByVal memberName As String _
)
    RaiseContractError ERROR_TYPE_MISMATCH, memberName, _
        "Incompatible Variant key types: " & _
        TypeNameForVarType(leftType) & " and " & _
        TypeNameForVarType(rightType) & "."
End Sub

Private Function CreateBinaryExpression( _
    ByVal expressionOperator As Long, _
    ByVal operand As Variant _
) As ROneCOne
    Dim expression As ROneCOne
    Dim rightValue As ROneCOne

    RequireExpression Me, "Expression receiver"
    Set rightValue = CoerceExpression(operand)
    Set expression = New ROneCOne
    expression.ConfigureExpression expressionOperator, Me, rightValue
    Set CreateBinaryExpression = expression
End Function

Private Function CreateUnaryExpression(ByVal expressionOperator As Long) As ROneCOne
    Dim expression As ROneCOne
    Dim emptyRight As ROneCOne

    RequireExpression Me, "Expression receiver"
    Set expression = New ROneCOne
    expression.ConfigureExpression expressionOperator, Me, emptyRight
    Set CreateUnaryExpression = expression
End Function

Private Function CoerceExpression(ByVal candidate As Variant) As ROneCOne
    Dim candidateObject As Object
    Dim node As ROneCOne

    If IsObject(candidate) Then
        Set candidateObject = candidate
        If TypeOf candidateObject Is ROneCOne Then
            Set node = candidateObject
            RequireExpression node, "Expression operand"
            Set CoerceExpression = node
            Exit Function
        End If
    End If

    Set node = New ROneCOne
    If IsObject(candidate) Then
        node.ConfigureObjectValue candidateObject
    Else
        node.ConfigureScalarValue candidate
    End If
    Set CoerceExpression = node
End Function

Private Function ValuesEqualWithComparer( _
    ByVal leftValue As ROneCOne, _
    ByVal rightValue As ROneCOne, _
    ByVal comparer As ROneCOne, _
    ByVal memberName As String _
) As Boolean
    Dim result As ROneCOne

    If comparer Is Nothing Then
        ValuesEqualWithComparer = ValuesEqual(leftValue, rightValue)
    Else
        Set result = InvokeBinary(comparer, leftValue, rightValue)
        ValuesEqualWithComparer = CBool(RequireScalar(result, memberName))
    End If
End Function

Private Function CompareWithComparer( _
    ByVal leftValue As ROneCOne, _
    ByVal rightValue As ROneCOne, _
    ByVal comparer As ROneCOne, _
    ByVal memberName As String _
) As Long
    Dim result As ROneCOne

    If comparer Is Nothing Then
        CompareWithComparer = CompareWrapped( _
            leftValue, rightValue, memberName)
    Else
        Set result = InvokeBinary(comparer, leftValue, rightValue)
        CompareWithComparer = CLng(RequireScalar(result, memberName))
    End If
End Function

Friend Function NormalizeContextUnary( _
    ByVal candidate As Variant, _
    ByVal memberName As String, _
    ByVal returnsObject As Boolean _
) As ROneCOne
    Dim expression As ROneCOne

    RequireSequence Me, memberName
    If Not IsObject(candidate) Then
        If VarTypeOf(candidate) = vbString Then
            Set expression = BuildMemberPath(CStr(candidate), returnsObject)
            Set NormalizeContextUnary = NormalizeUnary(expression, memberName)
            Exit Function
        End If
    End If
    Set NormalizeContextUnary = NormalizeUnary( _
        RequireROneCOne(candidate, memberName), memberName)
End Function

Private Function BindContextPredicate( _
    ByVal predicateDelegate As ROneCOne _
) As ROneCOne
    Dim bound As ROneCOne
    Dim descriptor As ROneCOne
    Dim descriptors As Collection

    RequireDelegate predicateDelegate, "Predicate"
    If predicateDelegate.IsAction Then
        RaiseContractError ERROR_INVALID_OPERATION, "Predicate", _
            "A predicate requires a Func result."
    End If
    If predicateDelegate.Arity <> 1 And predicateDelegate.Arity <> -1 Then
        RaiseContractError ERROR_ARITY_MISMATCH, "Predicate", _
            "The predicate must accept exactly one sequence element."
    End If
    If predicateDelegate.InternalHasReturnSignature Then
        If predicateDelegate.InternalReturnType <> vbBoolean Or _
            Len(predicateDelegate.InternalReturnObjectType) > 0 Then
            RaiseContractError ERROR_TYPE_MISMATCH, "Predicate", _
                "The predicate result must be Boolean."
        End If
    End If
    If predicateDelegate.InternalHasParameterSignature Then
        Set descriptor = predicateDelegate.InternalParameters.Item(1)
        ValidateContextPredicateParameter descriptor
    End If

    Set bound = CloneDelegate(predicateDelegate)
    If Not bound.InternalHasParameterSignature Then
        Set descriptors = New Collection
        Set descriptor = New ROneCOne
        descriptor.ConfigureParameter mElementType, mElementObjectType
        descriptors.Add descriptor
        bound.ConfigureParameterSignature descriptors
    End If
    If Not bound.InternalHasReturnSignature Then
        bound.ConfigureReturnSignature vbBoolean, vbNullString
    End If
    Set BindContextPredicate = bound
End Function

Private Sub ValidateContextPredicateParameter( _
    ByVal descriptor As ROneCOne _
)
    Dim expectedObjectType As String
    Dim expectedType As VbVarType

    If descriptor.InternalIsByRef Then
        RaiseContractError ERROR_TYPE_MISMATCH, "Predicate", _
            "A sequence predicate parameter must be passed ByVal."
    End If
    expectedType = descriptor.InternalExpectedType
    expectedObjectType = descriptor.InternalExpectedObjectType
    If expectedType = vbVariant Then Exit Sub

    If mElementType = vbObject Then
        If expectedType <> vbObject Then
            RaiseContractError ERROR_TYPE_MISMATCH, "Predicate", _
                "The predicate parameter does not accept the sequence type."
        End If
        If Len(expectedObjectType) = 0 Then Exit Sub
        If StrComp(expectedObjectType, mElementObjectType, _
            vbTextCompare) = 0 Then Exit Sub
    ElseIf expectedType = mElementType Then
        Exit Sub
    End If
    RaiseContractError ERROR_TYPE_MISMATCH, "Predicate", _
        "The predicate parameter does not accept the sequence type."
End Sub

Private Function BuildMemberPath( _
    ByVal memberPath As String, _
    ByVal returnsObject As Boolean _
) As ROneCOne
    Dim expression As ROneCOne
    Dim index As Long
    Dim normalized As String
    Dim parts As Variant
    Dim propagateNullSafe As Boolean
    Dim segment As String
    Dim nullSafe As Boolean

    normalized = NormalizeMemberPath(memberPath, "Condition")
    parts = VBA.Split(normalized, ".")
    Set expression = Element
    For index = LBound(parts) To UBound(parts)
        segment = CStr(parts(index))
        nullSafe = (Left$(segment, 1) = "?")
        If nullSafe Then
            segment = Mid$(segment, 2)
            propagateNullSafe = True
        End If
        Set expression = expression.Member( _
            segment, index < UBound(parts) Or returnsObject, _
            propagateNullSafe)
    Next index
    Set BuildMemberPath = expression
End Function

Private Function NormalizeMemberPath( _
    ByVal memberPath As String, _
    ByVal memberName As String _
) As String
    Dim index As Long
    Dim normalized As String
    Dim parts As Variant
    Dim segment As String

    memberPath = VBA.Replace(memberPath, "?.", ".?")
    memberPath = Trim$(memberPath)
    If Len(memberPath) = 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
            "The member path cannot be empty."
    End If
    parts = VBA.Split(memberPath, ".")
    For index = LBound(parts) To UBound(parts)
        segment = Trim$(CStr(parts(index)))
        If Left$(segment, 1) = "?" Then
            segment = "?" & Trim$(Mid$(segment, 2))
        End If
        If Len(segment) = 0 Then
            RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
                "The member path cannot contain an empty segment."
        End If
        If segment = "?" Then
            RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
                "A null-safe member segment must include a member name."
        End If
        If Len(normalized) > 0 Then normalized = normalized & "."
        normalized = normalized & segment
    Next index
    NormalizeMemberPath = normalized
End Function

Private Function ContextMemberExpression( _
    ByVal returnsObject As Boolean _
) As ROneCOne
    If mRole <> ROLE_WHERE_BUILDER Then
        RaiseContractError ERROR_INVALID_OPERATION, "Where", _
            "A contextual member condition must begin with Where(memberName)."
    End If
    If Len(mContextMemberPath) = 0 Then
        Set ContextMemberExpression = mSource.Element
    Else
        Set ContextMemberExpression = mSource.Condition( _
            mContextMemberPath, returnsObject)
    End If
End Function

Private Function ConditionReceiver( _
    ByVal returnsObject As Boolean, _
    ByVal memberName As String _
) As ROneCOne
    If mRole = ROLE_WHERE_BUILDER Then
        Set ConditionReceiver = ContextMemberExpression(returnsObject)
    Else
        RequireExpression Me, memberName & " receiver"
        Set ConditionReceiver = Me
    End If
End Function

Private Function CompleteWhereBinary( _
    ByVal expressionOperator As Long, _
    ByVal operand As Variant _
) As ROneCOne
    Dim expression As ROneCOne
    Dim receiver As ROneCOne
    Dim rightValue As ROneCOne

    Set receiver = ContextMemberExpression(False)
    Set rightValue = CoerceExpression(operand)
    Set expression = New ROneCOne
    expression.ConfigureExpression expressionOperator, receiver, rightValue
    Set CompleteWhereBinary = mSource.Where(expression)
End Function

Private Function CompleteStringCondition( _
    ByVal expressionOperator As Long, _
    ByVal operand As String, _
    ByVal compareMode As VbCompareMethod _
) As ROneCOne
    Dim expression As ROneCOne
    Dim receiver As ROneCOne
    Dim rightValue As ROneCOne

    ValidateStringCompareMode compareMode, "String condition"
    Set receiver = ConditionReceiver(False, "String condition")
    Set rightValue = CoerceExpression(operand)
    Set expression = New ROneCOne
    expression.ConfigureExpression expressionOperator, receiver, _
        rightValue, compareMode
    Set CompleteStringCondition = CompleteWhereOrExpression(expression)
End Function

Private Sub ValidateStringCompareMode( _
    ByVal compareMode As VbCompareMethod, _
    ByVal memberName As String _
)
    If compareMode <> vbBinaryCompare And compareMode <> vbTextCompare Then
        RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
            "Comparison mode must be vbBinaryCompare or vbTextCompare."
    End If
End Sub

Private Function CreateMembershipCondition( _
    ByVal values As Variant, _
    ByVal negate As Boolean, _
    Optional ByVal comparer As Variant _
) As ROneCOne
    Dim comparerValue As ROneCOne
    Dim expression As ROneCOne
    Dim operatorCode As Long
    Dim receiver As ROneCOne
    Dim sequenceValue As ROneCOne

    Set receiver = ConditionReceiver(False, "In")
    Set sequenceValue = NormalizeMembershipSequence(values, "In")
    If Not IsMissing(comparer) Then
        Set comparerValue = NormalizeComparer(comparer, True, "IsIn")
    End If
    operatorCode = OP_IN_SEQUENCE
    If negate Then operatorCode = OP_NOT_IN_SEQUENCE
    Set expression = New ROneCOne
    expression.ConfigureMembershipExpression _
        operatorCode, receiver, sequenceValue, comparerValue
    Set CreateMembershipCondition = CompleteWhereOrExpression(expression)
End Function

Private Function NormalizeMembershipSequence( _
    ByVal values As Variant, _
    ByVal memberName As String _
) As ROneCOne
    Dim candidate As Object
    Dim sequenceValue As ROneCOne

    If IsObject(values) Then
        Set candidate = values
        If Not candidate Is Nothing Then
            If TypeOf candidate Is ROneCOne Then
                Set sequenceValue = candidate
                RequireSequence sequenceValue, memberName
                Set NormalizeMembershipSequence = sequenceValue
                Exit Function
            End If
        End If
    End If
    Set sequenceValue = ROneCOne.ListOf(vbVariant)
    sequenceValue.AddRange values
    Set NormalizeMembershipSequence = sequenceValue
End Function

' The only place the runtime may call VBA's IsArray. IsArray evaluates its
' argument in a value context, so a Variant holding an object is
' dereferenced through that object's default member. This class's default
' member is Run, which rejects the zero arguments such a dereference hands
' it. The raise is swallowed wherever the caller has On Error Resume Next
' active, so a call still returns the right answer while leaving Err dirty
' for the caller to trip over later. An object is never an array, so
' testing IsObject first costs nothing and removes the hazard.
Private Function IsArrayValue(ByVal value As Variant) As Boolean
    If IsObject(value) Then Exit Function
    IsArrayValue = IsArray(value)
End Function

' The only place the runtime may call VBA's VarType, for the same reason.
' VarType also evaluates its argument in a value context, so a Variant
' holding one of these objects fires Run with no arguments, and the
' intrinsic reports vbObject only because that call failed. Answering
' vbObject directly keeps the result identical, drops the dirty Err, and
' removes the accidental dependency: a foreign object whose default member
' succeeds would otherwise report the type of that member's value and slip
' past the object branches the runtime relies on.
Private Function VarTypeOf(ByVal value As Variant) As VbVarType
    If IsObject(value) Then
        VarTypeOf = vbObject
        Exit Function
    End If
    VarTypeOf = VarType(value)
End Function

Private Function IsSequenceContainer(ByVal value As Variant) As Boolean
    Dim candidate As Object
    Dim sequenceValue As ROneCOne

    If IsArrayValue(value) Then
        IsSequenceContainer = True
        Exit Function
    End If
    If Not IsObject(value) Then Exit Function
    Set candidate = value
    If candidate Is Nothing Then Exit Function
    If TypeOf candidate Is ROneCOne Then
        ' The typed local matters. A Friend member is not on the IDispatch
        ' interface, so reading InternalRole through the Object local would
        ' compile and then raise 438 at run time, even from inside this very
        ' class. Declaring candidate As ROneCOne instead is not equivalent:
        ' Set would then reject a plain Collection before the Else runs.
        Set sequenceValue = candidate
        IsSequenceContainer = _
            sequenceValue.InternalRole = ROLE_LIST Or _
            sequenceValue.InternalRole = ROLE_QUERY Or _
            IsGenericCollectionRole(sequenceValue.InternalRole)
    Else
        IsSequenceContainer = (TypeName(candidate) = "Collection")
    End If
End Function

Private Function IsExpressionVariant(ByVal candidate As Variant) As Boolean
    Dim candidateObject As Object
    Dim runtimeValue As ROneCOne

    If Not IsObject(candidate) Then Exit Function
    Set candidateObject = candidate
    If candidateObject Is Nothing Then Exit Function
    If Not TypeOf candidateObject Is ROneCOne Then Exit Function
    Set runtimeValue = candidateObject
    Select Case runtimeValue.InternalRole
        Case ROLE_VALUE, ROLE_PARAMETER, ROLE_EXPRESSION
            IsExpressionVariant = True
    End Select
End Function

Private Function CreateQuantifierCondition( _
    ByVal operatorCode As Long, _
    ByVal predicate As Variant, _
    ByVal hasPredicate As Boolean _
) As ROneCOne
    Dim expression As ROneCOne
    Dim predicateValue As ROneCOne
    Dim receiver As ROneCOne

    Set receiver = ConditionReceiver(True, "Nested quantifier")
    If hasPredicate Then
        Set predicateValue = NormalizeStandaloneUnary( _
            RequireROneCOne(predicate, "Nested quantifier"), _
            "Nested quantifier")
    End If
    Set expression = New ROneCOne
    expression.ConfigureQuantifierExpression _
        operatorCode, receiver, predicateValue
    Set CreateQuantifierCondition = CompleteWhereOrExpression(expression)
End Function

Private Function CompleteWhereOrExpression( _
    ByVal expression As ROneCOne _
) As ROneCOne
    If mRole = ROLE_WHERE_BUILDER Then
        Set CompleteWhereOrExpression = mSource.Where(expression)
    Else
        Set CompleteWhereOrExpression = expression
    End If
End Function

Private Sub ExecuteTryOperation()
    Dim arguments As Collection
    Dim catchAction As ROneCOne
    Dim catchError As ROneCOne
    Dim errorValue As ROneCOne
    Dim finallyError As ROneCOne

    RequireTry Me, "Execute"
    Set arguments = New Collection
    If Not InvokeActionCaptured(mTryBody, arguments, errorValue) Then
        Set catchAction = FindCatchAction(errorValue.ErrorNumber)
        If Not catchAction Is Nothing Then
            Set arguments = New Collection
            If catchAction.Arity <> 0 Then
                arguments.Add WrapInvocationArgument(errorValue)
            End If
            If InvokeActionCaptured(catchAction, arguments, catchError) Then
                Set errorValue = Nothing
            Else
                Set errorValue = catchError
            End If
        End If
    End If

    If Not mFinallyAction Is Nothing Then
        Set arguments = New Collection
        If Not InvokeActionCaptured(mFinallyAction, arguments, finallyError) Then
            Set errorValue = finallyError
        End If
    End If
    If Not errorValue Is Nothing Then RaiseCapturedError errorValue
End Sub

Private Function ExecuteUsingScope( _
    ByVal suppliedArguments As Collection _
) As ROneCOne
    Dim cleanupError As ROneCOne
    Dim errors As Collection
    Dim invocationArguments As Collection
    Dim result As ROneCOne
    Dim work As ROneCOne
    Dim workArity As Long
    Dim workError As ROneCOne

    If suppliedArguments.Count <> 1 Then
        RaiseContractError ERROR_ARITY_MISMATCH, "Using.Run", _
            "Using.Run requires one Action or Func."
    End If
    Set work = RequireROneCOne( _
        Unwrap(suppliedArguments.Item(1)), "Using.Run")
    RequireDelegate work, "Using.Run"
    workArity = work.Arity
    If workArity <> 0 And workArity <> 1 And workArity <> -1 Then
        RaiseContractError ERROR_ARITY_MISMATCH, "Using.Run", _
            "The delegate must accept zero arguments or the resource."
    End If
    Set invocationArguments = New Collection
    If workArity = 1 Then invocationArguments.Add CreateObjectResult(mSource)
    Call InvokeDelegateCaptured(work, invocationArguments, result, workError)
    Call DisposeResourceCaptured(mSource, cleanupError)
    If Not workError Is Nothing Or Not cleanupError Is Nothing Then
        Set errors = New Collection
        If Not workError Is Nothing Then errors.Add workError
        If Not cleanupError Is Nothing Then errors.Add cleanupError
        If errors.Count = 1 Then
            RaiseCapturedError errors.Item(1)
        Else
            RaiseCapturedError CreateAggregateException(errors)
        End If
    End If
    Set ExecuteUsingScope = result
End Function

Private Function InvokeDelegateCaptured( _
    ByVal delegateValue As ROneCOne, _
    ByVal arguments As Collection, _
    ByRef result As ROneCOne, _
    ByRef errorValue As ROneCOne _
) As Boolean
    On Error GoTo InvocationFailed
    Set result = delegateValue.InvokeValue(arguments)
    InvokeDelegateCaptured = True
    Exit Function

InvocationFailed:
    Set errorValue = New ROneCOne
    errorValue.ConfigureError Err.Number, Err.Source, Err.Description, _
        Err.HelpFile, Err.HelpContext
    Err.Clear
    On Error GoTo 0
End Function

Private Function DisposeResourceCaptured( _
    ByVal resource As ROneCOne, _
    ByRef errorValue As ROneCOne _
) As Boolean
    On Error GoTo DisposeFailed
    resource.Dispose
    DisposeResourceCaptured = True
    Exit Function

DisposeFailed:
    Set errorValue = New ROneCOne
    errorValue.ConfigureError Err.Number, Err.Source, Err.Description, _
        Err.HelpFile, Err.HelpContext
    Err.Clear
    On Error GoTo 0
End Function

Private Function InvokeActionCaptured( _
    ByVal actionValue As ROneCOne, _
    ByVal arguments As Collection, _
    ByRef errorValue As ROneCOne _
) As Boolean
    Dim ignored As ROneCOne

    On Error GoTo InvocationFailed
    Set ignored = actionValue.InvokeValue(arguments)
    InvokeActionCaptured = True
    Exit Function

InvocationFailed:
    Set errorValue = New ROneCOne
    errorValue.ConfigureError Err.Number, Err.Source, Err.Description, _
        Err.HelpFile, Err.HelpContext
    Err.Clear
    On Error GoTo 0
End Function

Private Function FindCatchAction(ByVal errorNumber As Long) As ROneCOne
    Dim index As Long

    For index = 1 To mCatchActions.Count
        If CLng(mCatchNumbers.Item(index)) = 0 Or _
            CLng(mCatchNumbers.Item(index)) = errorNumber Then
            Set FindCatchAction = mCatchActions.Item(index)
            Exit Function
        End If
    Next index
End Function

Private Sub RaiseCapturedError(ByVal errorValue As ROneCOne)
    Err.Raise errorValue.ErrorNumber, errorValue.ErrorSource, _
        errorValue.Message, errorValue.HelpFile, errorValue.HelpContext
End Sub

Friend Function InvokeValue(ByVal arguments As Collection) As ROneCOne
    Dim delegateValue As ROneCOne
    Dim firstResult As ROneCOne
    Dim nextArguments As Collection
    Dim result As ROneCOne

    Select Case mRole
        Case ROLE_LAMBDA
            If arguments.Count <> mParameters.Count Then
                RaiseContractError ERROR_ARITY_MISMATCH, "Run", _
                    "Expected " & CStr(mParameters.Count) & " argument(s), received " & _
                    CStr(arguments.Count) & "."
            End If
            Set result = EvaluateNode(mLeft, mParameters, arguments)

        Case ROLE_METHOD
            ValidateDelegateArguments arguments
            Set result = InvokeMethod(arguments)

        Case ROLE_PROCEDURE
            ValidateDelegateArguments arguments
            Set result = InvokeProcedure(arguments)

        Case ROLE_NATIVE
            ValidateDelegateArguments arguments
            Set result = InvokeNative(arguments)

        Case ROLE_MULTICAST
            ValidateDelegateArguments arguments
            For Each delegateValue In mInvocationList
                Set result = delegateValue.InvokeValue(arguments)
            Next delegateValue

        Case ROLE_COMPOSITION
            Set firstResult = mLeft.InvokeValue(arguments)
            Set nextArguments = New Collection
            nextArguments.Add firstResult
            Set result = mNext.InvokeValue(nextArguments)

        Case Else
            RaiseContractError ERROR_INVALID_OPERATION, "Run", _
                "Only delegates can be invoked."
    End Select
    Set InvokeValue = FinalizeDelegateResult(result)
End Function

Private Function EvaluateNode( _
    ByVal node As ROneCOne, _
    ByVal parameters As Collection, _
    ByVal arguments As Collection _
) As ROneCOne
    Dim leftResult As ROneCOne
    Dim leftText As String
    Dim objectResult As Object
    Dim rightResult As ROneCOne
    Dim rightText As String
    Dim scalarResult As Variant
    Dim operatorCode As Long

    Select Case node.InternalRole
        Case ROLE_VALUE
            Set EvaluateNode = node

        Case ROLE_PARAMETER
            Set EvaluateNode = ResolveParameter(node, parameters, arguments)

        Case ROLE_EXPRESSION
            Set leftResult = EvaluateNode(node.InternalLeft, parameters, arguments)
            operatorCode = node.InternalOperator

            If operatorCode = OP_AND_ALSO Then
                If Not CBool(RequireScalar(leftResult, "AndAlso")) Then
                    Set EvaluateNode = CreateScalarResult(False)
                    Exit Function
                End If
            ElseIf operatorCode = OP_OR_ELSE Then
                If CBool(RequireScalar(leftResult, "OrElse")) Then
                    Set EvaluateNode = CreateScalarResult(True)
                    Exit Function
                End If
            End If

            If IsMemberOperator(operatorCode) Then
                Set EvaluateNode = ReadMember( _
                    leftResult, _
                    node.InternalMethodName, _
                    IsObjectMemberOperator(operatorCode), _
                    IsSafeMemberOperator(operatorCode))
                Exit Function
            End If

            If Not node.InternalRight Is Nothing Then
                Set rightResult = EvaluateNode(node.InternalRight, parameters, arguments)
            End If

            If operatorCode = OP_IN_SEQUENCE Or _
                operatorCode = OP_NOT_IN_SEQUENCE Then
                scalarResult = SequenceContainsWrapped( _
                    node.InternalQueryValue, leftResult, _
                    node.InternalQueryComparer)
                If operatorCode = OP_NOT_IN_SEQUENCE Then
                    scalarResult = Not CBool(scalarResult)
                End If
                Set EvaluateNode = CreateScalarResult(scalarResult)
                Exit Function
            End If

            If IsQuantifierOperator(operatorCode) Then
                scalarResult = EvaluateNestedQuantifier( _
                    leftResult, node.InternalQueryDelegate, operatorCode)
                Set EvaluateNode = CreateScalarResult(scalarResult)
                Exit Function
            End If

            Select Case operatorCode
                Case OP_ADD
                    scalarResult = RequireScalar(leftResult, "Add") + _
                        RequireScalar(rightResult, "Add")
                Case OP_SUBTRACT
                    scalarResult = RequireScalar(leftResult, "Subtract") - _
                        RequireScalar(rightResult, "Subtract")
                Case OP_MULTIPLY
                    scalarResult = RequireScalar(leftResult, "Multiply") * _
                        RequireScalar(rightResult, "Multiply")
                Case OP_DIVIDE
                    scalarResult = RequireScalar(leftResult, "Divide") / _
                        RequireScalar(rightResult, "Divide")
                Case OP_MODULO
                    scalarResult = RequireScalar(leftResult, "Modulo") Mod _
                        RequireScalar(rightResult, "Modulo")
                Case OP_NEGATE
                    scalarResult = -RequireScalar(leftResult, "Negate")
                Case OP_CONCAT
                    scalarResult = CStr(RequireScalar(leftResult, "Concat")) & _
                        CStr(RequireScalar(rightResult, "Concat"))
                Case OP_EQUAL
                    scalarResult = ValuesEqual(leftResult, rightResult)
                Case OP_NOT_EQUAL
                    scalarResult = Not ValuesEqual(leftResult, rightResult)
                Case OP_LESS_THAN
                    If HasNullScalar(leftResult) Or HasNullScalar(rightResult) Then
                        scalarResult = False
                    Else
                        scalarResult = RequireScalar(leftResult, "LessThan") < _
                            RequireScalar(rightResult, "LessThan")
                    End If
                Case OP_LESS_THAN_OR_EQUAL
                    If HasNullScalar(leftResult) Or HasNullScalar(rightResult) Then
                        scalarResult = False
                    Else
                        scalarResult = _
                            RequireScalar(leftResult, "LessThanOrEqual") <= _
                            RequireScalar(rightResult, "LessThanOrEqual")
                    End If
                Case OP_GREATER_THAN
                    If HasNullScalar(leftResult) Or HasNullScalar(rightResult) Then
                        scalarResult = False
                    Else
                        scalarResult = RequireScalar(leftResult, "GreaterThan") > _
                            RequireScalar(rightResult, "GreaterThan")
                    End If
                Case OP_GREATER_THAN_OR_EQUAL
                    If HasNullScalar(leftResult) Or HasNullScalar(rightResult) Then
                        scalarResult = False
                    Else
                        scalarResult = _
                            RequireScalar(leftResult, "GreaterThanOrEqual") >= _
                            RequireScalar(rightResult, "GreaterThanOrEqual")
                    End If
                Case OP_AND_ALSO
                    scalarResult = CBool(RequireScalar(rightResult, "AndAlso"))
                Case OP_OR_ELSE
                    scalarResult = CBool(RequireScalar(rightResult, "OrElse"))
                Case OP_NOT
                    scalarResult = Not CBool(RequireScalar(leftResult, "NotExpression"))
                Case OP_STARTS_WITH
                    If HasNullScalar(leftResult) Then GoTo FalseStringResult
                    leftText = CStr(RequireScalar(leftResult, "StartsWith"))
                    rightText = CStr(RequireScalar(rightResult, "StartsWith"))
                    scalarResult = (StrComp( _
                        Left$(leftText, Len(rightText)), rightText, _
                        node.InternalExpressionOption) = 0)
                Case OP_ENDS_WITH
                    If HasNullScalar(leftResult) Then GoTo FalseStringResult
                    leftText = CStr(RequireScalar(leftResult, "EndsWith"))
                    rightText = CStr(RequireScalar(rightResult, "EndsWith"))
                    scalarResult = (StrComp( _
                        Right$(leftText, Len(rightText)), rightText, _
                        node.InternalExpressionOption) = 0)
                Case OP_CONTAINS_TEXT
                    If HasNullScalar(leftResult) Then GoTo FalseStringResult
                    leftText = CStr(RequireScalar(leftResult, "ContainsText"))
                    rightText = CStr(RequireScalar(rightResult, "ContainsText"))
                    scalarResult = (InStr( _
                        1, leftText, rightText, _
                        node.InternalExpressionOption) > 0)
                Case OP_IS_NOTHING, OP_IS_NOT_NOTHING
                    If Not leftResult.InternalHasObjectValue Then
                        RaiseContractError ERROR_TYPE_MISMATCH, "IsNothing", _
                            "IsNothing requires an object-valued expression."
                    End If
                    Set objectResult = leftResult.InternalObjectValue
                    scalarResult = (objectResult Is Nothing)
                    If operatorCode = OP_IS_NOT_NOTHING Then
                        scalarResult = Not CBool(scalarResult)
                    End If
                Case OP_IS_NULL_OR_EMPTY
                    If leftResult.InternalHasObjectValue Then
                        RaiseContractError ERROR_TYPE_MISMATCH, "IsNullOrEmpty", _
                            "IsNullOrEmpty requires a scalar expression."
                    End If
                    scalarResult = IsNull(leftResult.InternalScalarValue)
                    If Not CBool(scalarResult) Then
                        scalarResult = (Len(CStr( _
                            leftResult.InternalScalarValue)) = 0)
                    End If
                Case OP_MATCHES_PATTERN
                    If HasNullScalar(leftResult) Then GoTo FalseStringResult
                    leftText = CStr(RequireScalar(leftResult, "MatchesPattern"))
                    rightText = CStr(RequireScalar(rightResult, "MatchesPattern"))
                    If node.InternalExpressionOption = vbTextCompare Then
                        leftText = LCase$(leftText)
                        rightText = LCase$(rightText)
                    End If
                    scalarResult = leftText Like rightText
                Case OP_STRING_EQUAL, OP_STRING_NOT_EQUAL
                    If HasNullScalar(leftResult) Or HasNullScalar(rightResult) Then
                        scalarResult = ValuesEqual(leftResult, rightResult)
                    Else
                        scalarResult = (StrComp( _
                            CStr(RequireScalar(leftResult, "EqualToIgnoreCase")), _
                            CStr(RequireScalar(rightResult, "EqualToIgnoreCase")), _
                            node.InternalExpressionOption) = 0)
                    End If
                    If operatorCode = OP_STRING_NOT_EQUAL Then
                        scalarResult = Not CBool(scalarResult)
                    End If
                Case Else
                    RaiseContractError ERROR_INVALID_OPERATION, "EvaluateNode", _
                        "Unknown expression operator."
            End Select

            Set EvaluateNode = CreateScalarResult(scalarResult)
            Exit Function

FalseStringResult:
            Set EvaluateNode = CreateScalarResult(False)

        Case Else
            RaiseContractError ERROR_INVALID_OPERATION, "EvaluateNode", _
                "The lambda body contains a non-expression value."
    End Select
End Function

Private Function ResolveParameter( _
    ByVal parameterValue As ROneCOne, _
    ByVal parameters As Collection, _
    ByVal arguments As Collection _
) As ROneCOne
    Dim index As Long
    Dim argumentValue As ROneCOne

    For index = 1 To parameters.Count
        If parameters.Item(index) Is parameterValue Then
            Set argumentValue = WidenScalarToward( _
                arguments.Item(index), parameterValue.InternalExpectedType)
            ValidateParameterValue parameterValue, argumentValue
            Set ResolveParameter = argumentValue
            Exit Function
        End If
    Next index

    RaiseContractError ERROR_UNBOUND_PARAMETER, "ResolveParameter", _
        "The expression contains an unbound parameter."
End Function

Private Sub ValidateParameterValue( _
    ByVal parameterValue As ROneCOne, _
    ByVal argumentValue As ROneCOne _
)
    Dim actualType As VbVarType
    Dim expectedObjectType As String
    Dim expectedType As VbVarType

    expectedType = parameterValue.InternalExpectedType
    expectedObjectType = parameterValue.InternalExpectedObjectType
    If expectedType = vbVariant Then Exit Sub

    If expectedType = vbObject Then
        If Not argumentValue.InternalHasObjectValue Then
            RaiseTypeMismatch parameterValue, argumentValue
        End If
        If Len(expectedObjectType) <> 0 Then
            If StrComp( _
                TypeName(argumentValue.InternalObjectValue), _
                expectedObjectType, _
                vbTextCompare _
            ) <> 0 Then
                RaiseTypeMismatch parameterValue, argumentValue
            End If
        End If
        Exit Sub
    End If

    If argumentValue.InternalHasObjectValue Then
        RaiseTypeMismatch parameterValue, argumentValue
    End If

    actualType = VarTypeOf(argumentValue.InternalScalarValue)
    If actualType <> expectedType Then
        RaiseTypeMismatch parameterValue, argumentValue
    End If
End Sub

Private Sub RaiseTypeMismatch( _
    ByVal parameterValue As ROneCOne, _
    ByVal argumentValue As ROneCOne _
)
    Dim expectedName As String
    Dim actualName As String

    If Len(parameterValue.InternalExpectedObjectType) <> 0 Then
        expectedName = parameterValue.InternalExpectedObjectType
    Else
        expectedName = TypeNameForVarType(parameterValue.InternalExpectedType)
    End If

    If argumentValue.InternalHasObjectValue Then
        actualName = TypeName(argumentValue.InternalObjectValue)
    Else
        actualName = TypeName(argumentValue.InternalScalarValue)
    End If

    RaiseContractError ERROR_TYPE_MISMATCH, "Run", _
        "Expected " & expectedName & ", received " & actualName & "."
End Sub

Private Function InvokeMethod(ByVal arguments As Collection) As ROneCOne
    Dim returnedObject As Object
    Dim returnedValue As Variant
    Dim values() As Variant

    If arguments.Count > 8 Then
        RaiseContractError ERROR_ARITY_MISMATCH, "Run", _
            "Object-method delegates currently support at most 8 arguments."
    End If
    PrepareInvocationValues arguments, values

    On Error GoTo InvocationFailed
    If mHasReturnSignature And mReturnType = vbObject Then
        Select Case arguments.Count
            Case 0
                Set returnedObject = CallByName(mTarget, mMethodName, VbMethod)
            Case 1
                Set returnedObject = CallByName( _
                    mTarget, mMethodName, VbMethod, values(0))
            Case 2
                Set returnedObject = CallByName( _
                    mTarget, mMethodName, VbMethod, _
                    values(0), values(1))
            Case 3
                Set returnedObject = CallByName( _
                    mTarget, mMethodName, VbMethod, _
                    values(0), values(1), values(2))
            Case 4
                Set returnedObject = CallByName( _
                    mTarget, mMethodName, VbMethod, _
                    values(0), values(1), values(2), values(3))
            Case 5
                Set returnedObject = CallByName( _
                    mTarget, mMethodName, VbMethod, _
                    values(0), values(1), values(2), values(3), values(4))
            Case 6
                Set returnedObject = CallByName( _
                    mTarget, mMethodName, VbMethod, _
                    values(0), values(1), values(2), values(3), values(4), values(5))
            Case 7
                Set returnedObject = CallByName( _
                    mTarget, mMethodName, VbMethod, _
                    values(0), values(1), values(2), values(3), values(4), _
                    values(5), values(6))
            Case 8
                Set returnedObject = CallByName( _
                    mTarget, mMethodName, VbMethod, _
                    values(0), values(1), values(2), values(3), values(4), _
                    values(5), values(6), values(7))
        End Select
        Set InvokeMethod = CreateObjectResult(returnedObject)
    Else
        Select Case arguments.Count
            Case 0
                returnedValue = CallByName(mTarget, mMethodName, VbMethod)
            Case 1
                returnedValue = CallByName( _
                    mTarget, mMethodName, VbMethod, values(0))
            Case 2
                returnedValue = CallByName( _
                    mTarget, mMethodName, VbMethod, _
                    values(0), values(1))
            Case 3
                returnedValue = CallByName( _
                    mTarget, mMethodName, VbMethod, _
                    values(0), values(1), values(2))
            Case 4
                returnedValue = CallByName( _
                    mTarget, mMethodName, VbMethod, _
                    values(0), values(1), values(2), values(3))
            Case 5
                returnedValue = CallByName( _
                    mTarget, mMethodName, VbMethod, _
                    values(0), values(1), values(2), values(3), values(4))
            Case 6
                returnedValue = CallByName( _
                    mTarget, mMethodName, VbMethod, _
                    values(0), values(1), values(2), values(3), values(4), values(5))
            Case 7
                returnedValue = CallByName( _
                    mTarget, mMethodName, VbMethod, _
                    values(0), values(1), values(2), values(3), values(4), _
                    values(5), values(6))
            Case 8
                returnedValue = CallByName( _
                    mTarget, mMethodName, VbMethod, _
                    values(0), values(1), values(2), values(3), values(4), _
                    values(5), values(6), values(7))
        End Select
        Set InvokeMethod = CreateScalarResult(returnedValue)
    End If
    Exit Function

InvocationFailed:
    RaiseContractError ERROR_METHOD_INVOCATION, "Run", _
        "Method '" & mMethodName & "' failed with " & CStr(Err.Number) & ": " & _
        Err.Description
End Function

Private Function InvokeProcedure(ByVal arguments As Collection) As ROneCOne
    Dim procedureName As String
    Dim returnedObject As Object
    Dim returnedValue As Variant
    Dim values() As Variant

    If arguments.Count > 30 Then
        RaiseContractError ERROR_ARITY_MISMATCH, "Run", _
            "Workbook procedure delegates support at most 30 arguments."
    End If
    PrepareInvocationValues arguments, values
    procedureName = ResolveProcedureName(mMethodName)

    On Error GoTo InvocationFailed
    If mHasReturnSignature And mReturnType = vbObject Then
        Set returnedObject = RunWorkbookProcedureObject(procedureName, values, arguments.Count)
        Set InvokeProcedure = CreateObjectResult(returnedObject)
    Else
        returnedValue = RunWorkbookProcedure(procedureName, values, arguments.Count)
        Set InvokeProcedure = CreateScalarResult(returnedValue)
    End If
    Exit Function

InvocationFailed:
    RaiseContractError ERROR_METHOD_INVOCATION, "Run", _
        "Procedure '" & mMethodName & "' failed with " & CStr(Err.Number) & ": " & _
        Err.Description
End Function

Private Function InvokeNative(ByVal arguments As Collection) As ROneCOne
    Dim argumentPointers() As LongPtr
    Dim argumentTypes() As Integer
    Dim argumentValues() As Variant
    Dim descriptor As ROneCOne
    Dim hResult As Long
    Dim index As Long
    Dim returnedValue As Variant
    Dim wrapped As ROneCOne

    If Not mHasParameterSignature Or _
        (Not mIsActionDelegate And Not mHasReturnSignature) Then
        RaiseContractError ERROR_NATIVE_INVOCATION, "Native", _
            "Native invocation requires complete Takes and Returns metadata."
    End If
    If arguments.Count > 0 Then
        ReDim argumentPointers(0 To arguments.Count - 1)
        ReDim argumentTypes(0 To arguments.Count - 1)
        ReDim argumentValues(0 To arguments.Count - 1)
        For index = 0 To arguments.Count - 1
            Set descriptor = mParameters.Item(index + 1)
            Set wrapped = arguments.Item(index + 1)
            If descriptor.InternalIsByRef Then
                argumentValues(index) = CLngLng(wrapped.InternalReferenceAddress)
                argumentTypes(index) = vbLongLong
            ElseIf wrapped.InternalHasObjectValue Then
                Set argumentValues(index) = wrapped.InternalObjectValue
                argumentTypes(index) = descriptor.InternalExpectedType
            Else
                argumentValues(index) = wrapped.InternalScalarValue
                argumentTypes(index) = descriptor.InternalExpectedType
            End If
            argumentPointers(index) = VarPtr(argumentValues(index))
        Next index
        hResult = DispCallFunc( _
            0, mNativeAddress, CALLCONV_STDCALL, CInt(mReturnType), _
            arguments.Count, VarPtr(argumentTypes(0)), _
            VarPtr(argumentPointers(0)), VarPtr(returnedValue))
    Else
        hResult = DispCallFunc( _
            0, mNativeAddress, CALLCONV_STDCALL, CInt(mReturnType), _
            0, 0, 0, VarPtr(returnedValue))
    End If
    If hResult <> HRESULT_SUCCESS Then
        RaiseContractError ERROR_NATIVE_INVOCATION, "Native", _
            "DispCallFunc failed with HRESULT " & CStr(hResult) & "."
    End If

    If mHasReturnSignature And mReturnType = vbObject Then
        Set InvokeNative = CreateObjectResult(returnedValue)
    Else
        Set InvokeNative = CreateScalarResult(returnedValue)
    End If
End Function

Private Sub PrepareInvocationValues( _
    ByVal arguments As Collection, _
    ByRef values() As Variant _
)
    Dim index As Long
    Dim wrapped As ROneCOne

    If arguments.Count = 0 Then Exit Sub
    ReDim values(0 To arguments.Count - 1)
    For index = 0 To arguments.Count - 1
        Set wrapped = arguments.Item(index + 1)
        If wrapped.InternalHasObjectValue Then
            Set values(index) = wrapped.InternalObjectValue
        Else
            values(index) = wrapped.InternalScalarValue
        End If
    Next index
End Sub

Private Function ResolveProcedureName(ByVal procedureName As String) As String
    If InStr(1, procedureName, "!", vbBinaryCompare) > 0 Then
        ResolveProcedureName = procedureName
    Else
        ResolveProcedureName = "'" & VBA.Replace(ThisWorkbook.Name, "'", "''") & _
            "'!" & procedureName
    End If
End Function

Private Sub ValidateDelegateArguments(ByVal arguments As Collection)
    Dim argumentValue As ROneCOne
    Dim descriptor As ROneCOne
    Dim index As Long

    If Not mHasParameterSignature Then
        If mRole = ROLE_NATIVE Then
            RaiseContractError ERROR_NATIVE_INVOCATION, "Native", _
                "Native invocation requires a complete Takes signature."
        End If
        Exit Sub
    End If
    If arguments.Count <> mParameters.Count Then
        RaiseContractError ERROR_ARITY_MISMATCH, "Run", _
            "Expected " & CStr(mParameters.Count) & " argument(s), received " & _
            CStr(arguments.Count) & "."
    End If

    For index = 1 To mParameters.Count
        Set descriptor = mParameters.Item(index)
        Set argumentValue = arguments.Item(index)
        If descriptor.InternalIsByRef Then
            If mRole <> ROLE_NATIVE Then
                RaiseContractError ERROR_BYREF_INVOCATION, "Run", _
                    "True ByRef requires a signature-bound native delegate."
            End If
            If argumentValue.InternalRole <> ROLE_REFERENCE Then
                RaiseContractError ERROR_TYPE_MISMATCH, "Run", _
                    "Parameter " & CStr(index) & " requires a ROneCOne reference."
            End If
            If argumentValue.InternalExpectedType <> _
                descriptor.InternalExpectedType Then
                RaiseContractError ERROR_TYPE_MISMATCH, "Run", _
                    "The ByRef argument type does not match its signature."
            End If
        Else
            If argumentValue.InternalRole = ROLE_REFERENCE Then
                RaiseContractError ERROR_TYPE_MISMATCH, "Run", _
                    "A reference was supplied to a ByVal parameter."
            End If
            Set argumentValue = WidenScalarToward( _
                argumentValue, descriptor.InternalExpectedType)
            ReplaceCollectionItemAt arguments, index, argumentValue
            ValidateParameterValue descriptor, argumentValue
        End If
    Next index
End Sub

Private Function FinalizeDelegateResult(ByVal result As ROneCOne) As ROneCOne
    Dim actualType As VbVarType

    If result Is Nothing Then
        RaiseContractError ERROR_INVALID_OPERATION, "Run", _
            "The delegate did not produce an invocation result."
    End If
    If mIsActionDelegate Then
        Set FinalizeDelegateResult = CreateScalarResult(Empty)
        Exit Function
    End If
    If Not mHasReturnSignature Or mReturnType = vbVariant Then
        Set FinalizeDelegateResult = result
        Exit Function
    End If

    If mReturnType = vbObject Then
        If Not result.InternalHasObjectValue Then
            RaiseContractError ERROR_TYPE_MISMATCH, "Run", _
                "The delegate result is not an object."
        End If
        If Len(mReturnObjectType) > 0 Then
            If StrComp(TypeName(result.InternalObjectValue), mReturnObjectType, _
                vbTextCompare) <> 0 Then
                RaiseContractError ERROR_TYPE_MISMATCH, "Run", _
                    "The delegate returned the wrong object type."
            End If
        End If
    Else
        If result.InternalHasObjectValue Then
            RaiseContractError ERROR_TYPE_MISMATCH, "Run", _
                "The delegate result is an object, not a scalar value."
        End If
        Set result = WidenScalarToward(result, mReturnType)
        actualType = VarTypeOf(result.InternalScalarValue)
        If actualType <> mReturnType Then
            RaiseContractError ERROR_TYPE_MISMATCH, "Run", _
                "Expected " & TypeNameForVarType(mReturnType) & ", received " & _
                TypeNameForVarType(actualType) & "."
        End If
    End If
    Set FinalizeDelegateResult = result
End Function

Private Function RunWorkbookProcedure( _
    ByVal procedureName As String, _
    ByRef values() As Variant, _
    ByVal argumentCount As Long _
) As Variant
    Select Case argumentCount
        Case 0
            RunWorkbookProcedure = Application.Run(procedureName)
        Case 1
            RunWorkbookProcedure = Application.Run( _
                procedureName, values(0))
        Case 2
            RunWorkbookProcedure = Application.Run( _
                procedureName, values(0), values(1))
        Case 3
            RunWorkbookProcedure = Application.Run( _
                procedureName, values(0), values(1), values(2))
        Case 4
            RunWorkbookProcedure = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3))
        Case 5
            RunWorkbookProcedure = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4))
        Case 6
            RunWorkbookProcedure = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5))
        Case 7
            RunWorkbookProcedure = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6))
        Case 8
            RunWorkbookProcedure = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7))
        Case 9
            RunWorkbookProcedure = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8))
        Case 10
            RunWorkbookProcedure = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9))
        Case 11
            RunWorkbookProcedure = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10))
        Case 12
            RunWorkbookProcedure = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11))
        Case 13
            RunWorkbookProcedure = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12))
        Case 14
            RunWorkbookProcedure = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13))
        Case 15
            RunWorkbookProcedure = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13), values(14))
        Case 16
            RunWorkbookProcedure = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13), values(14), values(15))
        Case 17
            RunWorkbookProcedure = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13), values(14), values(15), values(16))
        Case 18
            RunWorkbookProcedure = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13), values(14), values(15), values(16), _
                values(17))
        Case 19
            RunWorkbookProcedure = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13), values(14), values(15), values(16), _
                values(17), values(18))
        Case 20
            RunWorkbookProcedure = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13), values(14), values(15), values(16), _
                values(17), values(18), values(19))
        Case 21
            RunWorkbookProcedure = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13), values(14), values(15), values(16), _
                values(17), values(18), values(19), values(20))
        Case 22
            RunWorkbookProcedure = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13), values(14), values(15), values(16), _
                values(17), values(18), values(19), values(20), values(21))
        Case 23
            RunWorkbookProcedure = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13), values(14), values(15), values(16), _
                values(17), values(18), values(19), values(20), values(21), values(22))
        Case 24
            RunWorkbookProcedure = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13), values(14), values(15), values(16), _
                values(17), values(18), values(19), values(20), values(21), values(22), _
                values(23))
        Case 25
            RunWorkbookProcedure = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13), values(14), values(15), values(16), _
                values(17), values(18), values(19), values(20), values(21), values(22), _
                values(23), values(24))
        Case 26
            RunWorkbookProcedure = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13), values(14), values(15), values(16), _
                values(17), values(18), values(19), values(20), values(21), values(22), _
                values(23), values(24), values(25))
        Case 27
            RunWorkbookProcedure = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13), values(14), values(15), values(16), _
                values(17), values(18), values(19), values(20), values(21), values(22), _
                values(23), values(24), values(25), values(26))
        Case 28
            RunWorkbookProcedure = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13), values(14), values(15), values(16), _
                values(17), values(18), values(19), values(20), values(21), values(22), _
                values(23), values(24), values(25), values(26), values(27))
        Case 29
            RunWorkbookProcedure = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13), values(14), values(15), values(16), _
                values(17), values(18), values(19), values(20), values(21), values(22), _
                values(23), values(24), values(25), values(26), values(27), values(28))
        Case 30
            RunWorkbookProcedure = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13), values(14), values(15), values(16), _
                values(17), values(18), values(19), values(20), values(21), values(22), _
                values(23), values(24), values(25), values(26), values(27), values(28), _
                values(29))
        Case Else
            RaiseContractError ERROR_ARITY_MISMATCH, "Run", _
                "Workbook procedures support at most 30 arguments."
    End Select
End Function

Private Function NormalizeUnaryAction( _
    ByVal candidate As ROneCOne, _
    ByVal memberName As String _
) As ROneCOne
    RequireSequence Me, memberName
    ValidateAction candidate, memberName, False
    If candidate.Arity <> 1 And candidate.Arity <> -1 Then
        RaiseContractError ERROR_ARITY_MISMATCH, memberName, _
            "The Action must accept exactly one value."
    End If
    Set NormalizeUnaryAction = candidate
End Function

Private Function RunWorkbookProcedureObject( _
    ByVal procedureName As String, _
    ByRef values() As Variant, _
    ByVal argumentCount As Long _
) As Object
    Select Case argumentCount
        Case 0
            Set RunWorkbookProcedureObject = Application.Run(procedureName)
        Case 1
            Set RunWorkbookProcedureObject = Application.Run( _
                procedureName, values(0))
        Case 2
            Set RunWorkbookProcedureObject = Application.Run( _
                procedureName, values(0), values(1))
        Case 3
            Set RunWorkbookProcedureObject = Application.Run( _
                procedureName, values(0), values(1), values(2))
        Case 4
            Set RunWorkbookProcedureObject = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3))
        Case 5
            Set RunWorkbookProcedureObject = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4))
        Case 6
            Set RunWorkbookProcedureObject = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5))
        Case 7
            Set RunWorkbookProcedureObject = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6))
        Case 8
            Set RunWorkbookProcedureObject = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7))
        Case 9
            Set RunWorkbookProcedureObject = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8))
        Case 10
            Set RunWorkbookProcedureObject = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9))
        Case 11
            Set RunWorkbookProcedureObject = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10))
        Case 12
            Set RunWorkbookProcedureObject = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11))
        Case 13
            Set RunWorkbookProcedureObject = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12))
        Case 14
            Set RunWorkbookProcedureObject = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13))
        Case 15
            Set RunWorkbookProcedureObject = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13), values(14))
        Case 16
            Set RunWorkbookProcedureObject = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13), values(14), values(15))
        Case 17
            Set RunWorkbookProcedureObject = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13), values(14), values(15), values(16))
        Case 18
            Set RunWorkbookProcedureObject = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13), values(14), values(15), values(16), _
                values(17))
        Case 19
            Set RunWorkbookProcedureObject = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13), values(14), values(15), values(16), _
                values(17), values(18))
        Case 20
            Set RunWorkbookProcedureObject = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13), values(14), values(15), values(16), _
                values(17), values(18), values(19))
        Case 21
            Set RunWorkbookProcedureObject = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13), values(14), values(15), values(16), _
                values(17), values(18), values(19), values(20))
        Case 22
            Set RunWorkbookProcedureObject = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13), values(14), values(15), values(16), _
                values(17), values(18), values(19), values(20), values(21))
        Case 23
            Set RunWorkbookProcedureObject = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13), values(14), values(15), values(16), _
                values(17), values(18), values(19), values(20), values(21), values(22))
        Case 24
            Set RunWorkbookProcedureObject = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13), values(14), values(15), values(16), _
                values(17), values(18), values(19), values(20), values(21), values(22), _
                values(23))
        Case 25
            Set RunWorkbookProcedureObject = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13), values(14), values(15), values(16), _
                values(17), values(18), values(19), values(20), values(21), values(22), _
                values(23), values(24))
        Case 26
            Set RunWorkbookProcedureObject = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13), values(14), values(15), values(16), _
                values(17), values(18), values(19), values(20), values(21), values(22), _
                values(23), values(24), values(25))
        Case 27
            Set RunWorkbookProcedureObject = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13), values(14), values(15), values(16), _
                values(17), values(18), values(19), values(20), values(21), values(22), _
                values(23), values(24), values(25), values(26))
        Case 28
            Set RunWorkbookProcedureObject = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13), values(14), values(15), values(16), _
                values(17), values(18), values(19), values(20), values(21), values(22), _
                values(23), values(24), values(25), values(26), values(27))
        Case 29
            Set RunWorkbookProcedureObject = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13), values(14), values(15), values(16), _
                values(17), values(18), values(19), values(20), values(21), values(22), _
                values(23), values(24), values(25), values(26), values(27), values(28))
        Case 30
            Set RunWorkbookProcedureObject = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13), values(14), values(15), values(16), _
                values(17), values(18), values(19), values(20), values(21), values(22), _
                values(23), values(24), values(25), values(26), values(27), values(28), _
                values(29))
        Case Else
            RaiseContractError ERROR_ARITY_MISMATCH, "Run", _
                "Workbook procedures support at most 30 arguments."
    End Select
End Function

Private Function Unwrap(ByVal wrapped As ROneCOne) As Variant
    If wrapped.InternalHasObjectValue Then
        Set Unwrap = wrapped.InternalObjectValue
    Else
        Unwrap = wrapped.InternalScalarValue
    End If
End Function

Private Function CreateScalarResult(ByVal scalarValue As Variant) As ROneCOne
    Dim result As ROneCOne

    Set result = New ROneCOne
    result.ConfigureScalarValue scalarValue
    Set CreateScalarResult = result
End Function

Private Function CreateObjectResult(ByVal objectValue As Object) As ROneCOne
    Dim result As ROneCOne

    Set result = New ROneCOne
    result.ConfigureObjectValue objectValue
    Set CreateObjectResult = result
End Function

Private Function IsMemberOperator(ByVal operatorCode As Long) As Boolean
    Select Case operatorCode
        Case OP_MEMBER_VALUE, OP_MEMBER_OBJECT
            IsMemberOperator = True
        Case OP_SAFE_MEMBER_VALUE, OP_SAFE_MEMBER_OBJECT
            IsMemberOperator = True
    End Select
End Function

Private Function IsOrderingOperator(ByVal operatorCode As Long) As Boolean
    Select Case operatorCode
        Case QUERY_ORDER, QUERY_ORDER_DESCENDING
            IsOrderingOperator = True
        Case QUERY_ORDER_BY, QUERY_ORDER_BY_DESCENDING
            IsOrderingOperator = True
        Case QUERY_THEN_BY, QUERY_THEN_BY_DESCENDING
            IsOrderingOperator = True
    End Select
End Function

Private Function IsDescendingOrderingOperator( _
    ByVal operatorCode As Long _
) As Boolean
    IsDescendingOrderingOperator = _
        operatorCode = QUERY_ORDER_DESCENDING Or _
        operatorCode = QUERY_ORDER_BY_DESCENDING Or _
        operatorCode = QUERY_THEN_BY_DESCENDING
End Function

Private Function IsObjectMemberOperator(ByVal operatorCode As Long) As Boolean
    IsObjectMemberOperator = _
        operatorCode = OP_MEMBER_OBJECT Or _
        operatorCode = OP_SAFE_MEMBER_OBJECT
End Function

Private Function IsSafeMemberOperator(ByVal operatorCode As Long) As Boolean
    IsSafeMemberOperator = _
        operatorCode = OP_SAFE_MEMBER_VALUE Or _
        operatorCode = OP_SAFE_MEMBER_OBJECT
End Function

Private Function IsQuantifierOperator(ByVal operatorCode As Long) As Boolean
    Select Case operatorCode
        Case OP_ANY_MATCH, OP_ALL_MATCH, OP_NONE_MATCH
            IsQuantifierOperator = True
    End Select
End Function

Private Function SequenceContainsWrapped( _
    ByVal sequenceValue As ROneCOne, _
    ByVal candidate As ROneCOne, _
    Optional ByVal comparer As ROneCOne _
) As Boolean
    Dim values As Collection
    Dim wrapped As ROneCOne

    RequireSequence sequenceValue, "In"
    Set values = sequenceValue.MaterializeWrapped
    For Each wrapped In values
        If ValuesEqualWithComparer( _
            wrapped, candidate, comparer, "IsIn") Then
            SequenceContainsWrapped = True
            Exit Function
        End If
    Next wrapped
End Function

Private Function EvaluateNestedQuantifier( _
    ByVal nestedValue As ROneCOne, _
    ByVal predicate As ROneCOne, _
    ByVal operatorCode As Long _
) As Boolean
    Dim nestedObject As Object
    Dim sequenceValue As ROneCOne
    Dim values As Collection
    Dim wrapped As ROneCOne

    If Not nestedValue.InternalHasObjectValue Then
        RaiseContractError ERROR_TYPE_MISMATCH, "Nested quantifier", _
            "A nested quantifier requires an object-valued sequence member."
    End If
    Set nestedObject = nestedValue.InternalObjectValue
    If nestedObject Is Nothing Then
        EvaluateNestedQuantifier = (operatorCode <> OP_ANY_MATCH)
        Exit Function
    End If
    If Not TypeOf nestedObject Is ROneCOne Then
        RaiseContractError ERROR_TYPE_MISMATCH, "Nested quantifier", _
            "The nested member must return a ROneCOne sequence."
    End If
    Set sequenceValue = nestedObject
    RequireSequence sequenceValue, "Nested quantifier"
    Set values = sequenceValue.MaterializeWrapped

    If operatorCode = OP_ANY_MATCH Then
        If predicate Is Nothing Then
            EvaluateNestedQuantifier = (values.Count > 0)
            Exit Function
        End If
        For Each wrapped In values
            If InvokePredicate(predicate, wrapped, "AnyMatch") Then
                EvaluateNestedQuantifier = True
                Exit Function
            End If
        Next wrapped
        Exit Function
    End If

    If operatorCode = OP_ALL_MATCH Then
        EvaluateNestedQuantifier = True
        For Each wrapped In values
            If Not InvokePredicate(predicate, wrapped, "AllMatch") Then
                EvaluateNestedQuantifier = False
                Exit Function
            End If
        Next wrapped
        Exit Function
    End If

    EvaluateNestedQuantifier = True
    If predicate Is Nothing Then
        EvaluateNestedQuantifier = (values.Count = 0)
        Exit Function
    End If
    For Each wrapped In values
        If InvokePredicate(predicate, wrapped, "NoneMatch") Then
            EvaluateNestedQuantifier = False
            Exit Function
        End If
    Next wrapped
End Function

Private Function ReadMember( _
    ByVal targetValue As ROneCOne, _
    ByVal memberName As String, _
    ByVal returnsObject As Boolean, _
    ByVal nullSafe As Boolean _
) As ROneCOne
    Dim capturedDescription As String
    Dim returnedObject As Object
    Dim returnedValue As Variant
    Dim runtimeTarget As ROneCOne
    Dim target As Object

    If Not targetValue.InternalHasObjectValue Then
        RaiseContractError ERROR_MEMBER_ACCESS, "Member", _
            "Member access requires an object value."
    End If
    Set target = targetValue.InternalObjectValue
    If target Is Nothing Then
        If nullSafe Then
            If returnsObject Then
                Set ReadMember = CreateObjectResult(Nothing)
            Else
                Set ReadMember = CreateScalarResult(Null)
            End If
            Exit Function
        End If
        RaiseContractError ERROR_MEMBER_ACCESS, "Member", _
            "Cannot access '" & memberName & "' on Nothing."
    End If

    If TypeOf target Is ROneCOne Then
        Set runtimeTarget = target
        If runtimeTarget.InternalRole = ROLE_DATA_ROW Then
            Set ReadMember = runtimeTarget.DataFieldWrapped(memberName)
            Exit Function
        End If
    End If

    On Error GoTo AccessFailed
    If returnsObject Then
        Set returnedObject = CallByName(target, memberName, VbGet)
        Set ReadMember = CreateObjectResult(returnedObject)
    Else
        returnedValue = CallByName(target, memberName, VbGet)
        Set ReadMember = CreateScalarResult(returnedValue)
    End If
    On Error GoTo 0
    Exit Function

AccessFailed:
    capturedDescription = Err.Description
    On Error GoTo 0
    RaiseContractError ERROR_MEMBER_ACCESS, "Member", _
        "Cannot read '" & memberName & "' from " & TypeName(target) & _
        ". " & capturedDescription
End Function

Private Function HasNullScalar(ByVal wrapped As ROneCOne) As Boolean
    If wrapped Is Nothing Then Exit Function
    If wrapped.InternalHasObjectValue Then Exit Function
    HasNullScalar = IsNull(wrapped.InternalScalarValue)
End Function

Private Function RequireScalar( _
    ByVal wrapped As ROneCOne, _
    ByVal memberName As String _
) As Variant
    If wrapped Is Nothing Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            "The expression operand is missing."
    End If
    If wrapped.InternalHasObjectValue Then
        RaiseContractError ERROR_TYPE_MISMATCH, memberName, _
            "This expression requires scalar operands."
    End If
    RequireScalar = wrapped.InternalScalarValue
End Function

Private Function ValuesEqual( _
    ByVal leftValue As ROneCOne, _
    ByVal rightValue As ROneCOne _
) As Boolean
    If leftValue.InternalHasObjectValue Or rightValue.InternalHasObjectValue Then
        If Not leftValue.InternalHasObjectValue Or Not rightValue.InternalHasObjectValue Then
            ValuesEqual = False
        Else
            ValuesEqual = leftValue.InternalObjectValue Is rightValue.InternalObjectValue
        End If
    Else
        If IsNull(leftValue.InternalScalarValue) Or _
            IsNull(rightValue.InternalScalarValue) Then
            ValuesEqual = IsNull(leftValue.InternalScalarValue) And _
                IsNull(rightValue.InternalScalarValue)
        Else
            ValuesEqual = _
                (leftValue.InternalScalarValue = rightValue.InternalScalarValue)
        End If
    End If
End Function

Private Sub ValidateBindings( _
    ByVal node As ROneCOne, _
    ByVal parameters As Collection _
)
    Select Case node.InternalRole
        Case ROLE_VALUE
            Exit Sub
        Case ROLE_PARAMETER
            If Not CollectionContainsIdentity(parameters, node) Then
                RaiseContractError ERROR_UNBOUND_PARAMETER, "Lambda", _
                    "The expression contains a parameter not declared by this lambda."
            End If
        Case ROLE_EXPRESSION
            ValidateBindings node.InternalLeft, parameters
            If Not node.InternalRight Is Nothing Then
                ValidateBindings node.InternalRight, parameters
            End If
        Case Else
            RaiseContractError ERROR_INVALID_ARGUMENT, "Lambda", _
                "The lambda body must be an expression, parameter, or captured value."
    End Select
End Sub

Private Function RequireROneCOne( _
    ByVal candidate As Variant, _
    ByVal context As String _
) As ROneCOne
    Dim candidateObject As Object

    If Not IsObject(candidate) Then
        RaiseContractError ERROR_INVALID_ARGUMENT, context, _
            "Expected a ROneCOne value."
    End If
    Set candidateObject = candidate
    If candidateObject Is Nothing Then
        RaiseContractError ERROR_INVALID_ARGUMENT, context, _
            "The ROneCOne value cannot be Nothing."
    End If
    If Not TypeOf candidateObject Is ROneCOne Then
        RaiseContractError ERROR_INVALID_ARGUMENT, context, _
            "Expected a ROneCOne value."
    End If
    Set RequireROneCOne = candidateObject
End Function

Private Sub RequireExpression(ByVal candidate As ROneCOne, ByVal context As String)
    If candidate Is Nothing Then
        RaiseContractError ERROR_INVALID_ARGUMENT, context, _
            "The expression cannot be Nothing."
    End If
    Select Case candidate.InternalRole
        Case ROLE_VALUE, ROLE_PARAMETER, ROLE_EXPRESSION
            Exit Sub
        Case Else
            RaiseContractError ERROR_INVALID_OPERATION, context, _
                "This operation requires an expression value."
    End Select
End Sub

Private Sub RequireDelegate(ByVal candidate As ROneCOne, ByVal context As String)
    If candidate Is Nothing Then
        RaiseContractError ERROR_INVALID_ARGUMENT, context, _
            "The delegate cannot be Nothing."
    End If
    If IsDelegateRole(candidate.InternalRole) Then Exit Sub
    RaiseContractError ERROR_INVALID_OPERATION, context, _
        "This operation requires a delegate."
End Sub

Private Sub RequireEvent(ByVal candidate As ROneCOne, ByVal context As String)
    If candidate Is Nothing Then
        RaiseContractError ERROR_INVALID_ARGUMENT, context, _
            "The event cannot be Nothing."
    End If
    Select Case candidate.InternalRole
        Case ROLE_EVENT
            Exit Sub
    End Select
    RaiseContractError ERROR_INVALID_OPERATION, context, _
        "This operation requires an event."
End Sub

Private Sub RequireTry(ByVal candidate As ROneCOne, ByVal context As String)
    If candidate Is Nothing Then
        RaiseContractError ERROR_INVALID_ARGUMENT, context, _
            "The Try operation cannot be Nothing."
    End If
    Select Case candidate.InternalRole
        Case ROLE_TRY
            Exit Sub
    End Select
    RaiseContractError ERROR_INVALID_OPERATION, context, _
        "This operation requires a Try operation."
End Sub

Private Sub RequireError(ByVal candidate As ROneCOne, ByVal context As String)
    If candidate Is Nothing Then
        RaiseContractError ERROR_INVALID_ARGUMENT, context, _
            "The exception cannot be Nothing."
    End If
    Select Case candidate.InternalRole
        Case ROLE_ERROR
            Exit Sub
    End Select
    RaiseContractError ERROR_INVALID_OPERATION, context, _
        "This property requires a captured exception."
End Sub

Private Function CollectionContainsIdentity( _
    ByVal values As Collection, _
    ByVal candidate As ROneCOne _
) As Boolean
    Dim index As Long

    For index = 1 To values.Count
        If values.Item(index) Is candidate Then
            CollectionContainsIdentity = True
            Exit Function
        End If
    Next index
End Function

Private Function TypeNameForVarType(ByVal valueType As VbVarType) As String
    Select Case valueType
        Case vbEmpty
            TypeNameForVarType = "Empty"
        Case vbNull
            TypeNameForVarType = "Null"
        Case vbInteger
            TypeNameForVarType = "Integer"
        Case vbLong
            TypeNameForVarType = "Long"
        Case vbSingle
            TypeNameForVarType = "Single"
        Case vbDouble
            TypeNameForVarType = "Double"
        Case vbCurrency
            TypeNameForVarType = "Currency"
        Case vbDate
            TypeNameForVarType = "Date"
        Case vbString
            TypeNameForVarType = "String"
        Case vbObject
            TypeNameForVarType = "Object"
        Case vbError
            TypeNameForVarType = "Error"
        Case vbBoolean
            TypeNameForVarType = "Boolean"
        Case vbVariant
            TypeNameForVarType = "Variant"
        Case vbDecimal
            TypeNameForVarType = "Decimal"
        Case vbByte
            TypeNameForVarType = "Byte"
        Case vbLongLong
            TypeNameForVarType = "LongLong"
        Case Else
            TypeNameForVarType = "VarType " & CStr(valueType)
    End Select
End Function

' ---------------------------------------------------------------------------
' DateTime and TimeSpan core: instants are milliseconds since the Unix epoch
' in universal time, durations are signed milliseconds, and every calendar
' operation runs on exact integer Double arithmetic to keep clock components
' stable across the supported years 0100 through 9999.
' ---------------------------------------------------------------------------

Private Function EpochDate() As Date
    EpochDate = DateSerial(1970, 1, 1)
End Function

Private Function DateToMs(ByVal value As Date) As Double
    Dim dateOnly As Date
    Dim dayCount As Double
    Dim msOfDay As Double

    dateOnly = DateSerial(VBA.Year(value), VBA.Month(value), VBA.Day(value))
    dayCount = DateDiff("d", EpochDate(), dateOnly)
    ' Pre-1900 VBA dates carry the time of day as a negative fraction, so the
    ' magnitude of the fraction is the clock time in both eras.
    msOfDay = Int(Abs(CDbl(value) - CDbl(dateOnly)) * MS_PER_DAY + 0.5)
    DateToMs = dayCount * MS_PER_DAY + msOfDay
End Function

Private Function MsToDate(ByVal msValue As Double) As Date
    Dim clockDate As Date
    Dim msOfDay As Long

    MsToClockParts msValue, clockDate, msOfDay
    If CDbl(clockDate) < 0 Then
        MsToDate = CDate(CDbl(clockDate) - msOfDay / MS_PER_DAY)
    Else
        MsToDate = CDate(CDbl(clockDate) + msOfDay / MS_PER_DAY)
    End If
End Function

Private Function FloorDivDouble(ByVal value As Double, ByVal divisor As Double) As Double
    Dim result As Double

    result = Int(value / divisor)
    ' Floating division can land one step off at exact multiples; the integer
    ' comparisons below are exact because every quantity stays under 2^53.
    If (result + 1#) * divisor <= value Then result = result + 1#
    If result * divisor > value Then result = result - 1#
    FloorDivDouble = result
End Function

Private Function RoundToMs(ByVal value As Double) As Double
    If value >= 0 Then
        RoundToMs = Int(value + 0.5)
    Else
        RoundToMs = -Int(-value + 0.5)
    End If
End Function

Private Sub MsToClockParts( _
    ByVal msValue As Double, _
    ByRef clockDate As Date, _
    ByRef msOfDay As Long _
)
    Dim dayCount As Double

    dayCount = FloorDivDouble(msValue, MS_PER_DAY)
    clockDate = DateAdd("d", dayCount, EpochDate())
    msOfDay = CLng(msValue - dayCount * MS_PER_DAY)
End Sub

Private Function DateTimeClockMs() As Double
    DateTimeClockMs = mDateTimeMs + CDbl(mDateTimeOffsetMinutes) * MS_PER_MINUTE
End Function

Private Function SystemTimeToMs(ByRef value As SYSTEMTIME) As Double
    SystemTimeToMs = DateToMs(DateSerial(value.wYear, value.wMonth, value.wDay)) + _
        CDbl(value.wHour) * MS_PER_HOUR + CDbl(value.wMinute) * MS_PER_MINUTE + _
        CDbl(value.wSecond) * MS_PER_SECOND + CDbl(value.wMilliseconds)
End Function

Private Function MsToSystemTime( _
    ByVal msValue As Double, _
    ByVal memberName As String _
) As SYSTEMTIME
    Dim clockDate As Date
    Dim msOfDay As Long

    MsToClockParts msValue, clockDate, msOfDay
    If VBA.Year(clockDate) < 1601 Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            "Windows time zone conversion supports years 1601 and later."
    End If
    MsToSystemTime.wYear = CInt(VBA.Year(clockDate))
    MsToSystemTime.wMonth = CInt(VBA.Month(clockDate))
    MsToSystemTime.wDay = CInt(VBA.Day(clockDate))
    MsToSystemTime.wHour = CInt(msOfDay \ 3600000)
    MsToSystemTime.wMinute = CInt((msOfDay \ 60000) Mod 60)
    MsToSystemTime.wSecond = CInt((msOfDay \ 1000) Mod 60)
    MsToSystemTime.wMilliseconds = CInt(msOfDay Mod 1000)
End Function

Private Function UtcMsToLocalMs( _
    ByVal utcMs As Double, _
    ByVal memberName As String _
) As Double
    Dim localTime As SYSTEMTIME
    Dim utcTime As SYSTEMTIME

    utcTime = MsToSystemTime(utcMs, memberName)
    If SystemTimeToTzSpecificLocalTime(0, utcTime, localTime) = 0 Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            "Windows rejected the universal to local time conversion."
    End If
    UtcMsToLocalMs = SystemTimeToMs(localTime)
End Function

Private Function LocalClockMsToUtcMs( _
    ByVal clockMs As Double, _
    ByVal memberName As String _
) As Double
    Dim localTime As SYSTEMTIME
    Dim utcTime As SYSTEMTIME

    localTime = MsToSystemTime(clockMs, memberName)
    If TzSpecificLocalTimeToSystemTime(0, localTime, utcTime) = 0 Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            "Windows rejected the local to universal time conversion."
    End If
    LocalClockMsToUtcMs = SystemTimeToMs(utcTime)
End Function

Private Function LocalOffsetMinutesForUtc( _
    ByVal utcMs As Double, _
    ByVal memberName As String _
) As Long
    LocalOffsetMinutesForUtc = CLng( _
        (UtcMsToLocalMs(utcMs, memberName) - utcMs) / MS_PER_MINUTE)
End Function

Private Function BuildDateTime( _
    ByVal msValue As Double, _
    ByVal offsetMinutes As Long _
) As ROneCOne
    Dim instance As ROneCOne

    If msValue < MIN_DATETIME_MS Or msValue > MAX_DATETIME_MS Or _
        msValue + CDbl(offsetMinutes) * MS_PER_MINUTE < MIN_DATETIME_MS Or _
        msValue + CDbl(offsetMinutes) * MS_PER_MINUTE > MAX_DATETIME_MS Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "DateTime", _
            "The instant falls outside the supported years 0100 through 9999."
    End If
    Set instance = New ROneCOne
    instance.ConfigureDateTime msValue, offsetMinutes
    Set BuildDateTime = instance
End Function

Private Function BuildTimeSpan(ByVal msValue As Double) As ROneCOne
    Dim instance As ROneCOne

    Set instance = New ROneCOne
    instance.ConfigureTimeSpan msValue
    Set BuildTimeSpan = instance
End Function

Private Function AddClockInterval( _
    ByVal intervalCode As String, _
    ByVal amount As Long, _
    ByVal memberName As String _
) As ROneCOne
    Dim clockDate As Date
    Dim moved As Date
    Dim msOfDay As Long

    MsToClockParts DateTimeClockMs(), clockDate, msOfDay
    On Error GoTo Failed
    moved = DateAdd(intervalCode, amount, clockDate)
    On Error GoTo 0
    Set AddClockInterval = BuildDateTime(DateToMs(moved) + CDbl(msOfDay) - _
        CDbl(mDateTimeOffsetMinutes) * MS_PER_MINUTE, mDateTimeOffsetMinutes)
    Exit Function
Failed:
    RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
        "The result falls outside the supported years 0100 through 9999."
End Function

Private Function DateTimeIsoText() As String
    Dim clockDate As Date
    Dim msOfDay As Long
    Dim result As String

    MsToClockParts DateTimeClockMs(), clockDate, msOfDay
    result = ClockStampText(clockDate, msOfDay)
    If (msOfDay Mod 1000) <> 0 Then
        result = result & "." & PadNumber(msOfDay Mod 1000, 3)
    End If
    DateTimeIsoText = result & OffsetSuffixText(mDateTimeOffsetMinutes)
End Function

Private Function ClockStampText( _
    ByVal clockDate As Date, _
    ByVal msOfDay As Long _
) As String
    ClockStampText = PadNumber(VBA.Year(clockDate), 4) & "-" & _
        PadNumber(VBA.Month(clockDate), 2) & "-" & _
        PadNumber(VBA.Day(clockDate), 2) & "T" & _
        PadNumber(msOfDay \ 3600000, 2) & ":" & _
        PadNumber((msOfDay \ 60000) Mod 60, 2) & ":" & _
        PadNumber((msOfDay \ 1000) Mod 60, 2)
End Function

Private Function OffsetSuffixText(ByVal offsetMinutes As Long) As String
    Dim magnitude As Long
    Dim signText As String

    If offsetMinutes = 0 Then
        OffsetSuffixText = "Z"
        Exit Function
    End If
    magnitude = Abs(offsetMinutes)
    signText = "+"
    If offsetMinutes < 0 Then signText = "-"
    OffsetSuffixText = signText & PadNumber(magnitude \ 60, 2) & ":" & _
        PadNumber(magnitude Mod 60, 2)
End Function

Private Function FormatDateTimeTokens(ByVal pattern As String) As String
    Dim clockDate As Date
    Dim msOfDay As Long

    MsToClockParts DateTimeClockMs(), clockDate, msOfDay
    FormatDateTimeTokens = FormatClockPattern(pattern, clockDate, msOfDay)
End Function

Private Function FormatClockPattern( _
    ByVal pattern As String, _
    ByVal clockDate As Date, _
    ByVal msOfDay As Long _
) As String
    Dim index As Long
    Dim probe As String
    Dim result As String
    Dim runLength As Long

    index = 1
    Do While index <= Len(pattern)
        probe = Mid$(pattern, index, 1)
        runLength = 1
        Do While Mid$(pattern, index + runLength, 1) = probe
            runLength = runLength + 1
        Loop
        If (probe >= "a" And probe <= "z") Or (probe >= "A" And probe <= "Z") Then
            result = result & DateTimeTokenText(probe, runLength, clockDate, msOfDay)
        Else
            result = result & String$(runLength, probe)
        End If
        index = index + runLength
    Loop
    FormatClockPattern = result
End Function

Private Function DateTimeTokenText( _
    ByVal token As String, _
    ByVal runLength As Long, _
    ByVal clockDate As Date, _
    ByVal msOfDay As Long _
) As String
    Select Case token & CStr(runLength)
        Case "y4"
            DateTimeTokenText = PadNumber(VBA.Year(clockDate), 4)
        Case "M2"
            DateTimeTokenText = PadNumber(VBA.Month(clockDate), 2)
        Case "d2"
            DateTimeTokenText = PadNumber(VBA.Day(clockDate), 2)
        Case "H2"
            DateTimeTokenText = PadNumber(msOfDay \ 3600000, 2)
        Case "m2"
            DateTimeTokenText = PadNumber((msOfDay \ 60000) Mod 60, 2)
        Case "s2"
            DateTimeTokenText = PadNumber((msOfDay \ 1000) Mod 60, 2)
        Case "f3"
            DateTimeTokenText = PadNumber(msOfDay Mod 1000, 3)
        Case Else
            RaiseFormatError "The format token '" & String$(runLength, token) & _
                "' is not supported; use yyyy MM dd HH mm ss fff."
    End Select
End Function

Private Function TimeSpanText() As String
    Dim dayPart As Double
    Dim magnitude As Double
    Dim msPart As Double
    Dim result As String

    magnitude = Abs(mTimeSpanMs)
    dayPart = Fix(magnitude / MS_PER_DAY)
    result = PadNumber(CLng(Fix(magnitude / MS_PER_HOUR) - dayPart * 24#), 2) & _
        ":" & PadNumber(CLng(Fix(magnitude / MS_PER_MINUTE) - _
        Fix(magnitude / MS_PER_HOUR) * 60#), 2) & _
        ":" & PadNumber(CLng(Fix(magnitude / MS_PER_SECOND) - _
        Fix(magnitude / MS_PER_MINUTE) * 60#), 2)
    If dayPart > 0 Then result = CStr(dayPart) & "." & result
    msPart = magnitude - Fix(magnitude / MS_PER_SECOND) * 1000#
    If msPart > 0 Then result = result & "." & PadNumber(CLng(msPart), 3)
    If mTimeSpanMs < 0 Then result = "-" & result
    TimeSpanText = result
End Function

Private Function ParseDateTimeText( _
    ByVal text As String, _
    ByVal memberName As String _
) As ROneCOne
    Dim clockMs As Double
    Dim hasOffset As Boolean
    Dim offsetMinutes As Long
    Dim utcMs As Double

    If Not TryParseIsoDateTime(text, clockMs, offsetMinutes, hasOffset) Then
        RaiseFormatError "The text is not ISO 8601 date and time: " & _
            ClipForMessage(text)
    End If
    If hasOffset Then
        utcMs = clockMs - CDbl(offsetMinutes) * MS_PER_MINUTE
    Else
        utcMs = LocalClockMsToUtcMs(clockMs, memberName)
        offsetMinutes = CLng((clockMs - utcMs) / MS_PER_MINUTE)
    End If
    Set ParseDateTimeText = BuildDateTime(utcMs, offsetMinutes)
End Function

Private Function TryParseIsoDateTime( _
    ByVal text As String, _
    ByRef clockMs As Double, _
    ByRef offsetMinutes As Long, _
    ByRef hasOffset As Boolean _
) As Boolean
    Dim clockDate As Date
    Dim dayPart As Long
    Dim fractionMs As Long
    Dim hourPart As Long
    Dim minutePart As Long
    Dim monthPart As Long
    Dim offsetHours As Long
    Dim offsetMins As Long
    Dim offsetSign As Long
    Dim parsed As String
    Dim position As Long
    Dim probe As String
    Dim secondPart As Long
    Dim textLength As Long
    Dim yearPart As Long

    hasOffset = False
    offsetMinutes = 0
    parsed = Trim$(text)
    textLength = Len(parsed)
    position = 1
    If Not TryReadDigits(parsed, position, 4, yearPart) Then Exit Function
    If yearPart < 100 Then Exit Function
    If Not SkipExpectedChar(parsed, position, "-") Then Exit Function
    If Not TryReadDigits(parsed, position, 2, monthPart) Then Exit Function
    If Not SkipExpectedChar(parsed, position, "-") Then Exit Function
    If Not TryReadDigits(parsed, position, 2, dayPart) Then Exit Function
    If monthPart < 1 Or monthPart > 12 Then Exit Function
    If dayPart < 1 Or dayPart > 31 Then Exit Function
    clockDate = DateSerial(yearPart, monthPart, dayPart)
    If VBA.Year(clockDate) <> yearPart Or VBA.Month(clockDate) <> monthPart Or _
        VBA.Day(clockDate) <> dayPart Then Exit Function
    If position <= textLength Then
        probe = Mid$(parsed, position, 1)
        If probe <> "T" And probe <> "t" And probe <> " " Then Exit Function
        position = position + 1
        If Not TryReadDigits(parsed, position, 2, hourPart) Then Exit Function
        If Not SkipExpectedChar(parsed, position, ":") Then Exit Function
        If Not TryReadDigits(parsed, position, 2, minutePart) Then Exit Function
        If hourPart > 23 Or minutePart > 59 Then Exit Function
        If PeekChar(parsed, position) = ":" Then
            position = position + 1
            If Not TryReadDigits(parsed, position, 2, secondPart) Then Exit Function
            If secondPart > 59 Then Exit Function
            If PeekChar(parsed, position) = "." Then
                position = position + 1
                If Not TryReadFractionMs(parsed, position, fractionMs) Then
                    Exit Function
                End If
            End If
        End If
        probe = PeekChar(parsed, position)
        If probe = "Z" Or probe = "z" Then
            hasOffset = True
            position = position + 1
        ElseIf probe = "+" Or probe = "-" Then
            offsetSign = 1
            If probe = "-" Then offsetSign = -1
            position = position + 1
            If Not TryReadDigits(parsed, position, 2, offsetHours) Then
                Exit Function
            End If
            If PeekChar(parsed, position) = ":" Then position = position + 1
            If position <= textLength Then
                If Not TryReadDigits(parsed, position, 2, offsetMins) Then
                    Exit Function
                End If
            End If
            If offsetHours > 14 Or offsetMins > 59 Then Exit Function
            offsetMinutes = offsetSign * (offsetHours * 60 + offsetMins)
            If Abs(offsetMinutes) > 840 Then Exit Function
            hasOffset = True
        End If
    End If
    If position <= textLength Then Exit Function
    clockMs = DateToMs(clockDate) + CDbl(hourPart) * MS_PER_HOUR + _
        CDbl(minutePart) * MS_PER_MINUTE + CDbl(secondPart) * MS_PER_SECOND + _
        fractionMs
    TryParseIsoDateTime = True
End Function

Private Function ParseTimeSpanText(ByVal text As String) As ROneCOne
    Dim msValue As Double

    If Not TryParseTimeSpanCore(Trim$(text), msValue) Then
        RaiseFormatError "The text is not a d.hh:mm:ss duration: " & _
            ClipForMessage(text)
    End If
    Set ParseTimeSpanText = BuildTimeSpan(msValue)
End Function

Private Function TryParseTimeSpanCore( _
    ByVal text As String, _
    ByRef msValue As Double _
) As Boolean
    Dim dayCount As Double
    Dim firstRun As String
    Dim fractionMs As Long
    Dim hourPart As Long
    Dim minutePart As Long
    Dim negative As Boolean
    Dim position As Long
    Dim runStart As Long
    Dim secondPart As Long
    Dim textLength As Long

    textLength = Len(text)
    position = 1
    If PeekChar(text, position) = "-" Then
        negative = True
        position = position + 1
    End If
    runStart = position
    Do While position <= textLength
        If Mid$(text, position, 1) < "0" Or Mid$(text, position, 1) > "9" Then
            Exit Do
        End If
        position = position + 1
    Loop
    If position = runStart Then Exit Function
    firstRun = Mid$(text, runStart, position - runStart)
    If Len(firstRun) > 7 Then Exit Function
    If PeekChar(text, position) = "." Then
        dayCount = CDbl(firstRun)
        position = position + 1
        If Not TryReadDigits(text, position, 2, hourPart) Then Exit Function
    Else
        If Len(firstRun) > 2 Then Exit Function
        hourPart = CLng(firstRun)
    End If
    If Not SkipExpectedChar(text, position, ":") Then Exit Function
    If Not TryReadDigits(text, position, 2, minutePart) Then Exit Function
    If PeekChar(text, position) = ":" Then
        position = position + 1
        If Not TryReadDigits(text, position, 2, secondPart) Then Exit Function
        If PeekChar(text, position) = "." Then
            position = position + 1
            If Not TryReadFractionMs(text, position, fractionMs) Then
                Exit Function
            End If
        End If
    End If
    If position <= textLength Then Exit Function
    If hourPart > 23 Or minutePart > 59 Or secondPart > 59 Then Exit Function
    msValue = dayCount * MS_PER_DAY + CDbl(hourPart) * MS_PER_HOUR + _
        CDbl(minutePart) * MS_PER_MINUTE + CDbl(secondPart) * MS_PER_SECOND + _
        fractionMs
    If negative Then msValue = -msValue
    TryParseTimeSpanCore = True
End Function

Private Function TryReadDigits( _
    ByVal text As String, _
    ByRef position As Long, _
    ByVal digitCount As Long, _
    ByRef result As Long _
) As Boolean
    Dim index As Long
    Dim probe As String

    If position + digitCount - 1 > Len(text) Then Exit Function
    result = 0
    For index = 1 To digitCount
        probe = Mid$(text, position, 1)
        If probe < "0" Or probe > "9" Then Exit Function
        result = result * 10 + (Asc(probe) - 48)
        position = position + 1
    Next index
    TryReadDigits = True
End Function

Private Function TryReadFractionMs( _
    ByVal text As String, _
    ByRef position As Long, _
    ByRef result As Long _
) As Boolean
    Dim digitCount As Long
    Dim digits As String
    Dim probe As String

    Do While position <= Len(text)
        probe = Mid$(text, position, 1)
        If probe < "0" Or probe > "9" Then Exit Do
        If digitCount < 3 Then digits = digits & probe
        digitCount = digitCount + 1
        position = position + 1
    Loop
    If digitCount = 0 Or digitCount > 7 Then Exit Function
    Do While Len(digits) < 3
        digits = digits & "0"
    Loop
    result = CLng(digits)
    TryReadFractionMs = True
End Function

Private Function SkipExpectedChar( _
    ByVal text As String, _
    ByRef position As Long, _
    ByVal expected As String _
) As Boolean
    If position > Len(text) Then Exit Function
    If Mid$(text, position, 1) <> expected Then Exit Function
    position = position + 1
    SkipExpectedChar = True
End Function

Private Function PeekChar(ByVal text As String, ByVal position As Long) As String
    If position > Len(text) Then Exit Function
    PeekChar = Mid$(text, position, 1)
End Function

Private Function PadNumber(ByVal value As Long, ByVal width As Long) As String
    PadNumber = CStr(value)
    If Len(PadNumber) < width Then
        PadNumber = String$(width - Len(PadNumber), "0") & PadNumber
    End If
End Function

Private Function ClipForMessage(ByVal text As String) As String
    If Len(text) > 60 Then
        ClipForMessage = Left$(text, 57) & "..."
    Else
        ClipForMessage = text
    End If
End Function

Private Sub RaiseFormatError(ByVal message As String)
    Err.Raise ERROR_FORMAT, "ROneCOne.FormatException", message
End Sub

' ---------------------------------------------------------------------------
' Composite formatting core: String.Format's {index,alignment:format} grammar
' with invariant text output, shared by Strings.Format and AppendFormat.
' ---------------------------------------------------------------------------

Private Function FormatCompositeCore( _
    ByVal template As String, _
    ByRef values() As Variant _
) As String
    Dim alignment As Long
    Dim argIndex As Long
    Dim builder As JsonTextBuilder
    Dim closeAt As Long
    Dim formatText As String
    Dim hasAlignment As Boolean
    Dim index As Long
    Dim item As String
    Dim piece As String
    Dim probe As String
    Dim templateLength As Long

    JsonBuilderInit builder, Len(template) + 32
    templateLength = Len(template)
    index = 1
    Do While index <= templateLength
        probe = Mid$(template, index, 1)
        If probe = "{" Then
            If Mid$(template, index + 1, 1) = "{" Then
                JsonBuilderAppend builder, "{"
                index = index + 2
            Else
                closeAt = InStr(index + 1, template, "}")
                If closeAt = 0 Then
                    RaiseFormatError _
                        "The format template has an unclosed '{' item."
                End If
                item = Mid$(template, index + 1, closeAt - index - 1)
                ParseFormatItem item, argIndex, hasAlignment, alignment, _
                    formatText
                If argIndex < LBound(values) Or argIndex > UBound(values) Then
                    RaiseFormatError "The format item {" & CStr(argIndex) & _
                        "} has no matching argument."
                End If
                piece = FormatArgText(values(argIndex), formatText)
                If hasAlignment Then piece = AlignFormatPiece(piece, alignment)
                JsonBuilderAppend builder, piece
                index = closeAt + 1
            End If
        ElseIf probe = "}" Then
            If Mid$(template, index + 1, 1) = "}" Then
                JsonBuilderAppend builder, "}"
                index = index + 2
            Else
                RaiseFormatError "The format template has a stray '}'."
            End If
        Else
            JsonBuilderAppend builder, probe
            index = index + 1
        End If
    Loop
    FormatCompositeCore = JsonBuilderText(builder)
End Function

Private Sub ParseFormatItem( _
    ByVal item As String, _
    ByRef argIndex As Long, _
    ByRef hasAlignment As Boolean, _
    ByRef alignment As Long, _
    ByRef formatText As String _
)
    Dim colonAt As Long
    Dim commaAt As Long
    Dim head As String

    hasAlignment = False
    alignment = 0
    formatText = vbNullString
    colonAt = InStr(1, item, ":")
    If colonAt > 0 Then
        formatText = Mid$(item, colonAt + 1)
        head = Left$(item, colonAt - 1)
    Else
        head = item
    End If
    commaAt = InStr(1, head, ",")
    If commaAt > 0 Then
        hasAlignment = True
        alignment = RequireFormatNumber(Mid$(head, commaAt + 1), True)
        head = Left$(head, commaAt - 1)
    End If
    argIndex = RequireFormatNumber(head, False)
End Sub

Private Function RequireFormatNumber( _
    ByVal text As String, _
    ByVal allowSign As Boolean _
) As Long
    Dim index As Long
    Dim probe As String
    Dim startAt As Long

    text = Trim$(text)
    startAt = 1
    If allowSign And Left$(text, 1) = "-" Then startAt = 2
    If Len(text) < startAt Then
        RaiseFormatError "A format item needs a numeric index and alignment."
    End If
    For index = startAt To Len(text)
        probe = Mid$(text, index, 1)
        If probe < "0" Or probe > "9" Then
            RaiseFormatError "The format item part '" & text & _
                "' is not a number."
        End If
    Next index
    RequireFormatNumber = CLng(text)
End Function

Private Function AlignFormatPiece( _
    ByVal piece As String, _
    ByVal alignment As Long _
) As String
    Dim padCount As Long

    padCount = Abs(alignment) - Len(piece)
    If padCount <= 0 Then
        AlignFormatPiece = piece
    ElseIf alignment > 0 Then
        AlignFormatPiece = Space$(padCount) & piece
    Else
        AlignFormatPiece = piece & Space$(padCount)
    End If
End Function

Private Function FormatArgText( _
    ByVal value As Variant, _
    ByVal formatText As String _
) As String
    Dim clockDate As Date
    Dim msOfDay As Long
    Dim runtimeValue As ROneCOne

    If IsObject(value) Then
        If value Is Nothing Then Exit Function
        If TypeOf value Is ROneCOne Then
            Set runtimeValue = value
            Select Case runtimeValue.InternalRole
                Case ROLE_DATETIME
                    FormatArgText = runtimeValue.ToString(formatText)
                Case ROLE_TIMESPAN
                    If Len(formatText) > 0 Then
                        RaiseFormatError _
                            "A TimeSpan argument accepts no format."
                    End If
                    FormatArgText = runtimeValue.ToString()
                Case Else
                    RaiseContractError ERROR_TYPE_MISMATCH, "Format", _
                        "This ROneCOne role cannot be formatted as text."
            End Select
            Exit Function
        End If
        RaiseContractError ERROR_TYPE_MISMATCH, "Format", _
            "Objects other than ROneCOne values cannot be formatted."
    End If
    If IsEmpty(value) Or IsNull(value) Then Exit Function
    If VarTypeOf(value) = vbDate Then
        MsToClockParts DateToMs(CDate(value)), clockDate, msOfDay
        If Len(formatText) = 0 Then
            FormatArgText = ClockStampText(clockDate, msOfDay)
        Else
            FormatArgText = FormatClockPattern(formatText, clockDate, msOfDay)
        End If
        Exit Function
    End If
    If VarTypeOf(value) = vbBoolean Then
        If value Then
            FormatArgText = "True"
        Else
            FormatArgText = "False"
        End If
        Exit Function
    End If
    If VarTypeOf(value) = vbString Then
        ' Strings pass through untouched, exactly as String.Format treats
        ' arguments that do not implement IFormattable.
        FormatArgText = value
        Exit Function
    End If
    If IsNumeric(value) Then
        FormatArgText = NumericFormatText(CDbl(value), formatText)
        Exit Function
    End If
    RaiseContractError ERROR_TYPE_MISMATCH, "Format", _
        "The argument type " & TypeName(value) & " cannot be formatted."
End Function

Private Function NumericFormatText( _
    ByVal value As Double, _
    ByVal formatText As String _
) As String
    Dim digitCount As Long
    Dim hasDigits As Boolean
    Dim index As Long
    Dim letter As String
    Dim probe As String

    If Len(formatText) = 0 Then
        NumericFormatText = JsonNumberText(value)
        Exit Function
    End If
    letter = Left$(formatText, 1)
    For index = 2 To Len(formatText)
        probe = Mid$(formatText, index, 1)
        If probe < "0" Or probe > "9" Then
            RaiseFormatError "The number format '" & formatText & _
                "' is not supported; use G, N, F, D, X, or P."
        End If
        hasDigits = True
    Next index
    If hasDigits Then digitCount = CLng(Mid$(formatText, 2))
    Select Case letter
        Case "G", "g"
            If hasDigits Then
                RaiseFormatError "The G format accepts no digit count."
            End If
            NumericFormatText = JsonNumberText(value)
        Case "N", "n"
            If Not hasDigits Then digitCount = 2
            NumericFormatText = FixedPointText(value, digitCount, True)
        Case "F", "f"
            If Not hasDigits Then digitCount = 2
            NumericFormatText = FixedPointText(value, digitCount, False)
        Case "P", "p"
            If Not hasDigits Then digitCount = 2
            NumericFormatText = FixedPointText(value * 100#, digitCount, _
                True) & "%"
        Case "D", "d"
            If value <> Fix(value) Then
                RaiseFormatError "The D format requires a whole number."
            End If
            If Not hasDigits Then digitCount = 1
            NumericFormatText = SignedPaddedInteger(value, digitCount)
        Case "X", "x"
            If value <> Fix(value) Then
                RaiseFormatError "The X format requires a whole number."
            End If
            NumericFormatText = HexIntegerText(value, digitCount, _
                letter = "X")
        Case Else
            RaiseFormatError "The number format '" & formatText & _
                "' is not supported; use G, N, F, D, X, or P."
    End Select
End Function

Private Function FixedPointText( _
    ByVal value As Double, _
    ByVal decimalCount As Long, _
    ByVal grouped As Boolean _
) As String
    Dim fracText As String
    Dim intText As String
    Dim magnitude As Double
    Dim scaled As Double
    Dim scaleFactor As Double
    Dim text As String

    If decimalCount < 0 Or decimalCount > 15 Then
        RaiseFormatError "Decimal counts run 0 through 15."
    End If
    magnitude = Abs(value)
    scaleFactor = 10# ^ decimalCount
    scaled = Int(magnitude * scaleFactor + 0.5)
    If scaled >= 1E+15 Then
        RaiseFormatError "The value is too large for fixed-point text."
    End If
    text = CStr(scaled)
    If Len(text) < decimalCount + 1 Then
        text = String$(decimalCount + 1 - Len(text), "0") & text
    End If
    If decimalCount > 0 Then
        fracText = "." & Right$(text, decimalCount)
        intText = Left$(text, Len(text) - decimalCount)
    Else
        intText = text
    End If
    If grouped Then intText = GroupThousands(intText)
    FixedPointText = intText & fracText
    If value < 0 And scaled > 0 Then FixedPointText = "-" & FixedPointText
End Function

Private Function GroupThousands(ByVal digits As String) As String
    Dim cursor As Long

    GroupThousands = digits
    cursor = Len(digits) - 3
    Do While cursor >= 1
        GroupThousands = Left$(GroupThousands, cursor) & "," & _
            Mid$(GroupThousands, cursor + 1)
        cursor = cursor - 3
    Loop
End Function

Private Function SignedPaddedInteger( _
    ByVal value As Double, _
    ByVal width As Long _
) As String
    Dim text As String

    If Abs(value) >= 1E+15 Then
        RaiseFormatError "The value is too large for integer text."
    End If
    text = CStr(Abs(value))
    If Len(text) < width Then
        text = String$(width - Len(text), "0") & text
    End If
    If value < 0 Then text = "-" & text
    SignedPaddedInteger = text
End Function

Private Function HexIntegerText( _
    ByVal value As Double, _
    ByVal width As Long, _
    ByVal upperCase As Boolean _
) As String
    Dim text As String

    If Abs(value) > 9.2E+18 Then
        RaiseFormatError "The value is too large for hexadecimal text."
    End If
    text = Hex$(CLngLng(value))
    If Not upperCase Then text = LCase$(text)
    If Len(text) < width Then
        text = String$(width - Len(text), "0") & text
    End If
    HexIntegerText = text
End Function

Private Function GuidBytesText(ByRef guidBytes() As Byte) As String
    ' A GUID's first three groups are little-endian on the wire, so the text
    ' form reads their bytes back to front, exactly like Guid.ToString.
    GuidBytesText = LCase$( _
        HexByteText(guidBytes(3)) & HexByteText(guidBytes(2)) & _
        HexByteText(guidBytes(1)) & HexByteText(guidBytes(0)) & "-" & _
        HexByteText(guidBytes(5)) & HexByteText(guidBytes(4)) & "-" & _
        HexByteText(guidBytes(7)) & HexByteText(guidBytes(6)) & "-" & _
        HexByteText(guidBytes(8)) & HexByteText(guidBytes(9)) & "-" & _
        HexByteText(guidBytes(10)) & HexByteText(guidBytes(11)) & _
        HexByteText(guidBytes(12)) & HexByteText(guidBytes(13)) & _
        HexByteText(guidBytes(14)) & HexByteText(guidBytes(15)))
End Function

Private Function HexByteText(ByVal value As Byte) As String
    HexByteText = Right$("0" & Hex$(value), 2)
End Function

' ---------------------------------------------------------------------------
' Zip core: a pure-VBA engine. Reading parses the central directory and
' inflates RFC 1951 streams (stored, fixed, and dynamic Huffman blocks)
' with CRC-32 verification; writing emits store-only archives that every
' unzipper reads. No COM component exists for this work: the in-box Shell
' folder trick has no completion or error channel, so the arithmetic lives
' here, validated offline against python-generated fixtures.
' ---------------------------------------------------------------------------

Private Sub RaiseZipError(ByVal message As String)
    Err.Raise ERROR_ZIP, "ROneCOne.ZipException", message
End Sub

Private Function ZipWord16(ByVal at As Long) As Long
    ZipWord16 = CLng(mZipWork(at)) + CLng(mZipWork(at + 1)) * 256&
End Function

Private Function ZipWord32(ByVal at As Long) As Double
    ZipWord32 = CDbl(mZipWork(at)) + CDbl(mZipWork(at + 1)) * 256# + _
        CDbl(mZipWork(at + 2)) * 65536# + CDbl(mZipWork(at + 3)) * 16777216#
End Function

Private Function ZipSizeValue(ByVal at As Long) As Long
    Dim raw As Double

    raw = ZipWord32(at)
    If raw = 4294967295# Then
        RaiseZipError "Zip64 archives (4 GB or larger) are not supported."
    End If
    If raw > 2147483646# Then
        RaiseZipError "The archive holds an entry too large to read."
    End If
    ZipSizeValue = CLng(raw)
End Function

Private Function ZipReadDirectoryCore() As Collection
    Dim at As Long
    Dim commentLength As Long
    Dim entries As Collection
    Dim entryCount As Long
    Dim entryIndex As Long
    Dim entryName As String
    Dim extraLength As Long
    Dim flags As Long
    Dim method As Long
    Dim nameLength As Long

    ZipEnsureTables
    at = ZipFindDirectoryEnd()
    If ZipWord16(at + 4) <> 0 Or ZipWord16(at + 6) <> 0 Then
        RaiseZipError "Multi-disk archives are not supported."
    End If
    entryCount = ZipWord16(at + 10)
    If entryCount = 65535 Then
        RaiseZipError "Zip64 archives (4 GB or larger) are not supported."
    End If
    Set entries = New Collection
    at = ZipSizeValue(at + 16)
    For entryIndex = 1 To entryCount
        If at + 45 > UBound(mZipWork) Then
            RaiseZipError "The central directory is truncated."
        End If
        If ZipWord32(at) <> 33639248# Then
            RaiseZipError "The central directory is not where the " & _
                "archive says."
        End If
        flags = ZipWord16(at + 8)
        If (flags And 1) <> 0 Then
            RaiseZipError "Encrypted archives are not supported."
        End If
        method = ZipWord16(at + 10)
        If method <> 0 And method <> 8 Then
            RaiseZipError "Only stored and deflated entries are supported."
        End If
        nameLength = ZipWord16(at + 28)
        extraLength = ZipWord16(at + 30)
        commentLength = ZipWord16(at + 32)
        entryName = ZipEntryNameText(at + 46, nameLength, flags)
        entries.Add Array(entryName, method, flags, ZipWord32(at + 16), _
            ZipSizeValue(at + 20), ZipSizeValue(at + 24), _
            ZipSizeValue(at + 42), Right$(entryName, 1) = "/")
        at = at + 46 + nameLength + extraLength + commentLength
    Next entryIndex
    Set ZipReadDirectoryCore = entries
End Function

Private Function ZipFindDirectoryEnd() As Long
    Dim at As Long
    Dim lowest As Long
    Dim size As Long

    size = ByteArrayLength(mZipWork)
    If size < 22 Then
        RaiseZipError "The file is too small to be a zip archive."
    End If
    lowest = size - 22 - 65535
    If lowest < 0 Then lowest = 0
    For at = size - 22 To lowest Step -1
        If mZipWork(at) = 80 Then
            If mZipWork(at + 1) = 75 And mZipWork(at + 2) = 5 And _
                mZipWork(at + 3) = 6 Then
                ZipFindDirectoryEnd = at
                Exit Function
            End If
        End If
    Next at
    RaiseZipError "The end-of-directory record was not found; this is " & _
        "not a zip archive."
End Function

Private Function ZipEntryNameText( _
    ByVal at As Long, _
    ByVal nameLength As Long, _
    ByVal flags As Long _
) As String
    Dim index As Long
    Dim nameBytes() As Byte

    If nameLength = 0 Then RaiseZipError "An entry has an empty name."
    ReDim nameBytes(0 To nameLength - 1)
    For index = 0 To nameLength - 1
        nameBytes(index) = mZipWork(at + index)
    Next index
    If (flags And 2048) <> 0 Then
        ZipEntryNameText = IoDecodeText(nameBytes, "utf-8")
    Else
        ZipEntryNameText = IoDecodeText(nameBytes, "windows-1252")
    End If
End Function

Private Function ZipExtractEntryCore( _
    ByVal entryName As String, _
    ByVal method As Long, _
    ByVal crcValue As Double, _
    ByVal compressedSize As Long, _
    ByVal size As Long, _
    ByVal localOffset As Long, _
    ByVal isDirectory As Boolean _
) As Byte()
    Dim dataStart As Long
    Dim index As Long
    Dim result() As Byte

    ' Each entry is its own instance, so a reader that never parsed a
    ' directory still needs the inflate and CRC tables ensured here.
    ZipEnsureTables
    If isDirectory Or size = 0 Then
        ZipExtractEntryCore = ZipEmptyBytes()
        Exit Function
    End If
    dataStart = ZipEntryDataStart(localOffset)
    If CDbl(dataStart) + compressedSize - 1 > UBound(mZipWork) Then
        RaiseZipError "The entry data runs past the end of the archive."
    End If
    If method = 0 Then
        If compressedSize <> size Then
            RaiseZipError "A stored entry's sizes disagree."
        End If
        ReDim result(0 To size - 1)
        For index = 0 To size - 1
            result(index) = mZipWork(dataStart + index)
        Next index
    Else
        result = ZipInflate(dataStart, compressedSize, size)
    End If
    If ZipCrc32(result) <> crcValue Then
        RaiseZipError "The entry '" & entryName & _
            "' failed its CRC-32 check."
    End If
    ZipExtractEntryCore = result
End Function

Private Function ZipEntryDataStart(ByVal localOffset As Long) As Long
    If localOffset + 29 > UBound(mZipWork) Then
        RaiseZipError "A local header runs past the end of the archive."
    End If
    If ZipWord32(localOffset) <> 67324752# Then
        RaiseZipError "A local header is not where the directory says."
    End If
    ZipEntryDataStart = localOffset + 30 + ZipWord16(localOffset + 26) + _
        ZipWord16(localOffset + 28)
End Function

Private Function ZipEntryBytes(ByVal memberName As String) As Byte()
    mZipWork = IoReadAllBytesCore(mZipArchivePath, memberName)
    ZipEntryBytes = ZipExtractEntryCore(mName, mZipEntryMethod, _
        mZipEntryCrc, mZipEntryCompressedSize, mZipEntrySize, _
        mZipEntryOffset, mZipEntryIsDirectory)
    Erase mZipWork
End Function

Private Function ZipEntryText( _
    ByRef dataBytes() As Byte, _
    ByVal requestedCharset As String _
) As String
    Dim byteCount As Long
    Dim charset As String
    Dim index As Long
    Dim skipCount As Long
    Dim sliced() As Byte

    byteCount = ByteArrayLength(dataBytes)
    If byteCount = 0 Then Exit Function
    charset = requestedCharset
    If byteCount >= 3 Then
        If dataBytes(0) = 239 And dataBytes(1) = 187 And _
            dataBytes(2) = 191 Then
            charset = "utf-8"
            skipCount = 3
        End If
    End If
    If skipCount = 0 And byteCount >= 2 Then
        If dataBytes(0) = 255 And dataBytes(1) = 254 Then
            charset = "unicode"
            skipCount = 2
        ElseIf dataBytes(0) = 254 And dataBytes(1) = 255 Then
            charset = "bigendianunicode"
            skipCount = 2
        End If
    End If
    If skipCount >= byteCount Then Exit Function
    If skipCount = 0 Then
        ZipEntryText = IoDecodeText(dataBytes, charset)
    Else
        ReDim sliced(0 To byteCount - skipCount - 1)
        For index = 0 To byteCount - skipCount - 1
            sliced(index) = dataBytes(skipCount + index)
        Next index
        ZipEntryText = IoDecodeText(sliced, charset)
    End If
End Function

Private Function SafeRelativeEntryPath(ByVal entryName As String) As String
    Dim part As Variant
    Dim parts As Variant
    Dim result As String

    If InStr(1, entryName, ":") > 0 Or InStr(1, entryName, "\") > 0 Or _
        Left$(entryName, 1) = "/" Then
        RaiseZipError "The entry '" & entryName & _
            "' would escape the target folder."
    End If
    parts = VBA.Split(entryName, "/")
    For Each part In parts
        If part = ".." Or part = "." Then
            RaiseZipError "The entry '" & entryName & _
                "' would escape the target folder."
        End If
        If Len(part) > 0 Then
            If Len(result) > 0 Then result = result & "\"
            result = result & part
        End If
    Next part
    If Len(result) = 0 Then RaiseZipError "An entry has an empty name."
    SafeRelativeEntryPath = result
End Function

Private Function ZipInflate( _
    ByVal startAt As Long, _
    ByVal compressedLength As Long, _
    ByVal outputSize As Long _
) As Byte()
    Dim blockType As Long
    Dim finalBlock As Long
    Dim state As InflateState

    state.position = startAt
    state.limit = startAt + compressedLength
    ReDim state.output(0 To outputSize - 1)
    Do
        finalBlock = InflateBits(state, 1)
        blockType = InflateBits(state, 2)
        Select Case blockType
            Case 0
                InflateStored state
            Case 1
                InflateCodes state, mFixedLitCounts, mFixedLitSymbols, _
                    mFixedDistCounts, mFixedDistSymbols
            Case 2
                InflateDynamicBlock state
            Case Else
                RaiseZipError "The deflate stream holds an invalid " & _
                    "block type."
        End Select
    Loop While finalBlock = 0
    If state.outPosition <> outputSize Then
        RaiseZipError "The deflate stream ended before the entry was " & _
            "complete."
    End If
    ZipInflate = state.output
End Function

Private Function InflateBits( _
    ByRef state As InflateState, _
    ByVal needed As Long _
) As Long
    Do While state.bitCount < needed
        If state.position >= state.limit Then
            RaiseZipError "The deflate stream ended early."
        End If
        state.bitBuffer = state.bitBuffer + _
            CLng(mZipWork(state.position)) * mPow2(state.bitCount)
        state.position = state.position + 1
        state.bitCount = state.bitCount + 8
    Loop
    InflateBits = state.bitBuffer And (mPow2(needed) - 1)
    state.bitBuffer = state.bitBuffer \ mPow2(needed)
    state.bitCount = state.bitCount - needed
End Function

Private Sub InflateStored(ByRef state As InflateState)
    Dim index As Long
    Dim storedLength As Long

    state.bitBuffer = 0
    state.bitCount = 0
    If state.position + 3 >= state.limit Then
        RaiseZipError "The deflate stream ended early."
    End If
    storedLength = CLng(mZipWork(state.position)) + _
        CLng(mZipWork(state.position + 1)) * 256&
    If (storedLength Xor (CLng(mZipWork(state.position + 2)) + _
        CLng(mZipWork(state.position + 3)) * 256&)) <> 65535 Then
        RaiseZipError "A stored deflate block fails its length check."
    End If
    state.position = state.position + 4
    If state.position + storedLength > state.limit Then
        RaiseZipError "The deflate stream ended early."
    End If
    If storedLength > 0 Then
        If state.outPosition + storedLength - 1 > UBound(state.output) Then
            RaiseZipError "The deflate output exceeds the entry size."
        End If
        For index = 1 To storedLength
            state.output(state.outPosition) = mZipWork(state.position)
            state.outPosition = state.outPosition + 1
            state.position = state.position + 1
        Next index
    End If
End Sub

Private Sub InflateCodes( _
    ByRef state As InflateState, _
    ByRef litCounts() As Long, _
    ByRef litSymbols() As Long, _
    ByRef distCounts() As Long, _
    ByRef distSymbols() As Long _
)
    Dim copyFrom As Long
    Dim distance As Long
    Dim runLength As Long
    Dim symbol As Long

    Do
        symbol = InflateDecode(state, litCounts, litSymbols)
        If symbol < 256 Then
            If state.outPosition > UBound(state.output) Then
                RaiseZipError "The deflate output exceeds the entry size."
            End If
            state.output(state.outPosition) = symbol
            state.outPosition = state.outPosition + 1
        ElseIf symbol > 256 Then
            If symbol > 285 Then
                RaiseZipError "The deflate stream holds an invalid " & _
                    "length code."
            End If
            runLength = mZipLengthBase(symbol - 257) + _
                InflateBits(state, mZipLengthExtra(symbol - 257))
            symbol = InflateDecode(state, distCounts, distSymbols)
            If symbol > 29 Then
                RaiseZipError "The deflate stream holds an invalid " & _
                    "distance code."
            End If
            distance = mZipDistBase(symbol) + _
                InflateBits(state, mZipDistExtra(symbol))
            If distance > state.outPosition Then
                RaiseZipError "The deflate stream reaches before the start."
            End If
            If state.outPosition + runLength - 1 > UBound(state.output) Then
                RaiseZipError "The deflate output exceeds the entry size."
            End If
            copyFrom = state.outPosition - distance
            Do While runLength > 0
                state.output(state.outPosition) = state.output(copyFrom)
                state.outPosition = state.outPosition + 1
                copyFrom = copyFrom + 1
                runLength = runLength - 1
            Loop
        End If
    Loop While symbol <> 256
End Sub

Private Function InflateDecode( _
    ByRef state As InflateState, _
    ByRef lengthCounts() As Long, _
    ByRef symbols() As Long _
) As Long
    Dim code As Long
    Dim count As Long
    Dim first As Long
    Dim size As Long
    Dim tableIndex As Long

    For size = 1 To 15
        code = code + InflateBits(state, 1)
        count = lengthCounts(size)
        If code - first < count Then
            InflateDecode = symbols(tableIndex + code - first)
            Exit Function
        End If
        tableIndex = tableIndex + count
        first = (first + count) * 2
        code = code * 2
    Next size
    RaiseZipError "The deflate stream holds an invalid Huffman code."
End Function

Private Sub InflateConstruct( _
    ByRef lengths() As Long, _
    ByVal startAt As Long, _
    ByVal symbolCount As Long, _
    ByRef lengthCounts() As Long, _
    ByRef symbols() As Long _
)
    Dim index As Long
    Dim leftover As Long
    Dim offsets(0 To 15) As Long
    Dim size As Long

    ReDim lengthCounts(0 To 15)
    ReDim symbols(0 To symbolCount)
    For index = 0 To symbolCount - 1
        lengthCounts(lengths(startAt + index)) = _
            lengthCounts(lengths(startAt + index)) + 1
    Next index
    If lengthCounts(0) = symbolCount Then Exit Sub
    leftover = 1
    For size = 1 To 15
        leftover = leftover * 2 - lengthCounts(size)
        If leftover < 0 Then
            RaiseZipError "The deflate code lengths are oversubscribed."
        End If
    Next size
    offsets(1) = 0
    For size = 1 To 14
        offsets(size + 1) = offsets(size) + lengthCounts(size)
    Next size
    For index = 0 To symbolCount - 1
        If lengths(startAt + index) <> 0 Then
            symbols(offsets(lengths(startAt + index))) = index
            offsets(lengths(startAt + index)) = _
                offsets(lengths(startAt + index)) + 1
        End If
    Next index
End Sub

Private Sub InflateDynamicBlock(ByRef state As InflateState)
    Dim codeCounts() As Long
    Dim codeSymbols() As Long
    Dim distCounts() As Long
    Dim distSymbols() As Long
    Dim hclen As Long
    Dim hdist As Long
    Dim hlit As Long
    Dim index As Long
    Dim lengths(0 To 318) As Long
    Dim litCounts() As Long
    Dim litSymbols() As Long
    Dim orderValues As Variant
    Dim previous As Long
    Dim repeatCount As Long
    Dim symbol As Long

    hlit = InflateBits(state, 5) + 257
    hdist = InflateBits(state, 5) + 1
    hclen = InflateBits(state, 4) + 4
    If hlit > 286 Or hdist > 30 Then
        RaiseZipError "The deflate stream declares too many codes."
    End If
    orderValues = Array(16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, _
        13, 2, 14, 1, 15)
    For index = 0 To 18
        lengths(index) = 0
    Next index
    For index = 0 To hclen - 1
        lengths(CLng(orderValues(index))) = InflateBits(state, 3)
    Next index
    InflateConstruct lengths, 0, 19, codeCounts, codeSymbols
    index = 0
    Do While index < hlit + hdist
        symbol = InflateDecode(state, codeCounts, codeSymbols)
        If symbol < 16 Then
            lengths(index) = symbol
            index = index + 1
        Else
            Select Case symbol
                Case 16
                    If index = 0 Then
                        RaiseZipError "The deflate stream repeats nothing."
                    End If
                    previous = lengths(index - 1)
                    repeatCount = 3 + InflateBits(state, 2)
                Case 17
                    previous = 0
                    repeatCount = 3 + InflateBits(state, 3)
                Case Else
                    previous = 0
                    repeatCount = 11 + InflateBits(state, 7)
            End Select
            If index + repeatCount > hlit + hdist Then
                RaiseZipError "The deflate stream repeats too many lengths."
            End If
            Do While repeatCount > 0
                lengths(index) = previous
                index = index + 1
                repeatCount = repeatCount - 1
            Loop
        End If
    Loop
    If lengths(256) = 0 Then
        RaiseZipError "The deflate stream is missing its end-of-block code."
    End If
    InflateConstruct lengths, 0, hlit, litCounts, litSymbols
    InflateConstruct lengths, hlit, hdist, distCounts, distSymbols
    InflateCodes state, litCounts, litSymbols, distCounts, distSymbols
End Sub

Private Function ZipCrc32(ByRef dataBytes() As Byte) As Double
    Dim crc As LongLong
    Dim index As Long
    Dim tableIndex As Long

    ZipEnsureTables
    crc = 4294967295^
    If ByteArrayLength(dataBytes) > 0 Then
        For index = LBound(dataBytes) To UBound(dataBytes)
            tableIndex = CLng((crc Xor dataBytes(index)) And 255^)
            crc = ((crc \ 256^) And 16777215^) Xor mCrcTable(tableIndex)
        Next index
    End If
    ZipCrc32 = CDbl(crc Xor 4294967295^)
End Function

Private Sub ZipEnsureTables()
    Dim bit As Long
    Dim entry As LongLong
    Dim index As Long
    Dim lengths(0 To 287) As Long
    Dim values As Variant

    If mZipTablesReady Then Exit Sub
    ReDim mPow2(0 To 24)
    ReDim mCrcTable(0 To 255)
    ReDim mZipLengthBase(0 To 28)
    ReDim mZipLengthExtra(0 To 28)
    ReDim mZipDistBase(0 To 29)
    ReDim mZipDistExtra(0 To 29)
    mPow2(0) = 1
    For index = 1 To 24
        mPow2(index) = mPow2(index - 1) * 2
    Next index
    For index = 0 To 255
        entry = index
        For bit = 1 To 8
            If (entry And 1^) <> 0 Then
                entry = 3988292384^ Xor ((entry \ 2^) And 2147483647^)
            Else
                entry = (entry \ 2^) And 2147483647^
            End If
        Next bit
        mCrcTable(index) = entry
    Next index
    values = Array(3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, _
        31, 35, 43, 51, 59, 67, 83, 99, 115, 131, 163, 195, 227, 258)
    For index = 0 To 28
        mZipLengthBase(index) = values(index)
    Next index
    values = Array(0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, _
        3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0)
    For index = 0 To 28
        mZipLengthExtra(index) = values(index)
    Next index
    values = Array(1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, _
        193, 257, 385, 513, 769, 1025, 1537, 2049, 3073, 4097, 6145, _
        8193, 12289, 16385, 24577)
    For index = 0 To 29
        mZipDistBase(index) = values(index)
    Next index
    values = Array(0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, _
        8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13)
    For index = 0 To 29
        mZipDistExtra(index) = values(index)
    Next index
    For index = 0 To 143
        lengths(index) = 8
    Next index
    For index = 144 To 255
        lengths(index) = 9
    Next index
    For index = 256 To 279
        lengths(index) = 7
    Next index
    For index = 280 To 287
        lengths(index) = 8
    Next index
    InflateConstruct lengths, 0, 288, mFixedLitCounts, mFixedLitSymbols
    For index = 0 To 29
        lengths(index) = 5
    Next index
    InflateConstruct lengths, 0, 30, mFixedDistCounts, mFixedDistSymbols
    mZipTablesReady = True
End Sub

Private Sub ZipWriteArchiveFromDirectory( _
    ByVal sourceDirectory As String, _
    ByVal destinationArchive As String _
)
    Dim centralRecords As Collection
    Dim stream As Object

    ZipEnsureTables
    Set centralRecords = New Collection
    Set stream = CreateStreamObject(STREAM_TYPE_BINARY)
    ZipWriteFolderEntries stream, centralRecords, sourceDirectory, _
        vbNullString
    ZipWriteCentralDirectory stream, centralRecords
    On Error GoTo SaveFailed
    stream.SaveToFile destinationArchive, STREAM_SAVE_OVERWRITE
    On Error GoTo 0
    stream.Close
    Exit Sub

SaveFailed:
    RaiseIoError "The archive '" & destinationArchive & _
        "' could not be written: " & Err.Description
End Sub

Private Sub ZipWriteFolderEntries( _
    ByVal stream As Object, _
    ByVal centralRecords As Collection, _
    ByVal folderPath As String, _
    ByVal relativePrefix As String _
)
    Dim dataBytes() As Byte
    Dim fileItem As Object
    Dim folder As Object
    Dim subFolder As Object

    Set folder = CreateFileSystemObject().GetFolder(folderPath)
    For Each fileItem In folder.Files
        dataBytes = IoReadAllBytesCore(fileItem.Path, "CreateFromDirectory")
        ZipAppendEntry stream, centralRecords, _
            relativePrefix & fileItem.Name, dataBytes, _
            fileItem.DateLastModified
    Next fileItem
    For Each subFolder In folder.SubFolders
        If subFolder.Files.Count = 0 And subFolder.SubFolders.Count = 0 Then
            ZipAppendEntry stream, centralRecords, relativePrefix & _
                subFolder.Name & "/", ZipEmptyBytes(), _
                subFolder.DateLastModified
        Else
            ZipWriteFolderEntries stream, centralRecords, subFolder.Path, _
                relativePrefix & subFolder.Name & "/"
        End If
    Next subFolder
End Sub

Private Sub ZipAppendEntry( _
    ByVal stream As Object, _
    ByVal centralRecords As Collection, _
    ByVal entryName As String, _
    ByRef dataBytes() As Byte, _
    ByVal modifiedAt As Date _
)
    Dim crcValue As Double
    Dim dosDate As Long
    Dim dosTime As Long
    Dim flags As Long
    Dim nameBytes() As Byte
    Dim offset As Long
    Dim size As Long

    nameBytes = IoEncodeText(entryName, "utf-8", 3)
    flags = ZipNameFlags(entryName)
    size = ByteArrayLength(dataBytes)
    crcValue = ZipCrc32(dataBytes)
    dosDate = ZipDosDate(modifiedAt)
    dosTime = ZipDosTime(modifiedAt)
    offset = CLng(stream.Position)
    ZipWriteWord32 stream, 67324752#
    ZipWriteWord16 stream, 20
    ZipWriteWord16 stream, flags
    ZipWriteWord16 stream, 0
    ZipWriteWord16 stream, dosTime
    ZipWriteWord16 stream, dosDate
    ZipWriteWord32 stream, crcValue
    ZipWriteWord32 stream, CDbl(size)
    ZipWriteWord32 stream, CDbl(size)
    ZipWriteWord16 stream, ByteArrayLength(nameBytes)
    ZipWriteWord16 stream, 0
    stream.Write nameBytes
    If size > 0 Then stream.Write dataBytes
    centralRecords.Add Array(nameBytes, flags, crcValue, size, offset, _
        dosTime, dosDate)
End Sub

Private Sub ZipWriteCentralDirectory( _
    ByVal stream As Object, _
    ByVal centralRecords As Collection _
)
    Dim directorySize As Long
    Dim directoryStart As Long
    Dim nameBytes() As Byte
    Dim record As Variant

    directoryStart = CLng(stream.Position)
    For Each record In centralRecords
        nameBytes = record(0)
        ZipWriteWord32 stream, 33639248#
        ZipWriteWord16 stream, 20
        ZipWriteWord16 stream, 20
        ZipWriteWord16 stream, CLng(record(1))
        ZipWriteWord16 stream, 0
        ZipWriteWord16 stream, CLng(record(5))
        ZipWriteWord16 stream, CLng(record(6))
        ZipWriteWord32 stream, CDbl(record(2))
        ZipWriteWord32 stream, CDbl(record(3))
        ZipWriteWord32 stream, CDbl(record(3))
        ZipWriteWord16 stream, ByteArrayLength(nameBytes)
        ZipWriteWord16 stream, 0
        ZipWriteWord16 stream, 0
        ZipWriteWord16 stream, 0
        ZipWriteWord16 stream, 0
        ZipWriteWord32 stream, 0
        ZipWriteWord32 stream, CDbl(record(4))
        stream.Write nameBytes
    Next record
    directorySize = CLng(stream.Position) - directoryStart
    ZipWriteWord32 stream, 101010256#
    ZipWriteWord16 stream, 0
    ZipWriteWord16 stream, 0
    ZipWriteWord16 stream, centralRecords.Count
    ZipWriteWord16 stream, centralRecords.Count
    ZipWriteWord32 stream, CDbl(directorySize)
    ZipWriteWord32 stream, CDbl(directoryStart)
    ZipWriteWord16 stream, 0
End Sub

Private Sub ZipWriteWord16(ByVal stream As Object, ByVal value As Long)
    Dim pair(0 To 1) As Byte

    pair(0) = value And 255&
    pair(1) = (value \ 256&) And 255&
    stream.Write pair
End Sub

Private Sub ZipWriteWord32(ByVal stream As Object, ByVal value As Double)
    Dim index As Long
    Dim quad(0 To 3) As Byte
    Dim work As Double

    work = value
    For index = 0 To 3
        quad(index) = CByte(work - Int(work / 256#) * 256#)
        work = Int(work / 256#)
    Next index
    stream.Write quad
End Sub

Private Function ZipNameFlags(ByVal entryName As String) As Long
    Dim index As Long

    For index = 1 To Len(entryName)
        If AscW(Mid$(entryName, index, 1)) > 127 Or _
            AscW(Mid$(entryName, index, 1)) < 0 Then
            ZipNameFlags = 2048
            Exit Function
        End If
    Next index
End Function

Private Function ZipDosDate(ByVal value As Date) As Long
    Dim yearPart As Long

    yearPart = VBA.Year(value)
    If yearPart < 1980 Then yearPart = 1980
    ZipDosDate = (yearPart - 1980) * 512 + VBA.Month(value) * 32 + _
        VBA.Day(value)
End Function

Private Function ZipDosTime(ByVal value As Date) As Long
    ZipDosTime = VBA.Hour(value) * 2048 + VBA.Minute(value) * 32 + _
        VBA.Second(value) \ 2
End Function

Private Function ZipEmptyBytes() As Byte()
    Dim result() As Byte

    result = vbNullString
    ZipEmptyBytes = result
End Function

Private Function ByteArrayLength(ByRef dataBytes() As Byte) As Long
    On Error GoTo Unallocated
    ByteArrayLength = UBound(dataBytes) - LBound(dataBytes) + 1
    Exit Function
Unallocated:
End Function

' ---------------------------------------------------------------------------
' File watcher core: a watch task holds a name-size-time snapshot of one
' folder and rescans it, throttled, while awaited. The first difference
' found resolves the task as a Created, Changed, or Deleted value.
' ---------------------------------------------------------------------------

Private Function InternalWatchPoll() As Boolean
    Dim baselineEntry As Variant
    Dim currentEntry As Variant
    Dim currentSnapshot As Collection
    Dim entryName As String
    Dim knownValue As String
    Dim nowTick As LongLong

    nowTick = GetTickCount64()
    If mWatcherLastTick <> 0 Then
        If nowTick - mWatcherLastTick < WATCH_POLL_INTERVAL_MS Then
            Exit Function
        End If
    End If
    mWatcherLastTick = nowTick
    Set currentSnapshot = WatcherSnapshot(mSource.InternalWatcherFolder, _
        mSource.InternalWatcherFilter)
    For Each currentEntry In currentSnapshot
        entryName = WatcherEntryName(CStr(currentEntry))
        If Not WatcherTryValue(mWatcherSnapshot, entryName, knownValue) Then
            Set mTaskResult = CreateObjectResult( _
                BuildFileChange("Created", entryName))
            InternalWatchPoll = True
            Exit Function
        ElseIf knownValue <> CStr(currentEntry) Then
            Set mTaskResult = CreateObjectResult( _
                BuildFileChange("Changed", entryName))
            InternalWatchPoll = True
            Exit Function
        End If
    Next currentEntry
    For Each baselineEntry In mWatcherSnapshot
        entryName = WatcherEntryName(CStr(baselineEntry))
        If Not WatcherTryValue(currentSnapshot, entryName, knownValue) Then
            Set mTaskResult = CreateObjectResult( _
                BuildFileChange("Deleted", entryName))
            InternalWatchPoll = True
            Exit Function
        End If
    Next baselineEntry
End Function

Private Function WatcherSnapshot( _
    ByVal folderPath As String, _
    ByVal filterPattern As String _
) As Collection
    Dim entry As Object
    Dim patternLower As String

    ' The same escape-and-Like matching the Directory surface uses, so a
    ' watcher filter behaves exactly like a GetFiles pattern.
    patternLower = LCase$(VBA.Replace$(filterPattern, "[", "[[]"))
    Set WatcherSnapshot = New Collection
    For Each entry In CreateFileSystemObject().GetFolder(folderPath).Files
        If LCase$(entry.Name) Like patternLower Then
            WatcherSnapshot.Add entry.Name & "|" & CStr(entry.Size) & _
                "|" & CStr(CDbl(entry.DateLastModified)), _
                "k" & LCase$(entry.Name)
        End If
    Next entry
End Function

Private Function WatcherEntryName(ByVal entryText As String) As String
    WatcherEntryName = Left$(entryText, InStr(1, entryText, "|") - 1)
End Function

Private Function WatcherTryValue( _
    ByVal snapshot As Collection, _
    ByVal entryName As String, _
    ByRef valueText As String _
) As Boolean
    On Error GoTo Missing
    valueText = snapshot.Item("k" & LCase$(entryName))
    WatcherTryValue = True
    Exit Function
Missing:
End Function

Private Function BuildFileChange( _
    ByVal changeType As String, _
    ByVal fileName As String _
) As ROneCOne
    Dim instance As ROneCOne

    Set instance = New ROneCOne
    instance.ConfigureFileChange changeType, fileName
    Set BuildFileChange = instance
End Function

' ---------------------------------------------------------------------------
' Logger core: one UTC millisecond timestamp, a three-letter severity code,
' and a composite-formatted message per line. Every call appends through the
' file layer and closes, so written lines survive a crash.
' ---------------------------------------------------------------------------

Private Sub LoggerWrite( _
    ByVal level As Long, _
    ByVal template As String, _
    ByRef values() As Variant _
)
    Dim line As String

    If level < mLoggerMinimumLevel Then Exit Sub
    line = LoggerTimestampText() & " [" & LoggerLevelCode(level) & "] " & _
        FormatCompositeCore(template, values) & vbCrLf
    IoWriteTextFile mName, line, "utf-8", True, "Logger"
End Sub

Private Function LoggerTimestampText() As String
    Dim clockDate As Date
    Dim msOfDay As Long
    Dim systemNow As SYSTEMTIME

    GetSystemTime systemNow
    MsToClockParts SystemTimeToMs(systemNow), clockDate, msOfDay
    LoggerTimestampText = ClockStampText(clockDate, msOfDay) & "." & _
        PadNumber(msOfDay Mod 1000, 3) & "Z"
End Function

Private Function LoggerLevelCode(ByVal level As Long) As String
    Select Case level
        Case LOG_LEVEL_TRACE
            LoggerLevelCode = "TRC"
        Case LOG_LEVEL_DEBUG
            LoggerLevelCode = "DBG"
        Case LOG_LEVEL_INFORMATION
            LoggerLevelCode = "INF"
        Case LOG_LEVEL_WARNING
            LoggerLevelCode = "WRN"
        Case LOG_LEVEL_ERROR
            LoggerLevelCode = "ERR"
        Case Else
            LoggerLevelCode = "CRT"
    End Select
End Function

Private Function LoggerLevelValue( _
    ByVal levelName As String, _
    ByVal memberName As String _
) As Long
    Select Case LCase$(Trim$(levelName))
        Case "trace"
            LoggerLevelValue = LOG_LEVEL_TRACE
        Case "debug"
            LoggerLevelValue = LOG_LEVEL_DEBUG
        Case "information"
            LoggerLevelValue = LOG_LEVEL_INFORMATION
        Case "warning"
            LoggerLevelValue = LOG_LEVEL_WARNING
        Case "error"
            LoggerLevelValue = LOG_LEVEL_ERROR
        Case "critical"
            LoggerLevelValue = LOG_LEVEL_CRITICAL
        Case Else
            RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
                "Levels are Trace, Debug, Information, Warning, Error, " & _
                "and Critical."
    End Select
End Function

' ---------------------------------------------------------------------------
' Escaping core: RFC 3986 percent encoding over UTF-8 bytes, and the HTML
' entity dialect (the five markup entities plus numeric references).
' ---------------------------------------------------------------------------

Private Function UriIsUnreservedByte(ByVal value As Byte) As Boolean
    Select Case value
        Case 48 To 57, 65 To 90, 97 To 122, 45, 46, 95, 126
            UriIsUnreservedByte = True
    End Select
End Function

Private Function IsHexPairText(ByVal text As String) As Boolean
    If Len(text) <> 2 Then Exit Function
    IsHexPairText = IsHexDigitChar(Mid$(text, 1, 1)) And _
        IsHexDigitChar(Mid$(text, 2, 1))
End Function

Private Function IsHexDigitChar(ByVal character As String) As Boolean
    Select Case character
        Case "0" To "9", "A" To "F", "a" To "f"
            IsHexDigitChar = True
    End Select
End Function

Private Sub FlushPendingUtf8( _
    ByRef buffer As JsonTextBuilder, _
    ByRef pendingBytes() As Byte, _
    ByRef pendingCount As Long _
)
    Dim chunk() As Byte
    Dim index As Long

    If pendingCount = 0 Then Exit Sub
    ReDim chunk(0 To pendingCount - 1)
    For index = 0 To pendingCount - 1
        chunk(index) = pendingBytes(index)
    Next index
    JsonBuilderAppend buffer, IoDecodeText(chunk, "utf-8")
    pendingCount = 0
End Sub

Private Function TryAppendHtmlEntity( _
    ByRef buffer As JsonTextBuilder, _
    ByVal entity As String _
) As Boolean
    Dim code As Long
    Dim digits As String
    Dim index As Long

    Select Case entity
        Case "amp"
            JsonBuilderAppend buffer, "&"
        Case "lt"
            JsonBuilderAppend buffer, "<"
        Case "gt"
            JsonBuilderAppend buffer, ">"
        Case "quot"
            JsonBuilderAppend buffer, """"
        Case "apos"
            JsonBuilderAppend buffer, "'"
        Case Else
            If Left$(entity, 1) <> "#" Then Exit Function
            digits = Mid$(entity, 2)
            If Left$(digits, 1) = "x" Or Left$(digits, 1) = "X" Then
                digits = Mid$(digits, 2)
                If Len(digits) = 0 Or Len(digits) > 6 Then Exit Function
                For index = 1 To Len(digits)
                    If Not IsHexDigitChar(Mid$(digits, index, 1)) Then
                        Exit Function
                    End If
                Next index
                code = CLng(Val("&H" & digits))
            Else
                If Len(digits) = 0 Or Len(digits) > 7 Then Exit Function
                For index = 1 To Len(digits)
                    If Mid$(digits, index, 1) < "0" Or _
                        Mid$(digits, index, 1) > "9" Then Exit Function
                Next index
                code = CLng(digits)
            End If
            If code < 0 Or code > &H10FFFF Then Exit Function
            If code >= &HD800& And code <= &HDFFF& Then Exit Function
            JsonBuilderAppend buffer, CodePointText(code)
    End Select
    TryAppendHtmlEntity = True
End Function

Private Function CodePointText(ByVal code As Long) As String
    If code <= &HFFFF& Then
        CodePointText = ChrW$(code)
    Else
        CodePointText = ChrW$(&HD800& + (code - &H10000) \ &H400&) & _
            ChrW$(&HDC00& + (code - &H10000) Mod &H400&)
    End If
End Function

' ---------------------------------------------------------------------------
' Xml core: a secured MSXML6 document behind XPath queries, plus the bridges
' between element-per-column XML and typed DataTables. Column typing reuses
' the CSV inference so both exchange formats agree on every value shape.
' ---------------------------------------------------------------------------

Private Function CreateSecuredXmlDocument( _
    ByVal selectionNamespaces As String _
) As Object
    Dim dom As Object

    Set dom = CreateObject(XML_PROG_ID)
    dom.async = False
    dom.validateOnParse = False
    dom.resolveExternals = False
    If Len(selectionNamespaces) > 0 Then
        On Error GoTo BadNamespaces
        dom.setProperty "SelectionNamespaces", selectionNamespaces
        On Error GoTo 0
    End If
    Set CreateSecuredXmlDocument = dom
    Exit Function
BadNamespaces:
    RaiseXmlError "The selection namespaces are not valid: " & _
        selectionNamespaces
End Function

Private Function XmlParseCore( _
    ByVal text As String, _
    ByVal selectionNamespaces As String _
) As ROneCOne
    Dim dom As Object

    Set dom = CreateSecuredXmlDocument(selectionNamespaces)
    If Not dom.loadXML(text) Then RaiseXmlParseError dom
    If dom.documentElement Is Nothing Then
        RaiseXmlError "The XML has no document element."
    End If
    Set XmlParseCore = BuildXmlNode(dom.documentElement)
End Function

Private Function BuildXmlNode(ByVal nodeObject As Object) As ROneCOne
    Dim instance As ROneCOne

    Set instance = New ROneCOne
    instance.ConfigureXmlNode nodeObject
    Set BuildXmlNode = instance
End Function

Private Function XmlFindAttribute(ByVal attributeName As String) As Object
    If mProviderObject.Attributes Is Nothing Then Exit Function
    Set XmlFindAttribute = _
        mProviderObject.Attributes.getNamedItem(attributeName)
End Function

Private Function XmlRunXPath( _
    ByVal xpath As String, _
    ByVal memberName As String _
) As Object
    On Error GoTo Failed
    Set XmlRunXPath = mProviderObject.SelectNodes(xpath)
    Exit Function
Failed:
    RaiseXmlError memberName & " could not run the XPath query '" & _
        xpath & "': " & Err.Description
End Function

Private Sub RaiseXmlParseError(ByVal dom As Object)
    Dim reason As String

    reason = Trim$(VBA.Replace$(CStr(dom.parseError.reason), vbCrLf, " "))
    RaiseXmlError "The XML failed to parse at line " & _
        CStr(dom.parseError.Line) & ", position " & _
        CStr(dom.parseError.linepos) & ": " & reason
End Sub

Private Sub RaiseXmlError(ByVal message As String)
    Err.Raise ERROR_XML, "ROneCOne.XmlException", message
End Sub

Private Function XmlDeserializeTable( _
    ByVal text As String, _
    ByVal tableName As String, _
    ByVal rowsPath As String _
) As ROneCOne
    Dim cellText As String
    Dim columnIndex As Long
    Dim columnNames As Collection
    Dim columnTypes() As Long
    Dim dateValue As Date
    Dim dom As Object
    Dim nameKeys As Collection
    Dim rowNode As Object
    Dim rowNodes As Object
    Dim rowValues() As Variant
    Dim table As ROneCOne

    Set dom = CreateSecuredXmlDocument(vbNullString)
    If Not dom.loadXML(text) Then RaiseXmlParseError dom
    On Error Resume Next
    Set rowNodes = dom.SelectNodes(rowsPath)
    If Err.Number <> 0 Then
        On Error GoTo 0
        RaiseXmlError "The row path '" & rowsPath & "' is not a valid XPath."
    End If
    On Error GoTo 0
    If rowNodes.Length = 0 Then
        RaiseXmlError "The row path '" & rowsPath & _
            "' matched no row elements."
    End If

    ' Columns are the union, in first-seen order, of every row's attributes
    ' and simple child elements; nested elements never become columns.
    Set columnNames = New Collection
    Set nameKeys = New Collection
    For Each rowNode In rowNodes
        CollectXmlColumns rowNode, columnNames, nameKeys
    Next rowNode
    If columnNames.Count = 0 Then
        RaiseXmlError "The row elements have no attributes or simple " & _
            "child elements to become columns."
    End If

    ReDim columnTypes(1 To columnNames.Count)
    For Each rowNode In rowNodes
        For columnIndex = 1 To columnNames.Count
            If TryXmlCellText(rowNode, _
                CStr(columnNames.Item(columnIndex)), cellText) Then
                If Len(cellText) > 0 Then
                    MergeCsvColumnType columnTypes(columnIndex), _
                        CsvClassifyCell(cellText, False)
                End If
            End If
        Next columnIndex
    Next rowNode

    Set table = ROneCOne.DataTable(tableName)
    For columnIndex = 1 To columnNames.Count
        If columnTypes(columnIndex) = 0 Then
            columnTypes(columnIndex) = vbString
        End If
        table.Column CStr(columnNames.Item(columnIndex)), _
            columnTypes(columnIndex)
    Next columnIndex
    For Each rowNode In rowNodes
        ReDim rowValues(0 To columnNames.Count - 1)
        For columnIndex = 1 To columnNames.Count
            If Not TryXmlCellText(rowNode, _
                CStr(columnNames.Item(columnIndex)), cellText) Then
                rowValues(columnIndex - 1) = Null
            ElseIf Len(cellText) = 0 Then
                rowValues(columnIndex - 1) = Null
            ElseIf columnTypes(columnIndex) = vbString Then
                rowValues(columnIndex - 1) = cellText
            ElseIf columnTypes(columnIndex) = vbDate Then
                If CsvTryIsoDate(cellText, dateValue) Then
                    rowValues(columnIndex - 1) = dateValue
                End If
            Else
                rowValues(columnIndex - 1) = CsvConvertCell(cellText, _
                    columnTypes(columnIndex))
            End If
        Next columnIndex
        table.LoadRow rowValues
    Next rowNode
    Set XmlDeserializeTable = table
End Function

Private Sub CollectXmlColumns( _
    ByVal rowNode As Object, _
    ByVal columnNames As Collection, _
    ByVal nameKeys As Collection _
)
    Dim attributeNode As Object
    Dim child As Object

    If Not rowNode.Attributes Is Nothing Then
        For Each attributeNode In rowNode.Attributes
            AddXmlColumnName CStr(attributeNode.nodeName), _
                columnNames, nameKeys
        Next attributeNode
    End If
    For Each child In rowNode.childNodes
        If child.nodeType = 1 Then
            If Not XmlHasElementChildren(child) Then
                AddXmlColumnName CStr(child.nodeName), columnNames, nameKeys
            End If
        End If
    Next child
End Sub

Private Sub AddXmlColumnName( _
    ByVal columnName As String, _
    ByVal columnNames As Collection, _
    ByVal nameKeys As Collection _
)
    On Error Resume Next
    nameKeys.Add columnNames.Count + 1, "k" & UCase$(columnName)
    If Err.Number = 0 Then columnNames.Add columnName
    On Error GoTo 0
End Sub

Private Function XmlHasElementChildren(ByVal node As Object) As Boolean
    Dim child As Object

    For Each child In node.childNodes
        If child.nodeType = 1 Then
            XmlHasElementChildren = True
            Exit Function
        End If
    Next child
End Function

Private Function TryXmlCellText( _
    ByVal rowNode As Object, _
    ByVal columnName As String, _
    ByRef cellText As String _
) As Boolean
    Dim attributeNode As Object
    Dim child As Object

    ' A simple child element wins over an attribute of the same name.
    For Each child In rowNode.childNodes
        If child.nodeType = 1 Then
            If child.nodeName = columnName And _
                Not XmlHasElementChildren(child) Then
                cellText = child.Text
                TryXmlCellText = True
                Exit Function
            End If
        End If
    Next child
    If Not rowNode.Attributes Is Nothing Then
        Set attributeNode = rowNode.Attributes.getNamedItem(columnName)
        If Not attributeNode Is Nothing Then
            cellText = attributeNode.Text
            TryXmlCellText = True
            Exit Function
        End If
    End If
End Function

Private Function XmlSerializeTable( _
    ByVal table As ROneCOne, _
    ByVal rootName As String, _
    ByVal rowName As String _
) As String
    Dim builder As JsonTextBuilder
    Dim cells As Collection
    Dim columnIndex As Long
    Dim columnNames() As String
    Dim columns As Collection
    Dim row As ROneCOne
    Dim wrapped As ROneCOne

    If Len(rowName) = 0 Then rowName = table.TableName
    If Len(rowName) = 0 Then rowName = "Row"
    RequireXmlElementName rootName
    RequireXmlElementName rowName
    Set columns = table.InternalDataColumns
    If columns.Count = 0 Then
        RaiseXmlError "XML serialization requires at least one column."
    End If
    ReDim columnNames(1 To columns.Count)
    For columnIndex = 1 To columns.Count
        columnNames(columnIndex) = columns.Item(columnIndex).ColumnName
        RequireXmlElementName columnNames(columnIndex)
    Next columnIndex
    JsonBuilderInit builder, 512
    JsonBuilderAppend builder, "<" & rootName & ">"
    For Each row In table.InternalDataRows
        If row.RowState <> "Deleted" Then
            JsonBuilderAppend builder, "<" & rowName & ">"
            Set cells = row.InternalDataRows
            For columnIndex = 1 To cells.Count
                Set wrapped = cells.Item(columnIndex)
                ' A database null omits its element, like DataTable.WriteXml.
                If Not IsNullDataValue(wrapped) Then
                    JsonBuilderAppend builder, _
                        "<" & columnNames(columnIndex) & ">"
                    JsonBuilderAppend builder, _
                        XmlCellText(JsonUnwrapValue(wrapped))
                    JsonBuilderAppend builder, _
                        "</" & columnNames(columnIndex) & ">"
                End If
            Next columnIndex
            JsonBuilderAppend builder, "</" & rowName & ">"
        End If
    Next row
    JsonBuilderAppend builder, "</" & rootName & ">"
    XmlSerializeTable = JsonBuilderText(builder)
End Function

Private Function XmlCellText(ByVal value As Variant) As String
    Dim clockDate As Date
    Dim msOfDay As Long

    Select Case VarTypeOf(value)
        Case vbString
            XmlCellText = XmlEscapeText(CStr(value))
        Case vbBoolean
            If value Then
                XmlCellText = "true"
            Else
                XmlCellText = "false"
            End If
        Case vbByte, vbInteger, vbLong, vbLongLong
            XmlCellText = CStr(value)
        Case vbSingle, vbDouble, vbCurrency, vbDecimal
            XmlCellText = JsonNumberText(CDbl(value))
        Case vbDate
            MsToClockParts DateToMs(CDate(value)), clockDate, msOfDay
            XmlCellText = ClockStampText(clockDate, msOfDay)
        Case Else
            RaiseXmlError "XML cells must be scalar values; '" & _
                TypeName(value) & "' is not writable."
    End Select
End Function

Private Function XmlEscapeText(ByVal text As String) As String
    XmlEscapeText = VBA.Replace$(text, "&", "&amp;")
    XmlEscapeText = VBA.Replace$(XmlEscapeText, "<", "&lt;")
    XmlEscapeText = VBA.Replace$(XmlEscapeText, ">", "&gt;")
End Function

Private Sub RequireXmlElementName(ByVal elementName As String)
    Dim index As Long
    Dim probe As String

    If Len(elementName) = 0 Then
        RaiseXmlError "An XML element name cannot be empty."
    End If
    For index = 1 To Len(elementName)
        probe = Mid$(elementName, index, 1)
        If Not ((probe >= "A" And probe <= "Z") Or _
            (probe >= "a" And probe <= "z") Or probe = "_" Or _
            (index > 1 And ((probe >= "0" And probe <= "9") Or _
            probe = "-" Or probe = "."))) Then
            RaiseXmlError "The name '" & elementName & _
                "' cannot be an XML element name."
        End If
    Next index
End Sub

Private Function IsDateTimeInstance(ByVal value As Variant) As Boolean
    Dim runtimeValue As ROneCOne

    If Not IsObject(value) Then Exit Function
    If Not TypeOf value Is ROneCOne Then Exit Function
    Set runtimeValue = value
    If runtimeValue Is Nothing Then Exit Function
    IsDateTimeInstance = (runtimeValue.InternalRole = ROLE_DATETIME)
End Function

Private Function RequireDateTimeOperand( _
    ByVal value As Variant, _
    ByVal memberName As String _
) As Double
    Dim runtimeValue As ROneCOne

    If Not IsDateTimeInstance(value) Then
        RaiseContractError ERROR_TYPE_MISMATCH, memberName, _
            memberName & " requires a DateTime value."
    End If
    Set runtimeValue = value
    RequireDateTimeOperand = runtimeValue.InternalDateTimeMs
End Function

Private Function RequireTimeSpanOperandMs( _
    ByVal value As Variant, _
    ByVal memberName As String _
) As Double
    Dim runtimeValue As ROneCOne

    If IsObject(value) Then
        If TypeOf value Is ROneCOne Then
            Set runtimeValue = value
            If Not runtimeValue Is Nothing Then
                If runtimeValue.InternalRole = ROLE_TIMESPAN Then
                    RequireTimeSpanOperandMs = runtimeValue.InternalTimeSpanMs
                    Exit Function
                End If
            End If
        End If
    End If
    RaiseContractError ERROR_TYPE_MISMATCH, memberName, _
        memberName & " requires a TimeSpan value."
End Function

Private Function CoerceOffsetMinutes( _
    ByVal offsetValue As Variant, _
    ByVal memberName As String _
) As Long
    Dim minutesValue As Double

    If IsObject(offsetValue) Then
        minutesValue = RequireTimeSpanOperandMs(offsetValue, memberName) / _
            MS_PER_MINUTE
    ElseIf IsNumeric(offsetValue) And VarTypeOf(offsetValue) <> vbString Then
        minutesValue = CDbl(offsetValue)
    Else
        RaiseContractError ERROR_TYPE_MISMATCH, memberName, _
            memberName & " requires a TimeSpan or whole minutes."
    End If
    If minutesValue <> Fix(minutesValue) Or Abs(minutesValue) > 840 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
            "The offset must be whole minutes within 14 hours of UTC."
    End If
    CoerceOffsetMinutes = CLng(minutesValue)
End Function

Private Function RequireIntegralNumber( _
    ByVal value As Variant, _
    ByVal memberName As String _
) As Double
    If IsObject(value) Or VarTypeOf(value) = vbString Or Not IsNumeric(value) Then
        RaiseContractError ERROR_TYPE_MISMATCH, memberName, _
            memberName & " requires a whole number."
    End If
    RequireIntegralNumber = CDbl(value)
    If RequireIntegralNumber <> Fix(RequireIntegralNumber) Then
        RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
            memberName & " requires a whole number."
    End If
End Function

Private Sub RaiseContractError( _
    ByVal errorNumber As Long, _
    ByVal memberName As String, _
    ByVal message As String _
)
    Err.Raise errorNumber, "ROneCOne." & memberName, message
End Sub

Private Sub Class_Terminate()
    Set mLeft = Nothing
    Set mRight = Nothing
    Set mNext = Nothing
    Set mObjectValue = Nothing
    Set mParameters = Nothing
    Set mWatcherSnapshot = Nothing
    Set mZipEntries = Nothing
    Erase mZipWork
    Set mSqlQueryConnection = Nothing
    Set mSqlQueryWheres = Nothing
    Set mSqlQueryOrderings = Nothing
    Set mSqlQueryColumns = Nothing
    Set mSqlQueryLastValues = Nothing
    ' A session releases its handles without terminating the command; a
    ' caller who wants the process gone calls Kill.
    If mRole = ROLE_PROCESS_SESSION Then SessionCloseHandles
    Set mSessionWriteQueue = Nothing
    Erase mSessionOutBytes
    Erase mSessionErrBytes
    Erase mSessionWriteBuffer
    Set mTarget = Nothing
    Set mInvocationList = Nothing
    Set mTryBody = Nothing
    Set mCatchActions = Nothing
    Set mCatchNumbers = Nothing
    Set mFinallyAction = Nothing
    Erase mItems
    Erase mKeys
    Erase mPriorities
    Erase mOriginalItems
    Set mEnumerationValues = Nothing
    Set mElementParameter = Nothing
    Set mSource = Nothing
    Set mQueryDelegate = Nothing
    Set mQueryComparer = Nothing
    Set mQueryValue = Nothing
End Sub
