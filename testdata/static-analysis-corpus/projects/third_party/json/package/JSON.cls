VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JSON"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/**
' * JSON - JSON Reader and Writer
' * @description A high-performance zero-copy JSON reader and lightweight JSON writer for VBA (x86/x64 compatible).
' * Uses compact token trees, crash-safe character buffers, native InStr quote scanning, lazy node wrappers,
' * typed accessors, Exists/ExistsKey/ExistsIndex/Keys helpers, raw field access, token iteration for huge arrays, and Stringify support for parsed JSON,
' * primitive VBA values, arrays, Collections, Dictionaries, and nested JSON nodes.
' * Designed for fast parsing and traversal without Dictionary, Collection, or per-node object allocation during parse.
' * @author UesleiDev
' * @version 1.0.4
' */

Option Explicit

#If VBA7 Then
    '/** @description Fast memory block copy used to attach and clear SAFEARRAY descriptors. */
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal length As LongPtr)
    
    '/** @description Returns the internal address of a dynamic array descriptor. */
    Private Declare PtrSafe Function VarPtrArray Lib "vbe7" Alias "VarPtr" (ByRef Var() As Any) As LongPtr
    
    '/** @description Performs ordinal UTF-16 string comparison without allocating key substrings. */
    Private Declare PtrSafe Function CompareStringOrdinal Lib "kernel32" (ByVal lpString1 As LongPtr, ByVal cchCount1 As Long, ByVal lpString2 As LongPtr, ByVal cchCount2 As Long, ByVal bIgnoreCase As Long) As Long
    
    '/** @description Native pointer size used when copying SAFEARRAY descriptor pointers. */
    Private Const PTR_SIZE As LongPtr = 8
#Else
    '/** @description Fast memory block copy used to attach and clear SAFEARRAY descriptors. */
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
    
    '/** @description Returns the internal address of a dynamic array descriptor. */
    Private Declare Function VarPtrArray Lib "msvbvm60" Alias "VarPtr" (ByRef Var() As Any) As Long
    
    '/** @description Performs ordinal UTF-16 string comparison without allocating key substrings. */
    Private Declare Function CompareStringOrdinal Lib "kernel32" (ByVal lpString1 As Long, ByVal cchCount1 As Long, ByVal lpString2 As Long, ByVal cchCount2 As Long, ByVal bIgnoreCase As Long) As Long
    
    '/** @description Native pointer size used when copying SAFEARRAY descriptor pointers. */
    Private Const PTR_SIZE As Long = 4
#End If

'/** @description Default minimum token capacity used for parsed documents. */
Private Const JSON_MIN_TOKEN_CAPACITY As Long = 65536

'/** @description Maximum initial token capacity guessed from text length before growth begins. */
Private Const JSON_MAX_INITIAL_TOKEN_CAPACITY As Long = 3000000

'/** @description Token capacity threshold after which growth changes from 2x to 1.5x. */
Private Const JSON_LARGE_TOKEN_CAPACITY As Long = 1048576

'/** @description Number of spaces used by default when pretty-printing JSON. */
Private Const JSON_DEFAULT_INDENT_SIZE As Long = 2

'/** @description Minimum initial character capacity used by the JSON string writer. */
Private Const JSON_WRITER_MIN_CAPACITY As Long = 4096

'/** @description Initial writer capacity used when serializing external VBA values. */
Private Const JSON_VALUE_WRITER_CAPACITY As Long = 65536

'/** @description Minimum object width that justifies a lazy key index. */
Private Const JSON_KEY_INDEX_MIN_CHILDREN As Long = 32

'/**
' * @struct SAFEARRAY1D
' * @brief Minimal one-dimensional SAFEARRAY descriptor used to alias a VBA String as Integer().
' */
Private Type SAFEARRAY1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    #If VBA7 Then
        pvData As LongPtr
    #Else
        pvData As Long
    #End If
    cElements As Long
    lLbound As Long
End Type

'/**
' * @enum JSType
' * @brief Internal JSON token type identifiers.
' */
Private Enum JSType
    jsNone = 0
    jsObject = 1
    jsArray = 2
    jsString = 3
    jsNumber = 4
    jsBool = 5
    jsNull = 6
End Enum

'/**
' * @struct JSToken
' * @brief Compact parsed JSON token storing hierarchy links and key/value slices.
' */
Private Type JSToken
    Type As Integer
    Parent As Long
    NextSibling As Long
    FirstChild As Long
    LastChild As Long
    ChildCount As Long
    KeyStart As Long
    KeyLen As Long
    ValStart As Long
    ValLen As Long
    KeyIndexBuilt As Boolean
    KeyLookupCount As Integer
    ChildIndexStart As Long
    IndexLookupCount As Integer
End Type

Private m_Chars() As Integer
Private m_Text As String
Private m_Length As Long
Private m_Index As Long
Private m_Tokens() As JSToken
Private m_TokenCount As Long
Private m_TokenCap As Long
Private m_KeyIndex As Object
Private m_ChildIndexTokens() As Long
Private m_ChildIndexCount As Long
Private m_ChildIndexCap As Long
Private m_RootId As Long
Private m_NodeId As Long
Private m_Document As JSON
Private m_CharAliasActive As Boolean
Private m_DecodedValueCache As Object

'/**
' * @method Class_Terminate
' * @brief Releases the active SAFEARRAY alias before the object is destroyed.
' */
Private Sub Class_Terminate()
    ClearCharAlias
    Set m_KeyIndex = Nothing
    Set m_DecodedValueCache = Nothing
End Sub

'/**
' * @method Parse
' * @brief Parses JSON text into a tokenized JSON document.
' * @param Text Source JSON string.
' * @return Parsed JSON document.
' */
Public Function Parse(ByRef Text As String) As JSON
    Dim doc As JSON
    Set doc = New JSON
    doc.LoadText Text
    Set Parse = doc
End Function

'/**
' * @method Stringify
' * @brief Serializes the current JSON document or node to JSON text.
' * @param Pretty Enables formatted output with line breaks and indentation.
' * @param IndentSize Number of spaces per indentation level.
' * @return JSON text.
' */
Public Function Stringify(Optional ByVal Pretty As Boolean = False, Optional ByVal IndentSize As Long = JSON_DEFAULT_INDENT_SIZE) As String
    Stringify = StringifyCurrent(Pretty, Space$(IndentSize))
End Function

'/**
' * @method StringifyWithIndent
' * @brief Serializes the current JSON document or node using a custom indentation string.
' * @param Pretty Enables formatted output with line breaks and indentation.
' * @param IndentText Text used for each indentation level, such as spaces or vbTab.
' * @return JSON text.
' */
Public Function StringifyWithIndent(Optional ByVal Pretty As Boolean = False, Optional ByVal IndentText As String = "  ") As String
    StringifyWithIndent = StringifyCurrent(Pretty, IndentText)
End Function

'/**
' * @method StringifyValue
' * @brief Serializes an external VBA value to JSON text.
' * @param Value VBA value, array, Collection, Dictionary, or JSON node.
' * @param Pretty Enables formatted output with line breaks and indentation.
' * @param IndentSize Number of spaces per indentation level.
' * @return JSON text.
' */
Public Function StringifyValue(ByVal Value As Variant, Optional ByVal Pretty As Boolean = False, Optional ByVal IndentSize As Long = JSON_DEFAULT_INDENT_SIZE) As String
    StringifyValue = StringifyAny(Value, Pretty, Space$(IndentSize), 0)
End Function

'/**
' * @method StringifyValueWithIndent
' * @brief Serializes an external VBA value using a custom indentation string.
' * @param Value VBA value, array, Collection, Dictionary, or JSON node.
' * @param Pretty Enables formatted output with line breaks and indentation.
' * @param IndentText Text used for each indentation level, such as spaces or vbTab.
' * @return JSON text.
' */
Public Function StringifyValueWithIndent(ByVal Value As Variant, Optional ByVal Pretty As Boolean = False, Optional ByVal IndentText As String = "  ") As String
    StringifyValueWithIndent = StringifyAny(Value, Pretty, IndentText, 0)
End Function

'/**
' * @method LoadText
' * @brief Initializes this instance as a root JSON document.
' * @param Text Source JSON string.
' */
Friend Sub LoadText(ByRef Text As String)
    ClearCharAlias
    
    m_Text = Text
    m_Length = Len(m_Text)
    m_Index = 0
    m_RootId = 0
    m_NodeId = 0
    Set m_Document = Nothing
    Set m_KeyIndex = Nothing
    Set m_DecodedValueCache = Nothing
    
    Erase m_Chars
    Erase m_Tokens
    Erase m_ChildIndexTokens
    m_ChildIndexCount = 0
    m_ChildIndexCap = 0
    
    If m_Length = 0 Then Exit Sub
    
    m_TokenCap = m_Length \ 24
    
    If m_TokenCap < JSON_MIN_TOKEN_CAPACITY Then
        m_TokenCap = JSON_MIN_TOKEN_CAPACITY
    End If
    
    If m_TokenCap > JSON_MAX_INITIAL_TOKEN_CAPACITY Then
        m_TokenCap = JSON_MAX_INITIAL_TOKEN_CAPACITY
    End If
    
    m_TokenCount = 0
    ReDim m_Tokens(1 To m_TokenCap)
    
    ReDim m_Chars(0 To m_Length - 1)
    
    #If VBA7 Then
        CopyMemory m_Chars(0), ByVal StrPtr(m_Text), CLngPtr(m_Length) * 2
    #Else
        CopyMemory m_Chars(0), ByVal StrPtr(m_Text), m_Length * 2&
    #End If
    
    m_CharAliasActive = False
    
    m_RootId = ParseValue(0, 0, 0)
End Sub

'/**
' * @method InitNode
' * @brief Initializes this instance as a lightweight wrapper around a token from another document.
' * @param TokenId Token to wrap.
' * @param Document Root document that owns the token buffer.
' */
Friend Sub InitNode(ByVal TokenId As Long, ByVal Document As JSON)
    m_NodeId = TokenId
    Set m_Document = Document
End Sub

'/**
' * @property Item
' * @brief Gets a child value by object key or array index.
' * @param Key Object key as String or array index as Long.
' * @return Primitive Variant or JSON node wrapper.
' */
Public Property Get Item(ByVal key As Variant) As Variant
Attribute Item.VB_UserMemId = 0
    Dim doc As JSON
    Dim baseId As Long
    Dim id As Long
    
    ResolveBase doc, baseId
    
    If baseId = 0 Then Exit Property
    
    id = doc.ResolveToken(baseId, key)
    If id = 0 Then Exit Property
    
    If doc.TokenTypeOf(id) = jsObject Or doc.TokenTypeOf(id) = jsArray Then
        Dim node As JSON
        Set node = New JSON
        node.InitNode id, doc
        Set Item = node
    Else
        Item = doc.PrimitiveValue(id)
    End If
End Property

'/**
' * @property Value
' * @brief Gets the current node value. Objects and arrays return the node itself.
' * @return Primitive Variant or current JSON node.
' */
Public Property Get Value() As Variant
    Dim doc As JSON
    Dim baseId As Long
    
    ResolveBase doc, baseId
    
    If baseId = 0 Then Exit Property
    
    If doc.TokenTypeOf(baseId) = jsObject Or doc.TokenTypeOf(baseId) = jsArray Then
        Set Value = Me
    Else
        Value = doc.PrimitiveValue(baseId)
    End If
End Property

'/**
' * @property Count
' * @brief Gets the amount of direct children in the current object or array.
' * @return Direct child count.
' */
Public Property Get Count() As Long
    Dim doc As JSON
    Dim baseId As Long
    
    ResolveBase doc, baseId
    
    If baseId = 0 Then Exit Property
    
    Count = doc.TokenChildCountOf(baseId)
End Property

'/**
' * @property JsonType
' * @brief Gets the JSON type name of the current node.
' * @return object, array, string, number, boolean, null, or empty string.
' */
Public Property Get JsonType() As String
    Dim doc As JSON
    Dim baseId As Long
    
    ResolveBase doc, baseId
    
    If baseId = 0 Then Exit Property
    
    Select Case doc.TokenTypeOf(baseId)
        Case jsObject
            JsonType = "object"
        Case jsArray
            JsonType = "array"
        Case jsString
            JsonType = "string"
        Case jsNumber
            JsonType = "number"
        Case jsBool
            JsonType = "boolean"
        Case jsNull
            JsonType = "null"
    End Select
End Property

'/**
' * @property IsObject
' * @brief Returns True when the current node is a JSON object.
' * @return Boolean state.
' */
Public Property Get IsObject() As Boolean
    Dim doc As JSON
    Dim baseId As Long
    
    ResolveBase doc, baseId
    
    If baseId = 0 Then Exit Property
    
    IsObject = (doc.TokenTypeOf(baseId) = jsObject)
End Property

'/**
' * @property IsArray
' * @brief Returns True when the current node is a JSON array.
' * @return Boolean state.
' */
Public Property Get IsArray() As Boolean
    Dim doc As JSON
    Dim baseId As Long
    
    ResolveBase doc, baseId
    
    If baseId = 0 Then Exit Property
    
    IsArray = (doc.TokenTypeOf(baseId) = jsArray)
End Property

'/**
' * @property IsNull
' * @brief Returns True when the current node is JSON null.
' * @return Boolean state.
' */
Public Property Get IsNull() As Boolean
    Dim doc As JSON
    Dim baseId As Long
    
    ResolveBase doc, baseId
    
    If baseId = 0 Then Exit Property
    
    IsNull = (doc.TokenTypeOf(baseId) = jsNull)
End Property

'/**
' * @method Exists
' * @brief Returns True when an object key or array index exists.
' * @param Key Object key as String or array index as Long.
' * @return Boolean state.
' */
Public Function Exists(ByVal key As Variant) As Boolean
    Dim doc As JSON
    Dim baseId As Long
    
    ResolveBase doc, baseId
    
    If baseId = 0 Then Exit Function
    
    Exists = (doc.ResolveToken(baseId, key) <> 0)
End Function

'/**
' * @method ExistsKey
' * @brief Returns True when an object key exists without using a Variant key parameter.
' * @param Key Object key.
' * @return Boolean state.
' */
Public Function ExistsKey(ByVal key As String) As Boolean
    Dim doc As JSON
    Dim baseId As Long
    
    ResolveBase doc, baseId
    
    If baseId = 0 Then Exit Function
    If doc.TokenTypeOf(baseId) <> jsObject Then Exit Function
    
    ExistsKey = (doc.FindChildFast(baseId, key) <> 0)
End Function

'/**
' * @method ExistsIndex
' * @brief Returns True when an array/object child index exists without using a Variant key parameter.
' * @param Index Zero-based child index.
' * @return Boolean state.
' */
Public Function ExistsIndex(ByVal Index As Long) As Boolean
    Dim doc As JSON
    Dim baseId As Long
    
    ResolveBase doc, baseId
    
    If baseId = 0 Then Exit Function
    If Index < 0 Then Exit Function
    
    ExistsIndex = (Index < doc.TokenChildCountOf(baseId))
End Function

'/**
' * @method Keys
' * @brief Gets all direct object keys or array indexes as a zero-based Variant array.
' * @return String array for objects, Long array for arrays, or an empty Variant array.
' */
Public Function keys() As Variant
    Dim doc As JSON
    Dim baseId As Long
    Dim child As Long
    Dim i As Long
    Dim total As Long
    
    ResolveBase doc, baseId
    
    If baseId = 0 Then
        keys = Array()
        Exit Function
    End If
    
    total = doc.TokenChildCountOf(baseId)
    
    If total <= 0 Then
        keys = Array()
        Exit Function
    End If
    
    If doc.TokenTypeOf(baseId) = jsObject Then
        Dim objectKeys() As String
        ReDim objectKeys(0 To total - 1)
        
        child = doc.TokenFirstChildOf(baseId)
        
        Do While child > 0
            objectKeys(i) = doc.GetRawKey(child)
            i = i + 1
            child = doc.TokenNextSiblingOf(child)
        Loop
        
        keys = objectKeys
    ElseIf doc.TokenTypeOf(baseId) = jsArray Then
        Dim arrayKeys() As Long
        ReDim arrayKeys(0 To total - 1)
        
        For i = 0 To total - 1
            arrayKeys(i) = i
        Next i
        
        keys = arrayKeys
    Else
        keys = Array()
    End If
End Function

'/**
' * @method Node
' * @brief Gets a child object or array by key/index without returning through Variant.
' * @param Key Object key as String or array index as Long.
' * @return JSON node wrapper.
' */
Public Function node(ByVal key As Variant) As JSON
    Dim doc As JSON
    Dim baseId As Long
    Dim id As Long
    
    ResolveBase doc, baseId
    
    If baseId = 0 Then Exit Function
    
    id = doc.ResolveToken(baseId, key)
    If id = 0 Then Exit Function
    
    If doc.TokenTypeOf(id) = jsObject Or doc.TokenTypeOf(id) = jsArray Then
        Dim nodeObj As JSON
        Set nodeObj = New JSON
        nodeObj.InitNode id, doc
        Set node = nodeObj
    End If
End Function

'/**
' * @method NodeKey
' * @brief Gets an object child node by key without using Variant dispatch.
' * @param Key Object key.
' * @return JSON node wrapper.
' */
Public Function NodeKey(ByVal key As String) As JSON
    Dim doc As JSON
    Dim baseId As Long
    Dim id As Long
    
    ResolveBase doc, baseId
    
    If baseId = 0 Then Exit Function
    If doc.TokenTypeOf(baseId) <> jsObject Then Exit Function
    
    id = doc.FindChildFast(baseId, key)
    If id = 0 Then Exit Function
    
    If doc.TokenTypeOf(id) = jsObject Or doc.TokenTypeOf(id) = jsArray Then
        Dim nodeObj As JSON
        Set nodeObj = New JSON
        nodeObj.InitNode id, doc
        Set NodeKey = nodeObj
    End If
End Function

'/**
' * @method NodeIndex
' * @brief Gets an array/object child node by zero-based index without using Variant dispatch.
' * @param Index Zero-based child index.
' * @return JSON node wrapper.
' */
Public Function NodeIndex(ByVal Index As Long) As JSON
    Set NodeIndex = NodeAt(Index)
End Function

'/**
' * @method NodeAt
' * @brief Gets an array/object child node by zero-based child index.
' * @param Index Zero-based child index.
' * @return JSON node wrapper.
' */
Public Function NodeAt(ByVal Index As Long) As JSON
    Dim doc As JSON
    Dim baseId As Long
    Dim id As Long
    
    ResolveBase doc, baseId
    
    If baseId = 0 Then Exit Function
    
    id = doc.FindIndexChild(baseId, Index)
    If id = 0 Then Exit Function
    
    If doc.TokenTypeOf(id) = jsObject Or doc.TokenTypeOf(id) = jsArray Then
        Dim nodeObj As JSON
        Set nodeObj = New JSON
        nodeObj.InitNode id, doc
        Set NodeAt = nodeObj
    End If
End Function

'/**
' * @method ValueAt
' * @brief Gets any child value by zero-based child index.
' * @param Index Zero-based child index.
' * @return Primitive Variant or JSON node wrapper.
' */
Public Function ValueAt(ByVal Index As Long) As Variant
    Dim doc As JSON
    Dim baseId As Long
    Dim id As Long
    
    ResolveBase doc, baseId
    
    If baseId = 0 Then Exit Function
    
    id = doc.FindIndexChild(baseId, Index)
    If id = 0 Then Exit Function
    
    If doc.TokenTypeOf(id) = jsObject Or doc.TokenTypeOf(id) = jsArray Then
        Dim node As JSON
        Set node = New JSON
        node.InitNode id, doc
        Set ValueAt = node
    Else
        ValueAt = doc.PrimitiveValue(id)
    End If
End Function

'/**
' * @method KeyAt
' * @brief Gets the key of an object child by zero-based child index.
' * @param Index Zero-based child index.
' * @return Object key text.
' */
Public Function KeyAt(ByVal Index As Long) As String
    Dim doc As JSON
    Dim baseId As Long
    Dim id As Long
    
    ResolveBase doc, baseId
    
    If baseId = 0 Then Exit Function
    
    id = doc.FindIndexChild(baseId, Index)
    If id = 0 Then Exit Function
    
    KeyAt = doc.GetRawKey(id)
End Function

'/**
' * @method StringValue
' * @brief Gets a child value as String by key/index, applying basic JSON unescape.
' * @param Key Object key as String or array index as Long.
' * @return String value.
' */
Public Function StringValue(ByVal key As Variant) As String
    Dim doc As JSON
    Dim baseId As Long
    Dim id As Long
    
    ResolveBase doc, baseId
    
    If baseId = 0 Then Exit Function
    
    id = doc.ResolveToken(baseId, key)
    If id = 0 Then Exit Function
    
    StringValue = doc.ValueAsString(id)
End Function

'/**
' * @method NumberValue
' * @brief Gets a child value as Double by key/index.
' * @param Key Object key as String or array index as Long.
' * @return Double value.
' */
Public Function NumberValue(ByVal key As Variant) As Double
    Dim doc As JSON
    Dim baseId As Long
    Dim id As Long
    
    ResolveBase doc, baseId
    
    If baseId = 0 Then Exit Function
    
    id = doc.ResolveToken(baseId, key)
    If id = 0 Then Exit Function
    
    NumberValue = doc.ValueAsDouble(id)
End Function

'/**
' * @method BoolValue
' * @brief Gets a child value as Boolean by key/index.
' * @param Key Object key as String or array index as Long.
' * @return Boolean value.
' */
Public Function BoolValue(ByVal key As Variant) As Boolean
    Dim doc As JSON
    Dim baseId As Long
    Dim id As Long
    
    ResolveBase doc, baseId
    
    If baseId = 0 Then Exit Function
    
    id = doc.ResolveToken(baseId, key)
    If id = 0 Then Exit Function
    
    BoolValue = doc.ValueAsBool(id)
End Function

'/**
' * @method RawStringValue
' * @brief Gets a child value as raw String by key/index without unescaping.
' * @param Key Object key as String or array index as Long.
' * @return Raw string slice.
' */
Public Function RawStringValue(ByVal key As Variant) As String
    Dim doc As JSON
    Dim baseId As Long
    Dim id As Long
    
    ResolveBase doc, baseId
    
    If baseId = 0 Then Exit Function
    
    id = doc.ResolveToken(baseId, key)
    If id = 0 Then Exit Function
    
    RawStringValue = doc.GetRawSlice(id)
End Function

'/**
' * @method StringKey
' * @brief Gets a string field by object key without using Variant dispatch.
' * @param Key Object key.
' * @return String value.
' */
Public Function StringKey(ByVal key As String) As String
    Dim doc As JSON
    Dim baseId As Long
    Dim id As Long
    
    ResolveBase doc, baseId
    
    If baseId = 0 Then Exit Function
    If doc.TokenTypeOf(baseId) <> jsObject Then Exit Function
    
    id = doc.FindChildFast(baseId, key)
    If id = 0 Then Exit Function
    
    StringKey = doc.ValueAsString(id)
End Function

'/**
' * @method NumberKey
' * @brief Gets a numeric field by object key without using Variant dispatch.
' * @param Key Object key.
' * @return Double value.
' */
Public Function NumberKey(ByVal key As String) As Double
    Dim doc As JSON
    Dim baseId As Long
    Dim id As Long
    
    ResolveBase doc, baseId
    
    If baseId = 0 Then Exit Function
    If doc.TokenTypeOf(baseId) <> jsObject Then Exit Function
    
    id = doc.FindChildFast(baseId, key)
    If id = 0 Then Exit Function
    
    NumberKey = doc.ValueAsDouble(id)
End Function

'/**
' * @method BoolKey
' * @brief Gets a Boolean field by object key without using Variant dispatch.
' * @param Key Object key.
' * @return Boolean value.
' */
Public Function BoolKey(ByVal key As String) As Boolean
    Dim doc As JSON
    Dim baseId As Long
    Dim id As Long
    
    ResolveBase doc, baseId
    
    If baseId = 0 Then Exit Function
    If doc.TokenTypeOf(baseId) <> jsObject Then Exit Function
    
    id = doc.FindChildFast(baseId, key)
    If id = 0 Then Exit Function
    
    BoolKey = doc.ValueAsBool(id)
End Function

'/**
' * @method RawStringKey
' * @brief Gets a raw JSON field by object key without using Variant dispatch.
' * @param Key Object key.
' * @return Raw JSON slice.
' */
Public Function RawStringKey(ByVal key As String) As String
    Dim doc As JSON
    Dim baseId As Long
    Dim id As Long
    
    ResolveBase doc, baseId
    
    If baseId = 0 Then Exit Function
    If doc.TokenTypeOf(baseId) <> jsObject Then Exit Function
    
    id = doc.FindChildFast(baseId, key)
    If id = 0 Then Exit Function
    
    RawStringKey = doc.GetRawSlice(id)
End Function

'/**
' * @method StringAt
' * @brief Gets an indexed child value as String, applying basic JSON unescape.
' * @param Index Zero-based child index.
' * @return String value.
' */
Public Function StringAt(ByVal Index As Long) As String
    Dim doc As JSON
    Dim baseId As Long
    Dim id As Long
    
    ResolveBase doc, baseId
    
    If baseId = 0 Then Exit Function
    
    id = doc.FindIndexChild(baseId, Index)
    If id = 0 Then Exit Function
    
    StringAt = doc.ValueAsString(id)
End Function

'/**
' * @method NumberAt
' * @brief Gets an indexed child value as Double.
' * @param Index Zero-based child index.
' * @return Double value.
' */
Public Function NumberAt(ByVal Index As Long) As Double
    Dim doc As JSON
    Dim baseId As Long
    Dim id As Long
    
    ResolveBase doc, baseId
    
    If baseId = 0 Then Exit Function
    
    id = doc.FindIndexChild(baseId, Index)
    If id = 0 Then Exit Function
    
    NumberAt = doc.ValueAsDouble(id)
End Function

'/**
' * @method BoolAt
' * @brief Gets an indexed child value as Boolean.
' * @param Index Zero-based child index.
' * @return Boolean value.
' */
Public Function BoolAt(ByVal Index As Long) As Boolean
    Dim doc As JSON
    Dim baseId As Long
    Dim id As Long
    
    ResolveBase doc, baseId
    
    If baseId = 0 Then Exit Function
    
    id = doc.FindIndexChild(baseId, Index)
    If id = 0 Then Exit Function
    
    BoolAt = doc.ValueAsBool(id)
End Function

'/**
' * @method RawStringAt
' * @brief Gets an indexed child value as raw String without unescaping.
' * @param Index Zero-based child index.
' * @return Raw string slice.
' */
Public Function RawStringAt(ByVal Index As Long) As String
    Dim doc As JSON
    Dim baseId As Long
    Dim id As Long
    
    ResolveBase doc, baseId
    
    If baseId = 0 Then Exit Function
    
    id = doc.FindIndexChild(baseId, Index)
    If id = 0 Then Exit Function
    
    RawStringAt = doc.GetRawSlice(id)
End Function

'/**
' * @method FirstChildToken
' * @brief Gets the first direct child token of the current node.
' * @return First child token id.
' */
Public Function FirstChildToken() As Long
    Dim doc As JSON
    Dim baseId As Long
    
    ResolveBase doc, baseId
    
    If baseId = 0 Then Exit Function
    
    FirstChildToken = doc.TokenFirstChildOf(baseId)
End Function

'/**
' * @method LastChildToken
' * @brief Gets the last direct child token of the current node.
' * @return Last child token id.
' */
Public Function LastChildToken() As Long
    Dim doc As JSON
    Dim baseId As Long
    
    ResolveBase doc, baseId
    
    If baseId = 0 Then Exit Function
    
    LastChildToken = doc.TokenLastChildOf(baseId)
End Function

'/**
' * @method NextToken
' * @brief Gets the next sibling token for a token.
' * @param TokenId Current token id.
' * @return Next sibling token id.
' */
Public Function NextToken(ByVal TokenId As Long) As Long
    Dim doc As JSON
    Dim baseId As Long
    
    ResolveBase doc, baseId
    
    If baseId = 0 Then Exit Function
    If TokenId < 1 Or TokenId > doc.TokenCount Then Exit Function
    
    NextToken = doc.TokenNextSiblingOf(TokenId)
End Function

'/**
' * @method NodeFromToken
' * @brief Wraps a token as a JSON node when the token is an object or array.
' * @param TokenId Token id.
' * @return JSON node wrapper.
' */
Public Function NodeFromToken(ByVal TokenId As Long) As JSON
    Dim doc As JSON
    Dim baseId As Long
    
    ResolveBase doc, baseId
    
    If baseId = 0 Then Exit Function
    If TokenId < 1 Or TokenId > doc.TokenCount Then Exit Function
    
    If doc.TokenTypeOf(TokenId) = jsObject Or doc.TokenTypeOf(TokenId) = jsArray Then
        Dim node As JSON
        Set node = New JSON
        node.InitNode TokenId, doc
        Set NodeFromToken = node
    End If
End Function

'/**
' * @method TokenKey
' * @brief Gets the object key associated with a token.
' * @param TokenId Token id.
' * @return Token key text.
' */
Public Function TokenKey(ByVal TokenId As Long) As String
    Dim doc As JSON
    Dim baseId As Long
    
    ResolveBase doc, baseId
    
    If baseId = 0 Then Exit Function
    If TokenId < 1 Or TokenId > doc.TokenCount Then Exit Function
    
    TokenKey = doc.GetRawKey(TokenId)
End Function

'/**
' * @method TokenValue
' * @brief Gets the token value as Variant.
' * @param TokenId Token id.
' * @return Primitive Variant or JSON node wrapper.
' */
Public Function TokenValue(ByVal TokenId As Long) As Variant
    Dim doc As JSON
    Dim baseId As Long
    
    ResolveBase doc, baseId
    
    If baseId = 0 Then Exit Function
    If TokenId < 1 Or TokenId > doc.TokenCount Then Exit Function
    
    If doc.TokenTypeOf(TokenId) = jsObject Or doc.TokenTypeOf(TokenId) = jsArray Then
        Dim node As JSON
        Set node = New JSON
        node.InitNode TokenId, doc
        Set TokenValue = node
    Else
        TokenValue = doc.PrimitiveValue(TokenId)
    End If
End Function

'/**
' * @method TokenStringValue
' * @brief Gets the token value as String, applying basic JSON unescape.
' * @param TokenId Token id.
' * @return String value.
' */
Public Function TokenStringValue(ByVal TokenId As Long) As String
    Dim doc As JSON
    Dim baseId As Long
    
    ResolveBase doc, baseId
    
    If baseId = 0 Then Exit Function
    If TokenId < 1 Or TokenId > doc.TokenCount Then Exit Function
    
    TokenStringValue = doc.ValueAsString(TokenId)
End Function

'/**
' * @method TokenRawStringValue
' * @brief Gets the token value as raw String without unescaping.
' * @param TokenId Token id.
' * @return Raw string slice.
' */
Public Function TokenRawStringValue(ByVal TokenId As Long) As String
    Dim doc As JSON
    Dim baseId As Long
    
    ResolveBase doc, baseId
    
    If baseId = 0 Then Exit Function
    If TokenId < 1 Or TokenId > doc.TokenCount Then Exit Function
    
    TokenRawStringValue = doc.GetRawSlice(TokenId)
End Function

'/**
' * @method TokenNumberValue
' * @brief Gets the token value as Double.
' * @param TokenId Token id.
' * @return Double value.
' */
Public Function TokenNumberValue(ByVal TokenId As Long) As Double
    Dim doc As JSON
    Dim baseId As Long
    
    ResolveBase doc, baseId
    
    If baseId = 0 Then Exit Function
    If TokenId < 1 Or TokenId > doc.TokenCount Then Exit Function
    
    TokenNumberValue = doc.ValueAsDouble(TokenId)
End Function

'/**
' * @method TokenBoolValue
' * @brief Gets the token value as Boolean.
' * @param TokenId Token id.
' * @return Boolean value.
' */
Public Function TokenBoolValue(ByVal TokenId As Long) As Boolean
    Dim doc As JSON
    Dim baseId As Long
    
    ResolveBase doc, baseId
    
    If baseId = 0 Then Exit Function
    If TokenId < 1 Or TokenId > doc.TokenCount Then Exit Function
    
    TokenBoolValue = doc.ValueAsBool(TokenId)
End Function

'/**
' * @method TokenString
' * @brief Gets a field from an object token as String, applying basic JSON unescape.
' * @param TokenId Object token id.
' * @param Key Field key.
' * @return String value.
' */
Public Function TokenString(ByVal TokenId As Long, ByVal key As Variant) As String
    Dim doc As JSON
    Dim baseId As Long
    Dim id As Long
    
    ResolveBase doc, baseId
    
    If baseId = 0 Then Exit Function
    If TokenId < 1 Or TokenId > doc.TokenCount Then Exit Function
    
    id = doc.ResolveToken(TokenId, key)
    If id = 0 Then Exit Function
    
    TokenString = doc.ValueAsString(id)
End Function

'/**
' * @method TokenRawString
' * @brief Gets a field from an object token as raw String without unescaping.
' * @param TokenId Object token id.
' * @param Key Field key.
' * @return Raw string slice.
' */
Public Function TokenRawString(ByVal TokenId As Long, ByVal key As Variant) As String
    Dim doc As JSON
    Dim baseId As Long
    Dim id As Long
    
    ResolveBase doc, baseId
    
    If baseId = 0 Then Exit Function
    If TokenId < 1 Or TokenId > doc.TokenCount Then Exit Function
    
    id = doc.ResolveToken(TokenId, key)
    If id = 0 Then Exit Function
    
    TokenRawString = doc.GetRawSlice(id)
End Function

'/**
' * @method TokenRawField
' * @brief Gets a raw string field from an object token using a schema-known String key.
' * @param TokenId Object token id.
' * @param Key Field key.
' * @return Raw string slice.
' */
Public Function TokenRawField(ByVal TokenId As Long, ByVal key As String) As String
    Dim doc As JSON
    Dim baseId As Long
    Dim id As Long
    
    ResolveBase doc, baseId
    
    If baseId = 0 Then Exit Function
    If TokenId < 1 Or TokenId > doc.TokenCount Then Exit Function
    
    id = doc.FindChildFast(TokenId, key)
    If id = 0 Then Exit Function
    
    TokenRawField = doc.GetRawSlice(id)
End Function

'/**
' * @method TokenNumber
' * @brief Gets a field from an object token as Double.
' * @param TokenId Object token id.
' * @param Key Field key.
' * @return Double value.
' */
Public Function TokenNumber(ByVal TokenId As Long, ByVal key As Variant) As Double
    Dim doc As JSON
    Dim baseId As Long
    Dim id As Long
    
    ResolveBase doc, baseId
    
    If baseId = 0 Then Exit Function
    If TokenId < 1 Or TokenId > doc.TokenCount Then Exit Function
    
    id = doc.ResolveToken(TokenId, key)
    If id = 0 Then Exit Function
    
    TokenNumber = doc.ValueAsDouble(id)
End Function

'/**
' * @method TokenBool
' * @brief Gets a field from an object token as Boolean.
' * @param TokenId Object token id.
' * @param Key Field key.
' * @return Boolean value.
' */
Public Function TokenBool(ByVal TokenId As Long, ByVal key As Variant) As Boolean
    Dim doc As JSON
    Dim baseId As Long
    Dim id As Long
    
    ResolveBase doc, baseId
    
    If baseId = 0 Then Exit Function
    If TokenId < 1 Or TokenId > doc.TokenCount Then Exit Function
    
    id = doc.ResolveToken(TokenId, key)
    If id = 0 Then Exit Function
    
    TokenBool = doc.ValueAsBool(id)
End Function

'/**
' * @method TokenNode
' * @brief Gets a field from an object token as a JSON node.
' * @param TokenId Object token id.
' * @param Key Field key.
' * @return JSON node wrapper.
' */
Public Function TokenNode(ByVal TokenId As Long, ByVal key As Variant) As JSON
    Dim doc As JSON
    Dim baseId As Long
    Dim id As Long
    
    ResolveBase doc, baseId
    
    If baseId = 0 Then Exit Function
    If TokenId < 1 Or TokenId > doc.TokenCount Then Exit Function
    
    id = doc.ResolveToken(TokenId, key)
    If id = 0 Then Exit Function
    
    If doc.TokenTypeOf(id) = jsObject Or doc.TokenTypeOf(id) = jsArray Then
        Dim node As JSON
        Set node = New JSON
        node.InitNode id, doc
        Set TokenNode = node
    End If
End Function


'/**
' * @property TokenCount
' * @brief Gets the total parsed token count for internal wrapper calls.
' * @return Total token count.
' */
Friend Property Get TokenCount() As Long
    TokenCount = m_TokenCount
End Property

'/**
' * @method TokenTypeOf
' * @brief Gets the internal type of a token.
' * @param TokenId Token id.
' * @return Internal token type id.
' */
Friend Function TokenTypeOf(ByVal TokenId As Long) As Integer
    If TokenId < 1 Or TokenId > m_TokenCount Then Exit Function
    TokenTypeOf = m_Tokens(TokenId).Type
End Function

'/**
' * @method TokenChildCountOf
' * @brief Gets the direct child count of a token.
' * @param TokenId Token id.
' * @return Direct child count.
' */
Friend Function TokenChildCountOf(ByVal TokenId As Long) As Long
    If TokenId < 1 Or TokenId > m_TokenCount Then Exit Function
    TokenChildCountOf = m_Tokens(TokenId).ChildCount
End Function

'/**
' * @method TokenFirstChildOf
' * @brief Gets the first child token of a token.
' * @param TokenId Token id.
' * @return First child token id.
' */
Friend Function TokenFirstChildOf(ByVal TokenId As Long) As Long
    If TokenId < 1 Or TokenId > m_TokenCount Then Exit Function
    TokenFirstChildOf = m_Tokens(TokenId).FirstChild
End Function

'/**
' * @method TokenLastChildOf
' * @brief Gets the last child token of a token.
' * @param TokenId Token id.
' * @return Last child token id.
' */
Friend Function TokenLastChildOf(ByVal TokenId As Long) As Long
    If TokenId < 1 Or TokenId > m_TokenCount Then Exit Function
    TokenLastChildOf = m_Tokens(TokenId).LastChild
End Function

'/**
' * @method TokenNextSiblingOf
' * @brief Gets the next sibling token of a token.
' * @param TokenId Token id.
' * @return Next sibling token id.
' */
Friend Function TokenNextSiblingOf(ByVal TokenId As Long) As Long
    If TokenId < 1 Or TokenId > m_TokenCount Then Exit Function
    TokenNextSiblingOf = m_Tokens(TokenId).NextSibling
End Function

'/**
' * @method ClearCharAlias
' * @brief Clears the character SAFEARRAY alias if it is active.
' */
Private Sub ClearCharAlias()
    If m_CharAliasActive Then
        #If VBA7 Then
            Dim nullPtr As LongPtr
        #Else
            Dim nullPtr As Long
        #End If
        
        nullPtr = 0
        CopyMemory ByVal VarPtrArray(m_Chars), nullPtr, PTR_SIZE
        m_CharAliasActive = False
    End If
    
    Erase m_Chars
End Sub

'/**
' * @method ResolveBase
' * @brief Resolves the root document and base token for root instances or node wrappers.
' * @param Document Receives the real root document.
' * @param BaseId Receives the active token id.
' */
Private Sub ResolveBase(ByRef Document As JSON, ByRef baseId As Long)
    If m_NodeId > 0 Then
        Set Document = m_Document
        baseId = m_NodeId
    Else
        Set Document = Me
        baseId = m_RootId
    End If
End Sub

'/**
' * @method ResolveToken
' * @brief Resolves a direct child token by object key or array index.
' * @param ParentId Parent token id.
' * @param Key Object key as String or array index as Long.
' * @return Child token id.
' */
Friend Function ResolveToken(ByVal ParentId As Long, ByVal key As Variant) As Long
    If ParentId < 1 Or ParentId > m_TokenCount Then Exit Function
    
    If VarType(key) = vbString Then
        ResolveToken = FindChildFast(ParentId, CStr(key))
    Else
        ResolveToken = FindIndexChild(ParentId, CLng(key))
    End If
End Function

'/**
' * @method PrimitiveValue
' * @brief Converts a primitive token into a Variant.
' * @param TokenId Primitive token id.
' * @return Primitive Variant value.
' */
Friend Function PrimitiveValue(ByVal TokenId As Long) As Variant
    If TokenId < 1 Or TokenId > m_TokenCount Then Exit Function
    
    Select Case m_Tokens(TokenId).Type
        Case jsString
            PrimitiveValue = GetRawValue(TokenId)
        Case jsNumber
            PrimitiveValue = val(GetRawSlice(TokenId))
        Case jsBool
            PrimitiveValue = ValueAsBool(TokenId)
        Case jsNull
            PrimitiveValue = Null
    End Select
End Function

'/**
' * @method ValueAsString
' * @brief Reads a primitive token as String.
' * @param TokenId Token id.
' * @return String value.
' */
Friend Function ValueAsString(ByVal TokenId As Long) As String
    If TokenId < 1 Or TokenId > m_TokenCount Then Exit Function
    
    Select Case m_Tokens(TokenId).Type
        Case jsString
            ValueAsString = GetRawValue(TokenId)
        Case jsNumber, jsBool
            ValueAsString = GetRawSlice(TokenId)
        Case jsNull
            ValueAsString = vbNullString
    End Select
End Function

'/**
' * @method ValueAsDouble
' * @brief Reads a number token as Double.
' * @param TokenId Token id.
' * @return Double value.
' */
Friend Function ValueAsDouble(ByVal TokenId As Long) As Double
    If TokenId < 1 Or TokenId > m_TokenCount Then Exit Function
    
    If m_Tokens(TokenId).Type = jsNumber Then
        ValueAsDouble = val(GetRawSlice(TokenId))
    End If
End Function

'/**
' * @method ValueAsBool
' * @brief Reads a boolean token as Boolean.
' * @param TokenId Token id.
' * @return Boolean value.
' */
Friend Function ValueAsBool(ByVal TokenId As Long) As Boolean
    If TokenId < 1 Or TokenId > m_TokenCount Then Exit Function
    
    If m_Tokens(TokenId).Type = jsBool Then
        ValueAsBool = (m_Chars(m_Tokens(TokenId).ValStart - 1) = 116)
    End If
End Function

'/**
' * @method AddToken
' * @brief Adds a new token and grows the token buffer when required.
' * @return Created token id.
' */
Private Function AddToken() As Long
    m_TokenCount = m_TokenCount + 1
    
    If m_TokenCount > m_TokenCap Then
        If m_TokenCap < JSON_LARGE_TOKEN_CAPACITY Then
            m_TokenCap = m_TokenCap * 2
        Else
            m_TokenCap = m_TokenCap + (m_TokenCap \ 2)
        End If
        
        ReDim Preserve m_Tokens(1 To m_TokenCap)
    End If
    
    AddToken = m_TokenCount
End Function

'/**
' * @method ParseValue
' * @brief Parses any JSON value and returns the created token id.
' * @param ParentId Parent token id.
' * @param KeyStart One-based key start position in the original text.
' * @param KeyLen Key length.
' * @return Created token id.
' */
Private Function ParseValue(ByVal ParentId As Long, ByVal KeyStart As Long, ByVal KeyLen As Long) As Long
    SkipWhitespace
    
    If m_Index >= m_Length Then Exit Function
    
    Dim id As Long
    id = AddToken()
    ParseValue = id
    
    With m_Tokens(id)
        .Parent = ParentId
        .KeyStart = KeyStart
        .KeyLen = KeyLen
    End With
    
    If ParentId > 0 Then
        If m_Tokens(ParentId).FirstChild = 0 Then
            m_Tokens(ParentId).FirstChild = id
        Else
            m_Tokens(m_Tokens(ParentId).LastChild).NextSibling = id
        End If
        
        m_Tokens(ParentId).LastChild = id
        m_Tokens(ParentId).ChildCount = m_Tokens(ParentId).ChildCount + 1
    End If
    
    Select Case m_Chars(m_Index)
        Case 123
            m_Tokens(id).Type = jsObject
            ParseObject id
        
        Case 91
            m_Tokens(id).Type = jsArray
            ParseArray id
        
        Case 34
            m_Tokens(id).Type = jsString
            ParseString id
        
        Case 116
            m_Tokens(id).Type = jsBool
            m_Tokens(id).ValStart = m_Index + 1
            m_Tokens(id).ValLen = 4
            m_Index = m_Index + 4
        
        Case 102
            m_Tokens(id).Type = jsBool
            m_Tokens(id).ValStart = m_Index + 1
            m_Tokens(id).ValLen = 5
            m_Index = m_Index + 5
        
        Case 110
            m_Tokens(id).Type = jsNull
            m_Tokens(id).ValStart = m_Index + 1
            m_Tokens(id).ValLen = 4
            m_Index = m_Index + 4
        
        Case 45, 48 To 57
            m_Tokens(id).Type = jsNumber
            ParseNumber id
        
        Case Else
            m_Index = m_Length
    End Select
End Function

'/**
' * @method SkipWhitespace
' * @brief Skips JSON whitespace characters.
' */
Private Sub SkipWhitespace()
    Do While m_Index < m_Length
        Select Case m_Chars(m_Index)
            Case 32, 9, 10, 13
                m_Index = m_Index + 1
            Case Else
                Exit Do
        End Select
    Loop
End Sub

'/**
' * @method ParseObject
' * @brief Parses a JSON object.
' * @param id Object token id.
' */
Private Sub ParseObject(ByVal id As Long)
    m_Index = m_Index + 1
    
    Do While m_Index < m_Length
        SkipWhitespace
        
        If m_Index >= m_Length Then Exit Do
        
        If m_Chars(m_Index) = 125 Then
            m_Index = m_Index + 1
            Exit Sub
        End If
        
        Dim kStart As Long
        Dim kLen As Long
        
        If m_Chars(m_Index) = 34 Then
            ScanJSONString kStart, kLen
        End If
        
        SkipWhitespace
        
        If m_Index >= m_Length Then Exit Do
        
        If m_Chars(m_Index) = 58 Then
            m_Index = m_Index + 1
        End If
        
        ParseValue id, kStart, kLen
        
        SkipWhitespace
        
        If m_Index >= m_Length Then Exit Do
        
        If m_Chars(m_Index) = 44 Then
            m_Index = m_Index + 1
        End If
    Loop
End Sub

'/**
' * @method ParseArray
' * @brief Parses a JSON array.
' * @param id Array token id.
' */
Private Sub ParseArray(ByVal id As Long)
    m_Index = m_Index + 1
    
    Do While m_Index < m_Length
        SkipWhitespace
        
        If m_Index >= m_Length Then Exit Do
        
        If m_Chars(m_Index) = 93 Then
            m_Index = m_Index + 1
            Exit Sub
        End If
        
        ParseValue id, 0, 0
        
        SkipWhitespace
        
        If m_Index >= m_Length Then Exit Do
        
        If m_Chars(m_Index) = 44 Then
            m_Index = m_Index + 1
        End If
    Loop
End Sub

'/**
' * @method ParseString
' * @brief Parses a JSON string token.
' * @param id String token id.
' */
Private Sub ParseString(ByVal id As Long)
    ScanJSONString m_Tokens(id).ValStart, m_Tokens(id).ValLen
End Sub

'/**
' * @method ScanJSONString
' * @brief Scans a JSON string using native InStr for quote search and m_Chars for escape checks.
' * @param StartPos Receives one-based string content start.
' * @param LengthOut Receives string content length.
' */
Private Sub ScanJSONString(ByRef StartPos As Long, ByRef LengthOut As Long)
    Dim quotePos As Long
    Dim searchFrom As Long
    Dim slashCount As Long
    Dim p As Long
    Dim firstChar As Long
    
    m_Index = m_Index + 1
    StartPos = m_Index + 1
    searchFrom = StartPos
    firstChar = StartPos - 1
    
    Do
        quotePos = InStr(searchFrom, m_Text, """")
        
        If quotePos = 0 Then
            LengthOut = m_Length - StartPos + 1
            m_Index = m_Length
            Exit Sub
        End If
        
        slashCount = 0
        p = quotePos - 2
        
        Do While p >= firstChar
            If m_Chars(p) <> 92 Then Exit Do
            slashCount = slashCount + 1
            p = p - 1
        Loop
        
        If (slashCount And 1) = 0 Then
            LengthOut = quotePos - StartPos
            m_Index = quotePos
            Exit Sub
        End If
        
        searchFrom = quotePos + 1
    Loop
End Sub

'/**
' * @method ParseNumber
' * @brief Parses a JSON number as a raw slice.
' * @param id Number token id.
' */
Private Sub ParseNumber(ByVal id As Long)
    m_Tokens(id).ValStart = m_Index + 1
    
    Do While m_Index < m_Length
        Select Case m_Chars(m_Index)
            Case 45, 46, 48 To 57, 69, 101, 43
                m_Index = m_Index + 1
            Case Else
                Exit Do
        End Select
    Loop
    
    m_Tokens(id).ValLen = m_Index - m_Tokens(id).ValStart + 1
End Sub

'/**
' * @method FindChildFast
' * @brief Finds an object child by key without allocating key strings.
' * @param ParentId Parent token id.
' * @param KeyName Key name.
' * @return Child token id.
' */
Friend Function FindChildFast(ByVal ParentId As Long, ByRef KeyName As String) As Long
    Dim child As Long
    Dim indexKey As String
    
    If ParentId < 1 Or ParentId > m_TokenCount Then Exit Function
    If m_Tokens(ParentId).Type <> jsObject Then Exit Function
    
    If m_Tokens(ParentId).KeyIndexBuilt Then
        indexKey = ObjectKeyIndexKey(ParentId, KeyName)
        
        If m_KeyIndex.Exists(indexKey) Then
            FindChildFast = m_KeyIndex.Item(indexKey)
        End If
        Exit Function
    End If
    
    If m_Tokens(ParentId).ChildCount >= JSON_KEY_INDEX_MIN_CHILDREN Then
        If m_Tokens(ParentId).KeyLookupCount < 2 Then
            m_Tokens(ParentId).KeyLookupCount = m_Tokens(ParentId).KeyLookupCount + 1
        End If
        
        If m_Tokens(ParentId).KeyLookupCount >= 2 Then
            BuildObjectKeyIndex ParentId
            indexKey = ObjectKeyIndexKey(ParentId, KeyName)
            
            If m_KeyIndex.Exists(indexKey) Then
                FindChildFast = m_KeyIndex.Item(indexKey)
            End If
            Exit Function
        End If
    End If
    
    child = m_Tokens(ParentId).FirstChild
    
    Do While child > 0
        If KeyEquals(child, KeyName) Then
            FindChildFast = child
            Exit Function
        End If
        
        child = m_Tokens(child).NextSibling
    Loop
End Function

'/**
' * @method KeyEquals
' * @brief Compares a token key with a VBA string using CompareStringOrdinal.
' * @param TokenId Token id.
' * @param KeyName Key name.
' * @return Boolean comparison result.
' */
Private Function KeyEquals(ByVal TokenId As Long, ByRef KeyName As String) As Boolean
    If m_Tokens(TokenId).KeyLen = 0 Then Exit Function

    If m_Tokens(TokenId).KeyLen <> Len(KeyName) Then
        If KeyContainsEscape(TokenId) Then
            KeyEquals = (StrComp(GetRawKey(TokenId), KeyName, vbBinaryCompare) = 0)
        End If
        Exit Function
    End If

    #If VBA7 Then
        Dim pJson As LongPtr
        Dim pKey As LongPtr
    #Else
        Dim pJson As Long
        Dim pKey As Long
    #End If

    pJson = StrPtr(m_Text) + ((m_Tokens(TokenId).KeyStart - 1) * 2)
    pKey = StrPtr(KeyName)

    KeyEquals = (CompareStringOrdinal(pJson, m_Tokens(TokenId).KeyLen, pKey, Len(KeyName), 0) = 2)
End Function

'/**
' * @method KeyContainsEscape
' * @brief Reports whether an object key source slice contains a JSON escape marker.
' * @param TokenId Token id containing the object key.
' * @return True when the raw key needs JSON decoding before comparison.
' */
Private Function KeyContainsEscape(ByVal TokenId As Long) As Boolean
    Dim i As Long
    Dim lastChar As Long

    lastChar = m_Tokens(TokenId).KeyStart + m_Tokens(TokenId).KeyLen - 2
    
    For i = m_Tokens(TokenId).KeyStart - 1 To lastChar
        If m_Chars(i) = 92 Then
            KeyContainsEscape = True
            Exit Function
        End If
    Next i
End Function

'/**
' * @method BuildObjectKeyIndex
' * @brief Builds normalized key lookup entries for one object token.
' * @param ParentId Object token id.
' */
Private Sub BuildObjectKeyIndex(ByVal ParentId As Long)
    Dim child As Long
    Dim indexKey As String
    
    If m_Tokens(ParentId).KeyIndexBuilt Then Exit Sub
    
    If m_KeyIndex Is Nothing Then
        Set m_KeyIndex = CreateObject("Scripting.Dictionary")
        m_KeyIndex.CompareMode = vbBinaryCompare
    End If
    
    child = m_Tokens(ParentId).FirstChild
    
    Do While child > 0
        indexKey = ObjectKeyIndexKey(ParentId, GetRawKey(child))
        
        If Not m_KeyIndex.Exists(indexKey) Then
            m_KeyIndex.Add indexKey, child
        End If
        
        child = m_Tokens(child).NextSibling
    Loop
    
    m_Tokens(ParentId).KeyIndexBuilt = True
End Sub

'/**
' * @method ObjectKeyIndexKey
' * @brief Creates a collision-free dictionary key scoped to an object token.
' * @param ParentId Object token id.
' * @param KeyName Decoded member name.
' * @return Dictionary key for the member name in its parent object.
' */
Private Function ObjectKeyIndexKey(ByVal ParentId As Long, ByVal KeyName As String) As String
    ObjectKeyIndexKey = CStr(ParentId) & ChrW$(0) & KeyName
End Function

'/**
' * @method FindIndexChild
' * @brief Finds a direct child by zero-based child index.
' * @param ParentId Parent token id.
' * @param Index Zero-based child index.
' * @return Child token id.
' */
Friend Function FindIndexChild(ByVal ParentId As Long, ByVal Index As Long) As Long
    If ParentId < 1 Or ParentId > m_TokenCount Then Exit Function
    If Index < 0 Or Index >= m_Tokens(ParentId).ChildCount Then Exit Function
    
    If m_Tokens(ParentId).ChildIndexStart > 0 Then
        FindIndexChild = m_ChildIndexTokens(m_Tokens(ParentId).ChildIndexStart + Index)
        Exit Function
    End If
    
    If m_Tokens(ParentId).ChildCount >= JSON_CHILD_INDEX_MIN_CHILDREN Then
        If m_Tokens(ParentId).IndexLookupCount < 2 Then
            m_Tokens(ParentId).IndexLookupCount = m_Tokens(ParentId).IndexLookupCount + 1
        End If
        
        If m_Tokens(ParentId).IndexLookupCount >= 2 Then
            BuildChildIndex ParentId
            FindIndexChild = m_ChildIndexTokens(m_Tokens(ParentId).ChildIndexStart + Index)
            Exit Function
        End If
    End If
    
    Dim child As Long
    Dim i As Long
    
    child = m_Tokens(ParentId).FirstChild
    
    Do While child > 0
        If i = Index Then
            FindIndexChild = child
            Exit Function
        End If
        
        i = i + 1
        child = m_Tokens(child).NextSibling
    Loop
End Function

'/**
' * @method BuildChildIndex
' * @brief Stores direct child token ids in a shared lookup buffer for one container.
' * @param ParentId Container token id.
' */
Private Sub BuildChildIndex(ByVal ParentId As Long)
    Dim child As Long
    Dim i As Long
    Dim startAt As Long
    Dim total As Long
    
    If m_Tokens(ParentId).ChildIndexStart > 0 Then Exit Sub
    
    total = m_Tokens(ParentId).ChildCount
    If total <= 0 Then Exit Sub
    
    EnsureChildIndexCapacity total
    startAt = m_ChildIndexCount + 1
    child = m_Tokens(ParentId).FirstChild
    
    For i = 0 To total - 1
        m_ChildIndexTokens(startAt + i) = child
        child = m_Tokens(child).NextSibling
    Next i
    
    m_ChildIndexCount = m_ChildIndexCount + total
    m_Tokens(ParentId).ChildIndexStart = startAt
End Sub

'/**
' * @method EnsureChildIndexCapacity
' * @brief Grows the shared lazy child-index buffer as required.
' * @param Additional Number of child ids about to be appended.
' */
Private Sub EnsureChildIndexCapacity(ByVal Additional As Long)
    Dim required As Long
    Dim newCapacity As Long
    
    required = m_ChildIndexCount + Additional
    If required <= m_ChildIndexCap Then Exit Sub
    
    newCapacity = m_ChildIndexCap
    If newCapacity = 0 Then
        newCapacity = JSON_CHILD_INDEX_MIN_CHILDREN
        Do While newCapacity < required
            newCapacity = newCapacity * 2
        Loop
        
        ReDim m_ChildIndexTokens(1 To newCapacity)
        m_ChildIndexCap = newCapacity
        Exit Sub
    End If
    
    Do While newCapacity < required
        newCapacity = newCapacity * 2
    Loop
    
    ReDim Preserve m_ChildIndexTokens(1 To newCapacity)
    m_ChildIndexCap = newCapacity
End Sub

'/**
' * @method GetRawKey
' * @brief Gets an object key string from the original JSON text.
' * @param TokenId Token id.
' * @return Key string.
' */
Friend Function GetRawKey(ByVal TokenId As Long) As String
    If TokenId < 1 Or TokenId > m_TokenCount Then Exit Function
    If m_Tokens(TokenId).KeyLen <= 0 Then Exit Function
    
    GetRawKey = Mid$(m_Text, m_Tokens(TokenId).KeyStart, m_Tokens(TokenId).KeyLen)
    
    If InStr(1, GetRawKey, "\") > 0 Then
        GetRawKey = UnescapeJSONString(GetRawKey)
    End If
End Function

'/**
' * @method GetRawValue
' * @brief Gets a value string from the original JSON text, applying basic JSON unescape.
' * @param TokenId Token id.
' * @return Value string.
' */
Friend Function GetRawValue(ByVal TokenId As Long) As String
    If TokenId < 1 Or TokenId > m_TokenCount Then Exit Function
    If m_Tokens(TokenId).ValLen <= 0 Then Exit Function
    
    GetRawValue = Mid$(m_Text, m_Tokens(TokenId).ValStart, m_Tokens(TokenId).ValLen)
    
    If InStr(1, GetRawValue, "\") > 0 Then
        If Not m_DecodedValueCache Is Nothing Then
            If m_DecodedValueCache.Exists(CStr(TokenId)) Then
                GetRawValue = m_DecodedValueCache.Item(CStr(TokenId))
                Exit Function
            End If
        End If
        
        GetRawValue = UnescapeJSONString(GetRawValue)
        
        If m_DecodedValueCache Is Nothing Then
            Set m_DecodedValueCache = CreateObject("Scripting.Dictionary")
        End If
        
        m_DecodedValueCache.Add CStr(TokenId), GetRawValue
    End If
End Function

'/**
' * @method GetRawSlice
' * @brief Gets a value string slice without unescaping.
' * @param TokenId Token id.
' * @return Raw value string.
' */
Friend Function GetRawSlice(ByVal TokenId As Long) As String
    If TokenId < 1 Or TokenId > m_TokenCount Then Exit Function
    If m_Tokens(TokenId).ValLen <= 0 Then Exit Function
    
    GetRawSlice = Mid$(m_Text, m_Tokens(TokenId).ValStart, m_Tokens(TokenId).ValLen)
End Function

'/**
' * @method UnescapeJSONString
' * @brief Applies basic JSON string unescaping.
' * @param Text Escaped JSON string content.
' * @return Unescaped string content.
' */
Private Function UnescapeJSONString(ByVal Text As String) As String
    Dim Buffer As String
    Dim Position As Long
    Dim Capacity As Long
    Dim length As Long
    Dim slashPos As Long
    Dim scanPos As Long
    Dim nextChar As String
    Dim CodeUnit As Long
    Dim lowUnit As Long
    Dim CodePoint As Long
    
    length = Len(Text)
    If length = 0 Then Exit Function
    
    WriterStart Buffer, Position, Capacity, length
    scanPos = 1
    
    Do While scanPos <= length
        slashPos = InStr(scanPos, Text, "\")
        
        If slashPos = 0 Then
            WriterAppend Buffer, Position, Capacity, Mid$(Text, scanPos)
            Exit Do
        End If
        
        If slashPos > scanPos Then
            WriterAppend Buffer, Position, Capacity, Mid$(Text, scanPos, slashPos - scanPos)
        End If
        
        If slashPos = length Then
            WriterAppend Buffer, Position, Capacity, "\"
            Exit Do
        End If
        
        nextChar = Mid$(Text, slashPos + 1, 1)
        
        Select Case nextChar
            Case """"
                WriterAppend Buffer, Position, Capacity, """"
                scanPos = slashPos + 2
            
            Case "\"
                WriterAppend Buffer, Position, Capacity, "\"
                scanPos = slashPos + 2
            
            Case "/"
                WriterAppend Buffer, Position, Capacity, "/"
                scanPos = slashPos + 2
            
            Case "b"
                WriterAppend Buffer, Position, Capacity, Chr$(8)
                scanPos = slashPos + 2
            
            Case "f"
                WriterAppend Buffer, Position, Capacity, Chr$(12)
                scanPos = slashPos + 2
            
            Case "n"
                WriterAppend Buffer, Position, Capacity, vbLf
                scanPos = slashPos + 2
            
            Case "r"
                WriterAppend Buffer, Position, Capacity, vbCr
                scanPos = slashPos + 2
            
            Case "t"
                WriterAppend Buffer, Position, Capacity, vbTab
                scanPos = slashPos + 2
            
            Case "u"
                If slashPos + 5 <= length Then
                    CodeUnit = Hex4ToLong(Text, slashPos + 2)
                    
                    If CodeUnit >= 0 Then
                        If CodeUnit >= &HD800& And CodeUnit <= &HDBFF& Then
                            If slashPos + 11 <= length Then
                                If Mid$(Text, slashPos + 6, 2) = "\u" Then
                                    lowUnit = Hex4ToLong(Text, slashPos + 8)
                                    
                                    If lowUnit >= &HDC00& And lowUnit <= &HDFFF& Then
                                        CodePoint = &H10000 + ((CodeUnit - &HD800&) * &H400&) + (lowUnit - &HDC00&)
                                        WriterAppend Buffer, Position, Capacity, CodePointToString(CodePoint)
                                        scanPos = slashPos + 12
                                    Else
                                        WriterAppend Buffer, Position, Capacity, CodePointToString(CodeUnit)
                                        scanPos = slashPos + 6
                                    End If
                                Else
                                    WriterAppend Buffer, Position, Capacity, CodePointToString(CodeUnit)
                                    scanPos = slashPos + 6
                                End If
                            Else
                                WriterAppend Buffer, Position, Capacity, CodePointToString(CodeUnit)
                                scanPos = slashPos + 6
                            End If
                        Else
                            WriterAppend Buffer, Position, Capacity, CodePointToString(CodeUnit)
                            scanPos = slashPos + 6
                        End If
                    Else
                        WriterAppend Buffer, Position, Capacity, "\u"
                        scanPos = slashPos + 2
                    End If
                Else
                    WriterAppend Buffer, Position, Capacity, "\u"
                    scanPos = slashPos + 2
                End If
            
            Case Else
                WriterAppend Buffer, Position, Capacity, nextChar
                scanPos = slashPos + 2
        End Select
    Loop
    
    UnescapeJSONString = WriterFinish(Buffer, Position)
End Function

'/**
' * @method Hex4ToLong
' * @brief Converts four hexadecimal characters from Text into a Long value.
' * @param Text Source text.
' * @param StartPos One-based start position.
' * @return Parsed value, or -1 when invalid.
' */
Private Function Hex4ToLong(ByVal Text As String, ByVal StartPos As Long) As Long
    Dim i As Long
    Dim digit As Long
    Dim result As Long
    
    If StartPos < 1 Or StartPos + 3 > Len(Text) Then
        Hex4ToLong = -1
        Exit Function
    End If
    
    For i = 0 To 3
        digit = HexDigitValue(AscW(Mid$(Text, StartPos + i, 1)))
        
        If digit < 0 Then
            Hex4ToLong = -1
            Exit Function
        End If
        
        result = result * 16 + digit
    Next i
    
    Hex4ToLong = result
End Function

'/**
' * @method HexDigitValue
' * @brief Returns the numeric value of a hexadecimal character code.
' * @param CharCode UTF-16 character code.
' * @return Hex digit value, or -1 when invalid.
' */
Private Function HexDigitValue(ByVal CharCode As Long) As Long
    Select Case CharCode
        Case 48 To 57
            HexDigitValue = CharCode - 48
        Case 65 To 70
            HexDigitValue = CharCode - 55
        Case 97 To 102
            HexDigitValue = CharCode - 87
        Case Else
            HexDigitValue = -1
    End Select
End Function

'/**
' * @method CodePointToString
' * @brief Converts a Unicode code point into a VBA UTF-16 string.
' * @param CodePoint Unicode code point.
' * @return UTF-16 string.
' */
Private Function CodePointToString(ByVal CodePoint As Long) As String
    Dim highSurrogate As Long
    Dim lowSurrogate As Long
    Dim temp As Long
    
    If CodePoint < 0 Then Exit Function
    
    If CodePoint <= &HFFFF& Then
        CodePointToString = ChrWCompat(CodePoint)
    ElseIf CodePoint <= &H10FFFF Then
        temp = CodePoint - &H10000
        highSurrogate = &HD800& + (temp \ &H400&)
        lowSurrogate = &HDC00& + (temp And &H3FF&)
        CodePointToString = ChrWCompat(highSurrogate) & ChrWCompat(lowSurrogate)
    End If
End Function

'/**
' * @method ChrWCompat
' * @brief Converts a UTF-16 code unit into a VBA character without failing on values above 32767.
' * @param CodeUnit UTF-16 code unit.
' * @return Single UTF-16 code unit string.
' */
Private Function ChrWCompat(ByVal CodeUnit As Long) As String
    If CodeUnit < 0 Then Exit Function
    
    If CodeUnit <= &H7FFF& Then
        ChrWCompat = ChrW$(CodeUnit)
    ElseIf CodeUnit <= &HFFFF& Then
        ChrWCompat = ChrW$(CodeUnit - &H10000)
    End If
End Function

'/**
' * @method StringifyCurrent
' * @brief Serializes the current document or node to JSON text using a linear-growth writer buffer.
' * @param Pretty Enables formatted output.
' * @param IndentText Text used for each indentation level.
' * @return JSON text.
' */
Friend Function StringifyCurrent(ByVal Pretty As Boolean, ByVal IndentText As String) As String
    Dim doc As JSON
    Dim baseId As Long
    Dim Buffer As String
    Dim Position As Long
    Dim Capacity As Long
    Dim InitialCapacity As Long
    
    ResolveBase doc, baseId
    
    If baseId = 0 Then
        StringifyCurrent = "null"
        Exit Function
    End If
    
    InitialCapacity = doc.EstimateForStringifyCapacity(baseId, Pretty, IndentText)
    WriterStart Buffer, Position, Capacity, InitialCapacity
    doc.WriteToken Buffer, Position, Capacity, baseId, Pretty, IndentText, 0
    StringifyCurrent = WriterFinish(Buffer, Position)
End Function

'/**
' * @method StringifyToken
' * @brief Serializes a parsed token to JSON text.
' * @param TokenId Token id.
' * @param Pretty Enables formatted output.
' * @param IndentText Text used for each indentation level.
' * @param Depth Current nesting depth.
' * @return JSON text.
' */
Friend Function StringifyToken(ByVal TokenId As Long, ByVal Pretty As Boolean, ByVal IndentText As String, ByVal Depth As Long) As String
    Dim Buffer As String
    Dim Position As Long
    Dim Capacity As Long
    
    WriterStart Buffer, Position, Capacity, EstimateForStringifyCapacity(TokenId, Pretty, IndentText)
    WriteToken Buffer, Position, Capacity, TokenId, Pretty, IndentText, Depth
    StringifyToken = WriterFinish(Buffer, Position)
End Function

'/**
' * @method WriteToken
' * @brief Appends a parsed token to the shared writer buffer.
' * @param Buffer Target writer buffer.
' * @param Position Current used character count.
' * @param Capacity Current allocated character capacity.
' * @param TokenId Token id.
' * @param Pretty Enables formatted output.
' * @param IndentText Text used for each indentation level.
' * @param Depth Current nesting depth.
' */
Friend Sub WriteToken(ByRef Buffer As String, ByRef Position As Long, ByRef Capacity As Long, ByVal TokenId As Long, ByVal Pretty As Boolean, ByVal IndentText As String, ByVal Depth As Long)
    If TokenId < 1 Or TokenId > m_TokenCount Then
        WriterAppend Buffer, Position, Capacity, "null"
        Exit Sub
    End If
    
    Select Case m_Tokens(TokenId).Type
        Case jsObject
            WriteObjectToken Buffer, Position, Capacity, TokenId, Pretty, IndentText, Depth
        
        Case jsArray
            WriteArrayToken Buffer, Position, Capacity, TokenId, Pretty, IndentText, Depth
        
        Case jsString
            WriteRawStringValue Buffer, Position, Capacity, TokenId
        
        Case jsNumber
            WriterAppendRawValue Buffer, Position, Capacity, TokenId
        
        Case jsBool
            WriterAppendRawValue Buffer, Position, Capacity, TokenId
        
        Case jsNull
            WriterAppend Buffer, Position, Capacity, "null"
        
        Case Else
            WriterAppend Buffer, Position, Capacity, "null"
    End Select
End Sub

'/**
' * @method WriteObjectToken
' * @brief Appends an object token to the shared writer buffer.
' * @param Buffer Target writer buffer.
' * @param Position Current used character count.
' * @param Capacity Current allocated character capacity.
' * @param TokenId Object token id.
' * @param Pretty Enables formatted output.
' * @param IndentText Text used for each indentation level.
' * @param Depth Current nesting depth.
' */
Private Sub WriteObjectToken(ByRef Buffer As String, ByRef Position As Long, ByRef Capacity As Long, ByVal TokenId As Long, ByVal Pretty As Boolean, ByVal IndentText As String, ByVal Depth As Long)
    Dim child As Long
    Dim first As Boolean
    
    If m_Tokens(TokenId).ChildCount = 0 Then
        WriterAppend Buffer, Position, Capacity, "{}"
        Exit Sub
    End If
    
    WriterAppend Buffer, Position, Capacity, "{"
    first = True
    child = m_Tokens(TokenId).FirstChild
    
    Do While child > 0
        If first Then
            first = False
        Else
            WriterAppend Buffer, Position, Capacity, ","
        End If
        
        If Pretty Then
            WriterAppend Buffer, Position, Capacity, vbCrLf
            WriterAppendIndent Buffer, Position, Capacity, IndentText, Depth + 1
        End If
        
        WriteRawKey Buffer, Position, Capacity, child
        
        If Pretty Then
            WriterAppend Buffer, Position, Capacity, ": "
        Else
            WriterAppend Buffer, Position, Capacity, ":"
        End If
        
        WriteToken Buffer, Position, Capacity, child, Pretty, IndentText, Depth + 1
        child = m_Tokens(child).NextSibling
    Loop
    
    If Pretty Then
        WriterAppend Buffer, Position, Capacity, vbCrLf
        WriterAppendIndent Buffer, Position, Capacity, IndentText, Depth
    End If
    
    WriterAppend Buffer, Position, Capacity, "}"
End Sub

'/**
' * @method WriteArrayToken
' * @brief Appends an array token to the shared writer buffer.
' * @param Buffer Target writer buffer.
' * @param Position Current used character count.
' * @param Capacity Current allocated character capacity.
' * @param TokenId Array token id.
' * @param Pretty Enables formatted output.
' * @param IndentText Text used for each indentation level.
' * @param Depth Current nesting depth.
' */
Private Sub WriteArrayToken(ByRef Buffer As String, ByRef Position As Long, ByRef Capacity As Long, ByVal TokenId As Long, ByVal Pretty As Boolean, ByVal IndentText As String, ByVal Depth As Long)
    Dim child As Long
    Dim first As Boolean
    
    If m_Tokens(TokenId).ChildCount = 0 Then
        WriterAppend Buffer, Position, Capacity, "[]"
        Exit Sub
    End If
    
    WriterAppend Buffer, Position, Capacity, "["
    first = True
    child = m_Tokens(TokenId).FirstChild
    
    Do While child > 0
        If first Then
            first = False
        Else
            WriterAppend Buffer, Position, Capacity, ","
        End If
        
        If Pretty Then
            WriterAppend Buffer, Position, Capacity, vbCrLf
            WriterAppendIndent Buffer, Position, Capacity, IndentText, Depth + 1
        End If
        
        WriteToken Buffer, Position, Capacity, child, Pretty, IndentText, Depth + 1
        child = m_Tokens(child).NextSibling
    Loop
    
    If Pretty Then
        WriterAppend Buffer, Position, Capacity, vbCrLf
        WriterAppendIndent Buffer, Position, Capacity, IndentText, Depth
    End If
    
    WriterAppend Buffer, Position, Capacity, "]"
End Sub

'/**
' * @method StringifyAny
' * @brief Serializes a VBA value to JSON text.
' * @param Value VBA value.
' * @param Pretty Enables formatted output.
' * @param IndentText Text used for each indentation level.
' * @param Depth Current nesting depth.
' * @return JSON text.
' */
Friend Function StringifyAny(ByVal Value As Variant, ByVal Pretty As Boolean, ByVal IndentText As String, ByVal Depth As Long) As String
    Dim Buffer As String
    Dim Position As Long
    Dim Capacity As Long
    
    WriterStart Buffer, Position, Capacity, JSON_VALUE_WRITER_CAPACITY
    WriteAny Buffer, Position, Capacity, Value, Pretty, IndentText, Depth
    StringifyAny = WriterFinish(Buffer, Position)
End Function

'/**
' * @method WriteAny
' * @brief Appends a VBA value to the shared writer buffer.
' * @param Buffer Target writer buffer.
' * @param Position Current used character count.
' * @param Capacity Current allocated character capacity.
' * @param Value VBA value.
' * @param Pretty Enables formatted output.
' * @param IndentText Text used for each indentation level.
' * @param Depth Current nesting depth.
' */
Private Sub WriteAny(ByRef Buffer As String, ByRef Position As Long, ByRef Capacity As Long, ByVal Value As Variant, ByVal Pretty As Boolean, ByVal IndentText As String, ByVal Depth As Long)
    If VBA.IsObject(Value) Then
        Dim objectValue As Object
        Set objectValue = Value
        WriteObjectValue Buffer, Position, Capacity, objectValue, Pretty, IndentText, Depth
        Exit Sub
    End If
    
    If VBA.IsNull(Value) Then
        WriterAppend Buffer, Position, Capacity, "null"
        Exit Sub
    End If
    
    If VBA.IsEmpty(Value) Then
        WriterAppend Buffer, Position, Capacity, "null"
        Exit Sub
    End If
    
    If (VarType(Value) And vbArray) = vbArray Then
        WriteArrayValue Buffer, Position, Capacity, Value, Pretty, IndentText, Depth
        Exit Sub
    End If
    
    Select Case VarType(Value)
        Case vbString
            WriteQuotedJSONString Buffer, Position, Capacity, CStr(Value)
        
        Case vbBoolean
            If CBool(Value) Then
                WriterAppend Buffer, Position, Capacity, "true"
            Else
                WriterAppend Buffer, Position, Capacity, "false"
            End If
        
        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
            WriterAppend Buffer, Position, Capacity, NormalizeNumber(CStr(Value))
        
        Case vbDate
            WriteQuotedJSONString Buffer, Position, Capacity, Format$(CDate(Value), "yyyy-mm-dd\THH:nn:ss")
        
        Case Else
            WriteQuotedJSONString Buffer, Position, Capacity, CStr(Value)
    End Select
End Sub

'/**
' * @method StringifyObjectValue
' * @brief Serializes supported VBA object values to JSON text.
' * @param Value Object value.
' * @param Pretty Enables formatted output.
' * @param IndentText Text used for each indentation level.
' * @param Depth Current nesting depth.
' * @return JSON text.
' */
Private Function StringifyObjectValue(ByVal Value As Object, ByVal Pretty As Boolean, ByVal IndentText As String, ByVal Depth As Long) As String
    Dim Buffer As String
    Dim Position As Long
    Dim Capacity As Long
    
    WriterStart Buffer, Position, Capacity, JSON_VALUE_WRITER_CAPACITY
    WriteObjectValue Buffer, Position, Capacity, Value, Pretty, IndentText, Depth
    StringifyObjectValue = WriterFinish(Buffer, Position)
End Function

'/**
' * @method WriteObjectValue
' * @brief Appends a supported VBA object value to the shared writer buffer.
' * @param Buffer Target writer buffer.
' * @param Position Current used character count.
' * @param Capacity Current allocated character capacity.
' * @param Value Object value.
' * @param Pretty Enables formatted output.
' * @param IndentText Text used for each indentation level.
' * @param Depth Current nesting depth.
' */
Private Sub WriteObjectValue(ByRef Buffer As String, ByRef Position As Long, ByRef Capacity As Long, ByVal Value As Object, ByVal Pretty As Boolean, ByVal IndentText As String, ByVal Depth As Long)
    If Value Is Nothing Then
        WriterAppend Buffer, Position, Capacity, "null"
        Exit Sub
    End If
    
    Select Case TypeName(Value)
        Case "JSON"
            Dim node As JSON
            Set node = Value
            node.WriteCurrent Buffer, Position, Capacity, Pretty, IndentText, Depth
        
        Case "Collection"
            WriteCollection Buffer, Position, Capacity, Value, Pretty, IndentText, Depth
        
        Case "Dictionary", "Scripting.Dictionary"
            WriteDictionary Buffer, Position, Capacity, Value, Pretty, IndentText, Depth
        
        Case Else
            WriterAppend Buffer, Position, Capacity, "null"
    End Select
End Sub

'/**
' * @method WriteCurrent
' * @brief Appends the current document or node to an existing writer buffer.
' * @param Buffer Target writer buffer.
' * @param Position Current used character count.
' * @param Capacity Current allocated character capacity.
' * @param Pretty Enables formatted output.
' * @param IndentText Text used for each indentation level.
' */
Friend Sub WriteCurrent(ByRef Buffer As String, ByRef Position As Long, ByRef Capacity As Long, ByVal Pretty As Boolean, ByVal IndentText As String, Optional ByVal Depth As Long = 0)
    Dim doc As JSON
    Dim baseId As Long
    
    ResolveBase doc, baseId
    
    If baseId = 0 Then
        WriterAppend Buffer, Position, Capacity, "null"
        Exit Sub
    End If
    
    doc.WriteToken Buffer, Position, Capacity, baseId, Pretty, IndentText, Depth
End Sub

'/**
' * @method StringifyArrayValue
' * @brief Serializes a one-dimensional VBA array to JSON text.
' * @param Value VBA array.
' * @param Pretty Enables formatted output.
' * @param IndentText Text used for each indentation level.
' * @param Depth Current nesting depth.
' * @return JSON array text.
' */
Private Function StringifyArrayValue(ByVal Value As Variant, ByVal Pretty As Boolean, ByVal IndentText As String, ByVal Depth As Long) As String
    Dim Buffer As String
    Dim Position As Long
    Dim Capacity As Long
    
    WriterStart Buffer, Position, Capacity, JSON_VALUE_WRITER_CAPACITY
    WriteArrayValue Buffer, Position, Capacity, Value, Pretty, IndentText, Depth
    StringifyArrayValue = WriterFinish(Buffer, Position)
End Function

'/**
' * @method WriteArrayValue
' * @brief Appends a one-dimensional VBA array to the shared writer buffer.
' * @param Buffer Target writer buffer.
' * @param Position Current used character count.
' * @param Capacity Current allocated character capacity.
' * @param Value VBA array.
' * @param Pretty Enables formatted output.
' * @param IndentText Text used for each indentation level.
' * @param Depth Current nesting depth.
' */
Private Sub WriteArrayValue(ByRef Buffer As String, ByRef Position As Long, ByRef Capacity As Long, ByVal Value As Variant, ByVal Pretty As Boolean, ByVal IndentText As String, ByVal Depth As Long)
    Dim i As Long
    Dim first As Boolean
    
    WriterAppend Buffer, Position, Capacity, "["
    first = True
    
    For i = LBound(Value) To UBound(Value)
        If first Then
            first = False
        Else
            WriterAppend Buffer, Position, Capacity, ","
        End If
        
        If Pretty Then
            WriterAppend Buffer, Position, Capacity, vbCrLf
            WriterAppendIndent Buffer, Position, Capacity, IndentText, Depth + 1
        End If
        
        WriteAny Buffer, Position, Capacity, Value(i), Pretty, IndentText, Depth + 1
    Next i
    
    If Pretty And Not first Then
        WriterAppend Buffer, Position, Capacity, vbCrLf
        WriterAppendIndent Buffer, Position, Capacity, IndentText, Depth
    End If
    
    WriterAppend Buffer, Position, Capacity, "]"
End Sub

'/**
' * @method StringifyCollection
' * @brief Serializes a VBA Collection as a JSON array.
' * @param Value Collection object.
' * @param Pretty Enables formatted output.
' * @param IndentText Text used for each indentation level.
' * @param Depth Current nesting depth.
' * @return JSON array text.
' */
Private Function StringifyCollection(ByVal Value As Object, ByVal Pretty As Boolean, ByVal IndentText As String, ByVal Depth As Long) As String
    Dim Buffer As String
    Dim Position As Long
    Dim Capacity As Long
    
    WriterStart Buffer, Position, Capacity, JSON_VALUE_WRITER_CAPACITY
    WriteCollection Buffer, Position, Capacity, Value, Pretty, IndentText, Depth
    StringifyCollection = WriterFinish(Buffer, Position)
End Function

'/**
' * @method WriteCollection
' * @brief Appends a VBA Collection as a JSON array to the shared writer buffer.
' * @param Buffer Target writer buffer.
' * @param Position Current used character count.
' * @param Capacity Current allocated character capacity.
' * @param Value Collection object.
' * @param Pretty Enables formatted output.
' * @param IndentText Text used for each indentation level.
' * @param Depth Current nesting depth.
' */
Private Sub WriteCollection(ByRef Buffer As String, ByRef Position As Long, ByRef Capacity As Long, ByVal Value As Object, ByVal Pretty As Boolean, ByVal IndentText As String, ByVal Depth As Long)
    Dim i As Long
    
    If Value.Count = 0 Then
        WriterAppend Buffer, Position, Capacity, "[]"
        Exit Sub
    End If
    
    WriterAppend Buffer, Position, Capacity, "["
    
    For i = 1 To Value.Count
        If i > 1 Then WriterAppend Buffer, Position, Capacity, ","
        
        If Pretty Then
            WriterAppend Buffer, Position, Capacity, vbCrLf
            WriterAppendIndent Buffer, Position, Capacity, IndentText, Depth + 1
        End If
        
        WriteAny Buffer, Position, Capacity, Value.Item(i), Pretty, IndentText, Depth + 1
    Next i
    
    If Pretty Then
        WriterAppend Buffer, Position, Capacity, vbCrLf
        WriterAppendIndent Buffer, Position, Capacity, IndentText, Depth
    End If
    
    WriterAppend Buffer, Position, Capacity, "]"
End Sub

'/**
' * @method StringifyDictionary
' * @brief Serializes a Scripting.Dictionary as a JSON object.
' * @param Value Dictionary object.
' * @param Pretty Enables formatted output.
' * @param IndentText Text used for each indentation level.
' * @param Depth Current nesting depth.
' * @return JSON object text.
' */
Private Function StringifyDictionary(ByVal Value As Object, ByVal Pretty As Boolean, ByVal IndentText As String, ByVal Depth As Long) As String
    Dim Buffer As String
    Dim Position As Long
    Dim Capacity As Long
    
    WriterStart Buffer, Position, Capacity, JSON_VALUE_WRITER_CAPACITY
    WriteDictionary Buffer, Position, Capacity, Value, Pretty, IndentText, Depth
    StringifyDictionary = WriterFinish(Buffer, Position)
End Function

'/**
' * @method WriteDictionary
' * @brief Appends a Scripting.Dictionary as a JSON object to the shared writer buffer.
' * @param Buffer Target writer buffer.
' * @param Position Current used character count.
' * @param Capacity Current allocated character capacity.
' * @param Value Dictionary object.
' * @param Pretty Enables formatted output.
' * @param IndentText Text used for each indentation level.
' * @param Depth Current nesting depth.
' */
Private Sub WriteDictionary(ByRef Buffer As String, ByRef Position As Long, ByRef Capacity As Long, ByVal Value As Object, ByVal Pretty As Boolean, ByVal IndentText As String, ByVal Depth As Long)
    Dim keys As Variant
    Dim key As Variant
    Dim i As Long
    
    If Value.Count = 0 Then
        WriterAppend Buffer, Position, Capacity, "{}"
        Exit Sub
    End If
    
    keys = Value.keys
    WriterAppend Buffer, Position, Capacity, "{"
    
    For i = LBound(keys) To UBound(keys)
        If i > LBound(keys) Then WriterAppend Buffer, Position, Capacity, ","
        
        If Pretty Then
            WriterAppend Buffer, Position, Capacity, vbCrLf
            WriterAppendIndent Buffer, Position, Capacity, IndentText, Depth + 1
        End If
        
        key = keys(i)
        WriteQuotedJSONString Buffer, Position, Capacity, CStr(key)
        
        If Pretty Then
            WriterAppend Buffer, Position, Capacity, ": "
        Else
            WriterAppend Buffer, Position, Capacity, ":"
        End If
        
        WriteAny Buffer, Position, Capacity, Value.Item(key), Pretty, IndentText, Depth + 1
    Next i
    
    If Pretty Then
        WriterAppend Buffer, Position, Capacity, vbCrLf
        WriterAppendIndent Buffer, Position, Capacity, IndentText, Depth
    End If
    
    WriterAppend Buffer, Position, Capacity, "}"
End Sub

'/**
' * @method RepeatText
' * @brief Repeats an indentation string a fixed number of times.
' * @param Text Text to repeat.
' * @param Count Number of repetitions.
' * @return Repeated text.
' */
Private Function RepeatText(ByVal Text As String, ByVal Count As Long) As String
    Dim Buffer As String
    Dim Position As Long
    Dim Capacity As Long
    
    If Count <= 0 Then Exit Function
    WriterStart Buffer, Position, Capacity, Len(Text) * Count
    WriterAppendIndent Buffer, Position, Capacity, Text, Count
    RepeatText = WriterFinish(Buffer, Position)
End Function

'/**
' * @method QuoteJSONString
' * @brief Escapes and quotes a string as JSON string text.
' * @param Text Raw string text.
' * @return Quoted JSON string.
' */
Private Function QuoteJSONString(ByVal Text As String) As String
    Dim Buffer As String
    Dim Position As Long
    Dim Capacity As Long
    
    WriterStart Buffer, Position, Capacity, Len(Text) + 2
    WriteQuotedJSONString Buffer, Position, Capacity, Text
    QuoteJSONString = WriterFinish(Buffer, Position)
End Function

'/**
' * @method EstimateForStringifyCapacity
' * @brief Estimates the initial writer capacity for a parsed token.
' * @param TokenId Token id.
' * @param Pretty Enables formatted output.
' * @param IndentText Text used for each indentation level.
' * @return Estimated character capacity.
' */
Friend Function EstimateForStringifyCapacity(ByVal TokenId As Long, ByVal Pretty As Boolean, ByVal IndentText As String) As Long
    Dim estimate As Long
    
    estimate = EstimateTokenCapacity(TokenId, Pretty, Len(IndentText), 0) + 32
    
    If estimate < JSON_WRITER_MIN_CAPACITY Then estimate = JSON_WRITER_MIN_CAPACITY
    EstimateForStringifyCapacity = estimate
End Function

'/**
' * @method EstimateTokenCapacity
' * @brief Estimates writer capacity for one parsed token subtree.
' * @param TokenId Token id.
' * @param Pretty Enables formatted output.
' * @param IndentLen Indentation text length.
' * @param Depth Current nesting depth.
' * @return Estimated character count.
' */
Private Function EstimateTokenCapacity(ByVal TokenId As Long, ByVal Pretty As Boolean, ByVal IndentLen As Long, ByVal Depth As Long) As Long
    Dim child As Long
    Dim total As Long
    
    If TokenId < 1 Or TokenId > m_TokenCount Then
        EstimateTokenCapacity = 4
        Exit Function
    End If
    
    Select Case m_Tokens(TokenId).Type
        Case jsObject
            If m_Tokens(TokenId).ChildCount = 0 Then
                EstimateTokenCapacity = 2
                Exit Function
            End If
            
            total = 2
            child = m_Tokens(TokenId).FirstChild
            
            Do While child > 0
                If child <> m_Tokens(TokenId).FirstChild Then total = total + 1
                total = total + m_Tokens(child).KeyLen + 3
                total = total + EstimateTokenCapacity(child, Pretty, IndentLen, Depth + 1)
                
                If Pretty Then total = total + 2 + (IndentLen * (Depth + 1)) + 1
                child = m_Tokens(child).NextSibling
            Loop
            
            If Pretty Then total = total + 2 + (IndentLen * Depth)
            EstimateTokenCapacity = total
        
        Case jsArray
            If m_Tokens(TokenId).ChildCount = 0 Then
                EstimateTokenCapacity = 2
                Exit Function
            End If
            
            total = 2
            child = m_Tokens(TokenId).FirstChild
            
            Do While child > 0
                If child <> m_Tokens(TokenId).FirstChild Then total = total + 1
                total = total + EstimateTokenCapacity(child, Pretty, IndentLen, Depth + 1)
                
                If Pretty Then total = total + 2 + (IndentLen * (Depth + 1))
                child = m_Tokens(child).NextSibling
            Loop
            
            If Pretty Then total = total + 2 + (IndentLen * Depth)
            EstimateTokenCapacity = total
        
        Case jsString
            EstimateTokenCapacity = m_Tokens(TokenId).ValLen + 2
        
        Case jsNumber, jsBool
            EstimateTokenCapacity = m_Tokens(TokenId).ValLen
        
        Case jsNull
            EstimateTokenCapacity = 4
        
        Case Else
            EstimateTokenCapacity = 4
    End Select
End Function

'/**
' * @method WriterStart
' * @brief Initializes a string writer buffer.
' * @param Buffer Target writer buffer.
' * @param Position Current used character count.
' * @param Capacity Current allocated character capacity.
' * @param InitialCapacity Initial character capacity.
' */
Private Sub WriterStart(ByRef Buffer As String, ByRef Position As Long, ByRef Capacity As Long, ByVal InitialCapacity As Long)
    Capacity = InitialCapacity
    If Capacity < JSON_WRITER_MIN_CAPACITY Then Capacity = JSON_WRITER_MIN_CAPACITY
    Buffer = String$(Capacity, vbNullChar)
    Position = 0
End Sub

'/**
' * @method WriterFinish
' * @brief Returns the used part of a string writer buffer.
' * @param Buffer Source writer buffer.
' * @param Position Current used character count.
' * @return Written JSON text.
' */
Private Function WriterFinish(ByRef Buffer As String, ByVal Position As Long) As String
    If Position <= 0 Then Exit Function
    WriterFinish = Left$(Buffer, Position)
End Function

'/**
' * @method WriterAppend
' * @brief Appends text to a string writer buffer.
' * @param Buffer Target writer buffer.
' * @param Position Current used character count.
' * @param Capacity Current allocated character capacity.
' * @param Text Text to append.
' */
Private Sub WriterAppend(ByRef Buffer As String, ByRef Position As Long, ByRef Capacity As Long, ByVal Text As String)
    Dim length As Long
    
    length = Len(Text)
    If length = 0 Then Exit Sub
    
    WriterEnsure Buffer, Position, Capacity, length
    Mid$(Buffer, Position + 1, length) = Text
    Position = Position + length
End Sub

'/**
' * @method WriterAppendRawValue
' * @brief Appends a parsed value slice without creating a temporary normalized value.
' * @param Buffer Target writer buffer.
' * @param Position Current used character count.
' * @param Capacity Current allocated character capacity.
' * @param TokenId Token id.
' */
Private Sub WriterAppendRawValue(ByRef Buffer As String, ByRef Position As Long, ByRef Capacity As Long, ByVal TokenId As Long)
    If TokenId < 1 Or TokenId > m_TokenCount Then Exit Sub
    If m_Tokens(TokenId).ValLen <= 0 Then Exit Sub
    WriterAppend Buffer, Position, Capacity, Mid$(m_Text, m_Tokens(TokenId).ValStart, m_Tokens(TokenId).ValLen)
End Sub

'/**
' * @method WriterAppendIndent
' * @brief Appends indentation text Count times to a string writer buffer.
' * @param Buffer Target writer buffer.
' * @param Position Current used character count.
' * @param Capacity Current allocated character capacity.
' * @param Text Indentation text.
' * @param Count Number of repetitions.
' */
Private Sub WriterAppendIndent(ByRef Buffer As String, ByRef Position As Long, ByRef Capacity As Long, ByVal Text As String, ByVal Count As Long)
    Dim i As Long
    Dim length As Long
    
    If Count <= 0 Then Exit Sub
    length = Len(Text)
    If length = 0 Then Exit Sub
    
    WriterEnsure Buffer, Position, Capacity, length * Count
    
    For i = 1 To Count
        Mid$(Buffer, Position + 1, length) = Text
        Position = Position + length
    Next i
End Sub

'/**
' * @method WriterEnsure
' * @brief Ensures the string writer buffer has enough free capacity.
' * @param Buffer Target writer buffer.
' * @param Position Current used character count.
' * @param Capacity Current allocated character capacity.
' * @param Additional Additional required character count.
' */
Private Sub WriterEnsure(ByRef Buffer As String, ByVal Position As Long, ByRef Capacity As Long, ByVal Additional As Long)
    Dim required As Long
    Dim newCapacity As Long
    
    required = Position + Additional
    If required <= Capacity Then Exit Sub
    
    newCapacity = Capacity
    If newCapacity < JSON_WRITER_MIN_CAPACITY Then newCapacity = JSON_WRITER_MIN_CAPACITY
    
    Do While newCapacity < required
        If newCapacity < 1048576 Then
            newCapacity = newCapacity * 2
        Else
            newCapacity = newCapacity + (newCapacity \ 2)
        End If
    Loop
    
    Buffer = Left$(Buffer, Position) & String$(newCapacity - Position, vbNullChar)
    Capacity = newCapacity
End Sub

'/**
' * @method WriteRawKey
' * @brief Appends a parsed object key using the original already-escaped source slice.
' * @param Buffer Target writer buffer.
' * @param Position Current used character count.
' * @param Capacity Current allocated character capacity.
' * @param TokenId Token id.
' */
Private Sub WriteRawKey(ByRef Buffer As String, ByRef Position As Long, ByRef Capacity As Long, ByVal TokenId As Long)
    WriterAppend Buffer, Position, Capacity, """"
    If TokenId >= 1 And TokenId <= m_TokenCount Then
        If m_Tokens(TokenId).KeyLen > 0 Then
            WriterAppend Buffer, Position, Capacity, Mid$(m_Text, m_Tokens(TokenId).KeyStart, m_Tokens(TokenId).KeyLen)
        End If
    End If
    WriterAppend Buffer, Position, Capacity, """"
End Sub

'/**
' * @method WriteRawStringValue
' * @brief Appends a parsed string value using the original already-escaped source slice.
' * @param Buffer Target writer buffer.
' * @param Position Current used character count.
' * @param Capacity Current allocated character capacity.
' * @param TokenId Token id.
' */
Private Sub WriteRawStringValue(ByRef Buffer As String, ByRef Position As Long, ByRef Capacity As Long, ByVal TokenId As Long)
    WriterAppend Buffer, Position, Capacity, """"
    If TokenId >= 1 And TokenId <= m_TokenCount Then
        If m_Tokens(TokenId).ValLen > 0 Then
            WriterAppend Buffer, Position, Capacity, Mid$(m_Text, m_Tokens(TokenId).ValStart, m_Tokens(TokenId).ValLen)
        End If
    End If
    WriterAppend Buffer, Position, Capacity, """"
End Sub

'/**
' * @method WriteQuotedJSONString
' * @brief Escapes and appends a raw VBA string as JSON string text.
' * @param Buffer Target writer buffer.
' * @param Position Current used character count.
' * @param Capacity Current allocated character capacity.
' * @param Text Raw string text.
' */
Private Sub WriteQuotedJSONString(ByRef Buffer As String, ByRef Position As Long, ByRef Capacity As Long, ByVal Text As String)
    Dim length As Long
    Dim i As Long
    Dim segmentStart As Long
    Dim code As Long
    Dim escaped As String
    
    WriterAppend Buffer, Position, Capacity, """"
    length = Len(Text)
    
    If length > 0 Then
        segmentStart = 1
        
        For i = 1 To length
            code = AscW(Mid$(Text, i, 1))
            escaped = vbNullString
            
            Select Case code
                Case 34
                    escaped = "\"""
                Case 92
                    escaped = "\\"
                Case 8
                    escaped = "\b"
                Case 9
                    escaped = "\t"
                Case 10
                    escaped = "\n"
                Case 12
                    escaped = "\f"
                Case 13
                    escaped = "\r"
                Case 0 To 31
                    escaped = "\u" & Right$("000" & Hex$(code), 4)
            End Select
            
            If Len(escaped) > 0 Then
                If i > segmentStart Then
                    WriterAppend Buffer, Position, Capacity, Mid$(Text, segmentStart, i - segmentStart)
                End If
                
                WriterAppend Buffer, Position, Capacity, escaped
                segmentStart = i + 1
            End If
        Next i
        
        If segmentStart <= length Then
            WriterAppend Buffer, Position, Capacity, Mid$(Text, segmentStart, length - segmentStart + 1)
        End If
    End If
    
    WriterAppend Buffer, Position, Capacity, """"
End Sub

'/**
' * @method NormalizeNumber
' * @brief Normalizes a VBA numeric string to JSON decimal format.
' * @param Text Numeric text.
' * @return JSON-compatible number text.
' */
Private Function NormalizeNumber(ByVal Text As String) As String
    NormalizeNumber = Replace(Text, ",", ".")
End Function
