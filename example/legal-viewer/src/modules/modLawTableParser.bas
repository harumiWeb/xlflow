Attribute VB_Name = "modLawTableParser"
Option Explicit

Public Function FormattedTablesTextFromNode(ByVal rootNode As Object) As String
    On Error GoTo ErrHandler

    Dim tableBlocks As Collection
    Set tableBlocks = New Collection
    CollectTableBlocks rootNode, tableBlocks

    Dim result As String
    Dim blockIndex As Long
    Dim block As Object
    Dim tableValues As Variant
    For blockIndex = 1 To tableBlocks.Count
        Set block = tableBlocks(blockIndex)
        tableValues = block("values")

        If tableBlocks.Count > 1 Then
            result = AppendLine(result, "表" & CStr(blockIndex))
        End If
        result = AppendLine(result, "Markdownテーブル")
        result = AppendLine(result, modTableFormatter.MarkdownTableFromArray(tableValues, True))
        result = AppendLine(result, "ASCII罫線テーブル")
        result = AppendLine(result, modTableFormatter.AsciiTableFromArray(tableValues, True))
    Next blockIndex

    FormattedTablesTextFromNode = result
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "LawTableParser", "modLawTableParser.FormattedTablesTextFromNode", Err.description, "", "", NodeTagSafe(rootNode)
    Err.Raise Err.Number, Err.source, Err.description
End Function

Public Function LawTableParserSmoke() As String
    On Error GoTo ErrHandler

    Dim tableNode As Object
    Set tableNode = MakeTableSmokeNode()

    Dim formattedText As String
    formattedText = FormattedTablesTextFromNode(tableNode)
    If InStr(1, formattedText, "| 項目 | 値 |", vbBinaryCompare) = 0 _
        Or InStr(1, formattedText, "| 労働時間 | 8 |", vbBinaryCompare) = 0 _
        Or InStr(1, formattedText, "| 別表検索語 |  |", vbBinaryCompare) = 0 Then
        Err.Raise vbObjectError + 7111, "modLawTableParser.LawTableParserSmoke", "Markdown table extraction mismatch."
    End If

    If InStr(1, formattedText, "ASCII罫線テーブル", vbBinaryCompare) = 0 _
        Or InStr(1, formattedText, "+------------+----+", vbBinaryCompare) = 0 _
        Or InStr(1, formattedText, "別表検索語", vbBinaryCompare) = 0 Then
        Err.Raise vbObjectError + 7112, "modLawTableParser.LawTableParserSmoke", "ASCII table extraction mismatch."
    End If

    LawTableParserSmoke = "ok:" & CStr(Len(formattedText))
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "LawTableParser", "modLawTableParser.LawTableParserSmoke", Err.description, "", "", ""
    LawTableParserSmoke = "error:" & Err.description
End Function

Private Sub CollectTableBlocks(ByVal node As Object, ByVal tableBlocks As Collection)
    If node Is Nothing Then
        Exit Sub
    End If

    Dim tagName As String
    tagName = NodeTag(node)
    If IsTableContainerTag(tagName) Then
        Dim block As Object
        Set block = TableBlockFromNode(node)
        If Not block Is Nothing Then
            tableBlocks.Add block
            Exit Sub
        End If
    End If

    Dim children As Object
    Set children = NodeChildren(node)
    If children Is Nothing Then
        Exit Sub
    End If

    Dim child As Variant
    For Each child In children
        If IsObject(child) Then
            CollectTableBlocks child, tableBlocks
        End If
    Next child
End Sub

Private Function TableBlockFromNode(ByVal tableNode As Object) As Object
    Dim rows As Collection
    Set rows = New Collection
    CollectTableRows tableNode, rows
    If rows.Count = 0 Then
        Exit Function
    End If

    Dim values As Variant
    values = TableRowsToArray(rows)
    If Not IsArray(values) Then
        Exit Function
    End If

    Dim block As Object
    Set block = CreateObject("Scripting.Dictionary")
    block.Add "values", values
    Set TableBlockFromNode = block
End Function

Private Sub CollectTableRows(ByVal node As Object, ByVal rows As Collection)
    If node Is Nothing Then
        Exit Sub
    End If

    If IsTableRowTag(NodeTag(node)) Then
        rows.Add node
        Exit Sub
    End If

    Dim children As Object
    Set children = NodeChildren(node)
    If children Is Nothing Then
        Exit Sub
    End If

    Dim child As Variant
    For Each child In children
        If IsObject(child) Then
            CollectTableRows child, rows
        End If
    Next child
End Sub

Private Function TableRowsToArray(ByVal rows As Collection) As Variant
    Dim gridRows As Collection
    Set gridRows = New Collection

    Dim maxColumn As Long
    Dim rowIndex As Long
    For rowIndex = 1 To rows.Count
        PlaceTableRow rows(rowIndex), gridRows, rowIndex, maxColumn
    Next rowIndex

    If gridRows.Count = 0 Or maxColumn = 0 Then
        Exit Function
    End If

    Dim result As Variant
    ReDim result(1 To gridRows.Count, 1 To maxColumn)

    Dim columnIndex As Long
    Dim rowValues As Object
    For rowIndex = 1 To gridRows.Count
        Set rowValues = gridRows(rowIndex)
        For columnIndex = 1 To maxColumn
            If rowValues.Exists(CStr(columnIndex)) Then
                result(rowIndex, columnIndex) = rowValues(CStr(columnIndex))
            Else
                result(rowIndex, columnIndex) = ""
            End If
        Next columnIndex
    Next rowIndex

    TableRowsToArray = result
End Function

Private Sub PlaceTableRow(ByVal rowNode As Object, ByVal gridRows As Collection, ByVal rowIndex As Long, ByRef maxColumn As Long)
    Dim rowValues As Object
    Set rowValues = EnsureGridRow(gridRows, rowIndex)

    Dim cells As Collection
    Set cells = New Collection
    CollectTableCells rowNode, cells

    Dim columnIndex As Long
    columnIndex = 1

    Dim cell As Variant
    Dim cellTextValue As String
    Dim rowSpan As Long
    Dim colSpan As Long
    For Each cell In cells
        Do While rowValues.Exists(CStr(columnIndex))
            columnIndex = columnIndex + 1
        Loop

        cellTextValue = NormalizeText(NodeText(cell))
        rowSpan = AttributeLong(cell, Array("rowspan", "rowSpan", "RowSpan"), 1)
        colSpan = AttributeLong(cell, Array("colspan", "colSpan", "ColSpan"), 1)
        If rowSpan < 1 Then rowSpan = 1
        If colSpan < 1 Then colSpan = 1

        SetGridSpan gridRows, rowIndex, columnIndex, rowSpan, colSpan, cellTextValue
        If columnIndex + colSpan - 1 > maxColumn Then
            maxColumn = columnIndex + colSpan - 1
        End If
        columnIndex = columnIndex + colSpan
    Next cell
End Sub

Private Sub CollectTableCells(ByVal node As Object, ByVal cells As Collection)
    If node Is Nothing Then
        Exit Sub
    End If

    If IsTableCellTag(NodeTag(node)) Then
        cells.Add node
        Exit Sub
    End If

    Dim children As Object
    Set children = NodeChildren(node)
    If children Is Nothing Then
        Exit Sub
    End If

    Dim child As Variant
    For Each child In children
        If IsObject(child) Then
            CollectTableCells child, cells
        End If
    Next child
End Sub

Private Sub SetGridSpan(ByVal gridRows As Collection, ByVal startRow As Long, ByVal startColumn As Long, ByVal rowSpan As Long, ByVal colSpan As Long, ByVal cellTextValue As String)
    Dim rowOffset As Long
    Dim columnOffset As Long
    Dim rowValues As Object
    For rowOffset = 0 To rowSpan - 1
        Set rowValues = EnsureGridRow(gridRows, startRow + rowOffset)
        For columnOffset = 0 To colSpan - 1
            If rowOffset = 0 And columnOffset = 0 Then
                rowValues(CStr(startColumn + columnOffset)) = cellTextValue
            ElseIf Not rowValues.Exists(CStr(startColumn + columnOffset)) Then
                rowValues(CStr(startColumn + columnOffset)) = ""
            End If
        Next columnOffset
    Next rowOffset
End Sub

Private Function EnsureGridRow(ByVal gridRows As Collection, ByVal rowIndex As Long) As Object
    Dim rowValues As Object
    Do While gridRows.Count < rowIndex
        Set rowValues = CreateObject("Scripting.Dictionary")
        gridRows.Add rowValues
    Loop

    Set EnsureGridRow = gridRows(rowIndex)
End Function

Private Function AttributeLong(ByVal node As Object, ByVal attributeNames As Variant, ByVal defaultValue As Long) As Long
    Dim attributes As Object
    Set attributes = modJsonUtil.JsonObjectProperty(node, "attr")
    If attributes Is Nothing Then
        AttributeLong = defaultValue
        Exit Function
    End If

    Dim index As Long
    Dim valueText As String
    For index = LBound(attributeNames) To UBound(attributeNames)
        valueText = modJsonUtil.JsonTextProperty(attributes, CStr(attributeNames(index)))
        If Len(valueText) > 0 And IsNumeric(valueText) Then
            AttributeLong = CLng(valueText)
            Exit Function
        End If
    Next index

    AttributeLong = defaultValue
End Function

Private Function NodeTag(ByVal node As Object) As String
    NodeTag = modJsonUtil.JsonTextProperty(node, "tag")
End Function

Private Function NodeTagSafe(ByVal node As Object) As String
    On Error GoTo ErrHandler
    If node Is Nothing Then
        NodeTagSafe = ""
    Else
        NodeTagSafe = NodeTag(node)
    End If
    Exit Function

ErrHandler:
    NodeTagSafe = ""
End Function

Private Function NodeChildren(ByVal node As Object) As Object
    Set NodeChildren = modJsonUtil.JsonObjectProperty(node, "children")
End Function

Private Function NodeText(ByVal node As Variant) As String
    If IsObject(node) Then
        Dim children As Object
        Set children = NodeChildren(node)
        If children Is Nothing Then
            NodeText = ""
            Exit Function
        End If

        Dim result As String
        Dim child As Variant
        For Each child In children
            result = result & NodeText(child)
        Next child
        NodeText = result
    ElseIf IsNull(node) Or IsEmpty(node) Then
        NodeText = ""
    Else
        NodeText = CStr(node)
    End If
End Function

Private Function NormalizeText(ByVal value As String) As String
    value = Replace(value, vbCr, " ")
    value = Replace(value, vbLf, " ")
    value = Replace(value, vbTab, " ")
    value = Trim$(value)

    Do While InStr(1, value, "  ", vbBinaryCompare) > 0
        value = Replace(value, "  ", " ")
    Loop

    NormalizeText = value
End Function

Private Function IsTableContainerTag(ByVal tagName As String) As Boolean
    IsTableContainerTag = (StrComp(tagName, "TableStruct", vbBinaryCompare) = 0 _
        Or StrComp(tagName, "Table", vbBinaryCompare) = 0)
End Function

Private Function IsTableRowTag(ByVal tagName As String) As Boolean
    IsTableRowTag = (StrComp(tagName, "TableRow", vbBinaryCompare) = 0 _
        Or StrComp(tagName, "TableHeaderRow", vbBinaryCompare) = 0)
End Function

Private Function IsTableCellTag(ByVal tagName As String) As Boolean
    IsTableCellTag = (StrComp(tagName, "TableColumn", vbBinaryCompare) = 0 _
        Or StrComp(tagName, "TableHeaderColumn", vbBinaryCompare) = 0)
End Function

Private Function AppendLine(ByVal baseText As String, ByVal partText As String) As String
    If Len(partText) = 0 Then
        AppendLine = baseText
    ElseIf Len(baseText) = 0 Then
        AppendLine = partText
    Else
        AppendLine = baseText & vbCrLf & partText
    End If
End Function

Private Function MakeJsonNode(ByVal tagName As String, Optional ByVal textValue As String = "") As Object
    Dim node As Object
    Set node = CreateObject("Scripting.Dictionary")

    Dim children As Collection
    Set children = New Collection

    node.Add "tag", tagName
    node.Add "children", children
    If Len(textValue) > 0 Then
        children.Add textValue
    End If

    Set MakeJsonNode = node
End Function

Private Sub AddJsonChild(ByVal parentNode As Object, ByVal childNode As Object)
    Dim children As Object
    Set children = NodeChildren(parentNode)
    children.Add childNode
End Sub

Private Sub SetJsonAttr(ByVal node As Object, ByVal attrName As String, ByVal attrValue As String)
    Dim attributes As Object
    Set attributes = modJsonUtil.JsonObjectProperty(node, "attr")
    If attributes Is Nothing Then
        Set attributes = CreateObject("Scripting.Dictionary")
        node.Add "attr", attributes
    End If
    attributes(attrName) = attrValue
End Sub

Private Function MakeTableSmokeNode() As Object
    Dim tableStruct As Object
    Set tableStruct = MakeJsonNode("TableStruct")

    Dim headerRow As Object
    Set headerRow = MakeJsonNode("TableHeaderRow")
    AddJsonChild headerRow, MakeJsonNode("TableHeaderColumn", "項目")
    AddJsonChild headerRow, MakeJsonNode("TableHeaderColumn", "値")
    AddJsonChild tableStruct, headerRow

    Dim bodyRow As Object
    Set bodyRow = MakeJsonNode("TableRow")
    AddJsonChild bodyRow, MakeJsonNode("TableColumn", "労働時間")
    AddJsonChild bodyRow, MakeJsonNode("TableColumn", "8")
    AddJsonChild tableStruct, bodyRow

    Dim spanRow As Object
    Set spanRow = MakeJsonNode("TableRow")
    Dim spanCell As Object
    Set spanCell = MakeJsonNode("TableColumn", "別表検索語")
    SetJsonAttr spanCell, "colspan", "2"
    AddJsonChild spanRow, spanCell
    AddJsonChild tableStruct, spanRow

    Set MakeTableSmokeNode = tableStruct
End Function
