Attribute VB_Name = "AppSnippet"
' Module to handle reading snippets

Public Const SNIPPET_PREFIX_KEY As String = "snippetPrefix"

Public Function GetSnippets() As Variant
    Dim AppProperties As Variant
    AppProperties = AppProperty.GetProperties()
    
    Dim Sheets As Variant, Snippets As Variant, Index As Long, SheetSnippets As Variant, Sheet As Worksheet
    Sheets = Util.GetSheetsByNameGlob(GetSnippetPrefixProperty(AppProperties) & "*")
    
    Snippets = Array()
    For Index = 0 To UBound(Sheets)
        Set Sheet = Sheets(Index)
        
        SheetSnippets = GetSheetSnippets(Sheet)
        
        If DataType.IsNotNil(SheetSnippets) Then
            Snippets = ArrayUtil.JoinArrays(Snippets, SheetSnippets)
        End If
    Next
    
    GetSnippets = Snippets
End Function

Public Function GetSheetSnippets(Sheet As Worksheet) As Variant
    GetSheetSnippets = GetRangeSnippets(Sheet.UsedRange)
End Function

Public Function GetRangeSnippets(Rng As Range)
    Dim AppProperties As Variant, SheetProperties As Variant
    AppProperties = AppProperty.GetProperties()
    SheetProperties = PropertyType.GetSheetProperties(Rng.Worksheet)

    Dim BlockStart As String, BlockEnd As String, Snippets As Variant, Index As Long, Row As Range
    BlockStart = AppProperty.GetBlockStartProperty(AppProperties)
    BlockEnd = AppProperty.GetBlockEndProperty(AppProperties)
    
    Snippets = ArrayUtil.CreateWithSize(Rng.Rows.CountLarge)
    
    Dim RowValue As Variant, StartRow As Range
    Set StartRow = Nothing
    
    ' Simple parser, please do this functionally if possible
    For Each Row In Rng.Rows
        RowValue = Row.Columns(1).Value
        
        If (StartRow Is Nothing) And (RowValue = BlockStart) Then
            Set StartRow = Row
        ElseIf (Not StartRow Is Nothing) And (RowValue = BlockEnd) Then
            If StartRow.Address <> Row.Address Then
                Dim Snippet As Variant
                Snippet = ParseSnippetStartEndRow(StartRow, Row, SheetProperties)
                
                Snippets(Index) = Snippet
                Index = Index + 1
            End If
                        
            Set StartRow = Nothing
        End If
    Next
    
    If Index > 0 Then
        ReDim Preserve Snippets(0 To Index - 1)
    Else
        Snippets = Array()
    End If

    GetRangeSnippets = Snippets
End Function

Private Function ParseSnippetStartEndRow(StartRow As Range, EndRow As Range, Properties As Variant) As Variant
    Dim Sheet As Worksheet, SnippetRange As Range, SnippetStartRow As Range, SnippetEndRow As Range
    Set Sheet = StartRow.Worksheet
    Set SnippetStartRow = StartRow.Offset(1)
    Set SnippetEndRow = EndRow.Offset(-1)
    Set SnippetRange = Sheet.Range(SnippetStartRow, SnippetEndRow)
    
    Dim SnippetProperties As Variant
    SnippetProperties = PropertyType.GetRangeProperties(SnippetRange)
    SnippetProperties = ObjectType.Merge(SnippetProperties, Properties)
    
    Dim BlockRange As Range, BlockStartRow As Range
    For Each BlockStartRow In SnippetRange.Rows
        If Not PropertyType.IsPropertyRow(BlockStartRow) Then
            Set BlockRange = Sheet.Range(BlockStartRow, SnippetEndRow)
            ' Bad practice but better than the alternative
            Exit For
        End If
    Next
    
    ParseSnippetStartEndRow = SnippetType.Create(SnippetProperties, BlockRange)
End Function

Private Function GetSnippetSheets(Properties As Variant) As String
    GetSnippetSheets = Util.GetSheetsByNameGlob(GetSnippetPrefixProperty(Properties) & "*")
End Function

Private Function GetSnippetPrefixProperty(Properties As Variant) As String
    GetSnippetPrefixProperty = ObjectType.Take(SNIPPET_PREFIX_KEY, "PRT", Properties)
End Function

