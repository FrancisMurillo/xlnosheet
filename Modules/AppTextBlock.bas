Attribute VB_Name = "AppTextBlock"
' A module for the text blocks

Public Const TEXT_PREFIX_KEY As String = "textPrefix"

Public Sub a()
    Dim CurSheet As Worksheet
    Set CurSheet = ActiveSheet
    
    Dim SampleSheet As Worksheet
    Set SampleSheet = Worksheets("Sandbox")
    
    BuildSourceToTarget AppSnippet.GetSnippets, CurSheet, SampleSheet
    
    'Util.DeleteSheet SampleSheet
End Sub

Public Function BuildSourceToTarget(Snippets As Variant, SourceSheet As Worksheet, Optional TargetSheet As Worksheet)
    Application.ScreenUpdating = False
    
    Dim AppProperties As Variant, Blocks As Variant
    AppProperties = AppProperty.GetProperties
    Blocks = GetSheetTextBlocks(AppProperties, SourceSheet)
    
    Dim Index As Long
    For Index = 0 To UBound(Blocks)
        Dim Block As Variant
        Block = Blocks(Index)

        BuildBy Snippets, Block, TargetSheet
    Next
    
    Application.ScreenUpdating = True
End Function

Private Function BuildBy(Snippets As Variant, Block As Variant, Sheet As Worksheet) As Variant
    If TextBlockType.GetRowArea(Block) <> Empty Then
        BuildBy = BuildByRowArea(Snippets, Block, Sheet)
    Else
        ' Better reporting might come in handy
        MsgBox "This text block has no target. Please specify a target pattern before building this"
        End
    End If
End Function

Private Function BuildByRowArea(Snippets As Variant, Block As Variant, Sheet As Worksheet) As Variant
    Dim TargetRow As Long, Rng As Range, TempSheet As Worksheet
    TargetRow = TextBlockType.GetRowArea(Block)
    Set Rng = Sheet.UsedRange
    Set TempSheet = ThisWorkbook.Sheets.Add
    
    Dim BlockRow As Range, BlockRange As Range, InsertRow As Range, CopyBlock As Range, LastRow As Range
    Set BlockRange = TextBlockType.GetBlock(Block)

    For Each BlockRow In BlockRange.Rows
        ' NOTE: Quick hack to insert a row at the end
        Set LastRow = RangeUtil.GetLastRow(TempSheet.UsedRange)
        LastRow.Offset(1).EntireRow.Insert
        Set InsertRow = RangeUtil.GetLastRow(TempSheet.UsedRange)
        
        If LineType.IsSnippetLine(BlockRow) Then
            Dim SnippetName As String, Snippet As Variant
            SnippetName = LineType.ParseSnippetName(BlockRow)
            Snippet = FindSnippetByName(SnippetName, Snippets)
            
            If DataType.IsNil(Snippet) Then
                ' NOTE: Improve this
                Util.DeleteSheet TempSheet
                MsgBox SnippetName & " does not exist. Better recheck your definitions"
                End
            End If
            
            Set CopyBlock = SnippetType.GetBlock(Snippet)
        Else
            Set CopyBlock = BlockRow
        End If
        
        CopyBlock.Copy InsertRow
    Next
    
    'Render properties
    RenderBlockProperties Block, TempSheet

    
    ' NOTE: Dirty remove of range
    Dim Index As Long
    For Index = Sheet.UsedRange.Rows.CountLarge + TargetRow To TargetRow Step -1
        Sheet.Rows(Index).Delete
    Next
    
    Dim DestRow As Range
    Set DestRow = Sheet.UsedRange.Rows(TargetRow)
    
    TempSheet.UsedRange.Copy DestRow
    
    Util.DeleteSheet TempSheet
End Function

Private Function RenderBlockProperties(Block As Variant, Sheet As Worksheet)
    Dim Index As Long, Pairs As Variant, Rng As Range
    Pairs = ObjectType.Pairs(Block)
    Set Rng = Sheet.UsedRange
    
    Dim Key As Variant, Value As Variant, PropertyName As String, Pair As Variant
    For Index = 0 To UBound(Pairs)
        Pair = Pairs(Index)
        Key = ArrayUtil.First(Pair)
        
        If IsObject(Pair(1)) Then
            ' Do nothing for objects
        End If
        
        Value = ArrayUtil.Second(Pair)
        PropertyName = LineType.WrapPropertyName(Key)
        
        ' Naive replacement
        Rng.Replace What:=PropertyName, Replacement:=Value, _
            SearchOrder:=xlByRows, LookAt:=xlPart
    Next
    
    ' NOTE: Naive replacement for counter, should be done functionally and refactored
    Dim CounterName As String, FirstFind As Range, Counter As Long, NextFind As Range, CurrentValue As Variant
    CounterName = LineType.WrapPropertyName("_counter")
    Counter = 0
    Set FirstFind = Rng.Find(What:=CounterName, _
        SearchOrder:=xlByRows, LookAt:=xlPart, SearchDirection:=xlNext)
        Counter = 0
    If Not FirstFind Is Nothing Then
        Set NextFind = FirstFind
        Do
            Counter = Counter + 1
            CurrentValue = NextFind.Value
            CurrentValue = Replace(CurrentValue, CounterName, Counter)
            
            NextFind.Value = CurrentValue
            
            Set NextFind = Rng.FindNext(NextFind)
        Loop While Not NextFind Is Nothing
    End If
End Function

Private Function FindSnippetByName(Name As String, Snippets As Variant) As Variant
    Dim Index As Long, Snippet As Variant, SnippetName As String
    
    For Index = 0 To UBound(Snippets)
        Snippet = Snippets(Index)
        SnippetName = SnippetType.GetName(Snippet)
        
        If SnippetName = Name Then
            FindSnippetByName = Snippet
            Exit Function
        End If
    Next
    
    FindSnippetByName = ObjectType.Create
End Function

Private Function GetSheetTextBlocks(AppProperties As Variant, Sheet As Worksheet) As Variant
    Dim Rng As Range
    Set Rng = Sheet.UsedRange
    
    Dim BlockProperties As Variant
    BlockProperties = PropertyType.GetSheetProperties(Sheet)
    BlockProperties = ObjectType.Merge(BlockProperties, AppProperties)
    
    Dim BlockStart As String, BlockEnd As String
    BlockStart = AppProperty.GetBlockStartProperty(AppProperties)
    BlockEnd = AppProperty.GetBlockEndProperty(AppProperties)
    
    Dim Blocks As Variant, Index As Long, Row As Range
    Index = 0
    Blocks = ArrayUtil.CreateWithSize(Rng.Rows.CountLarge)
    
    Dim RowValue As Variant, StartRow As Range
    Set StartRow = Nothing
    
    ' Like with AppSnippet, ubiquity is nice
    For Each Row In Rng.Rows
        RowValue = Row.Columns(1).Value
        
        If (StartRow Is Nothing) And (RowValue = BlockStart) Then
            Set StartRow = Row
        ElseIf (Not StartRow Is Nothing) And (RowValue = BlockEnd) Then
            If StartRow.Address <> Row.Address Then
                Dim Block As Variant
                Block = ParseBlockStartEndRow(StartRow, Row, BlockProperties)
                
                Blocks(Index) = Block
                Index = Index + 1
            End If
                        
            Set StartRow = Nothing
        End If
    Next
    
    If Index > 0 Then
        ReDim Preserve Blocks(0 To Index - 1)
    Else
        Blocks = Array()
    End If

    GetSheetTextBlocks = Blocks
End Function

Private Function ParseBlockStartEndRow(StartRow As Range, EndRow As Range, Properties As Variant) As Variant
    Dim Sheet As Worksheet, TextBlockRange As Range, TextBlockStartRow As Range, TextBlockEndRow As Range
    Set Sheet = StartRow.Worksheet
    Set TextBlockStartRow = StartRow.Offset(1)
    Set TextBlockEndRow = EndRow.Offset(-1)
    Set TextBlockRange = Sheet.Range(TextBlockStartRow, TextBlockEndRow)
    
    Dim BlockProperties As Variant
    BlockProperties = PropertyType.GetRangeProperties(TextBlockRange)
    BlockProperties = ObjectType.Merge(BlockProperties, Properties)
    
    Dim BlockRange As Range, BlockStartRow As Range
    For Each BlockStartRow In TextBlockRange.Rows
        If Not PropertyType.IsPropertyRow(BlockStartRow) Then
            Set BlockRange = Sheet.Range(BlockStartRow, TextBlockEndRow)
            ' Copy pasting from AppSnippet
            Exit For
        End If
    Next
    
    ParseBlockStartEndRow = TextBlockType.Create(BlockProperties, BlockRange)
End Function

Private Function GetTextSheets(Properties As Variant) As String
    GetTextSheets = Util.GetSheetsByNameGlob(GetTextPrefixProperty(Properties) & "*")
End Function

Public Function IsTextBlockSheet(Sheet As Worksheet) As Boolean
    IsTextBlockSheet = Sheet.Name Like GetTextPrefixProperty(AppProperty.GetProperties()) & "*"
End Function

Private Function GetTextPrefixProperty(Properties As Variant) As String
    GetTextPrefixProperty = ObjectType.Take(TEXT_PREFIX_KEY, "TXT", Properties)
End Function


