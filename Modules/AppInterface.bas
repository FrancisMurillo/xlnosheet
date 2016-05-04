Attribute VB_Name = "AppInterface"
' Interface for key bindings

Public Sub InsertFilePath()
Attribute InsertFilePath.VB_ProcData.VB_Invoke_Func = "O\n14"
    Dim Cell  As Range, Path As String
    Set Cell = ActiveCell
    
    Path = FileUtil.BrowseFile
    
    If DataType.IsNotNil(Path) Then
        ActiveCell.Value = Path
    End If
End Sub

Public Sub TestCurrentTextBlock()
Attribute TestCurrentTextBlock.VB_ProcData.VB_Invoke_Func = "T\n14"
    Dim Snippets As Variant
    Snippets = AppSnippet.GetSnippets
    
    Dim CurSheet As Worksheet, CurName As String
    Set CurSheet = ActiveSheet
    CurName = CurSheet.Name
    
    If Not AppTextBlock.IsTextBlockSheet(CurSheet) Then
        MsgBox "Current sheet is not a text block sheet"
        End
    End If
    
    OptionUtil.SilenceUpdates
    
    Dim TestSheet As Worksheet, TestSheetName As String
    TestSheetName = AppConstant.TEST_SHEET_NAME
    Set TestSheet = Util.CreateSheetWithName(TestSheetName)

    AppTextBlock.BuildSourceToTarget Snippets, CurSheet, TestSheet
    
    OptionUtil.UnsilenceUpdates
End Sub

Public Sub CompileCurrentTextBlock()
Attribute CompileCurrentTextBlock.VB_ProcData.VB_Invoke_Func = "C\n14"
    Dim Snippets As Variant
    Snippets = AppSnippet.GetSnippets
    
    Dim CurSheet As Worksheet, CurName As String
    Set CurSheet = ActiveSheet
    CurName = CurSheet.Name
    
    If Not AppTextBlock.IsTextBlockSheet(CurSheet) Then
        MsgBox "Current sheet is not a text block sheet"
        End
    End If
    
    OptionUtil.SilenceUpdates
    
    Dim AppProperties As Variant, BlockProperties As Variant, BookPath As String, SheetPath As String
    AppProperties = AppProperty.GetProperties()
    BlockProperties = PropertyType.GetSheetProperties(CurSheet)
    BlockProperties = ObjectType.Merge(AppProperties, BlockProperties)
    BookPath = AppProperty.GetProjectPathProperty(BlockProperties)
    SheetPath = AppProperty.GetSheetPathProperty(BlockProperties)
    
    Dim TargetBook As Workbook, TargetSheet As Worksheet
    
    Set TargetSheet = Util.OpenSheetByPathAndName(BookPath, SheetPath)
    
    If TargetSheet Is Nothing Then
        RuntimeUtil.ThrowError BookPath & " or " & SheetPath & " cannot be open or found. Double check your config."
    End If
    
    AppTextBlock.BuildSourceToTarget Snippets, CurSheet, TargetSheet
    
    OptionUtil.UnsilenceUpdates
End Sub


Public Sub CompileAllTextBlocks()
Attribute CompileAllTextBlocks.VB_ProcData.VB_Invoke_Func = "A\n14"
    Dim Snippets As Variant
    Snippets = AppSnippet.GetSnippets
    
    
    Dim AppProperties As Variant, BlockProperties As Variant, BookPath As String, SheetPath As String, TextSheets As Variant
    AppProperties = AppProperty.GetProperties()
    TextSheets = AppTextBlock.GetTextSheets(AppProperties)
    
    If DataType.IsNil(TextSheets) Then
        RuntimeUtil.ThrowError "No text blocks found. Double check if you got the config right."
    End If
    
    
    OptionUtil.SilenceUpdates
    
    Dim Index As Long, TextSheet As Worksheet, TargetSheet As Worksheet
    
    For Index = 0 To UBound(TextSheets)
        Set TextSheet = TextSheets(Index)
        
        BlockProperties = PropertyType.GetSheetProperties(TextSheet)
        BlockProperties = ObjectType.Merge(AppProperties, BlockProperties)
        
        If TextBlockType.GetExcludeCompile(BlockProperties) Then
            ' Skip this block
        Else
            BookPath = AppProperty.GetProjectPathProperty(BlockProperties)
            SheetPath = AppProperty.GetSheetPathProperty(BlockProperties)
            
            Set TargetSheet = Util.OpenSheetByPathAndName(BookPath, SheetPath)
            
            If TargetSheet Is Nothing Then
                MsgBox "Sheet in target does not exist"
                TargetBook.Close True
            End If
        End If
        
        AppTextBlock.BuildSourceToTarget Snippets, TextSheet, TargetSheet
    Next
    
    OptionUtil.UnsilenceUpdates
End Sub


