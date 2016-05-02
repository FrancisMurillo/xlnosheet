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
    
    Dim TestSheet As Worksheet, TestSheetName As String
    TestSheetName = AppConstant.TEST_SHEET_NAME
    Set TestSheet = Util.CreateSheetWithName(TestSheetName)

    AppTextBlock.BuildSourceToTarget Snippets, CurSheet, TestSheet
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
    
    Dim AppProperties As Variant, BlockProperties As Variant, BookPath As String, SheetPath As String
    AppProperties = AppProperty.GetProperties()
    BlockProperties = PropertyType.GetSheetProperties(CurSheet)
    BlockProperties = ObjectType.Merge(AppProperties, BlockProperties)
    BookPath = AppProperty.GetProjectPathProperty(BlockProperties)
    SheetPath = AppProperty.GetSheetPathProperty(BlockProperties)
    
    Dim TargetBook As Workbook, TargetSheet As Worksheet
    
    ' NOTE: Open workbook, refactor this
    Set TargetBook = Util.GetBookByPath(BookPath)
    If TargetBook Is Nothing Then
        ' NOTE: Hack from Compare Report Tool to open a book
        DoEvents
        DoEvents
        Set TargetBook = Workbooks.Open(BookPath)
        DoEvents
        DoEvents
    End If
    
    
    ' Error handling for book
    
    Set TargetSheet = Util.GetSheetByName(SheetPath, TargetBook)
    
    If TargetSheet Is Nothing Then
        MsgBox "Sheet in target does not exist"
        TargetBook.Close True
        End
    End If
    
    AppTextBlock.BuildSourceToTarget Snippets, CurSheet, TargetSheet
    
    TargetBook.Save
End Sub

