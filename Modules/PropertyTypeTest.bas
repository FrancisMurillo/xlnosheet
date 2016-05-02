Attribute VB_Name = "PropertyTypeTest"
' Test Property

Public Sub TestGetSheetProperties()
    Dim TestSheet As Worksheet
    Set TestSheet = ActiveWorkbook.Sheets.Add
    
    TestSheet.Cells(1, 1).Value = ":1"
    TestSheet.Cells(1, 2).Value = "A"

    TestSheet.Cells(2, 1).Value = ":2"
    TestSheet.Cells(2, 2).Value = "B"
    
    TestSheet.Cells(3, 1).Value = ""
    
    TestSheet.Cells(4, 1).Value = ":3"
    TestSheet.Cells(4, 2).Value = "C"
    
    TestSheet.Cells(5, 1).Value = ""
    TestSheet.Cells(6, 1).Value = ""
    
    TestSheet.Cells(7, 1).Value = ":4"
    TestSheet.Cells(7, 2).Value = "D"
    
    TestSheet.Cells(8, 1).Value = "Start Script"
    
    TestSheet.Cells(9, 1).Value = ":5"
    TestSheet.Cells(9, 2).Value = "E"
    
    Dim Properties As Variant
    Properties = PropertyType.GetSheetProperties(TestSheet)
    
    VaseAssert.AssertTrue ObjectType.HasKey("1", Properties)
    VaseAssert.AssertEqual "A", ObjectType.Take("1", False, Properties)
    
    VaseAssert.AssertTrue ObjectType.HasKey("2", Properties)
    VaseAssert.AssertEqual "B", ObjectType.Take("2", False, Properties)
    
    VaseAssert.AssertTrue ObjectType.HasKey("3", Properties)
    VaseAssert.AssertEqual "C", ObjectType.Take("3", False, Properties)
    
    VaseAssert.AssertTrue ObjectType.HasKey("4", Properties)
    VaseAssert.AssertEqual "D", ObjectType.Take("4", False, Properties)
    VaseAssert.AssertFalse ObjectType.HasKey("5", Properties)
    
    Util.DeleteSheet TestSheet
End Sub
