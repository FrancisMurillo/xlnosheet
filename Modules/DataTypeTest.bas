Attribute VB_Name = "DataTypeTest"
' Test DataType module

Public Sub TestIsNil()
    VaseAssert.AssertTrue DataType.IsNil("")
    VaseAssert.AssertTrue DataType.IsNil(Empty)
    VaseAssert.AssertTrue DataType.IsNil(Array())
    VaseAssert.AssertTrue DataType.IsNil(Nothing)
    
    VaseAssert.AssertFalse DataType.IsNil("MEOW")
    VaseAssert.AssertFalse DataType.IsNil(Array(1))
    VaseAssert.AssertFalse DataType.IsNil(True)
    VaseAssert.AssertFalse DataType.IsNil(VBAProject.ThisWorkbook)
End Sub
