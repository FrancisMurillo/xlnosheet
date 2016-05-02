Attribute VB_Name = "PropertyType"
' The property module

Public Const KEY_PREFIX As String = ":"

Public Function GetRangeProperties(Rng As Range) As Variant
    Dim Properties As Variant, Index As Long, Row As Range
    Properties = ObjectType.Create()
    
    For Each Row In Rng.Rows
        Key = Row.Columns(1).Value
        Value = Row.Columns(2).Value
        
        If Key = "" Then
            ' Skip blank lines
        ElseIf IsKey(Key) Then
            Properties = ObjectType.Place(ParseKey(Key), Value, Properties)
        Else
            GetRangeProperties = Properties
            Exit Function
        End If
    Next
    
    GetRangeProperties = Properties
End Function

Public Function GetSheetProperties(Sheet As Worksheet) As Variant
    GetSheetProperties = GetRangeProperties(Sheet.UsedRange)
End Function

Public Function ParseKey(RawKey As Variant) As String
    ParseKey = Replace(RawKey, KEY_PREFIX, "")
End Function

Public Function IsPropertyRow(Row As Range) As Boolean
    Dim Values As Variant, Key As Variant
    
    Key = Row.Columns(1).Value

    IsPropertyRow = (Key = "") Or (IsKey(Key))
End Function

Public Function IsKey(RawKey As Variant) As String
    IsKey = (RawKey <> "") And (Left(RawKey, 1) = KEY_PREFIX)
End Function
