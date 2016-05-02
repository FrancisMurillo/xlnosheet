Attribute VB_Name = "DataType"
' Data type module

Public Function IsNil(Value As Variant)
    If IsObject(Value) Then
        IsNil = Value Is Nothing
    ElseIf IsArray(Value) Then
        IsNil = ArrayUtil.IsEmptyArray(Value)
    Else
        IsNil = (Value = "") Or (Value = Empty)
    End If
End Function

Public Function IsNotNil(Value As Variant)
    IsNotNil = Not IsNil(Value)
End Function
