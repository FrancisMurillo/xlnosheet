Attribute VB_Name = "ObjectType"
' The basic datatype of this project
'
' Since this does not have a native concept of objects like in JavaScript or dictionary of Python
' we act in a lispy manner and make arrays of pairs be our king

Public Function Create() As Variant
    Create = Array()
End Function

Public Function Place(Key As Variant, Value As Variant, Object As Variant) As Variant
    If ArrayUtil.IsEmptyArray(Object) Then
        Object = Array()
    End If
    
    Dim Pair As Variant
    Pair = Array(Key, Value)

    ' Insert to first
    Place = ArrayUtil.JoinArrays(Array(Pair), Object)
End Function

Public Function Take(Key As Variant, DefaultValue As Variant, Object As Variant) As Variant
    If ArrayUtil.IsEmptyArray(Object) Then
        Take = DefaultValue
    End If
    
    Dim Index As Long, Pair As Variant, PairKey As Variant, PairValue As Variant
    
    Take = DefaultValue
    For Index = 0 To UBound(Object)
        Pair = Object(Index)
        PairKey = ArrayUtil.First(Pair)
        PairValue = ArrayUtil.Last(Pair)
        
        If PairKey = Key Then
            Take = PairValue
            Exit Function
        End If
    Next
End Function

Public Function Merge(SourceObject As Variant, ExtendObject As Variant) As Variant
    If ArrayUtil.IsEmptyArray(SourceObject) Then
        Merge = ExtendObject
        Exit Function
    End If
    
    If ArrayUtil.IsEmptyArray(ExtendObject) Then
        Merge = SourceObject
        Exit Function
    End If
    
    Merge = ArrayUtil.JoinArrays(ExtendObject, SourceObject)
End Function

Public Function Keys(Object As Variant) As Variant
    Dim ObjectKeys As Variant, Index As Long
    
    ObjectKeys = ArrayUtil.CloneSize(Object)
    For Index = 0 To UBound(Object)
        ObjectKeys(Index) = ArrayUtil.First(Object(Index))
    Next
    
    Keys = ArrayUtil.RemoveDuplicates(ObjectKeys)
End Function

Public Function Values(Object As Variant) As Variant
    Dim ObjectKeys As Variant, ObjectValues As Variant, Index As Long
    
    ObjectKeys = Keys(Object)
    ObjectValues = ArrayUtil.CloneSize(ObjectKeys)
    
    For Index = 0 To UBound(ObjectKeys)
        ObjectValues(Index) = Take(ObjectKeys(Index), Empty, Object)
    Next
    
    Values = ObjectValues
End Function
