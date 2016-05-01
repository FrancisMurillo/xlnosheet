Attribute VB_Name = "ObjectTypeTest"
' Test module for Object

Public Sub TestTakePlace()
    Dim MyObject As Variant, Key As Variant, Value As Variant
    Key = "Meow"
    Value = 1
    
    MyObject = ObjectType.Create()
    MyObject = ObjectType.Place(Key, Value, MyObject)

    VaseAssert.AssertEqual Value, ObjectType.Take(Key, Empty, MyObject)
    
    Dim OtherKey As Variant, OtherValue As Variant
    OtherKey = 2
    OtherValue = "Roar"
    
    MyObject = ObjectType.Place(OtherKey, OtherValue, MyObject)

    VaseAssert.AssertEqual OtherValue, ObjectType.Take(OtherKey, Empty, MyObject)
End Sub

Public Sub TestKeysValues()
    Dim Object As Variant
    
    Object = ObjectType.Create()
    Object = ObjectType.Place("A", 1, Object)
    Object = ObjectType.Place("B", 2, Object)
    Object = ObjectType.Place("B", 3, Object)
    
    Dim Keys As Variant
    Keys = ObjectType.Keys(Object)
    
    VaseAssert.AssertArraySize 2, Keys
    
    VaseAssert.AssertInArray "A", Keys
    VaseAssert.AssertInArray "B", Keys
    
    Dim Values As Variant
    Values = ObjectType.Values(Object)
    
    VaseAssert.AssertArraySize 2, Values
    
    VaseAssert.AssertInArray 1, Values
    VaseAssert.AssertInArray 3, Values
End Sub

Public Sub TestPlaceTakeOverride()
    Dim Object As Variant, Key As String, Value As Variant, OverrideValue As Variant
    
    Key = "This"
    Value = "Value"
    OverrideValue = "Life"
    
    Object = ObjectType.Place(Key, Value, Array())
    Object = ObjectType.Place(Key, OverrideValue, Object)

    VaseAssert.AssertEqual OverrideValue, ObjectType.Take(Key, Empty, Object)
    
    Object = ObjectType.Place(Key, Value, Object)
    
    VaseAssert.AssertEqual Value, ObjectType.Take(Key, Empty, Object)
End Sub

Public Sub TestMerge()
    Dim LeftObject As Variant, RightObject As Variant, MergeObject As Variant
    
    LeftObject = Array()
    LeftObject = ObjectType.Place("A", 1, LeftObject)
    LeftObject = ObjectType.Place("B", 2, LeftObject)
    LeftObject = ObjectType.Place("C", 3, LeftObject)
    LeftObject = ObjectType.Place(1, True, LeftObject)
    
    RightObject = Array()
    RightObject = ObjectType.Place("a", 1, RightObject)
    RightObject = ObjectType.Place("b", 2, RightObject)
    RightObject = ObjectType.Place("c", 3, RightObject)
    RightObject = ObjectType.Place(1, False, RightObject)
    
    MergeObject = ObjectType.Merge(LeftObject, RightObject)
    
    VaseAssert.AssertEqual 1, ObjectType.Take("A", Empty, MergeObject)
    VaseAssert.AssertEqual 1, ObjectType.Take("a", Empty, MergeObject)
    VaseAssert.AssertEqual 2, ObjectType.Take("B", Empty, MergeObject)
    VaseAssert.AssertEqual 2, ObjectType.Take("b", Empty, MergeObject)
    VaseAssert.AssertEqual 3, ObjectType.Take("C", Empty, MergeObject)
    VaseAssert.AssertEqual 3, ObjectType.Take("c", Empty, MergeObject)
    
    VaseAssert.AssertEqual False, ObjectType.Take(1, Empty, MergeObject)
End Sub
