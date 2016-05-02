Attribute VB_Name = "SnippetType"
' Snippet Representation

Public Const NAME_KEY As String = "name"
Public Const BLOCK_KEY As String = "block"

Public Function Create(Properties As Variant, Block As Range) As Variant
    Dim Object As Variant
    
    Object = Properties
    Object = ObjectType.Place(BLOCK_KEY, Block, Object)
    
    Create = Object
End Function

Public Function GetName(Snippet As Variant) As String
    GetName = ObjectType.Take(NAME_KEY, Empty, Snippet)
End Function

Public Function GetBlock(Snippet As Variant) As Range
    Set GetBlock = ObjectType.Take(BLOCK_KEY, Empty, Snippet)
End Function
