Attribute VB_Name = "TextBlockType"
' Text block type module

Public Const NAME_KEY As String = "name"
Public Const BLOCK_KEY As String = "block"
Public Const ROW_AREA_KEY As String = "targetRowArea"
Public Const CELL_KEY As String = "targetCell"
Public Const EXCLUDE_COMPILE_KEY As String = "excludeCompile"

Public Function Create(Properties As Variant, Block As Range) As Variant
    Dim Object As Variant
    
    Object = Properties
    Object = ObjectType.Place(BLOCK_KEY, Block, Object)
    
    Create = Object
End Function

Public Function GetName(Block As Variant) As String
    GetName = ObjectType.Take(NAME_KEY, Empty, Block)
End Function

Public Function GetBlock(Block As Variant) As Range
    Set GetBlock = ObjectType.Take(BLOCK_KEY, Empty, Block)
End Function

Public Function GetRowArea(Block As Variant) As Long
    GetRowArea = ObjectType.Take(ROW_AREA_KEY, Empty, Block)
End Function

Public Function GetExcludeCompile(Block As Variant) As Boolean
    GetExcludeCompile = DataType.IsNil(ObjectType.Take(EXCLUDE_COMPILE_KEY, False, Block))
End Function

Public Function GetCell(Block As Variant) As String
    GetCell = ObjectType.Take(CELL_KEY, Empty, Block)
End Function
