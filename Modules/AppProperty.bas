Attribute VB_Name = "AppProperty"
' Application specific property

Public Const CONFIG_SHEET_NAME As String = "_Config"

Public Const PROJECT_PATH_KEY As String = "targetProject"
Public Const SHEET_PATH_KEY As String = "targetSheet"

Public Const BLOCK_START_KEY As String = "blockStart"
Public Const BLOCK_END_KEY As String = "blockEnd"

Public Function GetProperties()
    Dim ConfigSheet As Worksheet
    
    Set ConfigSheet = Util.GetSheetByName(CONFIG_SHEET_NAME)
    If ConfigSheet Is Nothing Then
        MsgBox "Config sheet(" & CONFIG_SHEET_NAME & ") is missing." & vbCrLf _
            & " Please restore the config sheet as it is needed in this application."
        End
    End If
    
    GetProperties = PropertyType.GetSheetProperties(ConfigSheet)
End Function

Public Function GetSheetPathProperty(Properties As Variant) As String
    GetSheetPathProperty = ObjectType.Take(SHEET_PATH_KEY, Empty, Properties)
End Function

Public Function GetProjectPathProperty(Properties As Variant) As String
    GetProjectPathProperty = ObjectType.Take(PROJECT_PATH_KEY, Empty, Properties)
End Function

Public Function GetBlockStartProperty(Properties As Variant) As String
    GetBlockStartProperty = ObjectType.Take(BLOCK_START_KEY, "<<<", Properties)
End Function

Public Function GetBlockEndProperty(Properties As Variant) As String
    GetBlockEndProperty = ObjectType.Take(BLOCK_END_KEY, ">>>", Properties)
End Function

