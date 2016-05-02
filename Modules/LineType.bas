Attribute VB_Name = "LineType"
' A type for reading blocks

Public Const SNIPPET_LINE_GLOB As String = "{{*}}"

Public Function IsSnippetLine(Row As Range) As Boolean
    Dim Line As Variant
    Line = Row.Columns(1).Value
    
    IsSnippetLine = Line Like SNIPPET_LINE_GLOB
End Function

Public Function ParseSnippetName(Row As Range) As String
    Dim Line As Variant, Name As String
    Line = Row.Columns(1).Value

    Name = Replace(Line, "{{", "")
    Name = Replace(Name, "}}", "")
    
    ParseSnippetName = Name
End Function

Public Function WrapPropertyName(Name As Variant) As String
    WrapPropertyName = "#{" & Name & "}"
End Function
