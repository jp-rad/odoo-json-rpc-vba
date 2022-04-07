Attribute VB_Name = "WScript"
Option Explicit

Public Function CreateObject(aClass As String) As Object
    Dim ret As Object
    Select Case aClass
    Case ""
        Set ret = Nothing
    Case Else
        Set ret = VBA.Interaction.CreateObject(aClass)
    End Select
    Set CreateObject = ret
End Function

Public Function ScriptFullName() As String
    ScriptFullName = ThisWorkbook.FullName
End Function
