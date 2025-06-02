' External API - odoo-JSON-RPC-VBA
'
' MIT License
'
' Copyright (c) 2022-2025 jp-rad
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.
'

Option Explicit

Private Function GetExcelApplication() 'As Excel.Application
    Set GetExcelApplication = WScript.CreateObject("Excel.Application")
End Function

Private Function GetFileSystemObject() 'As FileSystemObject
    Set GetFileSystemObject = WScript.CreateObject("Scripting.FileSystemObject")
End Function

Private Function CreateNewDictionary() 'As Dictionary
    Set CreateNewDictionary = WScript.CreateObject("Scripting.Dictionary")
End Function

Private Function GetScriptFolderName() 'As String
    GetScriptFolderName = GetFileSystemObject().GetParentFolderName(WScript.ScriptFullName)
End Function

Private Function GetVbaModules(t) 'As Dictionary
    Dim dic 'As Dictionary
    Dim fso 'As FileSystemObject
    Dim a 'As Variant
    
    Set dic = CreateNewDictionary()
    Set fso = GetFileSystemObject()
    
    With fso.OpenTextFile(fso.BuildPath(GetScriptFolderName(), "vba_modules.txt"))
        Do Until .AtEndOfLine
            a = Split(.ReadLine(), Chr(9))
            If UBound(a) = 1 Then
                If (a(0) = t) Or ("develop" = t) Then
                    dic(fso.GetBaseName(a(1))) = a(1)
                End If
            End If
        Loop
        .Close
    End With
    
    Set GetVbaModules = dic
End Function

'Private Function BuildUniqueFilePath(aPath As String, aName As String, aExtentionName As String) As String
Private Function BuildUniqueFilePath(aPath, aName, aExtentionName) 'As String
    Dim fso 'As FileSystemObject
    Dim tmp 'As String
    Dim cnt 'As Long
    Set fso = GetFileSystemObject()
    cnt = 0
    Do
        If cnt > 0 Then
            tmp = aName & "-" & cnt
        Else
            tmp = aName
        End If
        tmp = fso.BuildPath(aPath, tmp & "." & aExtentionName)
        cnt = cnt + 1
    Loop While fso.FileExists(tmp)
    BuildUniqueFilePath = tmp
End Function

Public Sub BuildWorkbookFile(t)
    Dim appExcel 'As Excel.Application
    Set appExcel = GetExcelApplication()
    appExcel.Visible = True
    
    Dim wbk 'As Workbook
    Set wbk = appExcel.Workbooks.Add()
    
    Dim fso 'As FileSystemObject
    Set fso = GetFileSystemObject()
    
    Dim cur 'As String
    cur = GetScriptFolderName()
    
    Dim dic 'As Dictionary
    Dim tmp 'As String
    Set dic = GetVbaModules(t)
    For Each tmp In dic.Items()
        wbk.VBProject.VBComponents.Import fso.BuildPath(cur, tmp)
    Next 'tmp

    Dim fnm 'As String
    If t = "library" Then
        wbk.VBProject.Name="OdooJsonRpcVBA"
        fnm = fso.BuildPath(cur, "../odoo-json-rpc-vba.xlam")
        appExcel.DisplayAlerts = False
        wbk.SaveAs fnm, 55 'xlOpenXMLAddIn
    ElseIf t = "example" Then
        fnm = BuildUniqueFilePath(cur, "../odoo-json-rpc-vba example", "xlsm")
        wbk.SaveAs fnm, 52 'xlOpenXMLWorkbookMacroEnabled
    Else    ' "develop"
        wbk.VBProject.Name="OdooJsonRpcVBADev"
        fnm = BuildUniqueFilePath(cur, "../odoo-json-rpc-vba develop", "xlsm")
        wbk.SaveAs fnm, 52 'xlOpenXMLWorkbookMacroEnabled
    End If
    
    wbk.Close
    appExcel.Quit
    
End Sub

Dim t
On Error Resume Next
Err.Clear
t = Wscript.Arguments(0)
If Err.Number = 0 Then
    On Error Goto 0
    BuildWorkbookFile t
Else
    WScript.CreateObject("WScript.Shell").PopUp  Err.Description & " (Err:" & Err.Number & ")", 5, "Do not call me directly!", 48
End If
