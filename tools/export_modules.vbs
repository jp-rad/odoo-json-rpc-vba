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

Private Function GetVbaModules() 'As Dictionary
    Dim dic 'As Dictionary
    Dim fso 'As FileSystemObject
    Dim a 'As Variant
    
    Set dic = CreateNewDictionary()
    Set fso = GetFileSystemObject()
    
    With fso.OpenTextFile(fso.BuildPath(GetScriptFolderName(), "vba_modules.txt"))
        Do Until .AtEndOfLine
            a = Split(.ReadLine(), Chr(9))
            If UBound(a) = 1 Then
                dic(fso.GetBaseName(a(1))) = a(1)
            End If
        Loop
        .Close
    End With
    
    Set GetVbaModules = dic
End Function

Public Sub ExportModules()
    Dim appExcel 'As Excel.Application
    Set appExcel = GetExcelApplication()
    appExcel.Visible = True
    
    Dim fso 'As FileSystemObject
    Set fso = GetFileSystemObject()
    
    Dim cur 'As String
    cur = GetScriptFolderName()
    
    Dim fnm 'As String
    fnm = fso.BuildPath(cur, "../odoo-json-rpc-vba develop.xlsm")
    
    Dim wbk 'As Workbook
    Set wbk = appExcel.Workbooks.Open(fnm, , True)
    
    Dim tmp 'As String
    Dim cmp 'As Variant
    Dim mods 'As Dictionary
    Set mods = GetVbaModules()
    
On Error Resume Next
    For Each cmp In mods.Keys()
        tmp = mods(cmp)
        wbk.VBProject.VBComponents(cmp).Export fso.BuildPath(cur, tmp)
    Next 'cmp
On Error GoTo 0
    
    wbk.Close
    appExcel.Quit
    
End Sub

ExportModules
