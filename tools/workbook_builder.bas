Attribute VB_Name = "workbook_builder"
' External API - odoo-JSON-RPC-VBA
'
' MIT License
'
' Copyright (c) 2022 jp-rad
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

Private Function GetScriptFolderName() 'As String
    GetScriptFolderName = GetFileSystemObject().GetParentFolderName(WScript.ScriptFullName)
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
    
    Dim tmp 'As String
    ' odoo-JSON-RPC-VBA
    tmp = "../OdooJsonRpc.bas"
    wbk.VBProject.VBComponents.Import fso.BuildPath(cur, tmp)
    If t Then
        tmp = "../OdooJsonRpcTest.bas"
        wbk.VBProject.VBComponents.Import fso.BuildPath(cur, tmp)
    End If
    tmp = "../OdDomainBuilder.cls"
    wbk.VBProject.VBComponents.Import fso.BuildPath(cur, tmp)
    tmp = "../OdDomainCriteria.cls"
    wbk.VBProject.VBComponents.Import fso.BuildPath(cur, tmp)
    tmp = "../OdResult.cls"
    wbk.VBProject.VBComponents.Import fso.BuildPath(cur, tmp)
    tmp = "../OdServiceCommon.cls"
    wbk.VBProject.VBComponents.Import fso.BuildPath(cur, tmp)
    tmp = "../OdServiceObject.cls"
    wbk.VBProject.VBComponents.Import fso.BuildPath(cur, tmp)
    tmp = "../OdServiceStart.cls"
    wbk.VBProject.VBComponents.Import fso.BuildPath(cur, tmp)
    tmp = "../OdWebClient.cls"
    wbk.VBProject.VBComponents.Import fso.BuildPath(cur, tmp)
    ' VBA-tools/VBA-Web
    tmp = "../imports/vba-web/src/WebHelpers.bas"
    wbk.VBProject.VBComponents.Import fso.BuildPath(cur, tmp)
    tmp = "../imports/vba-web/src/IWebAuthenticator.cls"
    wbk.VBProject.VBComponents.Import fso.BuildPath(cur, tmp)
    'tmp = "../imports/vba-web/src/WebAsyncWrapper.cls"
    'wbk.VBProject.VBComponents.Import fso.BuildPath(cur, tmp)
    tmp = "../imports/vba-web/src/WebClient.cls"
    wbk.VBProject.VBComponents.Import fso.BuildPath(cur, tmp)
    tmp = "../imports/vba-web/src/WebRequest.cls"
    wbk.VBProject.VBComponents.Import fso.BuildPath(cur, tmp)
    tmp = "../imports/vba-web/src/WebResponse.cls"
    wbk.VBProject.VBComponents.Import fso.BuildPath(cur, tmp)
    'VBA-tools/VBA -JSON
    tmp = "../imports/vba-json/JsonConverter.bas"
    wbk.VBProject.VBComponents.Import fso.BuildPath(cur, tmp)
    ' VBA-tools/VBA -Dictionary
    tmp = "../imports/vba-dictionary/Dictionary.cls"
    wbk.VBProject.VBComponents.Import fso.BuildPath(cur, tmp)
    
    Dim fnm 'As String
    If t Then
        fnm = BuildUniqueFilePath(cur, "../JSON-RPC Tutorial", "xlsm")
    Else
        fnm = BuildUniqueFilePath(cur, "../JSON-RPC Blank", "xlsm")
    End If
    wbk.SaveAs fnm, 52 'xlOpenXMLWorkbookMacroEnabled
    
    wbk.Close
    appExcel.Quit
    
End Sub
