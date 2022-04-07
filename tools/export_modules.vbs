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

Public Sub ExportModules()
    Dim appExcel 'As Excel.Application
    Set appExcel = GetExcelApplication()
    appExcel.Visible = True
    
    Dim fso 'As FileSystemObject
    Set fso = GetFileSystemObject()
    
    Dim cur 'As String
    cur = GetScriptFolderName()
    
    Dim fnm 'As String
    fnm = fso.BuildPath(cur, "../JSON-RPC Tutorial.xlsm")
    
    Dim wbk 'As Workbook
    Set wbk = appExcel.Workbooks.Open(fnm, , True)
    
    Dim tmp 'As String
    ' odoo-JSON-RPC-VBA
    tmp = "../OdooJsonRpc.bas"
    wbk.VBProject.VBComponents("OdooJsonRpc").Export fso.BuildPath(cur, tmp)
    tmp = "../OdooJsonRpcTest.bas"
    wbk.VBProject.VBComponents("OdooJsonRpcTest").Export fso.BuildPath(cur, tmp)
    tmp = "../OdDomainBuilder.cls"
    wbk.VBProject.VBComponents("OdDomainBuilder").Export fso.BuildPath(cur, tmp)
    tmp = "../OdResult.cls"
    wbk.VBProject.VBComponents("OdResult").Export fso.BuildPath(cur, tmp)
    tmp = "../OdServiceCommon.cls"
    wbk.VBProject.VBComponents("OdServiceCommon").Export fso.BuildPath(cur, tmp)
    tmp = "../OdServiceObject.cls"
    wbk.VBProject.VBComponents("OdServiceObject").Export fso.BuildPath(cur, tmp)
    tmp = "../OdServiceStart.cls"
    wbk.VBProject.VBComponents("OdServiceStart").Export fso.BuildPath(cur, tmp)
    tmp = "../OdWebClient.cls"
    wbk.VBProject.VBComponents("OdWebClient").Export fso.BuildPath(cur, tmp)
    ' VBA-tools/VBA-Web
    tmp = "../imports/vba-web/src/WebHelpers.bas"
    wbk.VBProject.VBComponents("WebHelpers").Export fso.BuildPath(cur, tmp)
    tmp = "../imports/vba-web/src/IWebAuthenticator.cls"
    wbk.VBProject.VBComponents("IWebAuthenticator").Export fso.BuildPath(cur, tmp)
    'tmp = "../imports/vba-web/src/WebAsyncWrapper.cls"
    'wbk.VBProject.VBComponents("WebAsyncWrapper").Export fso.BuildPath(cur, tmp)
    tmp = "../imports/vba-web/src/WebClient.cls"
    wbk.VBProject.VBComponents("WebClient").Export fso.BuildPath(cur, tmp)
    tmp = "../imports/vba-web/src/WebRequest.cls"
    wbk.VBProject.VBComponents("WebRequest").Export fso.BuildPath(cur, tmp)
    tmp = "../imports/vba-web/src/WebResponse.cls"
    wbk.VBProject.VBComponents("WebResponse").Export fso.BuildPath(cur, tmp)
    'VBA-tools/VBA -JSON
    tmp = "../imports/vba-json/JsonConverter.bas"
    wbk.VBProject.VBComponents("JsonConverter").Export fso.BuildPath(cur, tmp)
    ' VBA-tools/VBA -Dictionary
    tmp = "../imports/vba-dictionary/Dictionary.cls"
    wbk.VBProject.VBComponents("Dictionary").Export fso.BuildPath(cur, tmp)
    
    wbk.Close
    appExcel.Quit
    
End Sub

ExportModules
