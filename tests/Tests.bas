Attribute VB_Name = "Tests"
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

'
' External API - odoo docs
'
' Odoo is usually extended internally via modules, but many of its features and all of its data
' are also available from the outside for external analysis or integration with various tools.
' Part of the Models API is easily available over XML-RPC and accessible from a variety of languages.
'
' https://www.odoo.com/documentation/master/developer/reference/external_api.html
'

Public Sub Run(Optional OutputPath As Variant)
    Dim Suite As New TestSuite
    Suite.Description = "vba-test"
    
    Dim Immediate As New ImmediateReporter
    Immediate.ListenTo Suite
    
    If Not IsMissing(OutputPath) And CStr(OutputPath) <> "" Then
        Dim Reporter As New FileReporter
        Reporter.WriteTo OutputPath
        Reporter.ListenTo Suite
    End If
    
    Tests_OdFilter.RunTests Suite.Group("Tests_OdFilter")
    
End Sub

Sub DoSearchRead()
    Dim oc As OdClient
    Dim rs As OdResult
    
    ' OdClient
    Set oc = OdRpc.NewOdClient("https://localhost")
    oc.RefWebClient.Insecure = True
    
    ' Login
    oc.Common.Authenticate "dev_odoo", "admin", "admin"
    
    ' Search and read
    Set rs = oc.Model("res.partner").Method("search_read").ExecuteKw( _
        "[[['is_company', '=', true]]]", _
        "{'fields': ['name', 'country_id'], 'limit': 3}" _
    )
    
    ' (JSON)
    Debug.Print
    Debug.Print "JSON: >>>>>"
    Debug.Print JsonConverter.ConvertToJson(rs.Result, 2)
    Debug.Print "<<<<<"
    
End Sub


