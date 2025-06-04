Attribute VB_Name = "TestOdClient"
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

Private Const CBASEURL As String = "https://localhost"
Private Const CINSECURE As Boolean = True
Private Const CFOLLOWREDIRECTS As Boolean = False
Private Const CDBNAME As String = "dev_odoo"
Private Const CUSERNAME As String = "admin"
Private Const CPASSWORD As String = "admin"

Private mConn As New Collection

Public Sub InitAuthConn()
    Set mConn = New Collection
End Sub

Public Function GetAuthConn(Optional aConnName As String = "", Optional aForceFetch As Boolean = False, _
Optional aBaseUrl As String = CBASEURL, Optional aInsecure As Boolean = CINSECURE, Optional aFollowRedirects As Boolean = CFOLLOWREDIRECTS, _
Optional aDbName As String = CDBNAME, Optional aUserName As String = CUSERNAME, Optional aPassword As String = CPASSWORD) As OdClient
On Error Resume Next
    If aForceFetch Then
        mConn.Remove aConnName
    End If
    Set GetAuthConn = mConn.Item(aConnName)
On Error GoTo 0
    If GetAuthConn Is Nothing Then
        Dim oClient As OdClient
        Set oClient = NewOdClient
        
        With oClient
            .BaseUrl = aBaseUrl
            .SetInsecure aInsecure
            .SetFollowRedirects aFollowRedirects
            .DbName = aDbName
            .Username = aUserName
            .Password = aPassword
        End With
        
        oClient.Common.Authenticate
        
        mConn.Add oClient, aConnName
        Set GetAuthConn = oClient
    End If
End Function

Public Sub TestGetAuthConn()
    Dim oClient As OdClient
    Dim oTest As OdResult
    
    ' Initialize
    InitAuthConn
    
    ' create
    Set oClient = GetAuthConn(aConnName:="")
    Debug.Assert oClient.IsAuthenticated
    ' cached
    Set oClient = GetAuthConn(aConnName:="")
    Debug.Assert oClient.IsAuthenticated
    ' create - force
    Set oClient = GetAuthConn(aConnName:="", aForceFetch:=True)
    Debug.Assert oClient.IsAuthenticated
    
    ' --- Test Database ---
    Set oTest = NewOdClient.StartTestDatabase()
    ' create
    Set oClient = GetAuthConn(aConnName:="demo", aBaseUrl:=oTest.sHost, aInsecure:=False, aDbName:=oTest.sDatabase, aUserName:=oTest.sUser, aPassword:=oTest.sPassword)
    Debug.Assert oClient.IsAuthenticated
    ' cached
    Set oClient = GetAuthConn(aConnName:="demo")
    Debug.Assert oClient.IsAuthenticated
    ' cached - force
    Set oClient = GetAuthConn(aConnName:="demo", aBaseUrl:=oTest.sHost, aInsecure:=False, aDbName:=oTest.sDatabase, aUserName:=oTest.sUser, aPassword:=oTest.sPassword, aForceFetch:=True)
    Debug.Assert oClient.IsAuthenticated
    ' ---------------------
    
    ' cached
    Set oClient = GetAuthConn(aConnName:="")
    Debug.Assert oClient.IsAuthenticated
    
    On Error Resume Next
    ' (ERROR) cached - force
    Set oClient = GetAuthConn(aConnName:="", aInsecure:=False, aForceFetch:=True)
    Debug.Assert Err.Number <> 0

End Sub

Public Sub TestCommonVerstion()
    Dim oClient As OdClient
    Dim oRet As OdResult
    Set oClient = NewOdClient
    
    oClient.BaseUrl = CBASEURL
    ' Turn off SSL validation
    oClient.SetInsecure True
    ' Follow redirects (301, 302, 307) using Location header
    oClient.SetFollowRedirects False
        
    ' Version
    Set oRet = oClient.Common.Version()
    
    Debug.Print "---------"
    Debug.Print " version"
    Debug.Print "---------"
    Debug.Print JsonConverter.ConvertToJson(oRet.Result, 4)
    Debug.Print
    
End Sub
