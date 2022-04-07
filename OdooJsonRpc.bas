Attribute VB_Name = "OdooJsonRpc"
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

'
' External API - odoo docs
'
' Odoo is usually extended internally via modules, but many of its features and
' all of its data are also available from the outside for external analysis or
' integration with various tools. Part of the Models API is easily available over
' XML-RPC and accessible from a variety of languages.
'
' see also: https://www.odoo.com/documentation/15.0/developer/misc/api/odoo.html
'

Option Explicit

Public Const CURL_DEMO As String = "https://demo.odoo.com"

Public Const CERR_STATUSCODE   As Long = 2001 + vbObjectError   ' web response error
Public Const CERR_RESPONSE     As Long = 2002 + vbObjectError   ' JSON-RPC error
Public Const CERR_AUTHENTICATE As Long = 2004 + vbObjectError   ' authentication failed

Public Const CMETHOD_CHECK_ACCESS_RIGHTS As String = "check_access_rights"
Public Const CMETHOD_SEARCH              As String = "search"
Public Const CMETHOD_SEARCH_COUNT        As String = "search_count"
Public Const CMETHOD_READ                As String = "read"
Public Const CMETHOD_FIELDS_GET          As String = "fields_get"
Public Const CMETHOD_SEARCH_READ         As String = "search_read"
Public Const CMETHOD_CREATE              As String = "create"
Public Const CMETHOD_WRITE               As String = "write"
Public Const CMETHOD_NAME_GET            As String = "name_get"
Public Const CMETHOD_UNLINK              As String = "unlink"

Public Function CreateJsonRequest(aMethod As String, Optional aParams As Variant = Nothing, Optional aId As Long = -1) As Object
    Dim dic As New Dictionary
    If aId < 0 Then
        aId = CLng(Rnd * 100000000)
    End If
    With dic
        ' jsonrpc
        .Add "jsonrpc", "2.0"
        ' method
        .Add "method", aMethod
        If Not aParams Is Nothing Then
            ' params
            .Add "params", aParams
        End If
        ' id
        .Add "id", aId
    End With
    Set CreateJsonRequest = dic
End Function

Public Function CreateJsonRequestCall(aParamsService As String, aParamsMethod As String, Optional aParamsArgs As Collection = Nothing, Optional aId As Long = -1) As Object
    Dim params As New Dictionary
    If aParamsArgs Is Nothing Then
        Set aParamsArgs = New Collection
    End If
    With params
        ' service
        .Add "service", aParamsService
        ' method
        .Add "method", aParamsMethod
        ' args
        .Add "args", aParamsArgs
    End With
    Set CreateJsonRequestCall = CreateJsonRequest("call", params, aId)
End Function

Public Function PostJsonRpc(aJsonBody As Dictionary, aBaseUrl As String, Optional aNamespace As String = "jsonrpc") As Dictionary
    Dim postUrl As String
    Dim wc As New WebClient
    Dim wr As WebResponse
    Dim dic As Dictionary
    Dim errSrc As String
    Dim errDsc As String
    postUrl = WebHelpers.JoinUrl(aBaseUrl, aNamespace)
    Set wr = wc.PostJson(postUrl, aJsonBody)
    If wr.StatusCode <> WebStatusCode.Ok Then
        errSrc = postUrl
        errDsc = "web response error (status code: " & wr.StatusCode & " )" & vbCrLf & postUrl
        LogError errDsc, errSrc, CERR_STATUSCODE
        Err.Raise CERR_STATUSCODE, errSrc, errDsc
    End If
    Set dic = JsonConverter.ParseJson(wr.Content)
    If Not dic.Exists("result") Then
        With dic("error")
            errSrc = .Item("message")
            With .Item("data")
                errDsc = .Item("name") & vbCrLf & .Item("message")
            End With
        End With
        LogError errDsc, errSrc, CERR_RESPONSE
        Err.Raise CERR_RESPONSE, errSrc, errDsc
    End If
    Set PostJsonRpc = dic  ' exists "result" key.
End Function

Public Function StartStart(aDemoUrl As String) As Dictionary
    Set StartStart = PostJsonRpc(CreateJsonRequest("start"), aDemoUrl, "start")
End Function

Public Function CommonVersion(aBaseUrl As String) As Dictionary
    Set CommonVersion = PostJsonRpc(CreateJsonRequestCall("common", "version"), aBaseUrl)
End Function

Public Function CommonAuthenticate(aBaseUrl As String, aDbName As String, aUsername As String, aPassword As String) As Dictionary
    Dim args As New Collection
    Dim dic As Dictionary
    Dim errSrc As String
    Dim errDsc As String
    With args
        .Add aDbName        ' dbname
        .Add aUsername      ' username
        .Add aPassword      ' password
        .Add New Collection ' (empty list)
    End With
    Set dic = PostJsonRpc(CreateJsonRequestCall("common", "authenticate", args), aBaseUrl)
    If VarType(dic("result")) = vbBoolean Then
        Debug.Assert dic("result") = False
        errSrc = "common.authenticate"
        errDsc = "authentication failed"
        LogError errDsc, errSrc, CERR_AUTHENTICATE
        Err.Raise CERR_AUTHENTICATE, errSrc, errDsc
    End If
    Set CommonAuthenticate = dic ' Type of json("result") is Long.
End Function

Public Function ObjectExecuteKw(aBaseUrl As String, aDbName As String, aUserId As Long, aPassword As String, aModelName As String, aMethodName As String, aListParam As Variant, Optional aOptions As Variant = "") As Dictionary
    Dim args As New Collection
    With args
        .Add aDbName    ' the database to use, a string
        .Add aUserId    ' the user id (retrieved through authenticate), an integer
        .Add aPassword  ' the userÅfs password, a string
        
        .Add aModelName                             ' the model name, a string
        .Add aMethodName                            ' the method name, a string
        If IsObject(aListParam) Then                ' an array/list of parameters passed by position
            .Add aListParam
        Else
            .Add JsonConverter.ParseJson(aListParam)
        End If
        If IsObject(aOptions) Then                  ' a mapping/dict of parameters to pass by keyword (optional)
            If Not aOptions Is Nothing Then
                .Add aOptions
            End If
        Else
            If aOptions <> "" Then
                .Add JsonConverter.ParseJson(aOptions)
            End If
        End If
    End With
    Set ObjectExecuteKw = PostJsonRpc(CreateJsonRequestCall("object", "execute_kw", args), aBaseUrl)
End Function
