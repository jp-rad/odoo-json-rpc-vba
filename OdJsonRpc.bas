Attribute VB_Name = "OdJsonRpc"
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

Private mJsonRpcId As Long

Private Function createJsonRpcId() As Long
    createJsonRpcId = mJsonRpcId + 1
End Function

Private Function GetJsonRpcId(aJsonRpc As Dictionary) As Long
    If aJsonRpc.Exists("id") Then
        GetJsonRpcId = aJsonRpc.Item("id")
    Else
        Debug.Print "Missing 'id'"
        Debug.Assert False
        GetJsonRpcId = -1
    End If
End Function

Private Function CreateJsonRpc(aMethod As String, Optional aParams As Variant = Nothing) As Dictionary
    Dim dic As New Dictionary
    If aParams Is Nothing Then
        Set aParams = New Dictionary
    End If
    With dic
        ' jsonrpc
        .Add "jsonrpc", "2.0"
        ' method
        .Add "method", aMethod
        ' params
        .Add "params", aParams
        ' id
        .Add "id", createJsonRpcId()
    End With
    Set CreateJsonRpc = dic
End Function


Private Function GetHeaderFromWebResponse(wr As WebResponse, header As String) As String
    Dim dict As Dictionary
    
    For Each dict In wr.Headers
        If dict("Key") = header Then
            GetHeaderFromWebResponse = dict("Value")
            Exit Function
        End If
    Next dict
    
End Function

Public Function PostJson(aOdConnection As OdConnection, aUrlPath As String, aJsonRpc As Dictionary) As Dictionary
    Dim sUrl As String
    Dim wr As WebResponse
    Dim errSrc As String
    Dim errDsc As String
    sUrl = WebHelpers.JoinUrl(aOdConnection.BaseUrl, aUrlPath)
    Set wr = aOdConnection.RefWebClient.PostJson(sUrl, aJsonRpc)
    If wr.StatusCode = 301 Or wr.StatusCode = 302 Or wr.StatusCode = 307 Then
        sUrl = GetHeaderFromWebResponse(wr, "Location")
        Debug.Print "[" & wr.StatusCode & "]" & " Location:" & sUrl, "OdService.PostJson"
        Set wr = aOdConnection.RefWebClient.PostJson(sUrl, aJsonRpc)
    End If
    If wr.StatusCode <> WebStatusCode.Ok Then
        errSrc = sUrl
        errDsc = "web response error (status code: " & wr.StatusCode & " )" & vbCrLf & sUrl
        LogError errDsc, errSrc, Od.CERR_STATUSCODE
        Err.Raise Od.CERR_STATUSCODE, errSrc, errDsc
    End If
    Set PostJson = JsonConverter.ParseJson(wr.Content)
    If Not PostJson.Exists("result") Then
        With PostJson.Item("error")
            errSrc = .Item("message")
            With .Item("data")
                errDsc = .Item("name") & vbCrLf & .Item("message")
            End With
        End With
        LogError errDsc, errSrc, Od.CERR_RESPONSE
        Err.Raise Od.CERR_RESPONSE, errSrc, errDsc
    End If
    If GetJsonRpcId(aJsonRpc) <> GetJsonRpcId(PostJson) Then
        errSrc = "OdJsonRpc.PostJson"
        errDsc = "Invalid JSON-RPC ID." & vbCrLf & "Expected: " & GetJsonRpcId(aJsonRpc) & vbCrLf & "Actual: " & GetJsonRpcId(PostJson)
        Debug.Print errDsc
        Debug.Assert False
        LogError errDsc, errSrc, Od.CERR_JSONRPC_ID
        Err.Raise Od.CERR_JSONRPC_ID, errSrc, errDsc
    End If
End Function

Public Function TestDatabase(aOdConnection As OdConnection) As Dictionary
    Set TestDatabase = PostJson(aOdConnection, "start", CreateJsonRpc("start"))
End Function

Private Function PostJsonRpc(aOdConnection As OdConnection, aJsonRpc As Dictionary) As Dictionary
    Set PostJsonRpc = PostJson(aOdConnection, "jsonrpc", aJsonRpc)
End Function

Private Function CreateJsonRpcCall(aParamsService As String, aParamsMethod As String, Optional aParamsArgs As Collection = Nothing) As Dictionary
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
    Set CreateJsonRpcCall = CreateJsonRpc("call", params)
End Function

Public Function JsonRpcCommonVersion(aOdConnection As OdConnection) As Dictionary
    Set JsonRpcCommonVersion = PostJsonRpc(aOdConnection, CreateJsonRpcCall("common", "version"))
End Function

Public Function JsonRpcCommonAuthenticate(aOdConnection As OdConnection) As Dictionary
    Dim args As New Collection
    Dim dic As Dictionary
    Dim errSrc As String
    Dim errDsc As String
    With args
        .Add aOdConnection.DbName      ' dbname
        .Add aOdConnection.Username    ' username
        .Add aOdConnection.Password    ' password
        .Add New Collection            ' (empty list)
    End With
    Set dic = PostJsonRpc(aOdConnection, CreateJsonRpcCall("common", "authenticate", args))
    If VarType(dic("result")) = vbBoolean Then
        Debug.Assert dic("result") = False
        errSrc = "common.authenticate"
        errDsc = "authentication failed"
        LogError errDsc, errSrc, Od.CERR_AUTHENTICATE
        Err.Raise Od.CERR_AUTHENTICATE, errSrc, errDsc
    End If
    Set JsonRpcCommonAuthenticate = dic ' Type of json("result") is Long.
End Function

Public Function JsonRpcObjectExecuteKw(aOdConnection As OdConnection, aModelName As String, aMethodName As String, Optional aParams As Variant = "", Optional aOptions As Variant = "") As Dictionary
    Dim args As New Collection
    With args
        .Add aOdConnection.DbName    ' the database to use, a string
        .Add aOdConnection.UserId    ' the user id (retrieved through authenticate), an integer
        .Add aOdConnection.Password  ' the userÅfs password, a string
        
        .Add aModelName                 ' the model name, a string
        .Add aMethodName                ' the method name, a string
        If IsObject(aParams) Then       ' an array/list of parameters passed by position
            If aParams Is Nothing Then
                .Add New Collection
            Else
                .Add aParams
            End If
        Else
            If aParams = "" Then
                .Add New Collection
            Else
                .Add JsonConverter.ParseJson(aParams)
            End If
        End If
        If IsObject(aOptions) Then      ' a mapping/dict of parameters to pass by keyword (optional)
            If Not aOptions Is Nothing Then
                .Add aOptions
            End If
        Else
            If aOptions <> "" Then
                .Add JsonConverter.ParseJson(aOptions)
            End If
        End If
    End With
    Set JsonRpcObjectExecuteKw = PostJsonRpc(aOdConnection, CreateJsonRpcCall("object", "execute_kw", args))
End Function

