Attribute VB_Name = "OdRpc"
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
' --------
'  COMMON
' --------
Private Const CERR_STATUSCODE   As Long = 2001 + vbObjectError   ' web response error
Private Const CERR_RESPONSE     As Long = 2002 + vbObjectError   ' JSON-RPC error
Private Const CERR_JSONRPC_ID   As Long = 2003 + vbObjectError   ' JSON-RPC id error
Private Const CERR_AUTHENTICATE As Long = 2004 + vbObjectError   ' authentication failed

' ----------
'  JSON-RPC
' ----------
Private mJsonRpcId As Long

' ---------
'  XML-RPC
' ---------
Private mRegisteredXml As Boolean

' ---------
'  HELPERS
' ---------
Public Function FormatDate(aDate As Date) As String
    FormatDate = Format(aDate, "yyyy-mm-dd")
End Function

Public Function ParseDate(aDateString As String) As Date
    ParseDate = CDate(aDateString)
End Function

Public Function ConvertToIsoDatetime(aDatetime As Date) As String
    ConvertToIsoDatetime = JsonConverter.ConvertToIso(aDatetime)
End Function

Public Function ParseIsoDatetime(aIsoString As String) As Date
    ParseIsoDatetime = JsonConverter.ParseIso(aIsoString)
End Function

Public Function NewOdClient() As OdClient
    Set NewOdClient = New OdClient
End Function

Public Function NewDomain() As OdFilterDomain
    Set NewDomain = New OdFilterDomain
End Function

Public Function NewField(aFieldExpr As String) As OdFilterCriteria
    With New OdFilterCriteria
        Set NewField = .SetFieldExpr(aFieldExpr)
    End With
End Function

Public Function NewCriteria(aFieldExpr As String, aOperator As String, aValueExpr As Variant) As OdFilterCriteria
    With New OdFilterCriteria
        Set NewCriteria = .SetCriteria(aFieldExpr, aOperator, aValueExpr)
    End With
End Function

Public Function NewAnd(aArity1 As Object, aArity2 As Object) As OdFilterCombi
    With New OdFilterCombi
        Set NewAnd = .SetAndLogic(aArity1, aArity2)
    End With
End Function

Public Function NewOr(aArity1 As Object, aArity2 As Object) As OdFilterCombi
    With New OdFilterCombi
        Set NewOr = .SetOrLogic(aArity1, aArity2)
    End With
End Function

Public Function NewNot(aArity As Object) As OdFilterCombi
    With New OdFilterCombi
        Set NewNot = .SetNotLogic(aArity)
    End With
End Function

Public Function NewList() As Collection
    Set NewList = New Collection
End Function

Public Function AddList(aList As Collection) As Collection
    Set AddList = New Collection
    aList.Add AddList
End Function

Public Function AddDict(aList As Collection) As Dictionary
    Set AddDict = New Dictionary
    aList.Add AddDict
End Function

Public Function NewDict() As Dictionary
    Set NewDict = New Dictionary
End Function

Public Function SetList(aDict As Dictionary, aKey As String) As Collection
    Set SetList = New Collection
    Set aDict(aKey) = SetList
End Function

Public Function SetDict(aDict As Dictionary, aKey As String) As Dictionary
    Set SetDict = New Dictionary
    Set aDict(aKey) = SetDict
End Function

' ----------
'  JSON-RPC
' ----------
Private Function createJsonRpcId() As Long
    createJsonRpcId = mJsonRpcId + 1
End Function

Private Function GetJsonRpcId(aJsonRpc As Dictionary) As Long
    If aJsonRpc.Exists("id") Then
        GetJsonRpcId = aJsonRpc.Item("id")
    Else
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
        Set wr = aOdConnection.RefWebClient.PostJson(sUrl, aJsonRpc)
    End If
    If wr.StatusCode = 400 Then
        errSrc = sUrl
        errDsc = "web response error (status code: " & wr.StatusCode & " )" & vbCrLf & sUrl
        errDsc = errDsc & vbCrLf & "The redirect may have failed. Try setting 'OdClient.SetFollowRedirects' to 'False'."""
        LogError errDsc, errSrc, CERR_STATUSCODE
        Err.Raise CERR_STATUSCODE, errSrc, errDsc
    End If
    If wr.StatusCode <> WebStatusCode.Ok Then
        errSrc = sUrl
        errDsc = "web response error (status code: " & wr.StatusCode & " )" & vbCrLf & sUrl
        LogError errDsc, errSrc, CERR_STATUSCODE
        Err.Raise CERR_STATUSCODE, errSrc, errDsc
    End If
    Set PostJson = JsonConverter.ParseJson(wr.Content)
    If Not PostJson.Exists("result") Then
        With PostJson.Item("error")
            errSrc = .Item("message")
            With .Item("data")
                errDsc = .Item("name") & vbCrLf & .Item("message")
            End With
        End With
        LogError errDsc, errSrc, CERR_RESPONSE
        Err.Raise CERR_RESPONSE, errSrc, errDsc
    End If
    If GetJsonRpcId(aJsonRpc) <> GetJsonRpcId(PostJson) Then
        errSrc = "OdJsonRpc.PostJson"
        errDsc = "Invalid JSON-RPC ID." & vbCrLf & "Expected: " & GetJsonRpcId(aJsonRpc) & vbCrLf & "Actual: " & GetJsonRpcId(PostJson)
        LogError errDsc, errSrc, CERR_JSONRPC_ID
        Err.Raise CERR_JSONRPC_ID, errSrc, errDsc
    End If
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
        LogError errDsc, errSrc, CERR_AUTHENTICATE
        Err.Raise CERR_AUTHENTICATE, errSrc, errDsc
    End If
    Set JsonRpcCommonAuthenticate = dic ' Type of json("result") is Long.
End Function

Public Function JsonRpcObjectExecuteKw(aOdConnection As OdConnection, aModelName As String, aMethodName As String, Optional aParams As Variant = "", Optional aOptions As Variant = "") As Dictionary
    Dim args As New Collection
    With args
        .Add aOdConnection.DbName    ' the database to use, a string
        .Add aOdConnection.UserId    ' the user id (retrieved through authenticate), an integer
        .Add aOdConnection.Password  ' the user�fs password, a string
        
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

' ---------
'  XML-RPC
' ---------
Public Function CreatePostXmlWebRequest(Url As String, Body As Variant, Optional Options As Dictionary) As WebRequest
    ' https://github.com/VBA-tools/VBA-Web/wiki/XML-Support-in-4.0
    
    ' Register XML Converter
    If Not mRegisteredXml Then
        WebHelpers.RegisterConverter "xml", "application/xml", "OdRpc.ConvertToXml", "OdRpc.ParseXml"
        mRegisteredXml = True
    End If
    
    Set CreatePostXmlWebRequest = New WebRequest
    With CreatePostXmlWebRequest
        .Method = WebMethod.HttpPost
        .Resource = Url
        If VBA.IsObject(Body) Then
            Set .Body = Body
        Else
            .Body = Body
        End If
        .CreateFromOptions Options
        ' Use XML Converter in WebRequest
        .Format = WebFormat.Custom
        .CustomRequestFormat = "xml"
        .CustomResponseFormat = "xml"
    End With
    
End Function

Public Function ParseXml(Encoded As String) As Object ' MSXML2.DOMDocument
    ' https://github.com/VBA-tools/VBA-Web/wiki/XML-Support-in-4.0
    Set ParseXml = CreateObject("MSXML2.DOMDocument")
    ParseXml.Async = False
    ParseXml.LoadXML Encoded
End Function
 
Public Function ConvertToXml(Obj As Object) As String
    ' https://github.com/VBA-tools/VBA-Web/wiki/XML-Support-in-4.0
    ConvertToXml = Trim(Replace(Obj.Xml, vbCrLf, ""))
End Function

Public Function PostXml(aOdConnection As OdConnection, aUrlPath As String, aBody As Variant, Optional aOptions As Dictionary) As Object    ' MSXML2.DOMDocument
    ' https://github.com/VBA-tools/VBA-Web/wiki/XML-Support-in-4.0
    Dim sUrl As String
    Dim web_Request As WebRequest
    Dim web_Response As WebResponse
    Dim errSrc As String
    Dim errDsc As String
    
    sUrl = WebHelpers.JoinUrl(aOdConnection.BaseUrl, aUrlPath)
    Set web_Request = CreatePostXmlWebRequest(sUrl, aBody, aOptions)
    Set web_Response = aOdConnection.RefWebClient.Execute(web_Request)
    
    If web_Response.StatusCode <> WebStatusCode.Ok Then
        errSrc = sUrl
        errDsc = "web response error (status code: " & web_Response.StatusCode & " )" & vbCrLf & sUrl
        LogError errDsc, errSrc, CERR_STATUSCODE
        Err.Raise CERR_STATUSCODE, errSrc, errDsc
    End If
    
    Set PostXml = ParseXml(web_Response.Content)
End Function

Public Function PostXmlStart(aOdConnection As OdConnection, aXmlBody As Variant) As Object    ' MSXML2.DOMDocument
    Set PostXmlStart = PostXml(aOdConnection, "start", aXmlBody)
End Function

Public Function CreateXmlBody(aMethodName As String, Optional aParams As Variant = Nothing) As Object   ' MSXML2.DOMDocument
    Dim xmlDoc As Object
    Dim methodCall As Object
    Dim methodName As Object
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    
    ' methodCall
    Set methodCall = xmlDoc.createElement("methodCall")
    xmlDoc.appendChild methodCall
    
    ' methodCall > methodName
    Set methodName = xmlDoc.createElement("methodName")
    methodName.Text = aMethodName
    methodCall.appendChild methodName

    Set CreateXmlBody = xmlDoc
End Function

Public Function TestDatabase(aOdConnection As OdConnection) As Dictionary
    Dim xmlDoc As Object
    Dim dicResult As New Dictionary
    Set xmlDoc = PostXmlStart(aOdConnection, CreateXmlBody("start"))
    Dim memberNode As Object
    For Each memberNode In xmlDoc.SelectNodes("//member")
        dicResult.Add memberNode.SelectSingleNode("name").Text, memberNode.SelectSingleNode("value/string").Text
    Next
    Set TestDatabase = New Dictionary
    TestDatabase.Add "result", dicResult
End Function

