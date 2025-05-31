Attribute VB_Name = "OdXmlRpc"
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

Private mRegisteredXml As Boolean

Public Function CreatePostXmlWebRequest(Url As String, Body As Variant, Optional Options As Dictionary) As WebRequest
    ' https://github.com/VBA-tools/VBA-Web/wiki/XML-Support-in-4.0
    
    ' Register XML Converter
    If Not mRegisteredXml Then
        WebHelpers.RegisterConverter "xml", "application/xml", "OdXmlRpc.ConvertToXml", "OdXmlRpc.ParseXml"
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

Public Function ParseXml(Value As String) As Object ' MSXML2.DOMDocument
    ' https://github.com/VBA-tools/VBA-Web/wiki/XML-Support-in-4.0
    Set ParseXml = CreateObject("MSXML2.DOMDocument")
    ParseXml.Async = False
    ParseXml.LoadXML Value
End Function
 
Public Function ConvertToXml(Value As Object) As String
    ' https://github.com/VBA-tools/VBA-Web/wiki/XML-Support-in-4.0
    ConvertToXml = Trim(Replace(Value.Xml, vbCrLf, ""))
End Function

Public Function PostXml(aOdService As odService, aUrlPath As String, aBody As Variant, Optional aOptions As Dictionary) As Object    ' MSXML2.DOMDocument
    ' https://github.com/VBA-tools/VBA-Web/wiki/XML-Support-in-4.0
    Dim sUrl As String
    Dim web_Request As WebRequest
    Dim web_Response As WebResponse
    Dim errSrc As String
    Dim errDsc As String
    
    sUrl = WebHelpers.JoinUrl(aOdService.BaseUrl, aUrlPath)
    Set web_Request = CreatePostXmlWebRequest(sUrl, aBody, aOptions)
    Set web_Response = aOdService.RefWebClient.Execute(web_Request)
    
    If web_Response.StatusCode <> WebStatusCode.Ok Then
        errSrc = sUrl
        errDsc = "web response error (status code: " & web_Response.StatusCode & " )" & vbCrLf & sUrl
        LogError errDsc, errSrc, OdRpc.CERR_STATUSCODE
        Err.Raise OdRpc.CERR_STATUSCODE, errSrc, errDsc
    End If
    
    Debug.Print web_Response.Content
    Set PostXml = OdXmlRpc.ParseXml(web_Response.Content)
    
End Function

Public Function PostXmlStart(aOdService As odService, aXmlBody As Variant) As Object    ' MSXML2.DOMDocument
    Set PostXmlStart = PostXml(aOdService, "start", aXmlBody)
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
    
    Debug.Print xmlDoc.Xml
End Function

Public Function TestDatabase(aOdService As odService) As Dictionary
    Dim xmlDoc As Object
    Dim Result As New Dictionary
    Set xmlDoc = PostXmlStart(aOdService, CreateXmlBody("start"))
    Dim memberNode As Object
    For Each memberNode In xmlDoc.SelectNodes("//member")
        Result.Add memberNode.SelectSingleNode("name").Text, memberNode.SelectSingleNode("value/string").Text
    Next
    Set TestDatabase = New Dictionary
    TestDatabase.Add "result", Result
End Function
