VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OdService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' External API - odoo-JSON-RPC-VBA
'
' MIT License
'
' Copyright (c) 2022-2026 jp-rad
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

Private mOdConnection As OdConnection
Private mModelName As String
Private mMethodName As String

Public Function Assign(aOdConnection As OdConnection, Optional aModelName As String = "", Optional aMethodName As String = "") As OdService
    Set Assign = Me
    Set mOdConnection = aOdConnection
    mModelName = aModelName
    mMethodName = aMethodName
End Function

' ======================
'  Common - Version
' ======================
' Logging in:
' https://www.odoo.com/documentation/master/developer/reference/external_api.html#logging-in
Public Function Version() As OdResult
    With New OdResult
        Set Version = .Wv(mOdConnection, JsonRpcCommonVersion(mOdConnection))
    End With
End Function

' ======================
'  Common - Authenticate
' ======================
' Logging in:
' https://www.odoo.com/documentation/master/developer/reference/external_api.html#logging-in
Public Function Authenticate(Optional aDbName As String = "", Optional aUserName As String, Optional aPassword As String = "") As OdResult
    mOdConnection.SetDbUserInfo aDbName:=aDbName, aUserName:=aUserName, aPassword:=aPassword
    With New OdResult
        Set Authenticate = .Wa(mOdConnection, JsonRpcCommonAuthenticate(mOdConnection))
        mOdConnection.SetUserId .Result
    End With
End Function

' ======================
'  ExcecuteKw
' ======================
' Calling methods:
' https://www.odoo.com/documentation/master/developer/reference/external_api.html#calling-methods
Public Function ExecuteKw(Optional aParams, Optional aNamedParams) As OdResult
    Debug.Assert mModelName <> ""
    Debug.Assert mMethodName <> ""
    With New OdResult
        Set ExecuteKw = .W(mOdConnection, JsonRpcObjectExecuteKw(mOdConnection, mModelName, mMethodName, aParams, aNamedParams), mModelName)
    End With
End Function

' ======================
'  Model Method
' ======================
' Calling methods:
' https://www.odoo.com/documentation/master/developer/reference/external_api.html#calling-methods
Public Property Get Method(aMethodName As String) As OdService
    With New OdService
        Set Method = .Assign(mOdConnection, mModelName, aMethodName)
    End With
End Property

Public Function MethodExecuteKw(aMethodName As String, Optional aParams, Optional aNamedParams) As OdResult
    Debug.Assert mModelName <> ""
    Set MethodExecuteKw = Me.Method(aMethodName).ExecuteKw(aParams, aNamedParams)
End Function

' ======================
'  Syntax Sugar
' ======================
Public Function MethodSearch(aDomain As OdFilterDomain, Optional aModifiers) As OdResult
    Dim params As Collection
    Set params = NewList
    aDomain.BuildAndAppend params
    Set MethodSearch = MethodExecuteKw("search", params, aModifiers)
End Function

Public Function MethodRead(aIds As Variant, aFieldNames As Collection) As OdResult
    Dim params As Collection
    Dim named As Dictionary
    Set params = NewList
    params.Add aIds
    Set named = NewDict
    Set named("fields") = aFieldNames
    Set MethodRead = MethodExecuteKw("read", params, named)
End Function

Public Function MethodSearchRead(aDomain As OdFilterDomain, aFieldNames As Collection, Optional aModifiers As Dictionary = Nothing) As OdResult
    Dim params As Collection
    Dim named As Dictionary
    Set params = NewList
    aDomain.BuildAndAppend params
    If aModifiers Is Nothing Then
        Set named = NewDict
    Else
        Set named = aModifiers
    End If
    Set named.Item("fields") = aFieldNames
    Set MethodSearchRead = MethodExecuteKw("search_read", params, named)
End Function

Public Function MethodFieldsGet(Optional aFieldNames As Collection = Nothing, Optional aAttributes As Collection = Nothing) As OdResult
    Dim params As Collection
    Dim named As Dictionary
    Set params = NewList
    If Not aFieldNames Is Nothing Then
        params.Add aFieldNames
    End If
    Set named = NewDict
    If Not aAttributes Is Nothing Then
        Set named("attributes") = aAttributes
    End If
    Set MethodFieldsGet = MethodExecuteKw("fields_get", params, named)
End Function

Private Function GetFieldInfo(aFieldName As String) As Dictionary
    
    Dim f As Collection
    Dim attr As Collection
    Dim oRslt As OdResult
    
    Dim dic As Dictionary
        
    If Not mOdConnection.HasFieldsCache(mModelName) Then
        Set f = NewList()
        Set attr = NewList()
        attr.Add "type"
        attr.Add "relation"
        Set oRslt = MethodFieldsGet(f, attr)
        mOdConnection.SetFieldsCache oRslt.ModelName, oRslt.Result
    End If
    
    Set dic = mOdConnection.GetFieldsCache(mModelName)
    If IsObject(dic(aFieldName)) Then
        Set GetFieldInfo = dic(aFieldName)
    Else
        Set GetFieldInfo = New Dictionary
    End If
    
End Function

Public Function FieldType(aFieldName As String) As String
    Dim dic As Dictionary
    Set dic = GetFieldInfo(aFieldName)
    FieldType = dic("type")
End Function

Public Function FieldRelation(aFieldName As String) As String
    Dim dic As Dictionary
    Set dic = GetFieldInfo(aFieldName)
    FieldRelation = dic("relation")
End Function
