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
        Set Version = .Wv(JsonRpcCommonVersion(mOdConnection))
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
        Set Authenticate = .W(JsonRpcCommonAuthenticate(mOdConnection))
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
        Set ExecuteKw = .W(JsonRpcObjectExecuteKw(mOdConnection, mModelName, mMethodName, aParams, aNamedParams))
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
    With New OdResult
        Set MethodExecuteKw = Me.Method(aMethodName).ExecuteKw(aParams, aNamedParams)
    End With
End Function

