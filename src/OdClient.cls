VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OdClient"
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

Private Const CURL_TEST_DATABASE As String = "https://demo.odoo.com"

Private mOdConnection As New OdConnection

Public Property Get BaseUrl() As String
    BaseUrl = mOdConnection.BaseUrl
End Property

Public Property Let BaseUrl(newValue As String)
    mOdConnection.SetConnectionInfo aBaseUrl:=newValue
End Property

Public Property Get DbName() As String
    DbName = mOdConnection.DbName
End Property

Public Property Let DbName(newValue As String)
    mOdConnection.SetConnectionInfo aDbName:=newValue
End Property

Public Property Get Username() As String
    Username = mOdConnection.Username
End Property

Public Property Let Username(newValue As String)
    mOdConnection.SetConnectionInfo aUserName:=newValue
End Property

Public Property Get Password() As String
    Password = mOdConnection.Password
End Property

Public Property Let Password(newValue As String)
    mOdConnection.SetConnectionInfo aPassword:=newValue
End Property

Public Property Get UserId() As Long
    UserId = mOdConnection.UserId
End Property

Public Function IsAuthenticated() As Boolean
    IsAuthenticated = mOdConnection.IsAuthenticated
End Function

Public Property Get RefWebClient() As WebClient
    Set RefWebClient = mOdConnection.RefWebClient
End Property


' ======================
'  Start
' ======================
' Test database:
' https://www.odoo.com/documentation/master/developer/reference/external_api.html#test-database
Public Function StartTestDatabase() As OdResult
    Dim connDemo As New OdConnection
    connDemo.BaseUrl = CURL_TEST_DATABASE
    With New OdResult
        Set StartTestDatabase = .Ws(TestDatabase(connDemo))
    End With
End Function

' ======================
'  Common
' ======================
' Logging in:
' https://www.odoo.com/documentation/master/developer/reference/external_api.html#logging-in
Public Property Get Common() As OdService
    With New OdService
        Set Common = .Assign(mOdConnection)
    End With
End Property

Public Function CommonVersion() As OdResult
    Set CommonVersion = Me.Common.Version()
End Function

Public Function CommonAuthenticate() As OdResult
    Set CommonAuthenticate = Me.Common.Authenticate()
End Function

' ======================
'  Model (Object)
' ======================
' Calling methods:
' https://www.odoo.com/documentation/master/developer/reference/external_api.html#calling-methods
Public Property Get Model(aModelName As String) As OdService
    With New OdService
        Set Model = .Assign(mOdConnection, aModelName)
    End With
End Property

Public Function ModelExecuteKw(aModelName As String, aMethodName As String, Optional aParams, Optional aNamedParams) As OdResult
    With New OdResult
        Set ModelExecuteKw = Me.Model(aModelName).MethodExecuteKw(aMethodName, aParams, aNamedParams)
    End With
End Function

' ======================================
'  'ir.model' (Object)
' ======================================
Public Property Get ModelOfIrModel() As OdService
    Set ModelOfIrModel = Me.Model("ir.model")
End Property

' ======================================
'  'ir.model.access' (Object)
' ======================================
Public Property Get ModelOfIrModelAccess() As OdService
    Set ModelOfIrModelAccess = Me.Model("ir.model.access")
End Property

' ======================================
'  'ir.model.fields' (Object)
' ======================================
Public Property Get ModelOfIrModelFields() As OdService
    Set ModelOfIrModelFields = Me.Model("ir.model.fields")
End Property

