VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OdWebClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
' see also: https://www.odoo.com/documentation/master/developer/misc/api/odoo.html
'

Option Explicit

Private mOdService As New odService

Public Property Get BaseUrl() As String
    BaseUrl = mOdService.BaseUrl
End Property

Public Property Let BaseUrl(newValue As String)
    mOdService.SetConnectionInfo aBaseUrl:=newValue
End Property

Public Property Get DbName() As String
    DbName = mOdService.DbName
End Property

Public Property Let DbName(newValue As String)
    mOdService.SetConnectionInfo aDbName:=newValue
End Property

Public Property Get Username() As String
    Username = mOdService.Username
End Property

Public Property Let Username(newValue As String)
    mOdService.SetConnectionInfo aUsername:=newValue
End Property

Public Property Get Password() As String
    Password = mOdService.Password
End Property

Public Property Let Password(newValue As String)
    mOdService.SetConnectionInfo aPassword:=newValue
End Property

Public Property Get UserId() As Long
    UserId = mOdService.UserId
End Property

Public Function IsAuthenticated() As Boolean
    IsAuthenticated = mOdService.IsAuthenticated
End Function

Public Sub SetInsecure(aValue As Boolean)
    mOdService.SetInsecure aValue
End Sub

Public Sub SetFollowRedirects(aValue As Boolean)
    mOdService.SetFollowRedirects aValue
End Sub

' ======================
'  Start
' ======================
' Test database:
' https://www.odoo.com/documentation/master/developer/misc/api/odoo.html#test-database
Public Function TestDatabase() As OdResult
    Dim odService As New odService
    Dim dic As Dictionary
    odService.BaseUrl = OdRpc.CURL_DEMO
    With New OdResult
        Set TestDatabase = .Ws(OdXmlRpc.TestDatabase(odService))
    End With
End Function

' ======================
'  Common
' ======================
' Logging in:
' https://www.odoo.com/documentation/master/developer/misc/api/odoo.html#logging-in
Public Property Get Common() As OdServiceCommon
    With New OdServiceCommon
        Set Common = .Assign(mOdService)
    End With
End Property

' ======================
'  Model (Object)
' ======================
' Calling methods:
' https://www.odoo.com/documentation/master/developer/misc/api/odoo.html#calling-methods
Public Property Get Model(aModelName As String) As OdServiceObject
    With New OdServiceObject
        Set Model = .Assign(mOdService, aModelName)
    End With
End Property

' ======================================
'  Model of 'ir.model' (Object)
' ======================================
' Inspection and introspection:
' https://www.odoo.com/documentation/master/developer/misc/api/odoo.html#inspection-and-introspection
Public Property Get ModelOfIrModel() As OdServiceObject
    Set ModelOfIrModel = Me.Model("ir.model")
End Property

' ======================================
'  Model of 'ir.model.fields' (Object)
' ======================================
' Inspection and introspection:
' https://www.odoo.com/documentation/master/developer/misc/api/odoo.html#inspection-and-introspection
Public Property Get ModelOfIrModelFields() As OdServiceObject
    Set ModelOfIrModelFields = Me.Model("ir.model.fields")
End Property
