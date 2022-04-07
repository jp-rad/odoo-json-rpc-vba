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

Private mBaseUrl As String      ' Base URL (HOST)
Private mDbName As String       ' the database to use, a string
Private mUsername As String     ' the user name, a string
Private mPassword As String     ' the user's password, a string

Private mUserId As Long         ' the user id (retrieved through authenticate), an integer

Private Sub Class_Initialize()
    mUserId = 0
End Sub

Public Sub SetConnectionInfo(Optional aBaseUrl As String = "", Optional aDbName As String = "", Optional aUsername As String, Optional aPassword As String = "")
    Dim blnChanged As Boolean
    blnChanged = False
    If aBaseUrl <> "" Then
        If mBaseUrl <> aBaseUrl Then
            mBaseUrl = aBaseUrl
            blnChanged = True
        End If
    End If
    If aDbName <> "" Then
        If mDbName <> aDbName Then
            mDbName = aDbName
            blnChanged = True
        End If
    End If
    If aUsername <> "" Then
        If mUsername <> aUsername Then
            mUsername = aUsername
            blnChanged = True
        End If
    End If
    If aPassword <> "" Then
        If mPassword <> aPassword Then
            mPassword = aPassword
            blnChanged = True
        End If
    End If
    If blnChanged Then
        mUserId = 0
    End If
End Sub

Public Property Get BaseUrl() As String
    BaseUrl = mBaseUrl
End Property

Public Property Let BaseUrl(newValue As String)
    SetConnectionInfo aBaseUrl:=newValue
End Property

Public Property Get DbName() As String
    DbName = mDbName
End Property

Public Property Let DbName(newValue As String)
    SetConnectionInfo aDbName:=newValue
End Property

Public Property Get Username() As String
    Username = mUsername
End Property

Public Property Let Username(newValue As String)
    SetConnectionInfo aUsername:=newValue
End Property

Public Property Get Password() As String
    Password = mPassword
End Property

Public Property Let Password(newValue As String)
    SetConnectionInfo aPassword:=newValue
End Property

Public Property Get UserId() As Long
    UserId = mUserId
End Property

Public Property Let UserId(newValue As Long)
    If newValue > 0 Then
        mUserId = newValue
    Else
        mUserId = 0
    End If
End Property

Public Function IsAuthenticated() As Boolean
    IsAuthenticated = mUserId > 0
End Function

' ======================
'  Start
' ======================
' Test database:
' https://www.odoo.com/documentation/15.0/developer/misc/api/odoo.html#test-database
Public Property Get Start(Optional aDemoUrl As String = "") As OdServiceStart
    If aDemoUrl = "" Then
        aDemoUrl = OdooJsonRpc.CURL_DEMO
    End If
    With New OdServiceStart
        Set Start = .AssignStart(Me, aDemoUrl)
    End With
End Property

' ======================
'  Common
' ======================
' Logging in:
' https://www.odoo.com/documentation/15.0/developer/misc/api/odoo.html#logging-in
Public Property Get Common(Optional aBaseUrl As String = "") As OdServiceCommon
    SetConnectionInfo aBaseUrl:=aBaseUrl
    With New OdServiceCommon
        Set Common = .AssignCommon(Me)
    End With
End Property

' ======================
'  Model (Object)
' ======================
' Calling methods:
' https://www.odoo.com/documentation/15.0/developer/misc/api/odoo.html#calling-methods
Public Property Get Model(aModelName As String) As OdServiceObject
    With New OdServiceObject
        Set Model = .AssignObject(Me, aModelName)
    End With
End Property

' ======================================
'  Model of 'ir.model' (Object)
' ======================================
' Inspection and introspection:
' https://www.odoo.com/documentation/15.0/developer/misc/api/odoo.html#inspection-and-introspection
Public Property Get ModelOfIrModel() As OdServiceObject
    Set ModelOfIrModel = Me.Model("ir.model")
End Property

' ======================================
'  Model of 'ir.model.fields' (Object)
' ======================================
' Inspection and introspection:
' https://www.odoo.com/documentation/15.0/developer/misc/api/odoo.html#inspection-and-introspection
Public Property Get ModelOfIrModelFields() As OdServiceObject
    Set ModelOfIrModel = Me.Model("ir.model.fields")
End Property

