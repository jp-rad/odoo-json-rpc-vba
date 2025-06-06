VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OdConnection"
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

Private mBaseUrl As String      ' Base URL (HOST)
Private mDbName As String       ' the database to use, a string
Private mUsername As String     ' the user name, a string
Private mPassword As String     ' the user's password, a string

Private mUserId As Long         ' the user id (retrieved through authenticate), an integer

Private mWebClient As New WebClient

Private Sub Class_Initialize()
    SetConnectionInfo _
        aBaseUrl:="https://localhost", _
        aDbName:="odoo", _
        aUserName:="admin", _
        aPassword:="admin"
End Sub

Public Sub SetConnectionInfo(Optional aBaseUrl As String = "", Optional aDbName As String = "", Optional aUserName As String, Optional aPassword As String = "")
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
    If aUserName <> "" Then
        If mUsername <> aUserName Then
            mUsername = aUserName
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
    SetConnectionInfo aUserName:=newValue
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

Public Sub SetInsecure(aInsecure As Boolean)
    mWebClient.Insecure = aInsecure
End Sub

Public Sub SetFollowRedirects(aFollowRedirects As Boolean)
    mWebClient.FollowRedirects = aFollowRedirects
End Sub

Public Property Get RefWebClient() As WebClient
    Set RefWebClient = mWebClient
End Property
