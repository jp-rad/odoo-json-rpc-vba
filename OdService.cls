VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OdService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

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
        aUsername:="admin", _
        aPassword:="admin"
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

Public Sub SetInsecure(aValue As Boolean)
    mWebClient.Insecure = aValue
End Sub

Public Sub SetFollowRedirects(aValue As Boolean)
    mWebClient.FollowRedirects = aValue
End Sub

Public Property Get RefWebClient() As WebClient
    Set RefWebClient = mWebClient
End Property
