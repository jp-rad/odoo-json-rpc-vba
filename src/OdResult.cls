VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OdResult"
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

' result value
Private mResult As Variant

' Result of start.start
Public sHost As String
Public sDatabase As String
Public sUser As String
Public sPassword As String

' Result of common.version
Public vServerVersion As String
Public vServerVersionInfo As Collection
Public vServerSerie As String
Public vProtocolVersion As Long

Public Property Get Result() As Variant
    If IsObject(mResult) Then
        Set Result = mResult
    Else
        Result = mResult
    End If
End Property

Public Function W(dic As Dictionary) As OdResult
    Set W = Me
    If IsObject(dic("result")) Then
        Set mResult = dic("result")
    Else
        mResult = dic("result")
    End If
End Function

' Test Database:
' https://www.odoo.com/documentation/master/developer/misc/api/odoo.html#test-database
Public Function Ws(dic As Dictionary) As OdResult
    Set Ws = Me
    Set mResult = dic("result")
    ' "result":
    ' {
    '     "host":     "https://demo3.odoo.com",
    '     "database": "demo_150_1648963374",
    '     "user":     "admin",
    '     "password": "admin"
    ' }
    With mResult
        Me.sHost = .Item("host")
        Me.sDatabase = .Item("database")
        Me.sUser = .Item("user")
        Me.sPassword = .Item("password")
    End With
End Function

' version:
' https://www.odoo.com/documentation/master/developer/misc/api/odoo.html#logging-in
Public Function Wv(dic As Dictionary) As OdResult
    Dim v As Variant
    
    Set Wv = Me
    Set mResult = dic("result")
    
    ' "result":
    ' {
    '     "server_version": "15.0+e",
    '     "server_version_info": [
    '             15,
    '             0,
    '             0,
    '             "final",
    '             0,
    '             "e"
    '         ],
    '     "server_serie": "15.0",
    '     "protocol_version": 1
    ' }
    With mResult
        Me.vServerVersion = .Item("server_version")
        Set Me.vServerVersionInfo = New Collection
        For Each v In .Item("server_version_info")
            Me.vServerVersionInfo.Add v
        Next v
        Me.vServerSerie = .Item("server_serie")
        Me.vProtocolVersion = .Item("protocol_version")
    End With
End Function
