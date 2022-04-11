VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OdServiceCommon"
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
' Logging in:
' https://www.odoo.com/documentation/15.0/developer/misc/api/odoo.html#logging-in
'

Option Explicit

Private mClient As OdWebClient

Public Function AssignCommon(aClient As OdWebClient) As OdServiceCommon
    Set AssignCommon = Me
    Set mClient = aClient
End Function

' Version
Public Function Version() As OdResult
    With New OdResult
        Set Version = .Wv(OdooJsonRpc.CommonVersion(mClient.BaseUrl))
    End With
End Function

' Authenticate
Public Function Authenticate(Optional aDbName As String = "", Optional aUsername As String = "", Optional aPassword As String = "") As OdResult
    mClient.SetConnectionInfo aDbName:=aDbName, aUsername:=aUsername, aPassword:=aPassword
    With New OdResult
        Set Authenticate = .W(OdooJsonRpc.CommonAuthenticate(mClient.BaseUrl, mClient.DbName, mClient.Username, mClient.Password))
        mClient.UserId = .JsonResult
    End With
End Function

