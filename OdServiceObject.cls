VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OdServiceObject"
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
' Calling methods:
' https://www.odoo.com/documentation/15.0/developer/misc/api/odoo.html#calling-methods
'

Option Explicit

Private mClient As OdWebClient
Private mModelName As String

Private Function WL(v As Variant) As Variant
    Dim r As Collection
    If IsObject(v) Then
        Set r = New Collection
        r.Add v
        Set WL = r
    Else
        If Len(v) > 0 Then
            WL = "[" & v & "]"
        Else
            WL = "[[]]"
        End If
    End If
End Function

Public Function AssignObject(aClient As OdWebClient, aModelName As String) As OdServiceObject
    Set AssignObject = Me
    Set mClient = aClient
    mModelName = aModelName
End Function

Public Function ExecuteKw(aMethodName As String, aListParam As Variant, Optional aDictParam As Variant = "") As OdResult
    With New OdResult
        Set ExecuteKw = .W(OdooJsonRpc.ObjectExecuteKw(mClient.BaseUrl, mClient.DbName, mClient.UserId, mClient.Password, mModelName, aMethodName, aListParam, aDictParam))
    End With
End Function

Public Function MethodCheckAccessRights(aRights As Variant, aParameters As Variant) As OdResult
    Set MethodCheckAccessRights = Me.ExecuteKw(CMETHOD_CHECK_ACCESS_RIGHTS, aRights, aParameters)
End Function

Public Function MethodSearch(aDomain As Variant, Optional aParameters As Variant = "") As OdResult
    Set MethodSearch = Me.ExecuteKw(CMETHOD_SEARCH, WL(aDomain), aParameters)
End Function

Public Function MethodRead(aIds As Variant, Optional aParameters As Variant = "") As OdResult
    Set MethodRead = Me.ExecuteKw(CMETHOD_READ, WL(aIds), aParameters)
End Function

Public Function MethodSearchAndRead(aDomain As Variant, Optional aParameters As Variant = "") As OdResult
    Set MethodSearchAndRead = Me.ExecuteKw(CMETHOD_SEARCH_READ, WL(aDomain), aParameters)
End Function

