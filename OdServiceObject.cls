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
' Calling methods:
' https://www.odoo.com/documentation/master/developer/misc/api/odoo.html#calling-methods
'

Option Explicit

Private mOdService As odService
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

Private Function WL2(n As Long, Optional v As Variant = "") As Variant
    Dim r As Collection
    If IsObject(v) Then
        Set r = New Collection
        r.Add New Collection
        r.Item(1).Add n
        r.Add v
        Set WL2 = r
    Else
        If Len(v) > 0 Then
            WL2 = "[[" & n & "], " & v & "]"
        Else
            WL2 = "[[" & n & "]]"
        End If
    End If
End Function

Public Function Assign(aOdService As odService, aModelName As String) As OdServiceObject
    Set Assign = Me
    Set mOdService = aOdService
    mModelName = aModelName
End Function

Public Function ExecuteKw(aMethodName As String, aListParam As Variant, Optional aDictParam As Variant = "") As OdResult
    With New OdResult
        Set ExecuteKw = .W(OdJsonRpc.JsonRpcObjectExecuteKw(mOdService, mModelName, aMethodName, aListParam, aDictParam))
    End With
End Function

Public Function MethodCheckAccessRights(aRights As Variant, aParameters As Variant) As OdResult
    Set MethodCheckAccessRights = Me.ExecuteKw(OdRpc.CMETHOD_CHECK_ACCESS_RIGHTS, aRights, aParameters)
End Function

Public Function MethodSearch(aDomain As Variant, Optional aParameters As Variant = "") As OdResult
    Set MethodSearch = Me.ExecuteKw(OdRpc.CMETHOD_SEARCH, WL(aDomain), aParameters)
End Function

Public Function MethodRead(aIds As Variant, Optional aParameters As Variant = "") As OdResult
    Set MethodRead = Me.ExecuteKw(OdRpc.CMETHOD_READ, WL(aIds), aParameters)
End Function

Public Function MethodSearchAndRead(aDomain As Variant, Optional aParameters As Variant = "") As OdResult
    Set MethodSearchAndRead = Me.ExecuteKw(OdRpc.CMETHOD_SEARCH_READ, WL(aDomain), aParameters)
End Function

Public Function MethodCreate(aValues As Variant) As OdResult
    Set MethodCreate = Me.ExecuteKw(OdRpc.CMETHOD_CREATE, WL(aValues))
End Function

Public Function MethodWrite(aId As Long, aValues As Variant) As OdResult
    Set MethodWrite = Me.ExecuteKw(OdRpc.CMETHOD_WRITE, WL2(aId, aValues))
End Function

Public Function MethodNameGet(aId As Long) As OdResult
    Set MethodNameGet = Me.ExecuteKw(OdRpc.CMETHOD_NAME_GET, WL(WL(aId)))
End Function

Public Function MethodUnlink(aId As Long) As OdResult
    Set MethodUnlink = Me.ExecuteKw(OdRpc.CMETHOD_UNLINK, WL(WL(aId)))
End Function
