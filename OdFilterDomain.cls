VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OdFilterDomain"
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

'
' Search domains, ORM API - odoo docs
'
' A domain is a list of criteria, each criterion being a triple (either a list or a tuple)
' of (field_name, operator, value)
'
' see also: https://www.odoo.com/documentation/18.0/developer/reference/backend/orm.html#search-domains
'

Option Explicit

Private mList As New Collection ' list of criteria or combinations or number or string

Public Function ToFilterString() As String
    Dim a() As String
    Dim i As Integer
    
    If mList.Count = 0 Then
        ToFilterString = "[]"
        GoTo ExitProc
    End If
    
    ReDim a(1 To mList.Count)
    For i = LBound(a) To UBound(a)
        a(i) = VarToFilterString(mList(i))
    Next i
    
    ToFilterString = "[" & Join(a, ", ") & "]"
    
ExitProc:
    Exit Function
End Function

Private Function VarToFilterString(v As Variant) As String
On Error Resume Next
    Dim criteria As OdFilterCriteria
    Dim combi As OdFilterCombi
    
    Set criteria = v
    If Not criteria Is Nothing Then
        VarToFilterString = criteria.ToFilterString()
        GoTo ExitProc
    End If
    Set combi = v
    If Not combi Is Nothing Then
        VarToFilterString = combi.ToFilterString()
        GoTo ExitProc
    End If
    If VarType(v) = vbString Then
        VarToFilterString = "'" & v & "'"
        GoTo ExitProc
    End If
    VarToFilterString = v
ExitProc:
    Exit Function
End Function

Public Function AddArity(aArity As Object) As OdFilterDomain
    mList.Add aArity
    Set AddArity = Me
End Function

