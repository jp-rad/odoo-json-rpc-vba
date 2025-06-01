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

Option Explicit

'
' Search domains - ORM API
'
' A search domain is a first-order logical predicate used for filtering and searching recordsets.
' You combine simple conditions on a field expression with logical operators.
'
' https://www.odoo.com/documentation/master/developer/reference/backend/orm.html#search-domains
'

Private mList As New Collection ' list of criteria or combinations

Public Function AddArity(aArity As Object) As OdFilterDomain
    Set AddArity = Me
    mList.Add aArity
End Function

Private Function VarToJson(v As Object) As String
On Error Resume Next
    Dim criteria As OdFilterCriteria
    Set criteria = v
    If Not criteria Is Nothing Then
        VarToJson = criteria.ToJson()
        GoTo ExitProc
    End If
    Dim combi As OdFilterCombi
    Set combi = v
    If Not combi Is Nothing Then
        VarToJson = combi.ToJson()
        GoTo ExitProc
    End If
On Error GoTo ErrRaiseHandler
    VarToJson = v.ToJson()
ExitProc:
    Exit Function
ErrRaiseHandler:
    Err.Raise Err.Number, Err.Source, Err.Description & vbCrLf & "TypeName:" & TypeName(v)
End Function

Public Function ToJson() As String
    Dim a() As String
    Dim i As Integer
    
    If mList.Count = 0 Then
        ToJson = "[]"
        GoTo ExitProc
    End If
    
    ReDim a(1 To mList.Count)
    For i = LBound(a) To UBound(a)
        a(i) = VarToJson(mList(i))
    Next i
    
    ToJson = "[" & Join(a, ", ") & "]"
    
ExitProc:
    Exit Function
End Function

Public Function Build() As Collection
    Set Build = JsonConverter.ParseJson(Me.ToJson())
End Function

Public Function BuildAndAppend(aList As Collection) As Collection
    Set BuildAndAppend = Build
    aList.Add BuildAndAppend
End Function
