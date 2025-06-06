VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OdFilterCombi"
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

Private Const CLOGIC_NOT As String = "'!'"
Private Const CLOGIC_AND As String = "'&'"
Private Const CLOGIC_OR As String = "'|'"

Private mLogic As String
Private mArityList As Collection

Public Function ToJson() As String
    Dim s() As String
    Dim cnt As Long
    Dim i As Long
    Dim v As Variant
    
    If mLogic = CLOGIC_NOT Then
        ReDim s(0 To 1)
        s(0) = CLOGIC_NOT
        s(1) = mArityList(1).ToJson()
    Else
        cnt = mArityList.Count
        ReDim s(0 To (cnt - 1) * 2)
        For i = 0 To cnt - 2
            s(i) = mLogic
        Next i
        For Each v In mArityList
            s(i) = v.ToJson()
            i = i + 1
        Next v
    End If
    ToJson = Join(s, ", ")
End Function

Private Function SetLogic(aLogic As String, aArity1 As Object, aArity2 As Object) As OdFilterCombi
    Set SetLogic = Me
    mLogic = aLogic
    Set mArityList = New Collection
    mArityList.Add aArity1
    If mLogic <> CLOGIC_NOT Then
        mArityList.Add aArity2
    End If
End Function

Public Function SetAndLogic(aArity1 As Object, aArity2 As Object) As OdFilterCombi
    Set SetAndLogic = SetLogic(CLOGIC_AND, aArity1, aArity2)
End Function

Public Function SetOrLogic(aArity1 As Object, aArity2 As Object) As OdFilterCombi
    Set SetOrLogic = SetLogic(CLOGIC_OR, aArity1, aArity2)
End Function

Public Function SetNotLogic(aArity As Object) As OdFilterCombi
    Set SetNotLogic = SetLogic(CLOGIC_NOT, aArity, Nothing)
End Function

Public Function Add(aArity As Object) As OdFilterCombi
    Set Add = Me
    mArityList.Add aArity
End Function
