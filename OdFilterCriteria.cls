VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OdFilterCriteria"
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

Private mFieldExpr As String
Private mOperator As String
Private mValue As Variant   ' string, number, boolean, null, domain

Private Function COdFilterDomain(v As Variant) As OdFilterDomain
    Set COdFilterDomain = v
End Function

Public Function ToJson() As String
    Dim sValue As String
    If IsEmpty(mValue) Then
        Debug.Assert False
        sValue = "null"
    ElseIf IsObject(mValue) Then
        If mValue Is Nothing Then
            sValue = "null" ' null
        Else
            sValue = COdFilterDomain(mValue).ToJson()
        End If
    ElseIf VarType(mValue) = vbString Then
        sValue = "'" & mValue & "'"
    ElseIf VarType(mValue) = vbBoolean Then
        sValue = IIf(mValue, "true", "false")   ' true/false
    Else
        sValue = mValue
    End If

    ToJson = "['" & mFieldExpr & "', '" & mOperator & "', " & sValue & "]"
End Function

Public Sub SetValue(aValue As Variant)
    If IsObject(aValue) Then
        Set mValue = COdFilterDomain(aValue)
    Else
        mValue = aValue
    End If
End Sub

Public Function SetCriteria(aFieldExpr As String, aOperator As String, aValue As Variant) As OdFilterCriteria
    Set SetCriteria = Me
    mFieldExpr = aFieldExpr
    mOperator = aOperator
    SetValue aValue
End Function

Public Function SetFieldExpr(aFieldExpr As String) As OdFilterCriteria
    Set SetFieldExpr = Me
    mFieldExpr = aFieldExpr
    mOperator = ""
    mValue = Empty  ' Clear
End Function

Private Function SetOpeValue(aOperator As String, aValue As Variant) As OdFilterCriteria
    With New OdFilterCriteria
        Set SetOpeValue = .SetCriteria(mFieldExpr, aOperator, aValue)
    End With
End Function

Public Function Eq(aValue As Variant) As OdFilterCriteria
    Set Eq = SetOpeValue("=", aValue)
End Function

Public Function NotEq(aValue As Variant) As OdFilterCriteria
    Set NotEq = SetOpeValue("!=", aValue)
End Function

Public Function Gt(aValue As Variant) As OdFilterCriteria
    Set Gt = SetOpeValue(">", aValue)
End Function

Public Function Ge(aValue As Variant) As OdFilterCriteria
    Set Ge = SetOpeValue(">=", aValue)
End Function

Public Function Lt(aValue As Variant) As OdFilterCriteria
    Set Lt = SetOpeValue("<", aValue)
End Function

Public Function Le(aValue As Variant) As OdFilterCriteria
    Set Le = SetOpeValue("<=", aValue)
End Function

Public Function UnsetOrEq(aValue As Variant) As OdFilterCriteria
    Set UnsetOrEq = SetOpeValue("=?", aValue)
End Function

Public Function EqLike(aValue As Variant) As OdFilterCriteria
    Set EqLike = SetOpeValue("=like", aValue)
End Function

Public Function IsLike(aValue As Variant) As OdFilterCriteria
    Set IsLike = SetOpeValue("like", aValue)
End Function

Public Function NotLike(aValue As Variant) As OdFilterCriteria
    Set NotLike = SetOpeValue("not like", aValue)
End Function

Public Function IsILike(aValue As Variant) As OdFilterCriteria
    Set IsILike = SetOpeValue("ilike", aValue)
End Function

Public Function NotILike(aValue As Variant) As OdFilterCriteria
    Set NotILike = SetOpeValue("not ilike", aValue)
End Function

Public Function EqILike(aValue As Variant) As OdFilterCriteria
    Set EqILike = SetOpeValue("=ilike", aValue)
End Function

Public Function IsIn(aValue As Variant) As OdFilterCriteria
    Set IsIn = SetOpeValue("in", aValue)
End Function

Public Function NotIn(aValue As Variant) As OdFilterCriteria
    Set NotIn = SetOpeValue("not in", aValue)
End Function

Public Function ChildOf(aValue As Variant) As OdFilterCriteria
    Set ChildOf = SetOpeValue("child_of", aValue)
End Function

Public Function ParentOf(aValue As Variant) As OdFilterCriteria
    Set ParentOf = SetOpeValue("parent_of", aValue)
End Function

Public Function IsAny(aValue As Variant) As OdFilterCriteria
    Set IsAny = SetOpeValue("any", aValue)
End Function

Public Function NotAny(aValue As Variant) As OdFilterCriteria
    Set NotAny = SetOpeValue("not any", aValue)
End Function

