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
Private mValue As Variant   ' string, number, boolean, null or OdFilterDOmain

Private Const CERR_FIELD_EMPTY     As Long = vbObjectError + 1000
Private Const CERR_OPERATOR_EMPTY  As Long = vbObjectError + 1001
Private Const CERR_VALUE_MISSING   As Long = vbObjectError + 1002
Private Const CERR_INVALID_OBJECT  As Long = vbObjectError + 1003

Private Sub ValidateMembers()
    If Len(Trim(mFieldExpr)) = 0 Then
        Err.Raise CERR_FIELD_EMPTY, "OdFilterCriteria", "Field expression must not be empty."
    End If

    If Len(Trim(mOperator)) = 0 Then
        Err.Raise CERR_OPERATOR_EMPTY, "OdFilterCriteria", "Operator must not be empty." & vbCrLf & "[FieldExpr = '" & mFieldExpr & "']"
    End If

    If IsMissing(mValue) Then
        Err.Raise CERR_VALUE_MISSING, "OdFilterCriteria", "Value expression is Missing." & vbCrLf & "[FieldExpr = '" & mFieldExpr & "']"
    End If

    If IsObject(mValue) Then
        If Not TypeOf mValue Is OdFilterDomain Then
            Err.Raise CERR_INVALID_OBJECT, "OdFilterCriteria", "Object value must be of type OdFilterDomain." & vbCrLf & "[FieldExpr = '" & mFieldExpr & "']"
        End If
    End If
End Sub

Public Function ToJson() As String
    Dim sValue As String
    Dim filterDomain As OdFilterDomain
    
    ValidateMembers
    
    If IsNull(mValue) Then
        sValue = "null"
    ElseIf IsObject(mValue) Then
        If mValue Is Nothing Then
            sValue = "null" ' null
        Else
            Set filterDomain = mValue
            sValue = filterDomain.ToJson()
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

Public Sub SetValue(aValueExpr As Variant)
    If IsObject(aValueExpr) Then
        Set mValue = aValueExpr
    Else
        mValue = aValueExpr
    End If
End Sub

Public Function SetCriteria(aFieldExpr As String, aOperator As String, aValueExpr As Variant) As OdFilterCriteria
    Set SetCriteria = Me
    mFieldExpr = aFieldExpr
    mOperator = aOperator
    SetValue aValueExpr
End Function

Public Function SetFieldExpr(aFieldExpr As String) As OdFilterCriteria
    Set SetFieldExpr = Me
    mFieldExpr = aFieldExpr
    mOperator = ""
    mValue = Empty  ' Clear
End Function

Private Function SetOpeValue(aOperator As String, aValueExpr As Variant) As OdFilterCriteria
    With New OdFilterCriteria
        Set SetOpeValue = .SetCriteria(mFieldExpr, aOperator, aValueExpr)
    End With
End Function

Public Function Eq(aValueExpr As Variant) As OdFilterCriteria
    Set Eq = SetOpeValue("=", aValueExpr)
End Function

Public Function NotEq(aValueExpr As Variant) As OdFilterCriteria
    Set NotEq = SetOpeValue("!=", aValueExpr)
End Function

Public Function Gt(aValueExpr As Variant) As OdFilterCriteria
    Set Gt = SetOpeValue(">", aValueExpr)
End Function

Public Function Ge(aValueExpr As Variant) As OdFilterCriteria
    Set Ge = SetOpeValue(">=", aValueExpr)
End Function

Public Function Lt(aValueExpr As Variant) As OdFilterCriteria
    Set Lt = SetOpeValue("<", aValueExpr)
End Function

Public Function Le(aValueExpr As Variant) As OdFilterCriteria
    Set Le = SetOpeValue("<=", aValueExpr)
End Function

Public Function UnsetOrEq(aValueExpr As Variant) As OdFilterCriteria
    Set UnsetOrEq = SetOpeValue("=?", aValueExpr)
End Function

Public Function EqLike(aValueExpr As Variant) As OdFilterCriteria
    Set EqLike = SetOpeValue("=like", aValueExpr)
End Function

Public Function IsLike(aValueExpr As Variant) As OdFilterCriteria
    Set IsLike = SetOpeValue("like", aValueExpr)
End Function

Public Function NotLike(aValueExpr As Variant) As OdFilterCriteria
    Set NotLike = SetOpeValue("not like", aValueExpr)
End Function

Public Function IsILike(aValueExpr As Variant) As OdFilterCriteria
    Set IsILike = SetOpeValue("ilike", aValueExpr)
End Function

Public Function NotILike(aValueExpr As Variant) As OdFilterCriteria
    Set NotILike = SetOpeValue("not ilike", aValueExpr)
End Function

Public Function EqILike(aValueExpr As Variant) As OdFilterCriteria
    Set EqILike = SetOpeValue("=ilike", aValueExpr)
End Function

Public Function IsIn(aValueExpr As Variant) As OdFilterCriteria
    Set IsIn = SetOpeValue("in", aValueExpr)
End Function

Public Function NotIn(aValueExpr As Variant) As OdFilterCriteria
    Set NotIn = SetOpeValue("not in", aValueExpr)
End Function

Public Function ChildOf(aValueExpr As Variant) As OdFilterCriteria
    Set ChildOf = SetOpeValue("child_of", aValueExpr)
End Function

Public Function ParentOf(aValueExpr As Variant) As OdFilterCriteria
    Set ParentOf = SetOpeValue("parent_of", aValueExpr)
End Function

Public Function IsAny(aValueExpr As Variant) As OdFilterCriteria
    Set IsAny = SetOpeValue("any", aValueExpr)
End Function

Public Function NotAny(aValueExpr As Variant) As OdFilterCriteria
    Set NotAny = SetOpeValue("not any", aValueExpr)
End Function

Public Function EqNull() As OdFilterCriteria
    Set EqNull = SetOpeValue("=", Null)
End Function

Public Function NotNull() As OdFilterCriteria
    Set NotNull = SetOpeValue("!=", Null)
End Function

