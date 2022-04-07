VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OdDomainCriteria"
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
' ORM API - odoo docs
'
' Search domains:
' https://www.odoo.com/documentation/15.0/developer/reference/backend/orm.html#search-domains
'

Option Explicit

Private mFieldName As String
Private mDomainBuilder As OdDomainBuilder

Public Function AssignFieldName(aFieldName As String, aDomainBuilder As OdDomainBuilder) As OdDomainCriteria
    Set AssignFieldName = Me
    mFieldName = aFieldName
    Set mDomainBuilder = aDomainBuilder
End Function

Private Function SetOpeValue(aOperator As String, aValue As Variant) As OdDomainBuilder
    Set SetOpeValue = mDomainBuilder.AddCriteriaTriple(mFieldName, aOperator, aValue)
End Function

Public Function Eq(aValue As Variant) As OdDomainBuilder
    Set Eq = SetOpeValue("=", aValue)
End Function

Public Function NotEq(aValue As Variant) As OdDomainBuilder
    Set NotEq = SetOpeValue("!=", aValue)
End Function

Public Function Gt(aValue As Variant) As OdDomainBuilder
    Set Gt = SetOpeValue(">", aValue)
End Function

Public Function Ge(aValue As Variant) As OdDomainBuilder
    Set Ge = SetOpeValue(">=", aValue)
End Function

Public Function Lt(aValue As Variant) As OdDomainBuilder
    Set Lt = SetOpeValue("<", aValue)
End Function

Public Function Le(aValue As Variant) As OdDomainBuilder
    Set Le = SetOpeValue("<=", aValue)
End Function

Public Function UnsetOrEq(aValue As Variant) As OdDomainBuilder
    Set UnsetOrEq = SetOpeValue("=?", aValue)
End Function

Public Function EqLike(aValue As Variant) As OdDomainBuilder
    Set EqLike = SetOpeValue("=like", aValue)
End Function

Public Function IsLike(aValue As Variant) As OdDomainBuilder
    Set IsLike = SetOpeValue("like", aValue)
End Function

Public Function NotLike(aValue As Variant) As OdDomainBuilder
    Set NotLike = SetOpeValue("not like", aValue)
End Function

Public Function CILike(aValue As Variant) As OdDomainBuilder
    Set CILike = SetOpeValue("ilike", aValue)
End Function

Public Function NotCILike(aValue As Variant) As OdDomainBuilder
    Set NotCILike = SetOpeValue("not ilike", aValue)
End Function

Public Function EqCILike(aValue As Variant) As OdDomainBuilder
    Set EqCILike = SetOpeValue("=ilike", aValue)
End Function

Public Function IsIn(aValue As Variant) As OdDomainBuilder
    Set IsIn = SetOpeValue("in", aValue)
End Function

Public Function NotIn(aValue As Variant) As OdDomainBuilder
    Set NotIn = SetOpeValue("not in", aValue)
End Function

Public Function ChildOf(aValue As Variant) As OdDomainBuilder
    Set ChildOf = SetOpeValue("child_of", aValue)
End Function

Public Function ParentOf(aValue As Variant) As OdDomainBuilder
    Set ParentOf = SetOpeValue("parent_of", aValue)
End Function


