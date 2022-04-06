VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OdDomainBuilder"
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

Private mCollection As New Collection

Private Function AddItemsTo(args As Variant, target As Collection) As Collection
    Dim v As Variant
    If VarType(args) = vbError Then     ' not assign
    
    ElseIf VarType(args) = 8204 Then    ' Array() = 8204
        For Each v In args
            target.Add v
        Next v
    Else
        target.Add args
    End If
    Set AddItemsTo = target
End Function

Public Property Get RefCollection() As Collection
    Set RefCollection = mCollection
End Property

Public Property Get RefMe() As OdDomainBuilder
    Set RefMe = Me
End Property

Public Function GetDomain() As Collection
    Set GetDomain = mCollection
End Function

Public Function NewCollection(Optional args As Variant) As OdDomainBuilder
    With New OdDomainBuilder
        mCollection.Add AddItemsTo(args, .RefCollection)
        Set NewCollection = .RefMe
    End With
End Function

Private Function AddCollection(Optional args As Variant) As OdDomainBuilder
    mCollection.Add AddItemsTo(args, New Collection)
    Set AddCollection = Me
End Function

Public Function AddCriteria(aFieldName As String, aOperator As String, aValue As Variant) As OdDomainBuilder
    AddCollection Array(aFieldName, aOperator, aValue)
    Set AddCriteria = Me
End Function

Public Function AddCriteriaEq(aFieldName As String, aValue As Variant) As OdDomainBuilder
    Set AddCriteriaEq = Me.AddCriteria(aFieldName, "=", aValue)
End Function

Public Function AddCriteriaNotEq(aFieldName As String, aValue As Variant) As OdDomainBuilder
    Set AddCriteriaNotEq = Me.AddCriteria(aFieldName, "!=", aValue)
End Function

Public Function AddCriteriaGt(aFieldName As String, aValue As Variant) As OdDomainBuilder
    Set AddCriteriaGt = Me.AddCriteria(aFieldName, ">", aValue)
End Function

Public Function AddCriteriaGe(aFieldName As String, aValue As Variant) As OdDomainBuilder
    Set AddCriteriaGe = Me.AddCriteria(aFieldName, ">=", aValue)
End Function

Public Function AddCriteriaLt(aFieldName As String, aValue As Variant) As OdDomainBuilder
    Set AddCriteriaLt = Me.AddCriteria(aFieldName, "<", aValue)
End Function

Public Function AddCriteriaLe(aFieldName As String, aValue As Variant) As OdDomainBuilder
    Set AddCriteriaLe = Me.AddCriteria(aFieldName, "<=", aValue)
End Function

Public Function AddCriteriaUnsetOrEq(aFieldName As String, aValue As Variant) As OdDomainBuilder
    Set AddCriteriaUnsetOrEq = Me.AddCriteria(aFieldName, "=?", aValue)
End Function

Public Function AddCriteriaEqLike(aFieldName As String, aValue As Variant) As OdDomainBuilder
    Set AddCriteriaEqLike = Me.AddCriteria(aFieldName, "=like", aValue)
End Function

Public Function AddCriteriaLike(aFieldName As String, aValue As Variant) As OdDomainBuilder
    Set AddCriteriaLike = Me.AddCriteria(aFieldName, "like", aValue)
End Function

Public Function AddCriteriaNotLike(aFieldName As String, aValue As Variant) As OdDomainBuilder
    Set AddCriteriaNotLike = Me.AddCriteria(aFieldName, "not like", aValue)
End Function

Public Function AddCriteriaCILike(aFieldName As String, aValue As Variant) As OdDomainBuilder
    Set AddCriteriaCILike = Me.AddCriteria(aFieldName, "ilike", aValue)
End Function

Public Function AddCriteriaNotCILike(aFieldName As String, aValue As Variant) As OdDomainBuilder
    Set AddCriteriaNotCILike = Me.AddCriteria(aFieldName, "not ilike", aValue)
End Function

Public Function AddCriteriaEqCILike(aFieldName As String, aValue As Variant) As OdDomainBuilder
    Set AddCriteriaEqCILike = Me.AddCriteria(aFieldName, "=ilike", aValue)
End Function

Public Function AddCriteriaIn(aFieldName As String, aValue As Variant) As OdDomainBuilder
    Set AddCriteriaIn = Me.AddCriteria(aFieldName, "in", aValue)
End Function

Public Function AddCriteriaNotIn(aFieldName As String, aValue As Variant) As OdDomainBuilder
    Set AddCriteriaNotIn = Me.AddCriteria(aFieldName, "not in", aValue)
End Function

Public Function AddCriteriaChildOf(aFieldName As String, aValue As Variant) As OdDomainBuilder
    Set AddCriteriaChildOf = Me.AddCriteria(aFieldName, "child_of", aValue)
End Function

Public Function AddCriteriaParentOf(aFieldName As String, aValue As Variant) As OdDomainBuilder
    Set AddCriteriaParentOf = Me.AddCriteria(aFieldName, "parent_of", aValue)
End Function

Private Function AddLogical(v As String) As OdDomainBuilder
    mCollection.Add v
    Set AddLogical = Me
End Function

Public Function AddLogicalAnd() As OdDomainBuilder
    Set AddLogicalAnd = AddLogical("&")
End Function

Public Function AddLogicalOr() As OdDomainBuilder
    Set AddLogicalOr = AddLogical("|")
End Function

Public Function AddLogicalNot() As OdDomainBuilder
    Set AddLogicalNot = AddLogical("!")
End Function
