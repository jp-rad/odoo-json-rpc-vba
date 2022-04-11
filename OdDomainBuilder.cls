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

Private mDomain As New Collection

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

Public Property Get RefMe() As OdDomainBuilder
    Set RefMe = Me
End Property

Public Function GetDomain() As Collection
    Set GetDomain = mDomain
End Function

Public Function NewCollection(Optional args As Variant) As OdDomainBuilder
    With New OdDomainBuilder
        mDomain.Add AddItemsTo(args, .GetDomain())
        Set NewCollection = .RefMe
    End With
End Function

Public Function AddCriteria(aFieldName As String) As OdDomainCriteria
    With New OdDomainCriteria
        Set AddCriteria = .AssignFieldName(aFieldName, Me)
    End With
End Function

Private Function AddCollection(Optional args As Variant) As OdDomainBuilder
    mDomain.Add AddItemsTo(args, New Collection)
    Set AddCollection = Me
End Function

Public Function AddCriteriaTriple(aFieldName As String, aOperator As String, aValue As Variant) As OdDomainBuilder
    AddCollection Array(aFieldName, aOperator, aValue)
    Set AddCriteriaTriple = Me
End Function

Private Function AddLogical(v As String) As OdDomainBuilder
    mDomain.Add v
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
