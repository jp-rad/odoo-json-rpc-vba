Attribute VB_Name = "OdFilter"
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

Public Function NewDomain() As OdFilterDomain
    Set NewDomain = New OdFilterDomain
End Function

Public Function NewField(field_name As String) As OdFilterField
    With New OdFilterField
        Set NewField = .SetFieldName(field_name)
    End With
End Function

Public Function NewCriteria(field_name As String, operator As String, value As Variant) As OdFilterCriteria
    With New OdFilterCriteria
        Set NewCriteria = .SetCriteria(field_name, operator, value)
    End With
End Function

Public Function NewAnd(aArity1 As Object, aArity2 As Object) As OdFilterCombi
    With New OdFilterCombi
        Set NewAnd = .SetAndLogic(aArity1, aArity2)
    End With
End Function

Public Function NewOr(aArity1 As Object, aArity2 As Object) As OdFilterCombi
    With New OdFilterCombi
        Set NewOr = .SetOrLogic(aArity1, aArity2)
    End With
End Function

Public Function NewNot(aArity As Object) As OdFilterCombi
    With New OdFilterCombi
        Set NewNot = .SetNotLogic(aArity)
    End With
End Function

