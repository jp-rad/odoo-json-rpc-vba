VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OdFIlterCriteria"
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

Private mFieldName As String
Private mOperator As String
Private mValue As Variant   ' string, number, domain

Public Function ToFilterString() As String
    Dim sValue As String
    If IsObject(mValue) Then
        sValue = COdFilterDomain(mValue).ToFilterString()
    ElseIf VarType(mValue) = vbString Then
        sValue = "'" & mValue & "'"
    Else
        sValue = mValue
    End If

    ToFilterString = "('" & mFieldName & "', '" & mOperator & "', " & sValue & ")"
End Function

Private Function COdFilterDomain(v As Variant) As OdFilterDomain
    Set COdFilterDomain = v
End Function

Public Function SetCriteria(aFieldName As String, aOperator As String, aValue As Variant) As OdFilterCriteria
    Set SetCriteria = Me
    mFieldName = aFieldName
    mOperator = aOperator
    If IsObject(aValue) Then
        Set mValue = COdFilterDomain(aValue)
    Else
        mValue = aValue
    End If
End Function

