Attribute VB_Name = "Od"
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
' External API - odoo docs
'
' Odoo is usually extended internally via modules, but many of its features and all of its data
' are also available from the outside for external analysis or integration with various tools.
' Part of the Models API is easily available over XML-RPC and accessible from a variety of languages.
'
' https://www.odoo.com/documentation/master/developer/reference/external_api.html
'

Public Const CERR_STATUSCODE   As Long = 2001 + vbObjectError   ' web response error
Public Const CERR_RESPONSE     As Long = 2002 + vbObjectError   ' JSON-RPC error
Public Const CERR_JSONRPC_ID   As Long = 2003 + vbObjectError   ' JSON-RPC id error
Public Const CERR_AUTHENTICATE As Long = 2004 + vbObjectError   ' authentication failed

Public Function FormatDate(aDate As Date) As String
    FormatDate = Format(aDate, "yyyy-mm-dd")
End Function

Public Function ParseDate(aDateString As String) As Date
    ParseDate = CDate(aDateString)
End Function

Public Function ConvertToIsoDatetime(aDatetime As Date) As String
    ConvertToIsoDatetime = JsonConverter.ConvertToIso(aDatetime)
End Function

Public Function ParseIsoDatetime(aIsoString As String) As Date
    ParseIsoDatetime = JsonConverter.ParseIso(aIsoString)
End Function

Public Function NewOdClient() As OdClient
    Set NewOdClient = New OdClient
End Function

Public Function NewDomain() As OdFilterDomain
    Set NewDomain = New OdFilterDomain
End Function

Public Function NewField(aFieldExpr As String) As OdFilterCriteria
    With New OdFilterCriteria
        Set NewField = .SetFieldExpr(aFieldExpr)
    End With
End Function

Public Function NewCriteria(aFieldExpr As String, aOperator As String, aValue As Variant) As OdFilterCriteria
    With New OdFilterCriteria
        Set NewCriteria = .SetCriteria(aFieldExpr, aOperator, aValue)
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

Public Function NewList() As Collection
    Set NewList = New Collection
End Function

Public Function AddList(aList As Collection) As Collection
    Set AddList = New Collection
    aList.Add AddList
End Function

Public Function AddDict(aList As Collection) As Dictionary
    Set AddDict = New Dictionary
    aList.Add AddDict
End Function

Public Function NewDict() As Dictionary
    Set NewDict = New Dictionary
End Function

Public Function SetList(aDict As Dictionary, aKey As String) As Collection
    Set SetList = New Collection
    Set aDict(aKey) = SetList
End Function

Public Function SetDict(aDict As Dictionary, aKey As String) As Dictionary
    Set SetDict = New Dictionary
    Set aDict(aKey) = SetDict
End Function
