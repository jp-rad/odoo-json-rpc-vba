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

'
' Search domains, ORM API - odoo docs
'
' A domain is a list of criteria, each criterion being a triple (either a list or a tuple)
' of (field_name, operator, value)
'
' see also: https://www.odoo.com/documentation/18.0/developer/reference/backend/orm.html#search-domains
'

Private mLogic As String
Private mArity1 As Object
Private mArity2 As Object

Public Function ToFilterString() As String
    ToFilterString = "'" & mLogic & "', " & mArity1.ToFilterString()
    If Not mArity2 Is Nothing Then
        ToFilterString = ToFilterString & ", " & mArity2.ToFilterString()
    End If
End Function

Private Function SetLogic(aLogic As String, aArity1 As Object, aArity2 As Object) As OdFilterCombi
    Set SetLogic = Me
    mLogic = aLogic
    Set mArity1 = aArity1
    Set mArity2 = aArity2
End Function

Public Function SetAndLogic(aArity1 As Object, aArity2 As Object) As OdFilterCombi
    Set SetAndLogic = SetLogic("&", aArity1, aArity2)
End Function

Public Function SetOrLogic(aArity1 As Object, aArity2 As Object) As OdFilterCombi
    Set SetOrLogic = SetLogic("|", aArity1, aArity2)
End Function

Public Function SetNotLogic(aArity As Object) As OdFilterCombi
    Set SetNotLogic = SetLogic("!", aArity, Nothing)
End Function

