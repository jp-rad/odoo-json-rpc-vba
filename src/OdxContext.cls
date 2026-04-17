VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OdxContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' External API - odoo-JSON-RPC-VBA
'
' MIT License
'
' Copyright (c) 2022-2026 jp-rad
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

Private mClient As OdClient
Private mSchema As Dictionary

Public Sub InitContext(oClient As OdClient)
    Set mClient = oClient
    Set mSchema = New Dictionary
End Sub

Public Function AddRecordsetField(aRs As ADODB.Recordset, aModelName As String, aFieldName As String) As OdxModelView
    Dim dicModelFields As Dictionary
    Dim dicModelField As Dictionary
    
    Set dicModelFields = GetModelFields(aModelName)
    Debug.Assert Not dicModelFields Is Nothing
    Debug.Assert dicModelFields.Exists(aFieldName)
    Set dicModelField = dicModelFields(aFieldName)
    ' field
    OdxApi.AddRecordsetField aRs, dicModelField
    ' relations
    Set AddRecordsetField = CreateRelatedModelViewOrNothing(dicModelField)
End Function

Private Function GetModelFields(aModelName As String) As Dictionary
    Dim dicModelFields As Dictionary
    If mSchema.Exists(aModelName) Then
        Set dicModelFields = mSchema(aModelName)
        GoTo ExitProc
    End If
    Set dicModelFields = OdxApi.ExecuteModelFieldsGet(mClient, aModelName)
    Set mSchema(aModelName) = dicModelFields
ExitProc:
    Set GetModelFields = dicModelFields
    Exit Function
End Function

Private Function CreateRelatedModelViewOrNothing(dicModelField As Dictionary) As OdxModelView
    Dim oView As OdxModelView
    Dim sFieldType As String
    Dim sRelationModelName As String
    Dim sFKeyFieldName As String
    
    sRelationModelName = dicModelField(CODOO_ATTR_RELATION)
    
    sFieldType = dicModelField(CODOO_ATTR_TYPE)
    Select Case sFieldType
        Case "many2one"
            Set oView = New OdxModelView
            oView.InitMany2One Me, sRelationModelName
        Case "one2many"
            sFKeyFieldName = dicModelField(CODOO_ATTR_FKEY)
            Set oView = New OdxModelView
            oView.InitOne2Many Me, sRelationModelName, sFKeyFieldName
        Case "many2many"
            Set oView = New OdxModelView
            oView.InitMany2Many Me, sRelationModelName
        Case Else
            Set oView = Nothing
    End Select

    Set CreateRelatedModelViewOrNothing = oView
ExitProc:
    Exit Function
End Function

Public Function ExecuteSearchRead(aModelName As String, colFieldNames As Collection, aDomain As OdFilterDomain, limit As Long) As OdResult
    Set ExecuteSearchRead = OdxApi.ExecuteSearchReadModel(mClient, aModelName, colFieldNames, aDomain, limit)
End Function

Public Function ExecuteRead(aModelName As String, colFieldNames As Collection, colIds As Collection) As OdResult
    Set ExecuteRead = OdxApi.ExecuteReadModel(mClient, aModelName, colFieldNames, colIds)
End Function

Public Function NewModelView(aModelName As String) As OdxModelView
    Set NewModelView = New OdxModelView
    NewModelView.InitView Me, aModelName
End Function

