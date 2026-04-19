Attribute VB_Name = "OdxApi"
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

Private Const CBIT_DSP      As Long = &H100000  ' bit20 - display name, not model's field
Private Const CBIT_M2O      As Long = &H200000  ' bit21 - many2one
Private Const CBIT_O2M      As Long = &H400000  ' bit22 - one2many
Private Const CBIT_M2M      As Long = &H800000  ' bit23 - many2many
Private Const CBIT_NULLABLE As Long = FieldAttributeEnum.adFldMayBeNull ' Enable NULLs via adFldMayBeNull even for Odoo's required fields.

Private Const CATTR_PK_ID        As Long = FieldAttributeEnum.adFldKeyColumn
Private Const CATTR_FIELD        As Long = FieldAttributeEnum.adFldUpdatable _
                                            Or FieldAttributeEnum.adFldIsNullable   ' Force adFldIsNullable to allow NULLs in the Recordset regardless of DB constraints.
Private Const CATTR_DISPLAY_NAME As Long = CATTR_FIELD Or CBIT_DSP
Private Const CATTR_M2O          As Long = CATTR_FIELD Or CBIT_M2O
Private Const CATTR_O2M          As Long = CATTR_FIELD Or CBIT_O2M
Private Const CATTR_M2M          As Long = CATTR_FIELD Or CBIT_M2M

Public Const CODOO_ATTR_NAME     As String = "name"
Public Const CODOO_ATTR_STRING   As String = "string"
Public Const CODOO_ATTR_TYPE     As String = "type"
Public Const CODOO_ATTR_RELATION As String = "relation"
Public Const CODOO_ATTR_FKEY     As String = "relation_field"  ' foreign key
Public Const CODOO_ATTR_REQUIRED As String = "required"

Private Const CFIELD_SEARCHED As String = "__searched__"
Private Const CDISPLAY_NAME As String = "_display_name_"

Public Function IsOdooField(fld As ADODB.Field) As Boolean
    IsOdooField = 0 = (fld.attributes And CBIT_DSP)
End Function

Public Function IsOdooMany2OneField(fld As ADODB.Field) As Boolean
    IsOdooMany2OneField = 0 <> (fld.attributes And CBIT_M2O)
End Function

Public Function IsOdooOne2ManyField(fld As ADODB.Field) As Boolean
    IsOdooOne2ManyField = 0 <> (fld.attributes And CBIT_O2M)
End Function

Public Function IsOdooMany2ManyField(fld As ADODB.Field) As Boolean
    IsOdooMany2ManyField = 0 <> (fld.attributes And CBIT_M2M)
End Function

Public Function IsOdooNullableField(fld As ADODB.Field) As Boolean
    IsOdooNullableField = 0 <> (fld.attributes And CBIT_NULLABLE)
End Function

Public Function IsOdooRequiredField(fld As ADODB.Field) As Boolean
    IsOdooRequiredField = Not IsOdooNullableField(fld)
End Function

Public Function IsOdooDateField(fld As ADODB.Field) As Boolean
    IsOdooDateField = adDBDate = fld.Type
End Function

Public Function IsOdooDateTimeField(fld As ADODB.Field) As Boolean
    IsOdooDateTimeField = adDate = fld.Type
End Function

Public Function ExecuteModelFieldsGet(oClient As OdClient, aModelName As String) As Dictionary
    Dim colAttributes As Collection
    Set colAttributes = NewList
    With colAttributes
        .Add CODOO_ATTR_NAME
        .Add CODOO_ATTR_STRING
        .Add CODOO_ATTR_TYPE
        .Add CODOO_ATTR_RELATION
        .Add CODOO_ATTR_FKEY
        .Add CODOO_ATTR_REQUIRED
    End With
    Set ExecuteModelFieldsGet = oClient.Model(aModelName).MethodFieldsGet(, colAttributes).Result
End Function

Public Function ExecuteSearchReadModel(oClient As OdClient, aModelName As String, colFieldNames As Collection, aDomain As OdFilterDomain, limit As Long) As OdResult
    Dim params As Collection
    Dim named As Dictionary
    
    Set params = NewList
    aDomain.BuildAndAppend params
    Set named = NewDict
    named.Add "fields", colFieldNames
    If limit > 0 Then
        named.Add "limit", limit
    End If
    Set ExecuteSearchReadModel = oClient.Model(aModelName).Method("search_read").ExecuteKw(params, named)
End Function

Public Function ExecuteReadModel(oClient As OdClient, aModelName As String, colFieldNames As Collection, colIds As Collection) As OdResult
    Dim params As Collection
    Dim named As Dictionary
    
    Set params = NewList
    params.Add colIds
    Set named = NewDict
    named.Add "fields", colFieldNames
    
    Set ExecuteReadModel = oClient.Model(aModelName).Method("read").ExecuteKw(params, named)
End Function

Public Function CreateNewRecordSet() As ADODB.Recordset
    Set CreateNewRecordSet = New ADODB.Recordset
    CreateNewRecordSet.CursorLocation = adUseClient
    With CreateNewRecordSet.Fields
        ' <<PK>> id: Long
        .Append Name:="id", Type:=adInteger, Attrib:=CATTR_PK_ID
    End With
End Function

Private Function FormatDisplayName(aFieldName As String) As String
    FormatDisplayName = aFieldName & CDISPLAY_NAME
End Function

Public Function AddRecordsetField(rs As ADODB.Recordset, dicModelField As Dictionary) As ADODB.Field
    Dim sFieldName As String
    Dim sFieldType As String
    Dim attrOdooNullable As Long
    
    Debug.Assert dicModelField.Exists(CODOO_ATTR_NAME)
    sFieldName = dicModelField(CODOO_ATTR_NAME)
    If dicModelField.Exists(CODOO_ATTR_REQUIRED) Then
        attrOdooNullable = IIf(dicModelField(CODOO_ATTR_REQUIRED), 0, CBIT_NULLABLE)
    Else
        attrOdooNullable = CBIT_NULLABLE
    End If
    Debug.Assert dicModelField.Exists(CODOO_ATTR_TYPE)
    sFieldType = dicModelField(CODOO_ATTR_TYPE)
    Select Case sFieldType

        '----------------------------------------
        ' many2one: (id, display_name)
        '----------------------------------------
        Case "many2one"
            rs.Fields.Append Name:=sFieldName, Type:=adInteger, Attrib:=CATTR_M2O Or attrOdooNullable
            rs.Fields.Append Name:=FormatDisplayName(sFieldName), Type:=adVarWChar, DefinedSize:=-1, Attrib:=CATTR_DISPLAY_NAME

        '----------------------------------------
        ' one2many: list of id -> CSV
        '----------------------------------------
        Case "one2many"
            rs.Fields.Append Name:=sFieldName, Type:=adLongVarWChar, DefinedSize:=-1, Attrib:=CATTR_O2M Or attrOdooNullable

        '----------------------------------------
        ' many2many: list of id ¨ CSV
        '----------------------------------------
        Case "many2many"
            rs.Fields.Append Name:=sFieldName, Type:=adLongVarWChar, DefinedSize:=-1, Attrib:=CATTR_M2M Or attrOdooNullable

        '----------------------------------------
        ' char
        '----------------------------------------
        Case "char"
            rs.Fields.Append Name:=sFieldName, Type:=adVarWChar, DefinedSize:=-1, Attrib:=CATTR_FIELD Or attrOdooNullable

        '----------------------------------------
        ' text
        '----------------------------------------
        Case "text"
            rs.Fields.Append Name:=sFieldName, Type:=adLongVarWChar, DefinedSize:=-1, Attrib:=CATTR_FIELD Or attrOdooNullable

        '----------------------------------------
        ' html
        '----------------------------------------
        Case "html"
            rs.Fields.Append Name:=sFieldName, Type:=adLongVarWChar, DefinedSize:=-1, Attrib:=CATTR_FIELD Or attrOdooNullable

        '----------------------------------------
        ' json
        '----------------------------------------
        Case "json"
            rs.Fields.Append Name:=sFieldName, Type:=adLongVarWChar, DefinedSize:=-1, Attrib:=CATTR_FIELD Or attrOdooNullable

        '----------------------------------------
        ' boolean
        '----------------------------------------
        Case "boolean"
            rs.Fields.Append Name:=sFieldName, Type:=adBoolean, Attrib:=CATTR_FIELD Or attrOdooNullable

        '----------------------------------------
        ' integer
        '----------------------------------------
        Case "integer"
            rs.Fields.Append Name:=sFieldName, Type:=adInteger, Attrib:=CATTR_FIELD Or attrOdooNullable

        '----------------------------------------
        ' float
        '----------------------------------------
        Case "float"
            rs.Fields.Append Name:=sFieldName, Type:=adDouble, Attrib:=CATTR_FIELD Or attrOdooNullable

        '----------------------------------------
        ' monetary
        '----------------------------------------
        Case "monetary"
            rs.Fields.Append Name:=sFieldName, Type:=adCurrency, Attrib:=CATTR_FIELD Or attrOdooNullable

        '----------------------------------------
        ' date
        '----------------------------------------
        Case "date"
            rs.Fields.Append Name:=sFieldName, Type:=adDBDate, Attrib:=CATTR_FIELD Or attrOdooNullable
        
        '----------------------------------------
        ' datetime (UTC -> local datetime)
        '----------------------------------------
        Case "datetime"
            rs.Fields.Append Name:=sFieldName, Type:=adDate, Attrib:=CATTR_FIELD Or attrOdooNullable

        '----------------------------------------
        ' binary (base64 char)
        '----------------------------------------
        Case "binary"
            rs.Fields.Append Name:=sFieldName, Type:=adLongVarWChar, DefinedSize:=-1, Attrib:=CATTR_FIELD Or attrOdooNullable

        '----------------------------------------
        ' selection (char)
        '----------------------------------------
        Case "selection"
            rs.Fields.Append Name:=sFieldName, Type:=adVarWChar, DefinedSize:=-1, Attrib:=CATTR_FIELD Or attrOdooNullable

        '----------------------------------------
        ' properties (json char)
        '----------------------------------------
        Case "properties"
            rs.Fields.Append Name:=sFieldName, Type:=adLongVarWChar, DefinedSize:=-1, Attrib:=CATTR_FIELD Or attrOdooNullable

        '----------------------------------------
        ' Unknown (char)
        '----------------------------------------
        Case Else
            Debug.Print "Unknown type:", sFieldType, sFieldName
            Debug.Assert False
            rs.Fields.Append Name:=sFieldName, Type:=adVarWChar, DefinedSize:=-1, Attrib:=CATTR_FIELD Or attrOdooNullable

    End Select

ExitProc:
    Set AddRecordsetField = rs.Fields(sFieldName)
    Exit Function
End Function

Public Function GetDisplayNameField(rs As ADODB.Recordset, aFieldName As String) As ADODB.Field
    Set GetDisplayNameField = rs.Fields(FormatDisplayName(aFieldName))
End Function

Public Function NewContext(oClient As OdClient) As OdxContext
    Dim oCtx As OdxContext
    Set oCtx = New OdxContext
    oCtx.InitContext oClient
    Set NewContext = oCtx
End Function

