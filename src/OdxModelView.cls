VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OdxModelView"
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

Private mCtx As OdxContext
Private mRs As ADODB.Recordset
Private mModelName As String

Private mRelations As Dictionary    ' key: relation field name, value: OdxModelView

Private mFKeyFieldName As String    ' for one2many) fk field name

Private mRelationMap As Collection  ' for many2many) key: CStr(parent id), value: collection of id

Private Enum odxViewType
    vtDefault
    vtMany2One
    vtOne2Many
    vtMany2Many
End Enum

Private mViewType As odxViewType

Public Property Get IsMany2One() As Boolean
    IsMany2One = mViewType = vtMany2One
End Property

Public Property Get IsOne2Many() As Boolean
    IsOne2Many = mViewType = vtOne2Many
End Property

Public Property Get IsMany2Many() As Boolean
    IsMany2Many = mViewType = vtMany2Many
End Property

Public Sub InitView(aCtx As OdxContext, aModelName As String)
    Set mCtx = aCtx
    Set mRs = OdxApi.CreateNewRecordSet()
    mModelName = aModelName
    Set mRelations = New Dictionary
    mFKeyFieldName = ""
    Set mRelationMap = Nothing
    mViewType = vtDefault
End Sub

Public Sub InitMany2One(aCtx As OdxContext, aModelName As String)
    InitView aCtx, aModelName
    mViewType = vtMany2One
End Sub

Public Sub InitOne2Many(aCtx As OdxContext, aModelName As String, aFKeyFieldName As String)
    InitView aCtx, aModelName
    mFKeyFieldName = aFKeyFieldName
    AddField mFKeyFieldName
    mViewType = vtOne2Many
End Sub

Public Sub InitMany2Many(aCtx As OdxContext, aModelName As String)
    InitView aCtx, aModelName
    Set mRelationMap = New Collection
    mViewType = vtMany2Many
End Sub

Public Property Get ModelName() As String
    ModelName = mModelName
End Property

Private Sub SaftyOpenRecordset()
On Error Resume Next
    mRs.Open
End Sub

Private Sub SaftyMoveFirst()
On Error Resume Next
    mRs.MoveFirst
End Sub

Public Property Get Recordset() As ADODB.Recordset
    Set Recordset = mRs
End Property

Public Function AddField(aFieldName As String) As OdxModelView
    Dim oRelation As OdxModelView
    Set oRelation = mCtx.AddRecordsetField(mRs, mModelName, aFieldName)
    If Not oRelation Is Nothing Then
        Set mRelations(aFieldName) = oRelation
        Set AddField = oRelation
    End If
End Function

Public Sub ClearRecords()
    Dim ds As OdxModelView
    Dim v As Variant
    
    SaftyOpenRecordset  ' open recordset, no error
    SaftyMoveFirst
    
    Do Until mRs.EOF
        mRs.Delete
        mRs.MoveNext
    Loop
    
    For Each v In mRelations.Keys
        Set ds = mRelations(v)
        ds.ClearRecords
    Next

End Sub

Private Sub MoveFirstOrAddNewRecord(aId As Long)
    SaftyMoveFirst
    mRs.Find "id = " & aId

    If mRs.EOF Then
        mRs.AddNew
        mRs("id").Value = aId
    End If
End Sub

Public Sub UpsertMany2One(aId As Long)
    MoveFirstOrAddNewRecord aId
    mRs.Update
End Sub

Public Sub UpsertOne2Many(aId As Long, aFkId As Long)
    MoveFirstOrAddNewRecord aFkId
    mRs(mFKeyFieldName) = aId
    mRs.Update
End Sub

Public Sub UpsertMany2Many(aId As Long, aFkId As Long)
    Dim sId As String
    Dim colIds As Collection
    MoveFirstOrAddNewRecord aFkId
    mRs.Update
    
    ' map
    sId = CStr(aId)
    On Error Resume Next
    Set colIds = mRelationMap(sId)
    On Error GoTo 0
    If colIds Is Nothing Then
        Set colIds = New Collection
        mRelationMap.Add colIds, sId
    End If
    colIds.Add aFkId
    
End Sub

Private Function CollectionToCsv(col As Collection) As String
    Dim s As String
    Dim v As Variant
    
    s = ""
    For Each v In col
        If s <> "" Then s = s & ","
        s = s & CStr(v)
    Next v
    
    CollectionToCsv = s
End Function

Private Function UpsertRecordSet(colResult As Collection)
    Dim dicRecord As Dictionary
    Dim dsRelation As OdxModelView
    
    Dim nId As Long
    Dim fld As ADODB.Field
    Dim colValue As Collection  ' value of many2one/one2many/many2many
    Dim v As Variant
    
    Dim nRelatedId As Long
    Dim sRelatedDisplayName As String
    Dim sRelatedFkFieldName As String
        
    For Each dicRecord In colResult
        nId = dicRecord("id")
        MoveFirstOrAddNewRecord nId     ' Edit or AddNew
        
        For Each fld In mRs.Fields
            If Not IsOdooField(fld) Then
                ' SKIP
                GoTo CONTINUE_FOR_EACH
            End If
            If fld.Type = adBoolean Then
                ' boolean
                fld.Value = dicRecord(fld.Name)
                GoTo CONTINUE_FOR_EACH
            End If
            If VarType(dicRecord(fld.Name)) = vbBoolean Then
                ' Null
                GoTo CONTINUE_FOR_EACH
            End If
            If IsOdooMany2OneField(fld) Then
                ' many2one
                ' id (fk)
                Set colValue = dicRecord(fld.Name)
                nRelatedId = colValue(1)
                fld.Value = nRelatedId
                ' display_name
                sRelatedDisplayName = colValue(2)
                OdxApi.GetDisplayNameField(mRs, fld.Name).Value = sRelatedDisplayName
                ' relation
                Set dsRelation = mRelations(fld.Name)
                dsRelation.UpsertMany2One nRelatedId
            ElseIf IsOdooOne2ManyField(fld) Then
                ' one2many
                ' ids (list)
                Set colValue = dicRecord(fld.Name)
                fld.Value = "[" & CollectionToCsv(colValue) & "]"
                ' relation
                Set dsRelation = mRelations(fld.Name)
                For Each v In colValue
                    nRelatedId = v
                    dsRelation.UpsertOne2Many nId, nRelatedId
                Next v
            ElseIf IsOdooMany2ManyField(fld) Then
                ' many2many
                ' ids (list)
                Set colValue = dicRecord(fld.Name)
                fld.Value = "[" & CollectionToCsv(colValue) & "]"
                ' relation
                Set dsRelation = mRelations(fld.Name)
                For Each v In colValue
                    nRelatedId = v
                    dsRelation.UpsertMany2Many nId, nRelatedId
                Next v
            ElseIf IsOdooDateField(fld) Then
                ' date
                fld.Value = ParseDate(dicRecord(fld.Name))
            ElseIf IsOdooDateTimeField(fld) Then
                ' datetime
                fld.Value = ParseIsoDatetime(dicRecord(fld.Name))
            Else
                fld.Value = dicRecord(fld.Name)
            End If
CONTINUE_FOR_EACH:
        Next fld
        
        mRs.Update  ' Update, Edit or AddNew
    Next dicRecord
    
    SaftyMoveFirst
    
    ' fetch relations
    For Each v In mRelations.Keys
        Set dsRelation = mRelations(v)
        dsRelation.FetchRecordset
    Next v
    
End Function

Private Function GetFieldNames() As Collection
    Dim fld As ADODB.Field
    
    Set GetFieldNames = NewList
    For Each fld In mRs.Fields
        If IsOdooField(fld) Then
            GetFieldNames.Add fld.Name
        End If
    Next fld
End Function

Public Function ExecuteSearchRead(aDomain As OdFilterDomain, Optional limit As Long = -1) As Long
    Dim colFieldNames As Collection
    Dim colResult As Collection
    
    ClearRecords    ' open recordset(s), no error
    
    Set colFieldNames = GetFieldNames()
    
    Set colResult = mCtx.ExecuteSearchRead(mModelName, colFieldNames, aDomain, limit).Result
    
    UpsertRecordSet colResult
    
    ExecuteSearchRead = colResult.Count
End Function

Public Function FetchRecordset() As Long
    Dim colIds As Collection
    Dim nId As Long
    Dim colFieldNames As Collection
    Dim colResult As Collection
        
    Set colFieldNames = GetFieldNames()
    
    Set colIds = NewList
    If Not mRs.BOF Then
        mRs.MoveFirst
    End If
    Do Until mRs.EOF
        nId = mRs.Fields("id").Value
        colIds.Add nId
        mRs.MoveNext
    Loop
    
    Set colResult = mCtx.ExecuteRead(mModelName, colFieldNames, colIds).Result
    
    UpsertRecordSet colResult
    
    FetchRecordset = colResult.Count
End Function

Public Sub SetFilter(aExpr As String)
    mRs.Filter = aExpr
    SaftyMoveFirst
End Sub

Public Sub SetFilterMany2One(aId As Long)
    Debug.Assert IsMany2One
    mRs.Filter = "id = " & aId
    SaftyMoveFirst
End Sub

Public Sub SetFilterOne2Many(aId As Long)
    Debug.Assert IsOne2Many
    mRs.Filter = mFKeyFieldName & " = " & aId
    SaftyMoveFirst
End Sub

Public Sub SetFilterMany2Many(aId As Long)
    Dim colIds As Collection
    Dim a() As String
    Dim v As Variant
    Dim i As Long
    Dim s As String

    Debug.Assert IsMany2Many
    
    On Error Resume Next
    Set colIds = mRelationMap(CStr(aId))
    On Error GoTo 0
    If colIds Is Nothing Then
        mRs.Filter = "id = -1"  ' no records
    Else
        If colIds.Count > 0 Then
            ReDim a(1 To colIds.Count)
            For i = 1 To colIds.Count
                a(i) = "(id = " & CStr(colIds(i)) & ")"
            Next i
            s = Join(a, " OR ")
            mRs.Filter = s
        Else
            mRs.Filter = "id = -1"  ' no records
        End If
    End If
    SaftyMoveFirst
End Sub

Public Sub ClearFilter()
    mRs.Filter = adFilterNone
    SaftyMoveFirst
End Sub

Private Function GetFieldValueOrDefault(aFieldName As String, aDefault As Variant) As Variant
    Dim fld As ADODB.Field
    Set fld = mRs.Fields(aFieldName)
    If IsNull(fld.Value) Then
        GetFieldValueOrDefault = aDefault
    Else
        GetFieldValueOrDefault = fld.Value
    End If
End Function

Public Function GetRelatedModelView(aFieldName As String, Optional optNotFilterd As Boolean = False) As OdxModelView
    If Not mRelations.Exists(aFieldName) Then
        ' not exists
        Debug.Assert False
        GoTo ExitProc
    End If
    Set GetRelatedModelView = mRelations(aFieldName)
    If optNotFilterd Then
        GetRelatedModelView.ClearFilter
    Else
        With GetRelatedModelView
            If .IsMany2One Then
                .SetFilterMany2One GetFieldValueOrDefault(aFieldName, -1)
            ElseIf .IsOne2Many Then
                .SetFilterOne2Many GetFieldValueOrDefault("id", -1)
            ElseIf .IsMany2Many Then
                .SetFilterMany2Many GetFieldValueOrDefault("id", -1)
            Else
                ' not relations
                Debug.Assert False
                GoTo ExitProc
            End If
        End With
    End If
ExitProc:
    Exit Function
End Function

