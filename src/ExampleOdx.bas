Attribute VB_Name = "ExampleOdx"
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

Public Sub DoExampleOdx()
    Dim oModelView As OdxModelView
    Dim oDomain As OdFilterDomain
    
    Dim oClient As OdClient
    Dim oCtx As OdxContext
    
    Dim wb As Workbook
    Dim sht As Worksheet
    Dim rng As Range
    
    Dim colNameList As New Collection
    Dim colTagList As New Collection
    Dim v As Variant
    Dim s As String
    
    Set oClient = GetAuthConn()
    Set oCtx = NewContext(oClient)
    Set wb = Workbooks.Add
    Set sht = wb.Sheets(1)
    SetWorksheetName sht, "OdxModelView Example"
    Set rng = sht.Range("A1")
    
    s = "Gemini"
    colTagList.Add s
    colNameList.Add "Gemini Furniture", s
    s = "Mitchell"
    colTagList.Add s
    colNameList.Add "Mitchell Admin", s
    
    DebugRange rng, "--------------"
    DebugRange rng, " DoExampleOdx"
    DebugRange rng, "--------------"
    
    ' ==================
    '  New OdxModelView
    ' ==================
    Set oModelView = oCtx.NewModelView("res.partner")
    With oModelView
        ' -------------
        '  res.partner
        ' -------------
        DebugRange rng, .ModelName
        DebugRangeSchema rng, .ModelName, .GetModelSchema()
        
        .AddField "name"
        .AddField "is_company"
        .AddField "is_public"
        
        With .AddField("company_id")    ' many2one
            ' --------------------------
            '  res.company
            ' --------------------------
            DebugRange rng, "company_id: many2one --> " & .ModelName
            Debug.Assert .IsMany2One
            DebugRangeSchema rng, .ModelName, .GetModelSchema()
            
            .AddField "name"
            .AddField "city"
            
            With .AddField("currency_id")   ' many2one
                ' --------------------------
                '  res.currency
                ' --------------------------
                DebugRange rng, "company_id: many2one --> currency_id: many2one --> " & .ModelName
                Debug.Assert .IsMany2One
                DebugRangeSchema rng, .ModelName, .GetModelSchema()
                
                .AddField "name"
                .AddField "symbol"

            End With
            
            .AddField "layout_background"

        End With
        
        With .AddField("child_ids")     ' one2many
            ' --------------------------
            '  res.partner
            ' --------------------------
            DebugRange rng, "child_ids: one2many --> " & .ModelName
            Debug.Assert .IsOne2Many
            DebugRangeSchema rng, .ModelName, .GetModelSchema()
            
            .AddField "name"
        End With
                
        With .AddField("category_id")   ' many2many
            ' --------------------------
            '  res.partner.category
            ' --------------------------
            DebugRange rng, "category_id: many2many --> " & .ModelName
            Debug.Assert .IsMany2Many
            DebugRangeSchema rng, .ModelName, .GetModelSchema()
            
            .AddField "name"
        End With
    End With

    ' ==================
    '  Fetch data
    ' ==================
    Set oDomain = NewDomain
    oModelView.ExecuteSearchRead oDomain
    
    ' ==================
    '  All
    ' ==================
    oModelView.ClearFilter  ' Unfiltered
    With oModelView
        DebugPrintModelView wb, .RefMe
        With .GetRelatedModelView("company_id", True)
            DebugPrintModelView wb, .RefMe
            With .GetRelatedModelView("currency_id", True)
                DebugPrintModelView wb, .RefMe
            End With
        End With
        With .GetRelatedModelView("child_ids", True)
            DebugPrintModelView wb, .RefMe
        End With
        With .GetRelatedModelView("category_id", True)
            DebugPrintModelView wb, .RefMe
        End With
    End With
    
    
    ' ==================
    '  Filtered
    ' ==================
    For Each v In colTagList
        s = colNameList(CStr(v))
        
        oModelView.SetFilter "name = '" & s & "'"   ' Filtered
        With oModelView
            DebugPrintModelView wb, .RefMe, CStr(v)
            With .GetRelatedModelView("company_id")
                DebugPrintModelView wb, .RefMe, v & "(company_id)"
                With .GetRelatedModelView("currency_id")
                    DebugPrintModelView wb, .RefMe, v & "(currency_id)"
                End With
            End With
            With .GetRelatedModelView("child_ids")
                DebugPrintModelView wb, .RefMe, v & "(child_ids)"
            End With
            With .GetRelatedModelView("category_id")
                DebugPrintModelView wb, .RefMe, v & "(category_id)"
            End With
        End With
    Next v
    
    sht.Select
    wb.Saved = True
    wb.Activate
End Sub

Private Sub DebugRange(ByRef rng As Range, v As Variant)
    Debug.Print v
    rng = v
    Set rng = rng.Offset(1, 0)
End Sub

Private Sub DebugRangeSchema(ByRef rng As Range, aModelName As String, dicModelSchema As Dictionary)
    Debug.Print JsonConverter.ConvertToJson(dicModelSchema)
    
    Dim tag As String
    Dim i As Long
    Dim dic As Dictionary
    
    tag = aModelName
    
    ' header
    rng.Offset(0, 0) = tag
    rng.Offset(0, 1) = "#"
    rng.Offset(0, 2) = OdxApi.CODOO_ATTR_NAME
    rng.Offset(0, 3) = OdxApi.CODOO_ATTR_REQUIRED
    rng.Offset(0, 4) = OdxApi.CODOO_ATTR_TYPE
    rng.Offset(0, 5) = OdxApi.CODOO_ATTR_RELATION
    rng.Offset(0, 6) = OdxApi.CODOO_ATTR_FKEY
    rng.Offset(0, 7) = OdxApi.CODOO_ATTR_STRING
    ' next row
    Set rng = rng.Offset(1, 0)
    
    ' fields
    For i = 1 To dicModelSchema.Count
        Set dic = dicModelSchema.Items(i - 1)
        rng.Offset(0, 0) = tag
        rng.Offset(0, 1) = i
        rng.Offset(0, 2) = dic(OdxApi.CODOO_ATTR_NAME)
        rng.Offset(0, 3) = dic(OdxApi.CODOO_ATTR_REQUIRED)
        rng.Offset(0, 4) = dic(OdxApi.CODOO_ATTR_TYPE)
        rng.Offset(0, 5) = dic(OdxApi.CODOO_ATTR_RELATION)
        rng.Offset(0, 6) = dic(OdxApi.CODOO_ATTR_FKEY)
        rng.Offset(0, 7) = dic(OdxApi.CODOO_ATTR_STRING)
        ' next row
        Set rng = rng.Offset(1, 0)
    Next i
End Sub

Private Function TrySetWorksheetName(aWs As Worksheet, aName As String) As Boolean
On Error GoTo ErrHandler
    aWs.Name = aName
    TrySetWorksheetName = True
ExitProc:
    Exit Function
ErrHandler:
    Resume ExitProc
End Function

Private Function SetWorksheetName(aWs As Worksheet, aName As String)
    Dim i As Long
    If Not TrySetWorksheetName(aWs, aName) Then
        i = 2
        Do Until TrySetWorksheetName(aWs, aName & " (" & i & ")")
            i = i + 1
        Loop
    End If
End Function

Private Sub DebugPrintModelView(aWb As Workbook, aMv As OdxModelView, Optional optSheetName As String = "")
    Dim rs As ADODB.Recordset
    Dim sht As Worksheet
    Dim rng As Range
    Dim i As Long
    
    Dim sSheetName As String
    
    ' Clone and Filtered
    Set rs = aMv.Recordset.Clone
    rs.Filter = aMv.Recordset.Filter
    
    If optSheetName = "" Then
        sSheetName = aMv.ModelName
    Else
        sSheetName = optSheetName
    End If
    Set sht = aWb.Sheets.Add(After:=aWb.Worksheets(aWb.Worksheets.Count))
    SetWorksheetName sht, sSheetName
    
    Set rng = sht.Range("A1")
    For i = 0 To rs.Fields.Count - 1
        rng.Offset(0, i) = rs.Fields(i).Name
    Next i
    
    If Not rs.BOF Then
        rs.MoveFirst
    End If
    rng.Offset(1, 0).CopyFromRecordset rs
    
End Sub

