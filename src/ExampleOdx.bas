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
    Dim oClient As OdClient
    Dim oCtx As OdxContext
    Dim oMainView As OdxModelView
    Dim mv As OdxModelView
    Dim wb As Workbook
    Dim sht As Worksheet
    Dim rng As Range
    
    Set oClient = GetAuthConn()
    Set oCtx = NewContext(oClient)
    Set wb = Workbooks.Add
    Set sht = wb.Sheets(1)
    SetWorksheetName sht, "OdxModelView Example"
    Set rng = sht.Range("A1")
    
    DebugRange rng, "--------------"
    DebugRange rng, " DoExampleOdx"
    DebugRange rng, "--------------"
    
    ' ==================
    '  New OdxModelView
    ' ==================
    Set oMainView = oCtx.NewModelView("res.partner")
    With oMainView
        ' -------------
        '  res.partner
        ' -------------
        Debug.Print .ModelName
        DebugRange rng, .ModelName
        
        ' .AddField "id"
        .AddField "name"
        
        With .AddField("company_id")    ' many2one
            ' --------------------------
            '  res.company
            ' --------------------------
            DebugRange rng, "company_id: many2one --> " & .ModelName
            Debug.Assert .IsMany2One
            
            .AddField "name"
            .AddField "city"
        End With
        
        With .AddField("child_ids")     ' one2many
            ' --------------------------
            '  res.partner
            ' --------------------------
            DebugRange rng, "child_ids: one2many --> " & .ModelName
            Debug.Assert .IsOne2Many
            
            .AddField "name"
        End With
                
        With .AddField("category_id")   ' many2many
            ' --------------------------
            '  res.partner.category
            ' --------------------------
            DebugRange rng, "category_id: many2many --> " & .ModelName
            Debug.Assert .IsMany2Many
            
            .AddField "name"
        End With
    End With

    ' ==================
    '  Fetch data
    ' ==================
    oMainView.ExecuteSearchRead NewDomain
    
    ' ==================
    '  Filtered
    ' ==================
    oMainView.SetFilter "name = 'Gemini Furniture'"
    If oMainView.Recordset.EOF Then
        Debug.Print "No record"
        Debug.Assert False
    Else
        Set mv = oMainView
        DebugPrintModelView wb, mv, "Filtered"
        Set mv = oMainView.GetRelatedModelView("company_id")
        DebugPrintModelView wb, mv, "Filtered(company_id)"
        Set mv = oMainView.GetRelatedModelView("child_ids")
        DebugPrintModelView wb, mv, "Filtered(child_ids)"
        Set mv = oMainView.GetRelatedModelView("category_id")
        DebugPrintModelView wb, mv, "Filtered(category_id)"
    End If
    
    ' ==================
    '  All
    ' ==================
    oMainView.ClearFilter
    
    Set mv = oMainView
    DebugPrintModelView wb, mv
    
    Set mv = oMainView.GetRelatedModelView("company_id", True)
    DebugPrintModelView wb, mv
    
    Set mv = oMainView.GetRelatedModelView("child_ids", True)
    DebugPrintModelView wb, mv
    
    Set mv = oMainView.GetRelatedModelView("category_id", True)
    DebugPrintModelView wb, mv
    
    sht.Select
    wb.Saved = True
    wb.Activate
End Sub

Private Sub DebugRange(ByRef rng As Range, v As Variant)
    Debug.Print v
    rng = v
    Set rng = rng.Offset(1, 0)
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

