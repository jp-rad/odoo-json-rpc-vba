Attribute VB_Name = "OdWebClientTest"
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
' External API - odoo docs
'
' Odoo is usually extended internally via modules, but many of its features and
' all of its data are also available from the outside for external analysis or
' integration with various tools. Part of the Models API is easily available over
' XML-RPC and accessible from a variety of languages.
'
' see also: https://www.odoo.com/documentation/master/developer/misc/api/odoo.html
'

Option Explicit

Private Const CBASEURL As String = "https://localhost"
Private Const CDBNAME As String = "dev_odoo"
Private Const CUSERNAME As String = "admin"
Private Const CPASSWORD As String = "admin"

Private mClient As OdWebClient

Public Sub doSetupClient(Optional TestDatabase As Boolean = True)
    Set mClient = NewOdWebClient()
    If TestDatabase Then
        ' Test database
        ' https://www.odoo.com/documentation/master/developer/misc/api/odoo.html#test-database
        With mClient.TestDatabase()
            mClient.BaseUrl = .sHost
            mClient.DbName = .sDatabase
            mClient.Username = .sUser
            mClient.Password = .sPassword
        End With
    Else
        ' Turn off SSL validation
        mClient.SetInsecure True
        ' Follow redirects (301, 302, 307) using Location header
        mClient.SetFollowRedirects False
        
        mClient.BaseUrl = CBASEURL
        mClient.DbName = CDBNAME
        mClient.Username = CUSERNAME
        mClient.Password = CPASSWORD
    End If
    
    Debug.Print "---------------"
    Debug.Print " Setup client"
    Debug.Print "---------------"
    Debug.Print "BaseUrl:", mClient.BaseUrl
    Debug.Print "Database:", mClient.DbName
    Debug.Print "Username:", mClient.Username
    Debug.Print "Password:", mClient.Password
    Debug.Print

End Sub

Public Sub doConnectToTest()
    doSetupClient True
End Sub

Public Sub doConnectToLocal()
    doSetupClient False
End Sub

Public Sub doAll()
On Error GoTo ErrHandler
    Debug.Print "Start - doAll()"
    Debug.Print "Press F5 key to step over next."
    If mClient.BaseUrl = "" Then
        Debug.Assert False
        doSetupClient
    End If
    Debug.Assert False
    doVersion
    Debug.Assert False
    doAuthenticate
    Debug.Assert False
    doCheckAccessRights
    Debug.Assert False
    doListRecords
    Debug.Assert False
    doPagination
    Debug.Assert False
    doCountRecords
    Debug.Assert False
    doReadRecords
    Debug.Assert False
    doListRecordFields
    Debug.Assert False
    doSearchAndRead
    Debug.Assert False
    doCreateUpdateUnlinkRecords
    Debug.Assert False
    doInspectionAndIntrospection
    Debug.Assert False
    Debug.Print "Done! Retry,"
    Debug.Print "doAll"
ExitProc:
    Exit Sub
ErrHandler:
    Debug.Assert False
    MsgBox Err.Description, vbCritical, Err.Source
    Resume Next
End Sub

Public Sub doVersion()
    Debug.Print Now() & " - doVersion"
    ' version
    ' https://www.odoo.com/documentation/master/developer/misc/api/odoo.html#logging-in
    With mClient.Common().Version()
        
        Debug.Print "---------"
        Debug.Print " version"
        Debug.Print "---------"
        Debug.Print JsonConverter.ConvertToJson(.JsonResult, 4)
        Debug.Print
        
    End With
End Sub

Public Sub doAuthenticate()
    Debug.Print Now() & " - doAuthenticate"
    ' version
    ' https://www.odoo.com/documentation/master/developer/misc/api/odoo.html#logging-in
    With mClient.Common().Authenticate()
        
        Debug.Print "--------------"
        Debug.Print " authenticate"
        Debug.Print "--------------"
        Debug.Print "result(uid): " & .JsonResult
        Debug.Print
        
    End With
End Sub

Public Sub doCheckAccessRights()
    Debug.Print Now() & " - doCheckAccessRights"
    ' Calling methods
    ' https://www.odoo.com/documentation/master/developer/misc/api/odoo.html#calling-methods
    With mClient.Model("res.partner").MethodCheckAccessRights("['read']", "{'raise_exception': false}")
        
        Debug.Print "-------------------"
        Debug.Print " CheckAccessRights"
        Debug.Print "-------------------"
        Debug.Print "result: " & .JsonResult
        Debug.Print
        
    End With
End Sub

Public Sub doListRecords()
    Debug.Print Now() & " - doListRecords"
    ' List records
    ' https://www.odoo.com/documentation/master/developer/misc/api/odoo.html#list-records
    With mClient.Model("res.partner").MethodSearch("[['is_company', '=', true]]")
        
        Debug.Print "-------------------"
        Debug.Print " List records"
        Debug.Print "-------------------"
        Debug.Print "result: " & JsonConverter.ConvertToJson(.JsonResult)
        Debug.Print
        
    End With
End Sub

Public Sub doPagination()
    Debug.Print Now() & " - doPagination"
    ' Pagination
    ' https://www.odoo.com/documentation/master/developer/misc/api/odoo.html#pagination
    With mClient.Model("res.partner").MethodSearch("[['is_company', '=', true]]", "{'offset': 10, 'limit': 5}")
        
        Debug.Print "-------------------"
        Debug.Print " Pagination"
        Debug.Print "-------------------"
        Debug.Print "result: " & JsonConverter.ConvertToJson(.JsonResult)
        Debug.Print
        
    End With
End Sub

Public Sub doCountRecords()
    Debug.Print Now() & " - doCountRecords"
    ' Count records
    ' https://www.odoo.com/documentation/master/developer/misc/api/odoo.html#count-records
    With mClient.Model("res.partner").ExecuteKw(OdRpc.CMETHOD_SEARCH_COUNT, "[[['is_company', '=', true]]]")
        
        Debug.Print "-------------------"
        Debug.Print " Count records"
        Debug.Print "-------------------"
        Debug.Print "result: " & .JsonResult
        Debug.Print
        
    End With
End Sub

Public Sub doReadRecords()
    Debug.Print Now() & " - doReadRecords"
    ' Read records
    ' https://www.odoo.com/documentation/master/developer/misc/api/odoo.html#read-records
    Dim ids As Collection
    Set ids = mClient.Model("res.partner").MethodSearch("[['is_company', '=', true]]", "{'limit': 1}").JsonResult
    With mClient.Model("res.partner").MethodRead(ids)
        
        Debug.Print "-------------------"
        Debug.Print " Read records"
        Debug.Print "-------------------"
        'Debug.Print "result: " & JsonConverter.ConvertToJson(.JsonResult)
        Debug.Print "result: " & .JsonResult.Item(1).Count
        Debug.Print
        
    End With
    With mClient.Model("res.partner").MethodRead(ids, "{'fields': ['name', 'country_id', 'comment']}")
        Debug.Print "result: " & JsonConverter.ConvertToJson(.JsonResult)
        Debug.Print
        
    End With
    
End Sub

Public Sub doListRecordFields()
    Debug.Print Now() & " - doListRecordFields"
    ' List record fields
    ' https://www.odoo.com/documentation/master/developer/misc/api/odoo.html#list-record-fields
    With mClient.Model("res.partner").ExecuteKw(OdRpc.CMETHOD_FIELDS_GET, "[]", "{'attributes': ['string', 'help', 'type']}")
        
        Debug.Print "-------------------"
        Debug.Print " List record fields"
        Debug.Print "-------------------"
        Debug.Print "result: " & JsonConverter.ConvertToJson(.JsonResult, 4)
        Debug.Print
        
    End With
End Sub

Public Sub doSearchAndRead()
    Debug.Print Now() & " - doSearchAndRead"
    ' Search and read
    ' https://www.odoo.com/documentation/master/developer/misc/api/odoo.html#search-and-read
    With mClient.Model("res.partner").MethodSearchAndRead("[['is_company', '=', true]]", " {'fields': ['name', 'country_id', 'comment'], 'limit': 5}")
        
        Debug.Print "-------------------"
        Debug.Print " Search and read"
        Debug.Print "-------------------"
        Debug.Print "result: " & JsonConverter.ConvertToJson(.JsonResult, 4)
        Debug.Print
        
    End With
End Sub

Public Sub doCreateUpdateUnlinkRecords()
    Debug.Print Now() & " - doCreateUpdateUnlinkRecords"
    ' Create records
    ' https://www.odoo.com/documentation/master/developer/misc/api/odoo.html#create-records
    Dim n As Long
    With mClient.Model("res.partner").MethodCreate("{'name': 'New Partner'}")
        n = .JsonResult
        
        Debug.Print "-------------------"
        Debug.Print " Create records"
        Debug.Print "-------------------"
        Debug.Print "result: " & .JsonResult
        Debug.Print
        
    End With
    ' Update records
    ' https://www.odoo.com/documentation/master/developer/misc/api/odoo.html#update-records
    With mClient.Model("res.partner").MethodWrite(n, "{'name': 'Newer Partner'}")
        
        Debug.Print "-------------------"
        Debug.Print " Update records"
        Debug.Print "-------------------"
        Debug.Print "result: " & .JsonResult
        Debug.Print
        
    End With
    ' name_get
    With mClient.Model("res.partner").MethodNameGet(n)
        
        Debug.Print "---------------------------"
        Debug.Print " Update records - name_get"
        Debug.Print "---------------------------"
        Debug.Print "result: " & JsonConverter.ConvertToJson(.JsonResult, 4)
        Debug.Print
        
    End With
    ' Delete records
    ' https://www.odoo.com/documentation/master/developer/misc/api/odoo.html#delete-records
On Error Resume Next
    With mClient.Model("res.partner").MethodUnlink(n)
        
        Debug.Print "-------------------"
        Debug.Print " Delete records"
        Debug.Print "-------------------"
        If Err.Number = 0 Then
            Debug.Print "result: " & .JsonResult
        Else
            Debug.Print "error: " & Err.Description
        End If
        'Debug.Print
        
    End With
    With NewOdDomainBuilder
        .AddCriteria("id").Eq n
        With mClient.Model("res.partner").MethodSearch(.GetDomain)
            Debug.Print "search: " & JsonConverter.ConvertToJson(.JsonResult)
            Debug.Print
        End With
    End With
On Error GoTo 0
    
End Sub

Public Sub doInspectionAndIntrospection()
    Debug.Print Now() & " - doInspectionAndIntrospection"
    ' Inspection and introspection
    ' https://www.odoo.com/documentation/master/developer/misc/api/odoo.html#inspection-and-introspection
    Dim s As String
    Dim n As Long
    ' - Custom model
    s = "{" _
        & vbCrLf & "'name': 'Custom Model'," _
        & vbCrLf & "'model': 'x_custom_vba_model'," _
        & vbCrLf & "'state': 'manual'" _
        & vbCrLf & "}"
On Error Resume Next
    With mClient.ModelOfIrModel.MethodCreate(s)
        n = .JsonResult
        Debug.Print "-------------------"
        Debug.Print " Custom model"
        Debug.Print "-------------------"
        If Err.Number = 0 Then
            Debug.Print "result: " & .JsonResult
        Else
            Debug.Print "error: " & Err.Description
        End If
        Debug.Print
        
    End With
On Error GoTo 0
    With mClient.Model("x_custom_vba_model").ExecuteKw(OdRpc.CMETHOD_FIELDS_GET, "[]", "{'attributes': ['string', 'help', 'type']}")
        Debug.Print "fields_get: " & JsonConverter.ConvertToJson(.JsonResult, 4)
        Debug.Print
    End With
    ' - Custom field
    s = "{" _
        & vbCrLf & "'name': 'Custom Model'," _
        & vbCrLf & "'model': 'x_custom_vba'," _
        & vbCrLf & "'state': 'manual'" _
        & vbCrLf & "}"
On Error Resume Next
    Err.Clear
    With mClient.ModelOfIrModel.MethodCreate(s)
        n = .JsonResult
        Debug.Print "-------------------"
        Debug.Print " Custom field"
        Debug.Print "-------------------"
        If Err.Number = 0 Then
            Debug.Print "model: " & .JsonResult
        Else
            Debug.Print "error: " & Err.Description
        End If
        Debug.Print
        
    End With
    s = "{" _
    & vbCrLf & "'model_id': " & n & "," _
    & vbCrLf & "'name': 'x_name_vba'," _
    & vbCrLf & "'ttype': 'char'," _
    & vbCrLf & "'state': 'manual'," _
    & vbCrLf & "'required': true" _
    & vbCrLf & "}"
    Err.Clear
    With mClient.ModelOfIrModelFields.MethodCreate(s)
        If Err.Number = 0 Then
            Debug.Print "field: " & .JsonResult
        Else
            Debug.Print "error: " & Err.Description
        End If
        Debug.Print
    End With
On Error GoTo 0
    With mClient.Model("x_custom_vba").MethodCreate("{'x_name_vba': 'test record'}")
        n = .JsonResult
        
        Debug.Print "record_id: " & .JsonResult
        
    End With
    With NewOdDomainBuilder
        .AddCriteria("id").Eq n
        With mClient.Model("x_custom_vba").MethodSearch(.GetDomain)
            Debug.Print "search: " & JsonConverter.ConvertToJson(.JsonResult)
            Debug.Print
        End With
    End With
End Sub

Public Sub testDomainBuilder()
    With NewOdDomainBuilder
        
        ' criteria
        .AddCriteria("[=] equals to").Eq 1
        .AddCriteria("[!=] not equals to").NotEq 2
        .AddCriteria("[>] greater than").Gt 3
        .AddCriteria("[>=] greater than or equal to").Ge 4
        .AddCriteria("[<] less than").Lt 5
        .AddCriteria("[<=] less than or equal to").Le 6
        .AddCriteria("[=?] unset or equals to").UnsetOrEq 7
        
        .AddLogicalAnd
        
        .AddCriteria("[=like] matches field_name against the value pattern").EqLike 8
        .AddCriteria("[like] matches field_name against the %value% pattern").IsLike 9
        .AddCriteria("[not like] doesnft match against the %value% pattern").NotLike 10
        
        .AddLogicalOr
        
        .AddCriteria("[ilike] case insensitive like").CILike 11
        .AddCriteria("[not ilike] case insensitive not like").NotCILike 12
        .AddCriteria("[=ilike] case insensitive =like").EqCILike 13
        
        .AddLogicalNot
        
        .AddCriteria("[in] is equal to any of the items from value").IsIn 14
        .AddCriteria("[not in] is unequal to all of the items from value").NotIn 15
        .AddCriteria("[child_of] is a child (descendant) of a value record").ChildOf 16
        .AddCriteria("[parent_of] is a parent (ascendant) of a value record").ParentOf 17
        
        ' New Collection
        With .NewCollection()
            With .NewCollection()
                .AddLogicalOr
                Call .NewCollection().AddLogicalAnd().AddCriteria("flag1").Eq(True).AddCriteria("field1").Eq("a")
                Call .NewCollection().AddLogicalAnd().AddCriteria("flag2").Eq(False).AddCriteria("field2").Eq("b")
            End With
            .AddLogicalNot
            .AddCriteria("active").Eq False
        End With
        
        ' result
        Debug.Print
        Debug.Print ">>>>", "testDomainBuilder"
        Debug.Print JsonConverter.ConvertToJson(.GetDomain(), 2)
        'Debug.Print JsonConverter.ConvertToJson(.GetDomain())
        Debug.Print "<<<<"
        Debug.Print
        
    End With
End Sub

