Attribute VB_Name = "OdooJsonRpcTest"
' External API - odoo-JSON-RPC-VBA
'
' MIT License
'
' Copyright (c) 2022 jp-rad
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
' see also: https://www.odoo.com/documentation/15.0/developer/misc/api/odoo.html
'

Option Explicit

Private mClient As New OdWebClient

' Test database
Private mUrl As String
Private mDb As String
Private mUserName As String
Private mPassword As String

Public Sub doAll()
    Debug.Print "Start - doAll()"
    Debug.Print "Press F5 key to step over next."
    Debug.Assert False
    doTestDatabase
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
    Debug.Print "Done! Retry,"
    Debug.Print "doAll"
End Sub

Public Sub doTestDatabase()
    Debug.Print Now() & " - exTestDatabase"
    ' Test database
    ' https://www.odoo.com/documentation/15.0/developer/misc/api/odoo.html#test-database
    With mClient.Start().Start()
        
        mUrl = .SrHost
        mDb = .SrDatabase
        mUserName = .SrUser
        mPassword = .SrPassword
        
        Debug.Print "---------------"
        Debug.Print " Test Database"
        Debug.Print "---------------"
        Debug.Print JsonConverter.ConvertToJson(.JsonResult)
        Debug.Print
        
    End With
End Sub

Public Sub doVersion()
    Debug.Print Now() & " - exVersion"
    ' version
    ' https://www.odoo.com/documentation/15.0/developer/misc/api/odoo.html#logging-in
    mClient.BaseUrl = mUrl
    With mClient.Common().Version()
        
        Debug.Print "---------"
        Debug.Print " version"
        Debug.Print "---------"
        Debug.Print JsonConverter.ConvertToJson(.JsonResult, 4)
        Debug.Print
        
    End With
End Sub

Public Sub doAuthenticate()
    Debug.Print Now() & " - exAuthenticate"
    ' version
    ' https://www.odoo.com/documentation/15.0/developer/misc/api/odoo.html#logging-in
    mClient.BaseUrl = mUrl
    With mClient.Common().Authenticate(mDb, mUserName, mPassword)
        
        Debug.Print "--------------"
        Debug.Print " authenticate"
        Debug.Print "--------------"
        Debug.Print "result(uid): " & .JsonResult
        Debug.Print
        
    End With
End Sub

Public Sub doCheckAccessRights()
    Debug.Print Now() & " - exCheckAccessRights"
    ' Calling methods
    ' https://www.odoo.com/documentation/15.0/developer/misc/api/odoo.html#calling-methods
    With mClient.Model("res.partner").MethodCheckAccessRights("['read']", "{'raise_exception': false}")
        
        Debug.Print "-------------------"
        Debug.Print " CheckAccessRights"
        Debug.Print "-------------------"
        Debug.Print "result: " & .JsonResult
        Debug.Print
        
    End With
End Sub

Public Sub doListRecords()
    Debug.Print Now() & " - exListRecords"
    ' List records
    ' https://www.odoo.com/documentation/15.0/developer/misc/api/odoo.html#list-records
    With mClient.Model("res.partner").MethodSearch("[['is_company', '=', true]]")
        
        Debug.Print "-------------------"
        Debug.Print " List records"
        Debug.Print "-------------------"
        Debug.Print "result: " & JsonConverter.ConvertToJson(.JsonResult)
        Debug.Print
        
    End With
End Sub

Public Sub doPagination()
    Debug.Print Now() & " - exPagination"
    ' Pagination
    ' https://www.odoo.com/documentation/15.0/developer/misc/api/odoo.html#pagination
    With mClient.Model("res.partner").MethodSearch("[['is_company', '=', true]]", "{'offset': 10, 'limit': 5}")
        
        Debug.Print "-------------------"
        Debug.Print " Pagination"
        Debug.Print "-------------------"
        Debug.Print "result: " & JsonConverter.ConvertToJson(.JsonResult)
        Debug.Print
        
    End With
End Sub

Public Sub doCountRecords()
    Debug.Print Now() & " - exCountRecords"
    ' Count records
    ' https://www.odoo.com/documentation/15.0/developer/misc/api/odoo.html#count-records
    With mClient.Model("res.partner").ExecuteKw(CMETHOD_SEARCH_COUNT, "[[['is_company', '=', true]]]")
        
        Debug.Print "-------------------"
        Debug.Print " Count records"
        Debug.Print "-------------------"
        Debug.Print "result: " & .JsonResult
        Debug.Print
        
    End With
End Sub

Public Sub doReadRecords()
    Debug.Print Now() & " - exReadRecords"
    ' Read records
    ' https://www.odoo.com/documentation/15.0/developer/misc/api/odoo.html#read-records
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
    Debug.Print Now() & " - exListRecordFields"
    ' List record fields
    ' https://www.odoo.com/documentation/15.0/developer/misc/api/odoo.html#list-record-fields
    With mClient.Model("res.partner").ExecuteKw(CMETHOD_FIELDS_GET, "[]", "{'attributes': ['string', 'help', 'type']}")
        
        Debug.Print "-------------------"
        Debug.Print " List record fields"
        Debug.Print "-------------------"
        Debug.Print "result: " & JsonConverter.ConvertToJson(.JsonResult, 4)
        Debug.Print
        
    End With
End Sub

Public Sub doSearchAndRead()
    Debug.Print Now() & " - exSearchAndRead"
    ' Search and read
    ' https://www.odoo.com/documentation/15.0/developer/misc/api/odoo.html#search-and-read
    With mClient.Model("res.partner").MethodSearchAndRead("[['is_company', '=', true]]", " {'fields': ['name', 'country_id', 'comment'], 'limit': 5}")
        
        Debug.Print "-------------------"
        Debug.Print " Search and read"
        Debug.Print "-------------------"
        Debug.Print "result: " & JsonConverter.ConvertToJson(.JsonResult, 4)
        Debug.Print
        
    End With
End Sub

' Create records
' Update records
' Delete records
' Inspection and introspection
' - Custom model
' - Custom field

Public Sub testDomainBuilder()
    With New OdDomainBuilder
        
        ' criteria
        .AddCriteriaEq "[=] equals to", 1
        .AddCriteriaNotEq "[!=] not equals to", 2
        .AddCriteriaGt "[>] greater than", 3
        .AddCriteriaGe "[>=] greater than or equal to", 4
        .AddCriteriaLt "[<] less than", 5
        .AddCriteriaLe "[<=] less than or equal to", 6
        .AddCriteriaUnsetOrEq "[=?] unset or equals to", 7
        
        .AddLogicalAnd
        
        .AddCriteriaEqLike "[=like] matches field_name against the value pattern", 8
        .AddCriteriaLike "[like] matches field_name against the %value% pattern", 9
        .AddCriteriaNotLike "[not like] doesnft match against the %value% pattern", 10
        
        .AddLogicalOr
        
        .AddCriteriaCILike "[ilike] case insensitive like", 11
        .AddCriteriaNotCILike "[not ilike] case insensitive not like", 12
        .AddCriteriaEqCILike "[=ilike] case insensitive =like", 13
        
        .AddLogicalNot
        
        .AddCriteriaIn "[in] is equal to any of the items from value", 14
        .AddCriteriaNotIn "[not in] is unequal to all of the items from value", 15
        .AddCriteriaChildOf "[child_of] is a child (descendant) of a value record", 16
        .AddCriteriaParentOf "[parent_of] is a parent (ascendant) of a value record", 17
        
        ' New Collection
        With .NewCollection()
            With .NewCollection()
                .AddLogicalOr
                Call .NewCollection().AddLogicalAnd().AddCriteriaEq("flag1", True).AddCriteriaEq("field1", "a")
                Call .NewCollection().AddLogicalAnd().AddCriteriaEq("flag2", False).AddCriteriaEq("field2", "b")
            End With
            .AddLogicalNot
            .AddCriteriaEq "active", False
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

