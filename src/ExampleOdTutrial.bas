Attribute VB_Name = "ExampleOdTutrial"
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

' https://www.odoo.com/documentation/master/developer/reference/external_api.html#calling-methods
Public Sub DoTutorialExternalApi()
    Dim oClient As OdClient
    Dim oRet As OdResult
    Dim params As Collection
    Dim named As Dictionary
    Dim sJson As String
    Dim nId As Long

    ' Version
    Set oRet = GetCommonVersion()
    Debug.Print "---------"
    Debug.Print " version"
    Debug.Print "---------"
    Debug.Print JsonConverter.ConvertToJson(oRet.Result, 4)
    Debug.Print

    ' Logging in - Authenticate
    Set oClient = GetAuthConn()
    
    ' execute_kw
    ' python: models.execute_kw(db, uid, password, 'res.partner', 'name_search', ['foo'], {'limit': 10})
    Set params = NewList    ' ['foo']
    params.Add "Azure" ' "foo"
    Set named = NewDict     ' {'limit': 10}
    named.Add "limit", 10
    Set oRet = oClient.Model("res.partner").Method("name_search").ExecuteKw(params, named)
    Debug.Print "------------"
    Debug.Print " execute_kw"
    Debug.Print "------------"
    Debug.Print JsonConverter.ConvertToJson(oRet.Result, 4)
    Debug.Print
    
    ' List records - search
    ' python: models.execute_kw(db, uid, password, 'res.partner', 'search', [[['is_company', '=', True]]])
    Set params = NewList  ' [[['is_company', '=', True]]]
    With NewDomain
        .AddArity NewField("is_company").Eq(True)
        .BuildAndAppend params
    End With
    Set oRet = oClient.Model("res.partner").Method("search").ExecuteKw(params)
    Debug.Print "--------------"
    Debug.Print " List records"
    Debug.Print "--------------"
    Debug.Print JsonConverter.ConvertToJson(oRet.Result, 4)
    Debug.Print
    
    ' Pagination - search
    ' python: models.execute_kw(db, uid, password, 'res.partner', 'search', [[['is_company', '=', True]]], {'offset': 10, 'limit': 5})
    Set params = NewList ' [[['is_company', '=', True]]]
    With NewDomain
        .AddArity NewField("is_company").Eq(True)
        .BuildAndAppend params
    End With
    Set named = NewDict  ' {'offset': 10, 'limit': 5}
    With named
        .Add "offset", 3 ' 10
        .Add "limit", 5
    End With
    Set oRet = oClient.Model("res.partner").Method("search").ExecuteKw(params, named)
    Debug.Print "------------"
    Debug.Print " Pagination"
    Debug.Print "------------"
    Debug.Print JsonConverter.ConvertToJson(oRet.Result, 4)
    Debug.Print
    
    ' Count records - search_count
    ' python: models.execute_kw(db, uid, password, 'res.partner', 'search_count', [[['is_company', '=', True]]])
    Set params = NewList    ' [[['is_company', '=', True]]]
    With NewDomain
        .AddArity NewField("is_company").Eq(True)
        .BuildAndAppend params
    End With
    Set oRet = oClient.Model("res.partner").Method("search_count").ExecuteKw(params)
    Debug.Print "---------------"
    Debug.Print " Count records"
    Debug.Print "---------------"
    Debug.Print JsonConverter.ConvertToJson(oRet.Result, 4)
    Debug.Print
    
    ' Read records - search, read
    ' python: ids = models.execute_kw(db, uid, password, 'res.partner', 'search', [[['is_company', '=', True]]], {'limit': 1})
    '         [record] = models.execute_kw(db, uid, password, 'res.partner', 'read', [ids])
    '         # count the number of fields fetched by default
    '         len(record)
    '         models.execute_kw(db, uid, password, 'res.partner', 'read', [ids], {'fields': ['name', 'country_id', 'comment']})
    Set params = NewList ' [[['is_company', '=', True]]]
    With NewDomain
        .AddArity NewField("is_company").Eq(True)
        .BuildAndAppend params
    End With
    Set named = NewDict ' {'limit': 1}
    named.Add "limit", 1
    Set oRet = oClient.Model("res.partner").Method("search").ExecuteKw(params, named)
    nId = oRet.Result(1)
    Set params = NewList ' [ids]
    params.Add nId
    Set oRet = oClient.Model("res.partner").Method("read").ExecuteKw(params)
    Debug.Print "--------------"
    Debug.Print " Read records"
    Debug.Print "--------------"
    Debug.Print "ids:", nId
    Debug.Print JsonConverter.ConvertToJson(oRet.Result, 4)
    Debug.Print
    
    Set named = NewDict ' {'fields': ['name', 'country_id', 'comment']}
    With SetList(named, "fields")
        .Add "name"
        .Add "country_id"
        .Add "comment"
    End With
    Set oRet = oClient.Model("res.partner").Method("read").ExecuteKw(params, named)
    Debug.Print "----- fields -----"
    Debug.Print JsonConverter.ConvertToJson(oRet.Result, 4)
    Debug.Print
    
    ' List record fields - fields_get
    ' python: models.execute_kw(db, uid, password, 'res.partner', 'fields_get', [], {'attributes': ['string', 'help', 'type']})
    Set named = NewDict ' {'attributes': ['string', 'help', 'type']}
    With SetList(named, "attributes")
        .Add "string"
        .Add "help"
        .Add "type"
    End With
    Set oRet = oClient.Model("res.partner").Method("fields_get").ExecuteKw(aNamedParams:=named)
    Debug.Print "--------------------"
    Debug.Print " List record fields"
    Debug.Print "--------------------"
    Debug.Print JsonConverter.ConvertToJson(oRet.Result, 4)
    Debug.Print
    
    ' Search and read - search_read
    ' python: models.execute_kw(db, uid, password, 'res.partner', 'search_read', [[['is_company', '=', True]]], {'fields': ['name', 'country_id', 'comment'], 'limit': 5})
    Set params = NewList ' [[['is_company', '=', True]]]
    With NewDomain
        .AddArity NewField("is_company").Eq(True)
        .BuildAndAppend params
    End With
    Set named = NewDict  ' {'fields': ['name', 'country_id', 'comment'], 'limit': 5}
    With SetList(named, "fields")
        .Add "name"
        .Add "country_id"
        .Add "comment"
    End With
    named.Add "limit", 5
    Set oRet = oClient.Model("res.partner").Method("search_read").ExecuteKw(params, named)
    Debug.Print "-----------------"
    Debug.Print " Search and read"
    Debug.Print "-----------------"
    Debug.Print JsonConverter.ConvertToJson(oRet.Result, 4)
    Debug.Print
    
    ' Create records - create
    ' python: id = models.execute_kw(db, uid, password, 'res.partner', 'create', [{'name': "New Partner"}])
    Set params = NewList ' [{'name': "New Partner"}]
    With AddDict(params)
        .Add "name", "New Partner"
    End With
    Set oRet = oClient.Model("res.partner").Method("create").ExecuteKw(params)
    nId = oRet.Result
    Debug.Print "----------------"
    Debug.Print " Create records"
    Debug.Print "----------------"
    Debug.Print "id:", nId
    Debug.Print
    
    ' Update records - write
    ' python: models.execute_kw(db, uid, password, 'res.partner', 'write', [[id], {'name': "Newer partner"}])
    '         # get record name after having changed it
    '         models.execute_kw(db, uid, password, 'res.partner', 'read', [[id], ['display_name']])
    Set params = NewList    ' [[id], {'name': "Newer partner"}]
    With AddList(params)  ' [id]
        .Add nId
    End With
    With AddDict(params)  ' {'name': "Newer partner"}
        .Add "name", "Newer parther"
    End With
    Set oRet = oClient.Model("res.partner").Method("write").ExecuteKw(params)
    Debug.Print "----------------"
    Debug.Print " Update records"
    Debug.Print "----------------"
    Debug.Print JsonConverter.ConvertToJson(oRet.Result, 4)
    Debug.Print
    ' # get record name after having changed it
    Set params = NewList    ' [[id], ['display_name']]
    With AddList(params)    ' [id]
        .Add nId
    End With
    With AddList(params)    ' ['display_name']
        .Add "display_name"
    End With
    Set oRet = oClient.Model("res.partner").Method("read").ExecuteKw(params)
    Debug.Print JsonConverter.ConvertToJson(oRet.Result, 4)
    Debug.Print
    
    ' Delete records - unlink
    ' python: models.execute_kw(db, uid, password, 'res.partner', 'unlink', [[id]])
    '         # check if the deleted record is still in the database
    '         models.execute_kw(db, uid, password, 'res.partner', 'search', [[['id', '=', id]]])
    Set params = NewList ' [[id]]
    With AddList(params)
        .Add nId
    End With
    Set oRet = oClient.Model("res.partner").Method("unlink").ExecuteKw(params)
    Debug.Print "----------------"
    Debug.Print " Delete records"
    Debug.Print "----------------"
    Debug.Print JsonConverter.ConvertToJson(oRet.Result, 4)
    Debug.Print
    ' # check if the deleted record is still in the database
    Set params = NewList ' [[['id', '=', id]]]
    With NewDomain()
        .AddArity NewField("id").Eq(nId)
        .BuildAndAppend params
    End With
    Set oRet = oClient.Model("res.partner").Method("search").ExecuteKw(params)
    Debug.Print JsonConverter.ConvertToJson(oRet.Result, 4)
    Debug.Print
    
    ' Inspection and introspection - delete models, x_custom_model and x_custom
    Set params = NewList
    With NewDomain()
        .AddArity NewField("model").IsILike("x_custom")
        .BuildAndAppend params
    End With
    Set oRet = oClient.ModelOfIrModel.Method("search").ExecuteKw(params)
    Set params = NewList
    params.Add oRet.Result
    Set oRet = oClient.ModelOfIrModel.Method("unlink").ExecuteKw(params)
    
    ' Inspection and introspection - ir.model, fields_get
    ' python: models.execute_kw(db, uid, password, 'ir.model', 'create', [{
    '           'name': 'Custom Model',
    '           'model': 'x_custom_model',
    '           'state': 'manual',
    '         }])
    '         models.execute_kw(db, uid, password, 'x_custom_model', 'fields_get', [], {'attributes': ['string', 'help', 'type']})
    Set params = NewList ' [{'name':"Custom Model", 'model':"x_custom_model", 'state': 'manual',}]
    With AddDict(params)
        .Add "name", "Custom Model"
        .Add "model", "x_custom_model"
        .Add "state", "manual"
    End With
    Set oRet = oClient.ModelOfIrModel.Method("create").ExecuteKw(params)
    nId = oRet.Result
    Debug.Print "------------------------------"
    Debug.Print " Inspection and introspection"
    Debug.Print "------------------------------"
    Debug.Print JsonConverter.ConvertToJson(oRet.Result, 4)
    Debug.Print
    Set named = NewDict  ' {'attributes': ['string', 'help', 'type']}
    With SetList(named, "attributes")
        .Add "string"
        .Add "help"
        .Add "type"
    End With
    Set oRet = oClient.Model("x_custom_model").Method("fields_get").ExecuteKw(aNamedParams:=named)
    Debug.Print JsonConverter.ConvertToJson(oRet.Result, 4)
    Debug.Print
        
    ' Inspection and introspection - ir.model.fields
    ' python: id = models.execute_kw(db, uid, password, 'ir.model', 'create', [{
    '           'name': "Custom Model",
    '           'model': "x_custom",
    '           'state': 'manual',
    '         }])
    '         models.execute_kw(db, uid, password, 'ir.model.fields', 'create', [{
    '           'model_id': id,
    '           'name': 'x_name',
    '           'ttype': 'char',
    '           'state': 'manual',
    '           'required': True,
    '         }])
    '         record_id = models.execute_kw(db, uid, password, 'x_custom', 'create', [{'x_name': "test record"}])
    '         models.execute_kw(db, uid, password, 'x_custom', 'read', [[record_id]])
    ' ir.model - create
    Set params = NewList ' [{'name':"Custom Model", 'model':"x_custom", 'state': 'manual',}]
    With AddDict(params)
        .Add "name", "Custom"
        .Add "model", "x_custom"
        .Add "state", "manual"
    End With
    Set oRet = oClient.ModelOfIrModel.Method("create").ExecuteKw(params)
    nId = oRet.Result
    Debug.Print "------------------------------"
    Debug.Print " Inspection and introspection"
    Debug.Print "------------------------------"
    Debug.Print JsonConverter.ConvertToJson(oRet.Result, 4)
    Debug.Print
    ' ir.model.access - create
    Set params = NewList
    With AddDict(params)
        .Add "id", "acl.x_custom"
        .Add "name", "acl.x_custom"
        .Add "model_id", nId
        ' .Add "group_id", ""
        .Add "perm_read", 1
        .Add "perm_write", 1
        .Add "perm_create", 1
        .Add "perm_unlink", 1
    End With
    oClient.ModelOfIrModelAccess.Method("create").ExecuteKw params
    ' ir.model.fields - create
    Set params = NewList  ' [{'model_id':id, 'name':'x_name', 'ttype':'char', 'state':'manual', 'required': True,}]
    With AddDict(params)
        .Add "model_id", nId
        .Add "name", "x_name2"
        .Add "ttype", "char"
        .Add "state", "manual"
        .Add "required", True
    End With
    Set oRet = oClient.ModelOfIrModelFields.Method("create").ExecuteKw(params)
    Debug.Print JsonConverter.ConvertToJson(oRet.Result, 4)
    Debug.Print
    ' x_custom - create
    Set params = NewList ' [{'x_name': "test record"}]
    With AddDict(params)
        .Add "x_name2", "test record"
    End With
    Set oRet = oClient.Model("x_custom").Method("create").ExecuteKw(params)
    nId = oRet.Result
    Debug.Print "id:", nId
    Debug.Print
    ' x_custom -read
    Set params = NewList ' [ids]
    params.Add nId
    Set oRet = oClient.Model("x_custom").Method("read").ExecuteKw(params)
    Debug.Print JsonConverter.ConvertToJson(oRet.Result, 4)
    Debug.Print
    
End Sub

