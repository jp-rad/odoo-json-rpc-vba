Attribute VB_Name = "Tests_OdFilter"
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
' Search domains - ORM API
'
' A search domain is a first-order logical predicate used for filtering and searching recordsets.
' You combine simple conditions on a field expression with logical operators.
'
' https://www.odoo.com/documentation/master/developer/reference/backend/orm.html#search-domains
'

Public Sub RunTests(Suite As TestSuite)
'    Dim Tests As New TestSuite
'    Dim Test As TestCase
'
'    With Suite.Test("should pass if all assertions pass")
'        Set Test = Tests.Test("should pass")
'
'        Test.IsOk True
'
'        .IsEqual Test.Result, TestResultType.Pass
'    End With
    
    TestJson Suite
    TestCombi Suite
    TestField Suite
    TestCriteria Suite
    TestDomain Suite
    TestCoding Suite
    
End Sub

Private Function IsValidJson(aJson As String) As Boolean
On Error GoTo ErrHandler
    JsonConverter.ParseJson aJson
    IsValidJson = True
ExitProc:
    Exit Function
ErrHandler:
    Debug.Print "Invalid JSON:", aJson
    Debug.Print "----- Err.Description -----"
    Debug.Print Err.Description
    Debug.Print "---------------------------"
    Resume ExitProc
End Function

Private Sub TestJson(Suite As TestSuite)
    Dim Tests As New TestSuite
    Dim Test As TestCase
    Dim v As Variant

    With Suite.Test("TestJson - True/False/Nothing To true/false/null")
        Set Test = Tests.Test("should pass")
        With NewField("is_company")
            With .Eq(v) ' Empty
                Test.IsEqual "['is_company', '=', null]", .ToJson()
                Test.IsOk IsValidJson(.ToJson())
            End With
            With .Eq(True)
                Test.IsEqual "['is_company', '=', true]", .ToJson()
                Test.IsOk IsValidJson(.ToJson())
            End With
            With .Eq(False)
                Test.IsEqual "['is_company', '=', false]", .ToJson()
                Test.IsOk IsValidJson(.ToJson())
            End With
            With .Eq(Nothing)
                Test.IsEqual "['is_company', '=', null]", .ToJson()
                Test.IsOk IsValidJson(.ToJson())
            End With
        End With
        .IsEqual Test.Result, TestResultType.Pass
    End With
    
End Sub

Public Sub TestCombi(Suite As TestSuite)
    Dim Tests As New TestSuite
    Dim Test As TestCase

    With Suite.Test("TestCombi - 2, 3")
        Set Test = Tests.Test("should pass")

        With NewAnd(NewField("phone").IsILike("7620"), NewField("mobile").IsILike("7620"))
            Test.IsEqual "'&', ['phone', 'ilike', '7620'], ['mobile', 'ilike', '7620']", .ToJson()
            .Add NewField("fax").IsILike("7620")
            Test.IsEqual "'&', '&', ['phone', 'ilike', '7620'], ['mobile', 'ilike', '7620'], ['fax', 'ilike', '7620']", .ToJson()
        End With
        
        With NewOr(NewField("phone").IsILike("7620"), NewField("mobile").IsILike("7620"))
            Test.IsEqual "'|', ['phone', 'ilike', '7620'], ['mobile', 'ilike', '7620']", .ToJson()
            .Add NewField("fax").IsILike("7620")
            Test.IsEqual "'|', '|', ['phone', 'ilike', '7620'], ['mobile', 'ilike', '7620'], ['fax', 'ilike', '7620']", .ToJson()
        End With
        
        With NewNot(NewField("phone").IsILike("7620"))
            Test.IsEqual "'!', ['phone', 'ilike', '7620']", .ToJson()
            .Add NewField("mobile").IsILike("7620")
            Test.IsEqual "'!', ['phone', 'ilike', '7620']", .ToJson()
            .Add NewField("fax").IsILike("7620")
            Test.IsEqual "'!', ['phone', 'ilike', '7620']", .ToJson()
        End With
        
        .IsEqual Test.Result, TestResultType.Pass
    End With
        
End Sub

Public Sub TestField(Suite As TestSuite)
    '
    ' Example
    '
    ' To search for partners named ABC, with a phone or mobile number containing 7620:
    '
    '   [('name', '=', 'ABC'),
    '    '|', ('phone','ilike','7620'), ('mobile', 'ilike', '7620')]
    '
    ' To search sales orders to invoice that have at least one line with a product that is out of stock:
    '
    '   [('invoice_status', '=', 'to invoice'),
    '    ('order_line', 'any', [('product_id.qty_available', '<=', 0)])]
    '
    ' To search for all partners born in the month of February:
    '
    '   [('birthday.month_number', '=', 2)]

    Dim Tests As New TestSuite
    Dim Test As TestCase

    With Suite.Test("should pass if all assertions pass")
        Set Test = Tests.Test("should pass")

        With NewField("name").Eq("ABC")
            Test.IsEqual "['name', '=', 'ABC']", .ToJson()
        End With
        
        With NewField("phone").IsILike("7620")
            Test.IsEqual "['phone', 'ilike', '7620']", .ToJson()
        End With
        
        With NewField("mobile").IsILike("7620")
            Test.IsEqual "['mobile', 'ilike', '7620']", .ToJson()
        End With
        
        With NewField("invoice_status").Eq("to invoice")
            Test.IsEqual "['invoice_status', '=', 'to invoice']", .ToJson()
        End With
        
        With NewField("invoice_status").Eq("to invoice")
            Test.IsEqual "['invoice_status', '=', 'to invoice']", .ToJson()
        End With
        
        With NewField("product_id.qty_available").Le(0)
            Test.IsEqual "['product_id.qty_available', '<=', 0]", .ToJson()
        End With
        
        With NewField("order_line").IsAny(NewDomain().AddArity(NewField("product_id.qty_available").Le(0)))
            Test.IsEqual "['order_line', 'any', [['product_id.qty_available', '<=', 0]]]", .ToJson()
        End With
    
        With NewField("birthday.month_number").Eq(2)
            Test.IsEqual "['birthday.month_number', '=', 2]", .ToJson()
        End With

        .IsEqual Test.Result, TestResultType.Pass
    End With
End Sub

Public Sub TestCriteria(Suite As TestSuite)
    '
    ' Example
    '
    ' To search for partners named ABC, with a phone or mobile number containing 7620:
    '
    '   [('name', '=', 'ABC'),
    '    '|', ('phone','ilike','7620'), ('mobile', 'ilike', '7620')]
    '
    ' To search sales orders to invoice that have at least one line with a product that is out of stock:
    '
    '   [('invoice_status', '=', 'to invoice'),
    '    ('order_line', 'any', [('product_id.qty_available', '<=', 0)])]
    '
    ' To search for all partners born in the month of February:
    '
    '   [('birthday.month_number', '=', 2)]

    Dim Tests As New TestSuite
    Dim Test As TestCase

    With Suite.Test("TestCriteria")
        Set Test = Tests.Test("should pass")
        
        With NewCriteria("name", "=", "ABC")
            Test.IsEqual "['name', '=', 'ABC']", .ToJson()
        End With
        
        With NewCriteria("phone", "ilike", "7620")
            Test.IsEqual "['phone', 'ilike', '7620']", .ToJson()
        End With
        
        With NewCriteria("mobile", "ilike", "7620")
            Test.IsEqual "['mobile', 'ilike', '7620']", .ToJson()
        End With
        
        With NewCriteria("invoice_status", "=", "to invoice")
            Test.IsEqual "['invoice_status', '=', 'to invoice']", .ToJson()
        End With
        
        With NewCriteria("invoice_status", "=", "to invoice")
            Test.IsEqual "['invoice_status', '=', 'to invoice']", .ToJson()
        End With
        
        With NewCriteria("product_id.qty_available", "<=", 0)
            Test.IsEqual "['product_id.qty_available', '<=', 0]", .ToJson()
        End With
        
        With NewCriteria("order_line", "any", NewDomain().AddArity(NewCriteria("product_id.qty_available", "<=", 0)))
            Test.IsEqual "['order_line', 'any', [['product_id.qty_available', '<=', 0]]]", .ToJson()
        End With
    
        With NewCriteria("birthday.month_number", "=", 2)
            Test.IsEqual "['birthday.month_number', '=', 2]", .ToJson()
        End With
        
        .IsEqual Test.Result, TestResultType.Pass
    End With

End Sub

Public Sub TestDomain(Suite As TestSuite)
    '
    ' Example
    '
    ' To search for partners named ABC, with a phone or mobile number containing 7620:
    '
    '   [('name', '=', 'ABC'),
    '    '|', ('phone','ilike','7620'), ('mobile', 'ilike', '7620')]
    '
    ' To search sales orders to invoice that have at least one line with a product that is out of stock:
    '
    '   [('invoice_status', '=', 'to invoice'),
    '    ('order_line', 'any', [('product_id.qty_available', '<=', 0)])]
    '
    ' To search for all partners born in the month of February:
    '
    '   [('birthday.month_number', '=', 2)]
    
    Dim Tests As New TestSuite
    Dim Test As TestCase

    With Suite.Test("TestDomain")
        Set Test = Tests.Test("should pass")
        
        With NewDomain()
            .AddArity NewField("name").Eq("ABC")
            .AddArity NewOr(NewField("phone").IsILike("7620"), NewField("mobile").IsILike("7620"))
            Test.IsOk IsValidJson(.ToJson())
            Test.IsEqual Replace("[['name','=','ABC'],'|',['phone','ilike','7620'],['mobile','ilike','7620']]", "'", """"), JsonConverter.ConvertToJson(.Build)
        End With
        
        With NewDomain()
            .AddArity NewField("invoice_status").Eq("to invoice")
            .AddArity NewField("order_line").IsAny(NewDomain().AddArity(NewField("product_id.qty_available").Le(0)))
            Test.IsOk IsValidJson(.ToJson())
            Test.IsEqual Replace("[['invoice_status','=','to invoice'],['order_line','any',[['product_id.qty_available','<=',0]]]]", "'", """"), JsonConverter.ConvertToJson(.Build)
        End With
        
        With NewDomain()
            .AddArity NewCriteria("birthday.month_number", "=", 2)
            Test.IsOk IsValidJson(.ToJson())
            Test.IsEqual Replace("[['birthday.month_number','=',2]]", "'", """"), JsonConverter.ConvertToJson(.Build)
        End With
        
        .IsEqual Test.Result, TestResultType.Pass
    End With

End Sub

Public Sub TestCoding(Suite As TestSuite)
    Dim params As Collection
    Dim domain As OdFilterDomain
    Dim subdomain As OdFilterDomain
    Dim criteria As OdFilterCriteria
    Dim nId As Long
    Dim v As Variant
    
    Dim Tests As New TestSuite
    Dim Test As TestCase

    With Suite.Test("TestCoding")
        Set Test = Tests.Test("should pass")

        ' [[['is_company', '=', True]]]
        Set params = NewList
        With NewDomain
            .AddArity NewField("is_company").Eq(True)
            .BuildAndAppend params
        End With
        Test.IsEqual Replace("[[['is_company','=',true]]]", "'", """"), JsonConverter.ConvertToJson(params)
        
        ' [[['id', '=', id]]]
        nId = &H7FFFFFFF ' 2147483647
        Set params = NewList
        With NewDomain()
            .AddArity NewField("id").Eq(nId)
            .BuildAndAppend params
        End With
        Test.IsEqual Replace("[[['id','=',2147483647]]]", "'", """"), JsonConverter.ConvertToJson(params)
    
        ' [[['id', '=', id]]]
        ' id = 0, 1, 2, 3
        ' build
        Set domain = NewDomain
        Set criteria = NewField("id").Eq(Empty)
        domain.AddArity criteria
        ' loop
        For Each v In Split("0, 1, 2, 3", ",")
            nId = v
            criteria.SetValue nId
            Set params = NewList
            params.Add domain.Build
            Test.IsEqual Replace("[[['id','='," & nId & "]]]", "'", """"), JsonConverter.ConvertToJson(params)
        Next v
        
        '   [[['invoice_status', '=', 'to invoice',
        '    ['order_line', 'any', [['product_id.qty_available', '<=', 0]]]]]
        '
        Set params = NewList
        Set subdomain = NewDomain()
        subdomain.AddArity NewField("product_id.qty_available").Le(0)
        With NewDomain()
            .AddArity NewField("invoice_status").Eq("to invoice")
            .AddArity NewField("order_line").IsAny(subdomain)
            .BuildAndAppend params
        End With
        Test.IsEqual Replace("[[['invoice_status','=','to invoice'],['order_line','any',[['product_id.qty_available','<=',0]]]]]", "'", """"), JsonConverter.ConvertToJson(params)
        
        .IsEqual Test.Result, TestResultType.Pass
    End With

End Sub
