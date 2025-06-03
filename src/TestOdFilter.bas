Attribute VB_Name = "TestOdFilter"
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

Public Sub TestJson()
    With NewField("is_company")
        With .Eq(True)
            Debug.Assert "['is_company', '=', true]" = .ToJson()
            Debug.Assert IsValidJson(.ToJson())
            JsonConverter.ParseJson .ToJson
        End With
        With .Eq(False)
            Debug.Assert "['is_company', '=', false]" = .ToJson()
            Debug.Assert IsValidJson(.ToJson())
        End With
        With .Eq(Nothing)
            Debug.Assert "['is_company', '=', null]" = .ToJson()
            Debug.Assert IsValidJson(.ToJson())
        End With
    End With
End Sub

Public Sub TestCombi()
    
    With NewAnd(NewField("phone").IsILike("7620"), NewField("mobile").IsILike("7620"))
        Debug.Assert "'&', ['phone', 'ilike', '7620'], ['mobile', 'ilike', '7620']" = .ToJson()
        .Add NewField("fax").IsILike("7620")
        Debug.Assert "'&', '&', ['phone', 'ilike', '7620'], ['mobile', 'ilike', '7620'], ['fax', 'ilike', '7620']" = .ToJson()
    End With
    
    With NewOr(NewField("phone").IsILike("7620"), NewField("mobile").IsILike("7620"))
        Debug.Assert "'|', ['phone', 'ilike', '7620'], ['mobile', 'ilike', '7620']" = .ToJson()
        .Add NewField("fax").IsILike("7620")
        Debug.Assert "'|', '|', ['phone', 'ilike', '7620'], ['mobile', 'ilike', '7620'], ['fax', 'ilike', '7620']" = .ToJson()
    End With
    
    With NewNot(NewField("phone").IsILike("7620"))
        Debug.Assert "'!', ['phone', 'ilike', '7620']" = .ToJson()
        .Add NewField("mobile").IsILike("7620")
        Debug.Assert "'!', ['phone', 'ilike', '7620']" = .ToJson()
        .Add NewField("fax").IsILike("7620")
        Debug.Assert "'!', ['phone', 'ilike', '7620']" = .ToJson()
    End With
    
    
End Sub

Public Sub TestField()
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

    With NewField("name").Eq("ABC")
        Debug.Print .ToJson()
        Debug.Assert "['name', '=', 'ABC']" = .ToJson()
    End With
    
    With NewField("phone").IsILike("7620")
        Debug.Print .ToJson()
        Debug.Assert "['phone', 'ilike', '7620']" = .ToJson()
    End With
    
    With NewField("mobile").IsILike("7620")
        Debug.Print .ToJson()
        Debug.Assert "['mobile', 'ilike', '7620']" = .ToJson()
    End With
    
    With NewField("invoice_status").Eq("to invoice")
        Debug.Print .ToJson()
        Debug.Assert "['invoice_status', '=', 'to invoice']" = .ToJson()
    End With
    
    With NewField("invoice_status").Eq("to invoice")
        Debug.Print .ToJson()
        Debug.Assert "['invoice_status', '=', 'to invoice']" = .ToJson()
    End With
    
    With NewField("product_id.qty_available").Le(0)
        Debug.Print .ToJson()
        Debug.Assert "['product_id.qty_available', '<=', 0]" = .ToJson()
    End With
    
    With NewField("order_line").IsAny(NewDomain().AddArity(NewField("product_id.qty_available").Le(0)))
        Debug.Print .ToJson()
        Debug.Assert "['order_line', 'any', [['product_id.qty_available', '<=', 0]]]" = .ToJson()
    End With

    With NewField("birthday.month_number").Eq(2)
        Debug.Print .ToJson()
        Debug.Assert "['birthday.month_number', '=', 2]" = .ToJson()
    End With

End Sub

Public Sub TestCriteria()
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

    With NewCriteria("name", "=", "ABC")
        Debug.Print .ToJson()
        Debug.Assert "['name', '=', 'ABC']" = .ToJson()
    End With
    
    With NewCriteria("phone", "ilike", "7620")
        Debug.Print .ToJson()
        Debug.Assert "['phone', 'ilike', '7620']" = .ToJson()
    End With
    
    With NewCriteria("mobile", "ilike", "7620")
        Debug.Print .ToJson()
        Debug.Assert "['mobile', 'ilike', '7620']" = .ToJson()
    End With
    
    With NewCriteria("invoice_status", "=", "to invoice")
        Debug.Print .ToJson()
        Debug.Assert "['invoice_status', '=', 'to invoice']" = .ToJson()
    End With
    
    With NewCriteria("invoice_status", "=", "to invoice")
        Debug.Print .ToJson()
        Debug.Assert "['invoice_status', '=', 'to invoice']" = .ToJson()
    End With
    
    With NewCriteria("product_id.qty_available", "<=", 0)
        Debug.Print .ToJson()
        Debug.Assert "['product_id.qty_available', '<=', 0]" = .ToJson()
    End With
    
    With NewCriteria("order_line", "any", NewDomain().AddArity(NewCriteria("product_id.qty_available", "<=", 0)))
        Debug.Print .ToJson()
        Debug.Assert "['order_line', 'any', [['product_id.qty_available', '<=', 0]]]" = .ToJson()
    End With

    With NewCriteria("birthday.month_number", "=", 2)
        Debug.Print .ToJson()
        Debug.Assert "['birthday.month_number', '=', 2]" = .ToJson()
    End With

End Sub

Public Sub TestDomain()
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
    
    With NewDomain()
        .AddArity NewField("name").Eq("ABC")
        .AddArity NewOr(NewField("phone").IsILike("7620"), NewField("mobile").IsILike("7620"))
        Debug.Print .ToJson()
        Debug.Assert IsValidJson(.ToJson())
        Debug.Assert "[['name','=','ABC'],'|',['phone','ilike','7620'],['mobile','ilike','7620']]" = Replace(JsonConverter.ConvertToJson(.Build), """", "'")
    End With
    
    With NewDomain()
        .AddArity NewField("invoice_status").Eq("to invoice")
        .AddArity NewField("order_line").IsAny(NewDomain().AddArity(NewField("product_id.qty_available").Le(0)))
        Debug.Print .ToJson()
        Debug.Assert IsValidJson(.ToJson())
        Debug.Assert "[['invoice_status','=','to invoice'],['order_line','any',[['product_id.qty_available','<=',0]]]]" = Replace(JsonConverter.ConvertToJson(.Build), """", "'")
    End With
    
    With NewDomain()
        .AddArity NewCriteria("birthday.month_number", "=", 2)
        Debug.Print .ToJson()
        Debug.Assert IsValidJson(.ToJson())
        Debug.Assert "[['birthday.month_number','=',2]]" = Replace(JsonConverter.ConvertToJson(.Build), """", "'")
    End With

End Sub

Public Sub TestCoding()
    Dim params As Collection
    Dim domain As OdFilterDomain
    Dim criteria As OdFilterCriteria
    Dim nId As Long
    Dim v As Variant
    
    ' [[['is_company', '=', True]]]
    Set params = NewList
    With NewDomain
        .AddArity NewField("is_company").Eq(True)
        .BuildAndAppend params
    End With
    Debug.Print JsonConverter.ConvertToJson(params)
    Debug.Assert "[[['is_company','=',true]]]" = Replace(JsonConverter.ConvertToJson(params), """", "'")
    
    ' [[['id', '=', id]]]
    nId = &H7FFFFFFF ' 2147483647
    Set params = NewList
    With NewDomain()
        .AddArity NewField("id").Eq(nId)
        .BuildAndAppend params
    End With
    Debug.Print JsonConverter.ConvertToJson(params)
    Debug.Assert "[[['id','=',2147483647]]]" = Replace(JsonConverter.ConvertToJson(params), """", "'")

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
        Debug.Print JsonConverter.ConvertToJson(params)
        Debug.Assert "[[['id','='," & nId & "]]]" = Replace(JsonConverter.ConvertToJson(params), """", "'")
    Next v
End Sub

Public Sub TestAll()
    TestJson
    TestCombi
    TestField
    TestCriteria
    TestDomain
    TestCoding
End Sub
