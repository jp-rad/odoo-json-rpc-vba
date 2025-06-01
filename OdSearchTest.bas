Attribute VB_Name = "OdSearchTest"
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

Public Sub doLoginLocalhost()
    Dim ret As OdResult

    Set mClient = NewOdWebClient()
    Debug.Print TypeName(mClient)
    ' Turn off SSL validation
    mClient.SetInsecure True
    ' Follow redirects (301, 302, 307) using Location header
    mClient.SetFollowRedirects False
    
    mClient.BaseUrl = CBASEURL
    mClient.DbName = CDBNAME
    mClient.Username = CUSERNAME
    mClient.Password = CPASSWORD
    
    Debug.Print "---------------"
    Debug.Print " Setup client"
    Debug.Print "---------------"
    Debug.Print "BaseUrl:", mClient.BaseUrl
    Debug.Print "Database:", mClient.DbName
    Debug.Print "Username:", mClient.Username
    Debug.Print "Password:", mClient.Password
    Debug.Print

    Set ret = mClient.Common.Authenticate()
    Debug.Print "Login", ret.JsonResult
    
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
        Debug.Print .ToFilterString()
        Debug.Assert "('name', '=', 'ABC')" = .ToFilterString()
    End With
    
    With NewField("phone").IsILike("7620")
        Debug.Print .ToFilterString()
        Debug.Assert "('phone', 'ilike', '7620')" = .ToFilterString()
    End With
    
    With NewField("mobile").IsILike("7620")
        Debug.Print .ToFilterString()
        Debug.Assert "('mobile', 'ilike', '7620')" = .ToFilterString()
    End With
    
    With NewField("invoice_status").Eq("to invoice")
        Debug.Print .ToFilterString()
        Debug.Assert "('invoice_status', '=', 'to invoice')" = .ToFilterString()
    End With
    
    With NewField("invoice_status").Eq("to invoice")
        Debug.Print .ToFilterString()
        Debug.Assert "('invoice_status', '=', 'to invoice')" = .ToFilterString()
    End With
    
    With NewField("product_id.qty_available").Le(0)
        Debug.Print .ToFilterString()
        Debug.Assert "('product_id.qty_available', '<=', 0)" = .ToFilterString()
    End With
    
    With NewField("order_line").IsAny(NewDomain().AddArity(NewField("product_id.qty_available").Le(0)))
        Debug.Print .ToFilterString()
        Debug.Assert "('order_line', 'any', [('product_id.qty_available', '<=', 0)])" = .ToFilterString()
    End With

    With NewField("birthday.month_number").Eq(2)
        Debug.Print .ToFilterString()
        Debug.Assert "('birthday.month_number', '=', 2)" = .ToFilterString()
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
        Debug.Print .ToFilterString()
        Debug.Assert "('name', '=', 'ABC')" = .ToFilterString()
    End With
    
    With NewCriteria("phone", "ilike", "7620")
        Debug.Print .ToFilterString()
        Debug.Assert "('phone', 'ilike', '7620')" = .ToFilterString()
    End With
    
    With NewCriteria("mobile", "ilike", "7620")
        Debug.Print .ToFilterString()
        Debug.Assert "('mobile', 'ilike', '7620')" = .ToFilterString()
    End With
    
    With NewCriteria("invoice_status", "=", "to invoice")
        Debug.Print .ToFilterString()
        Debug.Assert "('invoice_status', '=', 'to invoice')" = .ToFilterString()
    End With
    
    With NewCriteria("invoice_status", "=", "to invoice")
        Debug.Print .ToFilterString()
        Debug.Assert "('invoice_status', '=', 'to invoice')" = .ToFilterString()
    End With
    
    With NewCriteria("product_id.qty_available", "<=", 0)
        Debug.Print .ToFilterString()
        Debug.Assert "('product_id.qty_available', '<=', 0)" = .ToFilterString()
    End With
    
    With NewCriteria("order_line", "any", NewDomain().AddArity(NewCriteria("product_id.qty_available", "<=", 0)))
        Debug.Print .ToFilterString()
        Debug.Assert "('order_line', 'any', [('product_id.qty_available', '<=', 0)])" = .ToFilterString()
    End With

    With NewCriteria("birthday.month_number", "=", 2)
        Debug.Print .ToFilterString()
        Debug.Assert "('birthday.month_number', '=', 2)" = .ToFilterString()
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
        Debug.Print .ToFilterString()
        Debug.Assert "[('name', '=', 'ABC'), '|', ('phone', 'ilike', '7620'), ('mobile', 'ilike', '7620')]" = .ToFilterString()
    End With
    
    With NewDomain()
        .AddArity NewField("invoice_status").Eq("to invoice")
        .AddArity NewField("order_line").IsAny(NewDomain().AddArity(NewField("product_id.qty_available").Le(0)))
        Debug.Print .ToFilterString()
        Debug.Assert "[('invoice_status', '=', 'to invoice'), ('order_line', 'any', [('product_id.qty_available', '<=', 0)])]" = .ToFilterString()
    End With
    
    With NewDomain()
        .AddArity NewCriteria("birthday.month_number", "=", 2)
        Debug.Print .ToFilterString()
        Debug.Assert "[('birthday.month_number', '=', 2)]" = .ToFilterString()
    End With

End Sub

