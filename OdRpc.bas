Attribute VB_Name = "OdRpc"
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
' see also: https://www.odoo.com/documentation/15.0/developer/misc/api/odoo.html
'

Option Explicit

Public Const CURL_DEMO As String = "https://demo.odoo.com"

Public Const CERR_STATUSCODE   As Long = 2001 + vbObjectError   ' web response error
Public Const CERR_RESPONSE     As Long = 2002 + vbObjectError   ' JSON-RPC error
Public Const CERR_JSONRPC_ID   As Long = 2003 + vbObjectError   ' JSON-RPC id error
Public Const CERR_AUTHENTICATE As Long = 2004 + vbObjectError   ' authentication failed

Public Const CMETHOD_CHECK_ACCESS_RIGHTS As String = "check_access_rights"
Public Const CMETHOD_SEARCH              As String = "search"
Public Const CMETHOD_SEARCH_COUNT        As String = "search_count"
Public Const CMETHOD_READ                As String = "read"
Public Const CMETHOD_FIELDS_GET          As String = "fields_get"
Public Const CMETHOD_SEARCH_READ         As String = "search_read"
Public Const CMETHOD_CREATE              As String = "create"
Public Const CMETHOD_WRITE               As String = "write"
Public Const CMETHOD_NAME_GET            As String = "name_get"
Public Const CMETHOD_UNLINK              As String = "unlink"

Public Function FormatDate(aDate As Date) As String
    FormatDate = Format(aDate, "yyyy-mm-dd")
End Function

Public Function ParseDate(aDateString As String) As Date
    ParseDate = CDate(aDateString)
End Function

Public Function ConvertToIsoDatetime(aDatetime As Date) As String
    ConvertToIsoDatetime = JsonConverter.ConvertToIso(aDatetime)
End Function

Public Function ParseIsoDatetime(aIsoString As String) As Date
    ParseIsoDatetime = JsonConverter.ParseIso(aIsoString)
End Function
