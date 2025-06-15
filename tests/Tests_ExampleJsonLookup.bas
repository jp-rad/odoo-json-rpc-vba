Attribute VB_Name = "Tests_ExampleJsonLookup"
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
    
    Test_Returns Suite
    Test_ErrValue Suite
    Test_ErrRef Suite
    Test_ErrNa Suite
    
End Sub

Private Sub Test_Returns(Suite As TestSuite)
    Dim Tests As New TestSuite
    Dim Test As TestCase
    Dim json As String
    With Suite.Test("JSONLOOKUP - ReturnsValueForSimplePath")
    
        Set Test = Tests.Test("ReturnsValueForSimplePath")
        With Test
            json = "{""foo"":123,""bar"":{""baz"":""hello""}}"
            .IsEqual 123, JSONLOOKUP(json, "foo")
        End With
        
        .IsEqual Test.Result, TestResultType.Pass
    End With
    
    With Suite.Test("JSONLOOKUP - ReturnsNestedValue")
        
        Set Test = Tests.Test("ReturnsNestedValue")
        With Test
            json = "{""foo"":123,""bar"":{""baz"":""hello""}}"
            .IsEqual "hello", JSONLOOKUP(json, "bar/baz")
        End With
        
        .IsEqual Test.Result, TestResultType.Pass
    End With
    
    With Suite.Test("JSONLOOKUP - ReturnsNested")
        
        Set Test = Tests.Test("ReturnsNested")
        With Test
            json = "{""foo"":123,""bar"":{""baz"":""hello""}}"
            .IsEqual "{""baz"":""hello""}", JSONLOOKUP(json, "bar")
        End With
        
        .IsEqual Test.Result, TestResultType.Pass
    End With
    
    With Suite.Test("JSONLOOKUP - ReturnsArrayElementByIndex0")
        
        Set Test = Tests.Test("ReturnsArrayElementByIndex0")
        With Test
            json = "{""items"": [""apple"", ""banana"", ""cherry""]}"
            .IsEqual "apple", JSONLOOKUP(json, "items[0]")
        End With
        
        .IsEqual Test.Result, TestResultType.Pass
    End With
    
    With Suite.Test("JSONLOOKUP - ReturnsArrayElementByIndex2")
        
        Set Test = Tests.Test("ReturnsArrayElementByIndex2")
        With Test
            json = "{""items"": [""apple"", ""banana"", ""cherry""]}"
            .IsEqual "cherry", JSONLOOKUP(json, "items[2]")
        End With
        
        .IsEqual Test.Result, TestResultType.Pass
    End With
    
    With Suite.Test("JSONLOOKUP - ReturnsArray")
        
        Set Test = Tests.Test("ReturnsArray")
        With Test
            json = "{""items"": [""apple"", ""banana"", ""cherry""]}"
            .IsEqual "[""apple"",""banana"",""cherry""]", JSONLOOKUP(json, "items")
        End With
        
        .IsEqual Test.Result, TestResultType.Pass
    End With
    
    With Suite.Test("JSONLOOKUP - ReturnsSimpleArrayElement")
        
        Set Test = Tests.Test("ReturnsSimpleArrayElement")
        With Test
            json = "[""apple"", ""banana"", ""cherry""]"
            .IsEqual "banana", JSONLOOKUP(json, "[1]")
        End With
        
        .IsEqual Test.Result, TestResultType.Pass
    End With
    
End Sub

Private Sub Test_ErrValue(Suite As TestSuite)
    Dim Tests As New TestSuite
    Dim Test As TestCase
    Dim json As String
    Dim Result As Variant
    With Suite.Test("JSONLOOKUP - ReturnsValueErrorForInvalidJson")
    
        Set Test = Tests.Test("ReturnsValueErrorForInvalidJson")
        With Test
            json = "{foo:123"
            Result = JSONLOOKUP(json, "foo")
            .IsOk IsError(Result)
            .IsOk CVErr(xlErrValue) = Result
        End With
        
        .IsEqual Test.Result, TestResultType.Pass
    End With
End Sub

Private Sub Test_ErrRef(Suite As TestSuite)
    Dim Tests As New TestSuite
    Dim Test As TestCase
    Dim json As String
    Dim Result As Variant
    With Suite.Test("JSONLOOKUP - ReturnsRefErrorForInvalidPath")
    
        Set Test = Tests.Test("ReturnsRefErrorForInvalidPath")
        With Test
            json = "{""foo"":123}"
            Result = JSONLOOKUP(json, "bar")
            .IsOk IsError(Result)
            .IsOk CVErr(xlErrRef) = Result
        End With
        
        .IsEqual Test.Result, TestResultType.Pass
    End With
    With Suite.Test("JSONLOOKUP - ReturnsRefErrorForInvalidPathNotArray")
    
        Set Test = Tests.Test("ReturnsRefErrorForInvalidPath")
        With Test
            json = "{""foo"":123}"
            Result = JSONLOOKUP(json, "foo[1]")
            .IsOk IsError(Result)
            .IsOk CVErr(xlErrRef) = Result
        End With
        
        .IsEqual Test.Result, TestResultType.Pass
    End With
    
    With Suite.Test("JSONLOOKUP - ReturnsRefErrorForOutOfRangeIndex")
        
        Set Test = Tests.Test("ReturnsRefErrorForOutOfRangeIndex")
        With Test
            json = "{""items"": [""apple"", ""banana""]}"
            Result = JSONLOOKUP(json, "items[2]")
            .IsOk IsError(Result)
            .IsOk CVErr(xlErrRef) = Result
        End With
        
        With Test
            json = "{""items"": [""apple"", ""banana""]}"
            Result = JSONLOOKUP(json, "items[-1]")
            .IsOk IsError(Result)
            .IsOk CVErr(xlErrRef) = Result
        End With
        
        .IsEqual Test.Result, TestResultType.Pass
    End With
End Sub

Private Sub Test_ErrNa(Suite As TestSuite)
    Dim Tests As New TestSuite
    Dim Test As TestCase
    Dim json As String
    Dim Result As Variant
    With Suite.Test("JSONLOOKUP - ReturnsNAForNull")
        
        Set Test = Tests.Test("ReturnsNAForNull")
        With Test
            json = "{""foo"":null}"
            Result = JSONLOOKUP(json, "foo")
            .IsOk IsError(Result)
            .IsOk CVErr(xlErrNA) = Result
        End With
        
        .IsEqual Test.Result, TestResultType.Pass
    End With
    
    With Suite.Test("JSONLOOKUP - ReturnsNAForNullNested")
        
        Set Test = Tests.Test("ReturnsNAForNullNested")
        With Test
            json = "{""foo"":null}"
            Result = JSONLOOKUP(json, "foo/bar")
            .IsOk IsError(Result)
            .IsOk CVErr(xlErrNA) = Result
        End With
        
        With Test
            json = "{""foo"":null}"
            Result = JSONLOOKUP(json, "foo[1]")
            .IsOk IsError(Result)
            .IsOk CVErr(xlErrNA) = Result
        End With
        
        .IsEqual Test.Result, TestResultType.Pass
    End With
    
End Sub
