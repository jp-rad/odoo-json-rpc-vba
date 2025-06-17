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
    
    Test_JSONCOUNT Suite
    
End Sub

Private Sub Test_Returns(Suite As TestSuite)
    Dim Tests As New TestSuite
    Dim Test As TestCase
    Dim json As String
    
    With Suite.Test("JSONLOOKUP - ReturnsValueForEmptyPath")
    
        Set Test = Tests.Test("ReturnsValueForEmptyPath")
        With Test
            json = "{""foo"":123,""bar"":{""baz"":""hello""}}"
            .IsEqual "{""foo"":123,""bar"":{""baz"":""hello""}}", JSONLOOKUP(json, "")
        End With
        
        .IsEqual Test.Result, TestResultType.Pass
    End With
    
    With Suite.Test("JSONLOOKUP - ReturnsValueForSimplePath")
    
        Set Test = Tests.Test("ReturnsValueForSimplePath")
        With Test
            json = "{""foo"":123,""bar"":{""baz"":""hello""}}"
            .IsEqual 123, JSONLOOKUP(json, "foo")
        End With
        
        .IsEqual Test.Result, TestResultType.Pass
    End With
    
    With Suite.Test("JSONLOOKUP - ReturnsValueForSplashKey")
    
        Set Test = Tests.Test("ReturnsValueForSplashKey")
        With Test
            json = "{""foo"":123,""foo/bar"":{""baz"":""hello""}}"
            .IsEqual "hello", JSONLOOKUP(json, "foo\/bar/baz")
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
    
    With Suite.Test("JSONLOOKUP - ReturnsNestedArrayValue")
        
        Set Test = Tests.Test("ReturnsNestedValue")
        With Test
            json = "[""apple0"", [""apple1"", [""apple2"", ""banana2"", ""cherry2""], ""cherry1""], ""cherry0""]"
            .IsEqual "banana2", JSONLOOKUP(json, "[1]/[1]/[1]")
        End With
        With Test
            json = "[""apple0"", [""apple1"", [""apple2"", ""banana2"", ""cherry2""], ""cherry1""], ""cherry0""]"
            .IsEqual "apple2", JSONLOOKUP(json, "[1][1][0]")
        End With
        
        .IsEqual Test.Result, TestResultType.Pass
    End With
    
    With Suite.Test("JSONLOOKUP - ReturnsNestedValueInArray")
        
        Set Test = Tests.Test("ReturnsNestedValueInArray")
        With Test
            json = "{""items"": [""apple"", {""foo"":123,""bar"":{""baz"":""hello""}}, ""cherry""]}"
            .IsEqual "hello", JSONLOOKUP(json, "items[1]/bar/baz")
        End With
        
        .IsEqual Test.Result, TestResultType.Pass
    End With
    
End Sub

Private Sub Test_ErrValue(Suite As TestSuite)
    Dim Tests As New TestSuite
    Dim Test As TestCase
    Dim json As String
    Dim Result As Variant
    
    With Suite.Test("JSONLOOKUP - ReturnsValueErrorForEmptyJson")
    
        Set Test = Tests.Test("ReturnsValueErrorForEmptyJson")
        With Test
            json = ""
            Result = JSONLOOKUP(json, "foo")
            .IsOk IsError(Result)
            .IsOk CVErr(xlErrValue) = Result
        End With
        
        .IsEqual Test.Result, TestResultType.Pass
    End With
    
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
        
        With Test
            json = "{""items"": [""apple"", ""banana""]}"
            Result = JSONLOOKUP(json, "items[string]")
            .IsOk IsError(Result)
            .IsOk CVErr(xlErrRef) = Result
        End With
        
        With Test
            json = "{""items"": [""apple"", ""banana""]}"
            Result = JSONLOOKUP(json, "items[1")
            .IsOk IsError(Result)
            .IsOk CVErr(xlErrRef) = Result
        End With
        
        With Test
            json = "{""items"": [""apple"", ""banana""]}"
            Result = JSONLOOKUP(json, "items[0]1")
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
    
    With Suite.Test("JSONLOOKUP - ReturnsNAForNullInArray")
    
        Set Test = Tests.Test("ReturnsNAForNullInArray")
        With Test
            json = "{""items"": [""apple"", null, ""cherry""]}"
            Result = JSONLOOKUP(json, "items[1]")
            .IsOk IsError(Result)
            .IsOk CVErr(xlErrNA) = Result
        End With
        
        With Test
            json = "{""items"": [""apple"", null, ""cherry""]}"
            Result = JSONLOOKUP(json, "items[1]/foo/bar")
            .IsOk IsError(Result)
            .IsOk CVErr(xlErrNA) = Result
        End With
        
        With Test
            json = "{""items"": [""apple"", null, ""cherry""]}"
            Result = JSONLOOKUP(json, "items[1]/[0]")
            .IsOk IsError(Result)
            .IsOk CVErr(xlErrNA) = Result
        End With
                
        .IsEqual Test.Result, TestResultType.Pass
    End With
    
End Sub

Private Sub Test_JSONCOUNT(Suite As TestSuite)
    Dim Tests As New TestSuite
    Dim Test As TestCase
    Dim json As String
    
    With Suite.Test("JSONCOUNT - Count0")
    
        Set Test = Tests.Test("Count0")
        With Test
            json = "[]"
            .IsEqual 0, JSONCOUNT(json)
        End With
        
        .IsEqual Test.Result, TestResultType.Pass
    End With
    
    With Suite.Test("JSONCOUNT - Count1")
    
        Set Test = Tests.Test("Count1")
        With Test
            json = "[""apple""]"
            .IsEqual 1, JSONCOUNT(json)
        End With
        
        .IsEqual Test.Result, TestResultType.Pass
    End With
    
    With Suite.Test("JSONCOUNT - Count2")
    
        Set Test = Tests.Test("Count2")
        With Test
            json = "[""apple"", null]"
            .IsEqual 2, JSONCOUNT(json)
        End With
        
        .IsEqual Test.Result, TestResultType.Pass
    End With
    
    With Suite.Test("JSONCOUNT - Count3")
    
        Set Test = Tests.Test("Count3")
        With Test
            json = "[""apple"", null, ""cherry""]"
            .IsEqual 3, JSONCOUNT(json)
        End With
        
        .IsEqual Test.Result, TestResultType.Pass
    End With
    
    With Suite.Test("JSONCOUNT - Err")
    
        Set Test = Tests.Test("Err")
        With Test
            json = "{""foo"":123,""bar"":{""baz"":""hello""}}"
            .IsOk CVErr(xlErrValue) = JSONCOUNT(json)
        End With
        
        .IsEqual Test.Result, TestResultType.Pass
    End With
    
End Sub
