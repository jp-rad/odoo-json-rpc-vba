Attribute VB_Name = "ExampleJsonLookup"
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

Private Const CESCAPED_SLASH As String = vbBack

' Helper function for JSONLOOKUP to extract a value from a JSON object using a path
Private Function ExtractJsonValue(jsonObject As Object, jsonPath As String, ByRef rResult As Variant) As Boolean
    Dim remainingPath As String
    Dim nestedJsonObject As Object
    Dim pathSegments() As String
    Dim Key As String
    Dim Index As Integer
    Dim hasIndex As Boolean
    Dim tempItem As Variant
    Dim dict As Dictionary
    Dim coll As Collection

    ' Split the JSON path by '/' to extract the first key or index
    pathSegments = Split(jsonPath, "/", 2)
    If UBound(pathSegments) >= 0 Then
        Key = Trim(pathSegments(0))
    Else
        Key = ""
    End If
    If UBound(pathSegments) > 0 Then
        remainingPath = pathSegments(1)
    Else
        remainingPath = ""
    End If

    ' Handle key with index notation, e.g., item[0]
    If InStr(Key, "[") > 0 Then
        If Right(Key, 1) <> "]" Then
            Err.Raise 9 ' Index out of bounds, returns CVErr(xlErrRef)
        End If
        Key = Replace(Key, "[", "/")
        Key = Replace(Key, "]", "")
        pathSegments = Split(Key, "/")
        Key = Trim(pathSegments(0))
        If Not IsNumeric(pathSegments(1)) Then
            Err.Raise 9 ' Index out of bounds, returns CVErr(xlErrRef)
        End If
        Index = pathSegments(1) + 1 ' Collections in VBA are 1-based
        hasIndex = True
    Else
        hasIndex = False
    End If

    ' Retrieve the value from the dictionary
    If Key = "" Then
        Set tempItem = jsonObject
    Else
        Key = Replace(Key, CESCAPED_SLASH, "/") ' Unescape slash
        Set dict = jsonObject
        If Not dict.Exists(Key) Then
            Err.Raise 9 ' Index out of bounds, returns CVErr(xlErrRef)
        End If
        If IsObject(dict.Item(Key)) Then
            Set tempItem = dict.Item(Key)
        Else
            tempItem = dict.Item(Key)
        End If
    End If
    If IsNull(tempItem) Then
        GoTo ExitProc
    End If

    ' Retrieve value from collection if indexed
    If hasIndex Then
        If TypeName(tempItem) <> "Collection" Then
            Err.Raise 9 ' Index out of bounds, returns CVErr(xlErrRef)
        End If
        Set coll = tempItem
        If IsObject(coll.Item(Index)) Then
            Set tempItem = coll.Item(Index)
        Else
            tempItem = coll.Item(Index)
        End If
    End If
    If IsNull(tempItem) Then
        GoTo ExitProc
    End If

    ' Recursively resolve the remaining path
    If remainingPath <> "" Then
        Set nestedJsonObject = tempItem
        ExtractJsonValue nestedJsonObject, remainingPath, tempItem
    End If

ExitProc:
    ' Assign result
    If IsObject(tempItem) Then
        Set rResult = tempItem
    Else
        rResult = tempItem
    End If

End Function

' Excel worksheet function to extract values from JSON data
' jsonInput: JSON string to parse
' jsonPath: Path to the target value in the JSON structure, using '/' as a separator
'           Array elements can be accessed with square brackets, e.g., "items[0]" for the first element
Public Function JSONLOOKUP(ByVal jsonInput As String, ByVal jsonPath As String) As Variant
On Error GoTo ErrHandler
    Dim jsonObject As Object

    ' Parse JSON string into an object
    Set jsonObject = JsonConverter.ParseJson(jsonInput)

    ' Escape "\/"
    jsonPath = Replace(jsonPath, "\/", CESCAPED_SLASH)

    ' Normalize jsonPath
    jsonPath = Replace(jsonPath, "][", "]/[")

    ' Extract value using the specified JSON path
    ExtractJsonValue jsonObject, jsonPath, JSONLOOKUP
    If IsNull(JSONLOOKUP) Then
        JSONLOOKUP = CVErr(xlErrNA)
        GoTo ExitProc
    End If
    If IsObject(JSONLOOKUP) Then
        JSONLOOKUP = JsonConverter.ConvertToJson(JSONLOOKUP)
        GoTo ExitProc
    End If

ExitProc:
    Exit Function
ErrHandler:
    Select Case Err.Number
        Case 9 ' Index out of bounds, Dictionary or Collection
            JSONLOOKUP = CVErr(xlErrRef)
        Case Else ' General fallback
            JSONLOOKUP = CVErr(xlErrValue)
    End Select
    Resume ExitProc
End Function

' Excel worksheet function to count the number of items in a JSON array
' jsonArray: JSON array string to parse
Public Function JSONCOUNT(ByVal jsonArray As String) As Variant
On Error GoTo ErrHandler
    Dim coll As Collection

    Set coll = JsonConverter.ParseJson(jsonArray)
    JSONCOUNT = coll.Count

ExitProc:
    Exit Function
ErrHandler:
    JSONCOUNT = CVErr(xlErrValue)
    Resume ExitProc
End Function
