Attribute VB_Name = "Tests"
Public Sub Run(Optional OutputPath As Variant)
    Dim Suite As New TestSuite
    Suite.Description = "vba-test"
    
    Dim Immediate As New ImmediateReporter
    Immediate.ListenTo Suite
    
    If Not IsMissing(OutputPath) And CStr(OutputPath) <> "" Then
        Dim Reporter As New FileReporter
        Reporter.WriteTo OutputPath
        Reporter.ListenTo Suite
    End If
    
    Tests_OdFilter.RunTests Suite.Group("Tests_OdFilter")
    
End Sub

