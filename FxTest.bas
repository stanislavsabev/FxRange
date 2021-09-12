Attribute VB_Name = "FxTest"
Option Explicit

Private Type TTest
    Failed  As Boolean
    Name    As String
    Details As String
End Type

Private Type TThis
    Tests()     As TTest
    TestsCount  As Long
    Last        As TTest
End Type
Private this As TThis

Sub FxTest_RunAll(Optional Verbose = False)
    Dim i As Long
    Dim Current As TTest
    
    DiscoverTests
    If this.TestsCount = 0 Then
        Debug.Print "0 tests found"
        Exit Sub
    End If
    
    For i = 0 To this.TestsCount - 1
        RunSingle this.Tests(i)
    Next
    
    this.Last = this.Tests(this.TestsCount - 1)
    
    Debug.Print PrintResults(Verbose)
End Sub

Public Sub FxTest_RunByName(Name As String)
    Dim Result    As TTest
    ' each test should return string with details if failed, or empty string if succeeded
    Result.Details = Application.Run(Name)
    Result.Name = Name
    Result.Failed = (Result.Details <> "")
    this.Last = Result
    PrintLastResult
End Sub

Public Function PrintLastResult() As String
    PrintVerboseSingle this.Last
End Function

Private Sub RunSingle(Test As TTest)
    
    ' each test should return string with details if failed, or empty string if succeeded
    Test.Details = Application.Run(Test.Name)
    Test.Failed = (Test.Details <> "")
        
End Sub

Private Sub DiscoverTests()
    Dim Module      As TModule
    Dim Name        As String
    Dim i           As Long
    Dim Count       As Long
    Const Criteria = "test_*"
    
    Module = ReadModule("tests")
    If Module.ProceduresCount = 0 Then
        Exit Sub
    End If

    For i = 0 To Module.ProceduresCount - 1
        Name = Module.Procedures(i).Name
        If LCase(Name) Like Criteria Then
            ReDim Preserve this.Tests(Count)
            this.Tests(Count).Name = Name
            Count = Count + 1
        End If
    Next
    this.TestsCount = Count
End Sub

Public Function PrintResults(Optional Verbose = False)
    Dim i           As Long
    Dim FailedCount As Long
    
    For i = 0 To this.TestsCount - 1
        If this.Tests(i).Failed Then
            FailedCount = FailedCount + 1
        End If
    Next
        
    If Verbose Then
        PrintVerbose
    Else
        PrintSimple
    End If
    
    Debug.Print vbNewLine & "=========="
    Debug.Print "Ran: "; this.TestsCount; IIf(this.TestsCount = 1, " test. ", " tests. ");
    Debug.Print " OK: "; this.TestsCount - FailedCount, " Failed: "; FailedCount
End Function

Private Sub PrintVerbose()
    Dim i       As Long
    For i = 0 To this.TestsCount - 1
        PrintVerboseSingle this.Tests(i)
    Next

End Sub

Private Sub PrintVerboseSingle(Test As TTest)
    If Test.Failed Then
        Debug.Print "[Failed]: ";
        Debug.Print Test.Name
        Debug.Print ">> "; Test.Details
    Else
        Debug.Print "[OK]: ";
        Debug.Print Test.Name
    End If
End Sub

Private Sub PrintSimple()
    Dim i       As Long
    For i = 0 To this.TestsCount - 1
        Debug.Print IIf(this.Tests(i).Failed, " Fail ", ".");
    Next
End Sub



