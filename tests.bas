Attribute VB_Name = "tests"

Function test_calling_Init_more_than_once_should_raise_InitError() As String
    Dim ErrNumber   As Long
    Dim Rng         As Range
    Dim sut         As New FxRange
    
    ' Setup
    Set Rng = Range("A1:C2") ' valid range
    
    ' Test
    On Error Resume Next
    sut.Init Rng
    sut.Init Rng
    ErrNumber = Err.Number
    On Error GoTo 0
   
    ' Verify
    If Not ErrNumber = FxRangeErrors.InitError Then
        test_calling_Init_more_than_once_should_raise_InitError = "Did not Raise InitError"
    End If
    
    ' TearDown
    '//
End Function

Function test_calling_Init_with_empty_Range_should_raise_ObjectNotSetError() As String
    Dim ErrNumber   As Long
    Dim ExpectedErrorNumber As Long
    Dim Rng         As Range
    Dim sut         As New FxRange
    
    ' Setup
     ExpectedErrorNumber = 91 ' Object variable or With block variable not set
     
    ' Test
    On Error Resume Next
    sut.Init Rng
    ErrNumber = Err.Number
    On Error GoTo 0
   
    ' Verify
    If Not ErrNumber = 91 Then
        test_calling_Init_with_empty_Range_should_raise_ObjectNotSetError = _
            fx.StrFormat("Did not raise Error {} {}", ExpectedErrorNumber, Error(ExpectedErrorNumber))
        
    End If
    
    ' TearDown
    '//
End Function

Function prototype() As String
    Dim Rng         As Range
    Dim sut         As New FxRange
    
    ' Setup
     
    ' Test
    On Error Resume Next

    On Error GoTo 0
   
    ' Verify
    '//
    
    ' TearDown
    '//
End Function
