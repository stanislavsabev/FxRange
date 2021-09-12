VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FxRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''' Class to manage Table-like Range
''' Version 0.1.0


Option Explicit
Private Const ModuleName = "FxRange"

Private Type TIndexItem
    Name        As Variant
    Ndx         As Long
    Count       As Long
    MultiNdx    As Collection
End Type

Public Enum FxRangeErrors
    [_First] = vbObjectError
    InitError = [_First] + 100
End Enum

Private Type TThis
    IsInitiated     As Boolean
    Rng             As Range
    ProcName        As String
End Type
Dim this As TThis

' , Optional KeyColumn, Optional RowIndex, Optional ColumnIndex = False, Optional AllowDuplicates = False, Optional AllowGaps
Public Sub Init(Rng As Range)
    this.ProcName = "Init"
    If this.IsInitiated Then
        Raise_ InitError, "Already initiated"
    End If
    
    If Rng Is Nothing Then Err.Raise 91, "Init"

    this.IsInitiated = True
End Sub

Private Sub Raise_(ErrCode As FxRangeErrors, Optional ErrDescr As String)
    Err.Raise ErrCode, this.ProcName, (IIf(ErrDescr <> "", ErrDescr, Error(ErrCode)))
End Sub
