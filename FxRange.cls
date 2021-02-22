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
''' Note: This version does not support

Option Explicit
Private Const ModuleName = "FxRange"

Private Type TIndexItem
    Name        As Variant
    Ndx         As Long
    Count       As Long
    MultiNdx    As Collection
End Type

Private Type TIndex
    Index       As Variant
    NameToNdx   As Scripting.Dictionary
    NdxToName   As Scripting.Dictionary
    Items()     As TIndexItem
    Size        As Long
    IsRange     As Boolean
End Type

Private Type TThis
    IsInitiated     As Boolean
    RowIndex        As TIndex
    ColumnIndex     As TIndex
    KeyColumn       As TIndexItem
    AllowEmpty      As Boolean
    AllowDuplicates As Boolean
    Rng             As Range
End Type
Dim this As TThis

'::region Constructor
Public Sub Init(ByVal Rng As Range, Optional KeyColumn, Optional RowIndex, Optional ColumnIndex, Optional AllowEmpty = False, Optional AllowDuplicates = False)
    If this.IsInitiated Then
        Debug.Print "[DEBUG] "; ModuleName; ".Init()"; ": Already initiated!"
        Exit Sub
    End If
    
    ValidateInput_ Rng, KeyColumn, RowIndex, ColumnIndex
    
    Set this.Rng = Rng
    this.AllowEmpty = AllowEmpty
    this.AllowDuplicates = AllowDuplicates
    CreateIndex_ Rng, KeyColumn, RowIndex, ColumnIndex

    this.IsInitiated = True
End Sub

'::region Public
Public Function InsertRows(NumRows As Long, Optional After, Optional EntireSheetRow = False) As FxRange
    Dim Result      As FxRange
    Dim AfterRow    As Range
    Dim NewRows     As Range

    If IsMissing(After) Then
        Set AfterRow = Me.LastRow
    Else
        Set AfterRow = Me.iLoc(After)
    End If

    Set NewRows = AfterRow.Offset(1, 0).Resize(RowSize:=NumRows)
    
    If EntireSheetRow Then
        NewRows.EntireSheetRow.Insert Shift:=xlShiftDown
    Else
        NewRows.Insert Shift:=xlShiftDown
    End If
    
    Remap FullRange.Resize(RowSize:=Me.AllRowsCount + NumRows)

    Set Result = New FxRange
    Result.Init AfterRow.Offset(1, 0).Resize(RowSize:=NumRows), _
        ColumnIndex:=this.ColumnIndex.Index, KeyColumn:=this.KeyColumn.Name
    Set InsertRows = Result
End Function

Public Function DeleteRows(NumRows As Long, Optional At, Optional EntireRow = False) As FxRange

    fx.Raise_NotImplementedError
End Function


Public Function InsertColumns(Columns As Variant, Optional Before, Optional EntireColumn = False) As FxRange

    fx.Raise_NotImplementedError
End Function

Public Function DeleteColumns(Columns As Variant, Optional EntireColumn = False) As FxRange

    fx.Raise_NotImplementedError
End Function

Public Sub CalculateWhenManual()
    If Application.Calculation = xlCalculationManual Then
        Me.Parent.Calculate
    End If
End Sub

Public Sub Remap(Optional Rng As Range)
    Dim ColumnIndex_ As Variant
    Dim RowIndex_ As Variant
    
    Set this.Rng = Rng

    If this.ColumnIndex.IsRange Then
        ColumnIndex_ = RngToArr_(Rng.Rows(1))
    End If

    this.ColumnIndex = ArrayToTIndex_(ColumnIndex_) ' 1
    this.KeyColumn = CreateKeyColumn_(this.KeyColumn.Name)    ' 2
    
    If this.RowIndex.IsRange Then
        With Rng.Columns(this.KeyColumn.Ndx)
            If this.ColumnIndex.IsRange Then
                RowIndex_ = RngToArr_(.Resize(RowSize:=.Cells.Count - 1).Offset(1, 0))
            Else
                RowIndex_ = RngToArr_(.Resize(RowSize:=.Cells.Count))
            End If
        End With
    End If
    this.RowIndex = ArrayToTIndex_(RowIndex_)
End Sub

Public Sub Reset()
    fx.Raise_NotImplementedError
End Sub

Public Sub Trim(ByRef Columns As Variant)
    Dim ColName As Variant
    Dim Arr     As Variant
    Dim Area    As Range
    Dim Rng     As Range
    Dim i       As Long
    
    For Each ColName In CArr_(Columns)
        Set Rng = ColumnLoc(ColName)
        For Each Area In Rng.Areas
            Arr = fx.Range_ToArray2D(Area)
           
            For i = LBound(Arr, 1) To UBound(Arr, 1)
                Arr(i, 1) = VBA.Trim(Arr(i, 1))
            Next
            
            Area = Arr
        Next Area
    Next ColName
End Sub

Public Sub ToValue(ByRef Columns As Variant)
    Dim Col     As Variant
    Dim Rng     As Range
    Dim Area    As Range
    
    CalculateWhenManual
    
    For Each Col In CArr_(Columns)
        Set Rng = Me.ColumnLoc(Col)
        For Each Area In Rng.Areas
            Area.Copy
            Area.PasteSpecial xlPasteValues
            Application.CutCopyMode = False
        Next
    Next
    
End Sub

Public Sub AddFormula(ByRef Columns As Variant, Formula As String, Optional StartRow, Optional EndRow, Optional IsArrayFormula = False)
    Dim Rng     As Range
    Dim Area    As Range
    Dim ColName As Variant
    Dim SR      As Long
    Dim ER      As Long

    If IsMissing(StartRow) Then StartRow = 1
    If IsMissing(EndRow) Then EndRow = this.RowIndex.Size

    If EndRow > this.RowIndex.Size Then Raise_InvalidArgumentError_ "EndRow"
    If StartRow < 1 Or StartRow > EndRow Then Raise_InvalidArgumentError_ "StartRow"

    For Each ColName In CArr_(Columns)
        Set Rng = Me.ColumnLoc(ColName)
        For Each Area In Rng.Areas
            If StartRow <= Area.Rows.Count Then SR = StartRow Else SR = Area.Rows.Count
            If EndRow <= Area.Rows.Count Then ER = EndRow Else ER = Area.Rows.Count

            With Rng.Rows(Span_(SR, ER))
                If IsArrayFormula Then
                    .Cells(1, 1).FormulaArray = Formula
                Else
                    .Cells(1, 1).Formula = Formula
                End If
                If .Rows.Count Then .FillDown
            End With
        Next Area

    Next ColName
End Sub

Public Function ColumnToArray(ColumnName As Variant) As Variant
    ColumnToArray = ToArr1D_(Me.ColumnLoc(ColumnName))
End Function

Public Function RowToArray(RowName As Variant) As Variant
    RowToArray = ToArr1D_(Me.RowLoc(RowName))
End Function

Public Function ColumnToUniqueArray(ColumnName As Variant) As Variant
    ColumnToUniqueArray = ToUnqArr_(ToArr1D_(Me.ColumnLoc(ColumnName)))
End Function

Public Function RowToUniqueArray(RowName As Variant) As Variant
    RowToUniqueArray = ToUnqArr_(ToArr1D_(Me.RowLoc(RowName)))
End Function

Public Function ColumnToUniqueCells(ColumnName As Variant) As Range
    Set ColumnToUniqueCells = ToUnqRng_(Me.ColumnLoc(ColumnName))
End Function

Public Function RowToUniqueCells(RowName As Variant) As Range
    Set RowToUniqueCells = ToUnqRng_(Me.RowLoc(RowName))
End Function

Public Function ToListObject(ByVal TblName As String, Optional StyleName As String) As ListObject
    If StyleName = "" Then StyleName = "TableStyleLight1"
    fx.Raise_NotImplementedError
End Function

'::region Properties
Public Property Get Parent() As Worksheet
    Set Parent = this.Rng.Parent
End Property

Public Property Get ColumnExists(ByVal ColumnName) As Boolean
    ColumnExists = this.ColumnIndex.NameToNdx.Exists(ColumnName)
End Property

Public Property Get RowExists(ByVal RowName) As Boolean
    RowExists = this.RowIndex.NameToNdx.Exists(RowName)
End Property

Public Property Get ColumnIndex() As Variant
    ColumnIndex = this.ColumnIndex.Index
End Property

Public Property Get RowIndex() As Variant
    RowIndex = this.RowIndex.Index
End Property

'::region Range Members
Public Property Get ColumnLoc(ColumnNames As Variant) As Range
Attribute ColumnLoc.VB_UserMemId = 0
    Dim Name    As Variant
    Dim Names   As Variant
    Dim NdxItem As TIndexItem
    Dim Result  As Range
    
    Names = CArr_(ColumnNames)
    For Each Name In Names
        NdxItem = GetIndexItem_(this.ColumnIndex, Name)
        Set Result = Union_(Result, ColumnItemToRange_(NdxItem))
    Next ' Name

    Set ColumnLoc = Result
End Property

Public Property Get RowLoc(RowNames As Variant) As Range
    Dim Name    As Variant
    Dim Names   As Variant
    Dim NdxItem As TIndexItem
    Dim Result  As Range
    
    Names = CArr_(RowNames)
    For Each Name In Names
        NdxItem = GetIndexItem_(this.RowIndex, Name)
        Set Result = Union_(Result, RowItemToRange_(NdxItem))
    Next ' Name

    Set RowLoc = Result
End Property

Public Property Get At(Optional RowValue, Optional ColumnName) As Range
    fx.Raise_NotImplementedError
End Property

Public Property Get Loc(Optional RowNames, Optional ColumnNames) As Range
    Dim Result As Range
    
    If IsMissing(RowNames) And IsMissing(ColumnNames) Then
        Raise_InvalidArgumentError_ "Loc requires Row Index and/or Column Index"
    ElseIf IsMissing(RowNames) Then
        Set Result = ColumnLoc(ColumnNames)
    ElseIf IsMissing(ColumnNames) Then
        Set Result = RowLoc(RowNames)
    Else
        Set Result = Intersect(RowLoc(RowNames), ColumnLoc(ColumnNames))
    End If
    
    Set Loc = Result
End Property

Public Property Get iLoc(Optional RowIndex, Optional ColumnIndex) As Range
    Dim Result As Range
    
    If IsMissing(RowIndex) And IsMissing(ColumnIndex) Then
        Raise_InvalidArgumentError_ "iLoc requires Row Index and/or Column Index"
    ElseIf IsMissing(RowIndex) Then
        
    ElseIf IsMissing(ColumnIndex) Then
        
    Else
'        Set Result = Intersect(RowLoc(RowIndex), ColumnLoc(ColumnIndex))
    End If
    
    fx.Raise_NotImplementedError "iLoc", ModuleName
    Set iLoc = Result
End Property

Public Property Get FullRange() As Range
    Set FullRange = this.Rng
End Property

Public Property Get Range() As Range
    With this.RowIndex
        Set Range = RngImpl_Rows_(1, -1)
    End With
End Property

Public Property Get FirstColumn() As Range
    With this.ColumnIndex
        Set FirstColumn = RngImpl_Columns_(1)
    End With
End Property

Public Property Get LastColumn() As Range
    With this.ColumnIndex
        Set LastColumn = RngImpl_Columns_(-1)
    End With
End Property

Public Property Get FirstRow() As Range
    With this.RowIndex
        Set FirstRow = RngImpl_Rows_(1)
    End With
End Property

Public Property Get LastRow() As Range
    With this.RowIndex
        Set LastRow = RngImpl_Rows_(-1)
    End With
End Property

Public Property Get FirstCell(Optional ByVal ColumnName) As Range
    Set FirstCell = Me.ColumnLoc(ColumnName).Cells(1, 1)
End Property

Public Property Get LastCell(Optional ByVal ColumnName) As Range
    Dim Result     As Range
    With Me.ColumnLoc(ColumnName)
         Set Result = .Areas(.Areas.Count)
         Set Result = Result.Cells(Result.Rows.Count)
    End With
    Set LastCell = Result
End Property

Public Property Get HeadRange(Optional ByVal ColumnNames) As Range
    If this.ColumnIndex.IsRange Then
        Set HeadRange = this.Rng.Rows(1)
    Else
        Set HeadRange = Nothing
    End If
End Property

'::region FxRange Members
Public Property Get FirstFxColumn() As FxRange
    Dim Result As New FxRange
    With this.ColumnIndex
        Result.Init Me.FirstColumn, , this.RowIndex.Index, _
            .Index(0), this.AllowEmpty, this.AllowDuplicates
    End With
    Set FirstFxColumn = Result
End Property

Public Property Get FirstFxRow() As FxRange
    Dim Result As New FxRange
    With this.RowIndex
        Result.Init Me.FirstRow, this.KeyColumn.Name, .Index(0), _
            this.ColumnIndex.Index, this.AllowEmpty, this.AllowDuplicates
    End With
    Set FirstFxRow = Result
End Property

Public Property Get LastFxColumn() As FxRange
    Dim Result As New FxRange
    With this.ColumnIndex
        Result.Init Me.LastColumn, , this.RowIndex.Index, _
            .Index(.Size - 1), this.AllowEmpty, this.AllowDuplicates
    End With
    Set LastFxColumn = Result
End Property

Public Property Get LastFxRow() As FxRange
    Dim Result As New FxRange
    With this.RowIndex
        Result.Init Me.LastRow, this.KeyColumn.Name, .Index(.Size - 1), _
            this.ColumnIndex.Index, this.AllowEmpty, this.AllowDuplicates
    End With
    Set LastFxRow = Result
End Property

Public Property Get NthFxColumn(ByVal Ndx As Long) As FxRange
    Dim Result As New FxRange
    With this.ColumnIndex
        Result.Init RngImpl_Columns_(Ndx), , this.RowIndex.Index, _
            .Index(Ndx - 1), this.AllowEmpty, this.AllowDuplicates
    End With
    Set NthFxColumn = Result
End Property

Public Property Get NthFxRow(ByVal Ndx As Long) As FxRange
    Dim Result As New FxRange
    With this.RowIndex
        Result.Init RngImpl_Rows_(Ndx), this.KeyColumn.Name, .Index(Ndx - 1), _
            this.ColumnIndex.Index, this.AllowEmpty, this.AllowDuplicates
    End With
    Set NthFxRow = Result
End Property

'::region Counts
Public Property Get RowsCount() As Long
    RowsCount = this.RowIndex.Size
End Property

Public Property Get ColumnsCount() As Long
    ColumnsCount = this.ColumnIndex.Size
End Property

Public Property Get AllRowsCount() As Long
    AllRowsCount = OffsetRow_(Me.RowsCount)
End Property

Public Property Get AllColumnsCount() As Long
    AllColumnsCount = Me.ColumnsCount
End Property

Public Function SheetAddress(Optional AbsRow = True, Optional AbsCol = True) As String
    SheetAddress = VBA.Replace(Me.Range.Address(AbsRow, AbsCol, external:=True), _
        "[" & this.Rng.Parent.Parent.Name & "]", "")
End Function

':: region Raise
Private Sub Raise_InvalidArgumentError_(Optional Description = "")
    Err.Raise -2147220985, ModuleName, _
        Join(Array(ModuleName, "Invalid function argument ", Description), vbNewLine)
End Sub

Private Sub Raise_ObjectNotSetError_(Optional Description = "", Optional Source = "")
    Err.Raise -2147220988, ModuleName, _
        Join(Array(ModuleName, "Object not set error ", Description), vbNewLine)
End Sub

'::region Private
Private Sub CreateIndex_(Rng As Range, Optional KeyColumn, _
        Optional RowIndex_, Optional ColumnIndex_)
    ' NOTE: Order is important as each depends on previous one:
    '    1. this.ColumnIndex
    '    2. this.KeyColumn
    '    3. this.RowIndex
    Dim ColumnIndexIsRange As Boolean
    Dim RowIndexIsRange As Boolean

    If IsMissing(ColumnIndex_) Then
        ColumnIndex_ = RngToArr_(Rng.Rows(1))
        ColumnIndexIsRange = True
    End If
    
    this.ColumnIndex = ArrayToTIndex_(ColumnIndex_) ' 1
    this.KeyColumn = CreateKeyColumn_(KeyColumn)    ' 2

    If IsMissing(RowIndex_) Then
        With Rng.Columns(this.KeyColumn.Ndx)
            If ColumnIndexIsRange Then
                RowIndex_ = RngToArr_(.Resize(RowSize:=.Cells.Count - 1).Offset(1, 0))
            Else
                RowIndex_ = RngToArr_(.Resize(RowSize:=.Cells.Count))
            End If
        End With
        RowIndexIsRange = True
    End If
    
    this.RowIndex = ArrayToTIndex_(RowIndex_) ' 3
    
    this.ColumnIndex.IsRange = ColumnIndexIsRange
    this.RowIndex.IsRange = RowIndexIsRange
End Sub

Private Function ArrayToTIndex_(ArrNames As Variant) As TIndex
    Dim Name            As Variant
    Dim i               As Long
    Dim NameToNdx       As Scripting.Dictionary
    Dim NdxToName       As Scripting.Dictionary
    Dim IndexItems()    As TIndexItem
    Dim Result          As TIndex

    Set NameToNdx = CreateObject("Scripting.Dictionary")
    Set NdxToName = CreateObject("Scripting.Dictionary")
    
    i = 1
    Result.Index = ArrNames
    For Each Name In ArrNames
        If NameToNdx.Exists(Name) Then ' deal with duplicates
            AddDuplicateItem_ IndexItems(NameToNdx(Name)), i
        Else
            AddNewItem_ NameToNdx, Name, IndexItems, i
        End If
        
        NdxToName.Add i, Name
        i = i + 1
    Next Name

    Set Result.NameToNdx = NameToNdx
    Set Result.NdxToName = NdxToName
    Result.Items = IndexItems
    Result.Size = NdxToName.Count
    
    ArrayToTIndex_ = Result
End Function

Private Function CreateKeyColumn_(Optional KeyColumn) As TIndexItem
    Dim Result  As TIndexItem
    
    If IsMissing(KeyColumn) Then KeyColumn = this.ColumnIndex.Items(0).Name
    With this.ColumnIndex
        Result = .Items(.NameToNdx(KeyColumn))
    End With
    CreateKeyColumn_ = Result
End Function

Private Sub AddDuplicateItem_(Item As TIndexItem, NextNdx As Long)
    If Item.Count = 1 Then
        Set Item.MultiNdx = New Collection
        Item.MultiNdx.Add Item.Ndx
    End If
    
    Item.MultiNdx.Add NextNdx
    Item.Count = Item.Count + 1
End Sub

Private Sub AddNewItem_(NameToNdx As Scripting.Dictionary, Name As Variant, _
        IndexItems() As TIndexItem, ByVal Ndx As Long)
    Dim Item    As TIndexItem
    Dim sz      As Long
    
    Item.Ndx = Ndx
    Item.Count = 1
    Item.Name = Name
    
    sz = NameToNdx.Count
    NameToNdx.Add Name, sz
    ReDim Preserve IndexItems(sz)
    IndexItems(sz) = Item
End Sub

Private Function RngToArr_(Rng As Range) As Variant
    Dim R       As Range
    Dim i       As Long
    Dim Arr     As Variant

    ReDim Arr(Rng.Cells.Count - 1)

    For Each R In Rng.Cells
        Arr(i) = R.Value
        i = i + 1
    Next R

    RngToArr_ = Arr
End Function

Private Sub ValidateInput_(Rng As Range, Optional KeyColumn, _
        Optional RowIndex, Optional ColumnIndex)

End Sub

Private Function CArr_(val As Variant) As Variant
    If IsArray(val) Then
        CArr_ = val
    Else
        CArr_ = Array(val)
    End If
End Function

Private Function GetIndexItem_(TIdx As TIndex, Name As Variant) As TIndexItem
    GetIndexItem_ = TIdx.Items(TIdx.NameToNdx(Name))
End Function

Private Function Union_(ParamArray Arg1_Arg2__ArgN() As Variant) As Range
    Dim Result As Range
    Dim i      As Long

    Set Result = Arg1_Arg2__ArgN(0)
    For i = 1 To UBound(Arg1_Arg2__ArgN)
        If Result Is Nothing Then
            Set Result = Arg1_Arg2__ArgN(i)
        ElseIf Not Arg1_Arg2__ArgN(i) Is Nothing Then
            Set Result = Application.Union(Result, Arg1_Arg2__ArgN(i))
        End If
    Next i
    Set Union_ = Result
End Function

Private Function ColumnItemToRange_(Item As TIndexItem) As Range
    Dim Result  As Range
    Dim i       As Long
    
    If Item.Count = 1 Then
        ' TODO: what if is multirange
        Set Result = RngImpl_Columns_(Item.Ndx)
    Else
        For i = 1 To Item.Count
            Set Result = Union_(Result, RngImpl_Columns_(Item.MultiNdx(i)))
        Next i
    End If

    Set ColumnItemToRange_ = Result
End Function

Private Function RowItemToRange_(Item As TIndexItem) As Range
    Dim Result  As Range
    Dim i       As Long
    
    If Item.Count = 1 Then
        ' TODO: what if is multirange
        Set Result = RngImpl_Rows_(Item.Ndx)
    Else
        For i = 1 To Item.Count
            Set Result = Union_(Result, RngImpl_Rows_(Item.MultiNdx(i)))
        Next i
    End If

    Set RowItemToRange_ = Result
End Function

'::region RngImpl

Private Function RngImpl_Rows_(first As Long, Optional Last) As Range
    Dim Result As Range
        
    If this.Rng.Areas.Count > 1 Then
        fx.Raise_NotImplementedError "RngImpl_Rows_", ModuleName
    Else
        first = OffsetRow_(RowNum_(first))
        If IsMissing(Last) Then
            Set Result = this.Rng.Rows(first)
        Else
            If Last < 0 Then Last = this.RowIndex.Size + Last + 1
            Set Result = this.Rng.Rows(Span_(first, Last))
        End If
    End If

    Set RngImpl_Rows_ = Result
End Function

Private Function RngImpl_Columns_(first As Long, Optional Last) As Range
    Dim Result      As Range
    
    If this.Rng.Areas.Count > 1 Then
        fx.Raise_NotImplementedError "RngImpl_Columns_", ModuleName
    Else
        Dim RowSpan  As String
        RowSpan = Span_(OffsetRow_(1), this.RowIndex.Size)
        
        If IsMissing(Last) Then
            Set Result = this.Rng.Columns(ColNum_(first)).Rows(RowSpan)
        Else
            Set Result = this.Rng.Columns( _
                Span_(ColNum_(first), ColNum_(Last))).Rows(RowSpan)
        End If
    End If

    Set RngImpl_Columns_ = Result
End Function

Private Sub RngImpl_AddFormula(Rng As Range)
    Dim Area    As Range

    For Each Area In Rng.Areas
        fx.Raise_NotImplementedError
        
    Next Area

End Sub

Private Function OffsetRow_(ByVal Row As Long) As Long
    If this.ColumnIndex.IsRange Then
        OffsetRow_ = Row + 1
    Else
        OffsetRow_ = Row
    End If
End Function

Private Function Span_(ByVal first As Long, ByVal Last As Long) As String
    Span_ = first & ":" & Last + first - 1
End Function

Private Function RowNum_(ByVal RowNumber As Long) As Long
    If RowNumber < 0 Then
        RowNum_ = this.RowIndex.Size + RowNumber + 1
    Else
        RowNum_ = RowNumber
    End If
End Function

Private Function ColNum_(ByVal RowNumber As Long) As Long
    If RowNumber < 0 Then
        ColNum_ = this.ColumnIndex.Size + RowNumber + 1
    Else
        ColNum_ = RowNumber
    End If
End Function

Private Function ToUnqArr_(Arr As Variant) As Variant
    Dim Ndx     As Long
    Dim Result  As Object
    Set Result = CreateObject("Scripting.Dictionary")
        
    For Ndx = 1 To UBound(Arr, 1)
        Result(Arr(Ndx, 1)) = Empty
    Next
    
    ToUnqArr_ = Result.Keys()
End Function

Private Function ToUnqRng_(Rng As Range) As Range
    Dim Cell    As Range
    Dim Result  As Range
    Dim Dict    As Object
    Set Dict = CreateObject("Scripting.Dictionary")
        
    For Each Cell In Rng.Cells
        If Not Dict.Exists(Cell.Value) Then
            Dict(Cell.Value) = Empty
            Set Result = Union_(Result, Cell)
        End If
    Next
    
    Set ToUnqRng_ = Result
End Function

Private Function ToArr1D_(Rng As Range, Optional StartIndex As Long = 0) As Variant
    Dim Arr      As Variant
    Dim Cell_     As Range
    Dim i        As Long
    
    ReDim Arr(StartIndex To StartIndex + Rng.Cells.Count - 1)
    i = StartIndex
    For Each Cell_ In Rng.Cells
        Arr(i) = Cell_.Value
        i = i + 1
    Next Cell_
    
    ToArr1D_ = Arr
End Function
