VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "fx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Class description here"
''' Description:
''' Public class to utilize common functions
'''
''' VERSION 4.3.0 Sep 2021


Option Explicit
'::region DECLARATIONS

' Used in fx.Idx and fx.Fml_IndexMatchT
Public Enum fxRefTyp
[_First] = -1
    fxAt = [_First] * 2
    FxRng = fxAt * 2
    fxRef = FxRng * 2
    fxNum = fxRef * 2
    fxStr = fxNum * 2
[_Last] = fxStr - 1
End Enum

' For Browse sub
Public Enum fxDialogType
[_First] = 1
    fxFileDialog = [_First] * 2
    fxFolderDialog = fxFileDialog * 2
[_Last]
End Enum

Public Enum fxFileFilter
[_First] = 1
    fxAllFiles = [_First] * 2
    fxXlFiles = fxAllFiles * 2
    fxCSVFiles = fxXlFiles * 2
    fxTxtFiles = fxCSVFiles * 2
    fxWordFiles = fxTxtFiles * 2
    fxPdfFiles = fxWordFiles * 2
    fxImagesFiles = fxPdfFiles * 2
[_Last]
End Enum

Public Enum fxBorderLayout
[_First] = 1
    fxClearBorders = [_First]
    fxBottomDoubleBorders = fxClearBorders
    fxRightBorders = fxBottomDoubleBorders
    fxLeftBorders = fxRightBorders
    fxBottomBorders = fxLeftBorders
    fxTopBorders = fxBottomBorders
    [_fx_InsideVert_] = fxTopBorders
    [_fx_InsideHoriz_] = [_fx_InsideVert_]
[_Last]
    ' compounds
    [_fx_InsideBorders] = [_fx_InsideVert_] + [_fx_InsideHoriz_]
    fxLeftRightBorders = fxRightBorders + fxLeftBorders
    fxTopBottomBorders = fxTopBorders + fxBottomBorders
    fxOutsideBorders = fxLeftRightBorders + fxTopBottomBorders
    fxAllBorders = fxOutsideBorders + [_fx_InsideBorders]
End Enum

' colors
Public Enum fxColor
    fxBlack = 0: fxRed = 255: fxYellow = 65535: fxBlue = 16711680: fxSatinBlue = 12419407
    fxGrey = 8421504: fxLightGrey = 14277081: fxWhite = 16777215: fxLightBlue = 14994616
End Enum

Public Enum fxErrCode
    fxErr_None = 0 ' no errors
    fxErr_Value = 5
[_First] = vbObjectError + 513
    fxErr_System = [_First] + 1
    fxErr_InvalidArgument
    fxErr_ObjNotSet
    fxErr_FxInternal
    fxErr_Unknown
    fxErr_NotImplemented
    fxErr_Fatal
[_Last]
End Enum

Public Enum fxDirection
    fxNoDir = 0 ' no direction
[_First] = 1
    fxDown = [_First] * 2
    fxRight = fxDown * 2
    fxLeft = fxRight * 2
    fxUp = fxLeft * 2
    fxAutoDir = fxUp * 2
[_Last]
End Enum

Private Type TLastErr
    Num         As Long
    Text        As String
    ProcNames   As Collection
End Type


Private Type TThis
    Version         As String
    VersionDate     As Date
    Name            As String
    LastErr         As TLastErr
    Fso             As Object
    xlCalcState     As XlCalculation    ' hold Excel's current calculations state
    DEBUG_MODE      As Boolean
' Helper constants to use as keys for the special ranges in a Table
    FRNG            As String
    DRNG            As String
    HRNG            As String
    KEY_COL         As String
    TTL_ROW         As String
    TTL_COL         As String
    DEF             As String
    USER            As String
    DOC_STR         As String
End Type
Private this As TThis

'::region PUBLIC Methods

''' Adds text at the end of a file name, before the extension
'''  Params:
'''   FName As String, ByVal AppendTxt As String:
Public Function AddSuffixToFileName(ByVal FileName As String, ByVal AppendTxt As String) As String
    fx_Push "AddSuffixToFileName"
    Dim Ext As String

    Ext = GetFileExt(FileName)
    FileName = DropFileExt(FileName)
    AddSuffixToFileName = NormalizeFileName(FileName & AppendTxt & Ext)
    fx_Pop
End Function

''' Returns true if FindTxt is in the SearchArray
'''  Params:
'''   ByVal FindTxt As String,
'''   ByRef SearchArray As Variant:
Public Function Arr_ContainsStr(ByRef Arr As Variant, ByVal SearchStr As String) As Boolean
    fx_Push "Arr_ContainsStr"
    Arr_ContainsStr = (UBound(Filter(Arr, SearchStr)) > -1)
    fx_Pop
End Function

''' Returns the number of dimensions for given Array
Public Function Arr_DimsCount(Arr As Variant) As Long
    fx_Push "Arr_DimsCount"
    Dim Ndx     As Long
    Dim Res     As Long

    ' validate input
    On Error GoTo Failed
    Res = UBound(Arr, 1)

    On Error Resume Next
    Do
        Ndx = Ndx + 1
        Res = UBound(Arr, Ndx)
    Loop Until Err.Number <> 0

    Arr_DimsCount = Ndx - 1
    fx_Pop
Exit Function
Failed:
    Arr_DimsCount = -1
    fx_SetErr fxErr_InvalidArgument, Err.Description
End Function

''' Returns true if Array is empty
Public Function Arr_IsEmpty(Arr As Variant) As Boolean
    fx_Push "Arr_IsEmpty"
    Dim Result  As Boolean
    Dim LB      As Long
    Dim UB      As Long

    On Error GoTo Failed
    If IsArray(Arr) = False Then
        Raise_InvalidArgumentError "'Arr' is not an array"
    End If

    On Error Resume Next
    UB = UBound(Arr, 1)
    If (Err.Number <> 0) Then
        Result = True
    Else
        ' On rare occasions, Err.Number will be 0 for an unallocated, empty array.
        ' On these occasions, LBound is 0 and UBound is -1.
        LB = LBound(Arr)
        Result = (LB > UB)
    End If

    Arr_IsEmpty = Result
    fx_Pop
Exit Function
Failed:
    Arr_IsEmpty = True
    fx_SetErr Err.Number, Err.Description
End Function

''' Joining two arrays and returning new Array
''' Supports 1D Arrays or 2D Arrays with the same number
''' of Columns (same size of 2nd dimension)
'''  Params:
'''   ArrFirst As Variant,
'''   ArrSecond As Variant:
Public Function Arr_Join1D(ArrFirst As Variant, ArrSecond As Variant) As Variant
    fx_Push "Arr_Join1D"
    Dim DimsFirst   As Long
    Dim DimsSecond  As Long
    Dim i           As Long
    Dim Ndx         As Long
    Dim Result      As Variant

    ' checking input
    On Error GoTo Failed
    fx_Validate_TwoArrayInput ArrFirst, ArrSecond

    DimsFirst = Arr_DimsCount(ArrFirst)
    DimsSecond = Arr_DimsCount(ArrSecond)

    If DimsFirst <> 1 Or DimsSecond <> 1 Then
        Raise_InvalidArgumentError "Only 1D Arrays allowed."
    End If

    ReDim Result(Me.LenOf(ArrFirst) + Me.LenOf(ArrSecond) - 1)
    Ndx = 0

    For i = LBound(ArrFirst) To UBound(ArrFirst)
        Result(Ndx) = ArrFirst(i)
        Ndx = Ndx + 1
    Next

    For i = LBound(ArrSecond) To UBound(ArrSecond)
        Result(Ndx) = ArrSecond(i)
        Ndx = Ndx + 1
    Next

    Arr_Join1D = Result
    fx_Pop
Exit Function
Failed:
    fx_SetErr Err.Number, Err.Description
End Function

''' Joining two 2D arrays and returning new 2D Array
''' Supports Arrays with the same number of Columns
''' (same size of 2nd dimension)
'''  Params:
'''   ArrFirst As Variant,
'''   ArrSecond As Variant:
Public Function Arr_Join2D(ArrFirst As Variant, ArrSecond As Variant) As Variant
    fx_Push "Arr_Join2D"
    Dim Result          As Variant
    Dim ColumnsFirst    As Long
    Dim ColumnsSecond   As Long
    Dim LBRowResult     As Long
    Dim LBColResult     As Long
    Dim UBRowResult     As Long
    Dim UBColResult     As Long
    Dim RNdx            As Long
    Dim CNdx            As Long
    Dim RNdxResult      As Long
    Dim CNdxResult      As Long

    ' checking input
    On Error GoTo Failed
    fx_Validate_TwoArrayInput ArrFirst, ArrSecond
    If Arr_DimsCount(ArrFirst) <> 2 Or Arr_DimsCount(ArrSecond) <> 2 Then
        Raise_InvalidArgumentError "Only 2D Arrays allowed."
    End If

    ColumnsFirst = LenOf(ArrFirst, 2)
    ColumnsSecond = LenOf(ArrSecond, 2)

    If ColumnsFirst <> ColumnsSecond Then
        Raise_InvalidArgumentError _
            "Both arrays must have same size of 2nd dimension (same number of columns)."
    End If

    LBRowResult = LBound(ArrFirst, 1)
    LBColResult = LBound(ArrFirst, 2)
    UBRowResult = LBRowResult + LenOf(ArrFirst, 1) + LenOf(ArrSecond, 1) - 1
    UBColResult = LBColResult + ColumnsFirst - 1

    ReDim Result(LBRowResult To UBRowResult, LBColResult To UBColResult)
    RNdxResult = LBRowResult
    CNdxResult = LBColResult

    ' Copy elements of ArrFirst to Result
    For RNdx = LBound(ArrFirst, 1) To UBound(ArrFirst, 1)
        For CNdx = LBound(ArrFirst, 2) To UBound(ArrFirst, 2)
            Result(RNdxResult, CNdxResult) = ArrFirst(RNdx, CNdx)
            CNdxResult = CNdxResult + 1
        Next CNdx

        CNdxResult = LBColResult
        RNdxResult = RNdxResult + 1
    Next RNdx

    ' Copy elements of ArrSecond to Result
    For RNdx = LBound(ArrSecond, 1) To UBound(ArrSecond, 1)
        For CNdx = LBound(ArrSecond, 2) To UBound(ArrSecond, 2)
            Result(RNdxResult, CNdxResult) = ArrSecond(RNdx, CNdx)
            CNdxResult = CNdxResult + 1
        Next CNdx

        CNdxResult = LBColResult
        RNdxResult = RNdxResult + 1
    Next RNdx

    Arr_Join2D = Result
    fx_Pop
Exit Function
Failed:
    fx_SetErr Err.Number, Err.Description
End Function

''' Prints array shape to the Immediate window
Public Sub Arr_PrintShape(Arr As Variant)
    Dim vShape As Variant
    Dim i As Long
    Dim vDims As Variant
    Dim Result As String
    fx_Push "Arr_PrintShape"

    On Error GoTo Failed
    vShape = fx.Arr_Shape(Arr)
    Result = "Dimensions: " & vShape(0)

    For i = LBound(vShape) + 1 To UBound(vShape)
        vDims = vShape(i)
        Result = Result & fx.StrFormat(", {}:({},{})", i, (vDims(0)), (vDims(1)))
    Next

    Debug.Print Result
    fx_Pop
Exit Sub
Failed:
    fx_SetErr Err.Number, Err.Description
End Sub

''' Set difference of two arrays
''' Returns array with unique set of elements in 'first' that are not in 'second'
Public Function Arr_SetDifference(LeftArr As Variant, RightArr As Variant) As Variant
    fx_Push "Arr_SetDifference"
    Dim ResultArr   As Object: CreateObject ("Scripting.Dictionary")
    Dim i           As Long
    
    If Not IsArray(LeftArr) Then fx.Raise_InvalidArgumentError "LeftArr is not an Array"
    If Not IsArray(RightArr) Then fx.Raise_InvalidArgumentError "RightArr is not an Array"

    On Error GoTo Failed
    For i = LBound(LeftArr) To UBound(LeftArr)
        If IsError(Application.Match(LeftArr(i), RightArr, 0)) Then
            ResultArr(LeftArr(i)) = 0
        End If
    Next i

    Arr_SetDifference = ResultArr.Keys
    fx_Pop
Exit Function
Failed:
    fx_SetErr Err.Number, Err.Description
End Function

''' Returns Array representing the shape of the Arr argument.
''' Shape is as follows:
'''     Array(NumDims, (LBound(D_1), UBound(D_1)), ... , (LBound(D_n), UBound(D_n)))
''' Example:
'''     ReDim Arr(0 to 10, 0 to 2, 1 to 3)
'''     Arr_Shape(Arr)
''' Returns:
'''     Array(3, Array(0, 10), Array(0, 2), Array(1, 3))
Public Function Arr_Shape(Arr As Variant) As Variant
    fx_Push "Arr_Shape"
    Dim Result()    As Variant
    Dim Dims        As Long
    Dim Ndx         As Long

    On Error GoTo Failed
    If Not IsArray(Arr) Then
        Raise_InvalidArgumentError "Arr is not an Array"
    End If

    If Arr_IsEmpty(Arr) Then
        Raise_InvalidArgumentError "Arr is Empty"
    End If

    Dims = Arr_DimsCount(Arr)
    ReDim Result(Dims)

    Result(0) = Dims
    For Ndx = 1 To Dims
        Result(Ndx) = Array(LBound(Arr, Ndx), UBound(Arr, Ndx))
    Next

    Arr_Shape = Result
    fx_Pop
Exit Function
Failed:
    fx_SetErr Err.Number, Err.Description
End Function

''' 1D Array to delimited string
Public Function Arr_ToString1D(Arr As Variant, Optional Delim = ", ") As String
    Dim Result As String
    fx_Push "Arr_ToString1D"

    On Error GoTo Failed
    Arr_ToString1D = Join(Arr, Delim)
    fx_Pop
Exit Function
Failed:
    fx_SetErr Err.Number, Err.Description
End Function

''' Displays Explorer window and Returns String with selected file / folder
''' NOTE: No support for multi-select
'''  Params:
'''   Optional ByVal StartPath As String,
'''   Optional ByVal Caption As String,
'''   Optional DialogTyp As fxDialogType = fxFileDialog,
'''   Optional FilterTyp As fxFileFilter = fxAllFiles,
'''            can be also .xl*, .csv*, .pdf, .doc, .txt
Public Function Browse(Optional ByVal StartPath As String, Optional ByVal Caption, Optional DialogTyp As fxDialogType = fxFileDialog, Optional FilterTyp As fxFileFilter = fxAllFiles) As String
    fx_Push "Browse"
    Dim Dialog As FileDialog
    Dim sItem As String

    If Not FolderExists(StartPath) Then StartPath = Application.DefaultFilePath
    If Caption = "" Then Caption = "Select a " & IIf(DialogTyp = fxFileDialog, "File...", "Folder...")

    If DialogTyp = fxFileDialog Then
        Set Dialog = Application.FileDialog(msoFileDialogFilePicker)
        Dialog.Filters.Clear
        Call fx_AddFileFilters(FilterTyp, Dialog)
    ElseIf DialogTyp = fxFolderDialog Then
        Set Dialog = Application.FileDialog(msoFileDialogFolderPicker)
    End If

    With Dialog
        .Title = Caption
        .AllowMultiSelect = False
        If .Show <> -1 Then GoTo NextCode
        sItem = .SelectedItems(1)
    End With
NextCode:
    Browse = sItem
    Set Dialog = Nothing
    fx_Pop
End Function

''' Casts Value to Array
'''  Params:
'''   Val As Variant:
Public Function CArr(val As Variant) As Variant
Attribute CArr.VB_Description = "Casts Value to Array  \nParams:\n Val As Variant"
    fx_Push "CArr"
    On Error GoTo Failed
    If IsArray(val) Then
        CArr = val
    Else
        CArr = Array(val)
    End If
    fx_Pop
Exit Function
Failed:
    fx_SetErr Err.Number, Err.Description
End Function

Private Sub Class_Initialize()
    ' Use it to raise errors when developing / testing
    this.Version = "4.3.0" ' Sep 2021
    this.VersionDate = #9/10/2021#
    this.DEBUG_MODE = True
    this.Name = "fx"
    this.FRNG = "[_fxFRNG]"
    this.DRNG = "[_fxDRNG]"
    this.HRNG = "[_fxHRNG]"
    this.KEY_COL = "[_fxKEY_COL]"
    this.TTL_ROW = "[_fxTTL_ROW]"
    this.TTL_COL = "[_fxTTL_COL]"
    this.DEF = "[_fxDef]"
    this.USER = "[_fxUser]"
    this.DOC_STR = "'''"
    Set this.LastErr.ProcNames = New Collection
    With Application
        If .Workbooks.Count = 0 Then
            this.xlCalcState = xlCalculationAutomatic
        Else
            this.xlCalcState = .Calculation
        End If
    End With
End Sub

Private Sub Class_Terminate()
    Reset
End Sub

''' Reset fxErrCode and Error Description
Public Sub ClearLastError()
    With this.LastErr
        .Num = fxErr_None
        .Text = vbNullString
        Set .ProcNames = New Collection
    End With
End Sub

Public Sub Collection_SortInPlace(Coll As Collection)
    fx_Push "Collection_SortInPlace"
    Dim Tmp     As Variant
    Dim i       As Long
    Dim j       As Long
    Dim lLast   As Long

    On Error GoTo Failed
    lLast = Coll.Count
    For i = 1 To lLast
        For j = 1 + i To Coll.Count
            If Coll.Item(i) > Coll.Item(j) Then
                Tmp = Coll.Item(j)
                Coll.Remove j
                Coll.Add Tmp, Tmp, i
                lLast = j
            End If
        Next j
    Next i

    fx_Pop
Exit Sub
Failed:
    fx_SetErr Err.Number, Err.Description
End Sub

Public Function Collection_ToArray(Coll As Object, Optional StartIndex As Long = 0) As Variant
    fx_Push "Collection_ToArray"
    Dim Arr      As Variant
    Dim Cell     As Range
    Dim i        As Long
    
    With WorksheetFunction
        i = .Min(0, .Max(Coll.Count - 1, StartIndex))
    End With
    ReDim Arr(i To Coll.Count - 1)
    
    For i = StartIndex To Coll.Count - 1
        If IsObject(Coll.Item(i + 1)) Then
            Set Arr(i) = Coll.Item(i + 1)
        Else
            Arr(i) = Coll.Item(i + 1)
        End If
    Next i
    
    Collection_ToArray = Arr
    fx_Pop
Exit Function
Failed:
    fx_SetErr Err.Number, Err.Description
End Function

''' Similar to CollectionToArray.
''' Creates array with collection's elements by calling Object member function/property
''' on each element in the collection
Public Function Collection_ToObjMemberArray(CollLike As Object, ByVal MemberName As String, Optional ByVal MemberType As VbCallType = VbGet, Optional Limit As Variant) As Variant
    fx_Push "Collection_ToObjMemberArray"
    Dim Tmp As Variant
    Dim Lim As Long
'    Dim v   As Variant
    Dim i   As Long
    Dim ln  As Long
    
    On Error GoTo Failed
    ln = CollLike.Count
    If ln = 0 Then
        fx_Pop
        Exit Function
    End If
    
    With WorksheetFunction
        Lim = (IIf(IsMissing(Limit), ln, .Max(.Min(Limit, ln), 1)))
    End With

    ReDim Tmp(0 To ln - 1)
    For i = 0 To ln - 1
        If i > Lim - 1 Then Exit For
        Tmp(i) = CStr(VBA.CallByName(CollLike.Item(i + 1), MemberName, MemberType))
    Next i
    Collection_ToObjMemberArray = Tmp ' Returning...
    fx_Pop
Exit Function
Failed:
    fx_SetErr fxErr_System, Err.Description
End Function

''' Creates delimited string with values in a collection
Public Function Collection_ToString(CollLike As Object, Optional ByVal Delim = ", ", Optional Limit As Variant) As String
    fx_Push "Collection_ToString"
    Dim Tmp As Variant
    Dim Lim As Long
    Dim i   As Long
    Dim s   As String
    Dim ln  As Long
    
    ln = CollLike.Count
    If ln = 0 Then
        fx_Pop
        Exit Function
    End If
    
    With WorksheetFunction
        Lim = (IIf(IsMissing(Limit), ln, .Max(.Min(Limit, ln), 1)))
    End With
    
    ReDim Tmp(1 To Lim)
    For i = 1 To ln
        If i > Lim Then
            s = "..."
            Exit For
        End If
        
        Tmp(i) = CollLike.Item(i)
    Next i
    Collection_ToString = (Join(Tmp, Delim) & vbNewLine & s)   ' Returning...
    fx_Pop
End Function

''' Similar to CollectionsToString.
''' Creates a delimited string with values by calling Object member function/property
''' on each element in the collection
Public Function Collection_ToObjMemberString(CollLike As Object, ByVal MemberName As String, Optional ByVal MemberType As VbCallType = VbGet, Optional ByVal Delim = ", ", Optional Limit As Variant) As String
    fx_Push "Collection_ToObjMemberString"
    Dim Tmp As Variant
    Dim Str As String
    Dim ln  As Long
    Dim Lim As Long
    
    With WorksheetFunction
        Lim = (IIf(IsMissing(Limit), ln, .Max(.Min(Limit, ln), 1)))
    End With
    
    Tmp = Collection_ToObjMemberArray(CollLike, MemberName, MemberType, Lim)
    If fx.Arr_IsEmpty(Tmp) Then
        fx_Pop
        Exit Function
    End If
    
    If UBound(Tmp) + 1 > Lim Then Str = "..."
    Collection_ToObjMemberString = (Join(Tmp, Delim) & vbNewLine & Str) ' Returning...
    fx_Pop
Exit Function
Failed:
    fx_SetErr fxErr_System, Err.Description
End Function

''' Similar to CollectionToString.
''' Creates delimited string with keys/values in a Scripting.Dictionary
Public Function Dict_ToString(DictLike As Object, Optional Format = "keys_and_values", Optional ByVal Delim = ", ", Optional Limit As Variant) As String
    fx_Push "Dict_ToString"
    Dim Tmp As Variant
    Dim Lim As Long
    Dim i   As Long
    Dim v   As Variant
    Dim s   As String
    Dim ln  As Long
    
    ln = DictLike.Count
    If ln = 0 Then
        fx_Pop
        Exit Function
    End If
    
    With WorksheetFunction
        Lim = (IIf(IsMissing(Limit), ln, .Max(.Min(Limit, ln), 1)))
    End With
        
    ReDim Tmp(1 To Lim)
    For Each v In DictLike.Keys
        i = i + 1
        If i > Lim Then
            s = "..."
            Exit For
        End If
        If Format = "keys_and_values" Then
            Tmp(i) = CStr(v) & " : " & CStr(DictLike(v))
        ElseIf "keys" Then
            Tmp(i) = CStr(v)
        ElseIf "values" Then
            Tmp(i) = CStr(DictLike(v))
        End If
    Next v
    Dict_ToString = (Join(Tmp, Delim) & vbNewLine & s) ' Returning...
    fx_Pop
End Function

''' Drawing Borders
Public Sub DrawBorders(Rng As Range, Optional Layout As fxBorderLayout = fxAllBorders, Optional LineWeight As XlBorderWeight = xlThin, Optional ByVal ColorRGB = 0)
    fx_Push "DrawBorders"
    Dim BIdx      As XlBordersIndex
    Dim LineStyle As XlLineStyle
    Dim lMax As Long, lMin As Long, lVal As Long

    If ColorRGB < fxBlack Then ColorRGB = fxBlack
    If ColorRGB > fxWhite Then ColorRGB = fxWhite
    LineStyle = xlContinuous ' Only clear fxClearBorders can change this one
    lVal = Layout
    lMin = fxBorderLayout.[_First]

newLoop:
    On Error GoTo Failed
    lMax = fxBorderLayout.[_Last] - 1
    Do While lMax > lMin And lVal > lMin
        Do While lVal >= lMax
            lVal = lVal - lMax
            GoSub drawBorder ' Drawing lMax
        Loop
        lMax = lMax / 2
    Loop
    fx_Pop
Exit Sub
Failed:
    fx_SetErr Err.Number, Err.Description
Exit Sub
drawBorder:
    Select Case lMax
        Case fxClearBorders: LineStyle = xlNone: lVal = fxAllBorders: GoTo newLoop
        Case fxBottomDoubleBorders: BIdx = xlEdgeBottom: LineStyle = xlDouble: LineWeight = xlThin
        Case fxTopBorders: BIdx = xlEdgeTop
        Case fxBottomBorders: BIdx = xlEdgeBottom
        Case [_fx_InsideVert_]: BIdx = xlInsideVertical
        Case [_fx_InsideHoriz__]: BIdx = xlInsideHorizontal
        Case fxLeftBorders: BIdx = xlEdgeLeft
        Case fxRightBorders: BIdx = xlEdgeRight
        Case Else
            Err.Raise fxErr_Value, , "Unknown Border Layout"
    End Select

    With Rng.Borders(BIdx)
        .LineStyle = LineStyle
        If LineStyle = xlContinuous Then
            .Color = ColorRGB
            .Weight = LineWeight
        End If
    End With
    Return
End Sub

''' Returns File Name without the extension
'''  Params:
'''   FName As String:
Public Function DropFileExt(FName As String) As String
    fx_Push "DropFileExt"
    Dim sExt As String

    On Error GoTo Failed
    sExt = GetFileExt(FName)
    If sExt <> "" Then
        DropFileExt = Left(FName, Len(FName) - Len(sExt))
    Else
        DropFileExt = FName
    End If
    fx_Pop
Exit Function
Failed:
    fx_SetErr Err.Number, Err.Description
End Function

''' Returns File Name without the path
'''  Params:
'''   FullFileName As String:
Public Function DropFilePath(FullFileName As String) As String
    fx_Push "DropFilePath"
    Dim v As Variant

    On Error GoTo Failed
    v = Split(FullFileName, Application.PathSeparator)
    DropFilePath = v(UBound(v))
    fx_Pop
Exit Function
Failed:
    fx_SetErr Err.Number, Err.Description
End Function

''' Terminates program execution / calls Err.Raise in DEBUG_MODE
Public Sub FatalError(Optional ByVal ErrMsg = "", Optional ByVal Source = "")
    Static Thrown   As Boolean
    Dim EndMsg      As String
    Dim v           As Variant

    If ErrMsg = "" And this.LastErr.Num Then
        ErrMsg = this.LastErr.Text
        Source = Empty
    End If
    If Source <> "" Then Source = "Source: " & Source

    If this.DEBUG_MODE And Not Thrown Then
        Thrown = True
        Call Err.Raise(vbObjectError = fxErrCode.[_First], , _
            "DEBUG MODE ERROR." & vbNewLine & Source & vbNewLine & ErrMsg)
    End If

    With Application
        ' resuming excel's initial calculation state only if there are workbooks open
        If Workbooks.Count Then .Calculation = this.xlCalcState
        .ScreenUpdating = True
        .EnableEvents = True
        .DisplayAlerts = True
        .StatusBar = False
        .Cursor = xlDefault
    End With

    For Each v In Array(Source, ErrMsg)
        If v <> vbNullString Then EndMsg = EndMsg & vbNewLine & v
    Next v

    If EndMsg <> vbNullString Then
        MsgBox IIf(Thrown, "", "ERROR: ") & EndMsg, vbCritical
    End If
    If this.LastErr.Num <> fxErr_None Then Me.ClearLastError
    End ' program execution
End Sub

''' Returns True if FullFileName exists
Public Function FileExists(ByVal FullFileName As String) As Boolean
    fx_Push "FileExists"
    FileExists = Fso.FileExists(FullFileName)
    fx_Pop
End Function


''' Incrementing a number at the end of a File Name, while File exists
'''  Params:
'''   ByVal FullFileName As String:
Public Function File_FreeFileName(ByVal FullFileName As String) As String
    fx_Push "File_FreeFileName"
    Dim Ext     As String
    Dim Name    As String
    Dim i       As Long

    If Not FileExists(FullFileName) Then GoTo Finish
    Ext = GetFileExt(FullFileName)
    Name = DropFileExt(FullFileName)

    Do
        i = i + 1
        FullFileName = Name & " (" & CStr(i) & ")" & Ext
    Loop While FileExists(FullFileName)
Finish:
    File_FreeFileName = FullFileName
    fx_Pop
End Function

''' Writing / Appending Log Message to Text File
'''  Params:
'''   LogMsg As String,
'''   Optional ByRef LogPath As String,
'''   Optional ByRef LogFile As String = "fx.log":
Public Function File_Log(LogMsg As String, Optional ByRef LogPath, Optional ByRef LogFile = "fx.log") As Boolean
    fx_Push "File_Log"
    Dim lFile As Integer
    Dim FName As String
    Dim PathName As String
    Dim bLogStarted As Boolean

    On Error GoTo Failed
    If Not bLogStarted Then
        ' check input
        FName = IIf(LogFile = "fx.log", LogFile, NormalizeFileName(LogFile))

        If Not FolderExists(LogPath) Then
            Err.Raise fxErr_InvalidArgument, , "LogPath not found"
        End If
        PathName = TrailingSlash(PathName)
        bLogStarted = True
    End If

    ' write to the file
    lFile = FreeFile
    Open PathName & FName For Append As #lFile
    Print #lFile, LogMsg
    Close #lFile
    File_Log = True
    fx_Pop
Exit Function
Failed:
    fx_SetErr Err.Number, Err.Description
    Close #lFile
    File_Log = False
End Function

''' Returns string containing INDEX / MATCH formula for a single column lookup
''' NOTE: final string does not have '=' sign
'''  Params:
'''   ReturnColumn As Range, LookupCell As Range,
'''   LookupColumn As Range,
'''   Optional EndOnError As Boolean = False:
Public Function Fml_IndexMatch(ReturnColumn As Range, LookupCell As Range, LookupColumn As Range, Optional EndOnError = False) As String
    fx_Push "Fml_IndexMatch"
    Dim sLookCell As String
    Dim sLookCol  As String
    Dim sRetCol   As String

    On Error GoTo Failed
    sLookCell = LookupCell.Address(False, True, xlA1, True)
    sLookCol = LookupColumn.Address(True, True, xlA1, True)
    sRetCol = ReturnColumn.Address(True, True, xlA1, True)

    Fml_IndexMatch = "INDEX(" & sRetCol & ",MATCH(" & sLookCell & "," & sLookCol & ",0))"
    fx_Pop
Exit Function
Failed:
    fx_SetErr fxErr_System, Err.Description
    If EndOnError Then FatalError
End Function

    ''' Returns String containing INDEX / MATCH formula for a Table-like Object \nNOTE: Result does not have '=' sign
'''  Params:
'''   TblObj As Object, ReturnColumn As String, LookupVal As Variant, LookupColumn
'''   Optional LookupTyp As fxRefTyp = fxAt,
'''   Optional EndOnError As Boolean = False:
Public Function Fml_IndexMatchT(TblObj As Object, ReturnColumn As String, ByVal LookupVal As Variant, ByVal LookupColumn As Variant, Optional LookupTyp As fxRefTyp = fxAt, Optional EndOnError = False) As String
    fx_Push "Fml_IndexMatchT"
    Dim sLookCell As String
    Dim sLookCol  As String
    Dim sRetCol   As String

    On Error GoTo Failed
    sLookCell = fx_RefToStr(LookupTyp, LookupVal)

    ' Strings for the return and lookup columns
    sLookCol = fx_TblColToAddr(TblObj, LookupColumn)
    If this.LastErr.Num Then GoTo Failed
    sRetCol = fx_TblColToAddr(TblObj, ReturnColumn)
    If this.LastErr.Num Then GoTo Failed

    Fml_IndexMatchT = "INDEX(" & sRetCol & ",MATCH(" & sLookCell & "," & sLookCol & ",0))"
    fx_Pop
Exit Function
Failed:
    If this.LastErr.Num = fxErr_None Then
        fx_SetErr Err.Number, Err.Description
    End If
    If EndOnError Then FatalError
End Function

    ''' Returns String containing SUMIFS formula
''' NOTE: final String does  not have '=' sign
'''  Params:
'''     ParamArray SumRng__CrRng__ValRef__RefT() As Variant:
'''     SumRng   As Range
'''     Pairs CriteriaRng : Criteria
'''     ReferenceType
Public Function Fml_SumIfs(ParamArray SumRng__CrRng__ValRef__RefT() As Variant) As String
    fx_Push "Fml_SumIfs"
    Dim v       As Variant
    Dim Ref     As fxRefTyp
    Dim i       As Long
    Dim sz      As Long
    Dim sFml    As String
    Dim sCrVal  As String
    Dim sCrTyp  As String
    Dim rSumRng As Range
    Dim rCrRng  As Range

    On Error GoTo Failed
    v = SumRng__CrRng__ValRef__RefT ' shorthand
    sz = UBound(v)

    If sz < 3 Or (sz Mod 2) = 0 Then ' checking the initial input
        Err.Raise fxErr_InvalidArgument, , _
            "ParamArray: Wrong number of arguments"
    End If

    Set rSumRng = v(0)
    Ref = v(sz)
    sCrTyp = TypeName(v(2)) ' data type of the first criteria, that's passed by user

    For i = 1 To sz - 1 Step 2
        Set rCrRng = Nothing
        Set rCrRng = v(i)
        sCrVal = fx_RefToStr(Ref, v(i + 1))

        ' validate both inputs
        If rCrRng.Rows.Count <> rSumRng.Rows.Count Then
            Err.Raise fxErr_FxInternal, , "CriteriaRng Rows.Count mismatch"
        ElseIf TypeName(v(i + 1)) <> sCrTyp Then
            Err.Raise fxErr_FxInternal, , "Criteria type mismatch"
        End If

        ' Add CriteriaRng - Criteria pair to the formula
        sFml = sFml & "," & rCrRng.Address(True, True, xlA1, True) & "," & sCrVal

    Next i

    Fml_SumIfs = "SUMIFS(" & rSumRng.Address(True, True, xlA1, True) & sFml & ")" ' Returns...
    fx_Pop
Exit Function
Failed:
    If this.LastErr.Num = fxErr_None Then
        fx_SetErr Err.Number, Err.Description
    End If
End Function


    ''' Returns String containing SUMIFS formula for a Table-like Object
''' NOTE: final String does not have '=' sign
'''  Params:
'''     ParamArray TblObj__SumCol__CrCol__ValRef__RefT() As Variant:
'''                Tbl-LikeObj As Object: ListObject, Dictionary, or clsSmartRange, fxRng
'''     SumColumn
'''     Pairs of CriteriaColumn : Criteria
'''     ReferenceType as fxRefTyp
Public Function Fml_SumIfsT(ParamArray TblObj__SumCol__CrCol__ValRef__RefT() As Variant) As String
    fx_Push "Fml_SumIfsT"
    Dim v       As Variant
    Dim Ref     As fxRefTyp
    Dim i       As Long
    Dim sz      As Long
    Dim sFml    As String
    Dim sCrVal  As String
    Dim sCrTyp  As String
    Dim sSumCol As String
    Dim sCrCol  As String
    Dim TblObj  As Object

    On Error GoTo Failed
    v = TblObj__SumCol__CrCol__ValRef__RefT ' shorthand
    sz = UBound(v)

    If sz < 4 Or (sz Mod 2) = 1 Then ' checking the initial input
        Err.Raise fxErr_InvalidArgument, , _
            "ParamArray: Wrong number of arguments"
    End If

    Set TblObj = v(0)
    sSumCol = fx_TblColToAddr(TblObj, CStr(v(1)))
    If this.LastErr.Num Then GoTo Failed

    Ref = v(sz)
    sCrTyp = TypeName(v(3)) ' data type of the first criteria, that's passed by user

    For i = 2 To sz - 1 Step 2
        sCrCol = Empty
        sCrCol = fx_TblColToAddr(TblObj, CStr(v(i)))
        If this.LastErr.Num Then GoTo Failed
        sCrVal = fx_RefToStr(Ref, v(i + 1))

        ' validate criteria input
        If TypeName(v(i + 1)) <> sCrTyp Then
            Err.Raise fxErr_FxInternal, , "Criteria type mismatch"
        End If

        ' add CriteriaRng : Criteria pair to the formula
        sFml = sFml & "," & sCrCol & "," & sCrVal
    Next i

    Fml_SumIfsT = "SUMIFS(" & sSumCol & sFml & ")" ' Return...
    fx_Pop
Exit Function
Failed:
    If this.LastErr.Num = fxErr_None Then
        fx_SetErr Err.Number, Err.Description
    End If
End Function

    ''' Returns String containing SUMIFS formula
''' NOTE: final String does not have '=' sign
'''  Params:
'''     ParamArray CrRng__ValRef__RefT() As Variant:
'''     Pairs CriteriaRng : Criteria
'''     ReferenceType
Public Function Fml_CountIfs(ParamArray CrRng__ValRef__RefT() As Variant) As String
    fx_Push "Fml_CountIfs"
    Dim v       As Variant
    Dim Ref     As fxRefTyp
    Dim i       As Long
    Dim sz      As Long
    Dim lRows   As Long
    Dim sFml    As String
    Dim sCrVal  As String
    Dim sCrTyp  As String
    Dim rCrRng  As Range

    On Error GoTo Failed
    v = CrRng__ValRef__RefT ' shorthand
    sz = UBound(v)

    If sz < 2 Or (sz Mod 2) = 1 Then ' checking the initial input
        Err.Raise fxErr_InvalidArgument, , _
            "ParamArray: Wrong number of arguments"
    End If

    Ref = v(sz)
    lRows = v(0).Rows.Count
    sCrTyp = TypeName(v(1)) ' data type of the first criteria, that's passed by user

    For i = 0 To sz - 1 Step 2
        Set rCrRng = Nothing
        Set rCrRng = v(i)
        sCrVal = fx_RefToStr(Ref, v(i + 1))

        ' validate both inputs
        If rCrRng.Rows.Count <> lRows Then
            Err.Raise fxErr_FxInternal, , "CriteriaRng Rows.Count mismatch"
        ElseIf TypeName(v(i + 1)) <> sCrTyp Then
            Err.Raise fxErr_FxInternal, , "Criteria type mismatch"
        End If

        ' Add CriteriaRng - Criteria pair to the formula
        sFml = sFml & IIf(i = 0, "", ",") & rCrRng.Address(True, True, xlA1, True) & "," & sCrVal

    Next i

    Fml_CountIfs = "COUNTIFS(" & sFml & ")" ' Returns...
    fx_Pop
Exit Function
Failed:
    If this.LastErr.Num = fxErr_None Then
        fx_SetErr Err.Number, Err.Description
    End If
End Function

''' Returns String containing COUNTIFS formula for a Table-like Object
''' NOTE: final String does not have '=' sign
'''  Params:
'''   ParamArray TblObj__CrCol__ValRef__RefT() As Variant:
'''              Tbl-LikeObj As Object: ListObject, Dictionary, or clsSmartRange, fxRng
'''   Pairs of CriteriaColumn : Criteria
'''   ReferenceType as fxRefTyp
Public Function Fml_CountIfsT(ParamArray TblObj__CrCol__ValRef__RefT() As Variant) As String
    fx_Push "Fml_CountIfsT"
    Dim v       As Variant
    Dim Ref     As fxRefTyp
    Dim i       As Long
    Dim sz      As Long
    Dim sFml    As String
    Dim sCrVal  As String
    Dim sCrTyp  As String
    Dim sCrCol  As String
    Dim TblObj  As Object

    On Error GoTo Failed
    v = TblObj__CrCol__ValRef__RefT ' shorthand
    sz = UBound(v)

    If sz < 3 Or (sz Mod 2) = 0 Then ' checking the initial input
        Err.Raise fxErr_InvalidArgument, , _
            "ParamArray: Wrong number of arguments"
    End If

    Set TblObj = v(0)
    If this.LastErr.Num Then GoTo Failed

    Ref = v(sz)
    sCrTyp = TypeName(v(2)) ' data type of the first criteria, that's passed by user

    For i = 1 To sz - 1 Step 2
        sCrCol = Empty
        sCrCol = fx_TblColToAddr(TblObj, CStr(v(i)))

        If this.LastErr.Num Then GoTo Failed
        sCrVal = fx_RefToStr(Ref, v(i + 1))

        ' validate criteria input
        If TypeName(v(i + 1)) <> sCrTyp Then
            Err.Raise fxErr_FxInternal, , "Criteria type mismatch"
        End If

        ' add CriteriaRng : Criteria pair to the formula
        sFml = sFml & IIf(i = 0, "", ",") & sCrCol & "," & sCrVal
    Next i

    Fml_CountIfsT = "COUNTIFS(" & sFml & ")" ' Return...
    fx_Pop
Exit Function
Failed:
    If this.LastErr.Num = fxErr_None Then
        fx_SetErr Err.Number, Err.Description
    End If
End Function

''' Returns String containing SUM formula \NOTE: final String does not have '=' sign
'''  Params:
'''   ParamArray Values_RefT() As Variant:
'''   Values
'''   Reference Type As fxRefTyp
Public Function Fml_Sum(ParamArray Values_RefT() As Variant) As String
    fx_Push "Fml_Sum"
    Dim i       As Long
    Dim sz      As Long
    Dim sFml    As String
    Dim sVal    As String
    Dim v       As Variant
    Dim Ref     As fxRefTyp

    On Error GoTo Failed
    v = Values_RefT ' shorthand
    sz = UBound(v)

    If sz < 1 Then ' checking the initial input
        Err.Raise fxErr_InvalidArgument, , _
            "ParamArray: Wrong number of arguments"
    End If

    Ref = v(sz)
    For i = 0 To sz - 1
        Select Case v(i)
            Case "-", "+"
                sFml = sFml & IIf(i = 0, "", ",") & v(i)
                sFml = sFml & fx_RefToStr(Ref, v(i + 1))
                i = i + 1
            Case ":"
                sFml = sFml & v(i) & fx_RefToStr(Ref, v(i + 1))
                i = i + 1
            Case Else
                sVal = fx_RefToStr(Ref, v(i))
                sFml = sFml & IIf(i = 0, "", ",") & sVal
        End Select
    Next i

    Fml_Sum = "SUM(" & sFml & ")" ' Returns...
    fx_Pop
Exit Function
Failed:
    If this.LastErr.Num = fxErr_None Then
        fx_SetErr Err.Number, Err.Description
    End If
End Function

''' Returns String containing SUM formula for column in a Table-like Object
''' NOTE: final String does not have '=' sign
'''  Params:
'''   ParamArray TblObj__Columns() As Variant:
Public Function Fml_SumT(ParamArray TblObj__Columns() As Variant) As String
    fx_Push "Fml_SumT"
    Dim i       As Long
    Dim sz      As Long
    Dim sVal    As String
    Dim Vals    As New Collection
    Dim v       As Variant
    Dim TblObj  As Object

    On Error GoTo Failed
    v = TblObj__Columns ' shorthand
    sz = UBound(v)

    If sz < 1 Then ' checking the initial input
        Err.Raise fxErr_InvalidArgument, , _
            "ParamArray: Wrong number of arguments"
    End If

    Set TblObj = v(0)
    For i = 1 To sz
        sVal = fx_TblColToAddr(TblObj, CStr(v(i)))
        If this.LastErr.Num Then GoTo Failed
        Vals.Add sVal
    Next i

    Fml_SumT = "SUM(" & Collection_ToString(Vals, ",") & ")" ' Returns...
    fx_Pop
Exit Function
Failed:
    If this.LastErr.Num = fxErr_None Then
        fx_SetErr Err.Number, Err.Description
    End If
End Function

''' Fills down formula in a column range
'''  Params:
'''   ColumnRng As Range
'''   FormulaTxt As String
'''   Optional IsArrayFormula As Boolean = False
'''   Optional EndOnError As Boolean = False:
Public Sub Fml_FillDown(ColumnRng As Range, FormulaTxt As String, Optional IsArrayFormula = False, Optional EndOnError = False)
    fx_Push "Fml_FillDown"
    ' checking input
    On Error GoTo Failed
    If ColumnRng Is Nothing Then
        Err.Raise fxErr_ObjNotSet, , "ColumnRng"
    ElseIf ColumnRng.Rows(1).Columns.Count > 1 Then
        Err.Raise fxErr_InvalidArgument, _
            "ColumnRng: Parameter must be a single column"
    ElseIf ColumnRng.Rows(1).Columns.Count > 1 Then
        Err.Raise fxErr_InvalidArgument, _
            "FormulaTxt: FormulaArray property has a character limit of 255"
    End If

    With ColumnRng
        ' check for array formula
        If IsArrayFormula Then
            .Cells(1, 1).FormulaArray = FormulaTxt
        Else
            .Cells(1, 1).Formula = FormulaTxt
        End If

        ' fill down is to avoid column calculation
        If .Rows.Count > 1 Then
            .Cells(1, 1).AutoFill Destination:=ColumnRng, Type:=xlFillValues
        End If
    End With
    fx_Pop
Exit Sub
Failed:
    If this.LastErr.Num = fxErr_None Then
        fx_SetErr Err.Number, Err.Description
    End If
End Sub

''' Wraps Formula test in IFERROR formula
''' NOTE: final String does not have '=" sign
'''  Params:
'''   FormulaTxt As String, Optional ByVal ValueIfErr As String = """""":
Public Function Fml_IfErr(FormulaTxt As String, Optional ByVal ValueIfErr As String = """""") As String
    fx_Push "Fml_IfErr"
    Fml_IfErr = "IFERROR(" & FormulaTxt & "," & ValueIfErr & ")"
    fx_Pop
End Function

''' Returns True if FolderName exists
Public Function FolderExists(ByVal FolderName As String) As Boolean
    fx_Push "FolderExists"
    FolderExists = Fso.FolderExists(FolderName)
    fx_Pop
End Function

''' Returns column number from String representing valid Column letter
'''  Params:
'''   ByVal ColumnLetter As String:
Public Function GetColumnNumberFromLetter(ByVal ColumnLetter As String) As Long
    fx_Push "GetColumnNumberFromLetter"
    On Error GoTo Failed
    GetColumnNumberFromLetter = Range(ColumnLetter & 1).Column
    fx_Pop
Exit Function
Failed:
    fx_SetErr Err.Number, Err.Description
End Function

''' Returns the last column in a worksheet, using Cells.Find
Public Function GetLastColumn(Optional TargetSheet) As Long
    fx_Push "GetLastColumn"

    If TargetSheet Is Nothing Then Set TargetSheet = ActiveSheet
    On Error Resume Next
    With TargetSheet
        GetLastColumn = .Cells.Find( _
            what:="*", _
            After:=.Range("A1"), _
            SearchOrder:=xlByColumns, _
            SearchDirection:=xlPrevious).Column
    End With
    If GetLastColumn = 0 Then GetLastColumn = 1
    fx_Pop
End Function

''' Returns the last row in a Worksheet, using Cells.Find()
Public Function GetLastRow(Optional TargetSheet) As Long
    fx_Push "GetLastRow"

    If TargetSheet Is Nothing Then Set TargetSheet = ActiveSheet
    On Error Resume Next
    With TargetSheet
        GetLastRow = .Cells.Find( _
            what:="*", _
            After:=.Range("A1"), _
            SearchOrder:=xlByRows, _
            SearchDirection:=xlPrevious).Row
    End With
    If GetLastRow = 0 Then GetLastRow = 1
    fx_Pop
End Function

''' Returns Last Row in Sheet Column
'''  Params:
'''   ByVal TargetColumn As Variant > Number or Letter,
'''   Optional ByRef TargetSheet As Worksheet,
'''   Optional EndOnError As Boolean = False:
Public Function GetLastRowInColumn(ByVal TargetColumn As Variant, Optional ByRef TargetSheet As Worksheet, Optional EndOnError = False) As Long
    fx_Push "GetLastRowInColumn"
    Dim LCol As Long

    ' validate input
    On Error GoTo Failed
    Select Case TypeName(TargetColumn)
        Case "String"
            LCol = GetColumnNumberFromLetter(TargetColumn)
        Case "Long", "Integer"
            LCol = CLng(TargetColumn)
        Case Else
            Err.Raise fxErr_InvalidArgument, , _
                StrFormat("{} is invalid type for TargetColumn.", _
                    TypeName(TargetColumn))
    End Select

    If LCol < 1 Or LCol > Columns.Count Then
        Err.Raise fxErr_InvalidArgument, , _
            StrFormat("'{}' is invalid Excel Column.", TargetColumn)
    End If

    If TargetSheet Is Nothing Then Set TargetSheet = ActiveSheet

    If TargetSheet.Cells(Rows.Count, LCol) <> "" Then
        GetLastRowInColumn = Rows.Count
    Else
        GetLastRowInColumn = TargetSheet.Cells(Rows.Count, LCol).End(xlUp).Row
    End If
    fx_Pop
Exit Function
Failed:
    If this.LastErr.Num = fxErr_None Then
        fx_SetErr Err.Number, Err.Description
    End If
    If EndOnError Then FatalError
    GetLastRowInColumn = -1
End Function

''' Returns file extension from file name
'''  Params:
'''   FName As String:
Public Function GetFileExt(FName As String) As String
    fx_Push "GetFileExt"
    Dim sExt As String

    On Error GoTo Failed
    sExt = Fso.GetExtensionName(FName)
    If sExt <> "" Then
        GetFileExt = "." & sExt
    End If
    fx_Pop
Exit Function
Failed:
    fx_SetErr Err.Number, Err.Description
End Function

''' Returns Folder Path from FullFileName
'''  Params:
'''   FullFileName As String:
Public Function GetFilePath(FullFileName As String) As String
    fx_Push "GetFilePath"
    Dim FName As String

    FName = DropFilePath(FullFileName)
    GetFilePath = Replace(FullFileName, FName, "")
    fx_Pop
End Function

''' Returns String with user's ShortDate Format
'''  Params:
'''   :
Public Function GetShortDateFormat() As String
    fx_Push "GetShortDateFormat"
    Dim sep As String, Res As String
    Dim d As String, m As String, Y As String

    With Application
        sep = .International(xlDateSeparator)
        d = .International(xlDayCode)
        If .International(xlDayLeadingZero) Then d = d + d
        m = .International(xlMonthCode)
        If .International(xlMonthLeadingZero) Then m = m + m
        Y = .International(xlYearCode) + .International(xlYearCode)
        If .International(xl4DigitYears) Then Y = Y + Y

        Select Case .International(xlDateOrder)
            Case 0
                Res = Join(Array(m, d, Y), sep)
            Case 1
                Res = Join(Array(d, m, Y), sep)
            Case 2
                Res = Join(Array(Y, m, d), sep)
        End Select
    End With

    GetShortDateFormat = Res
    fx_Pop
End Function

''' Finding a Workbook, opening if not opened and returning reference to it
Public Function GetWorkbook(ByVal WbkName As String, Optional ByVal WbkPath, Optional ByVal EndOnError = False, Optional ByVal ReadOnly = False, Optional ByVal PassWrd, Optional ByRef GetIsOpen) As Workbook
    fx_Push "GetWorkbook"
    Dim Wbk         As Workbook
    Dim bEvents     As Boolean
    Dim bAlerts     As Boolean
    Dim sMsg        As String

    With Application
        bAlerts = .DisplayAlerts
        bEvents = .EnableEvents
    End With

    ' If path is missing, then we assume, full filename is provided in WbkName
    If IsMissing(WbkPath) Then
        WbkPath = GetFilePath(WbkName)
        WbkName = DropFilePath(WbkName)
    End If

    ' checking if the folder path exists
    If Not FolderExists(WbkPath) Then
        Err.Raise fxErr_InvalidArgument, , "WbkPath: Folder not found"
    End If

    ' checking if the workbook exists
    WbkPath = TrailingSlash(WbkPath)
    If Not FileExists(WbkPath & WbkName) Then
        Err.Raise fxErr_InvalidArgument, , "WbkName: Workbook not found"
    End If

    ' If already opened - returning reference
    GetIsOpen = Workbook_IsOpen(WbkName)
    If GetIsOpen Then
        Set GetWorkbook = Workbooks(WbkName)
        fx_Pop
        Exit Function
    End If

    With Application
        .DisplayAlerts = False
        .EnableEvents = False
    End With

    ' attempting to open
    On Error GoTo Failed
    If Not IsMissing(PassWrd) Then
        Set Wbk = Workbooks.Open(FileName:=WbkPath & WbkName, UpdateLinks:=False, ReadOnly:=ReadOnly, Password:=PassWrd)
    Else
        Set Wbk = Workbooks.Open(FileName:=WbkPath & WbkName, UpdateLinks:=False, ReadOnly:=ReadOnly)
    End If

    If Wbk Is Nothing Then
        Err.Raise fxErr_Unknown, , "Failed to open " & WbkName
    End If

    With Application
        .DisplayAlerts = bAlerts
        .EnableEvents = bEvents
    End With

    Set GetWorkbook = Wbk
    fx_Pop
Exit Function
Failed:
    With Application
        .DisplayAlerts = bAlerts
        .EnableEvents = bEvents
    End With
    fx_SetErr Err.Number, Err.Description
    If EndOnError Then FatalError
End Function

''' Checking if worksheet exists and returning reference to it
'''  Params:
'''    SheetName > Sheet name to look for
'''    TargetWbk > The workbook to search in
'''    EndOnError:=False > if True, ending program execution
'''    UCaseCompare:=False > UCase before compare
'''    TrimCompare:=False > Trim before compare
'''    PartialMatch:=False > if True, using 'like' statement
'''                          to look for part of the sheet name
Public Function GetWorksheet(ByVal SheetName As String, ByRef TargetWbk As Workbook, Optional EndOnError = False, Optional UCaseCompare = False, Optional TrimCompare = False, Optional PartialMatch = False) As Worksheet
    fx_Push "GetWorksheet"
    Dim Wks         As Worksheet
    Dim bFound      As Boolean
    Dim sSearchName As String
    Dim sTargetName As String
    Dim sMsg        As String
    
    On Error GoTo Failed
    ' If TargetWbk argument is invalid
    If TargetWbk Is Nothing Then
        Err.Raise fxErr_ObjNotSet, , "TargetWbk: Object not set"
    ElseIf SheetName = vbNullString Then
        Err.Raise fxErr_InvalidArgument, , "SheetName: Param is Empty"
    End If

    ' Approximate match
    If UCaseCompare Or TrimCompare Or PartialMatch Then
        ' modifying search name
        sSearchName = SheetName
        If TrimCompare Then sSearchName = WorksheetFunction.Trim(sSearchName)
        If UCaseCompare Then sSearchName = UCase$(sSearchName)

        For Each Wks In TargetWbk.Worksheets
            ' modelling target name
            sTargetName = Wks.Name
            If TrimCompare Then sTargetName = WorksheetFunction.Trim(sTargetName)
            If UCaseCompare Then sTargetName = UCase$(sTargetName)

            ' comparing
            If PartialMatch Then
                bFound = (sTargetName Like "*" & sSearchName & "*")
            Else
                bFound = (sTargetName = sSearchName)
            End If

            If bFound Then
                Set GetWorksheet = Wks ' return
                fx_Pop
                Exit Function
            End If
        Next Wks
    ' Exact match
    Else
        On Error Resume Next
        Set Wks = TargetWbk.Worksheets(SheetName)
        On Error GoTo 0
    End If

    On Error GoTo Failed
    If Wks Is Nothing Then
        Err.Raise fxErr_Value, , _
            StrFormat("Failed to find sheet {} in {}", SheetName, TargetWbk.Name)
    End If

    Set GetWorksheet = Wks ' return
    fx_Pop
Exit Function
Failed:
    fx_SetErr Err.Number, Err.Description
    If EndOnError Then FatalError
End Function

''' Turn on everything
Public Sub GoLoud()
    fx_Push "GoLoud"
    With Application
        .StatusBar = False
        .ScreenUpdating = True
        .Calculation = this.xlCalcState
        If .Calculation <> xlCalculationAutomatic _
        And Application.CalculationState <> xlDone Then
            .Calculate
        End If
        .EnableEvents = True
        .DisplayAlerts = True
    End With
    fx_Pop
End Sub

''' Turn off everything
Public Sub GoSilent(Optional ByVal CalcState As XlCalculation = xlCalculationAutomatic)
    fx_Push "GoSilent"
    With Application
        .EnableEvents = False
        .DisplayAlerts = False
        .ScreenUpdating = False
        .Calculation = CalcState
    End With
    fx_Pop
End Sub

''' Returns length of Array-like object, Range or a Collection
'''  Params:
'''   Iterable As Variant
'''   Dimension=1: applicable only for Array and Range.
'''     - for Array: specifies the array's dimension
'''     - for Range: 0 = Range.Rows.Count, 1 = Range.Columns.Count
Public Function LenOf(Iterable As Variant, Optional Dimension = 1) As Long
    fx_Push "LenOf"
    On Error GoTo Failed
    Dim Result As Long
    Result = -1 ' Error flag

    If IsEmpty(Iterable) Then
        Result = 0
    ElseIf TypeName(Iterable) = "FxRange" Then
        Result = fx_LenOfFxRange(Iterable, Dimension)
    ElseIf TypeName(Iterable) = "Range" Then
        Result = fx_LenOfRange(Iterable, Dimension)
    Else
        On Error Resume Next
        Err.Clear
        Result = Iterable.Count
        If Err.Number = 0 Then GoTo Exit_Function ' does have Count property
    End If
            
    On Error GoTo Failed
        
    If VarType(Iterable) = vbString Then
        Result = VBA.Len(Iterable)
    ElseIf IsArray(Iterable) Then
        Result = UBound(Iterable, Dimension) - LBound(Iterable, Dimension) + 1
    Else
        Raise_InvalidArgumentError "Unsupported data type for 'Iterable'"
    End If
Exit_Function:
    LenOf = Result
    fx_Pop
Exit Function
Failed:
    Result = -1
    fx_SetErr fxErr_System, Err.Description
End Function

''' Converting ListObject to Range and returning reference to it
'''  Params:
'''   ListObj As ListObject,
'''   Optional ClearFormatting As Boolean = True:
Public Function ListObject_ToRange(ListObj As ListObject, Optional ClearFormatting = True) As Range
    fx_Push "ListObject_ToRange"
    Dim Rng As Range

    On Error GoTo Failed
    If ListObj Is Nothing Then
        Err.Raise fxErr_ObjNotSet, "ListObj: object not set"
    End If

    With ListObj
        Set Rng = .Range
        .Unlist ' convert back to range
    End With

    If ClearFormatting Then
        With Rng
            .Interior.ColorIndex = xlColorIndexNone
            .Font.ColorIndex = xlColorIndexAutomatic
            .Borders.LineStyle = XlLineStyleNone
        End With
    End If

    Set ListObject_ToRange = Rng
    fx_Pop
Exit Function
Failed:
    fx_SetErr Err.Number, Err.Description
End Function

''' Replacing all invalid characters in a FileName
'''  Params:
'''   ByVal FileName As String,
'''   Optional ByVal ReplaceWith As String = "_":
Public Function NormalizeFileName(ByVal FileName As String, Optional ByVal ReplaceWith = "_") As String
    fx_Push "NormalizeFileName"
    Dim v       As Variant
    Dim Invalid As Variant

    On Error GoTo Failed
    Invalid = Array("/", "\", ":", "*", "?", Chr$(34), ">", "<", "|")
    If Arr_ContainsStr(Invalid, ReplaceWith) Then ReplaceWith = "_"

    For Each v In Invalid
        FileName = Replace(FileName, v, ReplaceWith)
    Next v
    NormalizeFileName = FileName
    fx_Pop
Exit Function
Failed:
    fx_SetErr Err.Number, Err.Description
End Function

''' Returns FullFileName from list of folder Paths and FileName
'''  Params:
'''   Path1_PathN_FileName As ParamArray:
Public Function PathJoin(ParamArray Path1_PathN_FileName() As Variant) As String
    fx_Push "PathJoin"
    Dim i As Long

    On Error GoTo Failed
    For i = LBound(Path1_PathN_FileName) To UBound(Path1_PathN_FileName)
        If Right(Path1_PathN_FileName(i), 1) = Application.PathSeparator Then
            Path1_PathN_FileName(i) = Left(Path1_PathN_FileName(i), Len(Path1_PathN_FileName(i)) - 1)
        End If
    Next i

    PathJoin = Join(Path1_PathN_FileName, Application.PathSeparator)
    fx_Pop
Exit Function
Failed:
    fx_SetErr Err.Number, Err.Description
End Function

Public Sub PrintDebug(ParamArray Msg() As Variant)
    Debug.Print "[DEBUG] "; Join(Msg, " ")
End Sub

Public Sub PrintInfo(ParamArray Msg() As Variant)
    Debug.Print "[INFO] "; Join(Msg, " ")
End Sub

Public Sub PrintWarning(ParamArray Msg() As Variant)
    Debug.Print "[WARNING] "; Join(Msg, " ")
End Sub

Public Sub PrintError(ParamArray Msg() As Variant)
    Debug.Print "[ERROR] "; Join(Msg, " ")
End Sub

Public Sub Raise_NotImplementedError(Optional Description = "", Optional Source = "")
    Err.Raise fxErr_NotImplemented, Source, _
        Me.Error(fxErr_NotImplemented) & vbNewLine & Description
End Sub

Public Sub Raise_InvalidArgumentError(Optional Description = "", Optional Source = "")
    Err.Raise fxErr_InvalidArgument, Source, _
        Me.Error(fxErr_InvalidArgument) & vbNewLine & Description
End Sub

Public Sub Raise_ObjectNotSetError(Optional Description = "", Optional Source = "")
    Err.Raise fxErr_ObjNotSet, Source, _
        Me.Error(fxErr_ObjNotSet) & vbNewLine & Description
End Sub

''' Comparing all values in two ranges
Public Function Range_Compare(ByRef RngFirst As Range, ByRef RngSecond As Range, Optional ByVal UCaseCompare = False, Optional ByVal TrimCompare = False, Optional ByVal HCompare = True, Optional ByVal EndOnError = False) As Boolean
    fx_Push "Range_Compare"
    Dim i    As Long, sMsg  As String
    Dim r1   As Range, r2   As Range
    Dim v1   As Variant, v2 As Variant

    For Each v1 In Array(RngFirst, RngSecond)
        If v1 Is Nothing Then Err.Raise fxErr_ObjNotSet, , Error(fxErr_ObjNotSet)
        GoTo Failed
    Next v1

    If HCompare Then
        Set r1 = RngFirst.Rows(1)
        Set r2 = RngSecond.Rows(1)
    Else
        Set r1 = RngFirst.Columns(1)
        Set r2 = RngSecond.Columns(1)
    End If

    If r1.Cells.Count <> r2.Cells.Count Then
        Range_Compare = False
    fx_Pop
        Exit Function
    End If

    For i = 1 To r1.Cells.Count
        If HCompare Then ' looping horizontally
            v1 = r1.Cells(1, i).Value
            v2 = r2.Cells(1, i).Value
        Else ' looping vertically
            v1 = r1.Cells(i, 1).Value
            v2 = r2.Cells(i, 1).Value
        End If

        If TrimCompare Then
            v1 = Trim(v1)
            v2 = Trim(v2)
        End If
        If UCaseCompare Then
            v1 = UCase(v1)
            v2 = UCase(v2)
        End If

        If v1 <> v2 Then
            Range_Compare = False
            fx_Pop
            Exit Function
        End If
    Next i

    Range_Compare = True
    fx_Pop
Exit Function
Failed:
    fx_SetErr Err.Number, Err.Description
    If EndOnError Then FatalError
End Function

''' Reading Form-like Range to Dictionary (2 columns with label-value)
''' First column is used as Dict keys. Second column is read cells containing the value
Public Function Range_FormToDict(Rng As Range) As Variant
    fx_Push "Range_FormToDict"
    Dim Result  As Object
    Dim i       As Long

    If Rng Is Nothing Then Err.Raise fxErr_ObjNotSet, , _
        "Argument 'Rng' is not set"
    If Rng.Columns.Count <> 2 Then Err.Raise fxErr_InvalidArgument, , _
        "Invalid number of columns for argument 'Rng'. Expected 2"

    Set Result = CreateObject("Scripting.Dictionary")
    For i = 1 To Rng.Rows.Count
        Set Result(Rng.Cells(i, 1).Value) = Rng.Cells(i, 2)
    Next i

    Set Range_FormToDict = Result
    fx_Pop
Exit Function
Failed:
    fx_SetErr Err.Number, Err.Description
End Function

''' Finds Cell by Text and maps the Range associated with it
Public Function Range_GetByHeader(ByVal HeaderTxt As String, ByRef Wks As Worksheet, Optional ByVal EndOnError = False, Optional ByRef Dir As fxDirection = fxNoDir, Optional ByVal EntireCell As XlLookAt = xlWhole, Optional ByVal CaseSensitive = False) As Range
    fx_Push "Range_GetByHeader"
    Dim R     As Range
    Dim sMsg  As String, LR As Long, lC As Long
    Dim v     As Variant, vDirs As Variant
    On Error GoTo Failed
    If Wks Is Nothing Then Err.Raise fxErr_ObjNotSet, , "Argument 'Wks' not set"

    On Error Resume Next
    Set R = Wks.Cells.Find(what:=HeaderTxt, LookAt:=EntireCell, MatchCase:=CaseSensitive)
    If R Is Nothing Then
        fx_SetErr fxErr_Value, "Can't find cell " & HeaderTxt & " in sheet " & Wks.Name
        GoTo Failed
    End If

    On Error GoTo Failed
    If Dir <> fxNoDir Then
        vDirs = Array(Array(1, 0, xlDown, fxDown), _
                        Array(0, 1, xlToRight, fxRight), _
                        Array(0, -1, xlToLeft, fxLeft), _
                        Array(0, -1, xlUp, fxUp))
newDir:
        Select Case Dir
            Case fxDown And R.Row <> Rows.Count: v = vDirs(0)
            Case fxRight And R.Column <> Columns.Count: v = vDirs(1)
            Case fxLeft And R.Column <> 1: v = vDirs(2)
            Case fxUp And R.Row <> 1: v = vDirs(3)
            Case fxAutoDir
                For Each v In vDirs
                    If R.Offset(v(0), v(1)) <> "" Then
                        Dir = v(3)
                        GoTo newDir
                    End If
                Next v
                v = Array(0, 0, xlDown, fxNoDir)
                Err.Raise fxErr_FxInternal, , "fxAutoDir failed to find fxDirection value"
                If EndOnError Then GoTo Failed
            Case Else
                Err.Raise fxErr_InvalidArgument, , "fxDirection"
        End Select

        If v(3) <> fxNoDir And R.Offset(v(0), v(1)) <> "" Then
            Set R = Range(R.Offset(v(0), v(1)), R.End(v(2)))
        End If
    End If

    Set Range_GetByHeader = R
    fx_Pop
Exit Function
Failed:
    fx_SetErr Err.Number, Err.Description
    If EndOnError Then FatalError
End Function

''' Searching for list of values in the first row / column of a range
''' Returns True if all values are found
''' Use fx.LastErrText to read comma delimited string with missing values
'''  Params:
'''   Rng As Range,
'''   List As Variant,
'''   Optional ByVal LookIn As xlRowCol:= xlRows
Public Function Range_Isin(Rng As Range, List As Variant, Optional ByVal LookIn As xlRowCol = xlRows) As Boolean
    fx_Push "Range_Isin"
    Dim Rng_ As Range, v As Variant
    Dim bFound As Boolean
    Dim Missing As New Collection

    If LookIn = xlRows Then
        Set Rng_ = Rng.Rows(1)
    Else
        Set Rng_ = Rng.Columns(1)
    End If

    For Each v In List
        On Error Resume Next
        bFound = False
        bFound = (WorksheetFunction.Match(v, Rng_, 0) > 0)
        On Error GoTo 0
        If Not bFound Then Missing.Add CStr(v)
    Next v

    If Missing.Count Then
        Range_Isin = False
        this.LastErr.Text = Collection_ToString(Missing)
    Else
        Range_Isin = True
    End If
    fx_Pop
End Function

Public Function Range_ToArray1D(Rng As Range) As Variant
    fx_Push "Range_ToArray1D"
    Dim Arr      As Variant
    Dim Cell     As Range
    Dim i        As Long

    ReDim Arr(0 To Rng.Cells.Count - 1)
    
    For Each Cell In Rng.Cells
        Arr(i) = Cell.Value
        i = i + 1
    Next Cell

    Range_ToArray1D = Arr
    fx_Pop
Exit Function
Failed:
    fx_SetErr Err.Number, Err.Description
End Function

Public Function Range_ToArray2D(Rng As Range) As Variant
    fx_Push "Range_ToArray2D"
    Dim Arr      As Variant

    If Rng.Cells.Count = 1 Then
        ReDim Arr(1 To 1, 1 To 1)
        Arr(1, 1) = Rng.Value
    Else
        Arr = Rng
    End If

    Range_ToArray2D = Arr
    fx_Pop
Exit Function
Failed:
    fx_SetErr Err.Number, Err.Description
End Function

''' Returns delimited Sting from values in a Range
'''  Params:
'''   ByVal Rng As Range,
'''   Delim As String:
Public Function Range_ToString(ByVal Rng As Range, Delim As String) As String
    fx_Push "Range_ToString"
    Dim sTxt As String
    Dim R As Range

    On Error GoTo Failed
    If Rng Is Nothing Then
        Err.Raise fxErr_ObjNotSet, , "Rng: Object not set"
    End If

    For Each R In Rng.Cells
        sTxt = sTxt & CStr(R) & Delim
    Next R

    sTxt = Left(sTxt, Len(sTxt) - Len(Delim))
    Range_ToString = sTxt
    fx_Pop
Exit Function
Failed:
    fx_SetErr Err.Number, Err.Description
End Function

''' Mapping Table-like Range to a Scripting.Dictionary with Header:Data (String:Range) pair for each column
''' Adding Keys for FULL_RNG, HEAD_RNG,DATA_RNG
''' Params:
''' Rng As Range
''' EndOnError:=False
''' IgnoreEmptyHeaders:=True > when True, Assigning 'virtual header' for each Empty header cell,
'''                            but still the function is raising error message to fx.LastErrText
Public Function Range_ToTable(Rng As Range, Optional ByVal EndOnError = False, Optional ByVal IgnoreEmptyHeaders = True) As Object ' Scripting.Dictionary
    fx_Push "Range_ToTable"
    Const csEMPTY_H = "fx_Empty_Header"
    On Error GoTo Failed
    Dim Result As Object
    Dim lRows  As Long, lColumns As Long, i As Long, j As Long
    Dim sMsg   As String, sKey   As String, sDupl As String
    Dim R      As Range

    'Check Rng Input
    If Rng Is Nothing Then
        Raise_ObjectNotSetError "Rng: Object not set"
    ElseIf Rng.Rows.Count < 2 Then
        Raise_InvalidArgumentError "Rng: Rng.Rows.Count = 1"
    ElseIf CBool(WorksheetFunction.CountBlank(Rng.Rows(1))) Then
        If Not IgnoreEmptyHeaders Then
            Raise_InvalidArgumentError "Rng: Missing or Empty headers"
        Else
            fx_SetErr fxErr_InvalidArgument, "Rng: Missing or Empty headers"
        End If
    ElseIf Rng.Areas.Count > 1 Then
        Raise_InvalidArgumentError "Rng: Cannot use multi-area range"
    End If

    lRows = IIf(Rng.Rows.Count = 1, 2, Rng.Rows.Count)
    lColumns = Rng.Columns.Count

    ' mapping useful ranges
    Set Result = CreateObject("Scripting.Dictionary")
    Result.Add this.FRNG, Rng
    Result.Add this.HRNG, Rng.Rows(1)
    Result.Add this.DRNG, Range(Rng.Rows(2), Rng.Rows(lRows))

    ' mapping all columns - Header : DataRange pairs
    j = 1
    For i = 1 To lColumns
        sKey = CStr(Rng.Cells(1, i))
        If sKey = vbNullString Then
            sKey = csEMPTY_H & CStr(j) ' assigning virtual header
            j = j + 1
        End If

        If Result.Exists(sKey) Then
            sDupl = sDupl & IIf(sDupl = "", "", ", ") & sKey
        Else
            Result.Add Key:=sKey, _
                        Item:=Range(Rng.Cells(2, i), Rng.Cells(lRows, i))
        End If
    Next i

    If sDupl <> "" Then
        Err.Raise fxErr_Value, , "Duplicate columns: " & sDupl
    End If

    Set Range_ToTable = Result
    fx_Pop
Exit Function
Failed:
    fx_SetErr Err.Number, Err.Description
    If EndOnError Then FatalError
End Function

Sub Range_TrimInPlace(Rng As Range)
    fx_Push "Range_TrimInPlace"
    On Error GoTo Failed
    Dim i       As Long
    Dim Arr     As Variant
    Dim Area    As Range

    For Each Area In Rng.Areas
        Arr = Range_ToArray2D(Area)

        For i = LBound(Arr, 1) To UBound(Arr, 1)
            Arr(i, 1) = Trim(Arr(i, 1))
        Next

        Area = Arr
    Next
    fx_Pop
Exit Sub
Failed:
    fx_SetErr fxErr_System, Err.Description
End Sub

''' Much like Application.Union, but any of the range arguments can be 'Nothing'
'''  Params:
'''   ParamArray Rng1_Rng_2_RngN - param list of ranges to Union
Public Function Range_Union(ParamArray Rng1_Rng_2_RngN() As Variant) As Range
Attribute Range_Union.VB_Description = "Line Application.Union, but any of the range arguments can be 'Nothing'  \nParams:\n ParamArray Rng1_Rng_2_RngN - list of ranges to union"
    fx_Push "Range_Union"
    Dim Result  As Range
    Dim i       As Long

    On Error GoTo Failed
    Set Result = Rng1_Rng_2_RngN(0)
    For i = 1 To UBound(Rng1_Rng_2_RngN)
        If Result Is Nothing Then
            Set Result = Rng1_Rng_2_RngN(i)
        ElseIf Not Rng1_Rng_2_RngN(i) Is Nothing Then
            Set Result = Application.Union(Result, Rng1_Rng_2_RngN(i))
        End If
    Next i

    Set Range_Union = Result
    fx_Pop
Exit Function
Failed:
    fx_SetErr Err.Number, Err.Description

End Function


Public Sub Reset()
    Me.ClearLastError
    Set this.Fso = Nothing
    Application.Calculation = this.xlCalcState ' resuming Excel's calculation state
End Sub

Public Function SheetAddress(Rng As Range, Optional AbsRow = True, Optional AbsCol = True) As String
    fx_Push "SheetAddress"
    On Error GoTo Failed
    SheetAddress = VBA.Replace(Rng.Address(AbsRow, AbsCol, external:=True), _
        "[" & Rng.Parent.Parent.Name & "]", "")
    fx_Pop
Exit Function
Failed:
    fx_SetErr fxErr_System, Err.Description
End Function

Public Function StrFormat(ByVal Str As String, ParamArray Args() As Variant) As String
    fx_Push "StrFormat"
    Dim i               As Long
    Dim BeforeReplace   As String

    On Error GoTo Failed
    ' Validating Input
    If IsMissing(Args) Then
        If Not fx_ContainsFormatSeq(Str) Then
            StrFormat = Str
            Exit Function
        Else
            Err.Raise fxErr_InvalidArgument, , _
                "Missing arguments for the format function"
        End If
    ElseIf Not IsMissing(Args) And Not fx_ContainsFormatSeq(Str) Then
        Err.Raise fxErr_InvalidArgument, , _
            "'Str' does not contain format identifiers '{}'"
    End If

    i = 0
    Do While VBA.InStr(1, Str, "{}") > 0
        Str = VBA.Replace(Str, "{}", "{" & i & "}", Count:=1)
        i = i + 1
    Loop

    For i = 0 To UBound(Args)
        BeforeReplace = Str
        Str = VBA.Replace(Str, "{" & i & "}", CStr(Args(i)))

        If Str = BeforeReplace Then
            Err.Raise fxErr_InvalidArgument, , _
                "'Str' does not contain format identifier '{" & i & "}'"
        End If
    Next i

    If fx_ContainsFormatSeq(Str) Then
        Err.Raise fxErr_InvalidArgument, , _
            "Not enough values in 'Args' to format 'Str'"
    End If

    StrFormat = Str
    fx_Pop
Exit Function
Failed:
    fx_SetErr fxErr_System, Err.Description
End Function

''' Wraps String with Left and/or Right Strings
'''  Params:
'''   Val As String, LeftStr As String, Optional Right Str As Variant:
Public Function StrWrap(val As String, LeftStr As String, Optional RightStr) As String
    fx_Push "StrWrap"
    Dim RightVal As String

    On Error GoTo Failed
    RightVal = IIf(IsMissing(RightStr), LeftStr, RightStr)
    StrWrap = LeftStr & val & RightStr
    fx_Pop
Exit Function
Failed:
    fx_SetErr Err.Number, Err.Description
End Function

''' Adding new columns to existing Scripting.Dictionary (see fx.Range_ToTable)
'''  Params:
'''     TblObj        > must have keys for FULL_RNG, HEAD_RNG, DATA_RNG
'''     ColumnNames       > can be String or String array
'''     EndOnError:=False > if True, terminate the program execution
Public Sub Table_AppendColumns(TblObj As Object, ColumnNames As Variant, Optional ByVal EndOnError = False)
    fx_Push "Table_AppendColumns"
    Const csDelim = "_::_"
    Dim lAdded As Long, lRows As Long, lColumns As Long
    Dim rData  As Range, rHeader As Range, R As Range
    Dim Duplicates As New Collection
    Dim v As Variant, vIn As Variant

    If TblObj Is Nothing Then
        Err.Raise fxErr_ObjNotSet, , "TblObj: Object not set"
    End If

    ' checking input
    Select Case TypeName(ColumnNames)
        Case "String"
            vIn = Array(ColumnNames)
        Case "Range"
            Set R = ColumnNames
            If WorksheetFunction.CountBlank(R) Then
                Err.Raise fxErr_InvalidArgument, , _
                    "ColumnNames: Missing or Empty Cells"
            End If

            vIn = Range_ToString(R, csDelim)
            vIn = Split(vIn, csDelim)
        Case Else
            vIn = ColumnNames
    End Select

    ' attempting to parse it as valid Sting
    On Error Resume Next
    v = Join(vIn, csDelim)
    On Error GoTo 0
    If v = Empty Then
        Err.Raise fxErr_InvalidArgument, , "ColumnNames: Convert to String Array failed"
    ElseIf Len(v) = 0 Then
        Err.Raise fxErr_InvalidArgument, , "ColumnNames: Argument is empty"
    End If

    ' get last header and data rng in TblObj
    Set rHeader = TblObj(this.HRNG).Cells(1, lColumns)
    Set rData = TblObj(this.DRNG).Columns(lColumns)

    ' mapping new columns, adding header names
    For Each v In vIn
        If TblObj.Exists(v) Then Duplicates.Add CStr(v)
    Next v
    If Duplicates.Count Then
        Err.Raise fxErr_FxInternal, , "Duplicate columns not allowed" _
            & vbNewLine & Collection_ToString(Duplicates)
    End If

    For Each v In vIn
        ' offset first
        Set rHeader = rHeader.Offset(0, 1)
        Set rData = rData.Offset(0, 1)

        ' adding new data
        rHeader = v             ' header name
        TblObj.Add v, rData ' new dict item
        lAdded = lAdded + 1     ' Count added
    Next v

    ' remapping to include the new columns
    Set R = TblObj(this.FRNG).Resize(ColumnSize:=lColumns + lAdded)
    Set TblObj(this.FRNG) = R
    Set TblObj(this.HRNG) = R.Rows(1)
    Set TblObj(this.DRNG) = Range(R.Rows(2), R.Rows(lRows))

    fx_Pop
Exit Sub
Failed:
    fx_SetErr Err.Number, Err.Description
    If EndOnError Then FatalError
End Sub

''' Adding new rows to existing Scripting.Dictionary (see fx.Range_ToTable)
'''  Params:
'''     TblObj        > must have keys for FULL_RNG, HEAD_RNG, DATA_RNG
'''     RowsRng       > Range to append to the bottom of the current table
Public Sub Table_AppendRows(TblObj As Object, RowsRng As Object)
    fx_Push "Table_AppendRows"
    Dim v As Variant
    Dim rHeader As Range, R As Range

    On Error GoTo Failed
    If TblObj Is Nothing Then
        Err.Raise fxErr_ObjNotSet, , "TblObj: Object not set"
    End If

    ' checking input
    If RowsRng.Columns.Count <> TblObj(DATA_RNG).Columns.Count Then
        Err.Raise fxErr_InvalidArgument, , "Columns Count mismatch"
    End If

    ' Updating Special Ranges
    For Each v In Array(this.FRNG, this.DRNG, this.KEY_COL, this.TTL_COL)
        If TblObj.Exists(v) Then

        End If
    Next v

    ' Updating regular columns
    For Each v In Table_DropHelperKeysFrom(TblObj)

    Next v

    ' Moving TTl Row if any
    If TblObj.Exists(TTL_ROW) Then

    End If
    fx_Pop
Exit Sub
Failed:
    fx_SetErr Err.Number, Err.Description
End Sub

''' Returns only ColumnName Keys in Tbl-Like Object
Public Function Table_DropHelperKeysFrom(TblObj As Object) As Variant
    fx_Push "Table_DropHelperKeysFrom"
    On Error GoTo Failed
    Dim v As Variant
    Dim Result As New Collection

    For Each v In TblObj.Keys()
        If Not Arr_ContainsStr(Table_HelperKeys, CStr(v)) Then
            Result.Add CStr(v)
        End If
    Next v

    If Result.Count Then
        Table_DropHelperKeysFrom = Split(Collection_ToString(Result, Delim:=","), ",")
    Else
        Err.Raise fxErr_Value, , "TblObj does not contain ColumnName Keys"
    End If
    fx_Pop
Exit Function
Failed:
    fx_SetErr Err.Number, Err.Description
End Function

''' Returns only the special range Keys in Tbl-Like Object
Public Function Table_GetHelperKeysFrom(TblObj As Object) As Variant
    fx_Push "Table_GetHelperKeysFrom"
    On Error GoTo Failed
    Dim v As Variant
    Dim Result As New Collection

    For Each v In TblObj.Keys()
        If Arr_ContainsStr(Table_HelperKeys, CStr(v)) Then
            Result.Add CStr(v)
        End If
    Next v

    If Result.Count Then
        Table_GetHelperKeysFrom = Split(Collection_ToString(Result, Delim:=","), ",")
    Else
        Err.Raise fxErr_Value, , "TblObj does not contain Helper Keys"
    End If

    Set Table_GetHelperKeysFrom = Result
    fx_Pop
Exit Function
Failed:
    fx_SetErr Err.Number, Err.Description
End Function

''' Extended Version, including all constants
Public Function Table_HelperKeys() As Variant
    fx_Push "Table_HelperKeys"
    Table_HelperKeys = Array(this.FRNG, this.HRNG, this.DRNG, this.KEY_COL, this.TTL_COL, this.TTL_ROW)
    fx_Pop
End Function

''' Mapping a Row from Table-like Range to a Scripting.Dictionary
''' Params:
''' TblObj: Scripting.Dictionary
''' RowNumber: Long
Public Function Table_iAt(TblObj As Object, RowNumber As Long) As Object ' Scripting.Dictionary
    fx_Push "Table_iAt"
    Dim Key As Variant
    Dim Keys As Variant
    Dim TRow As Object

    Set TRow = CreateObject("Scripting.Dictionary")
    Keys = Table_DropHelperKeysFrom(TblObj)
    For Each Key In Keys
        TRow.Add CStr(Key), TblObj(Key).Rows(RowNumber)
    Next

    TRow.Add FULL_RNG, Union(TblObj(HEAD_RNG), TblObj(DATA_RNG).Rows(RowNumber))
    TRow.Add HEAD_RNG, TblObj(HEAD_RNG)
    TRow.Add DATA_RNG, TblObj(DATA_RNG).Rows(RowNumber)
    Set Table_iAt = TRow
    fx_Pop
End Function

''' Adding Path Separator at the end of a PathName
'''  Params:
'''   PathName As String:
Public Function TrailingSlash(ByVal PathName As String) As String
    fx_Push "TrailingSlash"
    With Application
        TrailingSlash = PathName & (IIf(Right(PathName, 1) = .PathSeparator, "", .PathSeparator))
    End With
    fx_Pop
End Function

''' Pauses the program execution
Public Sub WaitTime(MilliSeconds As Double)
    fx_Push "WaitTime"
    Dim dEndTime As Double
    dEndTime = Timer + MilliSeconds / 1000#
    Do While Timer < dEndTime
        DoEvents
    Loop
    fx_Pop
End Sub

''' Returns True if a Workbook is open in this instance of Excel
Public Function Workbook_IsOpen(ByVal WbkName As String) As Boolean
    fx_Push "Workbook_IsOpen"
    Dim Wbk As Workbook

    On Error Resume Next
    Set Wbk = Workbooks(WbkName)
    Workbook_IsOpen = (Not Wbk Is Nothing)
    fx_Pop
End Function

'::region Properties

''' Returns fxErrCode number if the last called function exited with error
''' Returns 0 (fxErr_None) if no error was raised
Public Property Get LastErrNumber() As Long
    LastErrNumber = CLng(this.LastErr.Num)
End Property

''' Returns String with the Error Message raised by the last fx function call
''' Returns Empty if no error was raised
Public Property Get LastErrText() As String
    LastErrText = this.LastErr.Text
End Property

''' If True, FatalError will call Err.Raise, instead of terminating"
Public Property Get DEBUG_MODE() As Boolean
    DEBUG_MODE = this.DEBUG_MODE
End Property

Public Property Get Name() As String
    Name = this.Name
End Property

Public Property Get Version() As Double
    Version = this.Version
End Property

Public Property Get VersionDate() As Date
    VersionDate = this.VersionDate
End Property

Public Property Get FULL_RNG() As String
    FULL_RNG = this.FRNG
End Property

Public Property Get DATA_RNG() As String
    DATA_RNG = this.DRNG
End Property

Public Property Get HEAD_RNG() As String
    HEAD_RNG = this.HRNG
End Property

Public Property Get TTL_ROW() As String
    TTL_ROW = this.TTL_ROW
End Property

Public Property Get TTL_COL() As String
    TTL_COL = this.TTL_COL
End Property

Public Property Get KEY_COL() As String
    KEY_COL = this.KEY_COL
End Property

''' Returns Scripting.FileSystemObject for file IO
Public Property Get Fso() As Object ' Scripting.FileSystemObject
    If this.Fso Is Nothing Then
        Set this.Fso = CreateObject("Scripting.FileSystemObject")
    End If
    Set Fso = this.Fso
End Property

''' Storing Excel's calculation state at the time of fx class initialization
Public Property Get CalcState() As XlCalculation
    CalcState = this.xlCalcState
End Property

Public Property Get Error(ErrorNumber As Long) As String
    Dim Result As String
    Select Case ErrorNumber
        Case fxErr_System
            Result = "System error"
        Case fxErr_InvalidArgument
            Result = "Invalid function argument"
        Case fxErr_ObjNotSet
            Result = "Object not set error"
        Case fxErr_FxInternal
            Result = "Fx internal error"
        Case fxErr_Unknown
            Result = "Unknown error"
        Case fxErr_NotImplemented
            Result = "Function or Sub not implemented"
        Case fxErr_Fatal
            Result = "Fatal error"
        Case Else
            Result = VBA.Error(ErrorNumber)
    End Select
    Error = Result
End Property

'::region Private Methods
Private Sub fx_Push(ProcName As String)
    With this.LastErr
        .ProcNames.Add ProcName
        .Num = fxErr_None
        .Text = vbNullString
    End With
End Sub

Private Sub fx_Pop()
    With this.LastErr.ProcNames
        .Remove .Count
    End With
End Sub

Private Sub fx_Validate_TwoArrayInput(ArrFirst As Variant, ArrSecond As Variant)

    If Not IsArray(ArrFirst) Then
        Raise_InvalidArgumentError "ArrFirst is not an Array"
    End If

    If Not IsArray(ArrSecond) Then
        Raise_InvalidArgumentError "ArrSecond is not an Array"
    End If

    If Me.Arr_IsEmpty(ArrFirst) Then
        Raise_InvalidArgumentError "ArrFirst is Empty"
    End If

    If Arr_IsEmpty(ArrSecond) Then
        Raise_InvalidArgumentError "ArrSecond is Empty"
    End If
End Sub

Private Function fx_LenOfRange(ByVal Rng As Range, ByVal Dimension As Long) As Long
    Dim Result  As Long
    Dim Area    As Range
    fx_ValidateRngDimension Dimension

    For Each Area In Rng.Areas
        If Dimension = 1 Then
            Result = Result + Area.Rows.Count
        Else
            Result = Result + Area.Columns.Count
        End If
    Next

    fx_LenOfRange = Result
End Function

Private Function fx_LenOfFxRange(ByVal FxRng As Object, ByVal Dimension As Long) As Long
    Dim Result  As Long
    fx_ValidateRngDimension Dimension

    If Dimension = 1 Then
        Result = FxRng.RowsCount
    Else
        Result = FxRng.ColumnsCount
    End If

    fx_LenOfFxRange = Result
End Function

Private Sub fx_ValidateRngDimension(ByVal Dimension As Long)
    If Dimension < 1 Or Dimension > 2 Then
        Raise_InvalidArgumentError "Incorrect value for 'Dimension'. Expected 1 = Range.Rows.Count, 2 = Range.Columns.Count"
    End If
End Sub

Private Function fx_ContainsFormatSeq(Str As String) As Boolean
    fx_ContainsFormatSeq = (Str Like "*{#}*") Or (Str Like "*{}*")
End Function

''' (Private) Helper to set fxErrCode and Error Description
Private Sub fx_SetErr(ByVal ErrNumber As Long, Optional ByVal Description = "")
    Dim Source As String
    Description = (IIf(Description = "", Me.Error(ErrNumber), Description))

    With this.LastErr
        .Num = ErrNumber
        Source = this.Name & "." & Me.Collection_ToString(.ProcNames, " > ")
    End With

    this.LastErr.Text = VBA.Join(Array( _
            "Source: " & Source, _
            "Description: " & Description), vbNewLine)
End Sub

''' (Private) Helper for Browse. Adding File Filters to FileDialog
Private Sub fx_AddFileFilters(FilterTyp As fxFileFilter, Dialog As FileDialog)
    Dim arrFltr As Variant
    Dim lMax    As Long, lMin As Long, lVal As Long

    lVal = FilterTyp
    lMin = fxFileFilter.[_First]
    lMax = fxFileFilter.[_Last] - 1

    Do While lMax > lMin And lVal > lMin
        Do While lVal >= lMax
            lVal = lVal - lMax
            ' Description, Expression pair
            Select Case lMax
                Case fxAllFiles
                    arrFltr = Array("All Files", "*.*")
                Case fxXlFiles
                    arrFltr = Array("Excel Files", "*.xl*")
                Case fxCSVFiles
                    arrFltr = Array("CSV Files", "*.csv")
                Case fxTxtFiles
                    arrFltr = Array("Text Files", "*.txt")
                Case fxPdfFiles
                    arrFltr = Array("Pdf Files", "*.pdf")
                Case fxWordFiles
                    arrFltr = Array("Word Files", "*.doc*")
                Case fxImagesFiles
                    arrFltr = Array("Images", "*.gif", "*.jpg", "*.jpeg")
            End Select

            Dialog.Filters.Add arrFltr(0), arrFltr(1), 1 ' push to the top
        Loop
        lMax = lMax / 2
    Loop
End Sub

''' (Private) Helper to convert fxRefTyp to String
'''  Params:
'''   Ref As fxRefTyp, Val As Variant:
Private Function fx_RefToStr(Ref As fxRefTyp, val As Variant) As String
    Dim Result As String
    Select Case Ref
        Case fxAt
            Result = StrWrap(Replace(CStr(val), "'", "''"), "[@[", "]]")
        Case FxRng
            Result = val.Address(False, True, , True)
        Case fxNum, fxRef
            Result = CStr(val)
        Case fxStr
            Result = StrWrap(CStr(val), Chr(34))
    End Select
    fx_RefToStr = Result ' returns
End Function

''' (Private) Helper to convert Column in Table-like object to Range Address
'''  Params:
'''   TblObj As Object, ColName As String:
Private Function fx_TblColToAddr(TblObj As Object, ByVal ColName As String) As String
    Dim Result As String

    ' Checking input
    On Error GoTo Failed
    Select Case TypeName(TblObj)
        Case "ListObject"
            Result = Replace(TblObj.Name & StrWrap(TblObj.ListColumns(ColName).Name, "[", "]"), "'", "''")
        Case "Dictionary"
            Result = TblObj(ColName).Address(True, True, xlA1, True)
        Case "clsSmartRange", "fxRng"
            Result = TblObj.DataColumn(ColName).Address(True, True, xlA1, True)
        Case Else
            Err.Raise fxErr_InvalidArgument, , "Invalid type for TblObj"
    End Select
    fx_TblColToAddr = Result
Exit Function
Failed:
    fx_SetErr Err.Number, Err.Description
End Function
