Attribute VB_Name = "FxReflection"
'---------------------------------------------------------------------------------------
' Purpose   :       Prints all subs and functions in a project
' Prerequisites:    Microsoft Visual Basic for Applications Extensibility 5.3 library
'                   CreateLogFile
' How to run:       Run GetFunctionAndSubNames, set a parameter to blnWithParentInfo
'                   If ComponentTypeToString(vbext_ct_StdModule) = "Code Module" Then
'
' Used:             ComponentTypeToString from -> http://www.cpearson.com/excel/vbe.aspx
'---------------------------------------------------------------------------------------

Option Explicit

Public Type TProc
    Name As String
    ModuleName As String
    Type As String
End Type

Public Type TModule
    Name            As String
    Procedures()    As TProc
    ProceduresCount As Long
    Type            As String
End Type

Public Type TProject
    Name         As String
    Modules()    As TModule
    ModulesCount As Long
End Type

Public Sub PrintWorkbookProcedures(WorkbookName As String)
    PrintProjectProcedures Workbooks(WorkbookName).VBProject
End Sub


Public Sub PrintProjectProcedures(VBProj As VBProject)
    Dim Proj            As TProject
    Dim i               As Long
    
    Proj = ReadProject(VBProj)
    If Proj.ModulesCount = 0 Then Exit Sub
    
    For i = 0 To Proj.ModulesCount - 1
        Debug.Print "Module:", Proj.Modules(i).Name
        PrintModuleProcedures Proj.Modules(i).Name, True
    Next
    
End Sub

Public Sub PrintModuleProcedures(ModuleName As String, Optional WithParentInfo = False)
    Dim Proc            As TProc
    Dim Module          As TModule
    Dim i               As Long
    
    Module = ReadModule(ModuleName)
    If Module.ProceduresCount = 0 Then Exit Sub
    For i = 0 To Module.ProceduresCount - 1
        Proc = Module.Procedures(i)
        Debug.Print IIf(WithParentInfo, Proc.ModuleName & ".", vbNullString) & Proc.Name
    Next
End Sub

Public Function ReadProject(VBProj As VBProject) As TProject
    Dim Modules()       As TModule
    Dim Proj            As TProject
    Dim ItemTypeName    As String
    Dim i               As Long
    Dim Item            As Variant
    Dim ModuleTypeNames As Variant
    
    ModuleTypeNames = Array("Code Module", "Class Module", "UserForm")
    
    For Each Item In VBProj.VBComponents
        ItemTypeName = ComponentTypeToString(Item.Type)
        If (UBound(Filter(ModuleTypeNames, ItemTypeName)) > -1) Then
            ReDim Preserve Modules(i)
            Modules(i) = ReadModule(Item.Name)
            i = i + 1
        End If
    Next Item
    
    Proj.ModulesCount = i
    Proj.Modules = Modules
    ReadProject = Proj
End Function

Public Function ReadModule(ModuleName As String) As TModule
    
    'Microsoft Visual Basic for Applications Extensibility 5.3 library

    Dim VBProj          As VBIDE.VBProject
    Dim VBComp          As VBIDE.VBComponent
    Dim CodeMod         As VBIDE.CodeModule
    Dim LineNum         As Long
    Dim ProcName        As String
    Dim ProcKind        As VBIDE.vbext_ProcKind
    
    Dim Procedures()    As TProc
    Dim Module          As TModule
    Dim i               As Long
    

    Set VBProj = ActiveWorkbook.VBProject
    Set VBComp = VBProj.VBComponents(ModuleName)
    Set CodeMod = VBComp.CodeModule
    
    With CodeMod
        
        LineNum = .CountOfDeclarationLines + 1
        
        Do Until LineNum >= .CountOfLines
            ReDim Preserve Procedures(i)
            ProcName = .ProcOfLine(LineNum, ProcKind)
            Procedures(i).Name = ProcName
            Procedures(i).ModuleName = ModuleName
            i = i + 1
            LineNum = .ProcStartLine(ProcName, ProcKind) + .ProcCountLines(ProcName, ProcKind) + 1
        Loop
    End With
    Module.Name = ModuleName
    Module.ProceduresCount = i
    Module.Procedures = Procedures
    ReadModule = Module
End Function

Function ComponentTypeToString(ComponentType As VBIDE.vbext_ComponentType) As String
    'ComponentTypeToString from http://www.cpearson.com/excel/vbe.aspx
    Select Case ComponentType
    
        Case vbext_ct_ActiveXDesigner
            ComponentTypeToString = "ActiveX Designer"
            
        Case vbext_ct_ClassModule
            ComponentTypeToString = "Class Module"
            
        Case vbext_ct_Document
            ComponentTypeToString = "Document Module"
            
        Case vbext_ct_MSForm
            ComponentTypeToString = "UserForm"
            
        Case vbext_ct_StdModule
            ComponentTypeToString = "Code Module"
            
        Case Else
            ComponentTypeToString = "Unknown Type: " & CStr(ComponentType)
            
    End Select
    
End Function
