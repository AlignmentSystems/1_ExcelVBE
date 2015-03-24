Attribute VB_Name = "CustomVBE2"
Option Explicit
Option Compare Text
Option Base 0
'============================================================================================================================
'
'
'   Author      :       John Greenan
'   Email       :       
'   Company     :       Alignment Systems Limited
'   Date        :       10th September 2014
'
'   Purpose     :       Matching Engine in Excel VBA for Alignment Systems Limited
'
'   References  :       See VB Module FL for list extracted from VBE
'   References  :
'============================================================================================================================
Const cstrTarget As String = "Const cstrMethodName As String"
Const cstrThisModule As String = "CustomVBE"
Const cstrFL As String = "FL"
Const cstrGlobals As String = "Globals"
Const cstrWTimer As String = "WTimer"
Const cstrJavaClass As String = "JavaClass"
Const cstrLogWrapper As String = "LogWrapper"
Const cstrMessageWrapper As String = "MessageWrapper"
Const cstrThisWorkbook As String = "ThisWorkbook"
Const cstrThisModule2 As String = "CustomVBE2"

Private Function ValidToProcessThisComponent(IncomingComponentName As String) As Boolean
'============================================================================================================================
'
'
'   Author      :       John Greenan
'   Email       :       
'   Company     :       Alignment Systems Limited
'   Date        :       10th September 2014
'
'   Purpose     :       Matching Engine in Excel VBA for Alignment Systems Limited
'
'   References  :       See VB Module FL for list extracted from VBE
'   References  :
'============================================================================================================================
'Variables
Dim ComponentsToIgnore(0 To 8) As String
Dim inc As Integer

ComponentsToIgnore(0) = cstrThisModule
ComponentsToIgnore(1) = cstrFL
ComponentsToIgnore(2) = cstrGlobals
ComponentsToIgnore(3) = cstrWTimer
ComponentsToIgnore(4) = cstrJavaClass
ComponentsToIgnore(5) = cstrLogWrapper
ComponentsToIgnore(6) = cstrMessageWrapper
ComponentsToIgnore(7) = cstrThisWorkbook
ComponentsToIgnore(8) = cstrThisModule2

ValidToProcessThisComponent = True

For inc = LBound(ComponentsToIgnore) To UBound(ComponentsToIgnore)
    If StrComp(ComponentsToIgnore(inc), IncomingComponentName, vbTextCompare) = 0 Then
        ValidToProcessThisComponent = False
        Exit For
    End If
Next

Erase ComponentsToIgnore

End Function

Private Function AddFunctionNamesToCodeModule(IncomingCodeModule As VBIDE.CodeModule, ProcedureType As VBIDE.vbext_ComponentType) As Boolean
'============================================================================================================================
'
'
'   Author      :       John Greenan
'   Email       :       
'   Company     :       Alignment Systems Limited
'   Date        :       10th September 2014
'
'   Purpose     :       Matching Engine in Excel VBA for Alignment Systems Limited
'
'   References  :       See VB Module FL for list extracted from VBE
'   References  :
'============================================================================================================================
'Constants
Const cstrMethodName As String = "CustomVBE2.AddFunctionNamesToCodeModule "
'Variables
Dim dblDifferenceInSeconds As Double
Dim lngFirstLineOfProcedure As Long
Dim lngLastLineOfProcedure As Long
Dim lngProcedureOffset As Long
Dim lngModuleCursorLine As Long
Dim lngModuleCountOfLines As Long
Dim lngModuleFirstCodeLine As Long
Dim strProcedureName As String
Dim strVBCodeString As String
Dim ProcType As VBIDE.vbext_ProcKind
Dim blnIgnore As Boolean

On Error GoTo ErrHandler
AddFunctionNamesToCodeModule = False

blnIgnore = False

If ProcedureType = VBIDE.vbext_ComponentType.vbext_ct_ActiveXDesigner Then
    blnIgnore = True
End If

If ProcedureType = VBIDE.vbext_ComponentType.vbext_ct_Document Then
    blnIgnore = True
End If

If ProcedureType = VBIDE.vbext_ComponentType.vbext_ct_MSForm Then
    blnIgnore = True
End If

If blnIgnore Then
    Debug.Print "Ignore..." & IncomingCodeModule.Name
Else
    Debug.Print "Examining module..." & IncomingCodeModule.Name
    
    With IncomingCodeModule
        lngModuleCountOfLines = .CountOfLines
        lngModuleFirstCodeLine = .CountOfDeclarationLines + 1
        lngModuleCursorLine = lngModuleFirstCodeLine
        strVBCodeString = .Lines(lngModuleCursorLine, 1)
    End With
'       We now have the first line and the last line in the module...
    Do While lngModuleCursorLine < lngModuleCountOfLines
        Do While StrComp(strVBCodeString, "", vbTextCompare) <> 0
            lngModuleCursorLine = lngModuleCursorLine + 1
            strVBCodeString = IncomingCodeModule.Lines(lngModuleCursorLine, 1)
        Loop
'       So, if we get here we now have a line at the start of a procedure that actually has something
'       meaningful...
        strProcedureName = IncomingCodeModule.ProcOfLine(lngModuleCursorLine, ProcType)
        lngFirstLineOfProcedure = IncomingCodeModule.ProcBodyLine(strProcedureName, ProcType)
        lngProcedureOffset = lngFirstLineOfProcedure - IncomingCodeModule.ProcStartLine(strProcedureName, ProcType)
        lngLastLineOfProcedure = lngFirstLineOfProcedure + IncomingCodeModule.ProcCountLines(strProcedureName, ProcType) - lngProcedureOffset - 1
        If ProcType = vbext_pk_Proc Then
            If Not FindProcedureNameConstant(IncomingCodeModule, strProcedureName, lngFirstLineOfProcedure, lngLastLineOfProcedure) Then
                If AddProcedureNameConstant(IncomingCodeModule, strProcedureName, lngFirstLineOfProcedure, lngLastLineOfProcedure) Then
                Else
                End If
            Else
                Debug.Print "Already added..." & IncomingCodeModule.Name & strProcedureName
            End If
        Else
            Debug.Print "Not a procedure..." & IncomingCodeModule.Name & strProcedureName
        End If
        lngModuleCursorLine = lngFirstLineOfProcedure + IncomingCodeModule.ProcCountLines(strProcedureName, ProcType) - lngProcedureOffset
    Loop
End If

On Error GoTo 0
AddFunctionNamesToCodeModule = True

Exit Function

ErrHandler:



End Function

Private Function AddProcedureNameConstant(CodeModule As VBIDE.CodeModule, ProcedureName As String, FirstLineOfProcedure As Long, LastLineOfProcedure As Long) As Boolean
'============================================================================================================================
'
'
'   Author      :       John Greenan
'   Email       :       
'   Company     :       Alignment Systems Limited
'   Date        :       10th September 2014
'
'   Purpose     :       Matching Engine in Excel VBA for Alignment Systems Limited
'
'   References  :       See VB Module FL for list extracted from VBE
'   References  :
'============================================================================================================================
'Constants
Const cstrMethodName As String = "CustomVBE2.AddProcedureNameConstant "
'Variables
Dim dblDifferenceInSeconds As Double
Dim strBuiltString As String
Dim strVBCodeString As String
Dim lngLineToAddFunctionNameConstant As Long

On Error GoTo ErrHandler
AddProcedureNameConstant = False

'Build Up the String
strBuiltString = cstrTarget & " = """ & CodeModule.Name & "." & ProcedureName & Chr(VBA.KeyCodeConstants.vbKeySpace) & " """
'   So, if we get to here we know we want to insert the methodname constant.
'   BUT - the oCodeModule.ProcBodyLine(strProcedureName, ProcType) returns the first line of the body of the procedure
'   So we need to bump down until we are not getting comment fields
'   The first line of the procedure is going to be the signature "public function blah" or whatever
'   So, we have to move one line down..
strVBCodeString = CodeModule.Lines(FirstLineOfProcedure + 1, 1)

lngLineToAddFunctionNameConstant = FirstLineOfProcedure + 1

Do While StrComp(Left(Trim(strVBCodeString), 1), "'", vbTextCompare) = 0
    lngLineToAddFunctionNameConstant = lngLineToAddFunctionNameConstant + 1
    strVBCodeString = CodeModule.Lines(lngLineToAddFunctionNameConstant, 1)
Loop

Debug.Print "Add...[" & CStr(lngLineToAddFunctionNameConstant) & "] " & strBuiltString

CodeModule.InsertLines lngLineToAddFunctionNameConstant, strBuiltString

On Error GoTo 0
AddProcedureNameConstant = True
Exit Function
ErrHandler:


End Function

Private Function FindProcedureNameConstant(CodeModule As VBIDE.CodeModule, ProcedureName As String, FirstLineOfProcedure As Long, LastLineOfProcedure As Long) As Boolean
'============================================================================================================================
'
'
'   Author      :       John Greenan
'   Email       :       
'   Company     :       Alignment Systems Limited
'   Date        :       10th September 2014
'
'   Purpose     :       Matching Engine in Excel VBA for Alignment Systems Limited
'
'   References  :       See VB Module FL for list extracted from VBE
'   References  :
'============================================================================================================================
'Constants
Const cstrMethodName As String = "CustomVBE2.FindProcedureNameConstant "
'Variables
Dim dblDifferenceInSeconds As Double
Dim strVBCodeString  As String
Dim lngProcedureIterator As Long
Dim blnFlagReturn As Boolean


On Error GoTo 0
FindProcedureNameConstant = False
blnFlagReturn = False

lngProcedureIterator = FirstLineOfProcedure
Do While lngProcedureIterator <= LastLineOfProcedure
'   Get the line of code at position lngProcedureIterator
    strVBCodeString = CodeModule.Lines(lngProcedureIterator, 1)
'   If the line of code at position lngProcedureIterator matches what we are looking for then
'   let blnFunctionNameConstantFound = True and exit the loop...
    If StrComp(Left(Trim(strVBCodeString), Len(cstrTarget)), cstrTarget, vbTextCompare) = 0 Then
        blnFlagReturn = True
        Exit Do
    End If
    lngProcedureIterator = lngProcedureIterator + 1
Loop

FindProcedureNameConstant = blnFlagReturn
On Error GoTo 0
Exit Function

ErrHandler:




End Function

Private Function EntryPointAddFunctionName() As Boolean
'============================================================================================================================
'
'
'   Author      :       John Greenan
'   Email       :       
'   Company     :       Alignment Systems Limited
'   Date        :       10th September 2014
'
'   Purpose     :       Matching Engine in Excel VBA for Alignment Systems Limited
'
'   References  :       See VB Module FL for list extracted from VBE
'   References  :
'============================================================================================================================
'Constants
Const cstrMethodName As String = "CustomVBE2.EntryPointAddFunctionName "
'Variables
Dim oCom As VBIDE.VBComponent


On Error GoTo 0
EntryPointAddFunctionName = False

'Iterate through all components...
For Each oCom In ThisWorkbook.VBProject.VBComponents
    If ValidToProcessThisComponent(oCom.Name) Then
        AddFunctionNamesToCodeModule oCom.CodeModule, oCom.Type
    End If
Next
Set oCom = Nothing

EntryPointAddFunctionName = True

On Error GoTo 0

Exit Function


ErrHandler:



End Function


Private Function IsThisAValidLine(TestValue As String, StringOfCode As String) As Boolean
'============================================================================================================================
'
'
'   Author      :       John Greenan
'   Email       :       
'   Company     :       Alignment Systems Limited
'   Date        :       10th September 2014
'
'   Purpose     :       Matching Engine in Excel VBA for Alignment Systems Limited
'
'   References  :       See VB Module FL for list extracted from VBE
'   References  :
'============================================================================================================================
'Constants
Const cstrMethodName As String = "CustomVBE2.IsThisAValidLine "
'Variables
Dim lngLengthOfCode As Long
Dim lngLengthOfTest As Long
Dim strTestString As String


On Error GoTo ErrHandler
IsThisAValidLine = False

lngLengthOfCode = Len(StringOfCode)
lngLengthOfTest = Len(TestValue)

If lngLengthOfCode >= lngLengthOfTest Then
    strTestString = Left(StringOfCode, lngLengthOfTest)
End If

If StrComp(strTestString, TestValue, vbTextCompare) = 0 Then
'   If's a line starting with "TestValue"..
    IsThisAValidLine = False
Else
    IsThisAValidLine = True
End If

On Error GoTo 0
Exit Function
ErrHandler:

End Function

