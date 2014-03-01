VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aexlgitClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Copyright (c) 2013 Peter F. Ennis
'This library is free software; you can redistribute it and/or
'modify it under the terms of the GNU Lesser General Public
'License as published by the Free Software Foundation;
'version 3.0.
'This library is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
'Lesser General Public License for more details.
'You should have received a copy of the GNU Lesser General Public
'License along with this library; if not, visit
'http://www.gnu.org/licenses/lgpl-3.0.txt


'=======================================================================
' Author:   Peter F. Ennis
' Date:     July 13, 2013
' Comment:  Create class for revision control
' History:  See comment details, basChangeLog, commit messages on github
'=======================================================================

Private Declare Sub Sleep Lib "kernel32" (ByVal lngMilliSeconds As Long)

Private Const aexlgitVERSION As String = "0.0.5"
Private Const aexlgitVERSION_DATE As String = "February 28, 2014"
Private Const THE_DRIVE As String = "C"
Private Const gcfHandleErrors As Boolean = True

Private Type mySetupType
    SourceFolder As String
    ImportFolder As String
    UseImportFolder As Boolean
End Type

Private aexlgitType As mySetupType
Private aexlgitSourceFolder As String
Private aexlgitImportFolder As String
Private aexlgitUseImportFolder As Boolean
Private aexlgitblnCustomSourceFolder As Boolean
Private aexlgitstrSourceLocation As String
Private aexlgitImportLocation As String

' Ref: http://www.rondebruin.nl/win/s9/win002.htm
' Requires reference to the VBA extensibility library. Click on Tools-References in the VBE, and
' scroll down and tick the entry for Microsoft Visual Basic for Applications Extensibility 5.3
' Scripting.FileSystemObject requires reference to Microsoft Scripting Runtime

Private Sub Class_Initialize()
' Ref: http://www.cadalyst.com/cad/autocad/programming-with-class-part-1-5050
' Ref: http://www.bigresource.com/Tracker/Track-vb-cyJ1aJEyKj/
' Ref: http://stackoverflow.com/questions/1731052/is-there-a-way-to-overload-the-constructor-initialize-procedure-for-a-class-in

    ' provide a default value for the SourceFolder and ImportFolder properties
    aexlgitSourceFolder = "default"
    aexlgitImportFolder = "default"
    aexlgitUseImportFolder = False
    aexlgitType.SourceFolder = "C:\ae\aexlgit\aerc\src"
    '''x aexlgitType.ImportFolder = "C:\ae\aexlgit\aerc\imp"
    aexlgitType.UseImportFolder = False

    Debug.Print "Class_Initialize"
    Debug.Print , "Default for aegitSourceFolder = " & aexlgitSourceFolder
    Debug.Print , "Default for aegitImportFolder = " & aexlgitImportFolder
    Debug.Print , "Default for aegitType.SourceFolder = " & aexlgitType.SourceFolder
    Debug.Print , "Default for aegitType.ImportFolder = " & aexlgitType.ImportFolder
    Debug.Print , "Default for aegitType.UseImportFolder = " & aexlgitType.UseImportFolder
End Sub

Private Sub Class_Terminate()
    Debug.Print
    Debug.Print "Class_Terminate"
    Debug.Print , "aexlgit VERSION: " & aexlgitVERSION
    Debug.Print , "aexlgit VERSION_DATE: " & aexlgitVERSION_DATE
End Sub

Property Get SourceFolder() As String
    SourceFolder = aexlgitSourceFolder
End Property

Property Let SourceFolder(ByVal strSourceFolder As String)
    ' Ref: http://www.techrepublic.com/article/build-your-skills-using-class-modules-in-an-access-database-solution/5031814
    ' Ref: http://www.utteraccess.com/wiki/index.php/Classes
    aexlgitSourceFolder = strSourceFolder
End Property

Property Get ImportFolder() As String
    ImportFolder = aexlgitImportFolder
End Property

Property Let ImportFolder(ByVal strImportFolder As String)
    aexlgitImportFolder = strImportFolder
End Property

Property Let UseImportFolder(ByVal blnUseImportFolder As Boolean)
    aexlgitUseImportFolder = blnUseImportFolder
End Property

Property Get DocumentTheExcelCode(Optional DebugTheCode As Variant) As Boolean
    'MsgBox "DocumentTheExcelCode"
    If IsMissing(DebugTheCode) Then
        Debug.Print "Get DocumentTheDatabase"
        Debug.Print , "DebugTheCode IS missing so no parameter is passed to aeDocumentTheDatabase"
        Debug.Print , "DEBUGGING IS OFF"
        DocumentTheExcelCode = aeDocumentTheExcelCode
    Else
        Debug.Print "Get DocumentTheDatabase"
        Debug.Print , "DebugTheCode IS NOT missing so a variant parameter is passed to aeDocumentTheDatabase"
        Debug.Print , "DEBUGGING TURNED ON"
        DocumentTheExcelCode = aeDocumentTheExcelCode(DebugTheCode)
    End If
End Property

Private Function aeDocumentTheExcelCode(Optional varDebug As Variant) As Boolean

    Dim bExport As Boolean
    Dim wkbSource As Excel.Workbook
    Dim strSourceWorkbook As String
    Dim strExportPath As String
    Dim strFileName As String
    Dim cmpComponent As VBIDE.VBComponent
    Dim blnDebug As Boolean

    Debug.Print "aeDocumentTheExcelCode"
    If IsMissing(varDebug) Then
        blnDebug = False
        Debug.Print , "varDebug IS missing so blnDebug of aeDocumentTheExcelCode is set to False"
        Debug.Print , "DEBUGGING IS OFF"
    Else
        blnDebug = True
        Debug.Print , "varDebug IS NOT missing so blnDebug of aeDocumentTheExcelCode is set to True"
        Debug.Print , "NOW DEBUGGING..."
    End If

    ' The code modules will be exported in a folder named
    ' VBAProjectFiles in the Documents folder.
    ' The code below creates this folder if it does not exist
    ' or deletes all files in the folder if it exists.
    ' The Const FOLDER_WITH_VBA_PROJECT_FILES hardcodes an
    ' alternative location.
    If FolderWithVBAProjectFiles = "Error" Then
        MsgBox "Export Folder does not exist"
        Exit Function
    End If

    Debug.Print , "FolderWithVBAProjectFiles = " & FolderWithVBAProjectFiles
    'Stop

    On Error Resume Next
        Kill FolderWithVBAProjectFiles & "\*.*"
    On Error GoTo 0

    ' NOTE: This workbook must be open in Excel.
    strSourceWorkbook = ActiveWorkbook.Name
    Set wkbSource = Application.Workbooks(strSourceWorkbook)

    If wkbSource.VBProject.Protection = 1 Then
    MsgBox "The VBA in this workbook is protected," & _
        "not possible to export the code"
    Exit Function
    End If

    strExportPath = FolderWithVBAProjectFiles & "\"

    For Each cmpComponent In wkbSource.VBProject.VBComponents

        bExport = True
        strFileName = cmpComponent.Name

        ' Concatenate the correct filename for export.
        Select Case cmpComponent.Type
            Case vbext_ct_ClassModule
                strFileName = strFileName & ".cls"
            Case vbext_ct_MSForm
                strFileName = strFileName & ".frm"
            Case vbext_ct_StdModule
                strFileName = strFileName & ".bas"
            Case vbext_ct_Document
                strFileName = strFileName & ".sht"
            Case Else
                MsgBox "Type not known: " & cmpComponent.Type
        End Select

        If bExport Then
            ' Export the component to a text file.
            cmpComponent.Export strExportPath & strFileName

            ' remove it from the project if you want
            'wkbSource.VBProject.VBComponents.Remove cmpComponent

        End If

    Next cmpComponent

    If Not IsMissing(varDebug) Then
        OutputListOfExcelProperties "DebugIt"
    Else
        OutputListOfExcelProperties
    End If

    aeDocumentTheExcelCode = True

End Function

Public Sub OutputListOfExcelProperties(Optional varDebug As Variant)
' Ref: http://stackoverflow.com/questions/17406585/vba-set-custom-document-property
' Workbook.BuiltinDocumentProperties Property
' Ref: http://msdn.microsoft.com/en-us/library/bb220896.aspx

    ' Use a call stack and global error handler
    'If gcfHandleErrors Then On Error GoTo PROC_ERR
    'PushCallStack "OutputListOfExcelProperties"

    On Error GoTo PROC_ERR

    Dim wb As Workbook
    Set wb = Application.ThisWorkbook

    Dim docProp As DocumentProperty
    Dim propExists As Boolean
    Dim strExportPathFileName As String
    Dim fle As Integer

    fle = FreeFile()
    strExportPathFileName = FolderWithVBAProjectFiles & "\OutputListOfExcelProperties.txt"
    Open strExportPathFileName For Output As #fle

    If Not IsMissing(varDebug) Then Debug.Print "OutputListOfExcelProperties"

    ' Ref: http://stackoverflow.com/questions/16642362/how-to-get-the-following-code-to-continue-on-error
    For Each docProp In wb.CustomDocumentProperties
        If Not IsMissing(varDebug) Then Debug.Print , docProp.Name & ": " & docProp.Value
        Print #fle, docProp.Name & ": " & docProp.Value
    Next

    For Each docProp In ActiveWorkbook.BuiltinDocumentProperties
        If Not IsMissing(varDebug) Then Debug.Print , docProp.Name & ": " & docProp.Value
        Print #fle, docProp.Name & ": " & docProp.Value
    Next

PROC_EXIT:
    Close fle
    Set wb = Nothing
    'PopCallStack
    Exit Sub

PROC_ERR:
    If Err = -2147467259 Then     ' Method 'Value' of object ... failed
        'MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputListOfExcelProperties of Class aexlgitClass"
        Resume Next
    Else
        MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputListOfExcelProperties of Class aexlgitClass"
        'GlobalErrHandler
        Resume PROC_EXIT
    End If

End Sub

Private Sub ImportModules()

    Dim wkbTarget As Excel.Workbook
    Dim objFSO As Scripting.FileSystemObject
    Dim objFile As Scripting.File
    Dim strTargetWorkbook As String
    Dim strImportPath As String
    Dim strFileName As String
    Dim cmpComponents As VBIDE.VBComponents

    If ActiveWorkbook.Name = ThisWorkbook.Name Then
        MsgBox "Select another destination workbook" & _
        "Not possible to import in this workbook "
        Exit Sub
    End If

    ' Get the path to the folder with modules
    If FolderWithVBAProjectFiles = "Error" Then
        MsgBox "Import Folder not exist"
        Exit Sub
    End If

    ' NOTE: This workbook must be open in Excel.
    strTargetWorkbook = ActiveWorkbook.Name
    Set wkbTarget = Application.Workbooks(strTargetWorkbook)

    If wkbTarget.VBProject.Protection = 1 Then
    MsgBox "The VBA in this workbook is protected," & _
        "not possible to Import the code"
    Exit Sub
    End If

    ' NOTE: Path where the code modules are located.
    strImportPath = FolderWithVBAProjectFiles & "\"

    Set objFSO = New Scripting.FileSystemObject
    If objFSO.GetFolder(strImportPath).Files.Count = 0 Then
       MsgBox "There are no files to import"
       Exit Sub
    End If

    ' Delete all modules/Userforms from the ActiveWorkbook
    Call DeleteVBAModulesAndUserForms

    Set cmpComponents = wkbTarget.VBProject.VBComponents

    ' Import all the code modules in the specified path
    ' to the ActiveWorkbook.
    For Each objFile In objFSO.GetFolder(strImportPath).Files

        If (objFSO.GetExtensionName(objFile.Name) = "cls") Or _
            (objFSO.GetExtensionName(objFile.Name) = "frm") Or _
            (objFSO.GetExtensionName(objFile.Name) = "bas") Then
            cmpComponents.Import objFile.Path
        End If
    
    Next objFile

    MsgBox "Import is ready"

End Sub

Private Function FolderWithVBAProjectFiles() As String

    Dim WshShell As Object
    Dim FSO As Object
    Dim SpecialPath As String

    Set WshShell = CreateObject("WScript.Shell")
    Set FSO = CreateObject("scripting.filesystemobject")

    SpecialPath = WshShell.SpecialFolders("MyDocuments")

    If aexlgitSourceFolder = "default" Then
    
        FolderWithVBAProjectFiles = aexlgitType.SourceFolder
    
    ElseIf IsNull(FOLDER_WITH_VBA_PROJECT_FILES) Then

        If Right(SpecialPath, 1) <> "\" Then
            SpecialPath = SpecialPath & "\"
        End If

        If FSO.FolderExists(SpecialPath & "VBAProjectFiles") = False Then
            On Error Resume Next
            MkDir SpecialPath & "VBAProjectFiles"
            On Error GoTo 0
        End If

        If FSO.FolderExists(SpecialPath & "VBAProjectFiles") = True Then
            FolderWithVBAProjectFiles = SpecialPath & "VBAProjectFiles"
        Else
            FolderWithVBAProjectFiles = "Error"
        End If

    Else

        FolderWithVBAProjectFiles = aexlgitSourceFolder

    End If

End Function

Private Function DeleteVBAModulesAndUserForms()

    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent

    Set VBProj = ActiveWorkbook.VBProject

    For Each VBComp In VBProj.VBComponents
        If VBComp.Type = vbext_ct_Document Then
            ' ThisWorkbook or Worksheet module we do nothing
        Else
            VBProj.VBComponents.Remove VBComp
        End If
    Next VBComp

End Function

