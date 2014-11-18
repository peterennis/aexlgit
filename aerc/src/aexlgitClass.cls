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
Option Compare Text

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
' History:  See comment details, basChangeLog, commit messages on GitHub
' GitHub:   https://github.com/peterennis/aexlgit
' Ref:      http://www.rondebruin.nl/win/s9/win002.htm
'=======================================================================

Private Declare Sub Sleep Lib "kernel32" (ByVal lngMilliSeconds As Long)

Private Const aexlgitVERSION As String = "0.1.0"
Private Const aexlgitVERSION_DATE As String = "November 14, 2014"

Private Type mySetupType
    SourceFolder As String
    ImportFolder As String
End Type

' Ref: http://www.pcreview.co.uk/forums/macro-save-macros-t1918984p2.html
Private Const vbext_ct_StdModule As Long = 1
Private Const vbext_ct_ClassModule As Long = 2
Private Const vbext_ct_MSForm As Long = 3
Private Const vbext_ct_Document As Long = 100

Private aexlgitType As mySetupType
Private aexlgitSourceFolder As String
''' Private aexlgitblnCustomSourceFolder As Boolean
''' Private aexlgitstrSourceLocation As String

Private Sub Class_Initialize()
' Ref: http://www.cadalyst.com/cad/autocad/programming-with-class-part-1-5050
' Ref: http://www.bigresource.com/Tracker/Track-vb-cyJ1aJEyKj/
' Ref: http://stackoverflow.com/questions/1731052/is-there-a-way-to-overload-the-constructor-initialize-procedure-for-a-class-in

    On Error GoTo 0
    'MsgBox Application.VBE.ActiveVBProject.Name
    'MsgBox Application.Caption
    If Application.VBE.ActiveVBProject.Name = "aexlgit" Then
        Application.Caption = Application.VBE.ActiveVBProject.Name & " " & aexlgitVERSION
    End If
    ' Provide a default value for the SourceFolder and ImportFolder properties
    aexlgitSourceFolder = "default"
    aexlgitType.SourceFolder = "C:\ae\aexlgit\aerc\src\"

    Debug.Print "Class_Initialize"
    Debug.Print , "Default for aegitSourceFolder = " & aexlgitSourceFolder
    Debug.Print , "Default for aegitType.SourceFolder = " & aexlgitType.SourceFolder
End Sub

Private Sub Class_Terminate()
    On Error GoTo 0
    Debug.Print
    Debug.Print "Class_Terminate"
    Debug.Print , "aexlgit VERSION: " & aexlgitVERSION
    Debug.Print , "aexlgit VERSION_DATE: " & aexlgitVERSION_DATE
End Sub

Public Property Get SourceFolder() As String
    On Error GoTo 0
    SourceFolder = aexlgitSourceFolder
End Property

Public Property Let SourceFolder(ByVal strSourceFolder As String)
    On Error GoTo 0
    ' Ref: http://www.techrepublic.com/article/build-your-skills-using-class-modules-in-an-access-database-solution/5031814
    ' Ref: http://www.utteraccess.com/wiki/index.php/Classes
    aexlgitSourceFolder = strSourceFolder
End Property

Public Property Get DocumentTheExcelCode(Optional ByVal varDebug As Variant) As Boolean

    On Error GoTo PROC_ERR

    'MsgBox "DocumentTheExcelCode"
    If IsMissing(varDebug) Then
        Debug.Print "Get DocumentTheDatabase"
        Debug.Print , "varDebug IS missing so no parameter is passed to aeDocumentTheDatabase"
        Debug.Print , "DEBUGGING IS OFF"
        DocumentTheExcelCode = aeDocumentTheExcelCode
    Else
        Debug.Print "Get DocumentTheDatabase"
        Debug.Print , "varDebug IS NOT missing so a variant parameter is passed to aeDocumentTheDatabase"
        Debug.Print , "DEBUGGING TURNED ON"
        DocumentTheExcelCode = aeDocumentTheExcelCode(varDebug)
    End If

PROC_EXIT:
    Exit Property

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure DocumentTheExcelCode of Class aexlgitClass"
    Resume PROC_EXIT

End Property

Private Function ThisIsAnAddIn(ByVal strAddIn As String, _
                                Optional ByVal varDebug As Variant) As Boolean
    Dim rw As Byte
    Dim varAddIn As Variant
    ThisIsAnAddIn = False
    rw = 1
    For Each varAddIn In Application.AddIns
        If Not IsMissing(varDebug) Then Debug.Print varAddIn.Name, varAddIn.Installed
        If varAddIn.Name = strAddIn Then
            ThisIsAnAddIn = True
            Exit Function
        Else
            rw = rw + 1
        End If
    Next
End Function

Private Function aeDocumentTheExcelCode(Optional ByVal varDebug As Variant) As Boolean

    Dim bExport As Boolean
    Dim wkbSource As Object
    Set wkbSource = ThisWorkbook

    Dim strSourceWorkbook As String
    Dim strExportPath As String
    Dim strFileName As String
    Dim blnDebug As Boolean

    'MsgBox "aeDocumentTheExcelCode"
    On Error GoTo PROC_ERR

    Dim objComponent As Object
    Set objComponent = ThisWorkbook.VBProject.VBComponents

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
    If FolderWithVBAProjectFiles = "Error" Then
        MsgBox "Export Folder does not exist"
        Exit Function
    End If

    Debug.Print , "FolderWithVBAProjectFiles = " & FolderWithVBAProjectFiles
    'Stop

    On Error Resume Next
        Kill FolderWithVBAProjectFiles & "\*.*"
    On Error GoTo PROC_ERR

    If ThisIsAnAddIn("DirTree.xla", varDebug) Then
    Else
        ' NOTE: This workbook must be open in Excel
        strSourceWorkbook = ActiveWorkbook.Name
        Set wkbSource = Application.Workbooks(strSourceWorkbook)

        If wkbSource.VBProject.Protection = 1 Then
            MsgBox "The VBA in this workbook is protected!" & vbCrLf & _
                    "It is not possible to export the code."
            Exit Function
        End If
    End If

    strExportPath = FolderWithVBAProjectFiles & "\"

    For Each objComponent In wkbSource.VBProject.VBComponents

        bExport = True
        strFileName = objComponent.Name

        ' Concatenate the correct filename for export.
        Select Case objComponent.Type
            Case vbext_ct_ClassModule
                strFileName = strFileName & ".cls"
            Case vbext_ct_MSForm
                strFileName = strFileName & ".frm"
            Case vbext_ct_StdModule
                strFileName = strFileName & ".bas"
            Case vbext_ct_Document
                strFileName = strFileName & ".sht"
            Case Else
                MsgBox "Type not known: " & objComponent.Type
        End Select

        If bExport Then
            ' Export the component to a text file.
            If Left$(strFileName, 3) <> "zzz" Then
                objComponent.Export strExportPath & strFileName
            End If

            ' remove it from the project if you want
            'wkbSource.VBProject.VBComponents.Remove objComponent

        End If

    Next objComponent

    If Not IsMissing(varDebug) Then
        OutputListOfExcelProperties "DebugIt"
    Else
        OutputListOfExcelProperties
    End If

    aeDocumentTheExcelCode = True

PROC_EXIT:
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeDocumentTheExcelCode of Class aexlgitClass"
    Resume PROC_EXIT

End Function

Public Sub OutputListOfExcelProperties(Optional ByVal varDebug As Variant)
' Ref: http://stackoverflow.com/questions/17406585/vba-set-custom-document-property
' Workbook.BuiltinDocumentProperties Property
' Ref: http://msdn.microsoft.com/en-us/library/bb220896.aspx

    On Error GoTo PROC_ERR

    Dim wbk As Workbook
    Set wbk = Application.ThisWorkbook

    Dim docProp As DocumentProperty
    Dim strExportPathFileName As String
    Dim fle As Integer

    fle = FreeFile()
    strExportPathFileName = FolderWithVBAProjectFiles & "\OutputListOfExcelProperties.txt"
    Open strExportPathFileName For Output As #fle

    If Not IsMissing(varDebug) Then Debug.Print "OutputListOfExcelProperties"

    ' Ref: http://stackoverflow.com/questions/16642362/how-to-get-the-following-code-to-continue-on-error
    For Each docProp In wbk.CustomDocumentProperties
        If Not IsMissing(varDebug) Then Debug.Print , docProp.Name & ": " & docProp.Value
        Print #fle, docProp.Name & ": " & docProp.Value
    Next

    If Not ThisIsAnAddIn("DirTree.xla", varDebug) Then
        For Each docProp In ActiveWorkbook.BuiltinDocumentProperties
            If Not IsMissing(varDebug) Then Debug.Print , docProp.Name & ": " & docProp.Value
            Print #fle, docProp.Name & ": " & docProp.Value
        Next
    End If

PROC_EXIT:
    Close fle
    Set wbk = Nothing
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

    On Error GoTo 0
    Dim wkbTarget As Excel.Workbook
    'Dim objFSO As Scripting.FileSystemObject
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")

    'Dim objFile As Scripting.File
    Dim objFile As Object
    
    Dim strTargetWorkbook As String
    Dim strImportPath As String

    Dim objComponent As Object
    Set objComponent = CreateObject("VBIDE.VBComponent")

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

    ' NOTE: This workbook must be open in Excel
    strTargetWorkbook = ActiveWorkbook.Name
    Set wkbTarget = Application.Workbooks(strTargetWorkbook)

    If wkbTarget.VBProject.Protection = 1 Then
    MsgBox "The VBA in this workbook is protected!" & vbCrLf & _
        "It is not possible to Import the code."
    Exit Sub
    End If

    ' NOTE: Path where the code modules are located.
    strImportPath = FolderWithVBAProjectFiles & "\"

    '''Set objFSO = New Scripting.FileSystemObject
    If objFSO.GetFolder(strImportPath).Files.Count = 0 Then
       MsgBox "There are no files to import"
       Exit Sub
    End If

    ' Delete all modules/Userforms from the ActiveWorkbook
    Call DeleteVBAModulesAndUserForms

    Set objComponent = wkbTarget.VBProject.VBComponents

    ' Import all the code modules in the specified path
    ' to the ActiveWorkbook.
    For Each objFile In objFSO.GetFolder(strImportPath).Files

        If (objFSO.GetExtensionName(objFile.Name) = "cls") Or _
            (objFSO.GetExtensionName(objFile.Name) = "frm") Or _
            (objFSO.GetExtensionName(objFile.Name) = "bas") Then
            objComponent.Import objFile.Path
        End If
    
    Next objFile

    MsgBox "Import is finished."

End Sub

Private Function FolderWithVBAProjectFiles() As String

    On Error GoTo 0
    Dim WshShell As Object
    Dim FSO As Object
    Dim SpecialPath As String

    Set WshShell = CreateObject("WScript.Shell")
    Set FSO = CreateObject("Scripting.FileSystemObject")

    SpecialPath = WshShell.SpecialFolders("MyDocuments")

    If aexlgitSourceFolder = "default" Then
    
        FolderWithVBAProjectFiles = aexlgitType.SourceFolder
    
    ElseIf IsNull(aexlgitSourceFolder) Then

        If Right$(SpecialPath, 1) <> "\" Then
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
    
    'MsgBox "FolderWithVBAProjectFiles=" & FolderWithVBAProjectFiles
    If Not FSO.FolderExists(FolderWithVBAProjectFiles) Then
        FolderWithVBAProjectFiles = "Error"
    End If

End Function

Private Function DeleteVBAModulesAndUserForms() As Boolean

    On Error GoTo 0
    Dim objProject As Object
    Set objProject = CreateObject("VBIDE.VBProject")

    Dim objComponent As Object
    Set objComponent = CreateObject("VBIDE.VBComponent")

    Set objProject = ActiveWorkbook.VBProject

    For Each objComponent In objProject.VBComponents
        If objComponent.Type = vbext_ct_Document Then
            ' ThisWorkbook or Worksheet module we do nothing
        Else
            objProject.VBComponents.Remove objComponent
        End If
    Next objComponent

End Function


