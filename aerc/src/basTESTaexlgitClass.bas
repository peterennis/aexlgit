Attribute VB_Name = "basTESTaexlgitClass"
Option Explicit
Option Compare Text
Option Private Module

' Default Usage:
' The following folders are used if no custom configuration is provided:
' aexlgitType.SourceFolder = "C:\ae\aexlgit\aerc\src\"
' Run in immediate window:                  MYXLPROJECT_TEST
' Show debug output in immediate window:    Uncomment aexlgitClassTest varDebug:="DebugIt"
'
' Custom Usage:
' Public Const FOLDER_FOR_VBA_PROJECT_FILES = "Z:\The\Source\Folder\srx.MYPROJECT\"
' For custom configuration of the output source folder in aexlClassTest use:
' oDbObjects.SourceFolder = FOLDER_FOR_VBA_PROJECT_FILES
' Run in immediate window: MYXLPROJECT_TEST
'

Public Function MYXLPROJECT_TEST() As Boolean
    On Error GoTo 0
    'aexlgitClassTest
    aexlgitClassTest varDebug:="Debugit"
End Function

Public Function aexlgitClassTest(Optional ByVal varDebug As Variant, _
                                    Optional ByVal varSrcFldr As Variant, _
                                    Optional ByVal varXmlFldr As Variant, _
                                    Optional ByVal varXmlData As Variant) As Boolean

    On Error GoTo PROC_ERR

    Dim oXlObjects As aexlgitClass
    Set oXlObjects = New aexlgitClass

    Dim bln1 As Boolean

    If Not IsMissing(varSrcFldr) Then oXlObjects.SourceFolder = varSrcFldr      ' THE_SOURCE_FOLDER
    '''If Not IsMissing(varXmlFldr) Then oXlObjects.XMLFolder = varXmlFldr         ' THE_XML_FOLDER

Test1:
    '=============
    ' TEST 1
    '=============
    Debug.Print
    Debug.Print "vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv"
    Debug.Print "1. aexlgitClassTest => DocumentTheExcelCode"
    Debug.Print "aexlgitClassTest"
    If IsMissing(varDebug) Then
        Debug.Print , "varDebug IS missing so no parameter is passed to DocumentTheExcelCode"
        Debug.Print , "DEBUGGING IS OFF"
        bln1 = oXlObjects.DocumentTheExcelCode()
    Else
        Debug.Print , "varDebug IS NOT missing so blnDebug is set to True"
        bln1 = oXlObjects.DocumentTheExcelCode("WithDebugging")
    End If
    Debug.Print "^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^"
    Debug.Print

PROC_EXIT:
    Exit Function

PROC_ERR:
    If Err = 1004 Then ' VBA Project Not Trusted - "Programmatic access to the Visual Basic Project is not trusted..."
        MsgBox "VBA Project Not Trusted", vbCritical, "aexlgitClassTest"
        Stop
        'Resume PROC_EXIT
    Else
        MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aexlgitClassTest of Module basTESTaexlgitClass"
        Resume PROC_EXIT
    End If

End Function

Public Sub ListAddInFiles()
    Dim rw As Byte
    Dim varAddIn As Variant
    rw = 1
    For Each varAddIn In Application.AddIns
        Debug.Print varAddIn.Name, varAddIn.Installed
        rw = rw + 1
    Next
End Sub


