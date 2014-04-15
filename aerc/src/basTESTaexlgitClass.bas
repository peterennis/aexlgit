Attribute VB_Name = "basTESTaexlgitClass"
Option Explicit

' Default Usage:
' The following folders are used if no custom configuration is provided:
' aexlgitType.SourceFolder = "C:\ae\aegit\aerc\srx\"
' Run in immediate window:                  MYXLPROJECT_TEST
' Show debug output in immediate window:    Uncomment aexlClassTest ("debug")
'
' Custom Usage:
' Public Const FOLDER_WITH_VBA_PROJECT_FILES = "Z:\The\Source\Folder\srx.MYPROJECT"
' For custom configuration of the output source folder in aexlClassTest use:
' oDbObjects.SourceFolder = FOLDER_WITH_VBA_PROJECT_FILES
' Run in immediate window: MYXLPROJECT_TEST
'

Public Function MYXLPROJECT_TEST() As Boolean
    On Error GoTo 0
    'aexlgitClassTest
    aexlgitClassTest ("debug")
End Function

Private Function aexlgitClassTest(Optional ByVal Debugit As Variant) As Boolean

    On Error GoTo PROC_ERR

    Dim oXlObjects As aexlgitClass
    Set oXlObjects = New aexlgitClass

    Dim bln1 As Boolean

    'oXlObjects.SourceFolder = FOLDER_WITH_VBA_PROJECT_FILES

Test1:
    '=============
    ' TEST 1
    '=============
    Debug.Print
    Debug.Print "vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv"
    Debug.Print "1. aexlgitClassTest => DocumentTheExcelCode"
    Debug.Print "aexlgitClassTest"
    If IsMissing(Debugit) Then
        Debug.Print , "Debugit IS missing so no parameter is passed to DocumentTheExcelCode"
        Debug.Print , "DEBUGGING IS OFF"
        bln1 = oXlObjects.DocumentTheExcelCode()
    Else
        Debug.Print , "Debugit IS NOT missing so blnDebug is set to True"
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
