Attribute VB_Name = "modMain"
Option Explicit

Sub ExportThem()
    Dim vbp As VBProject
    Dim oBookOrDocument As Object
    Dim oBookOrDocCollection As Object
    Dim vbc As Object
    Dim sAddition As String
    Dim sFilePath As String
    Dim sSubFolder As String
    Dim sFullFile As String
    Dim frmFolders As ufFolders
    Dim sSavePath As String

    On Error Resume Next
    Set vbp = ThisWorkbook.VBProject
    If Err.Number <> 0 Then
        MsgBox "In order for this tool to work, you must allow access to the VBA project (Excel Options, Trust center tab, Trust center settings, Macros tab)", vbExclamation + vbOKOnly, GCSAPPNAME
        Exit Sub
    End If
    Set vbp = Nothing

    Set frmFolders = New ufFolders
    With frmFolders
        .Initialize
        .Show
        If .OK Then
            If .WorkbookName Like "https:*" Then
                Set oBookOrDocument = Workbooks(Mid(.WorkbookName, InStrRev(.WorkbookName, "/") + 1, Len(.WorkbookName)))
            Else
                Set oBookOrDocument = Workbooks(Mid(.WorkbookName, InStrRev(.WorkbookName, Application.PathSeparator) + 1, Len(.WorkbookName)))
            End If
            sFilePath = .MainPath
            If oBookOrDocument.Saved = False Then
                If MsgBox("Before exporting modules and ribbonX," & vbNewLine & "your workbook must be saved." & vbNewLine & vbNewLine & _
                          "The workbook you selected appears to have unsaved changes." & vbNewLine & vbNewLine & _
                          "Continue without saving?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
            End If
            If .DeleteExisting Then
                EmptyFolder .MainPath
            End If
            Set vbp = oBookOrDocument.VBProject
            'Add all VBAProject components to relevant subfolders
            If Not vbp.Protection = vbext_pp_locked Then
                On Error Resume Next
                For Each vbc In vbp.VBComponents
                    Select Case vbc.Type
                    Case vbext_ct_MSForm
                        sAddition = ".frm"
                        sSubFolder = .UserformsFolder
                    Case vbext_ct_StdModule
                        sAddition = ".bas"
                        sSubFolder = .ModulesFolder
                    Case vbext_ct_ClassModule
                        sAddition = ".cls"
                        sSubFolder = .ClassModulesFolder
                    Case vbext_ct_Document
                        sAddition = ".cls"
                        sSubFolder = .ExcelObjectsFolder
                    End Select
                    ExportThisOne vbc, sFilePath, sAddition, sSubFolder
                Next
            End If
            'Generate txt file with all form control properties
            sSavePath = sFilePath & .UserformsFolder
            AddPath sSavePath
            ListForms oBookOrDocument, sSavePath
            If .cbxExcelFile Then
                'Add workbook itself to the binaries subfolder
                sSavePath = sFilePath & .ExcelFolder
                AddPath sSavePath
                oBookOrDocument.SaveCopyAs sSavePath & oBookOrDocument.Name
            End If
            If .CreateAddIn Then
                sSavePath = sFilePath & .AddInFolder
                AddPath sSavePath
                oBookOrDocument.SaveAs sSavePath & Replace(oBookOrDocument.Name, ".xlsm", ".xlam"), xlOpenXMLAddIn
            End If
            If .ExportRibbonX Then
                'Extract ribbonX from file container and store in Ribbon subfolder
                sSavePath = sFilePath & .RibbonXFolder
                sFullFile = oBookOrDocument.FullName
                AddPath sSavePath
                oBookOrDocument.Close False
                ExtractRibbonX sFullFile, sSavePath & "customUI.xml"
            End If
            MsgBox "Done exporting " & sFullFile, vbInformation + vbOKOnly, GCSAPPNAME
        Else
            MsgBox "Export cancelled", vbInformation + vbOKOnly, GCSAPPNAME
        End If

    End With
TidyUp:
    Exit Sub
LocErr:
    MsgBox "Error, code:" & Err.Number & "Description: " & Err.Description
    Stop
    Resume    'Next
End Sub

Sub ExportThisOne(vbc As Object, sFilePath As String, sAddition As String, sSubFolder As String)
    Dim sFullPath As String
    Dim sFileName As String
    Dim iCount As Integer
    If Right(sFilePath, 1) <> Application.PathSeparator Then sFilePath = sFilePath & Application.PathSeparator
    sFullPath = sFilePath & sSubFolder
    AddPath sFullPath
    sFileName = sFullPath & vbc.Name & sAddition
    vbc.Export sFileName
End Sub

Sub AddPath(ByRef sPath As String)
    
    Dim sTemp As String
    Dim iPos As Integer
    Dim sCurdir As String
    sCurdir = CurDir
    ChDrive sPath
    If Right(sPath, 1) <> Application.PathSeparator Then
        sPath = sPath & Application.PathSeparator
    End If
    If Dir(sPath, vbDirectory) <> "" Then Exit Sub
    iPos = 3
    While iPos > 0
        iPos = InStr(iPos + 1, sPath, Application.PathSeparator)
        sTemp = Left(sPath, iPos)
        If sTemp = "" Then Exit Sub
        If Dir(sTemp, vbDirectory) = "" Then
            MkDir sTemp
        Else
            ChDir sTemp
        End If
    Wend
    If sCurdir <> CurDir Then
        ChDrive sCurdir
        ChDir sCurdir
    End If
    Exit Sub
End Sub

Public Sub EmptyFolder(sFolder2Clear As String)
    Dim oFSO As Object

    If Right(sFolder2Clear, 1) = Application.PathSeparator Then
        sFolder2Clear = Left(sFolder2Clear, Len(sFolder2Clear) - 1)
    End If

    'Create FSO Object
    Set oFSO = CreateObject("Scripting.FileSystemObject")

    'Check Specified Folder exists or not
    If oFSO.FolderExists(sFolder2Clear) Then
        sFolder2Clear = sFolder2Clear & Application.PathSeparator
        'Delete All Files
        oFSO.DeleteFile sFolder2Clear & "*.*", True

        'Delete All Subfolders
        oFSO.DeleteFolder sFolder2Clear & "*.*", True

    End If
End Sub
