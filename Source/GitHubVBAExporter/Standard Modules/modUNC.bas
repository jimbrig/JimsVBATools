Attribute VB_Name = "modUNC"
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2008 VBnet, Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Const MAX_PATH As Long = 260
Private Const INVALID_HANDLE_VALUE As Long = -1
Private Const FILE_ATTRIBUTE_DIRECTORY As Long = &H10

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

#If VBA7 Then
    Private Declare PtrSafe Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As LongPtr
    Private Declare PtrSafe Function FindClose Lib "kernel32" (ByVal hFindFile As LongPtr) As Long
    Private Declare PtrSafe Function SetCurrentDirectoryA Lib "kernel32" (ByVal lpPathName As String) As Long
#Else
    Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
    Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
    Private Declare Function SetCurrentDirectoryA Lib "kernel32" (ByVal lpPathName As String) As Long
#End If

Function SetUNCPath(ByVal szPath As String) As Boolean
    Dim lReturn As Long
    lReturn = SetCurrentDirectoryA(szPath)
    If lReturn = 0 Then
        SetUNCPath = False
    Else
        SetUNCPath = True
    End If
End Function

Public Function FolderExists(ByVal sFolder As String) As Boolean

    #If VBA7 Then
        Dim hFile As LongPtr
    #Else
        Dim hFile As Long
    #End If
    Dim WFD As WIN32_FIND_DATA

    'remove training slash before verifying
    sFolder = UnQualifyPath(sFolder)

    'call the API pasing the folder
    hFile = FindFirstFile(sFolder, WFD)

    'if a valid file handle was returned,
    'and the directory attribute is set
    'the folder exists
    FolderExists = (hFile <> INVALID_HANDLE_VALUE) And _
                   (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY)

    'clean up
    Call FindClose(hFile)

End Function


Private Function UnQualifyPath(ByVal sFolder As String) As String

  'trim and remove any trailing slash
   sFolder = Trim$(sFolder)
   
   If Right$(sFolder, 1) = Application.PathSeparator Then
      UnQualifyPath = Left$(sFolder, Len(sFolder) - 1)
   Else
      UnQualifyPath = sFolder
   End If
   
End Function


