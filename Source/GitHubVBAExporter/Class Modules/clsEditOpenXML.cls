VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEditOpenXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mbCreateBackup As Boolean
Private mvSourceFile As Variant
Private msSheet2Change As String
Private msSheetId As String
Private msSheetFileName As String
Private mbAddedZip As Boolean
Private mvXMLFolderRoot As Variant
Private mvxmlfolder As Variant
'Private mvXMLFolderCustomUI As Variant

Public Enum XMLFolder
    'Date Created : 5/12/2009 21:34
    'Author       : Ken Puls (www.excelguru.ca)
    'Macro Purpose: Constants for XML Containers
    XMLFolder_root = 1
    XMLFolder_rels = 2
    XMLFolder_xl = 3
    XMLFolder_customUI = 4
    XMLFolder_docProps = 5
End Enum

Public Function GetXMLFromFile(sFileName As String, sXMLFolder As XMLFolder) As String
'-------------------------------------------------------------------------
' Procedure : GetXMLFromFile
' Company   : JKP Application Development Services (c)
' Author    : Jan Karel Pieterse
' Created   : 6-5-2009
' Purpose   : Gets the XML code from the foldername\filename
'-------------------------------------------------------------------------
    Dim oXMLDoc As MSXML2.DOMDocument
    If Len(XMLFolder(sXMLFolder)) = 0 Then
        GetXMLFromFile = ""
    Else
        Set oXMLDoc = New MSXML2.DOMDocument
        oXMLDoc.Load XMLFolder(sXMLFolder) & sFileName
        GetXMLFromFile = oXMLDoc.XML
        Set oXMLDoc = Nothing
    End If
End Function

Public Sub WriteXML2File(sXML As String, sFileName As String, sXMLFolder As XMLFolder)
'-------------------------------------------------------------------------
' Procedure : WriteXML2File
' Company   : JKP Application Development Services (c)
' Author    : Jan Karel Pieterse
' Created   : 6-5-2009
' Purpose   : Writes sXML to sFileName
'             Modified by Ken Puls 2009-05-12
'             Adjusted to add ability to write to customUI container
'-------------------------------------------------------------------------
    Dim oXMLDoc As MSXML2.DOMDocument
    Set oXMLDoc = New MSXML2.DOMDocument

    'If attempting to write a customUI component, test to see if one exists
    
    'Should probably test the .rels file to see if the CustomUI relationship exists...
    If sXMLFolder = XMLFolder_customUI Then
        If Not FolderExists(XMLFolder(XMLFolder_customUI)) Then

            MkDir XMLFolder(XMLFolder_customUI)
            'Write the XML to the file
            oXMLDoc.loadXML sXML
            oXMLDoc.Save XMLFolder(sXMLFolder) & sFileName
            'CustomUI has not been created yet.  Rels file needs to be adjusted
            AddCustomUIToRels
        End If
    End If

    'Write the XML to the file
    oXMLDoc.loadXML sXML
    oXMLDoc.Save XMLFolder(sXMLFolder) & sFileName
End Sub

Public Sub AddCustomUIToRels()
'Date Created : 5/14/2009 23:29
'Author       : Ken Puls (www.excelguru.ca)
'Macro Purpose: Add the customUI relationship to the rels file

    Dim oXMLDoc As MSXML2.DOMDocument
'    Dim oXMLElement As MSXML2.IXMLDOMElement
    Dim oXMLElement As MSXML2.IXMLDOMNode
    Dim oXMLAttrib As MSXML2.IXMLDOMAttribute
    Dim oNamedNodeMap As MSXML2.IXMLDOMNamedNodeMap
    Dim oXMLRelsList As MSXML2.IXMLDOMNodeList
    'Create a new XML document
    Set oXMLDoc = New MSXML2.DOMDocument
    'Attach to the root element of the .rels file
    oXMLDoc.Load XMLFolder(XMLFolder_rels) & ".rels"

    'Create a new relationship element in the .rels file
    Set oXMLElement = oXMLDoc.createNode(1, "Relationship", "http://schemas.openxmlformats.org/package/2006/relationships")
    Set oNamedNodeMap = oXMLElement.Attributes
    
    'Create ID attribute for the element
    Set oXMLAttrib = oXMLDoc.createAttribute("Id")
    oXMLAttrib.nodeValue = "cuID"
    oNamedNodeMap.setNamedItem oXMLAttrib

    'Create Type attribute for the element
'    Set oXMLAttrib = oXMLDoc.createAttribute("Type")
'    oXMLAttrib.nodeValue = "http://shemas.microsoft.com/office/2006/relationships/ui/extensibility"
    Set oXMLAttrib = oXMLDoc.createAttribute("Type")
    oXMLAttrib.nodeValue = "http://schemas.microsoft.com/office/2006/relationships/ui/extensibility"
    oNamedNodeMap.setNamedItem oXMLAttrib

    'Create Target element for the attribute
'    Set oXMLAttrib = oXMLDoc.createAttribute("Target")
'    oXMLAttrib.nodeValue = "customUI/customUI.xml"
'    oXMLElement.setAttributeNode oXMLAttrib
    Set oXMLAttrib = oXMLDoc.createAttribute("Target")
    oXMLAttrib.nodeValue = "customUI/customUI.xml"
    oNamedNodeMap.setNamedItem oXMLAttrib
    
    'Now insert the new node at the proper location
    Set oXMLRelsList = oXMLDoc.selectNodes("/Relationships")
    oXMLRelsList.Item(0).appendChild oXMLElement
    'Save the .rels file
    oXMLDoc.Save XMLFolder(XMLFolder_rels) & ".rels"

    Set oXMLAttrib = Nothing
    Set oXMLElement = Nothing
    Set oXMLDoc = Nothing
End Sub



Private Function GetSheetIdFromSheetName(sSheetName) As String
'-------------------------------------------------------------------------
' Procedure : GetSheetIdFromSheetName
' Company   : JKP Application Development Services (c)
' Author    : Jan Karel Pieterse
' Created   : 6-5-2009
' Purpose   : Finds out what the SheetId of sSheetname is
'             by reading Workbook.xml
'-------------------------------------------------------------------------
    Dim oXMLDoc As MSXML2.DOMDocument
    Dim oxmlNode As MSXML2.IXMLDOMNode
    Dim oXMLChildNode As MSXML2.IXMLDOMNode
    Dim oXMLTemp As MSXML2.IXMLDOMNode
    If XMLFolder(XMLFolder_xl) <> "" And Sheet2Change <> "" Then
        Set oXMLDoc = New MSXML2.DOMDocument
        oXMLDoc.Load XMLFolder(XMLFolder_xl) & "workbook.xml"
        For Each oxmlNode In oXMLDoc.ChildNodes
            For Each oXMLChildNode In oxmlNode.ChildNodes
                If oXMLChildNode.baseName = "sheets" Then
                    For Each oXMLTemp In oXMLChildNode.ChildNodes
                        If oXMLTemp.Attributes.getNamedItem("name").nodeValue = sSheetName Then
                            GetSheetIdFromSheetName = oXMLTemp.Attributes.getNamedItem("r:id").nodeValue
                            Exit Function
                        End If
                    Next
                End If
            Next
        Next
    End If
End Function

Public Function GetSheetFileNameFromId(sSheetId As String) As String
'-------------------------------------------------------------------------
' Procedure : GetSheetFileNameFromId
' Company   : JKP Application Development Services (c)
' Author    : Jan Karel Pieterse
' Created   : 6-5-2009
' Purpose   : Fetches the name of the xml file belonging to the sheet with id SheetId.
'-------------------------------------------------------------------------
    Dim oXMLDoc As MSXML2.DOMDocument
    Dim oxmlNode As MSXML2.IXMLDOMNode
    Dim oXMLChildNode As MSXML2.IXMLDOMNode
    If XMLFolder(XMLFolder_xl) <> "" And Sheet2Change <> "" Then
        Set oXMLDoc = New MSXML2.DOMDocument
        oXMLDoc.Load XMLFolder(XMLFolder_xl) & "_rels\workbook.xml.rels"
        For Each oxmlNode In oXMLDoc.ChildNodes
            For Each oXMLChildNode In oxmlNode.ChildNodes
                If oXMLChildNode.Attributes.getNamedItem("Id").nodeValue = sSheetId Then
                    GetSheetFileNameFromId = oXMLChildNode.Attributes.getNamedItem("Target").nodeValue
                    Exit Function
                End If
            Next
        Next
    End If
End Function

Private Function GetSheetNameFromId(sId As String) As String
'-------------------------------------------------------------------------
' Procedure : GetSheetNameFromId
' Company   : JKP Application Development Services (c)
' Author    : Jan Karel Pieterse
' Created   : 6-5-2009
' Purpose   : Returns the sheetname belonging to a sheetId
'-------------------------------------------------------------------------
    Dim oXMLDoc As MSXML2.DOMDocument
    Dim oxmlNode As MSXML2.IXMLDOMNode
    Dim oXMLChildNode As MSXML2.IXMLDOMNode
    Dim oXMLChildChildNode As MSXML2.IXMLDOMNode
    If mvxmlfolder(XMLFolder_xl) <> "" Then
        Set oXMLDoc = New MSXML2.DOMDocument
        oXMLDoc.Load XMLFolder(XMLFolder_xl) & "workbook.xml"
        For Each oxmlNode In oXMLDoc.ChildNodes
            For Each oXMLChildNode In oxmlNode.ChildNodes
                If oXMLChildNode.nodeName = "sheets" Then
                    For Each oXMLChildChildNode In oXMLChildNode.ChildNodes
                        If oXMLChildChildNode.Attributes.getNamedItem("r:id").nodeValue = "rId" & Val(sId) + 1 Then
                            GetSheetNameFromId = oXMLChildChildNode.Attributes.getNamedItem("name").nodeValue
                            'Got it, get out
                            Exit Function
                        End If
                    Next
                    'get out here, no point in doing the rest
                    Exit Function
                End If
            Next
        Next
    End If
End Function

Public Sub ZipAllFilesInFolder()
'-------------------------------------------------------------------------
' Procedure : ZipAllFilesInFolder
' Company   : JKP Application Development Services (c)
' Author    : Jan Karel Pieterse
' Created   : 6-5-2009
' Purpose   : Zips all files in a folder (including subfolders) whilst retaining the folder structure
'-------------------------------------------------------------------------
'Courtesy www.rondebruin.nl
    Dim oShellApp As Object
    Dim sDate As String
    Dim sDefPath As String
    Dim vFileNameZip As Variant
    Dim FSO As Object
    Dim lFileCt As Long
    Set FSO = CreateObject("scripting.filesystemobject")

    'To ensure a unique filename,
    'append date and time to the name of the current file

    sDate = Format(Now, " dd-mmm-yy h-mm-ss")
    vFileNameZip = SourceFile & sDate & ".zip"

    'Create empty Zip File
    NewZip vFileNameZip

    Set oShellApp = CreateObject("Shell.Application")

    'Count how many items are in the "old" folder
    lFileCt = oShellApp.Namespace(FolderName & "Unzipped " & FileName & Application.PathSeparator).items.Count

    'Copy the files to the compressed folder
    oShellApp.Namespace(vFileNameZip).CopyHere oShellApp.Namespace(FolderName & "Unzipped " & FileName & Application.PathSeparator).items

    'Keep script waiting until we have same # of files in the new folder
    On Error Resume Next
    Do Until oShellApp.Namespace(vFileNameZip).items.Count = lFileCt
        Application.Wait (Now + TimeValue("0:00:01"))
    Loop
    DoEvents

    'Remove original file
    Kill SourceFile

    'Rename new zipped file to same name as original file (with .zip appended)
    Name vFileNameZip As SourceFile
    On Error Resume Next

    'Now remove old folder, just in case something went haywire
    FSO.DeleteFolder FolderName & "Unzipped " & FileName, True
    On Error GoTo 0

    Set oShellApp = Nothing
End Sub

Public Sub UnzipFile()
'-------------------------------------------------------------------------
' Procedure : UnzipFile
' Company   : JKP Application Development Services (c)
' Author    : Jan Karel Pieterse
' Created   : 6-5-2009
' Purpose   : Unzips all files in a zip file to a designated folder
'             Modified by Ken Puls 2009-05-12
'             Adjusted to record customUI portion
'-------------------------------------------------------------------------
'Courtesy www.rondebruin.nl
    Dim FSO As Object
    Dim oShellApp As Object
    Set FSO = CreateObject("scripting.filesystemobject")

    'Derive the folder to unzip to from the location of the sourcefile
    XMLFolderRoot = FolderName

    'A dedicated unzip folder will be created in the same folder as the sourcefile,
    'called ..\Unzipped Filename\
    If Right(XMLFolderRoot, 1) <> Application.PathSeparator Then
        XMLFolderRoot = XMLFolderRoot & "\UnZipped " & FileName & Application.PathSeparator
    Else
        XMLFolderRoot = XMLFolderRoot & "UnZipped " & FileName & Application.PathSeparator
    End If
    On Error Resume Next
    'Remove all previous existing folders
    FSO.DeleteFolder XMLFolderRoot & "*", True
    Kill XMLFolderRoot & "*.*"

    'Create normal folder
    If FolderExists(XMLFolderRoot) = False Then
        MkDir XMLFolderRoot
    End If

    Set oShellApp = CreateObject("Shell.Application")
    'Copy the files in the newly created folder
    oShellApp.Namespace(XMLFolderRoot).CopyHere oShellApp.Namespace(SourceFile).items

    On Error Resume Next
    'Clean up temp folder
    FSO.DeleteFolder Environ("Temp") & "\Temporary Directory*", True

    Set oShellApp = Nothing
    Set FSO = Nothing
    Exit Sub
End Sub

Sub NewZip(sPath)
'Courtesy www.rondebruin.nl
'Create empty Zip File
'Changed by keepITcool Dec-12-2005
    If Len(Dir(sPath)) > 0 Then Kill sPath
    Open sPath For Output As #1
    Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
    Close #1
End Sub

Public Property Get CreateBackup() As Boolean
    CreateBackup = mbCreateBackup
End Property

Public Property Let CreateBackup(ByVal bCreateBackup As Boolean)
    mbCreateBackup = bCreateBackup
End Property

Private Sub Class_Initialize()
'Set defaults
    CreateBackup = True
End Sub

Public Property Get SourceFile() As Variant
    SourceFile = mvSourceFile
End Property

Public Property Let SourceFile(ByVal vSourceFile As Variant)
    mvSourceFile = vSourceFile
    If CreateBackup Then
        If Len(Dir(vSourceFile & "(backup)")) > 0 Then
            Kill vSourceFile & "(backup)"
        End If
        FileCopy vSourceFile, vSourceFile & "(backup)"
    End If
    If Not vSourceFile Like "*.zip" Then
        Name vSourceFile As vSourceFile & ".zip"
        mvSourceFile = mvSourceFile & ".zip"
        AddedZip = True
    End If
End Property

Public Property Get FolderName() As Variant
    FolderName = Mid(SourceFile, 1, InStrRev(SourceFile, Application.PathSeparator))
End Property

Public Property Get FileName() As Variant
    If SourceFile <> "" Then
        FileName = Mid(SourceFile, InStrRev(SourceFile, Application.PathSeparator) + 1, Len(SourceFile))
    End If
End Property



'Public Property Get xmlfolder(XMLFolder_xl)() As Variant
'    xmlfolder(XMLFolder_xl) = mvxmlfolder(XMLFolder_xl)
'End Property
'
'Public Property Let xmlfolder(XMLFolder_xl)(ByVal vxmlfolder(XMLFolder_xl) As Variant)
'    mvxmlfolder(XMLFolder_xl) = vxmlfolder(XMLFolder_xl)
'End Property
'
'Public Property Get XMLFolderCustomUI() As Variant
''Date Created : 5/12/2009 21:18
''Author       : Ken Puls (www.excelguru.ca)
''Macro Purpose: Retrieve customUI folder
'    XMLFolderCustomUI = mvXMLFolderCustomUI
'End Property
'
'Public Property Let XMLFolderCustomUI(ByVal vXMLFolderCustomUI As Variant)
''Date Created : 5/12/2009 21:18
''Author       : Ken Puls (www.excelguru.ca)
''Macro Purpose: Save customUI folder
'    mvXMLFolderCustomUI = vXMLFolderCustomUI
'End Property

Public Property Get XMLFolder(sXMLFolder As XMLFolder) As String
    Select Case sXMLFolder
        Case Is = XMLFolder_root
            XMLFolder = mvXMLFolderRoot
        Case Is = XMLFolder_customUI
            XMLFolder = mvXMLFolderRoot & "customUI" & Application.PathSeparator
        Case Is = XMLFolder_docProps
            XMLFolder = mvXMLFolderRoot & "docProps" & Application.PathSeparator
        Case Is = XMLFolder_rels
            XMLFolder = mvXMLFolderRoot & "_rels" & Application.PathSeparator
        Case Is = XMLFolder_xl
            XMLFolder = mvXMLFolderRoot & "xl" & Application.PathSeparator
    End Select
End Property

Public Property Get Sheet2Change() As String
    Sheet2Change = msSheet2Change
End Property

Public Property Let Sheet2Change(ByVal sSheet2Change As String)
    msSheet2Change = sSheet2Change
    SheetId = GetSheetIdFromSheetName(sSheet2Change)
    If SheetId <> "" Then
        SheetFileName = GetSheetFileNameFromId(SheetId)
    End If
End Property

Public Property Get SheetId() As String
    SheetId = msSheetId
End Property

Public Property Let SheetId(ByVal sSheetId As String)
    msSheetId = sSheetId
End Property

Public Property Get SheetFileName() As String
    SheetFileName = msSheetFileName
End Property

Public Property Let SheetFileName(ByVal sSheetFileName As String)
    msSheetFileName = sSheetFileName
End Property

Private Property Get AddedZip() As Boolean
    AddedZip = mbAddedZip
End Property

Private Property Let AddedZip(ByVal bAddedZip As Boolean)
    mbAddedZip = bAddedZip
End Property

Private Sub Class_Terminate()
    Dim FSO As Object
    If AddedZip Then
        'Remove .zip from file again
        Name SourceFile As Left(SourceFile, Len(SourceFile) - 4)
    End If
    'Remove zip folder
    On Error Resume Next
    FSO.DeleteFolder XMLFolderRoot, True
End Sub

Private Property Get XMLFolderRoot() As Variant
    XMLFolderRoot = mvXMLFolderRoot
End Property

Private Property Let XMLFolderRoot(ByVal vXMLFolderRoot As Variant)
    mvXMLFolderRoot = vXMLFolderRoot
End Property
