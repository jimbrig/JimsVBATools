Attribute VB_Name = "modExtractRibbonX"
Option Explicit

Public Sub ExtractRibbonX(sFullFile As String, sSaveFile As String)
'-------------------------------------------------------------------------
' Procedure : Demo
' Company   : JKP Application Development Services (c)
' Author    : Jan Karel Pieterse (www.jkp-ads.com)
' Created   : 06-05-2009
' Purpose   : Demonstrates getting something from an OpemXML file
'-------------------------------------------------------------------------
    Dim cEditOpenXML As clsEditOpenXML
    Dim sXML As String
    Dim oXMLDoc As MSXML2.DOMDocument

    Set cEditOpenXML = New clsEditOpenXML

    With cEditOpenXML
        .CreateBackup = False
        'Tell it which OpenXML file to process
        .SourceFile = sFullFile
        'Before you can access info in the file, it must be unzipped
        .UnzipFile

        'Get XML from the ribbonX file (Office 2007 compatible)
        sXML = .GetXMLFromFile("customUI.xml", XMLFolder_customUI)
        If Len(sXML) > 0 Then
            'Change the xml of the sheet here
            Set oXMLDoc = New DOMDocument
            oXMLDoc.loadXML sXML
            oXMLDoc.Save sSaveFile
        End If
        'RibbonX for Office 2010 and up
        sXML = .GetXMLFromFile("customUI14.xml", XMLFolder_customUI)
        If Len(sXML) > 0 Then
            'Change the xml of the sheet here
            Set oXMLDoc = New DOMDocument
            oXMLDoc.loadXML sXML
            oXMLDoc.Save Replace(sSaveFile, ".xml", "14.xml")
        End If
    End With

    'Only when you let the class go out of scope the zip file's .zip extension is removed
    'in the terminate event of the class.
    'Then the OpenXML file has its original filename back.
    Set cEditOpenXML = Nothing
End Sub
