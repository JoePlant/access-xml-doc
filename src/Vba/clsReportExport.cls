VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsReportExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private pReport As Report
Private xml As clsXmlWriter
Private dir As String
Private bypassFile As Boolean
Private ReportName As String

Property Set XmlWriter(writer As clsXmlWriter)
    Set xml = writer
    bypassFile = Not (xml Is Nothing)
End Property

Property Let Name(sName As String)
    ReportName = sName
End Property

Property Get GetFileName() As String
    If (pReport Is Nothing) Then
        GetFileName = "Report." & ReportName & ".xml"
    Else
        GetFileName = "Report." & pReport.Name & ".xml"
    End If
End Property

Property Let Directory(d As String)
    dir = d
End Property

Property Set Report(rpt As Report)
    Set pReport = rpt
End Property

Function ExportAsXml() As String
    
    If (xml Is Nothing) Then
        Set xml = New clsXmlWriter
    End If
    
    xml.WriteElementStart Name:="Report"
    If (pReport Is Nothing) Then
        xml.WriteAttribute Name:="name", Value:=ReportName
        xml.WriteAttribute Name:="fileName", Value:=GetFileName()
    Else
        xml.WriteAttribute Name:="name", Value:=pReport.Name
    End If
    
    xml.WriteElementEnd Name:="Report"
    
    If Not (pReport Is Nothing) Then
    
        Call xml.WritePropertiesAsXml(pReport.Properties)
        
        Call WriteGroup(0)
        Call WriteGroup(1)
        Call WriteGroup(2)
        Call WriteGroup(3)
        Call WriteGroup(4)
        Call WriteGroup(5)
        Call WriteGroup(6)
        Call WriteGroup(7)
        
        Call WriteSection(acPageHeader)
        Call WriteSection(acHeader)
        Call WriteSection(acGroupLevel1Header)
        Call WriteSection(acGroupLevel2Header)
        Call WriteSection(acDetail)
        Call WriteSection(acGroupLevel2Footer)
        Call WriteSection(acGroupLevel1Footer)
        Call WriteSection(acFooter)
        Call WriteSection(acPageFooter)
        
        Call WriteModule
    End If
    
    xml.WriteEndElement Name:="Report"
    ExportAsXml = xml.XmlOutput
    
    If Not (bypassFile) Then
        Dim fileName As String
        fileName = dir & "\" & GetFileName
        Call SaveToFile(fileName, xml.XmlOutput)
        Debug.Print "Output report: " & pReport.Name & " to file: " & fileName
    End If
    
    
End Function

Private Sub WriteSection(sectionId As Integer)
On Error Resume Next

    Dim sect As Section
    Set sect = pReport.Section(sectionId)
    
    If Not (sect Is Nothing) Then
        Call xml.WriteElementStart("Section")
        Call xml.WriteAttribute("name", sect.Name)
        Call xml.WriteElementEnd("Section")
        Call xml.WritePropertiesAsXml(sect.Properties)
        Call WriteControls(sect.Controls)

        Call xml.WriteEndElement("Section")
    End If
    
End Sub

Private Sub WriteGroup(level As Integer)
On Error Resume Next

    Dim grp As groupLevel
    Set grp = pReport.groupLevel(level)
    
    If Not (grp Is Nothing) Then
        Call xml.WriteElementStart("GroupLevel")
        Call xml.WriteAttribute("level", "" & level & "")
        Call xml.WriteElementEnd("GroupLevel")
        Call xml.WritePropertiesAsXml(grp.Properties)

        Call xml.WriteEndElement("GroupLevel")
    End If
    
End Sub

Private Sub SaveToFile(fileName As String, xml As String)
    Dim intFile As Integer

    intFile = FreeFile()
    Open fileName For Output As #intFile
    Print #intFile, xml
    Close #intFile
End Sub

Private Sub WriteModule()

    Dim m As Module
    
    If (pReport.HasModule) Then
        Set m = pReport.Module
        
        Dim mx As clsModuleExport
        
        Set mx = New clsModuleExport
        
        Set mx.Module = m
        mx.Directory = dir
        Set mx.XmlWriter = xml
        Dim mxml As String
        
        mxml = mx.ExportAsXml
    End If

End Sub

Private Sub WriteControls(cntrls As Variant)

Dim ctrl As Control
Dim ctrlExport As clsControlExport

    Set ctrlExport = New clsControlExport
    Set ctrlExport.XmlWriter = xml
    
    Call xml.WriteElementStart("Controls")
    Call xml.WriteAttribute("count", cntrls.Count)
    Call xml.WriteElementEnd("Controls")
    
    For Each ctrl In cntrls
        Set ctrlExport.Control = ctrl
        Call ctrlExport.WriteXml
    Next
    
    Call xml.WriteEndElement("Controls")
    
End Sub

Private Sub Class_Initialize()
    dir = Environ("TEMP")
End Sub

