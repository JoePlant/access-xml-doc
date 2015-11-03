VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFormExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private pFrm As Form
Private xml As clsXmlWriter
Private dir As String
Private bypassFile As Boolean
Private formName As String

Property Set XmlWriter(writer As clsXmlWriter)
    Set xml = writer
    bypassFile = Not (xml Is Nothing)
End Property

Property Let Name(sName As String)
    formName = sName
End Property

Property Get GetFileName() As String
    If (pFrm Is Nothing) Then
        GetFileName = "Form." & formName & ".xml"
    Else
        GetFileName = "Form." & pFrm.Name & ".xml"
    End If
End Property

Property Let Directory(d As String)
    dir = d
End Property

Property Set Form(frm As Form)
    Set pFrm = frm
End Property

Function ExportAsXml() As String
    
    If (xml Is Nothing) Then
        Set xml = New clsXmlWriter
    End If
    
    xml.WriteElementStart Name:="Form"
    If (pFrm Is Nothing) Then
        xml.WriteAttribute Name:="name", Value:=formName
        xml.WriteAttribute Name:="fileName", Value:=GetFileName()
    Else
        xml.WriteAttribute Name:="name", Value:=pFrm.Name
    End If
    
    xml.WriteElementEnd Name:="Form"
    
    If Not (pFrm Is Nothing) Then
    
        Call xml.WritePropertiesAsXml(pFrm.Properties)
        
        Call WriteSection(acDetail)
        Call WriteSection(acHeader)
        Call WriteSection(acFooter)
        Call WriteSection(acPageHeader)
        Call WriteSection(acPageFooter)
        
        Call WriteModule

    End If
    
    xml.WriteEndElement Name:="Form"
    ExportAsXml = xml.XmlOutput
    
    If Not (bypassFile) Then
        Dim fileName As String
        fileName = dir & "\" & GetFileName
        Call SaveToFile(fileName, xml.XmlOutput)
        Debug.Print "Output form: " & pFrm.Name & " to file: " & fileName
    End If
    
    
End Function

Private Sub WriteSection(sectionId As Integer)
On Error Resume Next

    Dim sect As Section
    Set sect = pFrm.Section(sectionId)
    
    If Not (sect Is Nothing) Then
        Call xml.WriteElementStart("Section")
        Call xml.WriteAttribute("name", sect.Name)
        Call xml.WriteElementEnd("Section")
        Call xml.WritePropertiesAsXml(sect.Properties)
        
        Call WriteControls(sect.Controls)

        Call xml.WriteEndElement("Section")
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
    
    If (pFrm.HasModule) Then
        Set m = pFrm.Module
        
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
