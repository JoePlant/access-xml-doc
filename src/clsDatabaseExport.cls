VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDatabaseExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private pDatabase As CurrentProject
Private xml As clsXmlWriter
Private dir As String

Property Let Directory(d As String)
    dir = d
End Property

Property Set Database(db As CurrentProject)
    Set pDatabase = db
End Property

Function ExportAsXml() As String
    
    Set xml = New clsXmlWriter
    
    xml.WriteElementStart Name:="Database"
    xml.WriteAttribute Name:="name", Value:=pDatabase.Name
       
    xml.WriteElementEnd Name:="Database"
    
    Call WriteProperties
    
    Call WriteForms
    Call WriteReports
    Call WriteModules
    
    xml.WriteEndElement Name:="Database"
    ExportAsXml = xml.XmlOutput
    
    Dim fileName As String
    fileName = dir & "\Database." & pDatabase.Name & ".xml"
    
    Call SaveToFile(fileName, xml.XmlOutput)
    
    Debug.Print "Output : " & pDatabase.Name & " to file: " & fileName
    
End Function

Private Sub SaveToFile(fileName As String, xml As String)
    Dim intFile As Integer

    intFile = FreeFile()
    Open fileName For Output As #intFile
    Print #intFile, xml
    Close #intFile
End Sub


Private Sub WriteForms()
    Dim accObj As AccessObject
    Dim strDoc As String
    
    Call xml.WriteElementStart("Forms")
    Call xml.WriteAttribute("count", pDatabase.AllForms.Count)
    Call xml.WriteElementEnd("Forms")
    
    For Each accObj In pDatabase.AllForms
        strDoc = accObj.Name
        
        Dim fx As clsFormExport
        Set fx = New clsFormExport
        fx.Name = accObj.Name
        'Set fx.Form = Nothing
        fx.Directory = dir
        Set fx.XmlWriter = xml
        
        Dim fxml As String
        fxml = fx.ExportAsXml
        
        DoEvents
       
    Next
    
    Call xml.WriteEndElement("Forms")
    
    Debug.Print "Forms count: " & pDatabase.AllForms.Count
End Sub

Private Sub WriteReports()
    Dim accObj As AccessObject
    Dim strDoc As String
    
    Call xml.WriteElementStart("Reports")
    Call xml.WriteAttribute("count", pDatabase.AllReports.Count)
    Call xml.WriteElementEnd("Reports")
    
    For Each accObj In pDatabase.AllReports
        strDoc = accObj.Name
        
        Dim rx As clsReportExport
        Set rx = New clsReportExport
        rx.Name = accObj.Name
        rx.Directory = dir
        Set rx.XmlWriter = xml
        
        Dim rxml As String
        rxml = rx.ExportAsXml
        
        DoEvents
        
    Next
    
    Call xml.WriteEndElement("Reports")
    
    Debug.Print "Reports count: " & pDatabase.AllReports.Count
End Sub


Private Sub WriteModules()
    Dim accObj As AccessObject
    Dim bWasOpen As Boolean
    Dim strDoc As String
    Dim m As Module
    
    Call xml.WriteElementStart("Modules")
    Call xml.WriteAttribute("count", pDatabase.AllModules.Count)
    Call xml.WriteElementEnd("Modules")
    
    For Each accObj In pDatabase.AllModules
        strDoc = accObj.Name
        
        Dim mx As clsModuleExport
        Set mx = New clsModuleExport
    
        mx.Name = accObj.Name
        mx.Directory = dir
        Set mx.XmlWriter = xml
        Dim mxml As String
    
        mxml = mx.ExportAsXml
    Next
    
    Call xml.WriteEndElement("Modules")
    
    Debug.Print "Modules count: " & pDatabase.AllModules.Count
    
End Sub

Private Sub WriteProperties()
    On Error Resume Next
    Dim prop As AccessObjectProperty

    Call xml.WriteElementStart("Properties")
    Call xml.WriteAttribute("count", pDatabase.Properties.Count)
    Call xml.WriteElementEnd("Properties")

    For Each prop In pDatabase.Properties
    
        Select Case prop.Name
            Case Else
                Call WriteProperty(prop)
        End Select
    Next
    
    Call xml.WriteEndElement("Properties")
    
End Sub

Private Sub WriteProperty(prop As AccessObjectProperty)
On Error Resume Next

Dim v As String
Dim n As String
Dim t As String

    n = prop.Name
    't = prop.Type
    v = prop.Value

    If (v = "") Then
    
    Else
    
        Call xml.WriteElementStart(Name:="Property")
        Call xml.WriteAttribute(Name:="name", Value:=n)
        Call xml.WriteElementEnd(Name:="Property", bCrLf:=False)
        Call xml.WriteText(Value:=v, bCrLf:=False)
        Call xml.WriteEndElement(Name:="Property", bPad:=False)
        
    End If
End Sub


Private Sub Class_Initialize()
    dir = Environ("TEMP")
End Sub

