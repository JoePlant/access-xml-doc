VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsModuleExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private pModule As Module
Private xml As clsXmlWriter
Private dir As String
Private bypassFile As Boolean
Private moduleName As String

Property Let Directory(d As String)
    dir = d
End Property

Property Let Name(sName As String)
    moduleName = sName
End Property

Property Get GetFileName() As String
    If (pModule Is Nothing) Then
        GetFileName = "Module." & moduleName & ".xml"
    Else
        GetFileName = "Module." & pModule.Name & ".xml"
    End If
End Property

Property Set Module(m As Module)
    Set pModule = m
End Property

Property Set XmlWriter(writer As clsXmlWriter)
    Set xml = writer
    bypassFile = Not (xml Is Nothing)
End Property

Function ExportAsXml() As String
    
    If (xml Is Nothing) Then
        Set xml = New clsXmlWriter
    End If
    xml.WriteElementStart Name:="Module"
    If (pModule Is Nothing) Then
        xml.WriteAttribute Name:="name", Value:=moduleName
        xml.WriteAttribute Name:="fileName", Value:=GetFileName()
    Else
        xml.WriteAttribute Name:="name", Value:=pModule.Name
        xml.WriteAttribute Name:="type", Value:=ModuleType()
        xml.WriteAttribute Name:="lines", Value:=pModule.CountOfLines
    End If
    xml.WriteElementEnd Name:="Module"

    If Not (pModule Is Nothing) Then
    
        Dim startSections() As Long
        
        startSections = Me.ArraySections
        
        Call WriteDeclaration
            
        Dim lineNo As Long
        Dim i As Integer
        For i = 0 To UBound(startSections)
            lineNo = startSections(i)
            Call WriteSection(lineNo)
        Next i
        
        '  Call WriteSource
        
    End If
    
    xml.WriteEndElement Name:="Module"
    ExportAsXml = xml.XmlOutput
    
    If Not (bypassFile) Then
        Dim fileName As String
        fileName = dir & "\" & GetFileName()
        
        Call SaveToFile(fileName, xml.XmlOutput)
        Debug.Print "Output module: " & pModule.Name & " to file: " & fileName
    End If

    
End Function

Private Sub SaveToFile(fileName As String, xml As String)
    Dim intFile As Integer

    intFile = FreeFile()
    Open fileName For Output As #intFile
    Print #intFile, xml
    Close #intFile
End Sub


Private Sub WriteModuleLines()
    Dim i As Integer
    Dim line As String

    For i = 1 To pModule.CountOfLines
        line = pModule.Lines(i, 1)
        Call xml.WriteElementStart("Line")
        Call xml.WriteAttribute("no", CStr(i))
        Call xml.WriteElementEnd("Line", False)
        Call xml.WriteText(line, False)
        Call xml.WriteEndElement("Line")
    Next i
    
End Sub

Private Sub WriteDeclaration()
    Dim i As Integer
    Dim line As String

    Call xml.WriteElementStart("Declarations")
    Call xml.WriteElementEnd("Declarations", True)
    For i = 1 To pModule.CountOfDeclarationLines
        line = pModule.Lines(i, 1)
        Call xml.WriteText(line, True)
    Next i
    Call xml.WriteEndElement("Declarations")

End Sub

Private Sub WriteSection(lineNo As Long)
    Dim i As Long
    Dim line As String

    Dim sProcName As String
    Dim sType As String
    Dim lngR As Long
    
    sProcName = pModule.ProcOfLine(lineNo, lngR)
    sType = ProcType(lngR)

    Call xml.WriteElementStart("Section")
    Call xml.WriteAttribute("name", sProcName)
    Call xml.WriteAttribute("lineNo", CStr(lineNo))
    Call xml.WriteAttribute("type", sType)
    Call xml.WriteElementEnd("Section", True)
    For i = lineNo To pModule.CountOfLines
        If (sProcName <> pModule.ProcOfLine(i, lngR)) Then
            Exit For
        End If
        line = pModule.Lines(i, 1)
        Call xml.WriteText(line, True)
    Next i
    Call xml.WriteEndElement("Section")

End Sub

Private Function ProcType(lType As Long) As String
    ProcType = "Unknown:" & lType
    Select Case lType
        Case 0: ProcType = "Sub or Function" 'vbext_pk_Proc
        Case 1: ProcType = "Property Let" ' vbext_pk_Let
        Case 2: ProcType = "Property Set" ' vbext_pk_Set
        Case 3: ProcType = "Property Get" ' vbext_pk_Get
    End Select
End Function

Private Sub WriteSource()
    Dim i As Integer
    Dim line As String

    Call xml.WriteElementStart("Source")
    Call xml.WriteElementEnd("Source", True)
    For i = 1 To pModule.CountOfLines
        line = pModule.Lines(i, 1)
        Call xml.WriteText(line, True)
    Next i
    Call xml.WriteEndElement("Source")

End Sub

Public Property Get ArraySections() As Long()

    Dim lngCount As Long
    Dim lngCountDecl As Long
    Dim lngI As Long
    Dim strProcName As String
    Dim alngLines() As Long
    Dim intI As Integer
    Dim strMsg As String
    Dim lngR As Long

    ' Count lines in module.
    lngCount = pModule.CountOfLines
    
    ' Count lines in Declaration section in module.
    lngCountDecl = pModule.CountOfDeclarationLines
    
    ' Determine name of first procedure.
    strProcName = pModule.ProcOfLine(lngCountDecl + 1, lngR)
    
    ' Initialize counter variable.
    intI = 0
    
    ' Redimension array.
    ReDim Preserve alngLines(intI)
    
    ' Store name of first procedure in array.
    alngLines(intI) = lngCountDecl + 1
    
    ' Determine procedure name for each line after declarations.
    For lngI = lngCountDecl + 1 To lngCount
        ' Compare procedure name with ProcOfLine property value.
        If strProcName <> pModule.ProcOfLine(lngI, lngR) Then
            ' Increment counter.
            intI = intI + 1
            strProcName = pModule.ProcOfLine(lngI, lngR)
            ReDim Preserve alngLines(intI)
            ' Assign unique procedure names to array.
            alngLines(intI) = lngI
        End If
    Next lngI
    
    ArraySections = alngLines

End Property

Public Property Get ModuleType() As String
    ModuleType = "unknown"
    Select Case pModule.Type
        Case acClassModule: ModuleType = "Class Module"
        Case acStandardModule: ModuleType = "Standard Module"
    End Select
End Property

Private Sub Class_Initialize()
    dir = Environ("TEMP")
End Sub
