VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsControlExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private xml As clsXmlWriter
Private ctrl As Control
Private useRef As Boolean

Property Set XmlWriter(writer As clsXmlWriter)
    Set xml = writer
End Property

Property Set Control(c As Control)
    Set ctrl = c
End Property

Property Let UseReference(bUseRef As Boolean)
    useRef = True
End Property

Public Sub WriteXml()

    If (useRef) Then
        Call WriteRefXml
    Else
        Call WriteControlXml
    End If

End Sub

Private Sub WriteRefXml()
    xml.WriteElementStart Name:="ControlRef"
    xml.WriteAttribute Name:="name", Value:=ctrl.Name
    xml.WriteAttribute Name:="type-id", Value:=ctrl.ControlType
    xml.WriteAttribute Name:="type-name", Value:=Me.ControlType()
    
    xml.WriteElementClose Name:="ControlRef"
End Sub

Private Sub WriteControlXml()
    xml.WriteElementStart Name:="Control"
    xml.WriteAttribute Name:="name", Value:=ctrl.Name
    xml.WriteAttribute Name:="type-id", Value:=ctrl.ControlType
    xml.WriteAttribute Name:="type-name", Value:=Me.ControlType()
    xml.WriteAttribute Name:="left", Value:=ctrl.Left
    xml.WriteAttribute Name:="width", Value:=ctrl.Width
    xml.WriteAttribute Name:="top", Value:=ctrl.Top
    xml.WriteAttribute Name:="height", Value:=ctrl.Height
    
    xml.WriteElementEnd Name:="Control"
        
    Call xml.WritePropertiesAsXml(ctrl.Properties)
    'Call WriteProperties
        
    Dim inner As Variant
    Set inner = ChildControls(ctrl)
    
    If (Not inner Is Nothing) Then
        If (inner.Count > 0) Then
            Call xml.WriteElementStart(Name:="Controls")
            Call xml.WriteAttribute("count", inner.Count)
            Call xml.WriteElementEnd(Name:="Controls")
            
            Dim c As Control
            Dim innerCtrl As clsControlExport
            
            Dim i, num As Integer
            num = inner.Count
            Set innerCtrl = New clsControlExport
            Set innerCtrl.XmlWriter = xml
            innerCtrl.UseReference = True
            Dim innerC As Control
        
            For i = 0 To num - 1
                Set innerC = inner(i)
                Set innerCtrl.Control = innerC
           
                Call innerCtrl.WriteXml
            Next
            xml.WriteEndElement Name:="Controls"
        End If
    End If
    
    xml.WriteEndElement Name:="Control"

End Sub

Private Function ChildControls(ByRef c As Object) As Object
    On Error Resume Next
    Set ChildControls = Nothing
    Set ChildControls = c.Controls
    If Err.Number = 438 Then
        Err.Clear
    End If
End Function

Public Function ControlType() As String
    Dim ct As String
    
    
    Select Case ctrl.ControlType
        Case acLabel:
            ct = "Label"
        Case acRectangle:
            ct = "Rectangle"
        Case acLine:
            ct = "Line"
        Case acImage:
            ct = "Image"
        Case acCommandButton:
            ct = "Command Button"
        Case acOptionButton:
            ct = "Option button"
        Case acCheckBox:
            ct = "Check box"
        Case acOptionGroup:
            ct = "Option group"
        Case acBoundObjectFrame:
            ct = "Bound object frame"
        Case acTextBox:
            ct = "Text Box"
        Case acListBox:
            ct = "List box"
        Case acComboBox:
            ct = "Combo box"
        Case acSubform:
            ct = "SubForm"
        Case acObjectFrame:
            ct = "Unbound object frame or chart"
        Case acPageBreak:
            ct = "Page break"
        Case acPage:
            ct = "Page"
        Case acCustomControl:
            ct = "ActiveX (custom) control"
        Case acToggleButton:
            ct = "Toggle Button"
        Case acTabCtl
            ct = "Tab Control"
        Case Else
            ct = ctrl.ControlType
    End Select
    ControlType = ct
End Function

