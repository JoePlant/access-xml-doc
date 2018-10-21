VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsXmlWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private xml As String
Private level As Integer
Private bWithAttribute As Boolean
Private line As String

Private Function pad(l As Integer) As String
    Select Case l
        Case 0:
            pad = ""
        Case 1:
            pad = "  "
        Case 2:
            pad = "    "
        Case 3:
            pad = "      "
        Case 4:
            pad = "        "
        Case 5:
            pad = "          "
        Case 6:
            pad = "            "
        Case 7:
            pad = "              "
        Case 8:
            pad = "                "
        Case 9:
            pad = "                  "
        Case 10:
            pad = "                    "
        Case Else
            Dim p As String
            Dim i As Integer
            For i = 1 To l
                p = p & "  "
            Next
            pad = p
    End Select
End Function

Property Get XmlOutput() As String
    XmlOutput = xml
End Property

Public Sub WriteElementStart(Name As String)
    line = pad(level) & "<" & Name & " "
    'xml = xml & pad(level) & "<" & name
    level = level + 1
End Sub

Public Sub WriteElementEnd(Name As String, Optional bCrLf As Boolean = True)
    'xml = xml & ">" & vbCrLf
    line = line & ">"
    xml = xml & line
    If (bCrLf) Then
        xml = xml & vbCrLf
    End If
    'Debug.Print line
    line = ""
    bWithAttribute = False
End Sub

Public Sub WriteElementClose(Name As String)
    'xml = xml & ">" & vbCrLf
    line = line & "/>"
    xml = xml & line & vbCrLf
    'Debug.Print line
    level = level - 1
    line = ""
    bWithAttribute = False
End Sub

Public Sub WriteAttribute(Name As String, Value As String)
    line = line & "" & Name & "='" & Value & "' "
    'xml = xml & " " & name & "='" & value & "' "
    bWithAttribute = True
End Sub

Public Sub WriteText(Value As String, Optional bCrLf As Boolean = True)
    line = line & Escape(Value)
    If (bCrLf) Then
        line = line & vbCrLf
    End If
End Sub

Public Sub WriteEndElement(Name As String, Optional bPad As Boolean = True)
    level = level - 1
    If (bPad) Then
        line = line & pad(level) & "</" & Name & ">" & vbCrLf
    Else
        line = line & "</" & Name & ">" & vbCrLf
    End If
    xml = xml & line
    line = ""
End Sub

Private Function Escape(Value As String) As String
    Dim escaped As String
    escaped = Replace(Value, "&", "&amp;")
    escaped = Replace(escaped, "<", "&lt;")
    escaped = Replace(escaped, ">", "&gt;")
    escaped = Replace(escaped, "'", "&apos;")
    escaped = Replace(escaped, """", "&quot;")
	
	' Values are sometimes null terminated 
	If (Len(escaped) > 0) Then
		If Mid$(escaped, Len(escaped), 1) = Chr$(0) Then 
			escaped = Left$(escaped, Len(escaped) - 1)
		End If
	End If
    
    Escape = escaped
End Function

Public Sub WritePropertiesAsXml(props As Properties)

Dim prop As Property

    Call WriteElementStart("Properties")
    Call WriteAttribute("count", props.Count)
    Call WriteElementEnd("Properties")
    
    For Each prop In props
        Call WriteProperty(prop)
    
'        Select Case prop.Type
'            Case 8209
'                Select Case prop.Name
'                    Case ""
'                    Case Else
'                        Call WriteProperty(prop)
'                End Select
                'ignore pictures
'            Case Else
'        End Select
    Next
    
    Call WriteEndElement("Properties")
End Sub

Private Sub WriteProperty(prop As Property)
On Error Resume Next

Dim v As String
Dim n As String
Dim t As String
Dim pt As String

    v = prop.Value

'    If (v = "") Then
    
'    Else
    
        n = prop.Name
        t = prop.Type
        If (t <> 8209) Then   ' vbByte + vbArray
            pt = LookupPropertyType(prop.Type)
                
            Call WriteElementStart(Name:="Property")
            Call WriteAttribute(Name:="name", Value:=n)
            Call WriteAttribute(Name:="type", Value:=pt)
            Call WriteElementEnd(Name:="Property", bCrLf:=False)
            Call WriteText(Value:=v, bCrLf:=False)
            Call WriteEndElement(Name:="Property", bPad:=False)
        End If
'    End If
End Sub

Private Function LookupPropertyType(iType As Integer) As String
Dim sType As String
    Select Case iType
        Case vbEmpty
            sType = "Empty"
        Case vbNull
            sType = "Null"
        Case vbInteger
            sType = "Integer"
        Case vbLong
            sType = "Long"
        Case vbSingle
            sType = "Single"
        Case vbDouble
            sType = "Double"
        Case vbCurrency
            sType = "Currency"
        Case vbDate
            sType = "Date"
        Case vbString
            sType = "String"
        Case vbObject
            sType = "Object"
        Case vbError
            sType = "Error"
        Case vbBoolean
            sType = "Boolean"
        Case vbVariant
            sType = "Variant"
        Case vbDataObject
            sType = "DataObject"
        Case vbDecimal
            sType = "Decimal"
        Case vbByte
            sType = "Byte"
        Case vbUserDefinedType
            sType = "UserDefinedType"
        Case Is > vbArray
            sType = LookupPropertyType(iType - vbArray) & " Array"
        Case Else
            sType = iType
    End Select
    LookupPropertyType = sType
End Function

Private Sub Class_Initialize()
 xml = ""
 level = 0
 line = ""
End Sub
