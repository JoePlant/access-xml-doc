Attribute VB_Name = "modExport"
Option Compare Database
Option Explicit

Sub ExportAll()

    Call ExportDatabaseAsXml
    Call ExportFormsAsXml
    Call ExportReportsAsXml
    Call ExportModulesAsXml
    
End Sub

Sub ExportModulesAsXml()
    Dim accObj As AccessObject
    Dim bWasOpen As Boolean
    Dim strDoc As String
    Dim m As Module
    
    For Each accObj In CurrentProject.AllModules
        strDoc = accObj.Name
        bWasOpen = accObj.IsLoaded
        If Not bWasOpen Then
            DoCmd.OpenModule strDoc
        End If
        
        DoEvents
        
        Debug.Print "Opening " & strDoc
        
        Set m = Modules(strDoc)
        
        Dim mx As clsModuleExport
        Set mx = New clsModuleExport
        Set mx.Module = m
        mx.Directory = "d:\xml"
        
        Dim mxml As String
        mxml = mx.ExportAsXml
        
        If Not bWasOpen Then
            DoCmd.Close acModule, strDoc
        End If
        
        DoEvents
        
    Next
    
End Sub

Sub ExportFormsAsXml()
    Dim accObj As AccessObject
    Dim bWasOpen As Boolean
    Dim strDoc As String
    Dim f As Form
    
    For Each accObj In CurrentProject.AllForms
        strDoc = accObj.Name
        bWasOpen = accObj.IsLoaded
        If Not bWasOpen Then
            DoCmd.OpenForm strDoc, View:=acDesign, WindowMode:=acWindowNormal
        End If
        
        DoEvents
        
        Debug.Print "Opening " & strDoc
        
        Set f = Forms(strDoc)
        
        Dim fx As clsFormExport
        Set fx = New clsFormExport
        Set fx.Form = f
        fx.Directory = "d:\xml"
        
        Dim fxml As String
        fxml = fx.ExportAsXml
        
        If Not bWasOpen Then
            DoCmd.Close acForm, strDoc, acSaveNo
        End If
        
        DoEvents
        
    Next
    
End Sub

Sub ExportMacrosAsXml()
    Dim accObj As AccessObject
    Dim bWasOpen As Boolean
    Dim strDoc As String
    Dim f As Form
    
    For Each accObj In CurrentProject.AllMacros
        strDoc = accObj.Name
        bWasOpen = accObj.IsLoaded
        If Not bWasOpen Then
            ''DoCmd.OP strDoc, View:=acDesign, WindowMode:=acWindowNormal
        End If
        
        DoEvents
        
        Debug.Print "Opening " & strDoc
        
        
        'Set f = Forms(strDoc)
        
        'Dim fx As clsFormExport
        'Set fx = New clsFormExport
        'Set fx.Form = f
        'fx.Directory = "d:\xml"
        
        'Dim fxml As String
        'fxml = fx.ExportAsXml
        
        'If Not bWasOpen Then
        '    DoCmd.Close acForm, strDoc, acSaveNo
        'End If
        
        DoEvents
        
    Next
    
End Sub


Sub ExportReportsAsXml()
    Dim accObj As AccessObject
    Dim bWasOpen As Boolean
    Dim strDoc As String
    Dim r As Report
    
    Dim iTotal As Integer
    Dim iCount As Integer
    
    iTotal = CurrentProject.AllReports.Count
    
    For Each accObj In CurrentProject.AllReports
        iCount = iCount + 1
        strDoc = accObj.Name
        bWasOpen = accObj.IsLoaded
        If Not bWasOpen Then
            DoCmd.OpenReport strDoc, View:=acDesign, WindowMode:=acWindowNormal
        End If
        
        DoEvents
        
        Debug.Print "Opening " & iCount & " of " & iTotal & " - " & strDoc
        
        Set r = Reports(strDoc)
        
        Dim rx As clsReportExport
        Set rx = New clsReportExport
        Set rx.Report = r
        rx.Directory = "d:\xml"
        
        Dim rxml As String
        rxml = rx.ExportAsXml
        
        If Not bWasOpen Then
            DoCmd.Close acReport, strDoc, acSaveNo
        End If
        
        DoEvents
        
    Next
    
End Sub

Sub ExportDatabaseAsXml()

    Dim dx As clsDatabaseExport
    Set dx = New clsDatabaseExport
    Set dx.Database = CurrentProject
    dx.Directory = "d:\xml"

    Dim dxml As String

    dx.ExportAsXml
    
    DoEvents
    
End Sub


Public Sub DocDatabase()
 '====================================================================
 ' Name:    DocDatabase
 ' Purpose: Documents the database to a series of text files
 '
 ' Author:  Arvin Meyer
 ' Date:    June 02, 1999
 ' Comment: Uses the undocumented [Application.SaveAsText] syntax
 '          To reload use the syntax [Application.LoadFromText]
 '      Modified to set a reference to DAO 8/22/2005
 '====================================================================
On Error GoTo Err_DocDatabase
Dim dbs As DAO.Database
Dim cnt As DAO.Container
Dim doc As DAO.Document
Dim i As Integer

Set dbs = CurrentDb() ' use CurrentDb() to refresh Collections

Set cnt = dbs.Containers("Forms")
For Each doc In cnt.Documents
    Application.SaveAsText acForm, doc.Name, "D:\xml\Form_" & doc.Name & ".txt"
Next doc

Set cnt = dbs.Containers("Reports")
For Each doc In cnt.Documents
    Application.SaveAsText acReport, doc.Name, "D:\xml\Report_" & doc.Name & ".txt"
Next doc

Set cnt = dbs.Containers("Scripts")
For Each doc In cnt.Documents
    Application.SaveAsText acMacro, doc.Name, "D:\xml\Scripts_" & doc.Name & ".txt"
Next doc

Set cnt = dbs.Containers("Modules")
For Each doc In cnt.Documents
    Application.SaveAsText acModule, doc.Name, "D:\xml\Module_" & doc.Name & ".txt"
Next doc

For i = 0 To dbs.QueryDefs.Count - 1
    Application.SaveAsText acQuery, dbs.QueryDefs(i).Name, "D:\xml\Query_" & dbs.QueryDefs(i).Name & ".txt"
Next i

Set doc = Nothing
Set cnt = Nothing
Set dbs = Nothing

Exit_DocDatabase:
    Exit Sub


Err_DocDatabase:
    Select Case Err

    Case Else
        MsgBox Err.Description
        Resume Exit_DocDatabase
    End Select

End Sub

