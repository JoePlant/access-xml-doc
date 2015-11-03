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

