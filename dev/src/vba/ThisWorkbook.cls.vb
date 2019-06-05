VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' PowerAuditor - A simple script to help report writing by https://github.com/1mm0rt41PC
'
' Filename: ThisWorkbook.cls.vb
'
' This program is free software; you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation; either version 2 of the License, or
' (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program; see the file COPYING. If not, write to the
' Free Software Foundation, Inc., 675 Mass Ave, Cambridge, MA 02139, USA.


Public G_naturalTableColor1 As Long
Public G_naturalTableColor2 As Long
Public G_ws As Worksheet
Public G_SaveAsOnGoing As Boolean
Public G_exportToProd As Boolean

Private Sub Workbook_AfterSave(ByVal Success As Boolean)
    If G_SaveAsOnGoing Then Exit Sub
    Versionning.exportVisualBasicCode
    Call Xls.updateTemplateList
End Sub


Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, ByRef Cancel As Boolean)
    If G_SaveAsOnGoing Then Exit Sub
    If SaveAsUI = False And IOFile.getPowerAuditorPath() & "PowerAuditor_last.xlsm" = ThisWorkbook.FullName Then
        Cancel = True
        MsgBox "Your are not allowed to save that file", vbOKOnly + vbSystemModal + vbInformation, "PowerAuditor"
        Exit Sub
    End If
    Debug.Print "Setting file properties for TEMPLATE"
    With ThisWorkbook
        .BuiltinDocumentProperties("Title") = "Security audit of <hidden> for <secret> by 1mm0rt41PC v" & Xls.getVersionDate()
        .BuiltinDocumentProperties("Subject") = .BuiltinDocumentProperties("Title")
        .BuiltinDocumentProperties("Author") = getFromO365("FriendlyName")
        .BuiltinDocumentProperties("Manager") = "1mm0rt41PC"
        .BuiltinDocumentProperties("Company") = "1mm0rt41PC"
        .BuiltinDocumentProperties("Category") = "Audit Documents"
        .BuiltinDocumentProperties("Keywords") = ""
        .BuiltinDocumentProperties("Comments") = .BuiltinDocumentProperties("Title")
    End With
    If getInfo("CLIENT") = "PowerAuditor" Then
        Application.DisplayAlerts = False
        ' Suppression des feuilles
        Dim ws As Worksheet
        For Each ws In ThisWorkbook.Worksheets
            If ws.name <> "PowerAuditor" Then
                ws.Delete
            End If
        Next ws
        Application.DisplayAlerts = True
    End If
    Call Xls.cleanUpInvalidExcelRef
    ' On cleanup la liste des templates
    Call Xls.cleanupTemplateList
End Sub


Private Sub Workbook_Open()
    ' On install les pré-requis
    If Not Common.isDevMode() Then
        If Not IOFile.isFile(IOFile.getPowerAuditorPath() & "\desktop.ini") Then
            MsgBox "It seems that the powerauditor dependencies are not installed." & vbNewLine & "The installation of dependencies ( git) and their configurations will start now....", vbOKOnly + vbSystemModal + vbInformation, "PowerAuditor"
            Call IOFile.runCmd(IOFile.getPowerAuditorPath() & "\install\setup.bat", 1, True)
        End If
    End If
        
    Call Xls.updateTemplateList
    
    ' On update les repos
    Call IOFile.runCmd(IOFile.getPowerAuditorPath() & "\bin\AutoUpdater.exe", 0, False)
    'Versionning.VBAFromCommonSrc
End Sub


Public Sub exportExcelToWordTemplate(control As Object)
    If MsgBox("Do you want generate the word template ?", vbYesNo + vbQuestion + vbSystemModal) = vbNo Then Exit Sub
    If Common.isEmptyString(Common.getInfo("LEVEL")) Then
        Worksheets("PowerAuditor").Activate
        Worksheets("PowerAuditor").Range("LEVEL").Select
        Call MsgBox("Please set the >Global level< !", vbOKOnly + vbInformation + vbSystemModal)
        Exit Sub
    End If
    
    ' On renome le template avec le bon nom
    If Not Common.isDevMode() Then
        renameDocument ThisWorkbook, "xlsm", "TEMPLATE", deleteOld:=True
    End If

    Dim i As Integer
    Dim iRow As Integer: iRow = 3
    Dim ws As Worksheet
    Dim wDoc As Object: Set wDoc = Word.getInstance()
    Dim nbVuln As Integer
    
    ' On defini les constantes
    Call Word.setCCVal(wDoc, "CLIENT", getInfo("CLIENT"))
    Call Word.setCCVal(wDoc, "TARGET", getInfo("TARGET"))
    Call Word.setCCVal(wDoc, "SCOPE", getInfo("SCOPE"))
    Call Word.setCCVal(wDoc, "VERSION_DATE", Xls.getVersionDate())
    Call Word.setCCVal(wDoc, "BEGIN_DATE", getInfo("BEGIN_DATE"))
    Call Word.setCCVal(wDoc, "END_DATE", getInfo("END_DATE"))
    Call Word.setCCVal(wDoc, "LEVEL", cleaupScoreMesg(getInfo("LEVEL")))
    Call copyExcelColor2Word(wDoc, "LEVEL", Worksheets("PowerAuditor").Range("LEVEL"))
    Call Word.setCCVal(wDoc, "LEVEL_higlight", "{" & getInfo("LEVEL") & "}")
    Dim aText As Variant: aText = Array(getInfo("LEVEL"))
    Call updateColorBoldSynthesis(wDoc, aText, True)

    Call Word.setCCVal(wDoc, "AUTHOR_EMAIL", Common.getFromO365("EmailAddress"))
    Call Word.setCCVal(wDoc, "AUTHOR", Common.getFromO365("FriendlyName"))
    
    Set ws = Worksheets(getInfo("REPORT_TYPE"))
    
    G_naturalTableColor1 = ws.Cells(2, 1).DisplayFormat.Interior.color
    G_naturalTableColor2 = ws.Cells(3, 1).DisplayFormat.Interior.color
    
    Call RT.initWordExport(wDoc, ws)
    While ws.Cells(iRow, 1).Value2 <> ""
        Call RT.insertVuln(wDoc, ws, iRow)
        iRow = iRow + 1
    Wend
    nbVuln = iRow - 3
    
    Call RT.finalizeWordExport(wDoc, ws, nbVuln)
    wDoc.Fields.Update
    MsgBox "Generation done :-)", vbSystemModal + vbInformation, "PowerAuditor"
End Sub



Sub genSynthesis(control As Object)
    If MsgBox("Do you want generate the SYTHESIS ?" & vbNewLine & "This action will >>>REMOVE<<< the current SYTHESIS !!!!!!!!", vbYesNo + vbQuestion + vbSystemModal, "PowerAuditor") = vbNo Then Exit Sub
    If Common.isEmptyString(Common.getInfo("LEVEL")) Then
        Worksheets("PowerAuditor").Activate
        Worksheets("PowerAuditor").Range("LEVEL").Select
        Call MsgBox("Please set the >Global level< !", vbOKOnly + vbInformation + vbSystemModal, "PowerAuditor")
        Exit Sub
    End If
    
    Dim ws As Worksheet: Set ws = Worksheets(getInfo("REPORT_TYPE"))
    Dim wDoc As Object: Set wDoc = Word.getInstance()
    Call RT.genSynthesis(wDoc, ws)
    MsgBox "Generated", vbInformation + vbSystemModal, "PowerAuditor"
End Sub


Sub exportFinalStaticsDocuments(control As Object)
    If MsgBox("Do you want export the template to finals documents ?", vbYesNo + vbQuestion + vbSystemModal, "PowerAuditor") = vbNo Then Exit Sub
    If Common.isEmptyString(Common.getInfo("LEVEL")) Then
        Worksheets("PowerAuditor").Activate
        Worksheets("PowerAuditor").Range("LEVEL").Select
        Call MsgBox("Please set the >Global level< !", vbOKOnly + vbInformation + vbSystemModal, "PowerAuditor")
        Exit Sub
    End If
    
    Dim ws As Worksheet: Set ws = Worksheets(getInfo("REPORT_TYPE"))
    Dim wDoc As Object: Set wDoc = Word.getInstance()
    Call Word.setCCVal(wDoc, "VERSION_DATE", Xls.getVersionDate())
    Call RT.exportFinalStaticsDocuments(wDoc, ws)
    Call IOFile.runCmd("explorer.exe " & ThisWorkbook.Path & "\output", 1, False)
    MsgBox "Generated", vbInformation + vbSystemModal, "PowerAuditor"
End Sub


Public Sub ToProd(control As Object)
    If Not Common.isDevMode() Then
        MsgBox "You do not use the xlsm development file", vbOKOnly + vbSystemModal + vbInformation, "PowerAuditor"
        Exit Sub
    End If
    Application.DisplayAlerts = False
    Dim sFilepath As String
    
    If Month(Now) < 10 Then
        Range("PowerAuditorVersion").Value2 = Year(Now) & "-0" & Month(Now) & "-" & Day(Now)
    Else
        Range("PowerAuditorVersion").Value2 = Year(Now) & "-" & Month(Now) & "-" & Day(Now)
    End If
    Range("TemplateVersion").Value2 = Range("PowerAuditorVersion").Value2
        
    ' On upgrade le template vers la bonne destination
    Dim wb_exp As Workbook: Set wb_exp = Workbooks.Add
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.name <> "PowerAuditor" Then
            ThisWorkbook.Sheets(ws.name).Copy After:=wb_exp.Sheets(1)
        End If
    Next ws
    wb_exp.Sheets(1).Delete
    sFilepath = Replace(IOFile.getPowerAuditorPath & "\template\" & Common.getInfo("REPORT_TYPE") & ".xlsm", "\\", "\")
    Call wb_exp.SaveAs(sFilepath, FileFormat:=xlOpenXMLWorkbookMacroEnabled)
    wb_exp.Close
    
    ' On export POWERAUDITOR vers le bon dossier de prod
    ThisWorkbook.Save ' On save la version de dev là où elle est
    sFilepath = ThisWorkbook.FullName ' On grade le path actuel pour y revenir
    For Each ws In ThisWorkbook.Worksheets
        If ws.name <> "PowerAuditor" Then
            ws.Delete
        End If
    Next ws
    With ThisWorkbook
        .BuiltinDocumentProperties("Title") = "Security audit of <hidden> for <secret> by 1mm0rt41PC v" & Year(Now) & Month(Now) & Day(Now)
        .BuiltinDocumentProperties("Subject") = .BuiltinDocumentProperties("Title")
        .BuiltinDocumentProperties("Author") = "1mm0rt41PC"
        .BuiltinDocumentProperties("Manager") = "1mm0rt41PC"
        .BuiltinDocumentProperties("Company") = "1mm0rt41PC"
        .BuiltinDocumentProperties("Category") = "Audit Documents"
        .BuiltinDocumentProperties("Keywords") = ""
        .BuiltinDocumentProperties("Comments") = .BuiltinDocumentProperties("Title")
    End With
    ' On cleanup la liste des templates
    G_exportToProd = True
    Xls.cleanupTemplateList
    Range("REPORT_TYPE").Value2 = ""
    
    Call Xls.cleanUpInvalidExcelRef
    G_SaveAsOnGoing = True
    Dim sNewPath As String: sNewPath = Replace(IOFile.getPowerAuditorPath() & "\PowerAuditor_", "\\", "\")
    Call ThisWorkbook.SaveAs(sNewPath & "v" & Year(Now) & Month(Now) & Day(Now) & ".xlsm", FileFormat:=xlOpenXMLWorkbookMacroEnabled)
    Call ThisWorkbook.SaveAs(sNewPath & "last.xlsm", FileFormat:=xlOpenXMLWorkbookMacroEnabled)
    G_SaveAsOnGoing = False
    Call ThisWorkbook.Application.Workbooks.Open(sFilepath)
    Application.DisplayAlerts = True
    ThisWorkbook.Close
End Sub


Public Sub fillExcelWithProof(control As Object)
    If MsgBox("Do you want fill this excel with your proof ?", vbYesNo + vbQuestion + vbSystemModal, "PowerAuditor") = vbNo Then Exit Sub
    Dim ws As Worksheet: Set ws = Worksheets(getInfo("REPORT_TYPE"))

    Dim COL_ID As Integer: COL_ID = Xls.getColLocation(ws, "id")
    Dim COL_NAME As Integer: COL_NAME = RT.getExportField_KeyColumn(ws)
    Dim toImportText As Variant: toImportText = Array("desc", "category", "fixtype", "risk", "fix")
    Dim i As Integer
    Dim vlnDir As String: vlnDir = ActiveWorkbook.Path & "\vuln\"
    Dim iRow As Integer: iRow = 3
    Dim pFile: pFile = Dir(vlnDir & "*", vbDirectory)
    Dim sPath As String
    Dim sFile As String
    Do While pFile <> ""
        sFile = pFile
        If Left(sFile, 1) <> "." Then
            ws.Cells(iRow, 1).EntireRow.Insert
            ws.Cells(iRow + 1, 1).EntireRow.Copy ws.Cells(iRow, 1)
            ws.Cells(iRow, COL_ID).Value2 = iRow - 2
            ws.Cells(iRow, COL_NAME).Value2 = sFile
            
            If IOFile.isFile(IOFile.getVulnDBPath(sFile) & "\desc.html") Then
                sPath = IOFile.getVulnDBPath(sFile)
                For i = 0 To UBound(toImportText)
                    ws.Cells(iRow, Xls.getColLocation(ws, toImportText(i))).Value2 = Common.trim(IOFile.fileGetContent(sPath & "\" & toImportText(i) & ".html"), Chr(10) & Chr(13))
                Next i
            End If
            iRow = iRow + 1
        End If
        pFile = Dir
    Loop
    
    MsgBox "Import done", vbSystemModal + vbInformation, "PowerAuditor"
End Sub


Public Sub exportVulnToGit(control As Object)
    If MsgBox("Do you want to export your vulnerabilities to the GIT ?", vbYesNo + vbQuestion + vbSystemModal, "PowerAuditor") = vbNo Then Exit Sub
    Dim iRow As Integer: iRow = 3
    Dim ws As Worksheet: Set ws = Worksheets(getInfo("REPORT_TYPE"))
    Dim wDoc As Object: Set wDoc = Word.getInstance()
    
    Dim COL_ID As Integer: COL_ID = Xls.getColLocation(ws, "id")
    Dim COL_NAME As Integer: COL_NAME = Xls.getColLocation(ws, "name")
    Dim name As String
    Dim sPath As String
    Dim toExportText As Variant: toExportText = RT.getExportFields_TXT
    Dim toExportHTML As Variant: toExportHTML = RT.getExportFields_HTML
    Dim toExportKeyCol As Integer: toExportKeyCol = RT.getExportField_KeyColumn(ws)
    Dim i As Integer
    If Not IOFile.git("pull") Then Exit Sub
    While ws.Cells(iRow, 1).Value2 <> ""
        name = ws.Cells(iRow, toExportKeyCol).Value2
        If MsgBox("Export >" & name & "< to the GIT ?", vbYesNo + vbQuestion + vbSystemModal, "PowerAuditor") = vbYes Then
            Debug.Print "Export VULN to GIT: " & name
            sPath = IOFile.getVulnDBPath(name, True)
            For i = 0 To UBound(toExportText)
                Call IOFile.fileSetContent(sPath & "\" & toExportText(i) & ".html", ws.Cells(iRow, Xls.getColLocation(ws, toExportText(i))).Value2)
            Next i
            For i = 0 To UBound(toExportHTML)
                ' Enregistre au format HTML avec un dossier séparé, avec le strict nécéssaire (img & css) (wdFormatFilteredHTML=10)
                wDoc.SelectContentControlsByTitle("VLN_" & toExportHTML(i) & "_" & ws.Cells(iRow, COL_ID).Value2)(1).Range.ExportFragment sPath & "\" & toExportHTML(i) & ".html", wdFormatHTML
            Next i
            sPath = IOFile.getNotableFile(name)
            If Not IOFile.isFile(sPath) Then
                Dim mo: mo = Month(Now())
                If mo < 10 Then mo = "0" & mo
                Call IOFile.fileSetContent(sPath, "---" & vbLf & _
                "title: " & name & vbLf & _
                "created: '" & Year(Now()) & "-" & mo & "-" & Day(Now()) & "T" & Split(Now(), " ")(1) & "Z'" & vbLf & _
                "modified: '" & Year(Now()) & "-" & mo & "-" & Day(Now()) & "T" & Split(Now(), " ")(1) & "Z'" & vbLf & _
                "tags: [Pentest/Fiche de vuln/A trier/]" & vbLf & _
                "---" & vbLf & _
                "" & vbLf & _
                "# 1. " & name & vbLf & _
                "" & vbLf & _
                "" & vbLf)
            End If
            If Not IOFile.git("add .") Then Exit Sub
            If Not IOFile.git("commit -am " & Chr(34) & "Update the vulnerability " & Replace(name, Chr(34), "") & Chr(34)) Then Exit Sub
        End If
        iRow = iRow + 1
    Wend
    If Not IOFile.git("push -u origin master") Then Exit Sub
    MsgBox "Export done", vbSystemModal + vbInformation, "PowerAuditor"
End Sub
