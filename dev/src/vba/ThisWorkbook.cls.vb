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
    Versionning.ExportVisualBasicCode
    Call Common.updateTemplateList
End Sub


Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    If G_SaveAsOnGoing Then Exit Sub
    Debug.Print "Setting file properties for TEMPLATE"
    With ThisWorkbook
        .BuiltinDocumentProperties("Title") = "Security audit of <hidden> for <secret> by 1mm0rt41PC v" & getInfo("VERSION_DATE")
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
    Call CleanUpInvalidExcelRef
    ' On cleanup la liste des templates
    Range("REPORT_TYPE_LIST").Value2 = ""
End Sub


Private Sub Workbook_Open()
    ' On install les pr�-requis
    If Not Common.isDevMode() Then
        If Not IOFile.isFile(Common.PowerAuditorPath() & "\desktop.ini") Then
            MsgBox "It seems that the powerauditor dependencies are not installed." & vbNewLine & "The installation of dependencies ( git) and their configurations will start now....", vbOKOnly, "PowerAuditor"
            Call VBA.CreateObject("WScript.Shell").Run(Common.PowerAuditorPath() & "\install\", 0, False)
        End If
    End If
        
    Call updateTemplateList
    
    ' On update les repos
    Call IOFile.git("pull", "vulndb")
    Call IOFile.git("pull", "template")
    Versionning.VBAFromCommonSrc
End Sub


Sub ExportExcelToWordTemplate(control As Object)
    If MsgBox("Do you want generate the word template ?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    
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
    Call Word.setCCVal(wDoc, "VERSION_DATE", getInfo("VERSION_DATE"))
    Call Word.setCCVal(wDoc, "BEGIN_DATE", getInfo("BEGIN_DATE"))
    Call Word.setCCVal(wDoc, "END_DATE", getInfo("END_DATE"))
    Call Word.setCCVal(wDoc, "LEVEL", CleaupScoreMesg(getInfo("LEVEL")))
    Call copyExcelColor2Word(wDoc, "LEVEL", Worksheets("PowerAuditor").Range("LEVEL"))
    Call Word.setCCVal(wDoc, "LEVEL_higlight", "{" & getInfo("LEVEL") & "}")
    Dim aText As Variant: aText = Array(getInfo("LEVEL"))
    Call updateColorBoldSynthesis(wDoc, aText, True)

    Call Word.setCCVal(wDoc, "AUTHOR_EMAIL", Common.getFromO365("EmailAddress"))
    Call Word.setCCVal(wDoc, "AUTHOR", Common.getFromO365("FriendlyName"))
    
    Set ws = Worksheets(getInfo("REPORT_TYPE") & "-" & getInfo("LANG"))
    
    G_naturalTableColor1 = ws.Cells(2, 1).DisplayFormat.Interior.color
    G_naturalTableColor2 = ws.Cells(3, 1).DisplayFormat.Interior.color
    
    Call Application.Run("RT_" & getInfo("REPORT_TYPE") & ".init", wDoc, ws)
    While ws.Cells(iRow, 1).Value2 <> ""
        Call Application.Run("RT_" & getInfo("REPORT_TYPE") & ".insertVuln", wDoc, ws, iRow)
        iRow = iRow + 1
    Wend
    nbVuln = iRow - 3
    
    Call Application.Run("RT_" & getInfo("REPORT_TYPE") & ".finish", wDoc, ws, nbVuln)
    wDoc.Fields.Update
    MsgBox "Generation done :-)"
End Sub



Sub GenSynthesis(control As Object)
    If MsgBox("Do you want generate the SYTHESIS ?" & vbNewLine & "This action will >>>REMOVE<<< the current SYTHESIS !!!!!!!!", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    
    Dim ws As Worksheet: Set ws = Worksheets(getInfo("REPORT_TYPE") & "-" & getInfo("LANG"))
    Dim wDoc As Object: Set wDoc = Word.getInstance()
    Call Application.Run("RT_" & getInfo("REPORT_TYPE") & ".GenSynthesis", wDoc, ws)
    MsgBox "Generated"
End Sub


Sub ExportFinalStaticsDocuments(control As Object)
    If MsgBox("Do you want export the template to finals documents ?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    Dim ws As Worksheet: Set ws = Worksheets(getInfo("REPORT_TYPE") & "-" & getInfo("LANG"))
    Dim wDoc As Object: Set wDoc = Word.getInstance()
    Call Application.Run("RT_" & getInfo("REPORT_TYPE") & ".ExportFinalStaticsDocuments", wDoc, ws)
    MsgBox "Generated"
End Sub



Public Sub toProd(control As Object)
    Application.DisplayAlerts = False
    Dim sFilepath As String
        
    ' On upgrade le template vers la bonne destination
    Dim wb_exp As Workbook: Set wb_exp = Workbooks.Add
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.name <> "PowerAuditor" Then
            ThisWorkbook.Sheets(ws.name).Copy after:=wb_exp.Sheets(1)
        End If
    Next ws
    wb_exp.Sheets(1).Delete
    sFilepath = Replace(Common.PowerAuditorPath & "\template\" & Common.getInfo("REPORT_TYPE") & "-" & Common.getInfo("LANG") & ".xlsm", "\\", "\")
    Call wb_exp.SaveAs(sFilepath, FileFormat:=xlOpenXMLWorkbookMacroEnabled)
    wb_exp.Close
    
    ' On export POWERAUDITOR vers le bon dossier de prod
    ThisWorkbook.Save ' On save la version de dev l� o� elle est
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
    Range("REPORT_TYPE_LIST").Value2 = ""
    Range("REPORT_TYPE").Value2 = ""
    
    Call CleanUpInvalidExcelRef
    G_SaveAsOnGoing = True
    Call ThisWorkbook.SaveAs(Replace(Common.PowerAuditorPath() & "\PowerAuditor_v" & Year(Now) & Month(Now) & Day(Now) & ".xlsm", "\\", "\"), FileFormat:=xlOpenXMLWorkbookMacroEnabled)
    G_SaveAsOnGoing = False
    Call ThisWorkbook.Application.Workbooks.Open(sFilepath)
    Application.DisplayAlerts = True
    ThisWorkbook.Close
End Sub


Public Sub WorkInProgress(control As Object)
    MsgBox "TODO !"
End Sub


Public Sub FillExcelWithProof(control As Object)
    If MsgBox("Do you want fill this excel with your proof ?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    Dim ws As Worksheet: Set ws = Worksheets(getInfo("REPORT_TYPE") & "-" & getInfo("LANG"))

    Dim COL_ID As Integer: COL_ID = Common.getColLocation(ws, "id")
    Dim COL_NAME As Integer: COL_NAME = Common.getColLocation(ws, "name")
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
            
            If IOFile.isFile(Common.VulnDBPath(sFile) & "\desc.html") Then
                sPath = Common.VulnDBPath(sFile)
                For i = 0 To UBound(toImportText)
                    ws.Cells(iRow, Common.getColLocation(ws, toImportText(i))).Value2 = Common.trim(IOFile.fileGetContent(sPath & "\" & toImportText(i) & ".html"), Chr(10) & Chr(13))
                Next i
            End If
            iRow = iRow + 1
        End If
        pFile = Dir
    Loop
    
    MsgBox "Import done"
End Sub


Public Sub ExportVulnToGit(control As Object)
    If MsgBox("Do you want to export your vulnerabilities to the GIT ?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    Dim iRow As Integer: iRow = 3
    Dim ws As Worksheet: Set ws = Worksheets(getInfo("REPORT_TYPE") & "-" & getInfo("LANG"))
    Dim wDoc As Object: Set wDoc = Word.getInstance()
    
    Dim COL_ID As Integer: COL_ID = Common.getColLocation(ws, "id")
    Dim COL_NAME As Integer: COL_NAME = Common.getColLocation(ws, "name")
    Dim name As String
    Dim sPath As String
    Dim toExportText As Variant: toExportText = Array("desc", "category", "fixtype", "risk", "fix")
    Dim toExportHTML As Variant: toExportHTML = Array("descDetails", "fixDetails", "fixDetails")
    Dim i As Integer
    Dim windowStyle As Integer: windowStyle = 0 ' Invisible Window
    Dim waitOnReturn As Boolean: waitOnReturn = True
    Dim wsh As Object: Set wsh = VBA.CreateObject("WScript.Shell")
    If Not IOFile.git("pull") Then Exit Sub
    While ws.Cells(iRow, 1).Value2 <> ""
        name = ws.Cells(iRow, COL_NAME).Value2
        If MsgBox("Export >" & name & "< to the GIT ?", vbYesNo + vbQuestion) = vbYes Then
            Debug.Print "Export VULN to GIT: " & name
            sPath = Common.VulnDBPath(name)
            Call IOFile.MyMkDir(sPath)
            For i = 0 To UBound(toExportText)
                Call IOFile.fileSetContent(sPath & "\" & toExportText(i) & ".html", ws.Cells(iRow, Common.getColLocation(ws, toExportText(i))).Value2)
            Next i
            For i = 0 To UBound(toExportHTML)
                ' Enregistre au format HTML avec un dossier s�par�, avec le strict n�c�ssaire (img & css) (wdFormatFilteredHTML=10)
                wDoc.SelectContentControlsByTitle("VLN_" & toExportHTML(i) & "_" & ws.Cells(iRow, COL_ID).Value2)(1).Range.ExportFragment sPath & "\" & toExportHTML(i) & ".html", wdFormatFilteredHTML
            Next i
            If Not IOFile.git("add .") Then Exit Sub
            If Not IOFile.git("commit -am " & Chr(34) & "Update the vulnerability " & Replace(name, Chr(34), "") & Chr(34)) Then Exit Sub
        End If
        iRow = iRow + 1
    Wend
    If Not IOFile.git("push -u origin master") Then Exit Sub
    MsgBox "Export done"
End Sub
