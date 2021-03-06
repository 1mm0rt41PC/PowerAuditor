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


'===============================================================================
' @brief Cleanup the document and export VBA code when saving
' @param...
' @return {NONE}
'
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
    Versionning.exportVisualBasicCode
End Sub


'===============================================================================
' @brief Init environnement at boot:
'        - install all packages
'        - avoid main template to be altered
'        - list all available plugins
'        - run auto-updater
' @return {NONE}
' @warning The system doesn't load anymore all vba code from repo. You need to load all source code manually.
'
Private Sub Workbook_Open()
    ' On install les pr�-requis
    If Not Common.isDevMode() Then
        If Not IOFile.isFile(IOFile.getPowerAuditorPath() & "\desktop.ini") Then
            MsgBox "It seems that the powerauditor dependencies are not installed." & vbNewLine & "The installation of dependencies ( git) and their configurations will start now....", vbOKOnly + vbSystemModal + vbInformation, "PowerAuditor"
            Call IOFile.runCmd(IOFile.getPowerAuditorPath() & "\install\setup.bat", 1, True)
        End If
    End If
    
    If IOFile.getPowerAuditorPath() & "PowerAuditor_last.xlsm" = ThisWorkbook.FullName Then
        MsgBox "Do not open this file directly !" & vbNewLine & "Please copy this file into a new folder !", vbOKOnly + vbSystemModal + vbInformation, "PowerAuditor"
        Application.Quit
        Exit Sub
    End If
        
    Call Xls.updateTemplateList
    
    ' On update les repos
    Call IOFile.runCmd(IOFile.getPowerAuditorPath() & "\bin\AutoUpdater.exe", 0, False)
    'Versionning.VBAFromCommonSrc
End Sub


'===============================================================================
' @brief This sub allows to fill a Word document with the Excel datas
'        This function will:
'        - set all fields, CLIENT, TARGET, SCOPE, ...
'        - call RT.exportExcel2Word_before, RT.exportExcel2Word_insertVuln, RT.exportExcel2Word_after.
' @return {NONE}
'
Public Sub exportExcelToWordTemplate(control As Object)
    Dim ws As Worksheet: Set ws = Xls.getWorksheetRT()
    If ws Is Nothing Then Exit Sub
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
    
    Call RT.exportExcel2Word_before(wDoc, ws)
    While ws.Cells(iRow, 1).Value2 <> ""
        Call RT.exportExcel2Word_insertVuln(wDoc, ws, iRow)
        iRow = iRow + 1
    Wend
    nbVuln = iRow - 3
    
    Call RT.exportExcel2Word_after(wDoc, ws, nbVuln)
    wDoc.Fields.Update
    MsgBox "Generation done :-)", vbSystemModal + vbInformation, "PowerAuditor"
End Sub


'===============================================================================
' @brief This sub allows to generate the synthesis of the word document with Excel datas.
'        This function is a basic wrapper of RT.genSynthesis
' @return {NONE}
'
Sub genSynthesis(control As Object)
    Dim ws As Worksheet: Set ws = Xls.getWorksheetRT()
    If ws Is Nothing Then Exit Sub
    If MsgBox("Do you want generate the SYTHESIS ?" & vbNewLine & "This action will >>>REMOVE<<< the current SYTHESIS !!!!!!!!", vbYesNo + vbQuestion + vbSystemModal, "PowerAuditor") = vbNo Then Exit Sub
    If Common.isEmptyString(Common.getInfo("LEVEL")) Then
        Worksheets("PowerAuditor").Activate
        Worksheets("PowerAuditor").Range("LEVEL").Select
        Call MsgBox("Please set the >Global level< !", vbOKOnly + vbInformation + vbSystemModal, "PowerAuditor")
        Exit Sub
    End If
    
    Dim wDoc As Object: Set wDoc = Word.getInstance()
    Call RT.genSynthesis(wDoc, ws)
    MsgBox "Generated", vbInformation + vbSystemModal, "PowerAuditor"
End Sub


'===============================================================================
' @brief This sub allows to generate finals documents.
'        This function is a basic wrapper of RT.exportFinalStaticsDocuments
' @return {NONE}
'
Sub exportFinalStaticsDocuments(control As Object)
    Dim ws As Worksheet: Set ws = Xls.getWorksheetRT()
    If ws Is Nothing Then Exit Sub
    If MsgBox("Do you want export the template to finals documents ?", vbYesNo + vbQuestion + vbSystemModal, "PowerAuditor") = vbNo Then Exit Sub
    If Common.isEmptyString(Common.getInfo("LEVEL")) Then
        Worksheets("PowerAuditor").Activate
        Worksheets("PowerAuditor").Range("LEVEL").Select
        Call MsgBox("Please set the >Global level< !", vbOKOnly + vbInformation + vbSystemModal, "PowerAuditor")
        Exit Sub
    End If
    
    Dim wDoc As Object: Set wDoc = Word.getInstance()
    Call Word.setCCVal(wDoc, "VERSION_DATE", Xls.getVersionDate())
    Call RT.exportFinalStaticsDocuments(wDoc, ws)
    Call IOFile.runCmd("explorer.exe " & ThisWorkbook.Path & "\output", 1, False)
    Call MsgBox("Generated", vbInformation + vbSystemModal, "PowerAuditor")
End Sub


'===============================================================================
' @brief This sub allows to send PowerAuditor in production to all users
' @return {NONE}
'
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
    
    If ThisWorkbook.Worksheets.Count > 1 Then
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
    End If
    
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
    Xls.cleanupTemplateList
    ThisWorkbook.Worksheets("PowerAuditor").Range("REPORT_TYPE").Value2 = ""
    ThisWorkbook.Worksheets("PowerAuditor").Range("CLIENT").Value2 = "My Client Name"
    ThisWorkbook.Worksheets("PowerAuditor").Range("TARGET").Value2 = "My App Name"
    ThisWorkbook.Worksheets("PowerAuditor").Range("SCOPE").Value2 = "http://target/ (127.0.0.1)"
    
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


'===============================================================================
' @brief This sub allows to fill the Excel with vulnerabilities from the VULN folder.
'        In addition, this sub uses datas from the shared database.
' @return {NONE}
'
Public Sub fillExcelWithProof(control As Object)
    Dim ws As Worksheet: Set ws = Xls.getWorksheetRT()
    If ws Is Nothing Then Exit Sub
    If MsgBox("Do you want fill this excel with your proof ?", vbYesNo + vbQuestion + vbSystemModal, "PowerAuditor") = vbNo Then Exit Sub

    Dim COL_ID As Integer: COL_ID = Xls.getColLocation(ws, "id")
    Dim COL_NAME As Integer: COL_NAME = RT.getExportField_KeyColumn(ws)
    Dim toImportText As Variant: toImportText = Array("desc", "category", "fixtype", "risk", "fix")
    Dim i As Integer
    Dim vlnDir As String: vlnDir = ActiveWorkbook.Path & "\vuln\"
    Dim iRow As Integer: iRow = nbNotEmptyRows(ws, COL_ID) + 3
    Dim pFile: pFile = Dir(vlnDir & "*", vbDirectory)
    Dim sPath As String
    Dim sFile As String
    Do While pFile <> ""
        sFile = pFile
        If Left(sFile, 1) <> "." Then
            If Xls.countOccurenceInCol(ws, COL_NAME, sFile) = 0 Then
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
        End If
        pFile = Dir
    Loop
    While ws.Cells(iRow, COL_ID).Value2 <> ""
        ws.Cells(iRow, COL_ID).Value2 = iRow - 2
        iRow = iRow + 1
    Wend
    
    MsgBox "Import done", vbSystemModal + vbInformation, "PowerAuditor"
End Sub


'===============================================================================
' @brief This sub allows to fill the Excel with vulnerabilities from the shared database.
' @return {NONE}
'
Public Sub importVulnFromDatabase(control As Object)
    Dim ws As Worksheet: Set ws = Xls.getWorksheetRT()
    If ws Is Nothing Then Exit Sub

    Dim COL_ID As Integer: COL_ID = Xls.getColLocation(ws, "id")
    Dim COL_NAME As Integer: COL_NAME = RT.getExportField_KeyColumn(ws)
    Dim toImportText As Variant: toImportText = Array("desc", "category", "fixtype", "risk", "fix")
    Dim i As Integer
    Dim j As Integer
    Dim sVuln As String: sVuln = Common.PowerImporter()
    If sVuln = "" Then Exit Sub
    Dim aVuln() As String: aVuln = Split(sVuln, vbCrLf)
    
    Dim iRow As Integer: iRow = nbNotEmptyRows(ws, COL_ID) + 3

    Dim sPath As String
    For j = 0 To UBound(aVuln)
        Call IOFile.myMkDir(ThisWorkbook.Path & "\vuln\" & aVuln(j))
        If Xls.countOccurenceInCol(ws, COL_NAME, aVuln(j)) = 0 Then
            ws.Cells(iRow, 1).EntireRow.Insert
            ws.Cells(iRow + 1, 1).EntireRow.Copy ws.Cells(iRow, 1)
            ws.Cells(iRow, COL_ID).Value2 = iRow - 2
            ws.Cells(iRow, COL_NAME).Value2 = aVuln(j)
            If IOFile.isFile(IOFile.getVulnDBPath(aVuln(j)) & "\desc.html") Then
                sPath = IOFile.getVulnDBPath(aVuln(j))
                For i = 0 To UBound(toImportText)
                    ws.Cells(iRow, Xls.getColLocation(ws, toImportText(i))).Value2 = Common.trim(IOFile.fileGetContent(sPath & "\" & toImportText(i) & ".html"), Chr(10) & Chr(13))
                Next i
            End If
            iRow = iRow + 1
        End If
    Next j
    While ws.Cells(iRow, COL_ID).Value2 <> ""
        ws.Cells(iRow, COL_ID).Value2 = iRow - 2
        iRow = iRow + 1
    Wend
    
    MsgBox "Import done", vbSystemModal + vbInformation, "PowerAuditor"
End Sub


'===============================================================================
' @brief This sub allows to export vulnerabilities to the shared database.
' @return {NONE}
'
Public Sub exportVulnToGit(control As Object)
    Dim ws As Worksheet: Set ws = Xls.getWorksheetRT()
    If ws Is Nothing Then Exit Sub
    Dim iRow As Integer: iRow = 3
    Dim COL_ID As Integer: COL_ID = Xls.getColLocation(ws, "id")
    Dim COL_NAME As Integer: COL_NAME = Xls.getColLocation(ws, "name")
    Dim name As String
    Dim sPath As String
    Dim toExportText As Variant: toExportText = RT.getExportFields_TXT
    Dim toExportHTML As Variant: toExportHTML = RT.getExportFields_HTML
    Dim toExportKeyCol As Integer: toExportKeyCol = RT.getExportField_KeyColumn(ws)
    Dim i As Integer
    Dim j As Integer

    Dim sData As String
    While ws.Cells(iRow, 1).Value2 <> ""
        name = ws.Cells(iRow, toExportKeyCol).Value2
        sData = sData & iRow & Chr(9) + name & vbCrLf
        iRow = iRow + 1
    Wend
    sData = Common.PowerExporter(sData)
    If sData = "" Then Exit Sub
    If MsgBox("Do you want EXPORT this excel to the shared database ?", vbYesNo + vbQuestion + vbSystemModal, "PowerAuditor") = vbNo Then Exit Sub
    
    Dim sRows() As String
    sRows = Split(sData, vbCrLf)
    
    Dim wDoc As Object: Set wDoc = Word.getInstance()
    If Not IOFile.git("pull") Then Exit Sub
    Dim aRow() As String
    Dim isDraft As Boolean
    Dim isExploit As Boolean
    For j = 0 To UBound(sRows)
        aRow = Split(sRows(j), "/")
        iRow = CInt(aRow(0))
        isDraft = (aRow(1) = "d")
        isExploit = (aRow(2) = "e")
        name = ws.Cells(iRow, toExportKeyCol).Value2
        
        Debug.Print "Export VULN to GIT: " & name
        sPath = IOFile.getVulnDBPath(name, True)
        For i = 0 To UBound(toExportText)
            Call IOFile.fileSetContent(sPath & "\" & toExportText(i) & ".html", ws.Cells(iRow, Xls.getColLocation(ws, toExportText(i))).Value2)
        Next i
        For i = 0 To UBound(toExportHTML)
            Dim eHtml As String: eHtml = toExportHTML(i)
            If isExploit Then
                eHtml = Replace(eHtml, "*", "")
            End If
            With wDoc.SelectContentControlsByTitle("VLN_" & eHtml & "_" & ws.Cells(iRow, COL_ID).Value2)
                If .Count = 1 Then
                    ' Cr�ation d'une ligne vide afin de permettre l'export HTML
                    .Item(1).Range.InsertAfter Chr(13)
                    ' Enregistre au format HTML avec un dossier s�par�, avec le strict n�c�ssaire (img & css) (wdFormatFilteredHTML=10)
                    .Item(1).Range.ExportFragment sPath & "\" & eHtml & ".html", wdFormatHTML
                    ' Suppression de la ligne vide
                    wDoc.Range(.Item(1).Range.End - 1, .Item(1).Range.End).text = ""
                End If
            End With
        Next i
        
        If isDraft Then
            Call IOFile.removeFile(sPath & "\.validated")
        Else
            Call IOFile.fileSetContent(sPath & "\.validated", Common.getFromO365("EmailAddress"))
        End If

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
    Next j
    If Not IOFile.git("push -u origin master") Then Exit Sub
    MsgBox "Export done", vbSystemModal + vbInformation, "PowerAuditor"
End Sub


'===============================================================================
' @brief This sub allows to highlight code and insert it into the word document.
' @return {NONE}
'
Public Sub syntaxHighlighter(control As Object)
    Dim ws As Worksheet: Set ws = Xls.getWorksheetRT()
    If ws Is Nothing Then Exit Sub
    Dim wDoc As Object: Set wDoc = Word.getInstance()
    Dim tmpFile As String: tmpFile = Environ("temp") & "\" & randomString(7) & ".html"
    Call IOFile.runCmd(IOFile.getPowerAuditorPath() & "\bin\SyntaxHighlighter-Helper.exe " & Chr(34) & tmpFile & Chr(34))
    Dim sLang As String: sLang = IOFile.fileGetContent(tmpFile & ".lang")
    Dim cc As Object
    With wDoc.ActiveWindow.Selection.Range
        .InsertParagraphAfter
        .InsertParagraphAfter
        Set cc = wDoc.Range(.End - 1, .End - 1).ContentControls.Add(wdContentControlRichText)
        cc.title = tmpFile
        cc.Appearance = wdContentControlHidden
    End With
    
    Call Word.pygmentizeMe(wDoc, cc, tmpFile, sLang, "xxx")
    Call cc.Delete(False)
End Sub
