Attribute VB_Name = "Xls"
Option Explicit
' PowerAuditor - A simple script to help report writing by https://github.com/1mm0rt41PC
'
' Filename: Xls.bas.vb
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


Public Function getColLocation(ws As Worksheet, ByVal sCol As String) As Integer
    Dim i As Integer: i = 1
    While ws.Cells(2, i).Value2 <> ""
        If ws.Cells(2, i).Value2 = sCol Then
            getColLocation = i
            Exit Function
        End If
        i = i + 1
    Wend
    Debug.Print "[!] Unable to find column " & sCol
    Debug.Print 0 / 0 ' Raise fatal error
End Function


Public Sub cleanUpInvalidExcelRef()
    ' N�toyage des noms de range
    Dim rangeName As name
    For Each rangeName In ThisWorkbook.Names
        If InStr(1, rangeName.RefersTo, "=[") > 0 Then
            rangeName.RefersTo = "=" + Split(rangeName.RefersTo, "]")(1)
        End If
        If InStr(1, rangeName.RefersTo, "='") > 0 Then
            rangeName.RefersTo = "=" + Split(rangeName.RefersTo, "'")(1)
        End If
        If InStr(1, rangeName.RefersTo, "#REF!") > 0 Then
            ThisWorkbook.Names(rangeName.Index).Delete
        End If
    Next
End Sub


Public Sub updateLevelCellColor()
    Dim rng As Range: Set rng = Range("LEVEL_LIST")
    Dim sLevel As String: sLevel = getInfo("LEVEL")
    Dim i As Integer
    For i = 1 To rng.Cells.Count
        If rng(i).Value2 = sLevel Then
            Range("LEVEL").Interior.color = rng(i).Interior.color
            Range("LEVEL").Font.color = rng(i).Font.color
            Exit Sub
        End If
    Next i
End Sub


Private Function getWorkbookByPath(sPath As String) As Workbook
    Dim i As Integer
    For i = 1 To Workbooks.Count
        If Workbooks(i).FullName = sPath Then
            Set getWorkbookByPath = Workbooks(i)
            Exit Function
        End If
    Next
End Function



Public Function loadExcelSheet()
    Dim RT As String: RT = getInfo("REPORT_TYPE")
    If MsgBox("Do you switch to the ReportType " & RT & " ?" & vbNewLine & "/!\ You will lost all information from this excel !!!", vbYesNo + vbQuestion + vbSystemModal, "PowerAuditor") = vbNo Then
        loadExcelSheet = False
        Exit Function
    End If
    loadExcelSheet = True
    Application.DisplayAlerts = False
    Dim ws_main As Worksheet: Set ws_main = ThisWorkbook.Worksheets("PowerAuditor")
       
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.name <> "PowerAuditor" Then
            ws.Delete
        End If
    Next ws
    
    Dim sPath As String: sPath = getPowerAuditorPath() & "\template\" & RT
    If IOFile.isFile(sPath & ".xlsx") Then
        sPath = sPath & ".xlsx"
    ElseIf IOFile.isFile(sPath & ".xlsm") Then
        sPath = sPath & ".xlsm"
    Else
        Call MsgBox("Template not found !", vbSystemModal + vbCritical, "PowerAuditor")
        Exit Function
    End If
    
    Dim wb As Workbook: Set wb = Workbooks.Open(fileName:=sPath, ReadOnly:=True, Notify:=False, AddToMru:=False, CorruptLoad:=xlNormalLoad)
    sPath = wb.FullName
    Set wb = Nothing
    Application.Wait (Now + TimeValue("0:00:02"))
    ' Copy des feuilles excel
    On Error GoTo reTry_loadExcelSheet
    For Each ws In Xls.getWorkbookByPath(sPath).Worksheets
reTry_loadExcelSheet:
        Xls.getWorkbookByPath(sPath).Worksheets(ws.name).Copy After:=ws_main
    Next ws
    On Error GoTo 0
    ' Suppression des r�f�rences invalides
    Call Xls.cleanUpInvalidExcelRef
    Dim rangeName As name
    ' Copy des Range nomm�s
    For Each rangeName In Xls.getWorkbookByPath(sPath).Names
        If InStr(1, rangeName.RefersTo, "#REF!") = 0 And InStr(1, rangeName.RefersTo, ".xlsm]") = 0 Then
            ThisWorkbook.Names.Add name:=rangeName.name, RefersTo:=rangeName.RefersTo
        End If
    Next
    Xls.getWorkbookByPath(sPath).Close
    Application.DisplayAlerts = True
    With ws_main.Range("LEVEL")
        .Value2 = ""
        .Interior.color = 0
        .Font.color = 0
    End With
    Call CustomRibbonTab.invalidAlltext
    Versionning.G_bDisableExportVBCode = True
    Call ws_main.Parent.Save
    Call Versionning.loadModule(Split(getInfo("REPORT_TYPE"), "-")(0))
    Call ws_main.Parent.Save
    Versionning.G_bDisableExportVBCode = False
End Function


'===============================================================================
' @brief Search in the column {iCol} of the Worksheet {ws} for the value {sSearch} (INSENSITIVE CASE)
' and return the number of occurence.
' @param[in] ws        {Worksheet} The sheet to use
' @param[in] iCol      {int} The column to lookat
' @param[in] sSearch   {String} The string to search (INSENSITIVE CASE)
' @return {int} The number of occurence
Public Function countOccurenceInCol(ws As Worksheet, iCol As Integer, sSearch As String, Optional bPartialMode As Boolean = False) As Integer
    Dim iRow As Integer
    iRow = 3
    countOccurenceInCol = 0
    If bPartialMode = False Then
        While ws.Cells(iRow, iCol).Value2 <> ""
            If StrComp(ws.Cells(iRow, iCol).Value2, sSearch, vbTextCompare) = 0 Then
                countOccurenceInCol = countOccurenceInCol + 1
            End If
            iRow = iRow + 1
        Wend
    Else
        While ws.Cells(iRow, iCol).Value2 <> ""
            If InStr(1, ws.Cells(iRow, iCol).Value2, sSearch, vbTextCompare) = 0 Then
                countOccurenceInCol = countOccurenceInCol + 1
            End If
            iRow = iRow + 1
        Wend
    End If
End Function



Public Sub updateTemplateList()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("PowerAuditor")
    Dim rReportTypeTbl As Range: Set rReportTypeTbl = ws.Range("REPORT_TYPE_TBL[REPORT TYPE]")
    Dim iRow As Integer: iRow = rReportTypeTbl.row
    Dim iCol As Integer: iCol = rReportTypeTbl.Column
    Dim sPath As String: sPath = IOFile.getPowerAuditorPath() & "\template\"
    Dim pFile: pFile = Dir(sPath & "*.xlsm")
    Dim sTmp As String
    Call Xls.cleanupTemplateList
    While pFile <> ""
        sTmp = Replace(pFile, ".xlsm", "")
        If Not isValueInExcelRange(sTmp, rReportTypeTbl) Then
            ws.Cells(iRow, iCol).Value2 = sTmp
            iRow = iRow + 1
        End If
        pFile = Dir
    Wend
    Exit Sub
End Sub


Public Sub cleanupTemplateList()
    ThisWorkbook.Worksheets("PowerAuditor").Range("REPORT_TYPE_TBL[REPORT TYPE]").Value2 = ""
End Sub


Public Function isValueInExcelRange(val As String, rng As Range) As Boolean
    Dim i As Integer
    For i = 1 To rng.Count
        If rng(i).Value2 = val Then
            isValueInExcelRange = True
            Exit Function
        End If
    Next i
    isValueInExcelRange = False
End Function



Public Sub exportPowerauditorToXlsx(Optional aSheetsNewName As Variant = Nothing)
    ThisWorkbook.Save ' Save this document in case of excel segfault
    Dim ws_ex As Workbook
    Set ws_ex = Workbooks.Add
    Dim ws As Worksheet
    Dim sFilename As String: sFilename = ThisWorkbook.Path & "\output\" & RT.getExcelFilename() & "xlsx"
    Dim sCorp As String: sCorp = RT.getCorp
    Dim iRow As Integer: iRow = 1
    If IsMissing(aSheetsNewName) Then aSheetsNewName = Array()
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.name <> "PowerAuditor" Then
            ws.Copy After:=ws_ex.Worksheets(ws_ex.Worksheets.Count)
            If UBound(aSheetsNewName) < ws_ex.Worksheets.Count Then
                ws_ex.Worksheets(ws_ex.Worksheets.Count).name = aSheetsNewName(ws_ex.Worksheets.Count - 2)
            End If
        End If
    Next ws
    
    Debug.Print "Setting file properties for " & sFilename
    With ws_ex
        .BuiltinDocumentProperties("Title") = "Security audit of " & getInfo("TARGET") & " for " & getInfo("CLIENT") & " by " & sCorp & " v" & Xls.getVersionDate()
        .BuiltinDocumentProperties("Subject") = .BuiltinDocumentProperties("Title")
        .BuiltinDocumentProperties("Author") = getFromO365("FriendlyName")
        .BuiltinDocumentProperties("Manager") = RT.getManager()
        .BuiltinDocumentProperties("Company") = sCorp
        .BuiltinDocumentProperties("Category") = "Audit Documents"
        .BuiltinDocumentProperties("Keywords") = sCorp & ", Audit, " & getInfo("TARGET") & ", " & getInfo("CLIENT")
        .BuiltinDocumentProperties("Comments") = .BuiltinDocumentProperties("Title")
    End With
    
    Call IOFile.myMkDir(ThisWorkbook.Path & "\output\")
    Application.DisplayAlerts = False
    ws_ex.Worksheets(1).Delete
    With ws_ex.Worksheets(1)
        .rows(2).EntireRow.Delete
        While .Cells(iRow, 1).Value2 <> ""
            iRow = iRow + 1
        Wend
        .rows(iRow).EntireRow.Delete
    End With
    ws_ex.SaveAs sFilename, FileFormat:=xlOpenXMLWorkbook
    ws_ex.Close
    Application.DisplayAlerts = True
End Sub


Public Function getVersionDate() As String
    Dim VERSION_DATE As String: VERSION_DATE = getInfo("VERSION_DATE")
    If Common.isEmptyString(VERSION_DATE) Then
        VERSION_DATE = Year(Date) & "-"
        If Month(Date) < 10 Then VERSION_DATE = VERSION_DATE & "0"
        VERSION_DATE = VERSION_DATE & Month(Date) & "-"
        If Day(Date) < 10 Then VERSION_DATE = VERSION_DATE & "0"
        VERSION_DATE = VERSION_DATE & Day(Date)
    End If
    getVersionDate = VERSION_DATE
End Function


Public Function nbNotEmptyRows(ws As Worksheet, iCol As Integer) As Integer
    nbNotEmptyRows = 3
    While ws.Cells(nbNotEmptyRows, iCol).Value2 <> ""
        nbNotEmptyRows = nbNotEmptyRows + 1
    Wend
    nbNotEmptyRows = nbNotEmptyRows - 3
End Function


Public Function getWorksheetRT() As Worksheet
    Dim sRT As String: sRT = Common.getInfo("REPORT_TYPE")
    If Common.isEmptyString(sRT) Then GoTo getWorksheetRT_fail
    Set getWorksheetRT = ThisWorkbook.Worksheets(sRT)
    Exit Function
getWorksheetRT_fail:
    Set getWorksheetRT = Nothing
    ThisWorkbook.Worksheets("PowerAuditor").Activate
    ThisWorkbook.Worksheets("PowerAuditor").Range("REPORT_TYPE").Select
    Call MsgBox("Please, select a >Report Type< before", vbInformation + vbOKOnly + vbSystemModal, "PowerAuditor")
End Function
