Attribute VB_Name = "Common"
Option Explicit
' PowerAuditor - A simple script to help report writing by https://github.com/1mm0rt41PC
'
' Filename: Common.bas.vb
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
Private m_isDevMode As Integer


Function uFirstLetter(sStr As String) As String
    uFirstLetter = UCase(Left(sStr, 1)) & LCase(Mid(sStr, 2))
End Function


Function cleaupScoreMesg(sStr As String) As String
    If InStr(sStr, " - ") Then
        sStr = Split(sStr, " - ")(1)
    End If
    cleaupScoreMesg = uFirstLetter(sStr)
End Function


Function getScoreValue(sMsg As String) As Integer
    getScoreValue = Split(sMsg, " - ")(0)
End Function


Function getScoreValue4Cell(ws As Worksheet, line As Integer, iCol As Integer) As Integer
    getScoreValue4Cell = getScoreValue(ws.Cells(line, iCol).Value2)
End Function


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


Public Function isDevMode() As Boolean
    If Common.m_isDevMode <= 0 Then
        Common.m_isDevMode = IOFile.isFolder(getVBAPath())
        If Common.m_isDevMode Then
            Common.m_isDevMode = 1
        Else
            Common.m_isDevMode = 2
        End If
    End If
    isDevMode = (Common.m_isDevMode = 1)
End Function


Function arrayAppendUniq(arr As Variant, val As String) As Variant
    Dim i As Integer
    For i = 0 To UBound(arr)
        If arr(i) = val Then
            arrayAppendUniq = arr
            Exit Function
        End If
    Next
    ReDim Preserve arr(0 To UBound(arr) + 1) As Variant
    arr(UBound(arr)) = val
    arrayAppendUniq = arr
End Function


Public Function dictContains(col As Collection, key As Variant) As Boolean
    On Error Resume Next
    col (key) ' Just try it. If it fails, Err.Number will be nonzero.
    dictContains = (Err.Number = 0)
    Err.Clear
End Function


Public Function trim(myStr As String, Optional rmChar As String) As String
    Dim strStart As Integer: strStart = 1
    Dim strEnd As Integer: strEnd = Len(myStr)
    Dim hasChanged As Boolean: hasChanged = True
    Dim i As Integer
    Dim old As String
    If IsMissing(rmChar) Or rmChar = "" Then
        rmChar = vbCr & vbLf & " "
    End If
    
    While hasChanged
        hasChanged = False
        For i = 1 To Len(rmChar)
            If Mid(myStr, 1, 1) = Mid(rmChar, i, 1) Then
                myStr = Mid(myStr, 2)
                hasChanged = True
            End If
            If Right(myStr, 1) = Mid(rmChar, i, 1) Then
                myStr = Mid(myStr, 1, Len(myStr) - 1)
                hasChanged = True
            End If
        Next
    Wend
    trim = myStr
End Function


Public Function isEmptyString(ByVal myStr As String) As Boolean
    myStr = Common.trim(myStr, Chr(10) & Chr(13) & " ")
    isEmptyString = (IsEmpty(myStr) Or myStr = " " Or myStr = vbNewLine Or myStr = vbLf Or myStr = vbCr Or myStr = "")
End Function


Public Function getFromO365(sType As String, Optional bFromO365) As String
    Dim iFileNum As Integer
    Dim sDataLine As String
    Dim sFilepath As String
    
    If IsMissing(bFromO365) Then
        sDataLine = getFromO365(sType, False)
        If sDataLine = "" Then sDataLine = getFromO365(sType, True)
        getFromO365 = sDataLine
        Exit Function
    End If
    
    If bFromO365 Then
        Call IOFile.removeFile(Environ("tmp") & "\O365.tmp")
        Call IOFile.runCmd("reg export HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\Identity\Identities %tmp%\O365.tmp /y", 0, True)
        sFilepath = Environ("tmp") & "\O365.tmp"
    Else
        sFilepath = Environ("USERPROFILE") & "\PowerAuditor\config.ini"
        If Not IOFile.isFile(sFilepath) Then
            getFromO365 = ""
            Exit Function
        End If
    End If
    If Not IOFile.isFile(sFilepath) Then
        getFromO365 = ""
        Exit Function
    End If
    iFileNum = FreeFile()
    Open sFilepath For Input As #iFileNum

    While Not EOF(iFileNum)
        Line Input #iFileNum, sDataLine ' read in data 1 line at a time
        If InStr(sDataLine, sType) > 0 Then
            '"EmailAddress"="xxxx@yyyyy.zzz"
            sDataLine = Replace(sDataLine, Chr(34), "")
            If bFromO365 Then Call IOFile.fileAppend(Environ("USERPROFILE") & "\PowerAuditor\config.ini", sDataLine)
            sDataLine = Replace(sDataLine, sType & "=", "")
            If sDataLine <> "" And sDataLine <> " " Then
                getFromO365 = sDataLine
                Close #iFileNum
                If bFromO365 Then Call IOFile.removeFile(Environ("tmp") & "\O365.tmp")
                Exit Function
            End If
        End If
    Wend
    Close #iFileNum
    If bFromO365 Then Call IOFile.removeFile(Environ("tmp") & "\O365.tmp")
    getFromO365 = ""
End Function


Function getInfo(rng As String) As String
    getInfo = ThisWorkbook.Worksheets("PowerAuditor").Range(rng).Value2
End Function


Sub setReportTypeList(RT As String)
    Call cleanupTemplateList
    Dim aList As Variant: aList = Split(RT, ",")
    ThisWorkbook.Worksheets("PowerAuditor").Range("REPORT_TYPE_TBL[REPORT TYPE]") = Application.Transpose(aList)
End Sub


Public Sub cleanUpInvalidExcelRef()
    ' Nétoyage des noms de range
    Dim rangeName As name
    For Each rangeName In ThisWorkbook.Names
        If InStr(1, rangeName.RefersTo, "#REF!") > 0 Then
            ThisWorkbook.Names(rangeName.index).Delete
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



Public Sub loadExcelSheet()
    Dim RT As String: RT = getInfo("REPORT_TYPE")
    If MsgBox("Do you switch to the ReportType " & RT & " ?" & vbNewLine & "/!\ You will lost all information from this excel !!!", vbYesNo + vbQuestion) = vbNo Then Exit Sub
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
        MsgBox ("Template not found !")
        Exit Sub
    End If
    
    Dim wb As Workbook: Set wb = Workbooks.Open(fileName:=sPath, ReadOnly:=True, Notify:=False, AddToMru:=False, CorruptLoad:=xlNormalLoad)
    sPath = wb.FullName
    Set wb = Nothing
    Application.Wait (Now + TimeValue("0:00:02"))
    ' Copy des feuilles excel
    On Error GoTo reTry_loadExcelSheet
    For Each ws In getWorkbookByPath(sPath).Worksheets
reTry_loadExcelSheet:
        getWorkbookByPath(sPath).Worksheets(ws.name).Copy after:=ws_main
    Next ws
    On Error GoTo 0
    ' Suppression des références invalides
    Call cleanUpInvalidExcelRef
    Dim rangeName As name
    ' Copy des Range nommés
    For Each rangeName In getWorkbookByPath(sPath).Names
        If InStr(1, rangeName.RefersTo, "#REF!") = 0 And InStr(1, rangeName.RefersTo, ".xlsm]") = 0 Then
            ThisWorkbook.Names.Add name:=rangeName.name, RefersTo:=rangeName.RefersTo
        End If
    Next
    getWorkbookByPath(sPath).Close
    Application.DisplayAlerts = True
    With ws_main.Range("LEVEL")
        .Value2 = ""
        .Interior.color = 0
        .Font.color = 0
    End With
    Call Versionning.loadModule(getInfo("REPORT_TYPE"))
    Call ws_main.Parent.Save
End Sub


'===============================================================================
' @brief Search in the column {col} of the Worksheet {ws} for the value {sVal} (INSENSITIVE CASE)
' and return the number of occurence.
' @param[in] ws     {Worksheet} The sheet to use
' @param[in] iCol   {int} The column to lookat
' @param[in] sVal   {String} The string to search (INSENSITIVE CASE)
' @return {int} The number of occurence
Public Function countOccurenceInCol(ws As Worksheet, iCol As Integer, sVal As String) As Integer
    Dim iRow As Integer
    Dim ret As Integer: ret = 0
    
    sVal = LCase(sVal)
    
    iRow = 3
    While Not IsEmpty(ws.Cells(iRow, 1).Value2)
        If LCase(ws.Cells(iRow, iCol).Value2) = sVal Then
            ret = ret + 1
        End If
        iRow = iRow + 1
    Wend
    countOccurenceInCol = ret
End Function




Public Function randomString(Length As Integer)
    'PURPOSE: Create a Randomized String of Characters
    'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault
    
    Dim CharacterBank As Variant
    Dim x As Long
    Dim str As String
    
    'Test Length Input
    If Length < 1 Then
        MsgBox "Length variable must be greater than 0"
        Exit Function
    End If
    
    CharacterBank = Array("a", "b", "c", "d", "e", "f", "g", "h", "i", "j", _
    "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", _
    "y", "z", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "_", "_", _
    "_", "_", "%", "_", "_", "_", "A", "B", "C", "D", "E", "F", "G", "H", _
    "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", _
    "W", "X", "Y", "Z")
    
    
    'Randomly Select Characters One-by-One
    For x = 1 To Length
        Randomize
        str = str & CharacterBank(Int((UBound(CharacterBank) - LBound(CharacterBank) + 1) * Rnd + LBound(CharacterBank)))
    Next x
    
    'Output Randomly Generated String
    randomString = str
End Function


Public Function CVSSReader(cvss As String) As String
    Dim sTmpFile As String: sTmpFile = Environ("TMP") & "\" & randomString(10)
    Call IOFile.runCmd("cmd /c " & Common.getPowerAuditorPath() & "\bin\CVSSEditor.exe " & cvss & " > " & sTmpFile, 0, True)
    CVSSReader = fileGetContent(sTmpFile)
    Call IOFile.removeFile(sTmpFile)
End Function


Public Sub updateTemplateList()
    Dim ws As Worksheet: Set ws = Worksheets("PowerAuditor")
    Dim rReportTypeTbl As Range: Set rReportTypeTbl = ws.Range("REPORT_TYPE_TBL[REPORT TYPE]")
    Dim iRow As Integer: iRow = rReportTypeTbl.Row
    Dim iCol As Integer: iCol = rReportTypeTbl.Column
    Dim sPath As String: sPath = IOFile.getPowerAuditorPath() & "\template\"
    Dim pFile: pFile = Dir(sPath & "*.xlsm")
    Dim sTmp As String
    Call cleanupTemplateList
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
    Range("REPORT_TYPE_TBL[REPORT TYPE]").Value2 = ""
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


Public Function getLang() As String
    getLang = Split(Worksheets(2).name, "-")(1)
End Function

