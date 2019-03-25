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


Function UFirstLetter(sStr As String) As String
    UFirstLetter = UCase(Left(sStr, 1)) & LCase(Mid(sStr, 2))
End Function


Function CleaupScoreMesg(sStr As String) As String
    If InStr(sStr, " - ") Then
        sStr = Split(sStr, " - ")(1)
    End If
    CleaupScoreMesg = UFirstLetter(sStr)
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
    isDevMode = IOFile.isFolder(getVBAPath())
End Function


Public Function PowerAuditorPath() As String
    If isDevMode() Then
        PowerAuditorPath = ThisWorkbook.Path & "\..\"
    Else
        PowerAuditorPath = Environ("USERPROFILE") & "\PowerAuditor\"
    End If
End Function


Public Function VulnDBPath(name As String) As String
    VulnDBPath = Common.PowerAuditorPath() & "\VulnDB\" & name
End Function


Public Function getVBAPath() As String
    getVBAPath = ThisWorkbook.Path & "\src\vba\"
End Function




Function ArrayAppendUniq(arr As Variant, val As String) As Variant
    Dim i As Integer
    For i = 0 To UBound(arr)
        If arr(i) = val Then
            ArrayAppendUniq = arr
            Exit Function
        End If
    Next
    ReDim Preserve arr(0 To UBound(arr) + 1) As Variant
    arr(UBound(arr)) = val
    ArrayAppendUniq = arr
End Function


Public Function DictContains(col As Collection, key As Variant) As Boolean
    On Error Resume Next
    col (key) ' Just try it. If it fails, Err.Number will be nonzero.
    DictContains = (Err.Number = 0)
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


Public Sub trimContentControl(cc)
    Dim found As Boolean: found = True
    Dim wDoc: Set wDoc = cc.Range.Document
    Dim rng, rng2, rng3
    While found
        Set rng = wDoc.Range(cc.Range.End - 1, cc.Range.End)
        Set rng2 = wDoc.Range(cc.Range.End - 2, cc.Range.End)
        If isEmptyString(rng.text) And isEmptyString(rng2.text) Then
            rng.text = ""
        Else
            found = False
        End If
    Wend
End Sub


Public Function isEmptyString(ByVal myStr As String) As Boolean
    myStr = Common.trim(myStr, Chr(10) & Chr(13) & " ")
    isEmptyString = (IsEmpty(myStr) Or myStr = " " Or myStr = vbNewLine Or myStr = vbLf Or myStr = vbCr Or myStr = "")
End Function





Public Function getFromO365(sType As String, Optional bFromO365) As String
    Dim wsh As Object: Set wsh = VBA.CreateObject("WScript.Shell")
    Dim waitOnReturn As Boolean: waitOnReturn = True
    Dim windowStyle As Integer: windowStyle = 0
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
        wsh.Run "reg export HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\Identity\Identities %tmp%\O365.tmp /y", windowStyle, waitOnReturn
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


Sub createExcelValidatorList(rng As Range, list As Variant)
    'Dim MyList As Variant: MyList = Array(1, "toto", 3, 4, 5, 6)
    With rng.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:=Join(list, ",")
    End With
End Sub


Sub setReportTypeList(RT As String)
    ThisWorkbook.Worksheets("PowerAuditor").Range("REPORT_TYPE_LIST") = ""
    Dim aList As Variant: aList = Split(RT, ",")
    With ThisWorkbook.Names.Item("REPORT_TYPE_LIST")
        .RefersTo = .RefersToRange.Resize(UBound(aList) + 1)
    End With
    ThisWorkbook.Worksheets("PowerAuditor").Range("REPORT_TYPE_LIST") = Application.Transpose(aList)
End Sub


Public Sub CleanUpInvalidExcelRef()
    ' Nétoyage des noms de range
    Dim rangeName As name
    For Each rangeName In ThisWorkbook.Names
        If InStr(1, rangeName.RefersTo, "#REF!") > 0 Then
            ThisWorkbook.Names(rangeName.index).Delete
        End If
    Next
End Sub


Public Sub UpdateLevelCellColor()
    Dim rng As Range: Set rng = Range("LEVEL_LIST")
    Dim sLevel As String: sLevel = getInfo("LEVEL")
    Dim i As Integer
    For i = 1 To rng.Cells.count
        If rng(i).Value2 = sLevel Then
            Range("LEVEL").Interior.color = rng(i).Interior.color
            Range("LEVEL").Font.color = rng(i).Font.color
            Exit Sub
        End If
    Next i
End Sub


Public Sub LoadExcelSheet()
    Dim RT As String: RT = getInfo("REPORT_TYPE") & "-" & getInfo("LANG")
    If MsgBox("Do you switch to the ReportType " & RT & " ?" & vbNewLine & "/!\ You will lost all information from this excel !!!", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    Application.DisplayAlerts = False
    
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.name <> "PowerAuditor" Then
            ws.Delete
        End If
    Next ws
    
    Dim sPath As String: sPath = PowerAuditorPath() & "\template\" & RT
    If IOFile.isFile(sPath & ".xlsx") Then
        sPath = sPath & ".xlsx"
    ElseIf IOFile.isFile(sPath & ".xlsm") Then
        sPath = sPath & ".xlsm"
    Else
        MsgBox ("Template not found !")
        Exit Sub
    End If
    
    
    Dim wb As Workbook: Set wb = Workbooks.Open(sPath)
    ' Copy des feuilles excel
    For Each ws In wb.Worksheets
        wb.Worksheets(ws.name).Copy after:=ThisWorkbook.Worksheets("PowerAuditor")
    Next ws
    ' Suppression des références invalides
    Call CleanUpInvalidExcelRef
    Dim rangeName As name
    ' Copy des Range nommés
    For Each rangeName In wb.Names
        If InStr(1, rangeName.RefersTo, "#REF!") = 0 And InStr(1, rangeName.RefersTo, ".xlsm]") = 0 Then
            ThisWorkbook.Names.Add name:=rangeName.name, RefersTo:=rangeName.RefersTo
        End If
    Next
    wb.Close
    Application.DisplayAlerts = True
    Call Versionning.loadModule(getInfo("REPORT_TYPE"))
    With Range("LEVEL")
        .Value2 = ""
        .Interior.color = 0
        .Font.color = 0
    End With
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




Public Function RandomString(Length As Integer)
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
    RandomString = str
End Function


Function selectCCInCC(oWRange As Object, sTitle As String) As Object
    Dim i As Integer
    For i = 1 To oWRange.ContentControls.count
        If oWRange.ContentControls(i).title = sTitle Then
            Set selectCCInCC = oWRange.ContentControls(i)
            Exit Function
        End If
    Next
    Set selectCCInCC = Nothing
End Function




Public Function cleanupFileName(ByVal sFilename As String) As String
    ' Suppression des espaces et \r\n au début du nom de fichier
    While Left(sFilename, 1) = " " Or Left(sFilename, 1) = vbCr Or Left(sFilename, 1) = vbLf
        sFilename = Mid(sFilename, 2)
    Wend
    ' Suppression de la numérotation
    While IsNumeric(Left(sFilename, 1))
        sFilename = Mid(sFilename, 2)
    Wend
    ' Suppression des tirets et des espaces
    While Left(sFilename, 1) = "." Or Left(sFilename, 1) = "-" Or Left(sFilename, 1) = " "
        sFilename = Mid(sFilename, 2)
    Wend
    cleanupFileName = sFilename
End Function


Public Function CVSSReader(cvss As String) As String
    'Run a shell command, returning the output as a string
    Dim oWSH As Object: Set oWSH = VBA.CreateObject("WScript.Shell")
    Dim sTmpFile As String: sTmpFile = Environ("TMP") & "\" & RandomString(10)
    oWSH.Run "cmd /c P:\Dev\PowerAuditor\bin\CVSSEditor.exe " & cvss & " > " & sTmpFile, 0, True
    CVSSReader = fileGetContent(sTmpFile)
    Kill sTmpFile
End Function


