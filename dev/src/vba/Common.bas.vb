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
Private m_getLang As String


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


Public Function isDevMode() As Boolean
    If Common.m_isDevMode <= 0 Then
        Common.m_isDevMode = IOFile.isFolder(IOFile.getVBAPath())
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


Public Function randomString(Length As Integer)
    'PURPOSE: Create a Randomized String of Characters
    'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault
    
    Dim CharacterBank As Variant
    Dim x As Long
    Dim str As String
    
    'Test Length Input
    If Length < 1 Then
        MsgBox "Length variable must be greater than 0", vbSystemModal + vbCritical + vbOKOnly, "PowerAuditor"
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
    Call IOFile.runCmd("cmd /c " & IOFile.getPowerAuditorPath() & "\bin\CVSSEditor.exe " & cvss & " > " & sTmpFile, 0, True)
    On Error GoTo CVSSReader_err
    CVSSReader = fileGetContent(sTmpFile)
    Call IOFile.removeFile(sTmpFile)
    Exit Function
CVSSReader_err:
    CVSSReader = ""
    Call IOFile.removeFile(sTmpFile)
End Function


Public Function getLang() As String
    If m_getLang = "" Then m_getLang = UCase(Split(ThisWorkbook.Worksheets(2).name, "-")(1))
    getLang = m_getLang
End Function

