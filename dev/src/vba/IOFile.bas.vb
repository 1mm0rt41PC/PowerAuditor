Attribute VB_Name = "IOFile"
Option Explicit
' PowerAuditor - A simple script to help report writing by https://github.com/1mm0rt41PC
'
' Filename: IOFile.bas.vb
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

Dim G_oWSH As Object
Private m_getPowerAuditorPath As String


Public Function filenameEncode(ByVal sFilename As String) As String
    Dim i As Integer
    Dim sRet As String
    Dim cTmp As String
    Dim dec As Variant: dec = Array("<", ">", ":", Chr(34), "/", "\", "|", "?", "*")  ' Chr(34) = "
    Dim enc As Variant: enc = Array(60, 62, 58, 34, 47, 92, 124, 63, 42)
    
    For i = 1 To UBound(dec)
        sFilename = Replace(sFilename, dec(i), "%" & enc(i))
    Next i
    filenameEncode = sFilename
End Function


Public Function filenameDecode(sFilename As String) As String
    Dim i As Integer
    Dim sRet As String
    Dim cTmp As String
    Dim dec As Variant: dec = Array("<", ">", ":", Chr(34), "/", "\", "|", "?", "*")  ' Chr(34) = "
    Dim enc As Variant: enc = Array(60, 62, 58, 34, 47, 92, 124, 63, 42)
    
    For i = 1 To UBound(dec)
        sFilename = Replace(sFilename, "%" & enc(i), dec(i))
    Next i
    filenameDecode = sFilename
End Function


Public Sub removeFile(mFile As String)
    On Error GoTo removeFile_err
    Kill mFile
    On Error GoTo 0
    Exit Sub
removeFile_err:
    On Error GoTo 0
    Debug.Print "[!] Unable to remove the file <" & mFile & ">"
End Sub


Public Function renameDocument(inst, ext As String, pType As String, Optional ByVal deleteOld As Boolean) As String
    Dim newFileName As String
    
    If IsMissing(deleteOld) Then
        deleteOld = True
    End If
    
    ' Renomage automatique des fichiers
    Dim fileName As String: fileName = RT.getReportFilename(pType) & ext
    Dim corp As String: corp = RT.getCorp()
    If InStr(ThisWorkbook.Path, "output") Then
        If pType = "TEMPLATE" Then
            newFileName = ThisWorkbook.Path & Application.PathSeparator & ".." & Application.PathSeparator & fileName
        Else
            newFileName = ThisWorkbook.Path & Application.PathSeparator & fileName
        End If
    Else
        If pType = "TEMPLATE" Then
            newFileName = ThisWorkbook.Path & Application.PathSeparator & fileName
        Else
            newFileName = ThisWorkbook.Path & Application.PathSeparator & "output" & Application.PathSeparator & fileName
            Call myMkDir(ThisWorkbook.Path & Application.PathSeparator & "output")
        End If
    End If
    
    Debug.Print "Setting file properties for " & fileName
    With inst
        .BuiltinDocumentProperties("Title") = "Security audit of " & getInfo("TARGET") & " for " & getInfo("CLIENT") & " by " & corp & " v" & getInfo("VERSION_DATE")
        .BuiltinDocumentProperties("Subject") = .BuiltinDocumentProperties("Title")
        .BuiltinDocumentProperties("Author") = getFromO365("FriendlyName")
        .BuiltinDocumentProperties("Manager") = RT.getManager()
        .BuiltinDocumentProperties("Company") = corp
        .BuiltinDocumentProperties("Category") = "Audit Documents"
        .BuiltinDocumentProperties("Keywords") = corp & ", Audit, " & getInfo("TARGET") & ", " & getInfo("CLIENT")
        .BuiltinDocumentProperties("Comments") = .BuiltinDocumentProperties("Title")
    End With
    
    Application.DisplayAlerts = False
    If newFileName <> inst.FullName Then
        Debug.Print "Renaming file <" & inst.name & "> into " & fileName
        Dim oldFileName As String: oldFileName = inst.FullName
        If ext = "xlsm" Then
            inst.SaveAs newFileName, FileFormat:=xlOpenXMLWorkbookMacroEnabled ' = 52
        ElseIf ext = "xlsx" Then
            inst.SaveAs newFileName, FileFormat:=xlOpenXMLWorkbook ' = 51
        Else
            inst.SaveAs newFileName
        End If
        If deleteOld Then removeFile oldFileName
    Else
        inst.Save
    End If
    Application.DisplayAlerts = True
    
    renameDocument = newFileName
End Function




Public Function git(sArgs As String, Optional sRepo As String = "vulndb") As Boolean
    Debug.Print "Git " & sArgs & " on <" & sRepo & ">"
    Dim tmpFile As String: tmpFile = Environ("temp") & "\" & randomString(7)
    Dim ret As String
    Call runCmd(IOFile.getPowerAuditorPath() & "\" & sRepo & "_git.bat " & tmpFile & " " & sArgs, 0, True)
    ret = Common.trim(fileGetContent(tmpFile & ".ret"))
    If ret <> "0" Then
        ret = Common.trim(fileGetContent(tmpFile & ".log"))
        If InStr(1, ret, "nothing to commit, working tree clean") = 0 Then
            MsgBox "The command git " & sArgs & vbCrLf & "Returned:" & vbCrLf & ret, vbOKOnly, "Error with git"
            git = False
        Else
            git = True
        End If
    Else
        git = True
    End If
    ' Then delete the file
    removeFile tmpFile & ".ret"
    removeFile tmpFile & ".log"
End Function


Public Function isFile(sPath As String) As Boolean
    isFile = CreateObject("Scripting.FileSystemObject").fileExists(sPath)
End Function

Public Function isFolder(sPath As String) As Boolean
    isFolder = CreateObject("Scripting.FileSystemObject").folderExists(sPath)
End Function




Public Sub fileAppend(sFilename As String, sData As String)
    Dim iFileNum As Integer: iFileNum = FreeFile()
    Open sFilename For Append As #iFileNum
    Print #iFileNum, sData
    Close #iFileNum
End Sub

Public Sub fileSetContent(sFilename As String, sData As String)
    Dim iFileNum As Integer: iFileNum = FreeFile()
    Open sFilename For Output As #iFileNum
    Print #iFileNum, sData
    Close #iFileNum
End Sub

Public Function fileGetContent(sFilename As String) As String
    Dim sData As String
    Dim oFSO As Object: Set oFSO = CreateObject("Scripting.FileSystemObject")
    Dim oTF As Object: Set oTF = oFSO.OpenTextFile(sFilename, 1)
    sData = oTF.readall()
    oTF.Close
    fileGetContent = sData
End Function

Public Function getFileExt(mFileName As String) As String
    Dim aPath() As String
    Dim realExt As String
    aPath = Split(mFileName, ".")
    getFileExt = aPath(UBound(aPath))
End Function

Public Function myMkDir(sPath As String)
    If Not IOFile.isFolder(sPath) Then
        Call MkDir(sPath)
    End If
End Function

' Run sCmd in a cmd.exe prompt and return the return_code (%ERROR_LEVEL%)
Public Function runCmd(sCmd As String, Optional iWindowStyle As Integer = 0, Optional bWaitOnReturn As Boolean = True) As Integer
    Dim ret As Integer
    If G_oWSH Is Nothing Then
        Set G_oWSH = VBA.CreateObject("WScript.Shell")
    End If
    On Error GoTo err_runCmd
    ret = G_oWSH.Run(sCmd, iWindowStyle, bWaitOnReturn)
    On Error GoTo 0
    runCmd = ret
    Exit Function

err_runCmd:
    On Error GoTo 0
    Set G_oWSH = VBA.CreateObject("WScript.Shell")
    ret = G_oWSH.Run(sCmd, iWindowStyle, bWaitOnReturn)
    runCmd = ret
    Exit Function
End Function


Public Function getOutpoutFromShellCmd(sCmd As String) As String
    'Run a shell command, returning the output as a string
    If G_oWSH Is Nothing Then
        Set G_oWSH = VBA.CreateObject("WScript.Shell")
    End If
    Dim oExec As Object
    Dim oOutput As Object
    
    On Error GoTo err_getOutpoutFromShellCmd
    Set oExec = G_oWSH.Exec(sCmd)
    Set oOutput = oExec.StdOut
    On Error GoTo 0
    GoTo readStdout_getOutpoutFromShellCmd
    
err_getOutpoutFromShellCmd:
    Set G_oWSH = VBA.CreateObject("WScript.Shell")
    Set oExec = G_oWSH.Exec(sCmd)
    Set oOutput = oExec.StdOut
    On Error GoTo 0

readStdout_getOutpoutFromShellCmd:

    'handle the results as they are written to and read from the StdOut object
    Dim s As String
    Dim sLine As String
    While Not oOutput.AtEndOfStream
        sLine = oOutput.ReadLine
        If sLine <> "" Then s = s & sLine & vbCrLf
    Wend
    getOutpoutFromShellCmd = s
End Function


Public Function getPowerAuditorPath() As String
    If IOFile.m_getPowerAuditorPath = "" Then
        If isDevMode() Then
            IOFile.m_getPowerAuditorPath = ThisWorkbook.Path & "\..\"
        Else
            IOFile.m_getPowerAuditorPath = Environ("USERPROFILE") & "\PowerAuditor\"
        End If
    End If
    getPowerAuditorPath = IOFile.m_getPowerAuditorPath
End Function


Public Function getVulnDBPath(sVulnerabilityName As String, Optional bCreateIfNotExist As Boolean = False) As String
    Dim sEncVulnName As String: sEncVulnName = IOFile.filenameEncode(sVulnerabilityName)
    Dim sPath As String: sPath = IOFile.getPowerAuditorPath() & "\VulnDB\" & Common.getLang() & "\" & sEncVulnName
    Dim oFS As Object
    Dim isFolder As Boolean: isFolder = IOFile.isFolder(sPath)
    If Not isFolder And Not bCreateIfNotExist Then
        getVulnDBPath = ""
        Exit Function
    End If
    Set oFS = CreateObject("scripting.filesystemobject")
    Dim sDesktopIni As String
    sDesktopIni = "[.ShellClassInfo]" & vbCrLf
    sDesktopIni = sDesktopIni & "ConfirmFileOp=1" & vbCrLf
    sDesktopIni = sDesktopIni & "NoSharing=1" & vbCrLf
    sDesktopIni = sDesktopIni & "LocalizedResourceName=" & sVulnerabilityName & vbCrLf
    sDesktopIni = sDesktopIni & "[ViewState]" & vbCrLf
    sDesktopIni = sDesktopIni & "Mode=" & vbCrLf
    sDesktopIni = sDesktopIni & "Vid=" & vbCrLf
    sDesktopIni = sDesktopIni & "FolderType=Generic" & vbCrLf
    sDesktopIni = sDesktopIni & "[DeleteOnCopy]" & vbCrLf
    sDesktopIni = sDesktopIni & "Personalized=5" & vbCrLf
    sDesktopIni = sDesktopIni & "PersonalizedName=" & sVulnerabilityName & vbCrLf
    
    If Not isFolder Then MkDir (sPath)
    Call IOFile.fileSetContent(sPath & "\desktop.ini", sDesktopIni)
    oFS.getfile(sPath & "\desktop.ini").Attributes = 39
    oFS.getfolder(sPath).Attributes = 17
    getVulnDBPath = sPath
End Function


Public Function getNotableFile(name As String) As String
    getNotableFile = IOFile.getPowerAuditorPath() & "\VulnDB\.notable\notes\" & IOFile.filenameEncode(name) & ".md"
End Function


Public Function getVBAPath() As String
    getVBAPath = ThisWorkbook.Path & "\src\vba\"
End Function

