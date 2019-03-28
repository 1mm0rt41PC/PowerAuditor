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

Public poney As String
Dim G_oWSH As Object

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
            Call MyMkDir(ThisWorkbook.Path & Application.PathSeparator & "output")
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




Public Function git(sArgs As String, Optional sRepo As String = "vulndb", Optional iRecurs As Integer = 5) As Boolean
    Debug.Print "Git " & sArgs & " on <" & sRepo & ">"
    Dim tmpFile As String: tmpFile = Environ("temp") & "\" & RandomString(7)
    Dim ret As String
    If iRecurs <= 0 Then
        MsgBox "Git is not installed or not initialised !?", vbOKOnly, "PowerAuditor"
        git = ""
        Exit Function
    End If
    If G_oWSH Is Nothing Then
        Set G_oWSH = VBA.CreateObject("WScript.Shell")
    End If
    On Error GoTo reloadWSH
    G_oWSH.Run Common.PowerAuditorPath() & "\" & sRepo & "_git.bat " & tmpFile & " " & sArgs, 0, True
    On Error GoTo 0
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
    Exit Function
reloadWSH:
    On Error GoTo 0
    Set G_oWSH = Nothing
    git = git(sArgs, sRepo, iRecurs - 1)
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

Public Function MyMkDir(sPath As String)
    If Not IOFile.isFolder(sPath) Then
        Call MkDir(sPath)
    End If
End Function


Public Function getOutpoutFromShellCmd(sCmd As String) As String
    'Run a shell command, returning the output as a string
    Dim oShell As Object: Set oShell = CreateObject("WScript.Shell")
    Dim oExec As Object
    Dim oOutput As Object
    
    Set oExec = oShell.Exec(sCmd)
    Set oOutput = oExec.StdOut

    'handle the results as they are written to and read from the StdOut object
    Dim s As String
    Dim sLine As String
    While Not oOutput.AtEndOfStream
        sLine = oOutput.ReadLine
        If sLine <> "" Then s = s & sLine & vbCrLf
    Wend
    ShellRun = s
End Function
