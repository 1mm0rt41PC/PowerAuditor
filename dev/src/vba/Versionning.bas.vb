Attribute VB_Name = "Versionning"
Option Explicit
' PowerAuditor - A simple script to help report writing by https://github.com/1mm0rt41PC
'
' Filename: Versionning.bas.vb
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

Public G_bDisableExportVBCode As Boolean


' Excel macro to export all VBA source code in this project to text files for proper source control versioning
' Requires enabling the Excel setting in Options/Trust Center/Trust Center Settings/Macro Settings/Trust access to the VBA project object model
Public Sub exportVisualBasicCode()
    'Exit Sub
    Const Module = 1
    Const ClassModule = 2
    Const Form = 3
    Const Document = 100
    Const Padding = 24
    
    Dim VBComponent As Object
    Dim Path As String
    Dim sPath As String: sPath = IOFile.getVBAPath
    Dim extension As String
    Dim REPORT_TYPE As String

    If sPath = "" Or Not Common.isDevMode() Or G_bDisableExportVBCode = True Then Exit Sub
    
    ' Require: reg ADD HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Excel\Security /v AccessVBOM /t REG_DWORD /d 1 /f
    For Each VBComponent In ActiveWorkbook.VBProject.VBComponents
        Select Case VBComponent.Type
            Case ClassModule, Document
                extension = ".cls.vb"
            Case Form
                extension = ".frm.vb"
            Case Module
                extension = ".bas.vb"
            Case Else
                extension = ".txt.vb"
        End Select
            
                
        On Error Resume Next
        Err.Clear
        
        If Left(VBComponent.name, 3) = "RT_" Then
            Path = IOFile.getPowerAuditorPath() & "\template\" & VBComponent.name & extension
        Else
            Path = sPath & "\" & VBComponent.name & extension
        End If
        Call VBComponent.Export(Path)
        If Err.Number <> 0 Then
            Call MsgBox("Failed to export " & VBComponent.name & " to " & Path, vbCritical + vbSystemModal, "PowerAuditor")
        Else
            Debug.Print "Exported " & Left$(VBComponent.name & ":" & Space(Padding), Padding) & Path
        End If
        ' List of report type
        If Left(VBComponent.name, 3) = "RT_" Then
            Debug.Print "Found report type: " & Mid(VBComponent.name, 4)
            REPORT_TYPE = REPORT_TYPE & Mid(VBComponent.name, 4) & ","
        End If

        On Error GoTo 0
    Next
End Sub


Public Static Sub VBAFromCommonSrc()
    Dim sFile As String
    Dim sData As String
    Dim pos
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim sPath As String: sPath = IOFile.getVBAPath
    Dim sModuleName As String
    Dim REPORT_TYPE As String

    If sPath = "" Or Not Common.isDevMode() Then Exit Sub
    
    If IOFile.isFile(ThisWorkbook.FullName & ".lock") Then
        Call IOFile.removeFile(ThisWorkbook.FullName & ".lock")
        MsgBox "Unable to call VBAFromCommonSrc, last call has crashed !", vbOKOnly + vbSystemModal + vbCritical, "PowerAuditor"
        Exit Sub
    End If
    Call IOFile.fileSetContent(ThisWorkbook.FullName & ".lock", "")

    Dim pFile: pFile = Dir(sPath & "*.vb")
    Do While pFile <> ""
        sData = Common.trim(fso.OpenTextFile(sPath & pFile, 1).readall())
        pos = InStr(sData, "Option Explicit")
        If pos Then
            sModuleName = Split(pFile, ".")(0)
            If isVBComponentsExist(sModuleName) Then
                Debug.Print "Load VBA from " & sPath & pFile
                If Left(sModuleName, 3) = "RT_" Then
                    Debug.Print "Found report type: " & Mid(sModuleName, 4)
                    REPORT_TYPE = REPORT_TYPE & Mid(sModuleName, 4) & ","
                End If
                ' Require: reg ADD HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Excel\Security /v AccessVBOM /t REG_DWORD /d 1 /f
                With ThisWorkbook.VBProject.VBComponents(sModuleName).CodeModule
                    .DeleteLines 1, .CountOfLines
                    .AddFromString Mid(sData, pos)
                End With
            End If
        Else
            Debug.Print "/!\ Unable to load VBA from " & sPath & pFile & " WITH the reason >Option Explicit< not found in the source file"
        End If
        ' Next file
        pFile = Dir
    Loop
    Debug.Print "Load complete"
    Call IOFile.removeFile(ThisWorkbook.FullName & ".lock")
End Sub


Private Function isVBComponentsExist(sModule As String) As Boolean
    On Error GoTo err_isVBComponentsExist
    Dim tmp: tmp = ThisWorkbook.VBProject.VBComponents(sModule).CodeModule
    Err.Clear
    isVBComponentsExist = True
    Exit Function
err_isVBComponentsExist:
    Err.Clear
    isVBComponentsExist = False
End Function



Public Sub loadModule(ByVal sModuleName As String)
    If Not Common.isDevMode() Then Exit Sub
    Dim sPath As String: sPath = IOFile.getPowerAuditorPath() & "\template\"
    Dim pFile: pFile = Dir(sPath & "\RT_" & sModuleName & ".*")
    Dim pos
    If pFile = "" Then Exit Sub
    Debug.Print "Found report type: " & sModuleName
    Dim sData As String: sData = IOFile.fileGetContent(sPath & "\" & pFile)
    pos = InStr(sData, "Option Explicit")
    With ThisWorkbook.VBProject.VBComponents("RT_" & sModuleName).CodeModule
        .DeleteLines 1, .CountOfLines
        .AddFromString Mid(sData, pos)
    End With
    Debug.Print "Successfully loaded report type: " & sModuleName
End Sub

