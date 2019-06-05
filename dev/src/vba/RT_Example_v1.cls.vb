VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RT_Example_v1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
' PowerAuditor - A simple script to help report writing by https://github.com/1mm0rt41PC
'
' Filename: RT_Example_v1.cls.vb
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


Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    If Not Application.Intersect(Target, Range("F3:G100")) Is Nothing Then
        Dim tmp: tmp = Common.CVSSReader(Range("G" & Target.Row).Value2)
        tmp = Split(tmp, vbCrLf)
        Range("G" & Target.Row).Value2 = tmp(0)
        Range("F" & Target.Row).Value2 = tmp(1)
        Cancel = True
    End If
End Sub


Public Function getCorp() As String
    RT.m_return = "Example Corp"
    getCorp = RT.m_return ' Bug dynamic call
End Function


Public Function getManager() As String
    RT.m_return = "God him self"
    getManager = RT.m_return ' Bug dynamic call
End Function


Public Function getReportFilename(pType As String) As String
    Dim auditType As String
    If Common.getLang() = "FR" Then
        auditType = "TI"
    Else
        auditType = "PT"
    End If
    RT.m_return = "myCorp-" & getInfo("CLIENT") & "-" & auditType & "-" & getInfo("TARGET") & "-" & pType & "-" & Mid(Replace(Xls.getVersionDate(), "-", ""), 3) & "."
    getReportFilename = RT.m_return
End Function


Public Function getExcelFilename() As String
    If Common.getLang() = "FR" Then
        RT.m_return = getReportFilename("RS")
    Else
        RT.m_return = getReportFilename("SR")
    End If
    getExcelFilename = RT.m_return
End Function


Public Function getExportFields_HTML() As Variant
    RT.m_return = Array("descDetails", "fixDetails", "fixDetails") ' Bug dynamic call
    getExportFields_HTML = RT.m_return ' Bug dynamic call
End Function


Public Function getExportFields_TXT() As Variant
    RT.m_return = Array("category", "desc", "fix", "risk", "fixtype") ' Bug dynamic call
    getExportFields_TXT = RT.m_return ' Bug dynamic call
End Function


Public Function getExportField_KeyColumn(ws As Worksheet) As Integer
    RT.m_return = Xls.getColLocation(ws, "name") ' Bug dynamic call
    getExportField_KeyColumn = RT.m_return ' Bug dynamic call
End Function


Public Sub initWordExport(wDoc As Object, ws As Worksheet)
    ' Do what you want here BEFORE export Excel to Word
End Sub


Public Sub finalizeWordExport(wDoc As Object, ws As Worksheet, nbVuln As Integer)
	' Do what you want here AFTER export Excel to Word
End Sub


Public Sub insertVuln(wDoc As Object, ws As Worksheet, line As Integer)
	' This function allow you to insert vulnerabilities
	' If you want, you can use the default function:
    Call Word.insertVuln(wDoc, ws, line)
End Sub


Public Sub genSynthesis(wDoc As Object, ws As Worksheet)
	' Do stuff to make a synthesis in the word document
End Sub


Public Sub exportFinalStaticsDocuments(wDoc As Object, ws As Worksheet)
	' Do stuff to export the Word
    ThisWorkbook.Save
    IOFile.renameDocument wDoc, "docx", "TEMPLATE", deleteOld:=True
    ThisWorkbook.Save
	
	Debug.Print "Generate all documents"
	Dim myDocx: myDocx = IOFile.renameDocument(wDoc, "docx", "DETAIL", deleteOld:=False)
	Dim wd_Exp as Object
	Set wd_Exp = wDoc.Application.Documents.Open(myDocx)
	
    Call Word.removeHiddenText(wd_Exp)
    wd_Exp.Fields.Update
    Call Word.removeAllContentControls(wd_Exp)
    wd_Exp.Fields.Update
    wd_Exp.Save
    wd_Exp.Close
	set wd_Exp = Nothing
	
	
    Debug.Print "Generate XLSX"
    If Common.getLang() = "EN" Then
        Call Xls.exportPowerauditorToXlsx(Array("Vulnerabilities"))
    Else
        Call Xls.exportPowerauditorToXlsx(Array("Vulnérabilités"))
    End If
    Application.DisplayAlerts = True
End Sub