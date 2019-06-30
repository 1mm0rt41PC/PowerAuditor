Attribute VB_Name = "RT"
Option Explicit
' PowerAuditor - A simple script to help report writing by https://github.com/1mm0rt41PC
'
' Filename: RT.bas.vb - This module bypass bug of vba dynamic function call.
' From 2019-03-22 the version of O365, when you call via "Application.Run"
' a function in "Microsoft Excel Object" part, the function will always return empty "".
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
Public m_return
Private m_RT As String ' ModuleName

Private Function getRT() As String
    If m_RT = "" Then m_RT = "RT_" & Split(getInfo("REPORT_TYPE"), "-")(0)
    getRT = m_RT
End Function

Public Function getCorp() As String
    Call Application.Run(getRT() & ".getCorp")
    getCorp = m_return
End Function


Public Function getManager() As String
    Call Application.Run(getRT() & ".getManager")
    getManager = m_return
End Function


Public Function getReportFilename(pType As String) As String
    Call Application.Run(getRT() & ".getReportFilename", pType)
    getReportFilename = m_return
End Function


Public Function getExportFields_HTML() As Variant
    Call Application.Run(getRT() & ".getExportFields_HTML")
    getExportFields_HTML = m_return
End Function


Public Function getExportFields_TXT() As Variant
    Call Application.Run(getRT() & ".getExportFields_TXT")
    getExportFields_TXT = m_return
End Function


Public Function getExportField_KeyColumn(ws As Worksheet) As Integer
    Call Application.Run(getRT() & ".getExportField_KeyColumn", ws)
    getExportField_KeyColumn = m_return
End Function


Public Sub exportExcel2Word_before(wDoc As Object, ws As Worksheet)
    Call Application.Run(getRT() & ".exportExcel2Word_before", wDoc, ws)
End Sub


Public Sub exportExcel2Word_insertVuln(wDoc As Object, ws As Worksheet, iRow As Integer)
    Call Application.Run(getRT() & ".exportExcel2Word_insertVuln", wDoc, ws, iRow)
End Sub


Public Sub exportExcel2Word_after(wDoc As Object, ws As Worksheet, nbVuln As Integer)
    Call Application.Run(getRT() & ".exportExcel2Word_after", wDoc, ws, nbVuln)
End Sub


Public Sub genSynthesis(wDoc As Object, ws As Worksheet)
    Call Application.Run(getRT() & ".GenSynthesis", wDoc, ws)
End Sub


Public Sub exportFinalStaticsDocuments(wDoc As Object, ws As Worksheet)
    Call Application.Run(getRT() & ".ExportFinalStaticsDocuments", wDoc, ws)
End Sub


Public Function getExcelFilename() As String
    Call Application.Run(getRT() & ".getExcelFilename")
    getExcelFilename = m_return
End Function
