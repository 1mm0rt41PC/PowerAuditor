VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PowerAuditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
' PowerAuditor - A simple script to help report writing by https://github.com/1mm0rt41PC
'
' Filename: PowerAuditor.cls.vb
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
Private Sub Worksheet_Change(ByVal Target As Range)
    If ThisWorkbook.G_exportToProd Then Exit Sub
    If Not Application.Intersect(Range("REPORT_TYPE"), Range(Target.Address)) Is Nothing Then
        Call Xls.loadExcelSheet
    End If
    
    If Not Application.Intersect(Range("LEVEL"), Range(Target.Address)) Is Nothing Then
        Call Xls.updateLevelCellColor
    End If
End Sub

