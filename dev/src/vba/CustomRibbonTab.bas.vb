Attribute VB_Name = "CustomRibbonTab"
Option Explicit
' PowerAuditor - A simple script to help report writing by https://github.com/1mm0rt41PC
'
' Filename: CustomRibbonTab.bas.vb
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

Private G_Ribbon As Object

Private Sub onLoad(rb As Object)
    Debug.Print "Custon ribbon loaded"
    Set CustomRibbonTab.G_Ribbon = rb
    CustomRibbonTab.G_Ribbon.ActivateTab "PowerAuditor"
End Sub


Private Sub nbReportType(control As Object, ByRef nbElt)
    'Définit le nombre d'éléments dans la combobox
    Dim rng As Range: Set rng = Worksheets("PowerAuditor").Range("REPORT_TYPE_TBL[REPORT TYPE]")
    If rng.Count <= 0 Then
        nbElt = 0
        Exit Sub
    End If
    nbElt = 1
    While rng.Value2(nbElt, 1) <> ""
        nbElt = nbElt + 1
    Wend
    nbElt = nbElt - 1
End Sub


Private Sub getReportTypeByIndex(control As Object, index As Integer, ByRef elt)
    elt = Range("REPORT_TYPE_TBL[REPORT TYPE]").Value2(index + 1, 1)
End Sub


Private Sub RibbonOnChange(control As Object, val)
    Worksheets("PowerAuditor").Range(control.id).Value2 = val
End Sub


Private Sub GetEnabled(control As Object, ByRef enabled)
    'enabled = Common.isDevMode()
    enabled = True
End Sub


Private Sub GetText(control As IRibbonControl, ByRef text)
    text = Worksheets("PowerAuditor").Range(control.id).Value2
End Sub


Public Sub invalidAlltext()
    On Error Resume Next
    CustomRibbonTab.G_Ribbon.InvalidateControl "CLIENT"
    CustomRibbonTab.G_Ribbon.InvalidateControl "TARGET"
    CustomRibbonTab.G_Ribbon.InvalidateControl "VERSION_DATE"
    CustomRibbonTab.G_Ribbon.InvalidateControl "BEGIN_DATE"
    CustomRibbonTab.G_Ribbon.InvalidateControl "END_DATE"
    CustomRibbonTab.G_Ribbon.InvalidateControl "LANG"
    CustomRibbonTab.G_Ribbon.InvalidateControl "REPORT_TYPE"
    Err.Clear
End Sub
