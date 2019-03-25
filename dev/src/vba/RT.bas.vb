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
Public m_return As String

Public Function getCorp() As String
    Call Application.Run("RT_" & getInfo("REPORT_TYPE") & ".getCorp")
    getCorp = m_return
End Function


Public Function getManager() As String
    Call Application.Run("RT_" & getInfo("REPORT_TYPE") & ".getManager")
    getManager = m_return
End Function


Public Function getReportFilename(pType As String) As String
    Call Application.Run("RT_" & getInfo("REPORT_TYPE") & ".getReportFilename", pType)
    getReportFilename = m_return
End Function
