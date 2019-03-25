Attribute VB_Name = "Word"
Option Explicit
' PowerAuditor - A simple script to help report writing by https://github.com/1mm0rt41PC
'
' Filename: Word.bas.vb
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

' List of functions:
' + Public Function getInstance() as ActiveDocument
' - Private Function getActiveInstance(docName As String) As ActiveDocument
' + Public Sub resizeUpTable(wDoc As Object, title As String, nbVuln As Integer)
' + Public Sub insertVuln(wDoc As Object, ws As Worksheet, iRow As Integer)
' - Private Sub createVulnId(wDoc As Object, ws As Worksheet, id As Integer)
' - Private Sub insertOrUpdateProof(wDoc As Object, ccExploit As Object, ByVal sFullPath As String, ByVal sLegend As String)
' + Public Sub setCCVal(wDoc As Object, sTitle As String, sVal As String)
' + Public Function getCCVal(wDoc As Object, sTitle As String) As String
' + Public Sub copyExcelColor2Word(wDoc As Object, sTitle As String, rCell As Range)
' + Public Sub updateColorBoldSynthesis(wDoc As Object, aText As Variant, Optional vColor As Variant = False)
' + Public Sub setCCVisibility(wDoc As Object, sTitle As String, bIsHidden As Boolean)
' + Public Sub removeHiddenText(wAd As Object)
' + Public Sub removeAllContentControls(wAd)

Public Const wdFindContinue As Long = 1
Public Const wdReplaceAll As Long = 2
Public Const wdStyleNormal As Long = -1
Public Const wdAlignParagraphCenter As Long = 1
Public Const wdCaptionPositionBelow As Long = 1
Public Const wdTextureNone As Long = 0
Public Const wdStyleHtmlPre As Long = -102
Public Const wdColorAutomatic = -16777216
Public Const wdColorBlack = 0
Public Const wdContentControlRichText = 0
Public Const wdFormatFilteredHTML = 10
Public Const wdContentControlHidden = 2
'Public wDoc As Object ' Instance Word


'===============================================================================
' @brief Get an instance of the word TEMPLATE. If no viable instance exists, this
'        function will create a new one.
' @return {Word.ActiveDocument} The instace to a valid TEMPLATE for the report
'
Public Function getInstance()
    Dim sPath As String: sPath = Environ("USERPROFILE") & "\PowerAuditor\template\"
    Dim wDoc As Object: Set wDoc = getActiveInstance(Application.ActiveWorkbook.Path & "\" & Dir(Application.ActiveWorkbook.Path & "\*TEMPLATE*.docx"))
    If wDoc Is Nothing Then
        If Dir(Application.ActiveWorkbook.Path & "\*TEMPLATE*.docx") = "" Then
            Debug.Print "[*] Using a new template from forge"
            Dim fso As Object: Set fso = VBA.CreateObject("Scripting.FileSystemObject")
            Call fso.CopyFile( _
                sPath & "\" & getInfo("REPORT_TYPE") & "-" & getInfo("LANG") & ".docx", _
                Application.ActiveWorkbook.Path & "\" & RT.getReportFilename("TEMPLATE") & ".docx", _
                True _
            )
        End If
        Set wDoc = CreateObject("word.Application")
        wDoc.Visible = True
        wDoc.Documents.Open Application.ActiveWorkbook.Path & "\" & Dir(Application.ActiveWorkbook.Path & "\*TEMPLATE*.docx")
    End If
    Set wDoc = wDoc.ActiveDocument
    Set getInstance = wDoc
End Function


'===============================================================================
' @brief This function searches for an active Word instance for the requested document {docName}
' @param[in] {String} docName:  The fullpath to the document
' @return {Nothing} if no instance has been found OR {Word.ActiveDocument} if a valid instance has been found
'
Private Function getActiveInstance(docName As String) As Object
    On Error GoTo MyErr
    Dim wApp
    While 1
        Set wApp = GetObject(, "Word.Application")
        If wApp.Documents.Count = 0 Or wApp.Visible = False Then
            wApp.Application.DisplayAlerts = False
            wApp.Quit
        ElseIf wApp.ActiveDocument.FullName = docName Then
            Set getActiveInstance = wApp
            Exit Function
        Else
            wApp.Application.DisplayAlerts = False
            wApp.Quit
        End If
    Wend
MyErr:
    ' Fail
    Set getActiveInstance = Nothing
    Exit Function
End Function


'===============================================================================
' @brief This procedure readjusts the number of rows in an array in order to make
'        it the same number of rows as the number of vulnerabilities.
' @param[in,out] {ActiveDocument} wDoc:  Handle to word instance
' @param[in] {String} title:             Title of the ContentControl that contain the table to alter
' @param[in] {Integer} nbVuln:           Number of vulnerabilities
' @return [NONE]
'
' @note In order for this procedure to work, it is required that the table to be modified is the only one in a ContentControl
'
Public Sub resizeUpTable(wDoc As Object, title As String, nbVuln As Integer)
    Dim myTab: Set myTab = wDoc.SelectContentControlsByTitle(title).Item(1).Range.Tables.Item(1)
    If nbVuln < 1 Then
        Debug.Print "[!][resizeUpTable] Not enough vulnerabilities !"
        Exit Sub
    End If
    While myTab.Rows.Count > nbVuln
        myTab.Rows.Item(3).Delete
    Wend
    While myTab.Rows.Count <= nbVuln
        myTab.Rows.Add myTab.Rows.Item(2)
    Wend
End Sub


'===============================================================================
' @brief Create or update the detail part {wDoc} of a vulnerability {ws[iRow]}
' @param[in,out] {ActiveDocument} wDoc:  Handle to word instance
' @param[in] {Worksheet} ws:             Handle to the Worksheet taht contain the list of all vulnerabilities
' @param[in] {Integer} iRow:             The row in the {ws} who contain a vulnerability
' @return {Nothing} if no instance has been found OR {Word.ActiveDocument} if a valid instance has been found
'
Public Sub insertVuln(wDoc As Object, ws As Worksheet, iRow As Integer)
    ' check si fichier existe return CC.RANGE
    ' Si exist pas insert puis return CC.RANGE
    Dim name As String: name = ws.Cells(iRow, Common.getColLocation(ws, "name")).Value2
    Debug.Print "Inserting the vulnerability: " & name
    Dim iCol As Integer: iCol = 1
    Dim id As Integer: id = iRow - 2
    Dim cellColor As Long
    Call createVulnId(wDoc, ws, id)
    Dim naturalTableColor1: naturalTableColor1 = ws.Cells(2, 1).DisplayFormat.Interior.color
    Dim naturalTableColor2: naturalTableColor2 = ws.Cells(3, 1).DisplayFormat.Interior.color
    Dim cc As Object
    Dim i As Integer
    While ws.Cells(2, iCol).Value2 <> ""
        cellColor = ws.Cells(iRow, iCol).DisplayFormat.Interior.color
        Set cc = wDoc.SelectContentControlsByTitle("VLN_" & ws.Cells(2, iCol).Value2 & "_" & id)
        For i = 1 To cc.Count
            With cc.Item(i).Range
                ' Copy color
                If cellColor <> ThisWorkbook.G_naturalTableColor1 And cellColor <> ThisWorkbook.G_naturalTableColor2 Then             ' Bleu du tableau
                    .Cells.Item(1).Shading.BackgroundPatternColor = cellColor
                End If
                .text = Common.CleaupScoreMesg(ws.Cells(iRow, iCol).Value2)
            End With
        Next
        iCol = iCol + 1
    Wend
    
    
    ' On insert les preuves qui proviennent du dossier VULNDB
    Dim toImportHTML As Variant: toImportHTML = Array("descDetails", "fixDetails", "fixDetails")
    Set cc = wDoc.SelectContentControlsByTitle("VLN_exploit_" & id)(1)
    Dim subCC As Object
    Dim sPath As String: sPath = Common.VulnDBPath(name)
    If IOFile.isFile(sPath & "\desc.html") Then
        For i = 0 To UBound(toImportHTML)
            If IOFile.isFile(sPath & "\" & toImportHTML(i) & ".html") Then
                Set subCC = wDoc.SelectContentControlsByTitle("VLN_" & toImportHTML(i) & "_" & id)(1)
                If Common.isEmptyString(Common.trim(subCC.Range.text, "x")) Then
                    Call subCC.Range.InsertFile(sPath & "\" & toImportHTML(i) & ".html", , , False, False)
                End If
            End If
        Next i
    End If
    
    ' On insert les preuves qui proviennent du dossier VULN
    Dim vlnDir As String: vlnDir = ActiveWorkbook.Path & "\vuln\" & name
    Dim pFile: pFile = Dir(vlnDir & "\*")
    Do While pFile <> ""
        Debug.Print "Inserting the proof: " & pFile
        Call insertOrUpdateProof(wDoc, cc, vlnDir & "\" & pFile, pFile)
        pFile = Dir
    Loop
    ' On insert les preuves qui proviennent des sous dossiers VULN\*
    Dim oFSO As Object: Set oFSO = CreateObject("Scripting.FileSystemObject")
    Dim oFilesList As Object
    Dim oFile As Object
    Dim sFullpath As String
    pFile = Dir(vlnDir & "\*", vbDirectory)
    Do While pFile <> ""
        sFullpath = vlnDir & "\" & pFile
        If pFile <> "." And pFile <> ".." And IOFile.isFolder(sFullpath) Then
            Set oFilesList = oFSO.GetFolder(sFullpath).Files
            
            Set subCC = Common.selectCCInCC(cc.Range, Replace(sFullpath, wDoc.Path, ""))
            If subCC Is Nothing Then
                cc.Range.InsertParagraphAfter
                cc.Range.InsertParagraphAfter
                Set subCC = wDoc.Range(cc.Range.End - 1, cc.Range.End - 1).ContentControls.Add(wdContentControlRichText)
                subCC.title = Replace(sFullpath, wDoc.Path, "")
                subCC.Appearance = wdContentControlHidden
                subCC.Range.text = pFile
                subCC.Range.Paragraphs.Alignment = 0
                subCC.Range.Paragraphs.Style = wDoc.Styles("Titre 3")
            End If
    
            For Each oFile In oFilesList
                Debug.Print "Inserting the proof: " & pFile & "\" & oFile.name
                Call insertOrUpdateProof(wDoc, subCC, sFullpath & "\" & oFile.name, oFile.name)
            Next
        End If
        pFile = Dir
    Loop
End Sub


'===============================================================================
' @brief Create the detail part {wDoc} of a vulnerability {ws[iRow]}
' @param[in,out] {ActiveDocument} wDoc:  Handle to word instance
' @param[in] {Worksheet} ws:             Handle to the Worksheet taht contain the list of all vulnerabilities
' @param[in] {Integer} id:               The row ID (id=iRow-2 (2 = the header (line 1) and for the script line (line 2)))
'                                        in the {ws} who contain a vulnerability
' @return [NONE]
'
Private Sub createVulnId(wDoc As Object, ws As Worksheet, id As Integer)
    Dim VLN_Template As Object
    Dim ccs As Object
    Dim cc_copy As Object
    Dim idName As String
    idName = ws.Cells(2, ColumnIndex:=1).Value2
    idName = idName & "_" & id
    
    If wDoc.SelectContentControlsByTitle("VLN_" & idName).Count <> 0 Then Exit Sub
   
    ' ID n'existe pas
    ' Copy le template
    Set VLN_Template = wDoc.SelectContentControlsByTitle("VLN_Template").Item(1)
    VLN_Template.Range.Copy
    
    ' Paste le new template
    Dim location: location = VLN_Template.Range.Start
    wDoc.Range(location - 1, location - 1).Paste
    
    Set cc_copy = wDoc.SelectContentControlsByTitle("VLN_Template")
    If cc_copy.Item(1).id = VLN_Template.id Then
        Set cc_copy = cc_copy.Item(2)
    Else
        Set cc_copy = cc_copy.Item(1)
    End If
    Set ccs = cc_copy.Range.ContentControls
    
    ' Renomage
    Dim i As Integer
    For i = 1 To ccs.Count
        ccs.Item(i).title = ccs.Item(i).title & "_" & id
    Next
    cc_copy.Delete False
End Sub


'===============================================================================
' @brief Insert a proof in the detail part of a vulnerability
' @param[in,out] {ActiveDocument} wDoc:      Handle to word instance
' @param[in,out] {ContentControl} ccExploit: Handle to the ContentControl "VLN_exploit_xxx"
' @param[in] {String} sFullPath:             Fullpath to the proof (image or text)
' @param[in] {String} sLegend:               Texts of the legend
' @return [NONE]
'
Private Sub insertOrUpdateProof(wDoc As Object, ccExploit As Object, ByVal sFullpath As String, ByVal sLegend As String)
    Dim waitOnReturn As Boolean: waitOnReturn = True
    Dim windowStyle As Integer: windowStyle = 0
    
    ' Création du ContentControl
    Dim cc: Set cc = Common.selectCCInCC(ccExploit.Range, Replace(sFullpath, wDoc.Path, ""))
    If cc Is Nothing Then
        ccExploit.Range.InsertParagraphAfter
        ccExploit.Range.InsertParagraphAfter
        Set cc = wDoc.Range(ccExploit.Range.End - 1, ccExploit.Range.End - 1).ContentControls.Add(wdContentControlRichText)
        cc.title = Replace(sFullpath, wDoc.Path, "")
        cc.Appearance = wdContentControlHidden
    Else
        Exit Sub ' On ne modifit pas les traces déjà en place
    End If
    ' Reset du CC
    cc.Range.text = vbNewLine
    cc.Range.Paragraphs.Alignment = 0
    cc.Range.Paragraphs.Style = wdStyleNormal
    
    Dim ext As String: ext = IOFile.getFileExt(sFullpath)
    If ext = "png" Or ext = "jpg" Or ext = "jpeg" Or ext = "bmp" Then
        Dim wrdPic: Set wrdPic = cc.Range.InlineShapes.AddPicture(fileName:=sFullpath, LinkToFile:=False, SaveWithDocument:=True)
        'wrdPic.ScaleHeight = 50
        'wrdPic.ScaleWidth = 50
        wrdPic.Range.Paragraphs.Alignment = wdAlignParagraphCenter
        wrdPic.Range.InsertCaption Label:="Figure", title:=" - " & Replace(sLegend, ".png", ""), Position:=wdCaptionPositionBelow
    Else
        Dim tmpFile As String: tmpFile = Environ("temp") & "\" & RandomString(7) & ".html"
        Dim pygmentize As String: pygmentize = "pygmentize"
        If IOFile.isFile(Common.PowerAuditorPath() & "\bin\pygmentize.exe") Then
            pygmentize = Chr(34) & Common.PowerAuditorPath() & "\bin\pygmentize.exe" & Chr(34)
        End If
        Debug.Print "Using pygmentize from: " & pygmentize
        
        Dim wsh As Object: Set wsh = VBA.CreateObject("WScript.Shell")
        wsh.Run pygmentize & " -f html -l " & IOFile.getFileExt(sFullpath) & " -O full,noclasses,style=monokai -o " & Chr(34) & tmpFile & Chr(34) & " " & Chr(34) & sFullpath & Chr(34), windowStyle, waitOnReturn
        If Not IOFile.isFile(tmpFile) Then ' Si pygmentize n'est pas trouvé ou en cas d'erreur => utilisation du fichier original
            tmpFile = sFullpath
            Debug.Print "pygmentize Failed !"
        End If
    
        ' Insertion du fichier
        Call cc.Range.InsertFile(fileName:=tmpFile, Link:=False, Attachment:=False)
        cc.Range.NoProofing = True
        If tmpFile = sFullpath Then
            With cc.Range
                .Style = wdStyleHtmlPre
                .Paragraphs.Style = wdStyleHtmlPre
                .Paragraphs.Shading.Texture = wdTextureNone
                .Paragraphs.Shading.ForegroundPatternColor = wdColorAutomatic
                .Paragraphs.Shading.BackgroundPatternColor = wdColorBlack
            End With
        End If
    
        ' Ajout de la legende
        Call Common.trimContentControl(cc)
        'If Not isEmptyString(wDoc.Range(cc.Range.End - 1, cc.Range.End).Text) Then
        '    cc.Range.InsertParagraphAfter
        'End If
        With wDoc.Range(cc.Range.End - 1, cc.Range.End)
            .InsertCaption Label:="Figure", title:=" - " & Replace(sLegend, "." & IOFile.getFileExt(sFullpath), ""), Position:=wdCaptionPositionBelow
        End With
        With wDoc.Range(cc.Range.End - 1, cc.Range.End)
            .Paragraphs.Alignment = wdAlignParagraphCenter
        End With
    End If
    Call trimContentControl(cc)
End Sub


'===============================================================================
' @brief Set the content text {sVal} for >ALL< ContentControls who have the title {sTitle} in the document {wDoc}
' @param[in,out] {ActiveDocument} wDoc: Handle to word instance
' @param[in] {String} sTitle:           Name of the ContentControl
' @param[in] {String} sVal:             Text to put in the ContentControl
' @return [NONE]
'
Public Sub setCCVal(wDoc As Object, sTitle As String, sVal As String)
    Dim i As Integer
    Dim ccs: Set ccs = wDoc.SelectContentControlsByTitle(sTitle)
    For i = 1 To ccs.Count
        ccs.Item(i).Range.text = sVal
    Next
End Sub


'===============================================================================
' @brief Get the text of the >FIRST< ContentControl who have the title {sTitle} in the document {wDoc}
' @param[in,out] {ActiveDocument} wDoc: Handle to word instance
' @param[in] {String} sTitle:           Name of the ContentControl
' @return {String} Text of the ContentControl
'
' @note /!\ Please check that the ContentControl exist before calling this function
'
Public Function getCCVal(wDoc As Object, sTitle As String) As String
    With wDoc.SelectContentControlsByTitle(sTitle)
        getCCVal = .Item(1).Range.text
    End With
End Function


'===============================================================================
' @brief Replicates the color of the Excel cell {rCell} to >ALL< Word cells (Tables)
'        that contain a ContentControl with the name {sTitle}
' @param[in,out] {ActiveDocument} wDoc: Handle to word instance
' @param[in] {String} sTitle:       Name of the ContentControl
' @param[in] {Range} rCell:         The Excel cell to copy
' @return [NONE]
'
Public Sub copyExcelColor2Word(wDoc As Object, sTitle As String, rCell As Range)
    Dim lColor As Long: lColor = rCell.DisplayFormat.Interior.color
    Dim ccs: Set ccs = wDoc.SelectContentControlsByTitle(sTitle)
    Dim sVal As String: sVal = Common.CleaupScoreMesg(rCell.Value2)
    Dim i As Integer
    For i = 1 To ccs.Count
        If ccs(1).Range.Cells.Count = 1 Then
            ccs(1).Range.Cells.Item(1).Shading.BackgroundPatternColor = lColor
        End If
        ccs.Item(i).Range.text = sVal
    Next
End Sub


'===============================================================================
' @brief Replaces the table texts with the associated colors
' @param[in,out] {ActiveDocument} wDoc:  Handle to word instance
' @param[in] {Array} aText:              Text to replace with color
' @param[in,opt] {Boolean/Array} vColor: If True, replace all text from {aText} by a color hardcoded in this sub.
'                                        If Array, replace all text from {aText} by a color from this variable.
' @return [NONE]
'
Public Sub updateColorBoldSynthesis(wDoc As Object, aText As Variant, Optional vColor As Variant = False)
    Dim i As Integer, j As Integer
    Dim color As Variant
    'color = Array("GRIS", "JAUNE", "ORANGE", "ROUGE")
    If vColor = True And VarType(vColor) = vbBoolean Then
        vColor = Array(RGB(133, 133, 133), RGB(255, 255, 0), RGB(255, 192, 0), RGB(255, 0, 0))
    End If
    For i = 0 To UBound(aText)
        With wDoc.Content.Find
            .ClearFormatting
            With .Replacement
                .ClearFormatting
                .Font.Bold = True
                If VarType(vColor) = vbArray Then
                    .Font.color = vColor(Left(aText(i), 1) - 1)
                End If
                .Font.Shadow = True
            End With
            If VarType(vColor) = vbArray Then
                .Execute FindText:="{" & aText(i) & "}", replacewith:=Split(aText(i), " - ")(1), Format:=True, Replace:=wdReplaceAll
            Else
                .Execute FindText:="**" & aText(i) & "**", replacewith:=aText(i), Format:=True, Replace:=wdReplaceAll
            End If
        End With
    Next
End Sub


'===============================================================================
' @brief Set the visibility of >ALL< ContentControls with the title {sTitle}
' @param[in,out] {ActiveDocument} wDoc:  Handle to word instance
' @param[in] {String} sTitle:            Name of the ContentControl
' @param[in] {Boolean} bIsHidden:        Hide the ContentControl ?
' @return [NONE]
'
Public Sub setCCVisibility(wDoc As Object, sTitle As String, bIsHidden As Boolean)
    Dim i As Integer
    Dim ccs As Object: Set ccs = wDoc.SelectContentControlsByTitle(sTitle)
    wDoc.ActiveWindow.View.ShowHiddenText = True
    For i = 1 To ccs.Count
        ccs.Item(i).Range.Font.Hidden = bIsHidden
    Next
    wDoc.ActiveWindow.View.ShowHiddenText = False
End Sub


'===============================================================================
' @brief Remove >ALL< hidden text in the ActiveDocument
' @param[in,out] {ActiveDocument} wAd:  Handle to word instance
' @return [NONE]
'
Public Sub removeHiddenText(wAd As Object)
    Dim hasBeenFound As Boolean: hasBeenFound = True
    Dim wdRange
    wAd.ActiveWindow.View.ShowHiddenText = True
    Dim maxLoop As Integer: maxLoop = 3
    While hasBeenFound
        hasBeenFound = False
        Set wdRange = wAd.Content
        With wdRange.Find
            .ClearFormatting
            .Format = True
            .Font.Hidden = True
            .Wrap = wdFindContinue ' wdFindContinue=1
            If maxLoop = 0 Then
                .Replacement.ClearFormatting
            End If
            .Execute replacewith:="", Replace:=wdReplaceAll ' = 2
            If .found Then
                Dim i As Integer: i = 1
                While i <= wdRange.Tables.Count
                    If wdRange.Tables(i).Range.Font.Hidden = -1 Or wdRange.Tables(i).Range.Cells(1).Range.Font.Hidden = -1 Then
                        wdRange.Tables(i).Delete
                    Else
                        i = i + 1
                    End If
                Wend
                i = 1
                While i <= wdRange.ContentControls.Count
                    If wdRange.ContentControls(i).Range.Font.Hidden = -1 Then
                        wdRange.ContentControls(i).Delete
                    Else
                        i = i + 1
                    End If
                Wend
                i = 1
                hasBeenFound = True
            End If
        End With
        maxLoop = maxLoop - 1
        If maxLoop < 0 Then Exit Sub
    Wend
    wAd.ActiveWindow.View.ShowHiddenText = False
End Sub


'===============================================================================
' @brief Remove >ALL< ContentControls in the ActiveDocument
' @param[in,out] {ActiveDocument} wAd:  Handle to word instance
' @return [NONE]
'
Public Sub removeAllContentControls(wAd)
    Dim ccs As Object: Set ccs = wAd.ContentControls
    While ccs.Count > 0
        ccs.Item(1).Delete False
    Wend
End Sub

