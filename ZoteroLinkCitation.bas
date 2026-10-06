Attribute VB_Name = "ZoteroLinkCitation"
Public Sub ZoteroLinkCitation()
Dim nStart&, nEnd&
nStart = Selection.Start
nEnd = Selection.End
Application.ScreenUpdating = False
Dim title As String
Dim titleAnchor As String
Dim style As String
Dim fieldCode As String
Dim numOrYear As String
Dim pos&, n1&, n2&
 
ActiveWindow.View.ShowFieldCodes = True
Selection.Find.ClearFormatting
With Selection.Find
    .Text = "^d ADDIN ZOTERO_BIBL"
    .Replacement.Text = ""
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = False
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute
With ActiveDocument.Bookmarks
    .Add Range:=Selection.Range, Name:="Zotero_Bibliography"
    .DefaultSorting = wdSortByName
    .ShowHidden = True
End With
ActiveWindow.View.ShowFieldCodes = False
 
 
For Each aField In ActiveDocument.Fields
' check if the field is a Zotero in-text reference
    If InStr(aField.Code, "ADDIN ZOTERO_ITEM") > 0 Then
        fieldCode = aField.Code
        pos = 0
        Paper_i = 1
        Do While InStr(fieldCode, """title"":""") > 0
            n1 = InStr(fieldCode, """title"":""") + Len("""title"":""")
            n2 = InStr(Mid(fieldCode, n1, Len(fieldCode) - n1), """,""") - 1 + n1
        
            title = Mid(fieldCode, n1, n2 - n1)
            
            titleAnchor = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(title, " ", "_"), "#", "_"), "&", "_"), ":", "_"), ",", "_"), "-", "_"), "‐", "_"), "'", "_"), ".", "_"), "(", "_"), ")", "_"), "?", "_"), "!", "_")
            titleAnchor = Left(titleAnchor, 40)
            
            Selection.GoTo What:=wdGoToBookmark, Name:="Zotero_Bibliography"
            Selection.Find.ClearFormatting
            With Selection.Find
                .Text = Left(title, 255)
                .Replacement.Text = ""
                .Forward = True
                .Wrap = wdFindAsk
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            '查找引文，Bibliography
            Selection.Find.Execute
            '选中对应引文的这一段
            Selection.Paragraphs(1).Range.Select
            
            If Selection.Range.Start <> Selection.Range.End Then
                If ActiveDocument.Bookmarks.Exists(titleAnchor) Then
                    ActiveDocument.Bookmarks(titleAnchor).Delete
                End If
                ActiveDocument.Bookmarks.Add Range:=Selection.Range, Name:=titleAnchor
            End If


            
            aField.Select
                        
            Selection.Find.ClearFormatting
                
            If pos = 0 Then
                ' 初始化起始位置和数组
                startPosition = 1
                ReDim commaPositions(1 To 1)
                    
                ' 查找逗号的位置(前提是作者和年份之间采用英文逗号分隔符，否则要改为其他符号)
                Do
                    commaPosition = InStr(startPosition, Selection, ",")
                    
                    If commaPosition > 0 Then
                        ' 将逗号的位置添加到数组
                        commaPositions(UBound(commaPositions)) = commaPosition
                        ' 更新起始位置，以便下一次查找
                        startPosition = commaPosition + 1
                        ReDim Preserve commaPositions(1 To UBound(commaPositions) + 1)
                    End If
                Loop While commaPosition > 0
            End If
                ' 输出记录的逗号位置
            'For j = 1 To UBound(commaPositions)
                'Debug.Print "Comma found at position: " & commaPositions(j)
            'Next j
                
            With Selection.Find
                .Text = "^#"
                .Replacement.Text = ""
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            
            Selection.Find.Execute
            
            Selection.MoveLeft Unit:=wdCharacter, Count:=1
            Selection.MoveRight Unit:=wdCharacter, Count:=pos
            
            Selection.Find.Execute
            Selection.MoveLeft Unit:=wdCharacter, Count:=1
            Selection.MoveRight Unit:=wdWord, Count:=1, Extend:=wdExtend
             
            numOrYear = Selection.Range.Text & ""
            
            pos = commaPositions(Paper_i) - 1
            Paper_i = Paper_i + 1
            
            style = Selection.style
            '如果为文中的参考文献引用设定了格式，那么需要取消下面的注释
            'Selection.style = ActiveDocument.Styles("CitationFormating")
            
            '插入超链接
            ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, Address:="", SubAddress:=titleAnchor, ScreenTip:="", TextToDisplay:="" & numOrYear
                       
            'Selection.style = style
            
            
            
            fieldCode = Mid(fieldCode, n2 + 1, Len(fieldCode) - n2 - 1)
        
        Loop
    End If
Next aField
ActiveDocument.Range(nStart, nEnd).Select
End Sub
