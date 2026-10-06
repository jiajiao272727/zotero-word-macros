Attribute VB_Name = "ZoteroLinkCitation2"
Public Sub ZoteroLinkCitation2()
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
    
    ' 补充声明原代码中缺失的变量
    Dim Paper_i As Long
    Dim startPosition As Long
    Dim commaPositions() As Long
    Dim tmpN2 As Long
    Dim remainingLength As Long
    
    ' -------------------------------
    ' 提前初始化正则表达式对象，避免在循环中反复创建，提高运行速度
    ' -------------------------------
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Pattern = "[^A-Za-z0-9_]"
    regEx.Global = True
    
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

    Dim aField As Field
    For Each aField In ActiveDocument.Fields
        ' 检查是否为 Zotero 引文
        If InStr(aField.Code, "ADDIN ZOTERO_ITEM") > 0 Then
            fieldCode = aField.Code
            pos = 0
            Paper_i = 1

            Do While InStr(fieldCode, """title"":""") > 0
                n1 = InStr(fieldCode, """title"":""") + Len("""title"":""")
                tmpN2 = InStr(Mid(fieldCode, n1, Len(fieldCode) - n1), """,""")
                
                If tmpN2 = 0 Then
                    n2 = Len(fieldCode) ' 如果没找到，就取到最后
                Else
                    n2 = tmpN2 - 1 + n1
                End If
                title = Mid(fieldCode, n1, n2 - n1)

                ' -------------------------------
                ' 安全生成书签名 (修复 5828 错误)
                ' -------------------------------
                titleAnchor = title
                ' 用正则替换所有非字母数字字符为 "_"
                titleAnchor = regEx.Replace(titleAnchor, "_")
                
                ' 【关键修复】：强制加上 "Z_" 前缀，保证书签绝对以字母开头！
                titleAnchor = "Z_" & titleAnchor
                
                ' 截取前 40 个字符，防止书签过长 (由于加了前缀，这里同样截取前 40)
                If Len(titleAnchor) > 40 Then titleAnchor = Left(titleAnchor, 40)
                
                ' -------------------------------
                ' 定位 Bibliography 引文段落
                ' -------------------------------
                Selection.GoTo What:=wdGoToBookmark, Name:="Zotero_Bibliography"
                Selection.Find.ClearFormatting
                With Selection.Find
                    .Text = Replace(Left(title, 255), "+", "") ' 清理 +，避免查找失败
                    .Replacement.Text = ""
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                End With

                If Not Selection.Find.Execute Then
                    ' 如果找不到匹配，跳过当前标题
                    GoTo NextTitle
                End If

                Selection.Paragraphs(1).Range.Select

                ' 删除已存在同名书签
                If ActiveDocument.Bookmarks.Exists(titleAnchor) Then
                    ActiveDocument.Bookmarks(titleAnchor).Delete
                End If
                
                ' 添加新书签
                ActiveDocument.Bookmarks.Add Range:=Selection.Range, Name:=titleAnchor

                ' -------------------------------
                ' 处理引文中的数字或年份
                ' -------------------------------
                aField.Select
                If pos = 0 Then
                    ' 初始化逗号数组
                    startPosition = 1
                    ReDim commaPositions(1 To 1)
                    Dim commaPosition As Long
                    Do
                        commaPosition = InStr(startPosition, Selection.Text, ",")
                        If commaPosition > 0 Then
                            commaPositions(UBound(commaPositions)) = commaPosition
                            startPosition = commaPosition + 1
                            ReDim Preserve commaPositions(1 To UBound(commaPositions) + 1)
                        End If
                    Loop While commaPosition > 0
                End If

                With Selection.Find
                    .Text = "^#"
                    .Replacement.Text = ""
                    .Forward = True
                    .Wrap = wdFindContinue
                End With
                Selection.Find.Execute

                Selection.MoveLeft Unit:=wdCharacter, Count:=1
                Selection.MoveRight Unit:=wdCharacter, Count:=pos

                Selection.Find.Execute
                Selection.MoveLeft Unit:=wdCharacter, Count:=1
                Selection.MoveRight Unit:=wdWord, Count:=1, Extend:=wdExtend

                numOrYear = Selection.Range.Text & ""

                ' 防止数组越界
                If Paper_i <= UBound(commaPositions) Then
                    pos = commaPositions(Paper_i) - 1
                Else
                    pos = 0
                End If
                Paper_i = Paper_i + 1

                ' -------------------------------
                ' 插入超链接
                ' -------------------------------
                ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, Address:="", SubAddress:=titleAnchor, ScreenTip:="", TextToDisplay:=numOrYear

NextTitle:
                ' -------------------------------
                ' 【修复】防止截取字符串时出现负数报错
                ' -------------------------------
                remainingLength = Len(fieldCode) - n2
                If remainingLength > 0 Then
                    fieldCode = Mid(fieldCode, n2 + 1, remainingLength)
                Else
                    Exit Do
                End If
            Loop
        End If
    Next aField

    ActiveDocument.Range(nStart, nEnd).Select
    Application.ScreenUpdating = True
End Sub
