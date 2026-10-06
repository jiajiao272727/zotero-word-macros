Attribute VB_Name = "CitingColor"
Sub CitingColor()
    For i = 1 To ActiveDocument.Fields.Count '遍历文档所有域
        If Left(ActiveDocument.Fields(i).Code, 4) = " REF" Then 'Word自带的交叉引用的域代码起始4位是" REF"（注意空格
        ActiveDocument.Fields(i).Select '选中上述几类域
        Selection.Font.Color = wdColorBlue '设置字体颜色为蓝色,也可以使用以下代码设置为任意颜色Selection.Font.Color = 12673797
        End If
    Next
End Sub
