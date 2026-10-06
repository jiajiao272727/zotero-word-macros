Attribute VB_Name = "ColorIEEERefs"
Sub ColorIEEERefs()
    Dim rng As Range
    Set rng = ActiveDocument.Content
    With rng.Find
        .ClearFormatting
        .Text = "\[[0-9]{1,}\]"
        .MatchWildcards = True
        Do While .Execute
            rng.Font.Color = wdColorBlue
            rng.Collapse wdCollapseEnd
        Loop
    End With
End Sub

