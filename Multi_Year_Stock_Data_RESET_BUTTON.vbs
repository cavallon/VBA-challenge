Sub Reset_Workbook()
'
' Reset_Workbook Macro
' Resets all calculations that were completed.
'

'
    Columns("I:Q").Select
    Selection.ClearContents
    Columns("J:J").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Sheets("2019").Select
    Columns("I:Q").Select
    Selection.ClearContents
    Columns("J:J").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Sheets("2020").Select
    Columns("I:Q").Select
    Selection.ClearContents
    Columns("J:J").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    Sheets("2018").Select
    
End Sub