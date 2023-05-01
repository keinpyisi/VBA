Option Explicit

Const STROW = 3
Const STCOL = 2

Sub まるまる表()

    Dim row As Integer
    Dim col As Integer
    Dim row_cell As Range
    Dim col_cell As Range
    Dim rangeStr As String
    
    
    Cells().Select     'Select All Cell'
    
    Selection.Clear   'Clear Cell'
 
    For row = 1 To 9
    
        'Insert Row Into Row 1,2,3,4'
        'Cells(STROW + row, 1) = row'
        
        For col = 1 To 9
            'Insert Row*Col Into Col 2,4,6,8'
            Cells(STROW + row, STCOL + col) = row * col
            
            Next col
        
        Next row
     rangeStr = Chr(Asc("A") + STCOL) & STROW + 1 & ":" & Chr(Asc("I") + STCOL) & STROW + 9
     Call Bracket(rangeStr)
     Call color(rangeStr, RGB(232, 235, 107))

     rangeStr = Chr(Asc("A") + STCOL) & STROW + 11
   
     'Move Cursor to A12'
     Range(rangeStr).Select
     
     For row = 1 To 9
         'Insert Row Into Row 1,2,3,4'
        ActiveCell.Offset(row - 1, 0) = row
        
        For col = 1 To 9
            'Insert Row*Col Into Col 2,4,6,8'
            ActiveCell.Offset(row - 1, col - 1) = row * col
            
            Next col
        
        Next row
    'A12:I20'
    rangeStr = Chr(Asc("A") + STCOL) & STCOL + 12 & ":" & Chr(Asc("I") + STCOL) & STROW + 19
    Call Bracket(rangeStr)
    Call color(rangeStr, RGB(112, 222, 108))
   

End Sub


Sub Bracket(rangeStr As String)
'
' Bracket Macro
'

'
    Range(rangeStr).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub



Sub color(rangeStr As String, color As Long)
'
' Color Macro
'

'
    Range(rangeStr).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .color = color
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub
