Attribute VB_Name = "Module1"
Option Explicit

Sub ‚Ü‚é‚Ü‚é•\()

    Dim row As Integer
    Dim col As Integer
    Dim row_cell As Range
    Dim col_cell As Range
    
    
    Cells().Select     'Select All Cell'
    
    Selection.Clear   'Clear Cell'
 
    For row = 1 To 9
    
        'Insert Row Into Row 1,2,3,4'
        Cells(row, 1) = row
        
        For col = 1 To 9
            'Insert Row*Col Into Col 2,4,6,8'
            Cells(row, col) = row * col
            
            Next col
        
        Next row
     Call Bracket("A1:I9")
     Call color("A1:I9", RGB(232, 235, 107))

    
     'Move Cursor to A12'
     Range("A12").Select
     
     For row = 1 To 9
         'Insert Row Into Row 1,2,3,4'
        ActiveCell.Offset(row - 1, 0) = row
        
        For col = 1 To 9
            'Insert Row*Col Into Col 2,4,6,8'
            ActiveCell.Offset(row - 1, col - 1) = row * col
            
            Next col
        
        Next row
    Call Bracket("A12:I20")
    Call color("A12:I20", RGB(112, 222, 108))
   

End Sub


Sub Bracket(rangeStr As String)
Attribute Bracket.VB_ProcData.VB_Invoke_Func = " \n14"
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
Attribute color.VB_ProcData.VB_Invoke_Func = " \n14"
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
