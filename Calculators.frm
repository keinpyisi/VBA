VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "Calculators"
   ClientHeight    =   5400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4320
   OleObjectBlob   =   "Calculators.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Dim num1 As Integer
    Dim num2 As Integer
    Dim ans As Integer
    Dim oper As String
    





Private Sub Equal_Click()

   Select Case oper:
        Case "+":
            ans = num1 + num2
            
        Case "-":
            ans = num1 - num2
            
        Case "*":
            ans = num1 * num2
            
        Case "/":
            ans = num1 / num2
            
    End Select
    
    Result = ans
    Result.Visible = True
  
End Sub




Private Sub Number1_Change()
    
    num1 = Val(Number1) 'Get First String From Textbox and change to Int'
    Label3 = Number1
    
End Sub
    
Private Sub Number2_Change()
    
     num2 = Val(Number2) 'Get Second String From Textbox and change to Int'
     Label2 = Number2
    
End Sub

Private Sub Plus_Click()
    
    oper = "+" 'Assign Operator'

End Sub

    
Private Sub Minus_Click()

    oper = "-" 'Assign Operator'

End Sub

Private Sub Multiply_Click()

    oper = "*" 'Assign Operator'
    
End Sub

Private Sub Divide_Click()

    oper = "/" 'Assign Operator'

End Sub




Private Sub Clsar_Click()

    oper = ""
    num1 = 0
    num2 = 0
    ans = 0
    Label3 = "Number 2"
    Label2 = "Number 1"
    Number1 = ""
    Number2 = ""
    Result = ""
    Result.Visible = False
     
End Sub

Private Sub Fin_Click()

Unload Me
End Sub







Private Sub UserForm_Initialize()

    Equal.BackColor = RGB(0, 255, 0)
    Result.Visible = False

End Sub
