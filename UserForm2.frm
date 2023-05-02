VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "Calculators"
   ClientHeight    =   5400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4320
   OleObjectBlob   =   "UserForm2.frx":0000
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
    Result.Visible = True
     
End Sub

Private Sub Fin_Click()

Unload Me
End Sub







Private Sub ToggleButton1_Change()
    
    Select Case ToggleButton1.Value:
     Case True:
            Equal.BackColor = RGB(0, 0, 0)
            Equal.ForeColor = RGB(255, 255, 255)
            Result.BackColor = RGB(0, 0, 0)
            Result.ForeColor = RGB(255, 255, 255)
            Number1.BackColor = RGB(0, 0, 0)
            Number1.ForeColor = RGB(255, 255, 255)
            Number2.BackColor = RGB(0, 0, 0)
            Number2.ForeColor = RGB(255, 255, 255)
            UserForm2.BackColor = RGB(0, 0, 0)
            UserForm2.ForeColor = RGB(255, 255, 255)
            Frame1.BackColor = RGB(0, 0, 0)
            Frame1.ForeColor = RGB(255, 255, 255)
            Frame2.BackColor = RGB(0, 0, 0)
            Frame2.ForeColor = RGB(255, 255, 255)
            Frame3.BackColor = RGB(0, 0, 0)
            Frame3.ForeColor = RGB(255, 255, 255)
            Label3.BackColor = RGB(0, 0, 0)
            Label3.ForeColor = RGB(255, 255, 255)
            Label2.BackColor = RGB(0, 0, 0)
            Label2.ForeColor = RGB(255, 255, 255)
            Plus.BackColor = RGB(0, 0, 0)
            Plus.ForeColor = RGB(255, 255, 255)
            Minus.BackColor = RGB(0, 0, 0)
            Minus.ForeColor = RGB(255, 255, 255)
            Multiply.BackColor = RGB(0, 0, 0)
            Multiply.ForeColor = RGB(255, 255, 255)
            Divide.BackColor = RGB(0, 0, 0)
            Divide.ForeColor = RGB(255, 255, 255)
            Clsar.BackColor = RGB(0, 0, 0)
            Clsar.ForeColor = RGB(255, 255, 255)
            Fin.BackColor = RGB(0, 0, 0)
            Fin.ForeColor = RGB(255, 255, 255)
            ToggleButton1.BackColor = RGB(153, 204, 255)
            ToggleButton1.ForeColor = RGB(0, 0, 0)
            
     Case False:
            ToggleButton1.BackColor = RGB(0, 0, 0)
            ToggleButton1.ForeColor = RGB(255, 255, 255)
            Equal.BackColor = RGB(153, 204, 255)
            Equal.ForeColor = RGB(0, 0, 0)
            Result.BackColor = RGB(153, 204, 255)
            Result.ForeColor = RGB(0, 0, 0)
            Number1.BackColor = RGB(153, 204, 255)
            Number1.ForeColor = RGB(0, 0, 0)
            Number2.BackColor = RGB(153, 204, 255)
            Number2.ForeColor = RGB(0, 0, 0)
            UserForm2.BackColor = RGB(153, 204, 255)
            UserForm2.ForeColor = RGB(0, 0, 0)
            Frame1.BackColor = RGB(153, 204, 255)
            Frame1.ForeColor = RGB(0, 0, 0)
            Frame2.BackColor = RGB(153, 204, 255)
            Frame2.ForeColor = RGB(0, 0, 0)
            Frame3.BackColor = RGB(153, 204, 255)
            Frame3.ForeColor = RGB(0, 0, 0)
            Label3.BackColor = RGB(153, 204, 255)
            Label3.ForeColor = RGB(0, 0, 0)
            Label2.BackColor = RGB(153, 204, 255)
            Label2.ForeColor = RGB(0, 0, 0)
            Plus.BackColor = RGB(153, 204, 255)
            Plus.ForeColor = RGB(0, 0, 0)
            Minus.BackColor = RGB(153, 204, 255)
            Minus.ForeColor = RGB(0, 0, 0)
            Multiply.BackColor = RGB(153, 204, 255)
            Multiply.ForeColor = RGB(0, 0, 0)
            Divide.BackColor = RGB(153, 204, 255)
            Divide.ForeColor = RGB(0, 0, 0)
            Clsar.BackColor = RGB(153, 204, 255)
            Clsar.ForeColor = RGB(0, 0, 0)
            Fin.BackColor = RGB(153, 204, 255)
            Fin.ForeColor = RGB(0, 0, 0)
            
     End Select
    
    
End Sub



Private Sub UserForm_Initialize()

            Equal.BackColor = RGB(153, 204, 255)
            Equal.ForeColor = RGB(0, 0, 0)
            Result.BackColor = RGB(153, 204, 255)
            Result.ForeColor = RGB(0, 0, 0)
            Number1.BackColor = RGB(153, 204, 255)
            Number1.ForeColor = RGB(0, 0, 0)
            Number2.BackColor = RGB(153, 204, 255)
            Number2.ForeColor = RGB(0, 0, 0)
            UserForm2.BackColor = RGB(153, 204, 255)
            UserForm2.ForeColor = RGB(0, 0, 0)
            Frame1.BackColor = RGB(153, 204, 255)
            Frame1.ForeColor = RGB(0, 0, 0)
            Frame2.BackColor = RGB(153, 204, 255)
            Frame2.ForeColor = RGB(0, 0, 0)
            Frame3.BackColor = RGB(153, 204, 255)
            Frame3.ForeColor = RGB(0, 0, 0)
            Label3.BackColor = RGB(153, 204, 255)
            Label3.ForeColor = RGB(0, 0, 0)
            Label2.BackColor = RGB(153, 204, 255)
            Label2.ForeColor = RGB(0, 0, 0)
            Plus.BackColor = RGB(153, 204, 255)
            Plus.ForeColor = RGB(0, 0, 0)
            Minus.BackColor = RGB(153, 204, 255)
            Minus.ForeColor = RGB(0, 0, 0)
            Multiply.BackColor = RGB(153, 204, 255)
            Multiply.ForeColor = RGB(0, 0, 0)
            Divide.BackColor = RGB(153, 204, 255)
            Divide.ForeColor = RGB(0, 0, 0)
            Clsar.BackColor = RGB(153, 204, 255)
            Clsar.ForeColor = RGB(0, 0, 0)
            Fin.BackColor = RGB(153, 204, 255)
            Fin.ForeColor = RGB(0, 0, 0)
            ToggleButton1.BackColor = RGB(0, 0, 0)
            ToggleButton1.ForeColor = RGB(255, 255, 255)
            Result.Visible = True

End Sub
