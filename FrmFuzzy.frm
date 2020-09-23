VERSION 5.00
Begin VB.Form FrmFuzzy 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Fuzzy Logic"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   222
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   286
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Move 
      Interval        =   500
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer Check 
      Interval        =   1000
      Left            =   3960
      Top             =   0
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      Caption         =   "a: left, d:right w:up, s:down"
      ForeColor       =   &H00C0C0FF&
      Height          =   495
      Left            =   1680
      TabIndex        =   7
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   3240
      TabIndex        =   6
      Top             =   2880
      Width           =   135
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1200
      TabIndex        =   5
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Your Health:"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2400
      Width           =   975
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0080FF80&
      BorderStyle     =   3  'Dot
      Height          =   255
      Left            =   1560
      Top             =   -120
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3480
      TabIndex        =   3
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Enemy State:"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3240
      TabIndex        =   2
      Top             =   2520
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H000000FF&
      BorderStyle     =   3  'Dot
      Height          =   735
      Left            =   1320
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Enemy"
      Height          =   255
      Left            =   1800
      TabIndex        =   1
      Top             =   240
      Width           =   555
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "You"
      Height          =   255
      Left            =   1920
      TabIndex        =   0
      Top             =   2160
      Width           =   375
   End
End
Attribute VB_Name = "FrmFuzzy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check_Timer()
    'Changes the color of the decision thing
    If Label7.BackColor = &HFF00& Then
        Label7.BackColor = &HFF&
    Else
        Label7.BackColor = &HFF00&
    End If
    'Calls the Calculate Function
    Calculate
End Sub

Private Sub Command1_Click()
    'makes you go forward
    Label1.Top = Label1.Top - 1
End Sub

Private Sub Command2_Click()
    'makes you go left
    Label1.Left = Label1.Left - 1
End Sub

Private Sub Command3_Click()
    'makes you go back
    Label1.Top = Label1.Top + 1
End Sub

Private Sub Command4_Click()
    'Makes you go right
    Label1.Left = Label1.Left + 1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 97 Then Label1.Left = Label1.Left - 1
    If KeyAscii = 100 Then Label1.Left = Label1.Left + 1
    If KeyAscii = 119 Then Label1.Top = Label1.Top - 1
    If KeyAscii = 115 Then Label1.Top = Label1.Top + 1
    'Used the following line for Debugging
    'MsgBox (Asc("a") & ";" & Asc("w") & ";" & Asc("d") & ";" & Asc("s") & ";")
End Sub

Private Sub Form_Load()
    'Sets all the variables to default
    Hp = 100
    Sleep = 5
    Attack = 1
    Gaurd = 10
    'Runs the calculate function
    Calculate
End Sub


Private Sub Move_Timer()
    'moves closer to you if its in attack mode
    If State = "Attack" Then
        If Label2.Left > Label1.Left Then
            Label2.Left = Label2.Left - 1
        End If
        If Label2.Left < Label1.Left Then
            Label2.Left = Label2.Left + 1
        End If
        If Label2.Top > Label1.Top Then
            Label2.Top = Label2.Top - 1
        End If
        If Label2.Top < Label1.Top Then
            Label2.Top = Label2.Top + 1
        End If
        Shape1.Left = Label2.Left - (Shape1.Width / 2) + (Label2.Width / 2)
        Shape1.Top = Label2.Top - (Shape1.Height / 2) + (Label2.Height / 2)
    End If
    'Checks if its in the gaurd state
    If State = "Gaurd" Then
        'checks to see if your in the red 'AGRESSIVE' box
         If Label1.Left >= Shape1.Left And Label1.Left <= Shape1.Left + Shape1.Width And Label1.Top >= Shape1.Top And Label1.Top <= (Shape1.Top + (Shape1.Height * 2)) Then
            'Checks to make sure the attack is 1
            'before lowering sleep and raising
            'attack
            If Attack <= 1 Then
                Attack = Attack + 14
                Gaurd = Gaurd - 9
                Sleep = Sleep - 4
            End If
        End If
    End If
    'Checks to see if you've won
    If Label1.Left >= Shape2.Left And Label1.Left <= (Shape2.Left + Shape2.Width) And Label1.Top >= Shape2.Top And Label1.Top <= (Shape2.Top + Shape2.Height) Then
        msg = MsgBox("YOU WIN!", vbOKOnly, "YOU WIN")
        End
    End If
    'checks to see if it should minus your HPs
    If Label1.Left >= Label2.Left And Label1.Left <= (Label2.Left + Label2.Width) And Label1.Top >= Label2.Top And Label1.Top <= (Label2.Top + Label2.Height) Then
        Hp = Hp - 30
    End If
    'Sees if hp is lower then 0
    If Hp <= 0 Then
        msg = MsgBox("YOU LOSE" & vbCrLf & "GAME OVER", vbOKOnly, "GAME OVER")
        End
    End If
    'Sets the HP label equal to your hp's
    Label6.Caption = Hp
End Sub
