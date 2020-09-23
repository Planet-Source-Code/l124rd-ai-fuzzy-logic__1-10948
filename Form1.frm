VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Hello"
   ClientHeight    =   5880
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6555
   LinkTopic       =   "Form1"
   ScaleHeight     =   5880
   ScaleWidth      =   6555
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   840
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    'Picture1.Cls
    Do Until i = 50
    Form1.Print "Working Hard?"
    Form1.Print "Or hardly Working?"
    i = i + 1
    Loop
End Sub

