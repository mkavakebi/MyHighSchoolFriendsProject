VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   5625
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10125
   LinkTopic       =   "Form1"
   ScaleHeight     =   5625
   ScaleWidth      =   10125
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timery 
      Interval        =   10
      Left            =   3330
      Top             =   3345
   End
   Begin VB.Timer Timerx 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3315
      Top             =   2880
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   315
      Left            =   7800
      TabIndex        =   1
      Top             =   0
      Width           =   1275
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   315
      Left            =   1080
      TabIndex        =   0
      Top             =   0
      Width           =   1275
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00404080&
      BackStyle       =   1  'Opaque
      FillStyle       =   2  'Horizontal Line
      Height          =   1080
      Left            =   9060
      Shape           =   4  'Rounded Rectangle
      Top             =   2175
      Width           =   240
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      X1              =   9165
      X2              =   9165
      Y1              =   0
      Y2              =   5625
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404080&
      BackStyle       =   1  'Opaque
      FillStyle       =   2  'Horizontal Line
      Height          =   1080
      Left            =   855
      Shape           =   4  'Rounded Rectangle
      Top             =   2175
      Width           =   240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      X1              =   960
      X2              =   960
      Y1              =   0
      Y2              =   5625
   End
   Begin VB.Shape Ball 
      BackStyle       =   1  'Opaque
      BorderStyle     =   6  'Inside Solid
      FillStyle       =   7  'Diagonal Cross
      Height          =   495
      Left            =   4440
      Shape           =   2  'Oval
      Top             =   2520
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim kx, ky
Dim N_a, N_b

Private Sub Form_Load()
kx = 1
ky = 1
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape2.Top = Y - Shape2.Height / 2
End Sub

Private Sub Timerx_Timer()
Ball.Left = Ball.Left + 100 * kx
If Ball.Left > Width - Ball.Width - 200 - 1000 Then
kx = -kx
If Abs((Ball.Top + Ball.Height / 2) - (Shape2.Top + Shape2.Height / 2)) - 500 > Shape2.Height / 2 Then Label1.Caption = Val(Label1.Caption) + 1
End If
If Ball.Left < 1000 Then kx = -kx: N_a = N_a + 1
End Sub

Private Sub Timery_Timer()
Ball.Top = Ball.Top + 100 * ky
If Ball.Top > Height - 500 - Ball.Height Then ky = -ky
If Ball.Top < 10 Then ky = -ky: N_b = N_b + 1
Shape1.Top = (Ball.Top + Ball.Height / 2) - Shape1.Height / 2
End Sub
