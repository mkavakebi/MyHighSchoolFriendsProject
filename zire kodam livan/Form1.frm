VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5730
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   ScaleHeight     =   5730
   ScaleWidth      =   6600
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "[_*-*}(()){*-*_]"
      Height          =   780
      Left            =   4455
      TabIndex        =   0
      Top             =   240
      Width           =   1590
   End
   Begin VB.Shape Shape 
      Height          =   3000
      Index           =   1
      Left            =   2250
      Shape           =   3  'Circle
      Top             =   1935
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.Shape Shape 
      Height          =   3000
      Index           =   2
      Left            =   0
      Shape           =   3  'Circle
      Top             =   1905
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.Shape Shape 
      Height          =   3000
      Index           =   3
      Left            =   1125
      Shape           =   3  'Circle
      Top             =   0
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1305
      Index           =   1
      Left            =   3150
      TabIndex        =   3
      Top             =   2910
      Width           =   1575
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1305
      Index           =   2
      Left            =   390
      TabIndex        =   2
      Top             =   2865
      Width           =   1575
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1305
      Index           =   3
      Left            =   1830
      TabIndex        =   1
      Top             =   540
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim T ' shomare mohre

Private Sub Command1_Click()
Randomize Timer
T = Fix(Rnd * 3) + 1
Shape(T).Visible = True
DoEvents
Sleep (500)
Shape(T).Visible = False
DoEvents
''''''''''charkhesh ha
For i = 1 To 5
    r = Fix(Rnd * 3) + 1
    Select Case r
    Case 1:
        Shape(2).Visible = True: Shape(3).Visible = True
        If T = 2 Then
        T = 3
        Else
        If T = 3 Then T = 2
        End If
    Case 2:
        Shape(1).Visible = True: Shape(3).Visible = True
        If T = 1 Then
        T = 3
        Else
        If T = 3 Then T = 1
        End If
    Case 3:
        Shape(1).Visible = True: Shape(2).Visible = True
        If T = 1 Then
        T = 2
        Else
        If T = 2 Then T = 1
        End If
    End Select
    Sleep (500)
    DoEvents
    Shape(1).Visible = False
    Shape(2).Visible = False
    Shape(3).Visible = False
    Sleep (500)
    DoEvents
Next
End Sub

Private Sub Label_Click(Index As Integer)
If T = Index Then
MsgBox "Aferin " + CStr(T)
Else
MsgBox "Khak to saret " + CStr(T)
End If
End Sub

