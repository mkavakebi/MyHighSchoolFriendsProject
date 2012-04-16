VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "comct232.ocx"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   4125
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9765
   LinkTopic       =   "Form1"
   ScaleHeight     =   275
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   651
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2580
      Left            =   7905
      TabIndex        =   0
      Top             =   90
      Width           =   1665
      Begin ComCtl2.UpDown UpDown1 
         Height          =   285
         Left            =   1276
         TabIndex        =   7
         Top             =   540
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   327681
         Value           =   30
         BuddyControl    =   "Text1"
         BuddyDispid     =   196614
         OrigLeft        =   240
         OrigTop         =   1500
         OrigRight       =   495
         OrigBottom      =   1530
         Max             =   360
         Min             =   -360
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   150
         TabIndex        =   2
         Text            =   "50"
         Top             =   1125
         Width           =   1380
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   150
         TabIndex        =   1
         Text            =   "0"
         Top             =   540
         Width           =   1125
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Reset"
         Height          =   330
         Left            =   135
         TabIndex        =   5
         Top             =   2115
         Width           =   1425
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Fire"
         Default         =   -1  'True
         Height          =   495
         Left            =   135
         TabIndex        =   6
         Top             =   1635
         Width           =   1425
      End
      Begin VB.Label Label2 
         Caption         =   "Velocity :"
         Height          =   240
         Left            =   60
         TabIndex        =   4
         Top             =   885
         Width           =   570
      End
      Begin VB.Label Label1 
         Caption         =   "Alpha :"
         Height          =   240
         Left            =   60
         TabIndex        =   3
         Top             =   240
         Width           =   570
      End
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00004080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   375
      Left            =   0
      Top             =   3765
      Width           =   9870
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   651
      Y1              =   250
      Y2              =   250
   End
   Begin VB.Shape Shape1 
      Height          =   150
      Left            =   1500
      Top             =   3615
      Width           =   150
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Const g = 10
Const Pi = 3.14159265358979
Dim V0
Dim Al As Single
Dim H As Single
Dim D As Single
Dim T As Single
1:
V0 = Val(Text2.Text)
Al = Val(Text1.Text)
Al = (Pi * Al) / 180
For T = 0 To 100 Step 0.02
vu = V0 * Sin(Al)
vd = V0 * Cos(Al)
H = -(g / 2) * (T ^ 2) + (vu * T)
D = vd * T
If H < 0 Then
Exit For
End If
Shape1.Left = D + 100
Shape1.Top = Line1.Y1 - H - Shape1.Height + 1
x = Shape1.Left + (Shape1.Width / 2)
y = Shape1.Top + Shape1.Height - 1
PSet (x, y)
DoEvents
Next
'Text1.Text = Val(Text1.Text) + 3
'If Text1.Text = "90" Then Exit Sub
'GoTo 1
End Sub

Private Sub Command2_Click()
Cls
Shape1.Left = 100
Shape1.Top = Line1.Y1 - Shape1.Height + 1
End Sub
