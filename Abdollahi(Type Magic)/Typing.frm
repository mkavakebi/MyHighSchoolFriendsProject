VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Type Magic"
   ClientHeight    =   6765
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8250
   LinkTopic       =   "Form1"
   ScaleHeight     =   6765
   ScaleWidth      =   8250
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   6495
      Left            =   6000
      TabIndex        =   1
      Top             =   45
      Width           =   2160
   End
   Begin VB.TextBox Text1 
      Height          =   6615
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   30
      Width           =   5940
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2835
      Top             =   900
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public h1 As Integer
Public m1 As Integer
Public s1 As Integer
Public h2 As Integer
Public m2 As Integer
Public s2 As Integer

Private Sub Text1_Change()
If Len(Text1.Text) = 1 Then
h1 = Hour(Now)
m1 = Minute(Now)
s1 = Second(Now)
End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
Static Plen As Integer
Dim T
If KeyCode = 13 Then
h2 = Hour(Now)
m2 = Minute(Now)
s2 = Second(Now)
T = T + Abs(h2 - h1) * 3600
T = T + Abs(m2 - m1) * 60
T = T + Abs(s2 - s1)
c = T / (Len(Text1.Text) - Plen)
Plen = Len(Text1.Text)
Select Case c
 Case Is > 3
 List1.AddItem "it's your first time typing"
Case Is > 2
 List1.AddItem "too bad"
Case Is > 1
List1.AddItem "Take Anouther I wish you get beter."
Case Is > 0.6
List1.AddItem "Nice."
Case Is > 0.5
List1.AddItem "good."
Case Is > 0.4
List1.AddItem "too nice."
Case Is > 0.3
List1.AddItem "Wounderfull!"
Case Is > 0.2
List1.AddItem "Prefect ! prefect!"
Case Is > 0.1
List1.AddItem "Ey Ostad!"
Case Is < 0.1
List1.AddItem ("Ey Ostad! Abar Ghodrat! Ghavi Heikal!")
End Select
End If
End Sub
