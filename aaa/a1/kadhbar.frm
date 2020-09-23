VERSION 5.00
Begin VB.Form kadhbar 
   Caption         =   "Form1"
   ClientHeight    =   2520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   ScaleHeight     =   2520
   ScaleWidth      =   4500
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Quit"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   2040
      Width           =   855
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   375
      Left            =   240
      Max             =   255
      TabIndex        =   2
      Top             =   1560
      Width           =   2655
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   375
      Left            =   240
      Max             =   255
      TabIndex        =   1
      Top             =   960
      Width           =   2655
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      Left            =   240
      Max             =   255
      TabIndex        =   0
      Top             =   360
      Width           =   2655
   End
End
Attribute VB_Name = "kadhbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim a As String
s = MsgBox("Are you sure?", vbYesNo)
If s = vbYes Then
End
End If
End Sub

Private Sub HScroll1_Change()
Form1.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
End Sub

Private Sub HScroll1_Scroll()
Form1.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
End Sub

Private Sub HScroll2_Change()
Form1.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
End Sub

Private Sub HScroll2_Scroll()
Form1.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
End Sub

Private Sub HScroll3_Change()
Form1.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
End Sub

Private Sub HScroll3_Scroll()
Form1.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
End Sub
