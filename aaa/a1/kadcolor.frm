VERSION 5.00
Begin VB.Form kadcolor 
   Caption         =   "Colour Timer"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   2
      Left            =   2160
      Top             =   1320
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   2895
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "kadcolor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ired As Integer, igreen As Integer, iblue As Integer

Private Sub Form_Load()

End Sub

Private Sub Timer1_Timer()
If ired < 255 Then
    ired = ired + 1
Else
    If igreen < 255 Then
        igreen = igreen + 1
    Else
        If iblue < 255 Then
            iblue = iblue + 1
        End If
    End If
End If
If ired >= 255 And igreen >= 255 And iblue >= 255 Then
ired = 0
igreen = 0
iblue = 0
End If
Label1.BackColor = RGB(ired, igreen, iblue)
End Sub
