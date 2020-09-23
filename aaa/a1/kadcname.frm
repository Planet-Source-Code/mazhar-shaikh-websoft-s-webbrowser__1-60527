VERSION 5.00
Begin VB.Form kadcname 
   Caption         =   "Form1"
   ClientHeight    =   990
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3600
   LinkTopic       =   "Form1"
   ScaleHeight     =   990
   ScaleWidth      =   3600
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Comp. Name"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "kadcname"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strbuffer As String
Dim lgbuffer As Long
Dim lgstatus As Long

Private Sub Command1_Click()
lgbuffer = 255
strbuffer = String$(lgbuffer, " ")
lgstatus = GetComputerName(strbuffer, lgbuffer)
If lgstatus <> 0 Then
    MsgBox "Computer name is " & (strbuffer)
End If

End Sub

Private Sub Command2_Click()
End
End Sub
