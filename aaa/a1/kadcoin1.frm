VERSION 5.00
Begin VB.Form kadcoin1 
   Caption         =   "Lucky Seven"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Palatino Linotype"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000A&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "End"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Spin"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   1965
      Left            =   1560
      Picture         =   "kadcoin1.frx":0000
      Stretch         =   -1  'True
      Top             =   960
      Width           =   2805
   End
   Begin VB.Label Label4 
      Caption         =   "Lucky Seven"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1095
      Left            =   360
      TabIndex        =   5
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "kadcoin1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Image1.Visible = False
Label1.Caption = Int(Rnd * 10)
Label2.Caption = Int(Rnd * 10)
Label3.Caption = Int(Rnd * 10)
If Label1.Caption = 7 Or Label2.Caption = 7 Or Label3.Caption = 7 Then
Image1.Visible = True
Beep
End If
End Sub

Private Sub Command2_Click()
End
End Sub
