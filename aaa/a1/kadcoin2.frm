VERSION 5.00
Begin VB.Form kadcoin2 
   Caption         =   "Lucky Seven"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4470
   BeginProperty Font 
      Name            =   "Palatino Linotype"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4470
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "End"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Coins"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   840
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
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "30"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1440
      TabIndex        =   8
      Top             =   1440
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1440
      TabIndex        =   7
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Lucky    Seven"
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
      Height          =   855
      Left            =   240
      TabIndex        =   6
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   240
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   1965
      Left            =   1440
      Picture         =   "kadcoin2.frx":0000
      Stretch         =   -1  'True
      Top             =   960
      Width           =   2805
   End
End
Attribute VB_Name = "kadcoin2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
a = Val(Label6.Caption)
Image1.Visible = False
Label1.Caption = Int(Rnd * 10)
Label2.Caption = Int(Rnd * 10)
Label3.Caption = Int(Rnd * 10)

If Label1.Caption = 7 Or Label2.Caption = 7 Or Label3.Caption = 7 Then
Image1.Visible = True
Beep
Label5.Caption = a + 10
Label6.Caption = Label5.Caption

Else
If a >= (Val(Label1.Caption) + Val(Label2.Caption) + Val(Label3.Caption)) Then
a = a - (Val(Label1.Caption) + Val(Label2.Caption) + Val(Label3.Caption))
Label5.Caption = a
Label6.Caption = Label5.Caption

Else
MsgBox (" Not Enough Coins ")
Label5.Caption = a
End If
End If
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Form_Load()
Label5.Caption = Label6.Caption
End Sub

