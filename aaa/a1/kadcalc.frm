VERSION 5.00
Begin VB.Form kadcalc 
   Caption         =   "Calculator"
   ClientHeight    =   3570
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6675
   LinkTopic       =   "Form2"
   ScaleHeight     =   3570
   ScaleWidth      =   6675
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option8 
      Caption         =   "Concatination"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   11
      Top             =   1560
      Width           =   1695
   End
   Begin VB.OptionButton Option7 
      Caption         =   "Exponentiation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   10
      Top             =   1200
      Width           =   1815
   End
   Begin VB.OptionButton Option6 
      Caption         =   "Remainder"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   9
      Top             =   480
      Width           =   1455
   End
   Begin VB.OptionButton Option5 
      Caption         =   "Integer divison"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   8
      Top             =   840
      Width           =   1695
   End
   Begin VB.OptionButton Option4 
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   480
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Substract"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   840
      Width           =   1455
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Divide"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   1560
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Multiply"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   3
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calculate"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      Picture         =   "kadcalc.frx":0000
      TabIndex        =   2
      Top             =   2880
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   1455
   End
   Begin VB.Frame frame1 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   240
      TabIndex        =   12
      Top             =   120
      Width           =   1935
      Begin VB.Label Label3 
         Caption         =   "Variable 2:"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Variable 1:"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   2280
      TabIndex        =   15
      Top             =   2040
      Width           =   4095
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label4 
         Caption         =   "Result:"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1815
      Left            =   2280
      TabIndex        =   18
      Top             =   240
      Width           =   4095
   End
End
Attribute VB_Name = "kadcalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Option1.Value = True Then
If IsNumeric(Text1.Text) = True And IsNumeric(Text2.Text) = True Then
Text3.Text = Val(Text1.Text) * Val(Text2.Text)
Else
MsgBox "error"
End If

End If
If Option2.Value = True Then
Text3.Text = Val(Text1.Text) / Val(Text2.Text)
End If
If Option3.Value = True Then
Text3.Text = Val(Text1.Text) - Val(Text2.Text)
End If
If Option4.Value = True Then
Text3.Text = Val(Text1.Text) + Val(Text2.Text)
End If
If Option5.Value = True Then
Text3.Text = Val(Text1.Text) \ Val(Text2.Text)
End If
If Option6.Value = True Then
Text3.Text = Val(Text1.Text) Mod Val(Text2.Text)
End If
If Option7.Value = True Then
Text3.Text = Val(Text1.Text) ^ Val(Text2.Text)
End If
If Option8.Value = True Then
Text3.Text = Text1.Text + Text2.Text
End If
End Sub

Private Sub Command2_Click()
End
End Sub


