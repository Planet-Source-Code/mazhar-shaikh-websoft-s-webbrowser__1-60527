VERSION 5.00
Begin VB.Form kadarc 
   Caption         =   "Circles"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7725
   LinkTopic       =   "Form1"
   ScaleHeight     =   5400
   ScaleWidth      =   7725
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   3495
      Left            =   120
      ScaleHeight     =   3435
      ScaleWidth      =   7395
      TabIndex        =   11
      Top             =   1800
      Width           =   7455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Circles Ellipses"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7455
      Begin VB.CommandButton Command1 
         Caption         =   "Clear"
         Height          =   975
         Left            =   6480
         TabIndex        =   10
         Top             =   360
         Width           =   615
      End
      Begin VB.HScrollBar HScroll3 
         Height          =   255
         Left            =   360
         Max             =   1000
         Min             =   1
         TabIndex        =   3
         Top             =   1080
         Value           =   1
         Width           =   3735
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   255
         Left            =   360
         Max             =   15
         Min             =   1
         TabIndex        =   2
         Top             =   720
         Value           =   1
         Width           =   3735
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   360
         Max             =   6
         Min             =   1
         TabIndex        =   1
         Top             =   360
         Value           =   1
         Width           =   3735
      End
      Begin VB.Label Label6 
         Caption         =   "Size"
         Height          =   255
         Left            =   5520
         TabIndex        =   9
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Eccentricity"
         Height          =   255
         Left            =   5520
         TabIndex        =   8
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Arc"
         Height          =   255
         Left            =   5520
         TabIndex        =   7
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1 to 1000"
         Height          =   255
         Left            =   4200
         TabIndex        =   6
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1 to 15"
         Height          =   255
         Left            =   4200
         TabIndex        =   5
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1 to 6"
         Height          =   255
         Left            =   4200
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "kadarc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const pi = 3.142
Dim r1, g1, b1 As Integer

Private Sub Command1_Click()
Picture1.Cls

End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Dim a, b, c As Integer
a = HScroll1.Value
b = HScroll2.Value
c = HScroll3.Value

r1 = 255 * Rnd
g1 = 255 * Rnd
b1 = 255 * Rnd

Picture1.Circle (X, Y), c, RGB(r1, g1, b1), , a, -pi / b
End If

End Sub
