VERSION 5.00
Begin VB.Form kadtime 
   Caption         =   "Comp Timer"
   ClientHeight    =   1350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3810
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   1350
   ScaleWidth      =   3810
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1680
      Top             =   720
   End
   Begin VB.OptionButton Option2 
      Caption         =   "12 Hrs."
      Height          =   270
      Left            =   2400
      TabIndex        =   7
      Top             =   840
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Caption         =   "24 Hrs."
      Height          =   270
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   240
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   240
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   240
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   ":"
      Height          =   375
      Index           =   1
      Left            =   1800
      TabIndex        =   5
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   ":"
      Height          =   375
      Index           =   0
      Left            =   840
      TabIndex        =   4
      Top             =   240
      Width           =   255
   End
End
Attribute VB_Name = "kadtime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X As Boolean
Dim h As Byte, s As Byte, m As Byte

Private Sub Form_Load()
h = Hour(Time)
m = Minute(Time)
s = Second(Time)
Text1.Text = h
Text2.Text = m
Text3.Text = s
Text4.Visible = False
End Sub

Private Sub Option1_Click()
If Text4.Text = "AM" Then
     If Text1.Text = "12" Then
         Text1.Text = Val(Text1.Text) - 12
     End If
  Else
       If Val(Text1.Text) < 12 Then
          Text1.Text = Val(Text1.Text) + 12
       End If
 End If
 Text4.Visible = False
End Sub

Private Sub Option2_Click()
Text4.Visible = True
 If h = 0 Then
  Text1.Text = "12"
  Text4.Text = "AM"
   X = False
  End If
  
  If h < 12 Then
  Text4.Text = "AM"
  X = False
  End If
  
  If h = 12 Then
  Text4.Text = "PM"
  X = True
  End If
  
  If h > 12 Then
  Text1.Text = Val(Text1.Text) - 12
  Text4.Text = "PM"
  X = True
  End If
  
End Sub

Private Sub Timer1_Timer()
s = s + 1
Text3.Text = s
If s = 60 Then
m = m + 1
s = 0
Text2.Text = m
Text3.Text = s
 If m = 60 Then
 h = h + 1
 m = 0
 Text1.Text = h
 Text2.Text = m
 If Option1.Value = True Then
 If h = 24 Then
 h = 0
 Text1.Text = h
 End If
 Else
 If h = 12 Then
   If X = True Then
   Text4.Text = "AM"
   X = False
   Else
    Text4.Text = "PM"
    X = True
    End If
    End If
    If h = 13 Then
    Text1.Text = "1"
    End If
    If h = 24 Then
    If X = True Then
    Text4.Text = "AM"
    X = False
    h = 0
    Text1.Text = h + 12
    End If
    End If
    End If
    End If
    End If
    
End Sub
