VERSION 5.00
Begin VB.Form kadconvt 
   Caption         =   "Conversion"
   ClientHeight    =   2640
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5160
   BeginProperty Font 
      Name            =   "Palatino Linotype"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   ScaleHeight     =   2640
   ScaleWidth      =   5160
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Quit"
      Height          =   375
      Left            =   3720
      TabIndex        =   10
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   2040
      TabIndex        =   9
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calculate"
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Height          =   1815
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Width           =   2415
      Begin VB.TextBox Text2 
         Height          =   390
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   2175
      End
      Begin VB.ComboBox Combo2 
         Height          =   390
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Output Format:"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
      Begin VB.ComboBox Combo1 
         Height          =   390
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Height          =   390
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Input Format:"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
   End
End
Attribute VB_Name = "kadconvt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim x As Long, i As Long, l As Long, ip As Long, op As Long
Dim ips As String, ops As String

Private Sub Command1_Click()
Select Case Combo1.ListIndex
    Case 0
        Select Case Combo2.ListIndex
            Case 0
                Text2.Text = Text1.Text
            Case 1
                Text2.Text = dtob(Val(Text1.Text))
            Case 2
                Text2.Text = dtoo(Val(Text1.Text))
            Case 3
                ip = Val(Text1.Text)
                ips = dtob(ip)
                Text2.Text = btog(ips)
            Case 4
                Text2.Text = dtoh(Val(Text1.Text))
            Case 5
                ip = Val(Text1.Text)
                ips = dtob(ip)
                Text2.Text = bto1(ips)
            Case 6
                ip = Val(Text1.Text)
                ips = dtob(ip)
                Text2.Text = bto2(ips)
        End Select
    Case 1
        Select Case Combo2.ListIndex
            Case 0
                Text2.Text = btod(Text1.Text)
            Case 1
                Text2.Text = Text1.Text
            Case 2
                ip = btod(Text1.Text)
                Text2.Text = dtoo(ip)
            Case 3
                Text2.Text = btog(Text1.Text)
            Case 4
                ip = btod(Text1.Text)
                Text2.Text = dtoh(ip)
            Case 5
                Text2.Text = bto1(Text1.Text)
            Case 6
                Text2.Text = bto2(Text1.Text)
        End Select
    Case 2
        Select Case Combo2.ListIndex
            Case 0
                Text2.Text = otod(Text1.Text)
            Case 1
                ip = otod(Text1.Text)
                Text2.Text = dtob(ip)
            Case 2
                Text2.Text = Text1.Text
            Case 3
                ip = otod(Text1.Text)
                ips = dtob(ip)
                Text2.Text = btog(ips)
            Case 4
                ip = otod(Text1.Text)
                Text2.Text = dtoh(ip)
            Case 5
                ip = otod(Text1.Text)
                ips = dtob(ip)
                Text2.Text = bto1(ips)
            Case 6
                ip = otod(Text1.Text)
                ips = dtob(ip)
                Text2.Text = bto2(ips)
        End Select
    Case 3
        Select Case Combo2.ListIndex
            Case 0
                ips = gtob(Text1.Text)
                Text2.Text = btod(ips)
            Case 1
                Text2.Text = gtob(Text1.Text)
            Case 2
                ips = gtob(Text1.Text)
                ip = btod(ips)
                Text2.Text = dtoo(ip)
            Case 3
                Text2.Text = Text1.Text
            Case 4
                ips = gtob(Text1.Text)
                ip = btod(ips)
                Text2.Text = dtoh(ip)
            Case 5
                ips = gtob(Text1.Text)
                Text2.Text = bto1(ips)
            Case 6
                ips = gtob(Text1.Text)
                Text2.Text = bto2(ips)
        End Select
    Case 4
        Select Case Combo2.ListIndex
            Case 0
                ips = htob(Text1.Text)
                Text2.Text = btod(ips)
            Case 1
                Text2.Text = htob(Text1.Text)
            Case 2
                ips = htob(Text1.Text)
                ip = btod(ips)
                Text2.Text = dtoo(ip)
            Case 3
                ips = htob(Text1.Text)
                Text2.Text = btog(ips)
            Case 4
                Text2.Text = Text1.Text
            Case 5
                ips = htob(Text1.Text)
                Text2.Text = bto1(ips)
            Case 6
                ips = htob(Text1.Text)
                Text2.Text = bto2(ips)
        End Select
End Select
End Sub

Private Sub Command2_Click()
ref
Text1.Text = ""
Text2.Text = ""
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Form_Load()
Combo1.AddItem "Decimal"
Combo1.AddItem "Binary"
Combo1.AddItem "Octal"
Combo1.AddItem "Gray"
Combo1.AddItem "Hex"
Combo2.AddItem "Decimal"
Combo2.AddItem "Binary"
Combo2.AddItem "Octal"
Combo2.AddItem "Gray"
Combo2.AddItem "Hex"
Combo2.AddItem "1's Complement"
Combo2.AddItem "2's Complement"
End Sub

Public Sub ref()
x = 0
i = 0
ip = 0
op = 0
ips = ""
ops = ""
End Sub

Public Function dtob(ip As Long) As Long
While (ip <> 0)
    x = ip Mod 2
    ip = ip \ 2
    op = op + x * 10 ^ i
    i = i + 1
Wend
dtob = op
ref
End Function

Public Function dtoo(ip As Long) As Long
While (ip <> 0)
    x = ip Mod 8
    ip = ip \ 8
    op = op + x * 10 ^ i
    i = i + 1
Wend
dtoo = op
ref
End Function
Public Function dtoh(ip As Long) As String
While (ip <> 0)
    x = ip Mod 16
    ip = ip \ 16
    If x > 9 Then
        Select Case x
            Case "10"
                ops = "A"
            Case "11"
                ops = "B"
            Case "12"
                ops = "C"
            Case "13"
                ops = "D"
            Case "14"
                ops = "E"
            Case "15"
                ops = "F"
        End Select
        ips = ips & ops
    Else
        ips = ips & Str(x)
    End If
Wend
ops = StrReverse(ips)
dtoh = ops
ref
End Function
Public Function btod(ips As String) As Long
l = Len(ips)
For i = 0 To l - 1
    x = CLng(Mid(ips, l - i, 1))
    op = op + x * (2 ^ i)
Next
btod = op
ref
End Function

Public Function bto1(ips As String) As String
l = Len(ips)
For i = 0 To l - 1
    x = Mid(ips, i + 1, 1)
    If x = 0 Then
        x = 1
    Else
        x = 0
    End If
    ops = ops & Str(x)
Next
bto1 = ops
ref
End Function
Public Function bto2(ips As String) As String
Dim c As Integer
l = Len(ips)
For i = 0 To l - 1
    x = Mid(ips, l - i, 1)
    If x = 1 Then
        If c = 0 Then
            op = 1
            c = 1
        Else
            x = 0
            op = op + x * 10 ^ i
        End If
    Else
        x = 1
        op = op + x * 10 ^ i
    End If
Next
ops = Str(op)
bto2 = ops
ref
End Function
Public Function btog(ips As String) As String
l = Len(ips)
x = Left(ips, 1)
ops = Str(x)
For i = 1 To l - 1
    x = x Xor Val(Mid(ips, i + 1, 1))
    ops = ops & Str(x)
Next
btog = ops
ref
End Function
Public Function otod(ips As String) As Long
l = Len(ips)
For i = 0 To l - 1
    x = CLng(Mid(ips, l - i, 1))
    op = op + x * (8 ^ i)
Next
otod = op
ref
End Function

Public Function gtob(ips As String) As String
l = Len(ips)
ops = Left(ips, 1)
For i = 1 To l - 1
    ops = ops & (Val(Mid(ips, i, 1)) Xor Val(Mid(ips, i + 1, 1)))
Next
gtob = ops
ref
End Function

Public Function htob(ips As String) As String
Dim c As Integer, y As Integer, a As Long
y = Len(ips)
For c = 1 To y
    a = Asc(Mid(ips, c, 1))
    Select Case a
        Case "65" Or "97"
            ops = ops & 1010
        Case "66" Or "98"
            ops = ops & 1011
        Case "67" Or "99"
            ops = ops & 1100
        Case "68" Or "100"
            ops = ops & 1101
        Case "69" Or "101"
            ops = ops & 1110
        Case "70" Or "102"
            ops = ops & 1111
        Case Else
            ops = dtob(Chr(a))
            x = Len(ops)
            If x < 4 Then
                For i = x + 1 To 4
                    ops = 0 & ops
                Next
            End If
    End Select
Next
htob = ops
ref
End Function
