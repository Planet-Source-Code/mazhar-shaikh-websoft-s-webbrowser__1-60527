VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form kadchart 
   Caption         =   "Chart"
   ClientHeight    =   2730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5130
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
   ScaleHeight     =   2730
   ScaleWidth      =   5130
   StartUpPosition =   3  'Windows Default
   Begin MSChart20Lib.MSChart chart 
      Height          =   2655
      Left            =   1560
      OleObjectBlob   =   "kadchart.frx":0000
      TabIndex        =   3
      Top             =   0
      Width           =   3375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Plot Three"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Plot Two"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "kadchart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ar() As Variant, br() As Variant
Dim a As Integer, b As Integer, c As Integer

Private Sub Command1_Click()
With Me.chart
    .Visible = True
    .RowCount = a
    .ChartData = ar
    .ShowLegend = True
    .chartType = VtChChartType2dStep
    .Legend.Location.LocationType = VtChLocationTypeTop
    .Plot.UniformAxis = False
    .ColumnCount = 2
    .ColumnLabelCount = 2
    .Column = 1
    .ColumnLabel = "Data set 1"
    .Column = 2
    .ColumnLabel = "Data set 2"
    .Refresh
End With
End Sub

Private Sub Command2_Click()
With Me.chart
    .Visible = True
    .RowCount = b
    .ChartData = br
    .ShowLegend = True
    .chartType = VtChChartType2dLine
    .Legend.Location.LocationType = VtChLocationTypeTop
    .Plot.UniformAxis = False
    .ColumnCount = 3
    .ColumnLabelCount = 3
    .Column = 1
    .ColumnLabel = "Data set 1"
    .Column = 2
    .ColumnLabel = "Data set 2"
    .Column = 3
    .ColumnLabel = "Data set 3"
    .Refresh
End With
        
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Form_Load()
chart.Visible = False
a = 0
ReDim ar(1 To 10, 2)
For c = 1 To 10
    ar(c, 0) = "" & c
    ar(c, 1) = c * 2
    ar(c, 2) = c + 6
    a = a + 1
Next
b = 0
ReDim br(1 To 10, 3)
For c = 1 To 10
    br(c, 0) = "" & c * 10
    br(c, 1) = c
    br(c, 2) = c * 2
    br(c, 3) = c * 3
    b = b + 1
Next
End Sub
