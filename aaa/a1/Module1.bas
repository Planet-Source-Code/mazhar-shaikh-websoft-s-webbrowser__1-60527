Attribute VB_Name = "Module1"
Option Explicit
Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal ipbuffer As String, nsize As Long) As Long

