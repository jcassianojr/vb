Attribute VB_Name = "wrptxvars"
Option Explicit
'requeridas wrpt/wrptx
Public zAppPath As String
Public zMANA5EMP As String
Public ZMANA5CON As String
Public ZMANA5IMP As String
Public ZRELAT As String
Public DBWRPT As String
Public Dbname As String
Public cARQERRO As String
Public zRPTARQ As String
Public zgrp As String
Public ZGRPSUB As String
Public cTIPORPT As String
Public zIDRPTUSR As Long
'requeridas wrpt
Public strCodeKey As String
Public dbuser As String
Public cMENU As String
Public cTIPO As String
Public iMENU As Long
Public iImage As Long

Public Sub Main()
  EnabledVisualStyles
  frmPRINCIPAL.Show
End Sub

Public Function ArqRPTUsr() As String
  ArqRPTUsr = "RPTUSR"
  Select Case cTIPORPT
  Case "R"
    ArqRPTUsr = "RPTUSR"
  Case "F"
    ArqRPTUsr = "RPTFOLUSR"
  Case "I"
    ArqRPTUsr = "RPTINTUSR"
  Case "D"
    ArqRPTUsr = "RTFUSR"
  End Select
End Function

