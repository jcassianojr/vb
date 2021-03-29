VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F22668DE-E08D-467B-8E41-13900013BD5F}#2.7#0"; "VBextra2.OCX"
Begin VB.Form escMANREL 
   Caption         =   "Escolha o grupo de Relatorio"
   ClientHeight    =   6330
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   ScaleHeight     =   6330
   ScaleWidth      =   8880
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   6330
      Left            =   7410
      TabIndex        =   1
      Top             =   0
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   11165
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      TextAlignment   =   1
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   9551
      _Version        =   393216
   End
   Begin vbExtra.FlexFn FlexFn1 
      Height          =   405
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   714
      PageNumbersFormat=   ""
      PageNumbersFormatIndex=   0
      BeginProperty PageNumbersFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "escMANREL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdSair_Click()
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub FilRelat()
 
    Dim cARQ As String
    Dim cSQL As String
    cARQ = ZMANA5CON
   
    If cTIPO = "R" Then

        cSQL = "SELECT MENU,CODIGO,NOME FROM MANREL WHERE MENU='" & zgrp & "' order by CODIGO"

    End If

    If cTIPO = "P" Then

        cSQL = "SELECT MENU,CODIGO,NOME FROM PADREL WHERE MENU='" & zgrp & "' order by CODIGO"

    End If
   
    cARQ = GeraConn(cARQ, "JETFOX")
    MontaGridUltra Grid, 3, Array(400, 1000, 6000), Array("GRP", "RPT", "Nome"), _
        Array("L$Menu", "L$COdigo", "Nome"), cARQ, cSQL

   

End Sub

Private Sub Form_Load()
    Center Me
    xmontatoolbar Me.Toolbar1, "escMANREL", True
    FilRelat
    If FixStr(eLOCALIZA) <> "" Then LocalizaGri1 Grid, eLOCALIZA, 1
End Sub

Private Sub Grid_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        imprima_click
    End If
    If KeyAscii > 31 And KeyAscii < 123 Then
        LocalizaGrid Grid, Chr(KeyAscii), 1, False
    End If


End Sub

'Private Sub Grid_SelChange()

 '   With Grid
  '      If .Rows > 2 Then
   '         .Col = .Cols - 1
    '        .ColSel = 0
     '       .TopRow = .Row
      '  End If
    'End With
'End Sub

Private Sub imprima_click()

    Dim eRUN             As String
    Grid.Col = 1
    zRPT = Grid
    eRUN = ZMANA5IMP & "IMPREL " & zgrp & " " & zRPT & " " & zUSER & " " & cTIPO & " " & ZMANA5IMP
    'Shell eRUN, vbNormalFocus
    ShellEx "IMPREL", essSW_SHOWDEFAULT, zgrp & " " & zRPT & " " & zUSER & " " & cTIPO, ZMANA5IMP, , Me.hWnd
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Dim sButton          As String
    sButton = Button
    sButton = Left(UCase(Replace(sButton, "&", "")), 3)
    If Not AcessaBtnOld("escMANREL", Button.Index) Then
        Exit Sub
    End If
    ''    TbrLog Me, Button

    Select Case sButton

    Case "IMP"
        imprima_click

    Case "LOC"
        LocalizaGri1 Grid

    Case "CON"
        FrmPrintSetup.Show vbModal, Me

    Case "SAI"
        CmdSair_Click

    End Select

End Sub


