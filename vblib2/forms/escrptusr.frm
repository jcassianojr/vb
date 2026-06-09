VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{F22668DE-E08D-467B-8E41-13900013BD5F}#2.7#0"; "VBextra2.OCX"
Object = "{379157C5-E9BD-43F1-9F83-B037496BED42}#1.3#0"; "VBCCR18.OCX"
Object = "{075212A8-C1CF-444E-939D-F6046CCDBC08}#1.5#0"; "VBFLXGRD18.OCX"
Begin VB.Form escrptusr 
   Caption         =   "Liberação Usuario"
   ClientHeight    =   6024
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   8952
   Icon            =   "escrptusr.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6024
   ScaleWidth      =   8952
   StartUpPosition =   2  'CenterScreen
   Begin VBFLXGRD18.VBFlexGrid gridrptusr 
      Height          =   5292
      Left            =   5280
      TabIndex        =   7
      Top             =   600
      Width           =   3492
      _ExtentX        =   6160
      _ExtentY        =   9335
   End
   Begin VBFLXGRD18.VBFlexGrid gridusr 
      Height          =   5175
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   3495
      _ExtentX        =   6160
      _ExtentY        =   9123
   End
   Begin vbExtra.FlexFn FlexFn1 
      Height          =   396
      Left            =   3720
      TabIndex        =   5
      Top             =   120
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   699
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
   Begin XPControls.XPButton incusrrpt 
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   720
      Width           =   1455
      _ExtentX        =   2561
      _ExtentY        =   656
      Picture         =   "escrptusr.frx":038A
      Caption         =   "incluir"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton excrptusr 
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   1200
      Width           =   1455
      _ExtentX        =   2561
      _ExtentY        =   656
      Picture         =   "escrptusr.frx":0724
      Caption         =   "excluir"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton edirptusr 
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   1680
      Width           =   1455
      _ExtentX        =   2561
      _ExtentY        =   656
      Picture         =   "escrptusr.frx":0ABE
      Caption         =   "direitos"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VBCCR18.CommandButtonW cmdsair 
      Height          =   492
      Left            =   3840
      TabIndex        =   8
      Top             =   2280
      Width           =   1212
      _ExtentX        =   2138
      _ExtentY        =   868
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ImageListAlignment=   1
      Caption         =   "Retornar"
      Alignment       =   0
      VerticalAlignment=   0
      Picture         =   "escrptusr.frx":1058
      PictureAndCaption=   -1  'True
      WordWrap        =   0   'False
      Style           =   1
   End
   Begin VB.Label Label2 
      Caption         =   "Usuario Liberados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   1
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Usuario Disponiveis"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "escrptusr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdSair_Click()
  Unload Me
End Sub

Private Sub edirptusr_Click()
  gridrptusr.Col = 0
  zIDRPTUSR = FixInt(gridrptusr, 0)
  gridrptusr.Col = 1
  ePASS03 = gridrptusr
  frmRPTUSR.Show vbModal, Me
End Sub

Private Sub encerrar_Click()

End Sub

Private Sub excrptusr_Click()
  Dim sSQL As String
  gridrptusr.Col = 0
  zIDRPTUSR = FixInt(gridrptusr)
  sSQL = "select * from " & ArqRPTUsr() & " WHERE GRP='" & zgrp & "' AND RPT='" & zRPT & "' AND IDUSUARIO=" & zIDRPTUSR
  If ApagaSQLP(DBWRPT, sSQL) Then
    filrptusr
  End If
End Sub

Private Sub filrptusr()
  Dim sSQL As String
  Dim cNOME As String

  cNOME = ArqRPTUsr


  sSQL = "select " & cNOME & ".IDUSUARIO,USUARIO.USUARIO "
  sSQL = sSQL & "from " & cNOME & ",USUARIO WHERE GRP='" & zgrp & "' AND RPT='" & zRPT & "'"
  sSQL = sSQL & " and " & cNOME & ".IDUSUARIO = USUARIO.IDUSUARIO"
  sSQL = sSQL & " Order By USUARIO.USUARIO  "

  MontaGridFast gridrptusr, 2, Array(500, 2000), Array("ID", "Usuario"), _
                Array("IDUSUARIO", "USUARIO"), DBWRPT, sSQL


End Sub

Private Sub filusr()
  Dim cARQ As String
  Dim cSQL As String

  cARQ = DBWRPT
  cSQL = "SELECT IDUSUARIO,USUARIO FROM USUARIO ORDER BY USUARIO"

  MontaGridUltra gridusr, 2, Array(500, 2000), Array("ID", "Nome"), _
                 Array("IDUSUARIO", "L$USUARIO"), cARQ, cSQL


End Sub

Private Sub Form_Load()
  CenterFormToScreen Me
  filusr
  filrptusr

End Sub

Private Sub gridrptusr_KeyPress(KeyAscii As Integer)
  If KeyAscii > 31 And KeyAscii < 123 Then
    LocalizaGrid gridrptusr, Chr(KeyAscii), 1, False
  End If

End Sub

'Private Sub gridrptusr_SelChange()
'    With gridrptusr
'        If .Rows > 2 Then
'            .Col = .Cols - 1
'            .ColSel = 0
'            .TopRow = .Row
'       End If
'   End With
'End Sub

Private Sub gridusr_KeyPress(KeyAscii As Integer)
  If KeyAscii > 31 And KeyAscii < 123 Then
    LocalizaGrid gridusr, Chr(KeyAscii), 1, False
  End If

End Sub

'Private Sub gridusr_SelChange()
'    With gridusr
'        If .Rows > 2 Then
'            .Col = .Cols - 1
'            .ColSel = 0
'            .TopRow = .Row
'        End If
'    End With
'End Sub

Private Sub incusrrpt_Click()
  Dim sSQL As String
  gridusr.Col = 0
  zIDRPTUSR = FixInt(gridusr)
  sSQL = "select * from " & ArqRPTUsr() & " WHERE GRP='" & zgrp & "' AND RPT='" & zRPT & "' AND IDUSUARIO=" & zIDRPTUSR
  IncluiSQL DBWRPT, sSQL, 4, Array("GRP", "RPT", "IDUSUARIO", "VISUALIZA"), _
            Array(zgrp, zRPT, zIDRPTUSR, True), True, False
  gridusr.Col = 1
  ePASS03 = gridusr
  frmRPTUSR.Show vbModal, Me
  filrptusr
End Sub


