VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Begin VB.Form frmMe04 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Me04"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11775
   Icon            =   "frmMe04.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Me04"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   11775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton escCodInt 
      Height          =   315
      Left            =   11160
      Picture         =   "frmMe04.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   82
      TabStop         =   0   'False
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox txt 
      Height          =   330
      Index           =   31
      Left            =   7320
      TabIndex        =   81
      Top             =   840
      Width           =   3735
   End
   Begin VB.CommandButton EscPFx 
      Height          =   315
      Left            =   10800
      Picture         =   "frmMe04.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   79
      TabStop         =   0   'False
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton EscTipIns 
      Height          =   315
      Index           =   2
      Left            =   5820
      Picture         =   "frmMe04.frx":109E
      Style           =   1  'Graphical
      TabIndex        =   78
      TabStop         =   0   'False
      Top             =   1320
      Width           =   375
   End
   Begin VB.CommandButton EscTipIns 
      Height          =   315
      Index           =   1
      Left            =   11040
      Picture         =   "frmMe04.frx":1628
      Style           =   1  'Graphical
      TabIndex        =   77
      TabStop         =   0   'False
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton CmdNorma 
      Caption         =   "CC"
      Height          =   315
      Index           =   1
      Left            =   10140
      TabIndex        =   76
      TabStop         =   0   'False
      ToolTipText     =   "Conforme Certificado"
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton CmdNorma 
      Caption         =   "CD"
      Height          =   315
      Index           =   0
      Left            =   9720
      TabIndex        =   75
      TabStop         =   0   'False
      ToolTipText     =   "Conforme Desenho"
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton cmdgrvdata 
      Caption         =   "Gravar"
      Height          =   315
      Left            =   10680
      TabIndex        =   74
      TabStop         =   0   'False
      Top             =   4380
      Width           =   615
   End
   Begin VB.CommandButton Cmd 
      Caption         =   "E."
      Height          =   315
      Left            =   10260
      TabIndex        =   73
      TabStop         =   0   'False
      Top             =   4380
      Width           =   375
   End
   Begin VB.CommandButton CmdCalib 
      Caption         =   "Externa"
      Height          =   255
      Index           =   2
      Left            =   3480
      TabIndex        =   72
      TabStop         =   0   'False
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton CmdCalib 
      Caption         =   "Cliente"
      Height          =   255
      Index           =   1
      Left            =   4440
      TabIndex        =   71
      TabStop         =   0   'False
      Top             =   4140
      Width           =   855
   End
   Begin VB.CommandButton CmdCalib 
      Caption         =   "Interna"
      Height          =   315
      Index           =   0
      Left            =   3480
      TabIndex        =   70
      TabStop         =   0   'False
      Top             =   4080
      Width           =   855
   End
   Begin VB.CommandButton CmdSitu 
      Caption         =   "Manutencao"
      Height          =   255
      Index           =   4
      Left            =   5220
      TabIndex        =   69
      TabStop         =   0   'False
      Top             =   3660
      Width           =   1275
   End
   Begin VB.CommandButton CmdSitu 
      Caption         =   "Inativo"
      Height          =   255
      Index           =   3
      Left            =   4260
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   3660
      Width           =   915
   End
   Begin VB.CommandButton CmdSitu 
      Caption         =   "Estoque"
      Height          =   255
      Index           =   2
      Left            =   3300
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   3660
      Width           =   915
   End
   Begin VB.CommandButton CmdSitu 
      Caption         =   "Calibrando"
      Height          =   255
      Index           =   1
      Left            =   2340
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   3660
      Width           =   915
   End
   Begin VB.CommandButton CmdSitu 
      Caption         =   "Ativo"
      Height          =   255
      Index           =   0
      Left            =   1380
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   3660
      Width           =   915
   End
   Begin VB.CommandButton Command6 
      Caption         =   "-->"
      Height          =   315
      Left            =   1380
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   2940
      Width           =   435
   End
   Begin VB.CommandButton Command7 
      Height          =   315
      Left            =   960
      Picture         =   "frmMe04.frx":1BB2
      Style           =   1  'Graphical
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   2940
      Width           =   435
   End
   Begin VB.CommandButton CmdCadTip 
      Caption         =   "Instrumentos"
      Height          =   315
      Index           =   1
      Left            =   2040
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton CmdCadTip 
      Caption         =   "Calibradores"
      Height          =   315
      Index           =   0
      Left            =   2040
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton EscTipIns 
      Height          =   315
      Index           =   0
      Left            =   1500
      Picture         =   "frmMe04.frx":213C
      Style           =   1  'Graphical
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox txtcodigo 
      BackColor       =   &H00C0FFFF&
      Height          =   330
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00C0FFFF&
      Height          =   330
      Index           =   0
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   360
      Width           =   390
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   1
      Left            =   3240
      TabIndex        =   0
      Top             =   360
      Width           =   2910
   End
   Begin VB.TextBox txt 
      Height          =   330
      Index           =   2
      Left            =   6180
      TabIndex        =   1
      Top             =   360
      Width           =   4110
   End
   Begin VB.TextBox txt 
      Height          =   330
      Index           =   3
      Left            =   10440
      TabIndex        =   2
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00C0FFFF&
      Height          =   330
      Index           =   4
      Left            =   540
      Locked          =   -1  'True
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   840
      Width           =   870
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00C0FFFF&
      Height          =   330
      Index           =   5
      Left            =   1980
      Locked          =   -1  'True
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   840
      Width           =   4110
   End
   Begin VB.TextBox txt 
      Height          =   330
      Index           =   6
      Left            =   6960
      TabIndex        =   3
      Top             =   1200
      Width           =   4110
   End
   Begin VB.TextBox txt 
      Height          =   330
      Index           =   7
      Left            =   1680
      TabIndex        =   4
      Top             =   1260
      Width           =   4110
   End
   Begin VB.TextBox txt 
      Height          =   330
      Index           =   8
      Left            =   6960
      TabIndex        =   5
      Top             =   1560
      Width           =   2670
   End
   Begin VB.TextBox txt 
      Height          =   330
      Index           =   9
      Left            =   1680
      TabIndex        =   6
      Top             =   1680
      Width           =   4110
   End
   Begin VB.TextBox txt 
      Height          =   330
      Index           =   10
      Left            =   6960
      TabIndex        =   7
      Top             =   1920
      Width           =   4050
   End
   Begin VB.TextBox txt 
      Height          =   330
      Index           =   11
      Left            =   60
      TabIndex        =   8
      Top             =   2580
      Width           =   3450
   End
   Begin VB.TextBox txt 
      Height          =   330
      Index           =   12
      Left            =   3540
      TabIndex        =   9
      Top             =   2580
      Width           =   3330
   End
   Begin VB.TextBox txt 
      Height          =   330
      Index           =   13
      Left            =   7020
      TabIndex        =   10
      Top             =   2580
      Width           =   3330
   End
   Begin VB.TextBox txt 
      Height          =   330
      Index           =   14
      Left            =   120
      TabIndex        =   11
      Top             =   3240
      Width           =   690
   End
   Begin VB.TextBox txt 
      Height          =   330
      Index           =   15
      Left            =   900
      TabIndex        =   12
      Top             =   3240
      Width           =   4110
   End
   Begin VB.TextBox txt 
      Height          =   330
      Index           =   17
      Left            =   6480
      TabIndex        =   14
      Top             =   3240
      Width           =   1230
   End
   Begin VB.TextBox txt 
      Height          =   330
      Index           =   18
      Left            =   7800
      TabIndex        =   15
      Top             =   3240
      Width           =   2130
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00C0FFFF&
      Height          =   330
      Index           =   19
      Left            =   900
      Locked          =   -1  'True
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   3600
      Width           =   390
   End
   Begin VB.TextBox txt 
      Height          =   330
      Index           =   20
      Left            =   120
      TabIndex        =   16
      Top             =   4320
      Width           =   1350
   End
   Begin VB.TextBox txt 
      Height          =   330
      Index           =   21
      Left            =   1560
      TabIndex        =   17
      Top             =   4320
      Width           =   1290
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00C0FFFF&
      Height          =   330
      Index           =   22
      Left            =   2940
      Locked          =   -1  'True
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   4320
      Width           =   450
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00C0FFFF&
      Height          =   330
      Index           =   23
      Left            =   5340
      Locked          =   -1  'True
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   4380
      Width           =   690
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00C0FFFF&
      Height          =   330
      Index           =   24
      Left            =   6060
      Locked          =   -1  'True
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   4380
      Width           =   1350
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00C0FFFF&
      Height          =   330
      Index           =   25
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   4380
      Width           =   1350
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00C0FFFF&
      Height          =   330
      Index           =   26
      Left            =   8820
      Locked          =   -1  'True
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   4380
      Width           =   1410
   End
   Begin VB.TextBox txt 
      Height          =   330
      Index           =   27
      Left            =   660
      TabIndex        =   18
      Top             =   4800
      Width           =   5910
   End
   Begin VB.TextBox txt 
      Height          =   330
      Index           =   28
      Left            =   660
      TabIndex        =   19
      Top             =   5160
      Width           =   5910
   End
   Begin VB.TextBox txt 
      Height          =   330
      Index           =   29
      Left            =   660
      TabIndex        =   20
      Top             =   5580
      Width           =   5910
   End
   Begin VB.TextBox txt 
      Height          =   330
      Index           =   16
      Left            =   5100
      TabIndex        =   13
      Top             =   3240
      Width           =   1290
   End
   Begin VB.TextBox txt 
      Height          =   330
      Index           =   30
      Left            =   8040
      TabIndex        =   21
      Top             =   4740
      Width           =   1410
   End
   Begin XPControls.XPButton Encerrar 
      Height          =   435
      Left            =   10080
      TabIndex        =   83
      Top             =   3600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "frmMe04.frx":26C6
      Caption         =   "Retornar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton cmdClose 
      Height          =   435
      Left            =   10080
      TabIndex        =   84
      Top             =   3000
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "frmMe04.frx":2C60
      Caption         =   "Salvar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton cmdFOTO 
      Height          =   435
      Left            =   6720
      TabIndex        =   85
      Top             =   5400
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "frmMe04.frx":31FA
      Caption         =   "Imagem"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CodigoInterno:"
      Height          =   195
      Index           =   6
      Left            =   6240
      TabIndex        =   80
      Top             =   840
      Width           =   1035
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo"
      Height          =   195
      Index           =   8
      Left            =   120
      TabIndex        =   59
      Top             =   900
      Width           =   315
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PF"
      Height          =   195
      Index           =   37
      Left            =   10440
      TabIndex        =   58
      Top             =   120
      Width           =   195
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   32
      Top             =   105
      Width           =   630
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo:"
      Height          =   195
      Index           =   1
      Left            =   1560
      TabIndex        =   33
      Top             =   120
      Width           =   630
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Marca:"
      Height          =   195
      Index           =   2
      Left            =   6240
      TabIndex        =   34
      Top             =   1920
      Width           =   630
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Capacidade"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   35
      Top             =   2340
      Width           =   855
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Divisão"
      Height          =   195
      Index           =   4
      Left            =   3540
      TabIndex        =   36
      Top             =   2340
      Width           =   525
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição"
      Height          =   195
      Index           =   5
      Left            =   3240
      TabIndex        =   37
      Top             =   120
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fornecedor"
      Height          =   195
      Index           =   7
      Left            =   60
      TabIndex        =   38
      Top             =   3000
      Width           =   810
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Compra:"
      Height          =   195
      Index           =   9
      Left            =   5100
      TabIndex        =   39
      Top             =   3000
      Width           =   630
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor:"
      Height          =   195
      Index           =   10
      Left            =   6540
      TabIndex        =   40
      Top             =   3000
      Width           =   630
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Situacao:"
      Height          =   195
      Index           =   11
      Left            =   120
      TabIndex        =   41
      Top             =   3660
      Width           =   630
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data Inicial"
      Height          =   195
      Index           =   12
      Left            =   180
      TabIndex        =   42
      Top             =   4080
      Width           =   795
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data Final"
      Height          =   195
      Index           =   13
      Left            =   1620
      TabIndex        =   43
      Top             =   4080
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Norma:"
      Height          =   195
      Index           =   14
      Left            =   6240
      TabIndex        =   44
      Top             =   1560
      Width           =   630
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Aplicação Produto"
      Height          =   195
      Index           =   15
      Left            =   6300
      TabIndex        =   45
      Top             =   120
      Width           =   1305
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Extensao"
      Height          =   195
      Index           =   16
      Left            =   8880
      TabIndex        =   46
      Top             =   4140
      Width           =   660
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Calibrar:"
      Height          =   195
      Index           =   17
      Left            =   2880
      TabIndex        =   47
      Top             =   4080
      Width           =   630
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ativo:"
      Height          =   195
      Index           =   18
      Left            =   7860
      TabIndex        =   48
      Top             =   3000
      Width           =   630
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Erroadm:"
      Height          =   195
      Index           =   19
      Left            =   7020
      TabIndex        =   49
      Top             =   2340
      Width           =   630
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Obs:"
      Height          =   195
      Index           =   20
      Left            =   120
      TabIndex        =   50
      Top             =   4740
      Width           =   330
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Modelo:"
      Height          =   195
      Index           =   23
      Left            =   900
      TabIndex        =   51
      Top             =   1740
      Width           =   630
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Preço Calibracao"
      Height          =   195
      Index           =   30
      Left            =   6660
      TabIndex        =   52
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Proxima"
      Height          =   195
      Index           =   31
      Left            =   7440
      TabIndex        =   53
      Top             =   4140
      Width           =   555
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ultima"
      Height          =   195
      Index           =   33
      Left            =   6120
      TabIndex        =   54
      Top             =   4140
      Width           =   435
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Setor Area Aplicacao"
      Height          =   195
      Index           =   34
      Left            =   60
      TabIndex        =   55
      Top             =   1320
      Width           =   1500
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Meses"
      Height          =   195
      Index           =   35
      Left            =   5400
      TabIndex        =   56
      Top             =   4140
      Width           =   465
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desenho:"
      Height          =   195
      Index           =   36
      Left            =   6240
      TabIndex        =   57
      Top             =   1200
      Width           =   630
   End
End
Attribute VB_Name = "frmMe04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cARQ As String
Dim cSQL As String
Dim aVAL As Variant
Dim aFOR As Variant
Dim aCAM As Variant
Dim aPAD As Variant
Dim nCAMPOS As Integer
Dim iLOOP As Integer

Private Sub Cmd_Click()
    cmdgrvdata.Visible = True
    TXT(23).Locked = False
    TXT(24).Locked = False
    TXT(26).Locked = False
    Cmd.Visible = False
End Sub

Private Sub CmdCalib_Click(Index As Integer)
    Select Case Index
    Case 0
        TXT(22).tEXT = "I"
    Case 1
        TXT(22).tEXT = "E"
    Case 2
        TXT(22).tEXT = "C"
    End Select

End Sub

Private Sub cmdClose_Click()
    On Error Resume Next
    If MDG("Gravar alteraçôes") Then
        For iLOOP = 0 To nCAMPOS - 1             ''tira o campos  nao gravar o numero,nome
            aVAL(iLOOP) = TXT(iLOOP)             ''vb matriz zero
        Next iLOOP
        GrvSQL cARQ, cSQL, nCAMPOS, aCAM, aVAL, aFOR
    End If
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub CmdCadTip_Click(Index As Integer)
    Select Case Index
    Case 0
        TXT(0).tEXT = "C"
    Case 1
        TXT(0).tEXT = "I"
      
    End Select
End Sub

Private Sub cmdFOTO_Click()
    Dim cSQL As String
    zgrp = FixStr(txtcodigo.tEXT)
    If Len(zgrp) = 0 Then
        Alert "Codigo nao preenchido"
        Exit Sub
    End If
    iImage = 5
    cARQRTF = PegPath("PATH", "IMGME04")
    cSQL = "select codigo from IMAGENS WHERE CODIGO='" & zgrp & "'"
    IncluiSQL cARQRTF, cSQL, 1, Array("CODIGO"), Array(zgrp), True, False
    Load frmIMAGENS
    frmIMAGENS.txtFields(0).Enabled = False
    frmIMAGENS.Escolher(0).Visible = False
    frmIMAGENS.Show vbModal, Me
End Sub

Private Sub cmdgrvdata_Click()
    cmdgrvdata.Visible = False
    TXT(23).Locked = True
    TXT(24).Locked = True
    TXT(26).Locked = True
    If IsDate(TXT(24)) And TXT(23) > 0 Then
        TXT(25).tEXT = DateAdd("m", TXT(23), TXT(24))
    End If
    Cmd.Visible = True
End Sub

Private Sub CmdNorma_Click(Index As Integer)
    Select Case Index
    Case 0
        TXT(8).tEXT = "CONFORME DESENHO"
    Case 1
        TXT(8).tEXT = "CONFORME CERTIFICADO"
    End Select

End Sub

Private Sub Cmdsitu_Click(Index As Integer)
    Select Case Index
    Case 0
        TXT(19).tEXT = "A"
    Case 1
        TXT(19).tEXT = "C"
    Case 2
        TXT(19).tEXT = "E"
    Case 3
        TXT(19).tEXT = "I"
    Case 4
        TXT(19).tEXT = "M"
    End Select
End Sub

Private Sub Command6_Click()
    Dim cARQ As String
    Dim aRETU As Variant
    Dim sSQL             As String
    Dim nNUMERO          As Long
    nNUMERO = FixInt(TXT(14), 0)
    cARQ = GeraConn(zMANA5EMP, "JETFOX")
    sSQL = "SELECT NOME FROM MB01 WHERE NUMERO=" & nNUMERO
    aRETU = PegSQL(cARQ, sSQL, 1, Array("NOME"), Array("C"), Array(""))
    If lRETU Then
        TXT(15) = aRETU(0)
    End If
End Sub

Private Sub Encerrar_Click()
    If Not MDG("Sair sem gravar") Then
        Exit Sub
    End If
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub Command7_Click()
    lRETU = False
    eRETU01 = ""
    eRETU02 = ""
    cARQESC = "MB01"
    escNUMNOM.Show vbModal, Me
    If lRETU Then
        TXT(14) = eRETU01
        TXT(15) = eRETU02
    End If
End Sub

Private Sub escCodInt_Click()
    iMU01 = 12
    escmp.Show vbModal
    If lRETU Then
        TXT(31) = eRETU01
    End If
End Sub

Private Sub EscPFx_Click()
    escpf.Show vbModal
    If lRETU Then
        TXT(3) = eRETU01
    End If
End Sub

Private Sub EscTipIns_Click(Index As Integer)
    Select Case Index
    Case 0
        iMD02 = 1
    Case 1
        iMD02 = 6
    Case 2
        iMD02 = 7
    End Select
    escMD02.Show vbModal, Me
    If lRETU Then
        Select Case Index
        Case 0
            TXT(4).tEXT = eRETU01
            TXT(5).tEXT = eRETU02
        Case 1
            TXT(10).tEXT = eRETU01
        Case 2
            TXT(7).tEXT = eRETU02
                
        End Select
       
      
    End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    TeclaEnter KeyCode

End Sub

Private Sub Form_Load()
    CenterFormToScreen Me
    cmdgrvdata.Visible = False
    cARQ = PegPath("PATH", "MANA5INS")
    cARQ = GeraConn(cARQ, "SDECDX")
    cSQL = "select CODIGO, "
    cSQL = cSQL & " CADTIP , NOMTIPO, APLICACAO, PF, CODTIPO,"
    cSQL = cSQL & " TIPO , DESENHO, APLIC, NORMA, MODELO,"
    cSQL = cSQL & " MARCA, CAPACI, DIVI, ERROADM, CODFOR,"
    cSQL = cSQL & " COGFOR , COMPRA, Valor, ATIVO, SITUACAO,"
    cSQL = cSQL & " DATAUSO , DATAFIM, TIPCAL, CALIBRAR, CALULT,"
    cSQL = cSQL & " CALPRO , DATAEXT, OBS01, OBS02, OBS03, PRECO, CODIGOINT "
    cSQL = cSQL & " FROM ME04 WHERE CODIGO='" & Trim(ePASS01) & "'"
    txtcodigo.tEXT = ePASS01
    nCAMPOS = 32
    aCAM = Array("CADTIP", "NOMTIPO", "APLICACAO", "PF", "CODTIPO", _
                 "TIPO", "DESENHO", "APLIC", "NORMA", "MODELO", _
                 "MARCA", "CAPACI", "DIVI", "ERROADM", "CODFOR", _
                 "COGFOR", "COMPRA", "VALOR", "ATIVO", "SITUACAO", _
                 "DATAUSO", "DATAFIM", "TIPCAL", "CALIBRAR", "CALULT", _
                 "CALPRO", "DATAEXT", "OBS01", "OBS02", "OBS03", "PRECO", "CODIGOINT")
    aFOR = Array("C", "C", "C", "NI", "C", _
                 "C", "C", "C", "C", "C", _
                 "C", "C", "C", "C", "NI", _
                 "C", "DC", "N", "C", "C", _
                 "DC", "DC", "C", "NI", "DC", _
                 "DC", "DC", "C", "C", "C", "N", "C")
    aPAD = Array("", "", "", 0, "", _
                 "", "", "", "", "", _
                 "", "", "", "", 0, _
                 "", "", 0, "", "", _
                 "", "", "", 0, "", _
                 "", "", "", "", "", 0, "")
    aVAL = PegSQL(cARQ, cSQL, nCAMPOS, aCAM, aFOR, aPAD)
    For iLOOP = 0 To nCAMPOS - 1
        TXT(iLOOP) = aVAL(iLOOP)
    Next iLOOP
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault
End Sub

Private Sub txt_GotFocus(Index As Integer)
    FocusMe
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
    Case 3, 14, 23
        KeyAscii = ValiText(KeyAscii, "#NI")
    Case 17, 30
        KeyAscii = ValiText(KeyAscii, "#N")
    Case 16, 20, 21, 24, 25, 26
        KeyAscii = ValidaDados(KeyAscii, "DL")
    End Select
End Sub


