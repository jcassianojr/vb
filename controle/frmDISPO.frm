VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{EA478B61-D9EC-47F6-BB21-95A533AF2251}#1.0#0"; "TabExC01.ocx"
Begin VB.Form frmDISPO 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MENU"
   ClientHeight    =   7590
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   11760
   Icon            =   "frmDISPO.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   11760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtNome 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   330
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   600
      Width           =   9735
   End
   Begin VB.TextBox TxtCodigo 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   330
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   120
      Width           =   3375
   End
   Begin VB.TextBox TXT 
      Height          =   330
      Index           =   8
      Left            =   840
      TabIndex        =   24
      Top             =   1440
      Width           =   4350
   End
   Begin TabExCtl.SSTabEx SSTab1 
      Height          =   4815
      Left            =   120
      TabIndex        =   11
      Top             =   2400
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   8493
      Tabs            =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabHeight       =   520
      TabCaption(0)   =   "Colaborador Elabarador"
      Tab(0).ControlCount=   12
      Tab(0).Control(0)=   "Duplicar(13)"
      Tab(0).Control(1)=   "Duplicar(3)"
      Tab(0).Control(2)=   "TXT(7)"
      Tab(0).Control(3)=   "TXT(6)"
      Tab(0).Control(4)=   "TXT(5)"
      Tab(0).Control(5)=   "GridPla"
      Tab(0).Control(6)=   "Command7"
      Tab(0).Control(7)=   "Command8"
      Tab(0).Control(8)=   "Command4"
      Tab(0).Control(9)=   "Command9"
      Tab(0).Control(10)=   "Label12"
      Tab(0).Control(11)=   "Label10"
      TabCaption(1)   =   "Revisao"
      Tab(1).ControlCount=   10
      Tab(1).Control(0)=   "Command6"
      Tab(1).Control(1)=   "CmdMotRev"
      Tab(1).Control(2)=   "TXT(4)"
      Tab(1).Control(3)=   "TXT(3)"
      Tab(1).Control(4)=   "NovaRev"
      Tab(1).Control(5)=   "gridrev"
      Tab(1).Control(6)=   "gridrevi"
      Tab(1).Control(7)=   "ComMotInc"
      Tab(1).Control(8)=   "ComMotAlt"
      Tab(1).Control(9)=   "ComMotExc"
      Begin XPControls.XPButton Duplicar 
         Height          =   375
         Index           =   13
         Left            =   8160
         TabIndex        =   29
         Top             =   3480
         Width           =   1095
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "Reordenar"
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
      Begin XPControls.XPButton Duplicar 
         Height          =   375
         Index           =   3
         Left            =   8160
         TabIndex        =   28
         Top             =   3000
         Width           =   1095
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "Duplicar"
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
      Begin XPControls.XPButton Command6 
         Height          =   495
         Left            =   -67200
         TabIndex        =   22
         Top             =   1620
         Width           =   1335
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "Distribuicao"
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
      Begin XPControls.XPButton CmdMotRev 
         Height          =   375
         Left            =   -67200
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1140
         Width           =   1335
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "Motivos"
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
      Begin VB.TextBox TXT 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   330
         Index           =   4
         Left            =   -72240
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   540
         Width           =   1095
      End
      Begin VB.TextBox TXT 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   330
         Index           =   3
         Left            =   -73440
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   540
         Width           =   1095
      End
      Begin XPControls.XPButton NovaRev 
         Height          =   375
         Left            =   -74880
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   540
         Width           =   1215
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "Nova Revisão"
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
      Begin VB.TextBox TXT 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   330
         Index           =   7
         Left            =   7680
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox TXT 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   330
         Index           =   6
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   600
         Width           =   4095
      End
      Begin VB.TextBox TXT 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   330
         Index           =   5
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   600
         Width           =   1095
      End
      Begin MSFlexGridLib.MSFlexGrid gridrev 
         Height          =   1815
         Left            =   -74880
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1020
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   3201
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid gridrevi 
         Height          =   1575
         Left            =   -74880
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   2940
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   2778
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid GridPla 
         Height          =   2895
         Left            =   240
         TabIndex        =   27
         Top             =   1080
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   5106
         _Version        =   393216
      End
      Begin XPControls.XPButton ComMotInc 
         Height          =   435
         Left            =   -67200
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   3000
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmDISPO.frx":058A
         Caption         =   "Novo"
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
      Begin XPControls.XPButton ComMotAlt 
         Height          =   435
         Left            =   -67200
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   3480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmDISPO.frx":0B24
         Caption         =   "Editar"
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
      Begin XPControls.XPButton ComMotExc 
         Height          =   435
         Left            =   -67200
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   3960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmDISPO.frx":10BE
         Caption         =   "Apagar"
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
      Begin XPControls.XPButton Command7 
         Height          =   375
         Left            =   2520
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   600
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Picture         =   "frmDISPO.frx":1658
         Caption         =   ""
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
      Begin XPControls.XPButton Command8 
         Height          =   435
         Left            =   8160
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmDISPO.frx":1BF2
         Caption         =   "Novo"
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
      Begin XPControls.XPButton Command4 
         Height          =   435
         Left            =   8160
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   1560
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmDISPO.frx":218C
         Caption         =   "Editar"
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
      Begin XPControls.XPButton Command9 
         Height          =   435
         Left            =   8160
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmDISPO.frx":2726
         Caption         =   "Apagar"
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
      Begin VB.Label Label12 
         Caption         =   "Em"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   7200
         TabIndex        =   15
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label10 
         Caption         =   "Elaborador"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.TextBox SSQ 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   375
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "0"
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox SEQ 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   375
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "0"
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox PF 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   372
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   120
      Width           =   732
   End
   Begin VB.TextBox TXT 
      Height          =   330
      Index           =   2
      Left            =   840
      TabIndex        =   4
      Top             =   1800
      Width           =   5175
   End
   Begin VB.TextBox TXT 
      Height          =   375
      Index           =   1
      Left            =   840
      TabIndex        =   2
      Top             =   960
      Width           =   3615
   End
   Begin VB.TextBox TXT 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   330
      Index           =   0
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   1095
   End
   Begin XPControls.XPButton Encerrar 
      Height          =   435
      Left            =   10080
      TabIndex        =   30
      Top             =   960
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "frmDISPO.frx":2CC0
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
      TabIndex        =   31
      Top             =   240
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "frmDISPO.frx":325A
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
   Begin XPControls.XPButton ESCMS01A 
      Height          =   375
      Index           =   0
      Left            =   4560
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   960
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Picture         =   "frmDISPO.frx":37F4
      Caption         =   "MANA"
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
   Begin XPControls.XPButton ESCMS01A 
      Height          =   375
      Index           =   1
      Left            =   5640
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   960
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Picture         =   "frmDISPO.frx":3D8E
      Caption         =   "LX"
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
   Begin VB.Label Label2 
      Caption         =   "SSQ"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   4920
      TabIndex        =   9
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Seq"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   3480
      TabIndex        =   7
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "PF"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2040
      TabIndex        =   5
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo:"
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   630
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Numero:"
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   630
   End
End
Attribute VB_Name = "frmDISPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cARQ As String
Dim cARQPF As String
Dim cSQL As String
Dim aVAL As Variant
Dim aFOR As Variant
Dim aCAM As Variant
Dim aPAD As Variant
Dim nCAMPOS As Integer
Dim iLOOP As Integer

Private Sub cmdClose_Click()
    On Error Resume Next
    Dim I As Integer
    If MDG("Gravar alteraçôes") Then
        For iLOOP = 0 To nCAMPOS - 1
            aVAL(iLOOP) = TXT(iLOOP)
        Next iLOOP
        GrvSQL cARQ, cSQL, nCAMPOS, aCAM, aVAL, aFOR
    End If
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub CmdMotRev_Click()
    gridrev.Col = 1
    nREV = FixInt(gridrev)
    ComMotAlt.Enabled = True
    ComMotExc.Enabled = True
    ComMotInc.Enabled = True
    Filgridrevi
End Sub

Private Sub Command4_Click()
    If GridPla.Row > 0 Then
        GridPla.Col = 0
        nORD = GridPla
        frmDISPOI.Show vbModal, Me
        FilRelat
    End If
End Sub

Private Sub Command6_Click()
    If nREV = 0 Then
        Alert "Revisão Nao Escolhida"
        Exit Sub
    End If
    ePASS01 = "select * from rev WHERE PF=" & nPPAP & " AND TIPO='DIS' AND REVISAO=" & nREV
    FrmRev.Show vbModal, Me
End Sub

Private Sub Command7_Click()
    TXT(5) = zIDFOLHA
    TXT(6) = zNOMEFOLHA
    TXT(7) = Date
End Sub

Private Sub PegCodigoDescricaoPf()
    Dim sSQL, aRETU, sARQ
    sARQ = PegPath("PATH", "PF")
    sSQL = "SELECT CODIGO,DESCR FROM PF WHERE PF=" & nPF
    aRETU = PegSQL(sARQ, sSQL, 2, Array("CODIGO", "DESCR"), Array("C", "C"), Array("", ""))
    If lRETU Then
        TxtCodigo = aRETU(0)
        TxtNome = aRETU(1)
    End If
End Sub

Private Sub Command8_Click()
    nNOVAPOS = FixInt(PegUltGrid(GridPla, 0), 0) + 1
    iMU01 = 9
    NewOrd.Show vbModal, Me
    FilRelat
End Sub

Private Sub Command9_Click()
    If GridPla.Row > 0 Then
        GridPla.Col = 0
        nORD = FixInt(GridPla)
        If ApagaSQLP(cARQ, "select * from DISPOI WHERE NUMERO=" & nPPAP & " AND ITEM=" & nORD) Then
            FilRelat
        End If
    End If
End Sub

Private Sub ComMotAlt_Click()
    If nREV = 0 Then
        Alert "Revisão Nao Escolhida"
        Exit Sub
    End If
    If gridrevi.Row > 0 Then
        gridrevi.Col = 3
        nREVI = FixInt(gridrevi)
        lEDITAR = True
        ePASS01 = "select * from revi WHERE PF=" & nPPAP & " AND TIPO='DIS' AND REVISAO=" & nREV & " AND ITEM=" & nREVI
        frmREVI.Show vbModal, Me
        Filgridrevi
    End If
End Sub

Private Sub ComMotExc_Click()
    Dim cSQL As String
    If nREV = 0 Then
        Alert "Revisão Nao Escolhida"
        Exit Sub
    End If
    If gridrevi.Row > 0 Then
        gridrevi.Col = 3
        nREVI = FixInt(gridrevi)
        cSQL = "select * from revi WHERE PF=" & nPPAP & " AND TIPO='DIS' AND REVISAO=" & nREV & " AND  ITEM=" & nREVI
        If ApagaSQLP(cARQPF, cSQL) Then
            Filgridrevi
        End If
    End If
End Sub

Private Sub ComMotInc_Click()
    Dim nITEM As Long
    If nREV = 0 Then
        Alert "Revisão Nao Escolhida"
        Exit Sub
    End If
    nITEM = 1
    If gridrevi.Rows > 1 Then
        gridrevi.Row = gridrevi.Rows - 1         ''//Posicao comeca com zero
        gridrevi.Col = 3
        nITEM = FixInt(gridrevi) + 1
    End If
    IncluiSQL cARQPF, "SELECT * FROM REVI WHERE PF=" & nPF & " AND TIPO='DIS' AND REVISAO=" & nREV & " AND ITEM=" & nITEM, 4, Array("PF", "REVISAO", "TIPO", "ITEM"), _
        Array(nPPAP, nREV, "DIS", nITEM), True, False
    Filgridrevi
End Sub

Private Sub Duplicar_Click(Index As Integer)
    Dim aVALTMP As Variant
    Dim aCAMTMP As Variant
    Dim aFORTMP As Variant
    Dim aPADTMP As Variant
    Dim sSQL As String
    Dim nCAMPOS As Integer
    Dim cARQ As String
    Dim NINDEXREF As Integer

    cARQ = PegPath("PATH", "DISPO")

    NINDEXREF = Index
    If NINDEXREF > 10 Then
        NINDEXREF = Index - 10
    End If

   
    
    GridPla.Col = 0
    nORD = FixInt(GridPla)

    If nORD = 0 Then
        Alert ("Clique no Item de Origem Antes")
        Exit Sub
    End If


    nNOVAPOS = FixInt(PegUltGrid(GridPla, 0), 0) + 1
    
    If Index = 13 Then                           'reordenacao
        nNOVAPOS = FixInt(Val(Busca("No", "Renumerando", CStr(nNOVAPOS), 8)))
        If nNOVAPOS = 0 Then
            Alert ("No em Branco")
            Exit Sub
        End If
        LocalizaGrid GridPla, nNOVAPOS
        If lRETU Then
            Alert "ordem ja existe"
            Exit Sub
        End If
    End If
    
    
    
    sSQL = "select * from DISPOI WHERE NUMERO=" & nPPAP & " AND ITEM=" & nORD
   
                              
    nCAMPOS = 1
    aCAMTMP = Array("DESCRICAO")
    aPADTMP = Array("")
    aFORTMP = Array("C")
                        
    
    aVALTMP = PegSQL(cARQ, sSQL, nCAMPOS, aCAMTMP, aFORTMP, aPADTMP)
    
    If Index = 13 Then                           'reordenacao apaga
        ApagaSQL cARQ, sSQL
    End If
    
    eRETU01 = 0
    iMU01 = 9
    
    NewOrd.Show vbModal, Me
    
    nORD = eRETU01
    sSQL = "select * from DISPOI WHERE NUMERO=" & nPPAP & " AND ITEM=" & nORD
    
    GrvSQL cARQ, sSQL, nCAMPOS, aCAMTMP, aVALTMP, aFORTMP
    FilRelat
    
End Sub

Private Sub Encerrar_Click()
    If Not MDG("Sair sem gravar") Then
        Exit Sub
    End If
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub ESCMS01A_Click(Index As Integer)
    Select Case Index
    Case 0
        ePASS01 = " WHERE CODTIPO='DC'"
        escME04.Show vbModal
    Case 1
        iMU01 = 12
        escmp.Show vbModal
        If lRETU Then
            TXT(8) = eRETU01
        End If
    End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    TeclaEnter KeyCode
End Sub

Private Sub Form_Load()
    Dim I As Integer
    CenterFormToScreen Me
    PF.Text = nPF
    SEQ.Text = nSEQ
    SSQ.Text = nSSQ
    cARQPF = PegPath("PATH", "PF")
    cARQ = PegPath("PATH", "DISPO")
    cSQL = "select * from DISPO WHERE numero=" & nPPAP
    nCAMPOS = 9
    aCAM = Array("NUMERO", "CODIGO", "NOME", "REVISAO", "DATAREV", "ELANUM", "ELANOM", "ELADAT", "CODIGOINT")
    aFOR = Array("NI", "C", "C", "C", "DZ", "NI", "C", "DZ", "C")
    aPAD = Array(0, "", "", "", "", 0, "", "", "")
    aVAL = PegSQL(cARQ, cSQL, nCAMPOS, aCAM, aFOR, aPAD)
    For iLOOP = 0 To nCAMPOS - 1
        TXT(iLOOP) = aVAL(iLOOP)
    Next iLOOP
    If nPF > 0 Then
        PegCodigoDescricaoPf
    End If
    Filgridrev
    If gridrev.Rows > 1 Then
        gridrev.Row = gridrev.Rows - 1
        CmdMotRev_Click
    End If
    FilRelat
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault
End Sub

Private Sub NovaRev_Click()
    Dim nREV As Integer
    If Len(TXT(4)) > 0 Then
        If TXT(4) > Today() Then
            Alert ("Data Revisao Anterior Maior que Data do Sistema")
            Exit Sub
        End If
    End If
    If MDG("Fazer Nova Revisao") Then
        nREV = FixNum(TXT(3)) + 1
        IncluiSQL cARQPF, "SELECT * FROM REV WHERE PF=" & nPF & " AND TIPO='DIS' AND REVISAO=" & nREV, 4, Array("PF", "REVISAO", "TIPO", "DATA"), _
        Array(nPPAP, nREV, "DIS", Today()), False, False
        TXT(3) = nREV
        TXT(4) = Date
    End If
End Sub

Private Sub Filgridrev()
    Dim cSQL As String
    cSQL = "select PF,REVISAO,TIPO,DATA from rev WHERE PF=" & nPPAP & " AND TIPO='DIS' ORDER BY REVISAO"
    MontaGridFast gridrev, 4, Array(600, 400, 400, 1200), Array("PF", "REV", "Tipo", "DATA"), _
        Array("PF", "REVISAO", "TIPO", "C$DATA"), cARQPF, cSQL
End Sub

Private Sub Filgridrevi()
    Dim cSQL As String
    cSQL = "select * from revI WHERE PF=" & nPPAP & " AND REVISAO=" & nREV & " AND TIPO='DIS' ORDER BY ITEM"
    MontaGridFast gridrevi, 5, Array(600, 400, 400, 400, 2000), Array("PF", "REV", "Tipo", "ITEM", "Motivo"), _
        Array("PF", "REVISAO", "TIPO", "ITEM", "Motivo"), cARQPF, cSQL
End Sub

Private Sub FilRelat()
    On Error Resume Next
    Dim cSQL As String
    cSQL = "select ITEM,DESCRICAO from DISPOI  WHERE NUMERO=" & nPPAP & " ORDER BY ITEM"
    MontaGridFast GridPla, 2, Array(400, 4000), Array("Ord", "Descricao"), _
        Array("L$ITEM", "L$DESCRicao"), cARQ, cSQL
End Sub



