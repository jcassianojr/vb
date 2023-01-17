VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{7020C36F-09FC-41FE-B822-CDE6FBB321EB}#1.2#0"; "vbccr17.ocx"
Object = "{451B73A5-1563-45D5-A6AC-7B2B7D30B778}#1.0#0"; "BSPrin10.ocx"
Begin VB.Form FrmTxl 
   Caption         =   "Gerador Relatorios Texto"
   ClientHeight    =   7380
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10425
   Icon            =   "FrmTxl.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7380
   ScaleWidth      =   10425
   StartUpPosition =   3  'Windows Default
   Begin BSPrinter.PrintPreview PrintPreview1 
      Left            =   9240
      Top             =   2160
      _ExtentX        =   1191
      _ExtentY        =   1191
      LcK1b           =   "yefT59bnyufI583n1ufV59HnyefN58nn0+fK58nnzefR58rn1+fW5w=="
      LcK2b           =   "reeI58zn+ue059jn1Of656fni+eN543nleeM54HnlueR5w=="
      AmbientBb       =   $"FrmTxl.frx":058A
   End
   Begin VBCCR17.RichTextBox RichTextBox1 
      Height          =   255
      Left            =   9240
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   4200
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   450
      MultiLine       =   -1  'True
      TextRTF         =   "FrmTxl.frx":062A
   End
   Begin MSFlexGridLib.MSFlexGrid Listview1 
      Height          =   1335
      Left            =   120
      TabIndex        =   18
      Top             =   5880
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   2355
      _Version        =   393216
   End
   Begin VB.Frame Destino 
      Caption         =   "Destino"
      Height          =   1095
      Left            =   120
      TabIndex        =   6
      Top             =   4680
      Width           =   9735
      Begin VB.TextBox TxtCaminho 
         Height          =   285
         Left            =   2040
         TabIndex        =   15
         Text            =   "TxtCaminho"
         Top             =   720
         Width           =   7575
      End
      Begin VB.OptionButton OptDestino 
         Caption         =   "&Rede Caminho"
         Height          =   375
         Index           =   7
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   1575
      End
      Begin VB.OptionButton OptDestino 
         Caption         =   "&Printer Objeto"
         Height          =   375
         Index           =   6
         Left            =   7320
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton OptDestino 
         Caption         =   "&Arquivo(Acima)"
         Height          =   375
         Index           =   5
         Left            =   5640
         TabIndex        =   12
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton OptDestino 
         Caption         =   "C&OM2"
         Height          =   375
         Index           =   4
         Left            =   4560
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton OptDestino 
         Caption         =   "&COM1"
         Height          =   375
         Index           =   3
         Left            =   3360
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton OptDestino 
         Caption         =   "LPT&3"
         Height          =   375
         Index           =   2
         Left            =   2280
         TabIndex        =   9
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton OptDestino 
         Caption         =   "LTP&2"
         Height          =   375
         Index           =   1
         Left            =   1200
         TabIndex        =   8
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton OptDestino 
         Caption         =   "LPT&1"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.TextBox TxtArquivo 
      Height          =   285
      Left            =   960
      TabIndex        =   5
      Text            =   "TxtArquivo"
      Top             =   1680
      Width           =   9255
   End
   Begin VB.CheckBox CmdEmail 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Enviar Email"
      Height          =   195
      Left            =   1920
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Height          =   915
      Left            =   4320
      TabIndex        =   4
      Top             =   720
      Width           =   5985
      Begin XPControls.XPButton cmdimp 
         Height          =   555
         Left            =   1680
         TabIndex        =   20
         Top             =   240
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   979
         Picture         =   "FrmTxl.frx":0790
         Caption         =   "Imprimir"
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
      Begin XPControls.XPButton CmdVisua 
         Height          =   555
         Left            =   3120
         TabIndex        =   21
         Top             =   240
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   979
         Picture         =   "FrmTxl.frx":0D2A
         Caption         =   "Visualizar"
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
      Begin XPControls.XPButton CmdConfImp 
         Height          =   555
         Left            =   120
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   979
         Picture         =   "FrmTxl.frx":12C4
         Caption         =   "Configurar Impressora"
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
      Begin XPControls.XPButton Encerrar 
         Height          =   555
         Left            =   4440
         TabIndex        =   29
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   979
         Picture         =   "FrmTxl.frx":185E
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
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Exportar/Salvar Como"
      Height          =   915
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   4095
      Begin XPControls.XPButton Salvar 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   360
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   661
         Picture         =   "FrmTxl.frx":1DF8
         Caption         =   "TXT"
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
      Begin XPControls.XPButton Salvar 
         Height          =   375
         Index           =   1
         Left            =   1080
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   360
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         Picture         =   "FrmTxl.frx":2192
         Caption         =   "Html"
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
      Begin XPControls.XPButton Salvar 
         Height          =   375
         Index           =   2
         Left            =   2160
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   360
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   661
         Picture         =   "FrmTxl.frx":272C
         Caption         =   "PDF"
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
      Begin XPControls.XPButton Salvar 
         Height          =   375
         Index           =   3
         Left            =   3120
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   360
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   661
         Picture         =   "FrmTxl.frx":2AC6
         Caption         =   "RTF"
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
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
      Height          =   675
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   10155
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "Label1 - Nome do Relatorio que Estamos Imprimindo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Top             =   240
         Width           =   8775
      End
   End
   Begin XPControls.XPButton CmdFiltro 
      Height          =   435
      Left            =   8760
      TabIndex        =   22
      Top             =   2760
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   767
      Picture         =   "FrmTxl.frx":2E60
      Caption         =   "filtrar"
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
   Begin XPControls.XPButton CmdOrdem 
      Height          =   435
      Left            =   8760
      TabIndex        =   23
      Top             =   3600
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   767
      Picture         =   "FrmTxl.frx":33FA
      Caption         =   "ordernar"
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
   Begin XPControls.XPButton CmdEditar 
      Height          =   315
      Left            =   960
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   2040
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      Picture         =   "FrmTxl.frx":3994
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
   Begin XPControls.XPButton CmdAbrirCom 
      Height          =   315
      Left            =   2040
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   2040
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Picture         =   "FrmTxl.frx":3F2E
      Caption         =   "Abrir Com"
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
   Begin XPControls.XPButton CmdShell 
      Height          =   315
      Left            =   3480
      TabIndex        =   32
      Top             =   2040
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   556
      Caption         =   "S"
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
      Caption         =   "Arquivo"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label filtro 
      Caption         =   "filtro"
      Height          =   735
      Left            =   960
      TabIndex        =   17
      Top             =   2760
      Width           =   7575
      WordWrap        =   -1  'True
   End
   Begin VB.Label Lblsql 
      Caption         =   "sql"
      Height          =   735
      Left            =   960
      TabIndex        =   16
      Top             =   3720
      Width           =   7575
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "FrmTxl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const nFORMID = 1181
Const cFORMID = "Imprimir Relatorios Script de Texto"
Dim oTXT As New Criatxt
Dim oIni As New cIniSettings
Dim nQTDELAY As Long
Dim nLINHAS As Long
Dim nCOLUNAS As Long
Dim cDESTINO As String
Dim aLAY As Variant
Dim oCONN As New ADODB.Connection
Dim oRS As New ADODB.Recordset
Dim cSQL As String
Dim cARQ As String
Dim cSETUP As String
Dim lARQTXT As Boolean

Private Sub gerar(ByVal cOPE As String)
    Dim X As Long
    Dim CLINHA As String
    Dim aUSO As Variant
    Dim aREG As Variant
    Dim cVAL As String
    'Dim oREG As Variant
    Dim cREG As String
    CLINHA = ""
  
    oTXT.nCOL = nCOLUNAS
    oTXT.nROW = nLINHAS
    oTXT.Destino = cDESTINO
    If cOPE = "A" Or cOPE = "P" Then
        oTXT.Destino = "ARQ"
        oTXT.arquivo = TxtArquivo
    End If
    If cDESTINO = "ARQ" Then
        oTXT.arquivo = TxtArquivo
    End If
    If cDESTINO = "IMPNET" Then
        oTXT.caminho = TxtCaminho
    End If
    If cSETUP <> "" Then
        oTXT.impsetup cSETUP
    End If
    oTXT.ABRIR
    cARQ = GeracArq(cARQ, , False)
    oCONN.ConnectionTimeout = 120
    oCONN.Open cARQ
    oRS.Open cSQL, oCONN, adOpenForwardOnly, adLockReadOnly
    While Not oRS.EOF
        For X = 1 To nQTDELAY
            CLINHA = aLAY(X)
            Select Case CLINHA
            Case "[NF]"
                oTXT.novafolha
            Case "[GF]"
                oTXT.gravafolha
            Case "{MN}"
                If Not oRS.EOF Then
                    oRS.MoveNext
                End If
            Case ""                              ''EM BRANCO NADA FAZ
            Case Else
                aUSO = Split(CLINHA, "|")
                If Left(aUSO(2), 1) = "{" Then
                    If Not oRS.EOF Then
                        aUSO(2) = Mid(aUSO(2), 2, Len(aUSO(2)) - 2)
                        aREG = Split(aUSO(2), "_")
                        
                        cVAL = aREG(2)
                        cREG = FixStr(oRS(aREG(1)))
                        aUSO(2) = FVar(cREG, cVAL)
                    Else
                        aUSO(2) = ""
                    End If
                End If
                oTXT.gravalincon aUSO(2), aUSO(0), aUSO(1)
            End Select
        Next X
    Wend
    oTXT.fechar
    oRS.Close
    oCONN.Close
End Sub

Private Sub CmdAbrirCom_Click()
cARQRTF = TxtArquivo.tEXT
   If FileExist(cARQRTF, True) Then
       Call OpenWith(cARQRTF, OAIF_ALLOW_REGISTRATION Or OAIF_EXEC Or OAIF_FORCE_REGISTRATION, Me.hWnd)
   End If
End Sub

Private Sub CmdConfImp_Click()
    FrmPrintSetup.Show vbModal, Me
End Sub

Private Sub CmdEditar_Click()
    Dim cEXTENSAO As String
    Dim nPOS As Long
    cEXTENSAO = "XXX"
    nPOS = InStrRev(TxtArquivo, ".")
    If nPOS > 0 Then
        cEXTENSAO = Mid(TxtArquivo, nPOS + 1)
    End If
    If cEXTENSAO = "RTF" Or lARQTXT Then
        cARQRTF = TxtArquivo
        FrmRTf.Show vbModal, Me
    End If

End Sub

Private Sub CmdFiltro_Click()
    On Error GoTo errhandler
    Dim cFILTRO As String
    
    If aRELCFG(11) Then
        ePASS01 = ""
        FrmFiltro.Show vbModal, Me
        filtro = Replace(Replace(eRETU01, "{", ""), "}", "")
    End If
    cFILTRO = FixStr(filtro)
    If Len(aRELCFG(15)) > 0 Then
        cSQL = TrocaSqlWhere(aRELCFG(15), FixStr(filtro))
        Lblsql = cSQL
    End If
    
    Exit Sub
 
errhandler:
    Select Case Err.Number
    Case 20515
        SayErro "Filtro"

    Case Else
        SayErro "Filtro"
    End Select

End Sub

Private Sub cmdimp_Click()
   If Extensao(TxtArquivo.tEXT, "TXT") Then
      imptxt
      Exit Sub
   End If
   If Extensao(TxtArquivo.tEXT, "PDF") Or Extensao(TxtArquivo.tEXT, "HTML") Or Extensao(TxtArquivo.tEXT, "RTF") Then
       CmdVisua_Click
       Exit Sub
   End If
    'If lARQTXT Then
    '    imptxt
    'Else
    '    gerar "I"
    'End If
End Sub

Private Sub imptxt()
    Dim CLINHA As String
    Dim fileFile As Integer
    Dim STRBUFFER As String
       
    If Not FileExist(cARQRTF, True) Then
        Exit Sub
    End If
    CLINHA = ""
    fileFile = FreeFile + 1
    oTXT.Destino = cDESTINO
    oTXT.ABRIR
    If cDESTINO = "IMPNET" Then
        oTXT.caminho = TxtCaminho
    End If
    If cSETUP <> "" Then
        oTXT.impsetup cSETUP
    End If
    Open cARQRTF For Input As #fileFile
    Do While Not EOF(fileFile)
        Input #fileFile, STRBUFFER
        CLINHA = STRBUFFER
        oTXT.implinha (CLINHA)
    Loop
    oTXT.fechar
    Close fileFile
End Sub

Private Sub CmdOrdem_Click()
    ePASS01 = cARQ
    ePASS02 = NomeTableSql(cSQL)
    EscArqOrdem.Show vbModal, Me
    If lRETU Then
        cSQL = TrocaSQLOrder(cSQL, FixStr(eRETU01))
        Lblsql = cSQL
    End If
End Sub


Private Sub CmdShell_Click()
    Dim cARQSHELL  As String
    cARQSHELL = FixStr(TxtArquivo)
    ShellEx cARQSHELL, essSW_SHOWDEFAULT, , , , Me.hWnd
End Sub

Private Sub CmdVisua_Click()
   cARQRTF = TxtArquivo.tEXT
   If Not FileExist(cARQRTF, True) Then
      Exit Sub
   End If
    If Extensao(TxtArquivo.tEXT, "TXT") Then
        ePASS03 = 1
        PrintPreview1.ShowPreview
    End If
    If Extensao(TxtArquivo.tEXT, "PDF") Then
        ShellEx cARQRTF, essSW_SHOWDEFAULT, , , , Me.hWnd
    End If
    If Extensao(TxtArquivo.tEXT, "HTML") Then
       ePASS01 = cARQRTF
       FrmPreview.Show vbModal, Me
    End If
    If Extensao(TxtArquivo.tEXT, "RTF") Then
        cARQRTF = TxtArquivo.tEXT
        RichTextBox1.LoadFile cARQRTF, RtfLoadSaveFormatRTF 'rtfRTF
        ePASS03 = 2
        PrintPreview1.ShowPreview
        RichTextBox1.tEXT = ""
    End If
End Sub
Private Sub PrintPreview1_PrepareReport(Cancel As Boolean)
   If ePASS03 = 1 Then
       MyPrintingTXT
   End If
   If ePASS03 = 2 Then
      MyPrintingRTF
   End If
End Sub
Public Sub MyPrintingRTF()
    PrinterEx.PrintRichTextBox RichTextBox1
    Printer.EndDoc
End Sub
Public Sub MyPrintingTXT()
    Dim fileFile As Integer
    Dim STRBUFFER As String
    If FileExist(cARQRTF, True) Then
        Exit Sub
    End If
    fileFile = FreeFile
    Open cARQRTF For Input As #fileFile
    Do While Not EOF(fileFile)
        'read line
        Input #fileFile, STRBUFFER
        Printer.Print STRBUFFER
    Loop
    Close fileFile
    Printer.EndDoc
End Sub
Private Sub Encerrar_Click()
    Unload Me
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    TeclaEnter KeyCode
End Sub

Private Sub Form_Load()
    Dim X As Long
    Dim cTMP As String
    CenterFormToScreen Me
    filtro = ""
    TxtCaminho = ""
    TxtCaminho.Enabled = False
    OptDestino(6).Value = True
    montaimp
  
    If Extensao(cARQRTF, "TXT") Then
        lARQTXT = True
    End If
    If Extensao(cARQRTF, "MAN") Then
        lARQTXT = True
    End If
    If InStr(UCase(cARQRTF), ".LST") > 0 Then    ''.LST .LTS2 .LTS3...
        lARQTXT = True
    End If
  
  
    If lARQTXT Then
        TxtArquivo = cARQRTF
    Else
        oIni.Path = cARQRTF
        nQTDELAY = oIni.GetSetting("CONFIGURACAO", "QTDELAY", 1)
        ReDim aLAY(nQTDELAY + 1)
        For X = 1 To nQTDELAY
            aLAY(X) = oIni.GetSetting("LAYOUT", StrZero(X, 3), "")
        Next X
        nCOLUNAS = oIni.GetSetting("CONFIGURACAO", "COLUNAS", 80)
        nLINHAS = oIni.GetSetting("CONFIGURACAO", "LINHAS", 60)
        cDESTINO = oIni.GetSetting("CONFIGURACAO", "DESTINO", "ARQ")
        TxtCaminho = oIni.GetSetting("CONFIGURACAO", "CAMINHO", "LPT1")
  
        Select Case cDESTINO
        Case "LPT1"
            OptDestino(0).Value = True
        Case "LPT2"
            OptDestino(1).Value = True
        Case "LPT3"
            OptDestino(2).Value = True
        Case "COM1"
            OptDestino(3).Value = True
        Case "COM2"
            OptDestino(4).Value = True
        Case "ARQ", "ARQUIVO"
            OptDestino(5).Value = True
        Case "PRINTER"
            OptDestino(6).Value = True
        Case "IMPNET"
            OptDestino(7).Value = True
            TxtCaminho.Enabled = True
        End Select

  
        cSETUP = oIni.GetSetting("CONFIGURACAO", "SETUP", "")
        TxtArquivo = oIni.GetSetting("CONFIGURACAO", "NOMEARQUIVO", "TXL_" & zUSER & Format(Now, "_DDMMYYYY_HHMMSS") & ".TXT")
        If aRELCFG(14) = "" Then
            aRELCFG(14) = oIni.GetSetting("CONFIGURACAO", "FILTRO", 1)
        End If
        filtro = aRELCFG(14)
        If aRELCFG(15) = "" Then
            aRELCFG(15) = oIni.GetSetting("CONFIGURACAO", "SQL", 1)
        End If
        cSQL = aRELCFG(15)
        If aARQUIVOS(0) = "" Then
            cTMP = oIni.GetSetting("CONFIGURACAO", "ARQ", 1)
            If Left(cTMP, 3) = "[T]" Then
                cTMP = PegTable("TABELAS", cTMP)
            End If
            If Left(cTMP, 3) = "[P]" Then
                cTMP = PegPath("PATH", cTMP)
            End If
            aARQUIVOS(0) = cTMP
        End If
        cARQ = Caminex(aARQUIVOS(0))
  
    End If


    ''Configura Help
    Me.Caption = cFORMID
    HelpContextID = nFORMID
    

    
    If Len(aRELCFG(6)) = 0 Then
        Label1 = aRELCFG(0)
    Else
        Label1 = aRELCFG(6)
    End If
    If Len(Label1.Caption) = 0 Then
        Label1.Caption = cARQRTF
    End If
    
    
    ' 0-Novo,1-Abrir,2-Salvar,3-Salvacomo
    ' 4-Editar,5-Visualizar,6-Imprimir,7-Exportar
    
    If lARQTXT Then
        aRELCFG(11) = False
        CmdOrdem.Visible = False
        TxtArquivo.Locked = True
        cDESTINO = "LPT1"
        OptDestino(5).Enabled = False
        OptDestino(0).Value = True
    End If

    If Not aDIREITOS(6) Then cmdimp.Visible = False
    If Not aDIREITOS(6) Then CmdConfImp.Visible = False
    If Not aDIREITOS(5) Then CmdVisua.Visible = False
    If Not aDIREITOS(7) Then Salvar(0).Visible = False
    If Not aDIREITOS(7) Then CmdEmail.Visible = False
    If Not aRELCFG(11) Then CmdFiltro.Visible = False
    If Not aRELCFG(11) Then filtro.Visible = False
     
    If Not aDIREITOS(4) Then CmdEditar.Visible = False
    If Not aDIREITOS(4) Then CmdShell.Visible = False

    aRELCFG(14) = FixStr(aRELCFG(14))
    If Len(aRELCFG(14)) > 0 Then
    Else
        If aRELCFG(11) Then
            CmdFiltro_Click
        End If
    End If
    If Len(aRELCFG(15)) > 0 Then
        cSQL = TrocaSqlWhere(aRELCFG(15), FixStr(filtro))
    End If
    
    Lblsql = cSQL

    If Not lARQTXT Then
        If aRELCFG(13) = "" Then                 'nome tabela nao preenchido pega do sql
            aRELCFG(13) = NomeTableSql(cSQL)
        End If
    Else
        TxtArquivo.Enabled = False
        '    Salvar(0).Enabled = False
    End If
  
  
End Sub

Private Sub Listview1_Click()
    Dim nPOS As Long
    Dim cPRINTER As String
    cPRINTER = Listview1
    nPOS = InStr(cPRINTER, "|")
    If nPOS = 0 Then
        cPRINTER = Trim(Mid(cPRINTER, nPOS + 1, 40))
    Else
        cPRINTER = Trim(cPRINTER)
    End If
    SelectPrinter (cPRINTER)
End Sub

Private Sub OptDestino_Click(Index As Integer)
    TxtCaminho.Enabled = False
    Select Case Index
    Case 0
        cDESTINO = "LPT1"
    Case 1
        cDESTINO = "LPT2"
    Case 2
        cDESTINO = "LPT3"
    Case 3
        cDESTINO = "COM1"
    Case 4
        cDESTINO = "COM2"
    Case 5
        cDESTINO = "ARQ"
    Case 6
        cDESTINO = "PRINTER"
    Case 7
        cDESTINO = "IMPNET"
        TxtCaminho.Enabled = True
    End Select
End Sub

Private Sub Salvar_Click(Index As Integer)
    On Error Resume Next
    Dim cARQUIVO As String
    Dim cEXTENSAO As String
    Dim sFILTER As String

    If lARQTXT Then
        cARQUIVO = TxtArquivo
    Else
        Select Case Index
        Case 0
            cEXTENSAO = "TXT"
        Case 1
            cEXTENSAO = "HTML"
        Case 2
            cEXTENSAO = "PDF"
        Case 3
            cEXTENSAO = "RTF"
            
        End Select
    
        sFILTER = "Formato (*." & cEXTENSAO & ")" & vbNullChar & "*." & cEXTENSAO
        cARQUIVO = FileSave(Me, sFILTER, 1, cEXTENSAO, TxtArquivo.tEXT, App.Path, "Salvar " & cEXTENSAO & " Como")
    
        If InStr(cARQUIVO, ".") > 0 Then
            cARQUIVO = Left(cARQUIVO, InStr(cARQUIVO, ".") - 1) + "." & cEXTENSAO
        Else
            cARQUIVO = cARQUIVO & cEXTENSAO
        End If
   
        If Len(cARQUIVO) > 0 Then
        Else
            Alert ("Arquivo Nao Preenchido")
            Exit Sub
        End If


        TxtArquivo = cARQUIVO
    
        Select Case Index
        Case 0
            gerar "A"
        Case 1
            TxtArquivo = TrocaExt(cARQUIVO, "TMP")
            gerar "A"
            txttohtml TxtArquivo, cARQUIVO
        Case 2
            TxtArquivo = TrocaExt(cARQUIVO, "TMP")
            gerar "A"
            ePASS01 = TxtArquivo
            formConvertToPDF.Show
        Case 3
            TxtArquivo = TrocaExt(cARQUIVO, "TMP")
            gerar "A"
            txttoRTF TxtArquivo, cARQUIVO
        End Select
    End If
    
    If CmdEmail.Value = 1 Then
        'servidor,porta,from,to,assunto,anexos,mensagem,enviar e sair
        ePASS01 = Array("", _
                        "", _
                        "", _
                        "", _
                        "", _
                        cARQUIVO, _
                        "", _
                        False)
        FrmSendMail.Show vbModal, Me
    End If

End Sub

Private Sub montaimp()
    Dim xPrinter As Printer
    'Dim sIcone As String
    Dim sDesc As String
    Dim sDriver As String
    Dim idx As Integer
    Dim cPAD As String
    On Error Resume Next

    Listview1.Clear
    Listview1.FixedCols = 0
    Listview1.Rows = 1
    Listview1.Row = 0
    Listview1.Cols = 3
    Listview1.ColWidth(0) = 3000
    Listview1.ColWidth(1) = 3000
    Listview1.ColWidth(2) = 3000
    Listview1.Col = 0
    Listview1 = "Impressora"
    Listview1.Col = 1
    Listview1 = "Porta"
    Listview1.Col = 2
    Listview1 = "Driver"


    For Each xPrinter In Printers
    
        cPAD = " "
        ''If xPrinter = Printer.DeviceName Then 'É a Impressora Padrão
        ''      sIcone = ImageList1.ListImages(1).key
        ''   cPAD = "*"
        ''Else 'Outras Impressoras Instaladas
        ''      sIcone = ImageList1.ListImages(2).key
        ''End If

        sDesc = xPrinter.DeviceName
        sDriver = xPrinter.DriverName
        
        Listview1.AddItem sDesc & vbTab & xPrinter.Port & vbTab & sDriver
        ''ListView1.ListItems.Add , , sDesc, sIcone, sIcone
        ''Listview1.ListItems.Add , , sDesc
        ''Listview1.ListItems(idx).SubItems(lvCad.ColumnHeaders(2).SubItemIndex) = sDriver
    
        idx = idx + 1

    Next xPrinter

End Sub

Private Function SelectPrinter(ByVal Nome As String) As Boolean
    Dim X As Printer
    Dim nLEN As Long
    Nome = UCase(Trim(Nome))
    nLEN = Len(Nome)
    For Each X In Printers
        If UCase(Mid(X.DeviceName, 1, nLEN)) = Nome Then
            Set Printer = X
            SelectPrinter = True
            Exit For
        End If
    Next
    SelectPrinter = False
End Function


