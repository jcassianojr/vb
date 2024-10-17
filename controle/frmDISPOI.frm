VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{451B73A5-1563-45D5-A6AC-7B2B7D30B778}#1.1#0"; "BSPrin10.ocx"
Object = "{66E63055-5A66-4C79-9327-4BC077858695}#9.0#0"; "newtab01.OCX"
Begin VB.Form frmDISPOI 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PFI"
   ClientHeight    =   6255
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   10395
   Icon            =   "frmDISPOI.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   10395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin BSPrinter.PrintPreview PrintPreview1 
      Left            =   7800
      Top             =   840
      _ExtentX        =   1191
      _ExtentY        =   1191
   End
   Begin NewTabCtl.NewTab SSTab1 
      Height          =   4815
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   8493
      Tabs            =   1
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
      ControlVersion  =   9
      TabCaption(0)   =   "Item"
      Tab(0).ControlCount=   9
      Tab(0).Control(0)=   "Picture1"
      Tab(0).Control(1)=   "Picture2"
      Tab(0).Control(2)=   "CmdConfImp"
      Tab(0).Control(3)=   "cmdimprimir"
      Tab(0).Control(4)=   "CmdPaste"
      Tab(0).Control(5)=   "Incluirimagem"
      Tab(0).Control(6)=   "DelImg"
      Tab(0).Control(7)=   "VerImg"
      Tab(0).Control(8)=   "Command4"
      Begin VB.PictureBox Picture1 
         Height          =   495
         Left            =   3840
         ScaleHeight     =   435
         ScaleWidth      =   555
         TabIndex        =   6
         Top             =   3360
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.PictureBox Picture2 
         Height          =   3375
         Left            =   120
         ScaleHeight     =   3315
         ScaleWidth      =   3555
         TabIndex        =   5
         Top             =   600
         Width           =   3615
      End
      Begin XPControls.XPButton CmdConfImp 
         Height          =   435
         Left            =   120
         TabIndex        =   9
         Top             =   4080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmDISPOI.frx":058A
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
      Begin XPControls.XPButton cmdimprimir 
         Height          =   435
         Left            =   1800
         TabIndex        =   10
         Top             =   4080
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   767
         Picture         =   "frmDISPOI.frx":0B24
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
      Begin XPControls.XPButton CmdPaste 
         Height          =   435
         Left            =   3480
         TabIndex        =   11
         Top             =   4080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   767
         Picture         =   "frmDISPOI.frx":10BE
         Caption         =   "Copia"
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
      Begin XPControls.XPButton Incluirimagem 
         Height          =   435
         Left            =   4200
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmDISPOI.frx":1658
         Caption         =   "Incluir Imagem"
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
      Begin XPControls.XPButton DelImg 
         Height          =   435
         Left            =   4200
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmDISPOI.frx":1AF2
         Caption         =   "Excluir Imagem"
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
      Begin XPControls.XPButton VerImg 
         Height          =   495
         Left            =   4200
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   2400
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         Picture         =   "frmDISPOI.frx":1F8C
         Caption         =   "Navegar Imagens"
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
         Left            =   4200
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmDISPOI.frx":2426
         Caption         =   "Salvar Imagem"
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
   Begin VB.TextBox TXTFIELDS 
      Height          =   555
      Index           =   1
      Left            =   2400
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   120
      Width           =   6135
   End
   Begin VB.TextBox TXTFIELDS 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   405
      Index           =   0
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   735
   End
   Begin XPControls.XPButton Encerrar 
      Height          =   435
      Left            =   8760
      TabIndex        =   7
      Top             =   600
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "frmDISPOI.frx":28C0
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
      Left            =   8760
      TabIndex        =   8
      Top             =   120
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "frmDISPOI.frx":2E5A
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
   Begin VB.Label lblLabels 
      Caption         =   "Descrição:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblLabels 
      Caption         =   "ORD:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmDISPOI"
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
Dim lTROCOU As Boolean

Private Sub cmdClose_Click()
  On Error Resume Next
  If MDG("Gravar alteraçôes") Then
    For iLOOP = 0 To nCAMPOS - 1
      aVAL(iLOOP) = TXTFIELDS(iLOOP)
    Next iLOOP
    GrvSQL cARQ, cSQL, nCAMPOS, aCAM, aVAL, aFOR, 1 '0 item chave

    If lTROCOU Then
     'cSQL = "select * from DISPOI WHERE NUMERO=" & nPPAP & " AND Item = " & nORD
      ADOGrvBlob cARQ, "DISPOI", Picture1, "IMAGEM", "NUMERO=" & nPPAP & " AND Item = " & nORD
    End If


  End If
  Screen.MousePointer = vbDefault
  Unload Me
End Sub

Private Sub CmdConfImp_Click()
 ' FrmPrintSetup.Show vbModal, Me
End Sub

Private Sub PrintPreview1_PrepareReport(Cancel As Boolean)
  Printer.PaintPicture Picture1, 0, 0
End Sub
Private Sub CmdImprimir_Click()
  If Picture1.Height = 0 Then
    Alert ("Sem Imagem")
  Else
    PrintPreview1.ShowPreview
  End If
End Sub

Private Sub CmdPaste_Click()
  If Picture1.Height = 0 Then
    Alert ("Sem Imagem")
  Else
    Clipboard.Clear
    Clipboard.SetData Picture1.Image, (vbCFBitmap)

    ' CopyEntirePicture Picture1
  End If

End Sub

Private Sub cmdpeg_Click(Index As Integer)
  iMU01 = 404                                  ' 5S
  eRETU02 = ""                                 'Evita Cancelamento Escolha
  escIED.Show vbModal, Me
  If lRETU Then
    TXTFIELDS(Index) = Mid(eRETU02, 1, 1)
  End If
End Sub

Private Sub Command4_Click()
  salvarpict Me, Picture1, "imagem"
End Sub

Private Sub DelImg_Click()
  Set Picture1.Picture = Nothing
  Set Picture2.Picture = Nothing
  lTROCOU = True

End Sub

Private Sub Encerrar_Click()
  If Not MDG("Sair sem gravar") Then
    Exit Sub
  End If
  Screen.MousePointer = vbDefault
  Unload Me
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  TeclaEnter KeyCode
End Sub

Private Sub Form_Load()
  CenterFormToScreen Me
  cARQ = PegPath("PATH", "DISPO")
  cSQL = "select * from DISPOI WHERE NUMERO=" & nPPAP & " AND Item = " & nORD
  nCAMPOS = 2
  aCAM = Array("ITEM", "DESCRICAO")
  aFOR = Array("NI", "C")
  aPAD = Array(0, "")
  aVAL = PegSQL(cARQ, cSQL, nCAMPOS, aCAM, aFOR, aPAD)
  For iLOOP = 0 To nCAMPOS - 1
    TXTFIELDS(iLOOP) = aVAL(iLOOP)
  Next iLOOP
  'ADOPegBlob(cPICURE, cARQ, cTABLE, cWHERE, cCAMPO)
  If ADOPegBlob(Picture1, cARQ, "DISPOI", "NUMERO=" & nPPAP & " AND Item = " & nORD, "IMAGEM") Then 'ADOPegBlob(cARQ, cSQL, Picture1, "IMAGEM") Then
    StretchSourcePictureFromPicture Picture1, Picture2
    If FixNum(eRETU01) > 500000 Then
      Alert ("Imagem Muito Grande,Ajuste o tamanho")
      salvarpict Me, Picture1, "Dispo_" & StrZero(nPPAP, 6) & "_item_" & StrZero(nORD, 3)
      Set Picture1.Picture = Nothing
      Set Picture2.Picture = Nothing
      lTROCOU = True
    End If
  Else
    Set Picture1.Picture = Nothing
    Set Picture2.Picture = Nothing
  End If
  PrintPreview1.AuxiliaryButtonVisible = PrintPreview1.PrinterExists("Microsoft Print to PDF")
  PrintPreview1.AuxiliaryButtonToolTipText = "Salvar como PDF"
End Sub
Public Sub PrintPreview1_AuxiliaryButtonClick(UpdateReport As Boolean)
  PrintPreview1.ShowSaveToFile "Microsoft Print to PDF", "*.pdf"
  UpdateReport = False  ' we don't need to update the report in the Print preview window after this action (the default value of UpdateReport parameter is True)
End Sub
Private Sub Incluirimagem_Click()
  Dim STMPFILE As String
  STMPFILE = OpenArqExt(Me, "", "JPG", "JPEG *.JPG")
  If lerarquivoimagem(STMPFILE, Picture1, Picture2) Then
    lTROCOU = True
  End If
End Sub

'Private Sub lerarquivoimagem(ByVal STMPFILE)
'     If Len(STMPFILE) > 0 Then
'        If FixInt(FileLen(STMPFILE)) > 500000 Then
'           Alert ("Imagem Muito Grande,Ajuste o tamanho")
'           If Not MDG("Anexar mesmo assim-NAO RECOMENDADO") Then
'              Exit Sub
'           Else
'              Exit Sub 'nao permitindo aumentando o banco e travando relatorio crystal
'           End If
'        End If
'        Picture1.Picture = LoadPicture(STMPFILE)
'        StretchSourcePictureFromPicture Picture1, Picture2
'        lTROCOU = True
'     End If
'End Sub

Private Sub txtFields_KeyPress(Index As Integer, KeyAscii As Integer)
  Select Case Index
  Case 0
    KeyAscii = ValiText(KeyAscii, "#NI")
  Case 5, 6, 7, 8, 9
    KeyAscii = ValiText(KeyAscii, "#N")
  End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub VerImg_Click()
  frmPicViewer.Show vbModal, Me
  If lerarquivoimagem(eRETU01, Picture1, Picture2) Then
    lTROCOU = True
  End If
End Sub

