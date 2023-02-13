VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{EA478B61-D9EC-47F6-BB21-95A533AF2251}#1.0#0"; "TabExC01.OCX"
Object = "{451B73A5-1563-45D5-A6AC-7B2B7D30B778}#1.0#0"; "BSPrin10.ocx"
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
      LcK1b           =   "yefT59bnyufI583n1ufV59HnyefN58nn0+fK58nnzefR58rn1+fW5w=="
      LcK2b           =   "reeI58zn+ue059jn1Of656fni+eN543nleeM54HnlueR5w=="
      AmbientBb       =   $"frmDISPOI.frx":058A
   End
   Begin TabExCtl.SSTabEx SSTab1 
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
         Picture         =   "frmDISPOI.frx":062A
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
         Picture         =   "frmDISPOI.frx":0BC4
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
         Picture         =   "frmDISPOI.frx":115E
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
         Picture         =   "frmDISPOI.frx":16F8
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
         Picture         =   "frmDISPOI.frx":1B92
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
         Picture         =   "frmDISPOI.frx":202C
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
         Picture         =   "frmDISPOI.frx":24C6
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
      Picture         =   "frmDISPOI.frx":2960
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
      Picture         =   "frmDISPOI.frx":2EFA
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
        GrvSQL cARQ, cSQL, nCAMPOS, aCAM, aVAL, aFOR
       
        If lTROCOU Then
            ADOGrvBlob cARQ, cSQL, Picture1, "IMAGEM"
        End If

       
    End If
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub CmdConfImp_Click()
    FrmPrintSetup.Show vbModal, Me
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
    If ADOPegBlob(cARQ, cSQL, Picture1, "IMAGEM") Then
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
    UpdateReport = False ' we don't need to update the report in the Print preview window after this action (the default value of UpdateReport parameter is True)
End Sub
Private Sub Incluirimagem_Click()
    Dim STMPFILE
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

