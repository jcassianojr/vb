VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{7020C36F-09FC-41FE-B822-CDE6FBB321EB}#1.2#0"; "vbccr17.ocx"
Begin VB.Form FrmCrwENG 
   Caption         =   "Crystal Report Relatorio"
   ClientHeight    =   5955
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11430
   Icon            =   "FrmCrwEng.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5955
   ScaleWidth      =   11430
   WindowState     =   2  'Maximized
   Begin VB.TextBox TxtArquivo 
      Height          =   285
      Left            =   120
      TabIndex        =   15
      Text            =   "TxtArquivo"
      Top             =   1920
      Width           =   7335
   End
   Begin VB.CheckBox CmdEmail 
      Caption         =   "Enviar Email (MAPI)"
      Height          =   435
      Left            =   120
      TabIndex        =   14
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CheckBox CmdEmail1 
      Caption         =   "Enviar Email (Winsock)"
      Height          =   555
      Left            =   120
      TabIndex        =   13
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox TxtAssunto 
      Height          =   375
      Left            =   2580
      TabIndex        =   9
      Top             =   3840
      Width           =   5775
   End
   Begin VB.TextBox TxtDestino 
      Height          =   375
      Left            =   2580
      TabIndex        =   8
      Top             =   4320
      Width           =   5775
   End
   Begin VB.TextBox TxtTexto 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   2520
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Filtro"
      Top             =   4800
      Width           =   7500
   End
   Begin VB.TextBox TxtComp 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3360
      Width           =   7455
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
      Height          =   675
      Left            =   240
      TabIndex        =   4
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
         Left            =   720
         TabIndex        =   5
         Top             =   240
         Width           =   8895
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Exportar/Salvar Como"
      Height          =   1155
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   6015
      Begin VB.CommandButton CmdExport 
         Caption         =   "Exportar Outros Formatos"
         Height          =   735
         Left            =   4920
         TabIndex        =   3
         Top             =   240
         Width           =   945
      End
      Begin XPControls.XPButton salvar 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   240
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   661
         Picture         =   "FrmCrwEng.frx":058A
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
      Begin XPControls.XPButton salvar 
         Height          =   375
         Index           =   1
         Left            =   960
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   240
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   661
         Picture         =   "FrmCrwEng.frx":0924
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
      Begin XPControls.XPButton salvar 
         Height          =   375
         Index           =   2
         Left            =   1800
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   240
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   661
         Picture         =   "FrmCrwEng.frx":0CBE
         Caption         =   "doc"
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
      Begin XPControls.XPButton salvar 
         Height          =   375
         Index           =   3
         Left            =   2640
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   240
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   661
         Picture         =   "FrmCrwEng.frx":1258
         Caption         =   "xls"
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
      Begin XPControls.XPButton salvar 
         Height          =   375
         Index           =   4
         Left            =   3480
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   240
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   661
         Picture         =   "FrmCrwEng.frx":17F2
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
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Height          =   1155
      Left            =   6360
      TabIndex        =   1
      Top             =   720
      Width           =   4065
      Begin XPControls.XPButton CmdVisua 
         Height          =   435
         Left            =   600
         TabIndex        =   17
         Top             =   120
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   767
         Picture         =   "FrmCrwEng.frx":1B8C
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
      Begin XPControls.XPButton cmdimp 
         Height          =   435
         Left            =   600
         TabIndex        =   18
         Top             =   600
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   767
         Picture         =   "FrmCrwEng.frx":2126
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
      Begin XPControls.XPButton CmdConfImp 
         Height          =   435
         Left            =   2280
         TabIndex        =   19
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "FrmCrwEng.frx":26C0
         Caption         =   "Propriedades"
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
         Height          =   435
         Left            =   2280
         TabIndex        =   20
         Top             =   120
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "FrmCrwEng.frx":2C5A
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
   Begin VB.TextBox FILTRO 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Filtro"
      Top             =   2400
      Width           =   7380
   End
   Begin XPControls.XPButton CmdFiltro 
      Height          =   435
      Left            =   7560
      TabIndex        =   16
      Top             =   2520
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   767
      Picture         =   "FrmCrwEng.frx":31F4
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
   Begin XPControls.XPButton CmdEditar 
      Height          =   315
      Left            =   7560
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   1920
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      Picture         =   "FrmCrwEng.frx":378E
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
      Left            =   8520
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   1920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      Picture         =   "FrmCrwEng.frx":3D28
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
      Left            =   9720
      TabIndex        =   28
      Top             =   1920
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
   Begin XPControls.XPButton CmdPreview 
      Height          =   315
      Left            =   10080
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   1920
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      Picture         =   "FrmCrwEng.frx":42C2
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
   Begin XPControls.XPButton CmdMudaFec 
      Height          =   375
      Left            =   7800
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   3360
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   661
      Picture         =   "FrmCrwEng.frx":485C
      Caption         =   "Competência "
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
   Begin VBCCR17.RichTextBox RichTextBox1 
      Height          =   255
      Left            =   10200
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   2400
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      MultiLine       =   -1  'True
      TextRTF         =   "FrmCrwEng.frx":4BF6
   End
   Begin VB.Label Label2 
      Caption         =   "Assunto"
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   1440
      TabIndex        =   12
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Destinatario(s)"
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   1440
      TabIndex        =   11
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Texto"
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   1440
      TabIndex        =   10
      Top             =   4800
      Width           =   1035
   End
End
Attribute VB_Name = "FrmCrwENG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CrystalDatabase As CRAXDRT.Database
Dim CrystalApp As CRAXDRT.Application
Dim CrystalReport As CRAXDRT.Report
Dim CRYSTALTABLE As CRAXDRT.DatabaseTable
''Dim CrystalParameter As CRAXDRT.ParameterFieldDefinition
''Dim CrystalFormula As CRAXDRT.FormulaFieldDefinition
Dim CrystalExportOptions As CRAXDRT.ExportOptions
Option Explicit
Const nFORMID = 1069
Const cFORMID = "Imprimir Relatorios Crystal-Engine 10"

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
    If cEXTENSAO = "RTF" Or cEXTENSAO = "TXT" Then
        cARQRTF = TxtArquivo
        FrmRTf.Show
    End If

End Sub

Private Sub CmdEmail_Click()
    If CmdEmail.Value = 1 Then
        TxtAssunto.Enabled = True
        TxtTexto.Enabled = True
        TxtDestino.Enabled = True
        CmdEmail1.Value = 0
    Else
        TxtAssunto.Enabled = False
        TxtTexto.Enabled = False
        TxtDestino.Enabled = False
    End If

End Sub

Private Sub CmdEmail1_Click()
    If CmdEmail1.Value = 1 Then
        TxtAssunto.Enabled = True
        TxtTexto.Enabled = True
        TxtDestino.Enabled = True
        CmdEmail.Value = 0
    Else
        TxtAssunto.Enabled = False
        TxtTexto.Enabled = False
        TxtDestino.Enabled = False
    End If

End Sub

Private Sub CmdExport_Click()
    On Error Resume Next
    CrystalReport.Export True
End Sub

Private Sub CmdFiltro_Click()
    On Error GoTo errhandler
    Dim cFILTRO As String
    
    If aRELCFG(11) Then
        ePASS01 = "CRYSTAL"
        FrmFiltro.Show vbModal, Me
        FILTRO = eRETU01
    End If
    cFILTRO = CStr(FILTRO.tEXT)
    
    If Len(cFILTRO) > 0 Then
        CrystalReport.RecordSelectionFormula = cFILTRO
    End If
    Exit Sub
 
errhandler:
    Select Case Err.Number
    Case 20515
        SayErro "Filtro"
        CrystalReport.RecordSelectionFormula = ""
    Case Else
        SayErro "Filtro"
    End Select
End Sub

Private Sub cmdimp_Click()
    On Error Resume Next
    CrystalReport.PrintOut
End Sub

Private Sub CmdMudaFec_Click()
    'Dim nARQUIVOS
    'Dim x
    'Mana5Fec
    'TxtComp.Text = aARQFEC(4) & " - " & aARQUIVOS(0)
    'nARQUIVOS = UBound(aARQUIVOS)
    'x = 0
    'For Each CRYSTALTABLE In CrystalDatabase.Tables
    '    If x <= nARQUIVOS Then
    '       If FileExist(aARQUIVOS(x), True) Then
    '          CRYSTALTABLE.location = aARQUIVOS(x)
    ''CrystalTable.SetLogOnInfo "PDBDAO", "", "userid", "senha"
    ''CrystalTable.SetLogOnInfo "crystalserver", "Databasename", "userid", "Password"
    '        End If
    '     End If
    'x = x + 1
    'Next
End Sub

Private Sub CmdPreview_Click()
   cARQRTF = TxtArquivo.tEXT
   If Not FileExist(cARQRTF, True) Then
      Exit Sub
   End If
    If Extensao(TxtArquivo.tEXT, "TXT") Then
        PrinterEx.ShowPrintPreview Me, "MyPrintingTXT"
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
        RichTextBox1.LoadFile cARQRTF, RtfLoadSaveFormatRTF '/ rtfRTF
        PrinterEx.ShowPrintPreview Me, "MyPrintingRTF"
        RichTextBox1.tEXT = ""
    End If
End Sub
Public Property Get Printer() As Printer
    Set Printer = vbExtra.Printer2
End Property
Public Property Set Printer(nPrinter As Printer)
    Set vbExtra.Printer2 = nPrinter
End Property
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
Private Sub CmdShell_Click()
    Dim cARQSHELL  As String
    cARQSHELL = FixStr(TxtArquivo)
    ShellEx cARQSHELL, essSW_SHOWDEFAULT, , , , Me.hWnd
End Sub

Private Sub CmdVisua_Click()
    On Error Resume Next
    FrmCrwENGViewer.Show vbModal
    'Form1.Show vbModal
    '    FrmCrwPrv.Show vbModal, Me
    ''CrystalReport.Preview
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Set CrystalApp = Nothing
    Set CrystalReport = Nothing
    ''Set CrystalPrevieW1 = Nothing
    'Set CrystalExportOptions = Nothing
    'Set CrystalDatabase = Nothing
    Set CRYSTALTABLE = Nothing
End Sub

Private Sub Encerrar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim X, nARQUIVOS As Integer
    Dim cFILTRO
    CenterFormToScreen Me
    TxtArquivo = zUSER & Format(Now, "_DDMMYYYY_HHMMSS") & ".TXT"
    ''Configura Help
    Me.Caption = cFORMID
    HelpContextID = nFORMID
    
     
    TxtComp.Visible = False
    CmdMudaFec.Visible = False
    
    
    If Len(aARQFEC(0)) > 0 Or Len(aARQFEC(1)) > 0 Or Len(aARQFEC(2)) Or Len(aARQFEC(3)) Then
        TxtComp.Visible = True
        CmdMudaFec.Visible = True
        CmdMudaFec.Enabled = True
        TxtComp.tEXT = aARQFEC(4) & " - " & aARQUIVOS(0)
    End If
    
    If Len(aRELCFG(6)) = 0 Then
        Label1 = aRELCFG(0)
    Else
        Label1 = aRELCFG(6)
    End If
    
    ' 0-Novo,1-Abrir,2-Salvar,3-Salvacomo
    ' 4-Editar,5-Visualizar,6-Imprimir,7-Exportar

    If Not aDIREITOS(6) Then cmdimp.Visible = False
    If Not aDIREITOS(6) Then CmdConfImp.Visible = False
    If Not aDIREITOS(5) Then CmdVisua.Visible = False
    If Not aDIREITOS(7) Then salvar(2).Visible = False
    If Not aDIREITOS(7) Then salvar(3).Visible = False
    If Not aDIREITOS(7) Then salvar(0).Visible = False
    If Not aDIREITOS(7) Then salvar(1).Visible = False
    If Not aRELCFG(11) Then CmdFiltro.Visible = False
    If Not aRELCFG(11) Then FILTRO.Visible = False
    
    If Not aDIREITOS(7) Then CmdEmail.Visible = False
    If Not aDIREITOS(7) Then CmdEmail1.Visible = False
    If Not aDIREITOS(7) Then TxtAssunto.Visible = False
    If Not aDIREITOS(7) Then TxtDestino.Visible = False
    If Not aDIREITOS(7) Then TxtTexto.Visible = False
    If Not aDIREITOS(7) Then Label2.Visible = False
    If Not aDIREITOS(7) Then Label3.Visible = False
    If Not aDIREITOS(7) Then Label4.Visible = False
    
    If Not aDIREITOS(4) Then CmdEditar.Visible = False
    If Not aDIREITOS(4) Then CmdShell.Visible = False

    
    CmdEmail_Click
     
    Set CrystalApp = New CRAXDRT.Application
    ''    Set CrystalApp = CreateObject("CrystalRuntime.Application")

    Set CrystalReport = CrystalApp.OpenReport(cARQRTF)
    Set CrystalExportOptions = CrystalReport.ExportOptions
    
    
    
    ''    Set CrystalPreview = CrystalReport.PrintWindowOptions

      
    ''    With CrystalPreview
    ''        .HasCancelButton = True
    ''        .HasCloseButton = True
    ''        .HasNavigationControls = True
    ''        .HasProgressControls = True
    ''        .HasZoomControl = True
    ''        .HasPrintButton = aDIREITOS(6)
    ''        .HasPrintSetupButton = aDIREITOS(6)
    ''        .HasGroupTree = aRELCFG(4)
    ''        .HasSearchButton = aRELCFG(5)
    ''        .HasExportButton = aDIREITOS(7)
    ''        .HasRefreshButton = True
    ''    End With
  
    '
  
  
  
  
    nARQUIVOS = UBound(aARQUIVOS)
    X = 0
    Set CrystalDatabase = CrystalReport.Database
    For Each CRYSTALTABLE In CrystalDatabase.Tables
        If X <= nARQUIVOS Then
            If Len(aARQUIVOS(X)) > 0 Then
                If FileExist(aARQUIVOS(X), True) Then
                    CRYSTALTABLE.location = aARQUIVOS(X)
                    
                    ''CrystalTable.SetLogOnInfo "PDBDAO", "", "userid", "senha"
                    ''CrystalTable.SetLogOnInfo "crystalserver", "Databasename", "userid", "Password"
                    
                End If
            End If
        End If
        X = X + 1
    Next
    aRELCFG(14) = FixStr(aRELCFG(14))
    If Len(aRELCFG(14)) > 0 Then
        cFILTRO = aRELCFG(14)
        FILTRO = cFILTRO
        CrystalReport.RecordSelectionFormula = cFILTRO
    Else
        If aRELCFG(11) Then
            CmdFiltro_Click
        End If
    End If
    
     
End Sub

Private Sub Salvar_Click(Index As Integer)
    On Error GoTo error
    Dim cARQUIVO As String
    Dim cEXTENSAO As String
    Dim sFILTER As String
    If CmdEmail.Value = 1 Then
        CrystalExportOptions.DestinationType = crEDTEMailMAPI
        CrystalExportOptions.MailMessage = FixStr(TxtTexto)
        CrystalExportOptions.MailSubject = FixStr(TxtAssunto)
        CrystalExportOptions.MailToList = FixStr(TxtDestino)
    Else
        CrystalExportOptions.DestinationType = crEDTDiskFile
    End If
    Select Case Index
    Case 0
        CrystalExportOptions.FormatType = crEFTRichText
        cEXTENSAO = "RTF"
    Case 1
        CrystalExportOptions.FormatType = crEFTText
        cEXTENSAO = "TXT"
    Case 2
        CrystalExportOptions.FormatType = crEFTWordForWindows
        cEXTENSAO = "DOC"
    Case 3
        CrystalExportOptions.FormatType = crEFTExcel80
        cEXTENSAO = "XLS"
    Case 4
        CrystalExportOptions.FormatType = crEFTPortableDocFormat
        cEXTENSAO = "PDF"
            
    End Select
    
    
  
    sFILTER = "Formato (*." & cEXTENSAO & ")" & vbNullChar & "*." & cEXTENSAO
    cARQUIVO = FileSave(Me, sFILTER, 1, cEXTENSAO, TxtArquivo.tEXT, App.Path, "Salvar " & cEXTENSAO & " Como")
          
        
      
    If InStr(cARQUIVO, ".") Then
        cARQUIVO = Left(cARQUIVO, InStr(cARQUIVO, ".") - 1) + "." & cEXTENSAO
    Else
        cARQUIVO = cARQUIVO & cEXTENSAO
    End If
   
    If Len(cARQUIVO) > 0 Then
        CrystalExportOptions.DiskFileName = cARQUIVO
        CrystalReport.Export False
    Else
        Alert ("Arquivo Nao Preenchido")
    End If
    TxtArquivo = cARQUIVO
    
    If CmdEmail1.Value = 1 Then
        'servidor,porta,from,to,assunto,anexos,mensagem,enviar e sair
        ePASS01 = Array("", _
                        "", _
                        "", _
                        FixStr(TxtDestino), _
                        FixStr(TxtAssunto), _
                        cARQUIVO, _
                        FixStr(TxtTexto), _
                        False)
        FrmSendMail.Show vbModal, Me
    End If

    Exit Sub
error:
    Select Case Err.Number
    Case Else
        SayErro
    End Select

End Sub

Private Sub TxtDestino_LostFocus()
    If Not CheckEmail(TxtDestino) Then
        TxtDestino.SetFocus
    End If
End Sub


