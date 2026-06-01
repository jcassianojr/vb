VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{451B73A5-1563-45D5-A6AC-7B2B7D30B778}#1.1#0"; "BSPrin10.ocx"
Object = "{379157C5-E9BD-43F1-9F83-B037496BED42}#1.3#0"; "vbccr18.ocx"
Begin VB.MDIForm frmPRINCIPAL 
   BackColor       =   &H8000000A&
   Caption         =   " "
   ClientHeight    =   6384
   ClientLeft      =   1572
   ClientTop       =   1896
   ClientWidth     =   10884
   Icon            =   "Principal.frx":0000
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VBCCR18.ImageList Imgmenunew 
      Left            =   2040
      Top             =   840
      _ExtentX        =   804
      _ExtentY        =   804
      ImageWidth      =   16
      ImageHeight     =   16
      InitListImages  =   "Principal.frx":058A
   End
   Begin BSPrinter.PrintPreview PrintPreview1 
      Left            =   480
      Top             =   960
      _ExtentX        =   953
      _ExtentY        =   953
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   780
      Top             =   60
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VBCCR18.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   372
      Left            =   0
      Top             =   6012
      Width           =   10884
      _ExtentX        =   19198
      _ExtentY        =   656
      InitPanels      =   "Principal.frx":05AA
   End
   Begin MSComctlLib.ImageList ImgMenu 
      Left            =   3600
      Top             =   1440
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   49
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":0C3E
            Key             =   "center"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":0FD8
            Key             =   "justificado"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":1372
            Key             =   "left"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":170C
            Key             =   "right"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":1AA6
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":1E40
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":23DA
            Key             =   "cut"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":2774
            Key             =   "diahora"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":2B0E
            Key             =   "editor"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":30A8
            Key             =   "find"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":3442
            Key             =   "marcado"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":359C
            Key             =   "letra01"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":3936
            Key             =   "tamanho"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":3CD0
            Key             =   "paragro"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":406A
            Key             =   "decrese"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":4404
            Key             =   "incrise"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":479E
            Key             =   "imagem"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":4B38
            Key             =   "italico"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":4ED2
            Key             =   "s_Key19"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":526C
            Key             =   "mail"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":5806
            Key             =   "new"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":5BA0
            Key             =   "open"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":613A
            Key             =   "xxx"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":64D4
            Key             =   "imagem2"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":686E
            Key             =   "colar"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":6E08
            Key             =   "printer"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":73A2
            Key             =   "printcfg"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":773C
            Key             =   "sair"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":7AD6
            Key             =   "salvar"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":7E70
            Key             =   "busca"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":820A
            Key             =   "proximabusca"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":85A4
            Key             =   "letra"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":893E
            Key             =   "abc"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":8CD8
            Key             =   "riscado"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":9072
            Key             =   "sub"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":940C
            Key             =   "sob"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":97A6
            Key             =   "minusculas"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":9B40
            Key             =   "maisculas"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":9EDA
            Key             =   "underline"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":A274
            Key             =   "refazer"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":A60E
            Key             =   "info"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":ABA8
            Key             =   "exit"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":B142
            Key             =   "calc"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":B6DC
            Key             =   "menusa"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":BA76
            Key             =   "usuerchave"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":BE10
            Key             =   "odbcinfo"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":BF6A
            Key             =   "todo"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":C404
            Key             =   "lampada"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":C89E
            Key             =   "user"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuMENU 
      Caption         =   "&Tabelas"
      Index           =   0
      Begin VB.Menu mnuSUBMENU 
         Caption         =   "&Sair"
         Index           =   0
      End
   End
   Begin VB.Menu mnuMENU 
      Caption         =   "&Cadastro"
      Index           =   1
      Begin VB.Menu mnuSUBMENU2 
         Caption         =   "&Tipo de títulos"
         Index           =   0
      End
   End
   Begin VB.Menu mnuMENU 
      Caption         =   "&Relatórios"
      Index           =   2
      Begin VB.Menu RELAT 
         Caption         =   "Wrpt"
         Index           =   0
      End
      Begin VB.Menu RELAT 
         Caption         =   "Documentos/Manuais"
         Index           =   1
      End
      Begin VB.Menu RELAT 
         Caption         =   "Gerador"
         Index           =   2
      End
      Begin VB.Menu RELAT 
         Caption         =   "Gerador Padrao"
         Index           =   3
      End
      Begin VB.Menu RELAT 
         Caption         =   "Folha/Ponto/RH"
         Index           =   4
      End
      Begin VB.Menu RELAT 
         Caption         =   "Integrados"
         Index           =   5
      End
      Begin VB.Menu RELAT 
         Caption         =   "Lst(Logix)"
         Index           =   6
      End
      Begin VB.Menu RELAT 
         Caption         =   "Wrpt Independente"
         Index           =   7
      End
   End
   Begin VB.Menu mnuMENU 
      Caption         =   "Con&figurações"
      Index           =   4
      Begin VB.Menu mnuSUBMENU4 
         Caption         =   "&Básica"
         Index           =   0
      End
   End
   Begin VB.Menu mnuMENU 
      Caption         =   "&Operações"
      Index           =   5
      Begin VB.Menu mnuSUBMENU5 
         Caption         =   "Operação 01"
         Index           =   0
      End
   End
   Begin VB.Menu mnuMENU 
      Caption         =   "Sai&r"
      Index           =   7
   End
End
Attribute VB_Name = "frmPRINCIPAL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Load()

  Dim pICONES As String
 ' Dim cLINHA As String
  Dim nRETU As Variant
  Dim zMDBPATH As String
  Dim cmdline As String
  Dim nPOS As Integer
  Dim cARQ As String
  Dim aRETU As Variant
  Dim bACESSO As Boolean
  Dim sSQL As String
  Dim DB As New ADODB.Connection
  Dim rs As New ADODB.Recordset
  Dim carqhelp As String


  StatusBar1.Panels(5).Text = ""

  'CopiarImageList ImgMenuold, Imgmenu

  CenterFormToScreen Me
  

  On Error GoTo ErrorHandler


  If App.PrevInstance Then
    Alert App.EXEName & " já está rodando!"
    End
  End If
  strCodeKey = "SoFtEc1SiStEmAs24055-2265"
  SysCodeKey = "SoFtEc1SiStEmAs24055-2265"
  pICONES = PegPath("PATH", "ICONES")
  zAppPath = PegPath("PATH", "APP")
  zMANA5EMP = PegPath("PATH", "MANA5EMP")
  dbuser = PegPath("PATH", "USER")
  DBWRPT = PegPath("PATH", "WRPT")
  ZMANA5CON = PegPath("PATH", "MANA5CON")
  ZMANA5IMP = PegPath("PATH", "IMPREL")
  ZRELAT = PegPath("PATH", "RELAT")
  zRPTARQ = PegPath("PATH", "RPT") 'agora no ini + "RPT.MDB"
  zusalx = PegPath("PATH", "USALX")
  zusamc = PegPath("PATH", "USAMC")
  cTIPORPT = "R"

 'se sql lite verifica se odbc esta instalado
 If InStr(LCase(dbuser), ".sqlite") > 0 Or InStr(LCase(dbuser), ".sqlite3") > 0 _
                 Or InStr(LCase(dbuser), ".fossil") > 0 Or InStr(LCase(dbuser), ".db3") > 0 _
                 Or (InStr(LCase(dbuser), ".db") > 0) Then
     If IsDriverInstalled("SQLite3 ODBC Driver") Then
       ' MsgBox "O driver SQLite ODBC está pronto para uso!", vbInformation
     Else
        MsgBox "Driver SQLite ODBC não encontrado. Por favor, instale o driver para continuar.", vbCritical
        End
     End If
 End If

  carqhelp = PegPath("PATH", "HELP")
  If FileConnExist(carqhelp) Then
    App.HelpFile = carqhelp
  End If

   'Cria o Arquivo log do mes
  cARQERRO = PegPath("PATH", "LOGCAM") & "LOG" & Format(Now, "MMYYYY") & ".sqlite"
  nRETU = CopyFileWindowsWay(PegPath("PATH", "LOG"), cARQERRO, False)
  If nRETU <> 0 Then
      Alert "Erro Criando Arquivo Log do Mes"
      End
  End If

  Dbname = PegPath("PATH", "MDB")
  zMDBPATH = Dbname
  ZGRPSUB = ""

  zANO = Year(Date)
  zMES = Month(Date)
  zEMPRESA = 1
  
  aRETU = pegue2delimitado(Trim(Command()), "__$", "%__")
  
  bACESSO = False
  If Len(aRETU(0)) > 0 Then
    cmdline = UCase(aRETU(0))
    sSQL = "SELECT IDUSUARIO,USUARIO,IDFOLHA,NOMEFOLHA FROM USUARIO WHERE USUARIO='" & cmdline & "'"
    aRETU = PegSQL(dbuser, sSQL, 3, Array("IDUSUARIO", "IDFOLHA", "NOMEFOLHA"), _
                   Array("NI", "NI", "C"), _
                   Array(0, 0, ""))
    If lRETU Then
      zUSERID = aRETU(0)
      zUSER = cmdline
      zWRPTID = zUSERID
      zIDFOLHA = aRETU(1)
      zNOMEFOLHA = aRETU(2)
      bACESSO = True
    End If
  End If

  If Not bACESSO Then
    frmSENHA.Show vbModal, Me
  End If

  If zIDFOLHA > 0 Then
    If demitido(zIDFOLHA) Then
      End
    End If
  End If




  Me.Caption = App.FileDescription & " Versao:" & App.Major & "." & App.Minor & " Usuário: " & zUSER
  If Not FillMenu("SUBMENU") Then
    Alert "O Sub-Menu não foi carregado!", "Sub Menu"
  End If
  If Not FillMenu("SUBMENU2") Then
    Alert "O Sub-Menu 2 não foi carregado!", "SubMenu 2"
  End If
  ''If Not FillMenu("SUBMENU3") Then
  ''  alert "O Sub-Menu 3 não foi carregado!",  "SubMenu 2"
  ''End If
  If Not FillMenu("SUBMENU4") Then
    Alert "O Sub-Menu 4 não foi carregado!", "SubMenu 4"
  End If
  If Not FillMenu("SUBMENU5") Then
    Alert "O Sub-Menu 5 não foi carregado!", "SubMenu 5"
  End If
  ''* Pega Identificacao do WRPT

  zWRPTID = 0
  sSQL = "SELECT IDUSUARIO FROM USUARIO WHERE USUARIO='" & zUSER & "'"
  aRETU = PegSQL(DBWRPT, sSQL, 1, Array("IDUSUARIO"), Array("NI"), Array(0))
  zWRPTID = aRETU(0)
  GravaLog 0, 0, "Entrar", "Sistema"

  '* Carrega imagens para o ImageList
  DB.ConnectionTimeout = 120
  DB.Open GeracArq(Dbname, , False)
  rs.Open "select * from controle where form='frmMDIPRINCIPAL' and controle='ImageList1'", DB, adOpenForwardOnly, adLockReadOnly
  While Not rs.EOF
    ImageList1.ListImages.Add , rs("Caption"), LoadPicture(pICONES & rs("toolTip"))
    rs.MoveNext
  Wend
  rs.Close
  DB.Close

  StatusBar1.Panels(6).Text = zUSER
  If Trim(PegPath("CITACAO", zUSER, "S")) = "S" Then
    frmDica.Show
  End If
  PrintPreview1.AuxiliaryButtonVisible = PrintPreview1.PrinterExists("Microsoft Print to PDF")
  PrintPreview1.AuxiliaryButtonToolTipText = "Salvar como PDF"
  Exit Sub
ErrorHandler:
  'Alert error$ & " - " & Str$(Err)
  Select Case Err
  Case 3024, 3044
    Alert "Verifiqie ini e banco de dados " & App.EXEName & ".MDB!"
  Case 53
    Alert "Verifique Arquivos mail.exe mail.ini"
    Resume Next
  Case Else
    Alert Err.Description
    Resume Next
  End Select
  Resume Next

End Sub
Public Sub PrintPreview1_AuxiliaryButtonClick(UpdateReport As Boolean)
  PrintPreview1.ShowSaveToFile "Microsoft Print to PDF", "*.pdf"
  UpdateReport = False  ' we don't need to update the report in the Print preview window after this action (the default value of UpdateReport parameter is True)
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)

  Unload Me
  End

End Sub

Private Sub mnuMENU_Click(Index As Integer)
  If Index = 7 Then Unload Me
End Sub

Private Sub mnuSUBMENU2_Click(Index As Integer)
  On Error Resume Next
  If Not AcessaMenu("mnuSUBMENU2", Index) Then Exit Sub
  zRPTARQ = PegPath("PATH", "RPT") ' agora no ini  + "RPT.MDB"
  cTIPORPT = "R"
  Sdb = ""
  zPFTIPO = ""
  ZENGTIP = ""
  iPPAP = 0
  Select Case Index
  Case 0
    zPFTIPO = "PFPRO"
    Sdb = PegPath("PATH", "PF")
    escpf.Show vbModal
  Case 1
    iPPAP = 1
    escPPAP.Show vbModal
  Case 2
    Sdb = PegPath("PATH", "PF")
    escfemea.Show vbModal
  Case 3
    Sdb = PegPath("PATH", "PF")
    escpc.Show vbModal
  Case 4
    Sdb = PegPath("PATH", "IE")
    escIE.Show vbModal
  Case 5
    iPPAP = 2
    escPPAP.Show vbModal
  Case 6
    escDes.Show vbModal
  Case 7
    ZENGTIP = "DES"
    escCLI.Show vbModal
  Case 8
    escPRO.Show vbModal
  Case 9
    iPPAP = 3
    escPPAP.Show vbModal
  Case 10
    ZENGTIP = "PRO"
    escCLI.Show vbModal
  Case 11
    zPFTIPO = "POA"
    Sdb = PegPath("PATH", "POA")
    escPOA.Show vbModal
  Case 12
    FrmFemeaPad.Show vbModal
  Case 13
    iPPAP = 11
    escPPAP.Show vbModal
  Case 14
    zPFTIPO = "PFPRE"
    Sdb = PegPath("PATH", "PFP")
    escpf.Show vbModal
  Case 15
    zPFTIPO = "PFGP12"
    Sdb = PegPath("PATH", "PFG")
    escpf.Show vbModal
  Case 16
    zPFTIPO = "POKA"
    Sdb = PegPath("PATH", "POKA")
    escPOA.Show vbModal
  Case 17
    zPFTIPO = "DISPO"
    Sdb = PegPath("PATH", "DISPO")
    escPOA.Show vbModal
  Case 18
    ePASS01 = "LXINSP"
    Sdb = PegPath("PATH", "PF")
    escms01.Show vbModal, Me
  End Select

End Sub

Private Sub mnuSUBMENU3_Click(Index As Integer)

End Sub

Private Sub mnuSUBMENU4_Click(Index As Integer)
 ' Dim cLINHA As String

  ''todos usurios aceeos
  ''os que necessitam permissao case abaixo apos o check
  Select Case Index

  Case 0
   ' FrmPrintSetup.Show vbModal
    Exit Sub

  Case 1
    cARQRTF = PegPath("HELP", "HELPTXT")
    If FileConnExist(cARQRTF, True) Then
      PrintPreview1.ShowPreview
    End If
    Exit Sub
  Case 6
    zIDTEMP = zUSERID
    frmUSUSENHA.Show vbModal
    Exit Sub
  Case 7
    WritePrivateProfileString "CITACAO", zUSER, "S", App.Path + "\" & App.EXEName & ".INI"
    frmDica.Show
    Exit Sub
  Case 8
    ShellEx "calc"
    Exit Sub
  Case 9
    ShellEx "emailprg.exe", essSW_SHOWDEFAULT, "$" & UCase(zUSER), PegPath("PATH", "MAIL"), , CLng(Me.hwnd)
    Exit Sub
  End Select


  If Not AcessaMenu("mnuSUBMENU4", Index) Then
    Exit Sub
  End If
  Select Case Index
  Case 2
    escuser.Show vbModal
  Case 3
    escMENU.Show vbModal
  Case 4
    escBTN.Show vbModal
  Case 5
    configuraodbc
  End Select
End Sub

Private Sub mnuSUBMENU5_Click(Index As Integer)

  If Not AcessaMenu("mnuSUBMENU5", Index) Then Exit Sub
  Select Case Index
  Case 0
    ExpMS03.Show vbModal
  Case 1
    imgsin.Show vbModal
  Case 2
    ImportaWrptUser
  Case 3
    Alert ("Nao e mais possivel apos migracao P12")
    'Shell "improtheus", vbNormalFocus
  End Select


End Sub

Private Sub ImportaWrptUser()
  Dim oDBORI As ADODB.Connection
  Dim oDBDES As ADODB.Connection
  Dim oRSORI As ADODB.Recordset
  Dim oRSDES As ADODB.Recordset
  Dim cSQL As String
  Dim cARQORI As String
  Dim cARQDES As String
  Dim aCAM As Variant
  Dim x As Integer


  aCAM = Array("EQUIVALENTE", "DATAULT", "IDFOLHA", "NOMEFOLHA", "HORAINI", "HORAFIM", _
               "ATIVO", "WEEKEND", "DATAVAL", "TROCAR", "USUARIO")

  cARQORI = PegPath("PATH", "WRPT")
  cARQDES = PegPath("PATH", "CONTROLE")
  cARQORI = GeracArq(cARQORI, , False)
  cARQDES = GeracArq(cARQDES, , True)
  Set oDBORI = New ADODB.Connection
  Set oDBDES = New ADODB.Connection
  Set oRSORI = New ADODB.Recordset
  Set oRSDES = New ADODB.Recordset

  oDBORI.ConnectionTimeout = 120
  oDBORI.Open cARQORI
  oDBDES.ConnectionTimeout = 120
  oDBDES.Open cARQDES
  oRSORI.Open "USUARIO", oDBORI, adOpenForwardOnly, adLockReadOnly
  While Not oRSORI.EOF
    cSQL = "SELECT * FROM USUARIO WHERE IDUSUARIO=" & FixInt(oRSORI("IDUSUARIO"))
    oRSDES.Open cSQL, oDBDES, adOpenDynamic, adLockOptimistic
    If oRSDES.EOF Then
      oRSDES.AddNew
      oRSDES("IDUSUARIO") = oRSORI("IDUSUARIO")
      oRSDES("SENHA") = oRSORI("SENHA")
    End If
    'For x = 0 To 10
    '   oRSDES(aCAM(x)) = oRSORI(aCAM(x))
    'Next x
    'nao trazer os campos pois pode ter trocado matricula no controle
    'e nao ter trocado no wrpt
    'bloqueando assim o acesso do usuario
    oRSDES.Update
    oRSDES.Close
    oRSORI.MoveNext
  Wend
  oRSORI.Close
  oDBDES.Close
  oDBORI.Close
End Sub

Private Sub mnuSUBMENU_Click(Index As Integer)
  zRPTARQ = PegPath("PATH", "RPT")
  cTIPORPT = "R"
  If UCase(Replace(mnuSUBMENU(Index).Caption, "&", "")) = "SAIR" Then
    Unload Me
  End If
  If Not AcessaMenu("mnuSUBMENU", Index) Then Exit Sub
  Select Case Index
  Case 0                                       'Clientes
    iImage = 1
    escIMG.Show vbModal
  Case 1                                       'Produtos
    iImage = 2
    escIMG.Show vbModal
  Case 2
    escFLX.Show vbModal, Me
  Case 3
    iMU01 = 999
    escIED.Show vbModal
  Case 4
    Zsetor = ""
    escPFD.Show vbModal, Me
  Case 5                                       'Ferramental
    escFER.Show vbModal
  Case 6                                       'Maquinas
    lRETU = False
    eRETU01 = ""
    eRETU02 = ""
    cARQESC = "ME01"
    escNUMNOM.Show vbModal, Me
    'iIMAGE = 4
    'escIMG.Show vbModal
  Case 7                                       'Equipamentos
    nPF = 0
    ePASS01 = ""
    escME04.Show vbModal
  Case 8                                       'Funcionarios
    'iImage = 6
    'escIMG.Show vbModal
    'ePASS01 = "LOGIX"
    'escMP04.Show vbModal, Me
    ePASS01 = ""
    escMP04.Show vbModal, Me
  Case 9                                       'Materia Prima
    iMU01 = 1
    escmu01.Show vbModal
    'iImage = 7
    'escIMG.Show vbModal
  Case 10                                      'componentes
    iMU01 = 2
    escmu01.Show vbModal

    'iImage = 8
    'escIMG.Show vbModal
  Case 11                                      'Consumiveis
    iMU01 = 3
    escmu01.Show vbModal

    'iImage = 9
    'escIMG.Show vbModal
  Case 12                                      ''ETI
    iMU01 = 4
    escmp.Show vbModal
  Case 13                                      ''AREAS
    escMP05.Show vbModal
  Case 14
    iImage = 11                              ''utilitarios
    escIMG.Show vbModal
  Case 15
    iMU01 = 4
    escmu01.Show vbModal
    'iImage = 10 ''Embalagens
    'escIMG.Show vbModal
  Case 16
    iMU01 = 6
    escmu01.Show vbModal
    'iImage = 12 ''Itens Manutencao
    'escIMG.Show vbModal
  Case 17
    iMU01 = 7
    escmu01.Show vbModal
    'iImage = 13 'sUB pRODUTOS
    'escIMG.Show vbModal
  Case 18
    iMU01 = 7
    escmp.Show vbModal, Me
  End Select

End Sub

Private Sub Relat_Click(Index As Integer)
  Select Case Index
  Case 0
    cTIPORPT = "R"
    cARQRTF = ""
    zRPTARQ = PegPath("PATH", "RPT")
    escRPTGRP.Show vbModal, Me
  Case 1
  cARQRTF = ""
    cTIPORPT = "D"
    zRPTARQ = PegPath("PATH", "DOC")
    escRPTGRP.Show vbModal, Me
  Case 2
    cTIPO = "R"
    escmanreg.Show vbModal, Me
  Case 3
    cTIPO = "P"
    escmanreg.Show vbModal, Me
  Case 4
    cARQRTF = ""
    cTIPO = "F"
    zRPTARQ = PegPath("PATH", "FOL")
    escRPTGRP.Show vbModal, Me
  Case 5
    cARQRTF = ""
    cTIPO = "I"
    zRPTARQ = PegPath("PATH", "INT")
    escRPTGRP.Show vbModal, Me
  Case 6
    FrmLstView.Show vbModal, Me
  Case 7
    ShellEx "WRPT", essSW_SHOWDEFAULT, "$" + Trim(zUSER) + "%#", PegPath("PATH", "WRPTX"), , CLng(Me.hwnd)
  End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub configuraodbc()
  Dim lCONTINUA As Boolean
  Dim cCFG As String
  Dim aDAD As Variant
  Dim x As Integer
  lCONTINUA = True
  x = 1
  While lCONTINUA
    cCFG = FixStr(PegPath("ODBCCFG", "ODBCCFG" & StrZero(x, 3)))
    If InStr(cCFG, "|") = 0 Then
      lCONTINUA = False
    Else
      aDAD = Split(cCFG, "|")
      AddDSN aDAD(0), aDAD(1), aDAD(2), aDAD(3), True
    End If
    x = x + 1
  Wend
End Sub
Public Sub PrintPreview1_PrepareReport(Cancel As Boolean)
  Dim fileFile As Integer
  Dim strBuffer As String
  fileFile = FreeFile
  Open cARQRTF For Input As #fileFile
  Do While Not EOF(fileFile)
    'read line
    Input #fileFile, strBuffer
    Printer.Print strBuffer
  Loop
  Close fileFile
End Sub

Public Sub CopiarImageList(Origem, Destino)
    Dim i As Long
    Dim Pic As StdPicture
    Dim Chave As String
    
    ' Limpa o destino antes de copiar (opcional)
    Destino.ListImages.Clear
    
    ' Loop por todos os ícones da origem
    For i = 1 To Origem.ListImages.Count
        ' Pega a imagem da origem
        Set Pic = Origem.ListImages(i).Picture
        Chave = Origem.ListImages(i).Key
        
        ' Adiciona ao vbCCR.ImageList
        ' O vbCCR.ImageList é muito mais flexível com formatos de imagem
        Destino.ListImages.Add , Chave, Pic
    Next i
    
    MsgBox "Cópia concluída! " & Destino.ListImages.Count & " ícones migrados.", vbInformation
End Sub
