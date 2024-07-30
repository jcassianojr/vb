VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmPRINCIPAL 
   BackColor       =   &H8000000A&
   Caption         =   " Menu Inicial"
   ClientHeight    =   3990
   ClientLeft      =   1575
   ClientTop       =   1890
   ClientWidth     =   10530
   Icon            =   "Principa.frx":0000
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   780
      Top             =   180
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   3675
      Width           =   10530
      _ExtentX        =   18574
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            Object.Width           =   1588
            MinWidth        =   1587
            Picture         =   "Principa.frx":058A
            TextSave        =   "20:13"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            Object.Width           =   2302
            MinWidth        =   2293
            Picture         =   "Principa.frx":0B24
            TextSave        =   "30/07/2024"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Picture         =   "Principa.frx":10BE
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8281
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgMenu 
      Left            =   5850
      Top             =   180
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   49
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":1458
            Key             =   "center"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":17F2
            Key             =   "justificado"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":1B8C
            Key             =   "left"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":1F26
            Key             =   "right"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":22C0
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":265A
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":2BF4
            Key             =   "cut"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":2F8E
            Key             =   "diahora"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":3328
            Key             =   "editor"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":38C2
            Key             =   "find"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":3C5C
            Key             =   "marcado"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":3DB6
            Key             =   "letra01"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":4150
            Key             =   "tamanho"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":44EA
            Key             =   "paragro"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":4884
            Key             =   "decrese"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":4C1E
            Key             =   "incrise"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":4FB8
            Key             =   "imagem"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":5352
            Key             =   "italico"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":56EC
            Key             =   "s_Key19"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":5A86
            Key             =   "mail"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":6020
            Key             =   "new"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":63BA
            Key             =   "open"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":6954
            Key             =   "xxx"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":6CEE
            Key             =   "imagem2"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":7088
            Key             =   "colar"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":7622
            Key             =   "printer"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":7BBC
            Key             =   "printcfg"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":7F56
            Key             =   "sair"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":82F0
            Key             =   "salvar"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":868A
            Key             =   "busca"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":8A24
            Key             =   "proximabusca"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":8DBE
            Key             =   "letra"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":9158
            Key             =   "abc"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":94F2
            Key             =   "riscado"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":988C
            Key             =   "sub"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":9C26
            Key             =   "sob"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":9FC0
            Key             =   "minusculas"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":A35A
            Key             =   "maisculas"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":A6F4
            Key             =   "underline"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":AA8E
            Key             =   "refazer"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":AE28
            Key             =   "info"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":B3C2
            Key             =   "exit"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":B95C
            Key             =   "calc"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":BEF6
            Key             =   "menusa"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":C290
            Key             =   "usuerchave"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":C62A
            Key             =   "odbcinfo"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":C784
            Key             =   "todo"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":CC1E
            Key             =   "lampada"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":D0B8
            Key             =   "user"
         EndProperty
      EndProperty
   End
   Begin VB.Menu Menu1 
      Caption         =   "&Gerador"
      Tag             =   "&Gerador|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      Begin VB.Menu mnuSUBMENU1 
         Caption         =   "Especificos"
         Index           =   0
         Tag             =   "Especificos|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu mnuSUBMENU1 
         Caption         =   "Padrao"
         Index           =   1
         Tag             =   "Padrao|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
   End
   Begin VB.Menu Menu2 
      Caption         =   "&Relatórios"
      Tag             =   "&Relatórios|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
   End
   Begin VB.Menu Menu3 
      Caption         =   "&Documentos/Manuais"
      Tag             =   "&Documentos/Manuais|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
   End
   Begin VB.Menu Menu4 
      Caption         =   "&Folha/Pto/RH"
      Tag             =   "&Folha/Pto/RH|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
   End
   Begin VB.Menu Menu5 
      Caption         =   "&Integrados"
      Tag             =   "&Integrados|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
   End
   Begin VB.Menu Menu6 
      Caption         =   "&Lst(Logix)"
      Tag             =   "&Lst(Logix)|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
   End
   Begin VB.Menu Menu7 
      Caption         =   "&Configurações"
      Tag             =   "&Configurações|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      Begin VB.Menu mnuSUBMENU4 
         Caption         =   "Impressoras"
         Index           =   0
         Tag             =   "Impressoras|#printcfg|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu mnuSUBMENU4 
         Caption         =   "Alteracoes"
         Index           =   1
         Tag             =   "Alteracoes|#info|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu mnuSUBMENU4 
         Caption         =   "Usuarios"
         Index           =   2
         Tag             =   "Usuarios|#user|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu mnuSUBMENU4 
         Caption         =   "Menus"
         Index           =   3
         Tag             =   "Menus|#menusa|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu mnuSUBMENU4 
         Caption         =   "Botoes"
         Index           =   4
         Tag             =   "Botoes|#todo|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu mnuSUBMENU4 
         Caption         =   "Configurar ODBC"
         Index           =   5
         Tag             =   "Configurar ODBC|#odbcinfo|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu mnuSUBMENU4 
         Caption         =   "Troca Senha"
         Index           =   6
         Tag             =   "Troca Senha|#usuerchave|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu mnuSUBMENU4 
         Caption         =   "Frase do Dia"
         Index           =   7
         Tag             =   "Frase do Dia|#lampada|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu mnuSUBMENU4 
         Caption         =   "Calculadora"
         Index           =   8
         Tag             =   "Calculadora|#calc|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
   End
   Begin VB.Menu Menu8 
      Caption         =   "&Sair"
      Tag             =   "&Sair|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
   End
End
Attribute VB_Name = "frmPRINCIPAL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents MenuEvents As CEvents
Attribute MenuEvents.VB_VarHelpID = -1
Private Sub MDIForm_Load()

'/ If gbSubClassMenu is False, the menu is not subclassed
  gbSubClassMenu = True

  If gbSubClassMenu Then SubClassMenuXP

  Dim cLINHA As String
  Dim pICONES As String
  Dim nRETU
  Dim cmdline As String
  Dim bACESSO As Boolean
  Dim nPOS As Long
  Dim sSQL As String
  Dim cARQ As String
  Dim aRETU As Variant
  Dim DAODB As ADODB.Connection
  Dim DAORS As ADODB.Recordset
  Dim cARQICO As String
  Dim carqhelp As String


  On Error GoTo ErrorHandler

  CenterFormToScreen Me

  '    aDIREITOS = Array(True, True, True, True, True, True, True, True)
  '    FrmRTf.Show vbModal

  '   aDIREITOS = Array(True, True, True, True, True, True, True, True)
  '   frmIniEditor.Show vbModal


  If App.PrevInstance Then

    MsgBox App.EXEName & " já está rodando!", 48
    End

  End If

  '* Pegando configuracoes do arq .INI
  strCodeKey = "SoFtEc1SiStEmAs24055-2265"
  SysCodeKey = "SoFtEc1SiStEmAs24055-2265"

  'Caminhos
  pICONES = PegPath("PATH", "ICONES")
  zAppPath = PegPath("PATH", "APP")
  zMANA5EMP = PegPath("PATH", "MANA5EMP")
  ZMANA5CON = PegPath("PATH", "MANA5CON")
  ZMANA5IMP = PegPath("PATH", "IMPREL")
  ZRELAT = PegPath("PATH", "RELAT")
  zusalx = PegPath("PATH", "USALX")
  zusamc = PegPath("PATH", "USAMC")
  dbuser = PegPath("PATH", "USER")
  DBWRPT = PegPath("PATH", "WRPT")
  Dbname = PegPath("PATH", "MDB") ' agora no ini+ "WRPT.MDB"

  If Not FileExist(Dbname, True) Then
    Alert ("Erro Ini [PATH] MDB=")
    End
  End If
  If Not FileExist(dbuser, True) Then
    Alert ("Erro Ini [PATH] USER=")
    End
  End If
  If Not FileExist(DBWRPT, True) Then
    Alert ("Erro Ini [PATH] WRPT=")
    End
  End If


  carqhelp = PegPath("PATH", "HELP")
  If FileExist(carqhelp) Then
    App.HelpFile = carqhelp
  End If

  Set DAODB = New ADODB.Connection
  Set DAORS = New ADODB.Recordset

    'Cria o Arquivo log do mes
  cARQERRO = PegPath("PATH", "LOGCAM") & "LOG" & Format(Now, "MMYYYY") & ".sqlite"
  nRETU = CopyFileWindowsWay(PegPath("PATH", "LOG"), cARQERRO, False)
  If nRETU <> 0 Then
      Alert "Erro Criando Arquivo Log do Mes"
      End
  End If

  zEMPRESA = 1
  zMES = Month(Date)
  zANO = Year(Date)

  ZGRPSUB = ""

  cmdline = Trim(Command())
  bACESSO = False
  If InStr(cmdline, "__$") > 0 And InStr(cmdline, "%__") > 0 Then
    nPOS = InStr(cmdline, "__$")
    cmdline = Mid(cmdline, 4)
    nPOS = InStr(cmdline, "%__")
    cmdline = UCase(Left(cmdline, Len(cmdline) - 3))
    DAODB.ConnectionTimeout = 120
    DAODB.Open GeracArq(dbuser, , False)
    sSQL = "SELECT IDUSUARIO,USUARIO,IDFOLHA,NOMEFOLHA FROM USUARIO WHERE USUARIO='" & cmdline & "'"
    DAORS.Open sSQL, DAODB, adOpenForwardOnly, adLockReadOnly
    If Not DAORS.EOF Then
      zUSERID = FixNum(DAORS("IDUSUARIO"))
      zUSER = cmdline
      zWRPTID = zUSERID
      zIDFOLHA = FixNum(DAORS("IDFOLHA"))
      zNOMEFOLHA = FixStr(DAORS("NOMEFOLHA"))
      bACESSO = True
    End If
    DAORS.Close
    DAODB.Close
  End If

  If Not bACESSO Then
    frmSENHA.Show vbModal, Me
  End If

  If zIDFOLHA > 0 Then
    If demitido(zIDFOLHA) Then
      End
    End If
  End If

  Me.Caption = App.FileDescription & " Versao:" & App.Major & "." & App.Minor & "Usuário: " & zUSER


  ''cLINHA = "MAIL $" & UCase(zUSER)
  ' CLINHA = "$" & UCase(zUSER)
  ' ShellEx "MAIL.EXE", , CLINHA, , , Me.hWnd
  ShellEx "MAIL", essSW_SHOWDEFAULT, "$" & UCase(zUSER), PegPath("PATH", "MAIL"), , Me.hWnd


  '* Carrega imagens para o ImageList

  DAODB.CursorLocation = adUseClient
  DAODB.ConnectionTimeout = 120
  DAODB.Open GeracArq(Dbname, , False)



  DAORS.Open "select * from controle where form='frmMDIPRINCIPAL'", DAODB, adOpenForwardOnly, adLockReadOnly

  With DAORS
    Do While Not DAORS.EOF  '!controle = "ImageList1"
      cARQICO = pICONES & DAORS.Fields("TOOLTIP")
      If FileExist(cARQICO, True) Then
        ImageList1.ListImages.Add , DAORS.Fields("Caption"), LoadPicture(cARQICO)
      Else
        End
      End If
      .MoveNext
    Loop
  End With
  DAORS.Close
  DAODB.Close
  Set DAODB = Nothing
  Set DAORS = Nothing



  StatusBar1.Panels(6).tEXT = zUSER
  If Trim(PegPath("CITACAO", zUSER, "S")) = "S" Then
    frmDica.Show
  End If
  Exit Sub

ErrorHandler:


  '* MsgBox Error$ & " - " & str$(Err), vbExclamation, "Erro!"

  Select Case Err

  Case 3024, 3044

    'set CommonDialog properties before showing


    MsgBox "Verifiqie ini e banco de dados " & App.EXEName & ".MDB!", vbExclamation, "Erro!"

  Case 53


    MsgBox "Verifique Arquivos mail.exe mail.ini", vbExclamation, "Erro!"
    Resume Next



  Case Else

    Resume Next

  End Select

End Sub
Private Sub MDIForm_Unload(Cancel As Integer)


  If gbSubClassMenu Then
    '/ prevent error if the menu is not subclassed
    On Error Resume Next
    '/ release object
    Call objMenuEx.Uninstall(Me.hWnd, ImgMenu, MenuEvents)
    Set MenuEvents = Nothing
    Set objMenuEx = Nothing
  End If


  Unload Me
  End
End Sub
Private Sub Menu2_Click()
  cTIPORPT = "R"
  zRPTARQ = PegPath("PATH", "RPT") ' agora no ini + "RPT.MDB"
  escRPTGRP.Show vbModal, Me
End Sub

Private Sub Menu3_Click()
  cTIPORPT = "D"
  zRPTARQ = PegPath("PATH", "DOC") ' agora no ini + "RTF.MDB"
  escRPTGRP.Show vbModal, Me
End Sub

Private Sub Menu4_Click()
  cTIPORPT = "F"
  zRPTARQ = PegPath("PATH", "FOL") 'agora no ini + "RPTFOL.MDB"
  escRPTGRP.Show vbModal, Me
End Sub

Private Sub Menu5_Click()
  cTIPORPT = "I"
  zRPTARQ = PegPath("PATH", "INT") 'agora no ini+ "RPTINT.MDB"
  escRPTGRP.Show vbModal, Me
End Sub

Private Sub Menu6_Click()
  FrmLstView.Show
End Sub

Private Sub Menu8_Click()
  Unload Me
End Sub

Private Sub mnuSUBMENU1_Click(Index As Integer)
  Select Case Index
  Case 0
    cTIPO = "R"
    escmanreg.Show vbModal, Me
  Case 1
    cTIPO = "P"
    escmanreg.Show vbModal, Me
  End Select
End Sub

Private Sub mnuSUBMENU4_Click(Index As Integer)

  If Not AcessaMenu("mnuSUBMENU4", Index) Then Exit Sub
  ''alert CStr(Index)
  Select Case Index

  Case 0

    FrmPrintSetup.Show vbModal, Me

  Case 1
    cARQRTF = PegPath("HELP", "CAMINHO") + "WRPT.txt"
    ' 0-Novo,1-Abrir,2-Salvar,3-Salvacomo
    ' 4-Editar,5-Visualizar,6-Imprimir,7-Exportar
    aDIREITOS = Array(False, False, False, False, True, True, True, False)
    FrmRTf.Show vbModal, Me

  Case 2
    escuser.Show vbModal, Me


  Case 3
    escMENU.Show vbModal, Me

  Case 4
    escBTN.Show vbModal

  Case 5
    configuraodbc
  Case 6
    zIDTEMP = zUSERID
    frmUSUSENHA.Show vbModal

  Case 7
    frmDica.Show

  Case 8
    'Calculadora.Show
    ShellEx "calc"

  End Select




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


Private Sub MenuEvents_MenuHelp(ByVal MenuText As String, ByVal MenuHelp As String, ByVal Enabled As Boolean)
  If gbSubClassMenu Then
    '/ this event show Decription menu-item to a StatusBar control
    '/ customize for your project, if you need this.
    '/ Note: MenuText contain the menu Caption.
    '/       MenuHelp contain the Description string.
    '/ example: sbStatusBar.Panels(1).text = MenuHelp
    '/          lblMenuHelp.Caption = MenuHelp
  End If
End Sub
Public Sub SubClassMenuXP()

'/ this code is made by MenuCreator add-in

'/ prepare the caption for subclassing. Warning! Don't remove this comment!!!
  Menu1.Caption = "&Gerador"
  mnuSUBMENU1(0).Caption = "Especificos"
  mnuSUBMENU1(1).Caption = "Padrao"
  Menu2.Caption = "&Relatórios"
  Menu3.Caption = "&Documentos/Manuais"
  Menu4.Caption = "&Folha/Pto/RH"
  Menu5.Caption = "&Integrados"
  Menu6.Caption = "&Lst(Logix)"
  Menu7.Caption = "&Configurações"
  mnuSUBMENU4(0).Caption = "Impressoras|#printcfg"
  mnuSUBMENU4(1).Caption = "Alteracoes|#info"
  mnuSUBMENU4(2).Caption = "Usuarios|#user"
  mnuSUBMENU4(3).Caption = "Menus|#menusa"
  mnuSUBMENU4(4).Caption = "Botoes|#todo"
  mnuSUBMENU4(5).Caption = "Configurar ODBC|#odbcinfo"
  mnuSUBMENU4(6).Caption = "Troca Senha|#usuerchave"
  mnuSUBMENU4(7).Caption = "Frase do Dia|#lampada"
  mnuSUBMENU4(8).Caption = "Calculadora|#calc"
  Menu8.Caption = "&Sair"

  '/ Subclassing menu. Warning! Don't remove this comment!!!

  Set MenuEvents = New CEvents
  Set objMenuEx = New cMenuEx
  Call objMenuEx.Install(Me.hWnd, App.Path & "\" & Me.Name, ImgMenu, 2, MenuEvents)

End Sub

Public Sub MenuDesigner()
'/ Open Menu Designer tool
  objMenuEx.MenuDesigner Me.hWnd
End Sub
