VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{379157C5-E9BD-43F1-9F83-B037496BED42}#1.3#0"; "vbccr18.ocx"
Begin VB.MDIForm frmPRINCIPAL 
   BackColor       =   &H8000000A&
   Caption         =   " Menu Inicial"
   ClientHeight    =   3984
   ClientLeft      =   1572
   ClientTop       =   1896
   ClientWidth     =   10536
   Icon            =   "Principa.frx":0000
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   780
      Top             =   180
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImgMenu 
      Left            =   5850
      Top             =   180
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
            Picture         =   "Principa.frx":058A
            Key             =   "center"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":0924
            Key             =   "justificado"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":0CBE
            Key             =   "left"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":1058
            Key             =   "right"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":13F2
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":178C
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":1D26
            Key             =   "cut"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":20C0
            Key             =   "diahora"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":245A
            Key             =   "editor"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":29F4
            Key             =   "find"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":2D8E
            Key             =   "marcado"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":2EE8
            Key             =   "letra01"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":3282
            Key             =   "tamanho"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":361C
            Key             =   "paragro"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":39B6
            Key             =   "decrese"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":3D50
            Key             =   "incrise"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":40EA
            Key             =   "imagem"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":4484
            Key             =   "italico"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":481E
            Key             =   "s_Key19"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":4BB8
            Key             =   "mail"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":5152
            Key             =   "new"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":54EC
            Key             =   "open"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":5A86
            Key             =   "xxx"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":5E20
            Key             =   "imagem2"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":61BA
            Key             =   "colar"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":6754
            Key             =   "printer"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":6CEE
            Key             =   "printcfg"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":7088
            Key             =   "sair"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":7422
            Key             =   "salvar"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":77BC
            Key             =   "busca"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":7B56
            Key             =   "proximabusca"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":7EF0
            Key             =   "letra"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":828A
            Key             =   "abc"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":8624
            Key             =   "riscado"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":89BE
            Key             =   "sub"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":8D58
            Key             =   "sob"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":90F2
            Key             =   "minusculas"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":948C
            Key             =   "maisculas"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":9826
            Key             =   "underline"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":9BC0
            Key             =   "refazer"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":9F5A
            Key             =   "info"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":A4F4
            Key             =   "exit"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":AA8E
            Key             =   "calc"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":B028
            Key             =   "menusa"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":B3C2
            Key             =   "usuerchave"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":B75C
            Key             =   "odbcinfo"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":B8B6
            Key             =   "todo"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":BD50
            Key             =   "lampada"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principa.frx":C1EA
            Key             =   "user"
         EndProperty
      EndProperty
   End
   Begin VBCCR18.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   372
      Left            =   0
      Top             =   3615
      Width           =   10530
      _ExtentX        =   18584
      _ExtentY        =   656
      InitPanels      =   "Principa.frx":C784
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
  Dim ctmpline As String


  On Error GoTo ErrorHandler

  CenterFormToScreen Me
  
 
  
  'cARQRTF = "C:/temp/citacao.pdf"
  'frmPDFView.Show vbModal
  'End
  
  ' teste frmrtf
 ' aDIREITOS = Array(True, True, True, True, True, True, True, True)
  ' FrmRTf.Show vbModal
  ' End
   
   'teste frmprevies
   'cARQRTF = "d:\temp\xharbour320.avi"
   'FrmPreview.Show vbModal, Me
   'End
   

 ' cARQRTF = "c:\temp\SQLRDD Manual - EN.pdf"
 '  frmPDFView.Show vbModal, Me
 ' End

  If App.PrevInstance Then

    MsgBox App.EXEName & " já esta rodando!", 48
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
  Dbname = PegPath("PATH", "MDB")

  If Not FileConnExist(Dbname, True) Then
    Alert ("Erro Ini [PATH] MDB=")
    End
  End If
  If Not FileConnExist(dbuser, True) Then
    Alert ("Erro Ini [PATH] USER=")
    End
  End If
  If Not FileConnExist(DBWRPT, True) Then
    Alert ("Erro Ini [PATH] WRPT=")
    End
  End If

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

  Set DAODB = New ADODB.Connection
  Set DAORS = New ADODB.Recordset

    'Cria o Arquivo log do mes
  cARQERRO = PegPath("PATH", "LOGCAM") & "LOG" & Format(Now, "MMYYYY") & ".sqlite"
  If Not FileExists(cARQERRO) Then
     CopyFileWindowsWay PegPath("PATH", "LOG"), cARQERRO, True
  End If
  If Not FileExists(cARQERRO) Then
      Alert "Erro Criando Arquivo Log do Mes"
      End
  End If

  zEMPRESA = 1
  zMES = Month(Date)
  zANO = Year(Date)
  ZGRPSUB = ""
  
  zgrp = ""
  ZGRPSUB = ""
  zUSER = ""
  eLOCALIZA = ""
  cARQRTF = ""
  
  
  
  
  
   'pegando o usuario
   cmdline = Command()
   cmdline = Replace(cmdline, "__", "")
   
  If InStr(UCase(cmdline), "LSTVIEW") Then
     FrmLstView.Show vbModal, Me
     End
  End If
  
   If InStr(UCase(cmdline), ".ZPL") _
     Or InStr(UCase(cmdline), ".JPG") Then
     cARQRTF = cmdline
     escRPT.Show vbModal, Me
     End
  End If
  
  If InStr(UCase(cmdline), ".TXT") Or InStr(UCase(cmdline), ".MAN") _
     Or InStr(UCase(cmdline), ".RTF") Then
     cARQRTF = cmdline
     escRPT.Show vbModal, Me
     'criar opcao de preview direto
     'FrmRtfView.Show vbModal, Me
     End
  End If
   
   
   cTIPORPT = ""
   zRPTARQ = ""
   If Mid(cmdline, 1, 1) = "X" Or Mid(cmdline, 1, 1) = "R" Then
      cTIPORPT = "R"
      cmdline = Mid(cmdline, 2)
      zRPTARQ = PegPath("PATH", "RPT")
   End If
   If Mid(cmdline, 1, 1) = "D" Then
      cTIPORPT = "D"
      cmdline = Mid(cmdline, 2)
      zRPTARQ = PegPath("PATH", "DOC")
   End If
   If Mid(cmdline, 1, 1) = "I" Then
      cTIPORPT = "I"
      cmdline = Mid(cmdline, 2)
      zRPTARQ = PegPath("PATH", "INT")
   End If
   If Mid(cmdline, 1, 1) = "F" Then
      cTIPORPT = "F"
      cmdline = Mid(cmdline, 2)
      zRPTARQ = PegPath("PATH", "FOL")
   End If
   ''especifico ou padrao manrel mana5
   If Mid(cmdline, 1, 1) = "E" Then
      cTIPORPT = "E"
   End If
   If Mid(cmdline, 1, 1) = "P" Then
      cTIPORPT = "P"
   End If
   
   
  aRETU = pegue2delimitado(cmdline, "$", "%")
  zUSER = aRETU(0)
  ctmpline = aRETU(1)

  'pegando o grupo
  aRETU = pegue2delimitado(cmdline, "%", "#")
  zgrp = aRETU(0)
  ctmpline = aRETU(1)
  
   'pegando o subgrupo
  aRETU = pegue2delimitado(cmdline, "#", "_")
  ZGRPSUB = aRETU(0)
  ctmpline = aRETU(1)
   
   
  'pegando o relatorio
  aRETU = pegue2delimitado(cmdline, "_", "")
  eLOCALIZA = aRETU(0)
  ctmpline = aRETU(1)
  

  zUSER = UCase(zUSER)
  zgrp = UCase(zgrp)
  ZGRPSUB = UCase(ZGRPSUB)

  If zUSER = "SUPERVISOR" Then
    zUSER = "ADMIN"
  End If
  
  'aRETU = pegue2delimitado(Trim(Command()), "__$", "%__")
  
  
  bACESSO = False
  'dados do usuario
  zWRPTID = 0
  zUSERID = 0
  zIDFOLHA = 0
  zIDUNI = 0
  zNOMEFOLHA = ""
  If Len(zUSER) > 0 Then
    sSQL = "SELECT IDUSUARIO,USUARIO,IDFOLHA,NOMEFOLHA FROM USUARIO WHERE USUARIO='" & zUSER & "'"
    aRETU = PegSQL(dbuser, sSQL, 3, Array("IDUSUARIO", "IDFOLHA", "NOMEFOLHA"), _
                   Array("NI", "NI", "C"), _
                   Array(0, 0, ""))
    If lRETU Then
      zUSERID = aRETU(0)
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

  Me.Caption = App.FileDescription & " Versao:" & App.Major & "." & App.Minor & "Usuário: " & zUSER


  If Trim(PegPath("EMAILINT", zUSER, "S")) = "S" Then
     ShellEx "MAIL", essSW_SHOWDEFAULT, "$" & UCase(zUSER), PegPath("PATH", "MAIL"), , Me.hwnd
  End If


  '* Carrega imagens para o ImageList

  DAODB.CursorLocation = adUseClient
  DAODB.ConnectionTimeout = 120
  DAODB.Open GeracArq(Dbname, , False)



  DAORS.Open "select * from controle where form='frmMDIPRINCIPAL'", DAODB, adOpenForwardOnly, adLockReadOnly

  With DAORS
    Do While Not DAORS.EOF  '!controle = "ImageList1"
      cARQICO = pICONES & DAORS.Fields("TOOLTIP")
      If FileConnExist(cARQICO, True) Then
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
  
  
  'abre caso o tipo nao esteja em branco
  If cTIPORPT <> "" Then
     If cTIPORPT = "E" Or cTIPORPT = "P" Then
       'especificos ou padrao manrel mana5 imprel via shell
       If cTIPORPT = "E" Then
          mnuSUBMENU1_Click (0)
       End If
        If cTIPORPT = "P" Then
          mnuSUBMENU1_Click (1)
       End If
     Else
        If eLOCALIZA = "" Then
           If zgrp = "" Then
              cARQRTF = ""
              eLOCALIZA = ""
              escRPTGRP.Show vbModal
              End
           Else
              cARQRTF = ""
              eLOCALIZA = ""
              escRPT.Show vbModal
              End
           End If
        Else
            escRPT.Show vbModal
            End
        End If
      End If
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
    Call objMenuEx.Uninstall(Me.hwnd, ImgMenu, MenuEvents)
    Set MenuEvents = Nothing
    Set objMenuEx = Nothing
  End If


  Unload Me
  End
End Sub
Private Sub Menu2_Click()
  cARQRTF = ""
  cTIPORPT = "R"
  zRPTARQ = PegPath("PATH", "RPT") ' agora no ini + "RPT.MDB"
  escRPTGRP.Show vbModal, Me
End Sub

Private Sub Menu3_Click()
  cARQRTF = ""
  cTIPORPT = "D"
  zRPTARQ = PegPath("PATH", "DOC") ' agora no ini + "RTF.MDB"
  escRPTGRP.Show vbModal, Me
End Sub

Private Sub Menu4_Click()
  cARQRTF = ""
  cTIPORPT = "F"
  zRPTARQ = PegPath("PATH", "FOL")
  escRPTGRP.Show vbModal, Me
End Sub

Private Sub Menu5_Click()
  cARQRTF = ""
  cTIPORPT = "I"
  zRPTARQ = PegPath("PATH", "INT")
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
    cTIPO = "R"  'especifico
    escmanreg.Show vbModal, Me
  Case 1
    cTIPO = "P"  'padrao
    escmanreg.Show vbModal, Me
  End Select
End Sub

Private Sub mnuSUBMENU4_Click(Index As Integer)

  If Not AcessaMenu("mnuSUBMENU4", Index) Then Exit Sub
  ''alert CStr(Index)
  Select Case Index

  Case 0

   ' FrmPrintSetup.Show vbModal, Me

  Case 1
    cARQRTF = PegPath("HELP", "CAMINHO") + "WRPT.txt"
    ' 0-Novo,1-Abrir,2-Salvar,3-Salvacomo
    ' 4-Editar,5-Visualizar,6-Imprimir,7-Exportar
    aDIREITOS = Array(False, False, False, False, True, True, True, False)
    FrmRtfView.Show vbModal, Me

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
  Call objMenuEx.Install(Me.hwnd, App.Path & "\" & Me.Name, ImgMenu, 2, MenuEvents)

End Sub

Public Sub MenuDesigner()
'/ Open Menu Designer tool
  objMenuEx.MenuDesigner Me.hwnd
End Sub
