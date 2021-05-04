VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{7B4BC5ED-76A0-4FFB-9D26-6EE54C082750}#2.5#0"; "CoolXPMenu.ocx"
Begin VB.MDIForm frmPRINCIPAL 
   BackColor       =   &H8000000A&
   Caption         =   " "
   ClientHeight    =   6390
   ClientLeft      =   1575
   ClientTop       =   1890
   ClientWidth     =   10875
   Icon            =   "Principa.frx":0000
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin CoolXPMenu.xpMenu xpMenu1 
      Left            =   4200
      Top             =   2340
      _ExtentX        =   900
      _ExtentY        =   900
      BmpCount        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   780
      Top             =   60
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
      Top             =   6075
      Width           =   10875
      _ExtentX        =   19182
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
            TextSave        =   "11:10"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            Object.Width           =   2302
            MinWidth        =   2293
            Picture         =   "Principa.frx":0B24
            TextSave        =   "04/05/2021"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Picture         =   "Principa.frx":10BE
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8890
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
    Dim CLINHA As String
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
  '  Dim nPOS As Integer
    
   'Dim lngReturn As Long
'      SetSysColors 1, COLOR_ACTIVECAPTION, RGB(191, 230, 255)
'      SetSysColors 1, COLOR_CAPTIONTEXT, RGB(191, 230, 255)
      
      
      
    StatusBar1.Panels(5).tEXT = ""

    
'carqhelp = "SELECT FEMEAEF FROM"
'carqhelp = UCase(carqhelp)
'If InStr(carqhelp, "SELECT") > 0 And InStr(carqhelp, "FROM") > 0 Then
'   carqhelp = Replace(carqhelp, "SELECT ", "")
'   carqhelp = Mid(carqhelp, 1, InStr(carqhelp, "FROM") - 1)
'   carqhelp = Trim(carqhelp)
'End If
    'SayErro "abrindo"

    CenterFormToScreen Me

    On Error GoTo ErrorHandler
    
    'SayErro "instance"

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
    zRPTARQ = PegPath("PATH", "RPT") + "RPT.MDB"
    zusalx = PegPath("PATH", "USALX")
    zusamc = PegPath("PATH", "USAMC")
    cTIPORPT = "R"
    
    carqhelp = PegPath("PATH", "HELP")
    If FileExist(carqhelp) Then
       App.HelpFile = carqhelp
    End If
    
    'Cria o Arquivo log do mes
    cARQERRO = PegPath("PATH", "LOGCAM") & "LOG" & Format(Now, "MMYYYY") & ".MDB"
    If Not FileExist(cARQERRO) Then
        If Not FileExist(PegPath("PATH", "LOG"), True) Then
           Alert PegPath("PATH", "LOG")
           End
        End If
        nRETU = CopyFileWindowsWay(PegPath("PATH", "LOG"), cARQERRO)
        If nRETU = 0 Then
            Alert "Erro Criando Arquivo Log do Mes"
            End
        End If
    End If
    Dbname = PegPath("PATH", "MDB")
    zMDBPATH = Dbname
    ZGRPSUB = ""
    
    
    'FrmSplash.Show vbModal
    zANO = Year(Date)
    zMES = Month(Date)
    cmdline = Trim(Command())
    bACESSO = False
    If InStr(cmdline, "__$") > 0 And InStr(cmdline, "%__") > 0 Then
        nPOS = InStr(cmdline, "__$")
        cmdline = Mid(cmdline, 4)
        nPOS = InStr(cmdline, "%__")
        cmdline = UCase(Left(cmdline, Len(cmdline) - 3))
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
            zEMPRESA = 1
            zANO = Year(Date)
            zMES = Month(Date)
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
    
    'opcao somente via menu
    'CLINHA = "MAIL $" & UCase(zUSER)
    'Shell CLINHA, vbNormalFocus
    'ShellEx "MAIL", essSW_SHOWDEFAULT, "$" & UCase(zUSER),PegPath("PATH", "MAIL") , , Me.hWnd
    
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
    
    StatusBar1.Panels(6).tEXT = zUSER
    If Trim(PegPath("CITACAO", zUSER, "S")) = "S" Then
        frmDica.Show
    End If
    
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
    zRPTARQ = PegPath("PATH", "RPT") + "RPT.MDB"
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
        ESCCLI.Show vbModal
    Case 8
        ESCpro.Show vbModal
    Case 9
        iPPAP = 3
        escPPAP.Show vbModal
    Case 10
        ZENGTIP = "PRO"
        ESCCLI.Show vbModal
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
    Dim CLINHA As String
    
    ''nao verifica acesso todos entram
    ''os que necessitam case abaixo
    Select Case Index
    
    Case 0
        FrmPrintSetup.Show vbModal
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
        'Calculadora.Show
        ShellEx "calc"
        Exit Sub
    Case 9
        ''CLINHA = "MAIL $" & UCase(zUSER)
        ''Shell CLINHA, vbNormalFocus
        ShellEx "MAIL", essSW_SHOWDEFAULT, "$" & UCase(zUSER), PegPath("PATH", "MAIL"), , Me.hWnd
        Exit Sub
    End Select
    

    If Not AcessaMenu("mnuSUBMENU4", Index) Then Exit Sub
    Select Case Index
    Case 1
        cARQRTF = PegPath("HELP", "HELPTXT") '+ "controle.txt"
        If FileExist(cARQRTF, True) Then
            ' 0-Novo,1-Abrir,2-Salvar,3-Salvacomo
            ' 4-Editar,5-Visualizar,6-Imprimir,7-Exportar
'            aDIREITOS = Array(False, False, False, False, True, True, True, False)
'            FrmRTf.Show
             PrinterEx.ShowPrintPreview Me, "MyPrintingTXT"
        End If
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
    Dim cSQL, cARQORI, cARQDES As String
    Dim aCAM As Variant
    Dim X As Integer


    aCAM = Array("EQUIVALENTE", "DATAULT", "IDFOLHA", "NOMEFOLHA", "HORAINI", "HORAFIM", _
                 "ATIVO", "WEEKEND", "DATAVAL", "TROCAR", "USUARIO")

    cARQORI = PegPath("PATH", "WRPT")            '"p:\novell\ITAESBRA\WRPT\WRPT.MDB"
    cARQDES = PegPath("PATH", "CONTROLE")        '"p:\novell\ITAESBRA\CONTROLE\CONTROLE.MDB"
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
    zRPTARQ = PegPath("PATH", "RPT") + "RPT.MDB"
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
        zRPTARQ = PegPath("PATH", "RPT") + "RPT.MDB"
        escRPTGRP.Show vbModal, Me
    Case 1
        cTIPORPT = "D"
        zRPTARQ = PegPath("PATH", "RTF") + "RTF.MDB"
        escRPTGRP.Show vbModal, Me
    Case 2
        cTIPO = "R"
        escmanreg.Show vbModal, Me
    Case 3
        cTIPO = "P"
        escmanreg.Show vbModal, Me
    Case 4
        cTIPO = "F"
        zRPTARQ = PegPath("PATH", "FOL") + "RPTFOL.MDB"
        escRPTGRP.Show vbModal, Me
    Case 5
        cTIPO = "I"
        zRPTARQ = PegPath("PATH", "INT") + "RPTINT.MDB"
        escRPTGRP.Show vbModal, Me
    Case 6
        FrmLstView.Show vbModal, Me
    Case 7
        'zRPTARQ = PegPath("PATH", "WRPTX") + " $" + Trim(zUSER) + "%#"
        'Shell zRPTARQ, vbNormalFocus
        ShellEx "WRPTX", essSW_SHOWDEFAULT, "$" + Trim(zUSER) + "%#", PegPath("PATH", "WRPTX"), , Me.hWnd
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault
End Sub

Private Sub configuraodbc()
    Dim lCONTINUA As Boolean
    Dim cCFG As String
    Dim aDAD As Variant
    Dim X As Integer
    lCONTINUA = True
    X = 1
    While lCONTINUA
        cCFG = FixStr(PegPath("ODBCCFG", "ODBCCFG" & StrZero(X, 3)))
        If InStr(cCFG, "|") = 0 Then
            lCONTINUA = False
        Else
            aDAD = Split(cCFG, "|")
            AddDSN aDAD(0), aDAD(1), aDAD(2), aDAD(3), True
        End If
        X = X + 1
    Wend
End Sub
Public Property Get Printer() As Printer
    Set Printer = vbExtra.Printer2
End Property
Public Property Set Printer(nPrinter As Printer)
    Set vbExtra.Printer2 = nPrinter
End Property
Public Sub MyPrintingTXT()
    Dim fileFile As Integer
    Dim STRBUFFER As String
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

