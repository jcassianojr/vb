VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmPRINCIPAL 
   BackColor       =   &H8000000A&
   Caption         =   " "
   ClientHeight    =   3990
   ClientLeft      =   1575
   ClientTop       =   1890
   ClientWidth     =   10530
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
            Picture         =   "Principa.frx":0000
            TextSave        =   "12:18"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            Object.Width           =   2302
            MinWidth        =   2293
            Picture         =   "Principa.frx":059A
            TextSave        =   "15/03/2021"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Picture         =   "Principa.frx":0B34
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8281
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuMENU 
      Caption         =   "&Gerador"
      Index           =   0
      Begin VB.Menu mnuSUBMENU1 
         Caption         =   "Especificos"
         Index           =   0
      End
      Begin VB.Menu mnuSUBMENU1 
         Caption         =   "Padrao"
         Index           =   1
      End
   End
   Begin VB.Menu mnuMENU 
      Caption         =   "&Relatórios"
      Index           =   2
   End
   Begin VB.Menu mnuMENU 
      Caption         =   "&Documentos/Manuais"
      Index           =   3
   End
   Begin VB.Menu mnuMENU 
      Caption         =   "&Folha/Pto/RH"
      Index           =   4
   End
   Begin VB.Menu mnuMENU 
      Caption         =   "&Integrados"
      Index           =   5
   End
   Begin VB.Menu mnuMENU 
      Caption         =   "&Lst(Logix)"
      Index           =   6
   End
   Begin VB.Menu mnuMENU 
      Caption         =   "&Configurações"
      Index           =   7
      Begin VB.Menu mnuSUBMENU4 
         Caption         =   "Impressoras"
         Index           =   0
      End
      Begin VB.Menu mnuSUBMENU4 
         Caption         =   "Alteracoes"
         Index           =   1
      End
      Begin VB.Menu mnuSUBMENU4 
         Caption         =   "Usuarios"
         Index           =   2
      End
      Begin VB.Menu mnuSUBMENU4 
         Caption         =   "Menus"
         Index           =   3
      End
      Begin VB.Menu mnuSUBMENU4 
         Caption         =   "Botoes"
         Index           =   4
      End
      Begin VB.Menu mnuSUBMENU4 
         Caption         =   "Configurar ODBC"
         Index           =   5
      End
      Begin VB.Menu mnuSUBMENU4 
         Caption         =   "Troca Senha"
         Index           =   6
      End
      Begin VB.Menu mnuSUBMENU4 
         Caption         =   "Frase do Dia"
         Index           =   7
      End
      Begin VB.Menu mnuSUBMENU4 
         Caption         =   "Calculadora"
         Index           =   8
      End
   End
   Begin VB.Menu mnuMENU 
      Caption         =   "&Sair"
      Index           =   8
   End
End
Attribute VB_Name = "frmPRINCIPAL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub MDIForm_Load()
Dim CLINHA As String
Dim pICONES          As String
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
    
    Center Me

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
    Dbname = PegPath("PATH", "MDB") + "WRPT.MDB"
   
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
    cARQERRO = PegPath("PATH", "LOGCAM") & "LOG" & Format(Now, "MMYYYY") & ".MDB"
    If Not FileExist(cARQERRO) Then
        If Not FileExist(PegPath("PATH", "LOG"), True) Then
            Alert ("Erro Ini [PATH] LOG=")
            End
        End If
        nRETU = CopyFileWindowsWay(PegPath("PATH", "LOG"), cARQERRO)
        If nRETU <> 0 Then
            Alert "Erro Criando Arquivo Log do Mes"
            End
        End If
    End If
    
   
    zEMPRESA = 1
    zMES = Month(Date)
    zANO = Year(Date)
   
    ZGRPSUB = ""

    FrmSplash.Show vbModal

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

    

    DAORS.Open "select * from controle where form='frmmdiprincipal'", DAODB, adOpenForwardOnly, adLockReadOnly
    
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
    
        
       
    StatusBar1.Panels(6).Text = zUSER
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

    Unload Me
    End
End Sub
Private Sub mnuMENU_Click(Index As Integer)
Dim MenuClick        As String
    MenuClick = Left(UCase(Replace(mnuMENU(Index).Caption, "&", "")), 4)
    Select Case MenuClick
        Case "LST("
           FrmLstView.Show
        Case "SAIR"
            Unload Me
        Case "DOCU"
            cTIPORPT = "D"
            zRPTARQ = PegPath("PATH", "RTF") + "RTF.MDB"
            escRPTGRP.Show vbModal, Me
        Case "RELA"
            cTIPORPT = "R"
            zRPTARQ = PegPath("PATH", "RPT") + "RPT.MDB"
            escRPTGRP.Show vbModal, Me
        Case "INTE"
            cTIPORPT = "I"
            zRPTARQ = PegPath("PATH", "INT") + "RPTINT.MDB"
            escRPTGRP.Show vbModal, Me
        Case "FOLH"
            cTIPORPT = "F"
            zRPTARQ = PegPath("PATH", "FOL") + "RPTFOL.MDB"
            escRPTGRP.Show vbModal, Me
    End Select
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
           Calculadora.Show
        
    End Select




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

