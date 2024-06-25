VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{7020C36F-09FC-41FE-B822-CDE6FBB321EB}#1.2#0"; "vbccr18.ocx"
Begin VB.Form frmSENHA 
   Caption         =   "Acesso ao Sistema"
   ClientHeight    =   2220
   ClientLeft      =   2610
   ClientTop       =   3825
   ClientWidth     =   4830
   Icon            =   "frmSenha.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2220
   ScaleMode       =   0  'User
   ScaleWidth      =   4830
   Begin vbccr18.ProgressBar barra 
      Height          =   255
      Left            =   840
      Top             =   1800
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   450
   End
   Begin XPControls.XPButton cmdOK 
      Height          =   495
      Left            =   3480
      TabIndex        =   11
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Picture         =   "frmSenha.frx":038A
      Caption         =   "acessar"
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
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   0
      Picture         =   "frmSenha.frx":0724
      ScaleHeight     =   2175
      ScaleWidth      =   405
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   405
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.TextBox txtSENHA 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1680
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   420
      Width           =   1695
   End
   Begin VB.TextBox txtUSUARIO 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1680
      MaxLength       =   10
      TabIndex        =   0
      Top             =   60
      Width           =   1695
   End
   Begin vbccr18.SpinBox txtempresa 
      Height          =   495
      Left            =   840
      TabIndex        =   8
      Top             =   1080
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Min             =   1
      Max             =   999
      Value           =   1
      AllowOnlyNumbers=   -1  'True
   End
   Begin vbccr18.SpinBox mes 
      Height          =   495
      Left            =   2160
      TabIndex        =   9
      Top             =   1080
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Min             =   1
      Max             =   12
      Value           =   1
   End
   Begin vbccr18.SpinBox ano 
      Height          =   495
      Left            =   3360
      TabIndex        =   10
      Top             =   1080
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Min             =   1950
      Max             =   2200
      Value           =   1950
      ThousandsSeparator=   0   'False
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   660
      TabIndex        =   7
      Top             =   60
      Width           =   810
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Senha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   720
      TabIndex        =   6
      Top             =   360
      Width           =   810
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Empresa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   720
      TabIndex        =   5
      Top             =   780
      Width           =   810
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Mês"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   1980
      TabIndex        =   3
      Top             =   780
      Width           =   810
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Ano"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   3240
      TabIndex        =   2
      Top             =   780
      Width           =   810
   End
End
Attribute VB_Name = "frmSENHA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'Dim lOK As Integer
Const nFORMID = 1070
Const cFORMID = "Acesso ao Sistema"
Dim nTENTA As Integer
Dim nTEMPO As Integer

Private Sub ano_GotFocus()
    FocusMe
End Sub

Private Sub ano_KeyPress(KeyAscii As Integer)
    KeyAscii = ValiText(KeyAscii, "#NI")
End Sub

Private Sub cmdOK_Click()
    Dim cARQ As String
    Dim cSQL As String
    Dim cSENHA As String
    Dim cSENHA2 As String
    Dim cSENHA3 As String
    Dim cSENHA4 As String
    Dim aRETU As Variant
    Dim USALX As String
    Dim cCHAVE As String
    Dim cSENHAPOS As String

    On Error GoTo errhandler

    USALX = PegPath("PATH", "USALX")
    
    
    
    If UCase(txtUSUARIO) = "ADMLOG" Or UCase(txtUSUARIO) = "ADMINISTRADOR" Or UCase(txtUSUARIO) = "SUPERVISOR" Then
       txtUSUARIO = "ADMIN"
    End If
    
    
    ''Armazena Codigo empresa
1   zEMPRESA = FixInt(TxtEmpresa.Value)
2   If zEMPRESA < 0 Then zEMPRESA = 1
    
    
3   cARQ = PegPath("PATH", "SYSUSER")
4   zUSER = FVar(txtUSUARIO, "CU", "")
5   cSENHA = FVar(txtSENHA, "CU", "")
6   cSENHA = XOREncryption(SysCodeKey, cSENHA)
7   cSENHA3 = XOREncryption(SysCodeKey, txtSENHA) ''Senhas Antigas maisculas minusculas
8   cSENHA4 = XOREncryption(SysCodeKey, LCase(txtSENHA)) ''Senhas Antigas minusculas
    
9   If Len(zUSER) = 0 Then
10      TimedMsgBox "Nome Usuario Não Preenchido" 'Alert ("Nome Usuario Não Preenchido")
11      txtUSUARIO.SetFocus
12      If nTENTA >= 3 Then
13          End
14      Else
15          nTENTA = nTENTA + 1
16          Exit Sub
17      End If
18  End If
    
19  cSENHA2 = Format(Date, "ddmmyyyy") + UCase(txtUSUARIO)
20  cSENHA2 = XOREncryption(SysCodeKey, cSENHA2)
    
    
21  cSQL = "select * from USUARIO WHERE USUARIO='" & zUSER & "'"
        
22  aRETU = PegSQL(cARQ, cSQL, 15, _
                   Array("SENHA", "ATIVO", "DATAVAL", "WEEKEND", "HORAINI", _
                      "HORAFIM", "EQUIVALENTE", "IDUSUARIO", "IDFOLHA", "NOMEFOLHA", _
                      "TROCAR", "ID", "POSTELAB", "CHAVEH", "CHAVEV"), _
                   Array("C", "BF", "DN", "BF", "", "", _
                          "", "N", "N", "C", "DH", "N", "C", "C", "C"), _
                   Array(Space(8), False, NullDate(), False, Null, Null, _
                          "", 0, 0, "", Today(), 0, "", "", ""))
                
    ' ID DO USUARIO NO SISTEMA
23  zUSERID = aRETU(7)
    zUSERCHV = aRETU(12)
    
    
24  If Not lRETU Then
25      If UCase(txtUSUARIO) = "ADMIN" Or UCase(txtUSUARIO) = "ADMINISTRADOR" Or UCase(txtUSUARIO) = "SUPERVISOR" Then
            ''cria admin caso nao exista
26          cARQ = PegPath("PATH", "SYSUSER")
27          zIDTEMP = FixInt(PegMAXSQL(cARQ, "USUARIO", "IDUSUARIO", 0)) + 1
28          ePASS01 = "select * from USUARIO WHERE IDUSUARIO=" & zIDTEMP
29          IncluiSQL cARQ, ePASS01, 4, Array("IDUSUARIO", "USUARIO", "ATIVO", "DATAVAL"), _
        Array(zIDTEMP, UCase(txtUSUARIO), True, Today() + 90), True
30          frmUSUSENHA.Show vbModal, Me
31          TimedMsgBox "Necessario Reiniciar Sistema" 'Alert ("Necessario Reiniciar Sistema")
32          End
33      End If
34      TimedMsgBox "Usuario Nao Cadastrado" 'Alert ("Usuario Nao Cadastrado")
35      If nTENTA >= 3 Then
36          End
37      Else
38          nTENTA = nTENTA + 1
39          Exit Sub
40      End If
41  End If
42  If Len(aRETU(0)) = 0 And UCase(txtUSUARIO) <> "ADMIN" Then
43      TimedMsgBox " Necessário Cadastrar sua Senha" 'Alert (" Necessário Cadastrar de Senha")
44      zIDTEMP = aRETU(7)
45      frmUSUSENHA.Show vbModal, Me
46      TimedMsgBox "Necessario Reiniciar Sistema" 'Alert ("Necessario Reiniciar Sistema")
47      End
48  End If
    
   cCHAVE = UCase(CreateSHA256HashString(UCase(Trim(txtUSUARIO)) + UCase(Trim(txtSENHA))))
 
 '   If aRETU(0) = cSENHA Then
'       Alert ("Senha")
'   End If
'       Alert ("senha 2")
'   End If
''      If aRETU(0) = cSENHA3 Then
'       Alert ("senha 3")
'   End If
'   If aRETU(0) = cSENHA4 Then
'       Alert ("senha 4")
'   End If
'   If aRETU(13) = cCHAVE Then
'       Alert ("HASH H")
'   End If
'   If aRETU(14) = cCHAVE Then
'       Alert ("HASH V")
'   End If
cSENHAPOS = senhapos(zUSERCHV)
        
49  If aRETU(0) = cSENHA Or _
       txtSENHA = cSENHA2 Or _
       aRETU(0) = cSENHA3 Or _
       aRETU(0) = cSENHA4 Or _
       cSENHAPOS = cSENHA Or _
       cSENHAPOS = cSENHA2 Or _
       cSENHAPOS = cSENHA3 Or _
       cSENHAPOS = cSENHA4 Or _
       aRETU(13) = cCHAVE Or _
       aRETU(14) = cCHAVE Then
    Else
50      If nTENTA >= 3 Then
51          TimedMsgBox "Usuário - numero de tentativas esgotadas"
52          End
53      Else
54          nTENTA = nTENTA + 1
55          txtUSUARIO.SetFocus
56          TimedMsgBox "Senha Incorreta", 1
57          Exit Sub
58      End If
59  End If

        
60  If zUSER <> "ADMIN" And zUSER <> "SUPERVISOR" Then
            
61      If aRETU(8) = 0 Then
62          TimedMsgBox "No Folha de Pagamento Nao Preenchido" 'Alert ("No Folha de Pagamento Nao Preenchido")
63          End
64      End If
           
65      If Len(aRETU(9)) = 0 Then
66          TimedMsgBox "Nome Folha de Pagamento Nao Preenchido"  'Alert ("Nome Folha de Pagamento Nao Preenchido")
67          End
68      End If
    
69      If Not aRETU(1) Then
70          TimedMsgBox "Usuario nao esta Ativo" 'Alert("Usuario nao esta Ativo")
71          End
72      End If
73      If IsNull(aRETU(2)) Then
74          TimedMsgBox "Data Limite Acesso Nao Preenchida" 'Alert("Data Limite Acesso Nao Preenchida")
75          End
76      End If
    
77      If aRETU(2) < Today() Then
78          TimedMsgBox "Data Limite Acesso Expirada" 'Alert("Data Limite Acesso Expirada")
79          End
80      End If
81      If (Weekday(Date) = vbSaturday Or Weekday(Date) = vbSunday) _
        And Not aRETU(3) Then
82          TimedMsgBox "Acesso Final Semana Bloqueado" 'Alert("Acesso Final Semana Bloqueado")
83          End
84      End If
85      If Not IsNull(aRETU(4)) And Not IsNull(aRETU(5)) Then
86          If DateValue(Format(Date, "HH:MM:SS")) < DateValue(aRETU(4)) Or _
                                                                         DateValue(Format(Date, "HH:MM:SS")) > DateValue(aRETU(5)) Then
87              TimedMsgBox "Horario nao Autorizado" 'Alert("Horario nao Autorizado")
88              End
89          End If
90      End If
91  End If
92  zUSERID = aRETU(7)
93  zWRPTID = aRETU(7)
94  zIDFOLHA = aRETU(8)
95  zNOMEFOLHA = aRETU(9)
       
       
96  GrvSQL cARQ, cSQL, 1, Array("DATAULT"), Array(Today()), Array("D")
       
       
97  If aRETU(10) < Today() Then
98      TimedMsgBox " Necessario Troca de Senha"
99      zIDTEMP = zUSERID
100     frmUSUSENHA.Show vbModal, Me
101     End
102 End If

103 Unload Me
    
    Exit Sub
errhandler:
    Select Case Err.Number
    Case Else
        SayErro "FrmSenha"
        End
    End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    TeclaEnter KeyCode
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    zMES = mes.Value
    zANO = ano.Value
    zEMPRESA = TxtEmpresa.Value
End Sub

'Private Sub maisemp_Click()
'    txtempresa = CStr(IncDec(txtempresa, 1, 1, 999))
'End Sub

'Private Sub menosemp_Click()
'    txtempresa = CStr(IncDec(txtempresa, -1, 1, 999))
'End Sub

Private Sub Form_Load()
    'CentralizaJanela Me
    CenterFormToScreen Me
    EnableCloseButton Me.hWnd, False
    txtUSUARIO = NetworkUserName()
    nTENTA = 0
    TxtEmpresa.Value = 1
    mes.Value = Month(Date)
    ano.Value = Year(Date)
    nTEMPO = 0
End Sub

Private Sub mes_GotFocus()
    FocusMe
End Sub

Private Sub mes_KeyPress(KeyAscii As Integer)
    KeyAscii = ValiText(KeyAscii, "#NI")
End Sub

Private Sub Timer1_Timer()
    nTEMPO = nTEMPO + 1
    If nTEMPO >= 61 Then
        End
    End If
    If nTENTA >= 3 Then
       End
    End If
    Barra.Value = nTEMPO
End Sub

Private Sub TXTempresa_GotFocus()
    FocusMe
End Sub

Private Sub TxtEmpresa_KeyPress(KeyAscii As Integer)
    KeyAscii = ValiText(KeyAscii, "#NI")
End Sub

Private Sub TXTempresa_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Or KeyCode = 40 Then
        SendKeys Chr(9)
    End If
    If KeyCode = 38 Then
        SendKeys "+" + Chr(9)
    End If
End Sub

Private Sub txtSENHA_GotFocus()
    fixuser
    FocusMe
End Sub

Private Sub txtSENHA_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then

        cmdOK.SetFocus

    End If

End Sub

Private Sub txtSENHA_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Or KeyCode = 40 Then
        SendKeys Chr(9)
    End If
    If KeyCode = 38 Then
        SendKeys "+" + Chr(9)
    End If

End Sub

Private Sub txtUSUARIO_GotFocus()
    FocusMe
End Sub

Private Sub txtUSUARIO_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Or KeyCode = 40 Then
        SendKeys Chr(9)
    End If
    If KeyCode = 38 Then
        SendKeys "+" + Chr(9)
    End If
End Sub

'Private Sub maisano_Click()
'    ano = CStr(IncDec(ano, 1, 1900, 2300))
'End Sub

'Private Sub maismes_Click()
'    mes = CStr(IncDec(mes, 1, 1, 12))
'End Sub

'Private Sub menosano_Click()
'    ano = CStr(IncDec(ano, -1, 1900, 2300))
'End Sub

'Private Sub menosmes_Click()
'    mes = CStr(IncDec(mes, -1, 1, 12))
'End Sub

Private Sub fixuser()
    If UCase(txtUSUARIO.tEXT) = "ADMLOG" Then
        txtUSUARIO.tEXT = "admin"
    End If
    If UCase(txtUSUARIO.tEXT) = "ADMINISTRA" Then
        txtUSUARIO.tEXT = "admin"
    End If
    If UCase(txtUSUARIO.tEXT) = "SUPERVISOR" Then
        txtUSUARIO.tEXT = "admin"
    End If

End Sub

Private Sub txtUSUARIO_LostFocus()
    fixuser
End Sub


Private Sub XPButton1_Click()

End Sub
