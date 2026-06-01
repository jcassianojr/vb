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
   Begin MSComctlLib.ImageList ImgMenu 
      Left            =   0
      Top             =   0
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
            Picture         =   "Principal.frx":058A
            Key             =   "center"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":0924
            Key             =   "justificado"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":0CBE
            Key             =   "left"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":1058
            Key             =   "right"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":13F2
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":178C
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":1D26
            Key             =   "cut"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":20C0
            Key             =   "diahora"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":245A
            Key             =   "editor"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":29F4
            Key             =   "find"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":2D8E
            Key             =   "marcado"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":2EE8
            Key             =   "letra01"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":3282
            Key             =   "tamanho"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":361C
            Key             =   "paragro"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":39B6
            Key             =   "decrese"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":3D50
            Key             =   "incrise"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":40EA
            Key             =   "imagem"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":4484
            Key             =   "italico"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":481E
            Key             =   "s_Key19"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":4BB8
            Key             =   "mail"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":5152
            Key             =   "new"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":54EC
            Key             =   "open"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":5A86
            Key             =   "xxx"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":5E20
            Key             =   "imagem2"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":61BA
            Key             =   "colar"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":6754
            Key             =   "printer"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":6CEE
            Key             =   "printcfg"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":7088
            Key             =   "sair"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":7422
            Key             =   "salvar"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":77BC
            Key             =   "busca"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":7B56
            Key             =   "proximabusca"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":7EF0
            Key             =   "letra"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":828A
            Key             =   "abc"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":8624
            Key             =   "riscado"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":89BE
            Key             =   "sub"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":8D58
            Key             =   "sob"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":90F2
            Key             =   "minusculas"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":948C
            Key             =   "maisculas"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":9826
            Key             =   "underline"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":9BC0
            Key             =   "refazer"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":9F5A
            Key             =   "info"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":A4F4
            Key             =   "exit"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":AA8E
            Key             =   "calc"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":B028
            Key             =   "menusa"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":B3C2
            Key             =   "usuerchave"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":B75C
            Key             =   "odbcinfo"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":B8B6
            Key             =   "todo"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":BD50
            Key             =   "lampada"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":C1EA
            Key             =   "user"
         EndProperty
      EndProperty
   End
   Begin VBCCR18.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   372
      Left            =   0
      Top             =   6012
      Width           =   10884
      _ExtentX        =   19198
      _ExtentY        =   656
      InitPanels      =   "Principal.frx":C784
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


20:   StatusBar1.Panels(5).text = ""



24:   CenterFormToScreen Me
  

27:   On Error GoTo ErrorHandler


30:   If App.PrevInstance Then
31:     Alert App.EXEName & " já está rodando!"
32:     End
33:   End If
34:   strCodeKey = "SoFtEc1SiStEmAs24055-2265"
35:   SysCodeKey = "SoFtEc1SiStEmAs24055-2265"
36:   pICONES = PegPath("PATH", "ICONES")
37:   zAppPath = PegPath("PATH", "APP")
38:   zMANA5EMP = PegPath("PATH", "MANA5EMP")
39:   dbuser = PegPath("PATH", "USER")
40:   DBWRPT = PegPath("PATH", "WRPT")
41:   ZMANA5CON = PegPath("PATH", "MANA5CON")
42:   ZMANA5IMP = PegPath("PATH", "IMPREL")
43:   ZRELAT = PegPath("PATH", "RELAT")
44:   zRPTARQ = PegPath("PATH", "RPT") 'agora no ini + "RPT.MDB"
45:   zusalx = PegPath("PATH", "USALX")
46:   zusamc = PegPath("PATH", "USAMC")
47:   cTIPORPT = "R"

 'se sql lite verifica se odbc esta instalado
 If InStr(LCase(dbuser), ".sqlite") > 0 Or InStr(LCase(dbuser), ".sqlite3") > 0 _
                 Or InStr(LCase(dbuser), ".fossil") > 0 Or InStr(LCase(dbuser), ".db3") > 0 _
                 Or (InStr(LCase(dbuser), ".db") > 0) Then
53:      If IsDriverInstalled("SQLite3 ODBC Driver") Then
       ' MsgBox "O driver SQLite ODBC está pronto para uso!", vbInformation
55:      Else
56:         MsgBox "Driver SQLite ODBC não encontrado. Por favor, instale o driver para continuar.", vbCritical
57:         End
58:      End If
59:  End If

61:   carqhelp = PegPath("PATH", "HELP")
62:   If FileConnExist(carqhelp) Then
63:     App.HelpFile = carqhelp
64:   End If

   'Cria o Arquivo log do mes
67:   cARQERRO = PegPath("PATH", "LOGCAM") & "LOG" & Format(Now, "MMYYYY") & ".sqlite"
68:   nRETU = CopyFileWindowsWay(PegPath("PATH", "LOG"), cARQERRO, False)
69:   If nRETU <> 0 Then
70:       Alert "Erro Criando Arquivo Log do Mes"
71:       End
72:   End If

74:   Dbname = PegPath("PATH", "MDB")
75:   zMDBPATH = Dbname
76:   ZGRPSUB = ""

78:   zANO = Year(Date)
79:   zMES = Month(Date)
80:   zEMPRESA = 1
  
82:   aRETU = pegue2delimitado(Trim(Command()), "__$", "%__")
  
84:   bACESSO = False
85:   If Len(aRETU(0)) > 0 Then
86:     cmdline = UCase(aRETU(0))
87:     sSQL = "SELECT IDUSUARIO,USUARIO,IDFOLHA,NOMEFOLHA FROM USUARIO WHERE USUARIO='" & cmdline & "'"
    aRETU = PegSQL(dbuser, sSQL, 3, Array("IDUSUARIO", "IDFOLHA", "NOMEFOLHA"), _
                   Array("NI", "NI", "C"), _
                   Array(0, 0, ""))
91:     If lRETU Then
92:       zUSERID = aRETU(0)
93:       zUSER = cmdline
94:       zWRPTID = zUSERID
95:       zIDFOLHA = aRETU(1)
96:       zNOMEFOLHA = aRETU(2)
97:       bACESSO = True
98:     End If
99:   End If

101:   If Not bACESSO Then
102:     frmSENHA.Show vbModal, Me
103:   End If

105:   If zIDFOLHA > 0 Then
106:     If demitido(zIDFOLHA) Then
107:       End
108:     End If
109:   End If




114:   Me.Caption = App.FileDescription & " Versao:" & App.Major & "." & App.Minor & " Usuário: " & zUSER
115:   If Not FillMenu("SUBMENU") Then
116:     Alert "O Sub-Menu não foi carregado!", "Sub Menu"
117:   End If
118:   If Not FillMenu("SUBMENU2") Then
119:     Alert "O Sub-Menu 2 não foi carregado!", "SubMenu 2"
120:   End If
  ''If Not FillMenu("SUBMENU3") Then
  ''  alert "O Sub-Menu 3 não foi carregado!",  "SubMenu 2"
  ''End If
124:   If Not FillMenu("SUBMENU4") Then
125:     Alert "O Sub-Menu 4 não foi carregado!", "SubMenu 4"
126:   End If
127:   If Not FillMenu("SUBMENU5") Then
128:     Alert "O Sub-Menu 5 não foi carregado!", "SubMenu 5"
129:   End If
  ''* Pega Identificacao do WRPT

132:   zWRPTID = 0
133:   sSQL = "SELECT IDUSUARIO FROM USUARIO WHERE USUARIO='" & zUSER & "'"
134:   aRETU = PegSQL(DBWRPT, sSQL, 1, Array("IDUSUARIO"), Array("NI"), Array(0))
135:   zWRPTID = aRETU(0)
136:   GravaLog 0, 0, "Entrar", "Sistema"

  '* Carrega imagens para o ImageList
139:   DB.ConnectionTimeout = 120
140:   DB.Open GeracArq(Dbname, , False)
141:   rs.Open "select * from controle where form='frmMDIPRINCIPAL' and controle='ImageList1'", DB, adOpenForwardOnly, adLockReadOnly
142:   While Not rs.EOF
143:     ImageList1.ListImages.Add , rs("Caption"), LoadPicture(pICONES & rs("toolTip"))
144:     rs.MoveNext
145:   Wend
146:   rs.Close
147:   DB.Close

149:   StatusBar1.Panels(6).text = zUSER
150:   If Trim(PegPath("CITACAO", zUSER, "S")) = "S" Then
151:     frmDica.Show
152:   End If
153:   PrintPreview1.AuxiliaryButtonVisible = PrintPreview1.PrinterExists("Microsoft Print to PDF")
154:   PrintPreview1.AuxiliaryButtonToolTipText = "Salvar como PDF"
155:   Exit Sub
ErrorHandler:
  'Alert error$ & " - " & Str$(Err)
158:   Select Case Err
  Case 3024, 3044
160:     Alert "Verifiqie ini e banco de dados " & App.EXEName & ".MDB!"
  Case 53
162:     Alert "Verifique Arquivos mail.exe mail.ini"
163:     Resume Next
  Case Else
165:     Alert Err.Description
166:     Resume Next
167:   End Select
168:   Resume Next

End Sub
Public Sub PrintPreview1_AuxiliaryButtonClick(UpdateReport As Boolean)
172:   PrintPreview1.ShowSaveToFile "Microsoft Print to PDF", "*.pdf"
173:   UpdateReport = False  ' we don't need to update the report in the Print preview window after this action (the default value of UpdateReport parameter is True)
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)

178:   Unload Me
179:   End

End Sub

Private Sub mnuMENU_Click(Index As Integer)
184:   If Index = 7 Then Unload Me
End Sub

Private Sub mnuSUBMENU2_Click(Index As Integer)
188:   On Error Resume Next
189:   If Not AcessaMenu("mnuSUBMENU2", Index) Then Exit Sub
190:   zRPTARQ = PegPath("PATH", "RPT") ' agora no ini  + "RPT.MDB"
191:   cTIPORPT = "R"
192:   Sdb = ""
193:   zPFTIPO = ""
194:   ZENGTIP = ""
195:   iPPAP = 0
196:   Select Case Index
  Case 0
198:     zPFTIPO = "PFPRO"
199:     Sdb = PegPath("PATH", "PF")
200:     escpf.Show vbModal
  Case 1
202:     iPPAP = 1
203:     escPPAP.Show vbModal
  Case 2
205:     Sdb = PegPath("PATH", "PF")
206:     escfemea.Show vbModal
  Case 3
208:     Sdb = PegPath("PATH", "PF")
209:     escpc.Show vbModal
  Case 4
211:     Sdb = PegPath("PATH", "IE")
212:     escIE.Show vbModal
  Case 5
214:     iPPAP = 2
215:     escPPAP.Show vbModal
  Case 6
217:     escDes.Show vbModal
  Case 7
219:     ZENGTIP = "DES"
220:     ESCCLI.Show vbModal
  Case 8
222:     ESCpro.Show vbModal
  Case 9
224:     iPPAP = 3
225:     escPPAP.Show vbModal
  Case 10
227:     ZENGTIP = "PRO"
228:     ESCCLI.Show vbModal
  Case 11
230:     zPFTIPO = "POA"
231:     Sdb = PegPath("PATH", "POA")
232:     escPOA.Show vbModal
  Case 12
234:     FrmFemeaPad.Show vbModal
  Case 13
236:     iPPAP = 11
237:     escPPAP.Show vbModal
  Case 14
239:     zPFTIPO = "PFPRE"
240:     Sdb = PegPath("PATH", "PFP")
241:     escpf.Show vbModal
  Case 15
243:     zPFTIPO = "PFGP12"
244:     Sdb = PegPath("PATH", "PFG")
245:     escpf.Show vbModal
  Case 16
247:     zPFTIPO = "POKA"
248:     Sdb = PegPath("PATH", "POKA")
249:     escPOA.Show vbModal
  Case 17
251:     zPFTIPO = "DISPO"
252:     Sdb = PegPath("PATH", "DISPO")
253:     escPOA.Show vbModal
  Case 18
255:     ePASS01 = "LXINSP"
256:     Sdb = PegPath("PATH", "PF")
257:     escms01.Show vbModal, Me
258:   End Select

End Sub

Private Sub mnuSUBMENU3_Click(Index As Integer)

End Sub

Private Sub mnuSUBMENU4_Click(Index As Integer)
 ' Dim cLINHA As String

  ''todos usurios aceeos
  ''os que necessitam permissao case abaixo apos o check
271:   Select Case Index

  Case 0
   ' FrmPrintSetup.Show vbModal
275:     Exit Sub

  Case 1
278:     cARQRTF = PegPath("HELP", "HELPTXT")
279:     If FileConnExist(cARQRTF, True) Then
280:       PrintPreview1.ShowPreview
281:     End If
282:     Exit Sub
  Case 6
284:     zIDTEMP = zUSERID
285:     frmUSUSENHA.Show vbModal
286:     Exit Sub
  Case 7
288:     WritePrivateProfileString "CITACAO", zUSER, "S", App.Path + "\" & App.EXEName & ".INI"
289:     frmDica.Show
290:     Exit Sub
  Case 8
292:     ShellEx "calc"
293:     Exit Sub
  Case 9
295:     ShellEx "emailprg.exe", essSW_SHOWDEFAULT, "$" & UCase(zUSER), PegPath("PATH", "MAIL"), , CLng(Me.hwnd)
296:     Exit Sub
297:   End Select


300:   If Not AcessaMenu("mnuSUBMENU4", Index) Then
301:     Exit Sub
302:   End If
303:   Select Case Index
  Case 2
305:     escuser.Show vbModal
  Case 3
307:     escMENU.Show vbModal
  Case 4
309:     escBTN.Show vbModal
  Case 5
311:     configuraodbc
312:   End Select
End Sub

Private Sub mnuSUBMENU5_Click(Index As Integer)

317:   If Not AcessaMenu("mnuSUBMENU5", Index) Then Exit Sub
318:   Select Case Index
  Case 0
320:     ExpMS03.Show vbModal
  Case 1
322:     imgsin.Show vbModal
  Case 2
324:     ImportaWrptUser
  Case 3
326:     Alert ("Nao e mais possivel apos migracao P12")
    'Shell "improtheus", vbNormalFocus
328:   End Select


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

348:   cARQORI = PegPath("PATH", "WRPT")
349:   cARQDES = PegPath("PATH", "CONTROLE")
350:   cARQORI = GeracArq(cARQORI, , False)
351:   cARQDES = GeracArq(cARQDES, , True)
352:   Set oDBORI = New ADODB.Connection
353:   Set oDBDES = New ADODB.Connection
354:   Set oRSORI = New ADODB.Recordset
355:   Set oRSDES = New ADODB.Recordset

357:   oDBORI.ConnectionTimeout = 120
358:   oDBORI.Open cARQORI
359:   oDBDES.ConnectionTimeout = 120
360:   oDBDES.Open cARQDES
361:   oRSORI.Open "USUARIO", oDBORI, adOpenForwardOnly, adLockReadOnly
362:   While Not oRSORI.EOF
363:     cSQL = "SELECT * FROM USUARIO WHERE IDUSUARIO=" & FixInt(oRSORI("IDUSUARIO"))
364:     oRSDES.Open cSQL, oDBDES, adOpenDynamic, adLockOptimistic
365:     If oRSDES.EOF Then
366:       oRSDES.AddNew
367:       oRSDES("IDUSUARIO") = oRSORI("IDUSUARIO")
368:       oRSDES("SENHA") = oRSORI("SENHA")
369:     End If
    'For x = 0 To 10
    '   oRSDES(aCAM(x)) = oRSORI(aCAM(x))
    'Next x
    'nao trazer os campos pois pode ter trocado matricula no controle
    'e nao ter trocado no wrpt
    'bloqueando assim o acesso do usuario
376:     oRSDES.Update
377:     oRSDES.Close
378:     oRSORI.MoveNext
379:   Wend
380:   oRSORI.Close
381:   oDBDES.Close
382:   oDBORI.Close
End Sub

Private Sub mnuSUBMENU_Click(Index As Integer)
386:   zRPTARQ = PegPath("PATH", "RPT")
387:   cTIPORPT = "R"
388:   If UCase(Replace(mnuSUBMENU(Index).Caption, "&", "")) = "SAIR" Then
389:     Unload Me
390:   End If
391:   If Not AcessaMenu("mnuSUBMENU", Index) Then Exit Sub
392:   Select Case Index
  Case 0                                       'Clientes
394:     iImage = 1
395:     escIMG.Show vbModal
  Case 1                                       'Produtos
397:     iImage = 2
398:     escIMG.Show vbModal
  Case 2
400:     escFLX.Show vbModal, Me
  Case 3
402:     iMU01 = 999
403:     escIED.Show vbModal
  Case 4
405:     Zsetor = ""
406:     escPFD.Show vbModal, Me
  Case 5                                       'Ferramental
408:     escFER.Show vbModal
  Case 6                                       'Maquinas
410:     lRETU = False
411:     eRETU01 = ""
412:     eRETU02 = ""
413:     cARQESC = "ME01"
414:     escNUMNOM.Show vbModal, Me
    'iIMAGE = 4
    'escIMG.Show vbModal
  Case 7                                       'Equipamentos
418:     nPF = 0
419:     ePASS01 = ""
420:     escME04.Show vbModal
  Case 8                                       'Funcionarios
    'iImage = 6
    'escIMG.Show vbModal
    'ePASS01 = "LOGIX"
    'escMP04.Show vbModal, Me
426:     ePASS01 = ""
427:     escMP04.Show vbModal, Me
  Case 9                                       'Materia Prima
429:     iMU01 = 1
430:     escmu01.Show vbModal
    'iImage = 7
    'escIMG.Show vbModal
  Case 10                                      'componentes
434:     iMU01 = 2
435:     escmu01.Show vbModal

    'iImage = 8
    'escIMG.Show vbModal
  Case 11                                      'Consumiveis
440:     iMU01 = 3
441:     escmu01.Show vbModal

    'iImage = 9
    'escIMG.Show vbModal
  Case 12                                      ''ETI
446:     iMU01 = 4
447:     escmp.Show vbModal
  Case 13                                      ''AREAS
449:     escMP05.Show vbModal
  Case 14
451:     iImage = 11                              ''utilitarios
452:     escIMG.Show vbModal
  Case 15
454:     iMU01 = 4
455:     escmu01.Show vbModal
    'iImage = 10 ''Embalagens
    'escIMG.Show vbModal
  Case 16
459:     iMU01 = 6
460:     escmu01.Show vbModal
    'iImage = 12 ''Itens Manutencao
    'escIMG.Show vbModal
  Case 17
464:     iMU01 = 7
465:     escmu01.Show vbModal
    'iImage = 13 'sUB pRODUTOS
    'escIMG.Show vbModal
  Case 18
469:     iMU01 = 7
470:     escmp.Show vbModal, Me
471:   End Select

End Sub

Private Sub Relat_Click(Index As Integer)
476:   Select Case Index
  Case 0
478:     cTIPORPT = "R"
479:     cARQRTF = ""
480:     zRPTARQ = PegPath("PATH", "RPT")
481:     escRPTGRP.Show vbModal, Me
  Case 1
483:   cARQRTF = ""
484:     cTIPORPT = "D"
485:     zRPTARQ = PegPath("PATH", "DOC")
486:     escRPTGRP.Show vbModal, Me
  Case 2
488:     cTIPO = "R"
489:     escmanreg.Show vbModal, Me
  Case 3
491:     cTIPO = "P"
492:     escmanreg.Show vbModal, Me
  Case 4
494:     cARQRTF = ""
495:     cTIPO = "F"
496:     zRPTARQ = PegPath("PATH", "FOL")
497:     escRPTGRP.Show vbModal, Me
  Case 5
499:     cARQRTF = ""
500:     cTIPO = "I"
501:     zRPTARQ = PegPath("PATH", "INT")
502:     escRPTGRP.Show vbModal, Me
  Case 6
504:     FrmLstView.Show vbModal, Me
  Case 7
506:     ShellEx "WRPT", essSW_SHOWDEFAULT, "$" + Trim(zUSER) + "%#", PegPath("PATH", "WRPTX"), , CLng(Me.hwnd)
507:   End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
511:   Screen.MousePointer = vbDefault
End Sub

Private Sub configuraodbc()
  Dim lCONTINUA As Boolean
  Dim cCFG As String
  Dim aDAD As Variant
  Dim x As Integer
519:   lCONTINUA = True
520:   x = 1
521:   While lCONTINUA
522:     cCFG = FixStr(PegPath("ODBCCFG", "ODBCCFG" & StrZero(x, 3)))
523:     If InStr(cCFG, "|") = 0 Then
524:       lCONTINUA = False
525:     Else
526:       aDAD = Split(cCFG, "|")
527:       AddDSN aDAD(0), aDAD(1), aDAD(2), aDAD(3), True
528:     End If
529:     x = x + 1
530:   Wend
End Sub
Public Sub PrintPreview1_PrepareReport(Cancel As Boolean)
  Dim fileFile As Integer
  Dim strBuffer As String
535:   fileFile = FreeFile
536:   Open cARQRTF For Input As #fileFile
537:   Do While Not EOF(fileFile)
    'read line
539:     Input #fileFile, strBuffer
540:     Printer.Print strBuffer
541:   Loop
542:   Close fileFile
End Sub

