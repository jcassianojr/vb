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
6:   gbSubClassMenu = True

8:   If gbSubClassMenu Then SubClassMenuXP

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


26:   On Error GoTo ErrorHandler

28:   CenterFormToScreen Me
  
 
  
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

51:   If App.PrevInstance Then

53:     MsgBox App.EXEName & " já esta rodando!", 48
54:     End

56:   End If

  '* Pegando configuracoes do arq .INI
59:   strCodeKey = "SoFtEc1SiStEmAs24055-2265"
60:   SysCodeKey = "SoFtEc1SiStEmAs24055-2265"

  'Caminhos
63:   pICONES = PegPath("PATH", "ICONES")
64:   zAppPath = PegPath("PATH", "APP")
65:   zMANA5EMP = PegPath("PATH", "MANA5EMP")
66:   ZMANA5CON = PegPath("PATH", "MANA5CON")
67:   ZMANA5IMP = PegPath("PATH", "IMPREL")
68:   ZRELAT = PegPath("PATH", "RELAT")
69:   zusalx = PegPath("PATH", "USALX")
70:   zusamc = PegPath("PATH", "USAMC")
71:   dbuser = PegPath("PATH", "USER")
72:   DBWRPT = PegPath("PATH", "WRPT")
73:   Dbname = PegPath("PATH", "MDB")

75:   If Not FileConnExist(Dbname, True) Then
76:     Alert ("Erro Ini [PATH] MDB=")
77:     End
78:   End If
79:   If Not FileConnExist(dbuser, True) Then
80:     Alert ("Erro Ini [PATH] USER=")
81:     End
82:   End If
83:   If Not FileConnExist(DBWRPT, True) Then
84:     Alert ("Erro Ini [PATH] WRPT=")
85:     End
86:   End If

  'se sql lite verifica se odbc esta instalado
 If InStr(LCase(dbuser), ".sqlite") > 0 Or InStr(LCase(dbuser), ".sqlite3") > 0 _
                 Or InStr(LCase(dbuser), ".fossil") > 0 Or InStr(LCase(dbuser), ".db3") > 0 _
                 Or (InStr(LCase(dbuser), ".db") > 0) Then
92:      If IsDriverInstalled("SQLite3 ODBC Driver") Then
       ' MsgBox "O driver SQLite ODBC está pronto para uso!", vbInformation
94:      Else
95:         MsgBox "Driver SQLite ODBC não encontrado. Por favor, instale o driver para continuar.", vbCritical
96:         End
97:      End If
98:  End If

100:   carqhelp = PegPath("PATH", "HELP")
101:   If FileConnExist(carqhelp) Then
102:     App.HelpFile = carqhelp
103:   End If

105:   Set DAODB = New ADODB.Connection
106:   Set DAORS = New ADODB.Recordset

    'Cria o Arquivo log do mes
109:   cARQERRO = PegPath("PATH", "LOGCAM") & "LOG" & Format(Now, "MMYYYY") & ".sqlite"
110:   If Not FileExists(cARQERRO) Then
111:      CopyFileWindowsWay PegPath("PATH", "LOG"), cARQERRO, True
112:   End If
113:   If Not FileExists(cARQERRO) Then
114:       Alert "Erro Criando Arquivo Log do Mes"
115:       End
116:   End If

118:   zEMPRESA = 1
119:   zMES = Month(Date)
120:   zANO = Year(Date)
121:   ZGRPSUB = ""
  
123:   zgrp = ""
124:   ZGRPSUB = ""
125:   zUSER = ""
126:   eLOCALIZA = ""
127:   cARQRTF = ""
  
  
  
  
  
   'pegando o usuario
134:    cmdline = Command()
135:    cmdline = Replace(cmdline, "__", "")
   
137:   If InStr(UCase(cmdline), "LSTVIEW") Then
138:      FrmLstView.Show vbModal, Me
139:      End
140:   End If
  
   If InStr(UCase(cmdline), ".ZPL") _
     Or InStr(UCase(cmdline), ".JPG") Then
144:      cARQRTF = cmdline
145:      escRPT.Show vbModal, Me
146:      End
147:   End If
  
  If InStr(UCase(cmdline), ".TXT") Or InStr(UCase(cmdline), ".MAN") _
     Or InStr(UCase(cmdline), ".RTF") Then
151:      cARQRTF = cmdline
152:      escRPT.Show vbModal, Me
     'criar opcao de preview direto
     'FrmRtfView.Show vbModal, Me
155:      End
156:   End If
   
   
159:    cTIPORPT = ""
160:    zRPTARQ = ""
161:    If Mid(cmdline, 1, 1) = "X" Or Mid(cmdline, 1, 1) = "R" Then
162:       cTIPORPT = "R"
163:       cmdline = Mid(cmdline, 2)
164:       zRPTARQ = PegPath("PATH", "RPT")
165:    End If
166:    If Mid(cmdline, 1, 1) = "D" Then
167:       cTIPORPT = "D"
168:       cmdline = Mid(cmdline, 2)
169:       zRPTARQ = PegPath("PATH", "DOC")
170:    End If
171:    If Mid(cmdline, 1, 1) = "I" Then
172:       cTIPORPT = "I"
173:       cmdline = Mid(cmdline, 2)
174:       zRPTARQ = PegPath("PATH", "INT")
175:    End If
176:    If Mid(cmdline, 1, 1) = "F" Then
177:       cTIPORPT = "F"
178:       cmdline = Mid(cmdline, 2)
179:       zRPTARQ = PegPath("PATH", "FOL")
180:    End If
   ''especifico ou padrao manrel mana5
182:    If Mid(cmdline, 1, 1) = "E" Then
183:       cTIPORPT = "E"
184:    End If
185:    If Mid(cmdline, 1, 1) = "P" Then
186:       cTIPORPT = "P"
187:    End If
   
   
190:   aRETU = pegue2delimitado(cmdline, "$", "%")
191:   zUSER = aRETU(0)
192:   ctmpline = aRETU(1)

  'pegando o grupo
195:   aRETU = pegue2delimitado(cmdline, "%", "#")
196:   zgrp = aRETU(0)
197:   ctmpline = aRETU(1)
  
   'pegando o subgrupo
200:   aRETU = pegue2delimitado(cmdline, "#", "_")
201:   ZGRPSUB = aRETU(0)
202:   ctmpline = aRETU(1)
   
   
  'pegando o relatorio
206:   aRETU = pegue2delimitado(cmdline, "_", "")
207:   eLOCALIZA = aRETU(0)
208:   ctmpline = aRETU(1)
  

211:   zUSER = UCase(zUSER)
212:   zgrp = UCase(zgrp)
213:   ZGRPSUB = UCase(ZGRPSUB)

215:   If zUSER = "SUPERVISOR" Then
216:     zUSER = "ADMIN"
217:   End If
  
  'aRETU = pegue2delimitado(Trim(Command()), "__$", "%__")
  
  
222:   bACESSO = False
  'dados do usuario
224:   zWRPTID = 0
225:   zUSERID = 0
226:   zIDFOLHA = 0
227:   zIDUNI = 0
228:   zNOMEFOLHA = ""
229:   If Len(zUSER) > 0 Then
230:     sSQL = "SELECT IDUSUARIO,USUARIO,IDFOLHA,NOMEFOLHA FROM USUARIO WHERE USUARIO='" & zUSER & "'"
    aRETU = PegSQL(dbuser, sSQL, 3, Array("IDUSUARIO", "IDFOLHA", "NOMEFOLHA"), _
                   Array("NI", "NI", "C"), _
                   Array(0, 0, ""))
234:     If lRETU Then
235:       zUSERID = aRETU(0)
236:       zWRPTID = zUSERID
237:       zIDFOLHA = aRETU(1)
238:       zNOMEFOLHA = aRETU(2)
239:       bACESSO = True
240:     End If

242:   End If

244:   If Not bACESSO Then
245:     frmSENHA.Show vbModal, Me
246:   End If

248:   If zIDFOLHA > 0 Then
249:     If demitido(zIDFOLHA) Then
250:       End
251:     End If
252:   End If

254:   Me.Caption = App.FileDescription & " Versao:" & App.Major & "." & App.Minor & "Usuário: " & zUSER


257:   If Trim(PegPath("EMAILINT", zUSER, "S")) = "S" Then
258:      ShellEx "MAIL", essSW_SHOWDEFAULT, "$" & UCase(zUSER), PegPath("PATH", "MAIL"), , Me.hwnd
259:   End If


  '* Carrega imagens para o ImageList

264:   DAODB.CursorLocation = adUseClient
265:   DAODB.ConnectionTimeout = 120
266:   DAODB.Open GeracArq(Dbname, , False)



270:   DAORS.Open "select * from controle where form='frmMDIPRINCIPAL'", DAODB, adOpenForwardOnly, adLockReadOnly

272:   With DAORS
273:     Do While Not DAORS.EOF  '!controle = "ImageList1"
274:       cARQICO = pICONES & DAORS.Fields("TOOLTIP")
275:       If FileConnExist(cARQICO, True) Then
276:         ImageList1.ListImages.Add , DAORS.Fields("Caption"), LoadPicture(cARQICO)
277:       Else
278:         End
279:       End If
280:       .MoveNext
281:     Loop
282:   End With
283:   DAORS.Close
284:   DAODB.Close
285:   Set DAODB = Nothing
286:   Set DAORS = Nothing



290:   StatusBar1.Panels(6).tEXT = zUSER
291:   If Trim(PegPath("CITACAO", zUSER, "S")) = "S" Then
292:     frmDica.Show
293:   End If
  
  
  'abre caso o tipo nao esteja em branco
297:   If cTIPORPT <> "" Then
298:      If cTIPORPT = "E" Or cTIPORPT = "P" Then
       'especificos ou padrao manrel mana5 imprel via shell
300:        If cTIPORPT = "E" Then
301:           mnuSUBMENU1_Click (0)
302:        End If
303:         If cTIPORPT = "P" Then
304:           mnuSUBMENU1_Click (1)
305:        End If
306:      Else
307:         If eLOCALIZA = "" Then
308:            If zgrp = "" Then
309:               cARQRTF = ""
310:               eLOCALIZA = ""
311:               escRPTGRP.Show vbModal
312:               End
313:            Else
314:               cARQRTF = ""
315:               eLOCALIZA = ""
316:               escRPT.Show vbModal
317:               End
318:            End If
319:         Else
320:             escRPT.Show vbModal
321:             End
322:         End If
323:       End If
324:     End If
325:   Exit Sub

ErrorHandler:


  '* MsgBox Error$ & " - " & str$(Err), vbExclamation, "Erro!"

332:   Select Case Err

  Case 3024, 3044

    'set CommonDialog properties before showing


339:     MsgBox "Verifiqie ini e banco de dados " & App.EXEName & ".MDB!", vbExclamation, "Erro!"

  Case 53


344:     MsgBox "Verifique Arquivos mail.exe mail.ini", vbExclamation, "Erro!"
345:     Resume Next



  Case Else

351:     Resume Next

353:   End Select

End Sub
Private Sub MDIForm_Unload(Cancel As Integer)


359:   If gbSubClassMenu Then
    '/ prevent error if the menu is not subclassed
361:     On Error Resume Next
    '/ release object
363:     Call objMenuEx.Uninstall(Me.hwnd, ImgMenu, MenuEvents)
364:     Set MenuEvents = Nothing
365:     Set objMenuEx = Nothing
366:   End If


369:   Unload Me
370:   End
End Sub
Private Sub Menu2_Click()
373:   cARQRTF = ""
374:   cTIPORPT = "R"
375:   zRPTARQ = PegPath("PATH", "RPT") ' agora no ini + "RPT.MDB"
376:   escRPTGRP.Show vbModal, Me
End Sub

Private Sub Menu3_Click()
380:   cARQRTF = ""
381:   cTIPORPT = "D"
382:   zRPTARQ = PegPath("PATH", "DOC") ' agora no ini + "RTF.MDB"
383:   escRPTGRP.Show vbModal, Me
End Sub

Private Sub Menu4_Click()
387:   cARQRTF = ""
388:   cTIPORPT = "F"
389:   zRPTARQ = PegPath("PATH", "FOL")
390:   escRPTGRP.Show vbModal, Me
End Sub

Private Sub Menu5_Click()
394:   cARQRTF = ""
395:   cTIPORPT = "I"
396:   zRPTARQ = PegPath("PATH", "INT")
397:   escRPTGRP.Show vbModal, Me
End Sub

Private Sub Menu6_Click()
401:   FrmLstView.Show
End Sub

Private Sub Menu8_Click()
405:   Unload Me
End Sub

Private Sub mnuSUBMENU1_Click(Index As Integer)
409:   Select Case Index
  Case 0
411:     cTIPO = "R"  'especifico
412:     escmanreg.Show vbModal, Me
  Case 1
414:     cTIPO = "P"  'padrao
415:     escmanreg.Show vbModal, Me
416:   End Select
End Sub

Private Sub mnuSUBMENU4_Click(Index As Integer)

421:   If Not AcessaMenu("mnuSUBMENU4", Index) Then Exit Sub
  ''alert CStr(Index)
423:   Select Case Index

  Case 0

   ' FrmPrintSetup.Show vbModal, Me

  Case 1
430:     cARQRTF = PegPath("HELP", "CAMINHO") + "WRPT.txt"
    ' 0-Novo,1-Abrir,2-Salvar,3-Salvacomo
    ' 4-Editar,5-Visualizar,6-Imprimir,7-Exportar
433:     aDIREITOS = Array(False, False, False, False, True, True, True, False)
434:     FrmRtfView.Show vbModal, Me

  Case 2
437:     escuser.Show vbModal, Me


  Case 3
441:     escMENU.Show vbModal, Me

  Case 4
444:     escBTN.Show vbModal

  Case 5
447:     configuraodbc
  Case 6
449:     zIDTEMP = zUSERID
450:     frmUSUSENHA.Show vbModal

  Case 7
453:     frmDica.Show

  Case 8
    'Calculadora.Show
457:     ShellEx "calc"

459:   End Select




End Sub

Private Sub configuraodbc()
  Dim lCONTINUA As Boolean
  Dim cCFG As String
  Dim aDAD As Variant
  Dim x As Integer
471:   lCONTINUA = True
472:   x = 1
473:   While lCONTINUA
474:     cCFG = FixStr(PegPath("ODBCCFG", "ODBCCFG" & StrZero(x, 3)))
475:     If InStr(cCFG, "|") = 0 Then
476:       lCONTINUA = False
477:     Else
478:       aDAD = Split(cCFG, "|")
479:       AddDSN aDAD(0), aDAD(1), aDAD(2), aDAD(3), True
480:     End If
481:     x = x + 1
482:   Wend

End Sub


Private Sub MenuEvents_MenuHelp(ByVal MenuText As String, ByVal MenuHelp As String, ByVal Enabled As Boolean)
488:   If gbSubClassMenu Then
    '/ this event show Decription menu-item to a StatusBar control
    '/ customize for your project, if you need this.
    '/ Note: MenuText contain the menu Caption.
    '/       MenuHelp contain the Description string.
    '/ example: sbStatusBar.Panels(1).text = MenuHelp
    '/          lblMenuHelp.Caption = MenuHelp
495:   End If
End Sub
Public Sub SubClassMenuXP()

'/ this code is made by MenuCreator add-in

'/ prepare the caption for subclassing. Warning! Don't remove this comment!!!
502:   Menu1.Caption = "&Gerador"
503:   mnuSUBMENU1(0).Caption = "Especificos"
504:   mnuSUBMENU1(1).Caption = "Padrao"
505:   Menu2.Caption = "&Relatórios"
506:   Menu3.Caption = "&Documentos/Manuais"
507:   Menu4.Caption = "&Folha/Pto/RH"
508:   Menu5.Caption = "&Integrados"
509:   Menu6.Caption = "&Lst(Logix)"
510:   Menu7.Caption = "&Configurações"
511:   mnuSUBMENU4(0).Caption = "Impressoras|#printcfg"
512:   mnuSUBMENU4(1).Caption = "Alteracoes|#info"
513:   mnuSUBMENU4(2).Caption = "Usuarios|#user"
514:   mnuSUBMENU4(3).Caption = "Menus|#menusa"
515:   mnuSUBMENU4(4).Caption = "Botoes|#todo"
516:   mnuSUBMENU4(5).Caption = "Configurar ODBC|#odbcinfo"
517:   mnuSUBMENU4(6).Caption = "Troca Senha|#usuerchave"
518:   mnuSUBMENU4(7).Caption = "Frase do Dia|#lampada"
519:   mnuSUBMENU4(8).Caption = "Calculadora|#calc"
520:   Menu8.Caption = "&Sair"
  
  
  'imgmenu.index
'    AplicarIconeMenu hMenuConfiguracoes, 0, 10 'mnuSUBMENU4(0).Caption = "Impressoras|#printcfg" 27
 '   AplicarIconeMenu hMenuConfiguracoes, 1, 0 'mnuSUBMENU4(1).Caption = "Alteracoes|#info"       41
 '   AplicarIconeMenu hMenuConfiguracoes, 2, 0 'mnuSUBMENU4(2).Caption = "Usuarios|#user"         49
 '   AplicarIconeMenu hMenuConfiguracoes, 3, 0 'mnuSUBMENU4(3).Caption = "Menus|#menusa"          44
 '   AplicarIconeMenu hMenuConfiguracoes, 4, 0 'mnuSUBMENU4(4)"Botoes|#todo"                      47
 '   AplicarIconeMenu hMenuConfiguracoes, 5, 0 'mnuSUBMENU4(5)"Configurar ODBC|#odbcinfo"         46
 '   AplicarIconeMenu hMenuConfiguracoes, 6, 0 'mnuSUBMENU4(6)"Troca Senha|#usuerchave"           45
 '   AplicarIconeMenu hMenuConfiguracoes, 7, 0 'mnuSUBMENU4(7)"Frase do Dia|#lampada"             48

  '/ Subclassing menu. Warning! Don't remove this comment!!!

535:   Set MenuEvents = New CEvents
536:   Set objMenuEx = New cMenuEx
537:   Call objMenuEx.Install(Me.hwnd, App.Path & "\" & Me.Name, ImgMenu, 2, MenuEvents)

End Sub

Public Sub MenuDesigner()
'/ Open Menu Designer tool
543:   objMenuEx.MenuDesigner Me.hwnd
End Sub
