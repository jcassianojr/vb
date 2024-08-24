VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmPRINCIPAL 
   BackColor       =   &H8000000C&
   Caption         =   "Menu Inicial"
   ClientHeight    =   7950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12180
   Icon            =   "Principal.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   840
      Top             =   120
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
      Top             =   7635
      Width           =   12180
      _ExtentX        =   21484
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            AutoSize        =   2
            Enabled         =   0   'False
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
            Picture         =   "Principal.frx":058A
            TextSave        =   "12:54"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            Object.Width           =   2302
            MinWidth        =   2293
            Picture         =   "Principal.frx":0B24
            TextSave        =   "24/08/2024"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Picture         =   "Principal.frx":10BE
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11192
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmPRINCIPAL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub MDIForm_Load()
  Dim nPOS As Integer
  Dim nPOS2 As Integer
  Dim nPOS3 As Integer
  Dim cmdline As String
  Dim pICONES As String
  Dim nRETU
  Dim DAODB As ADODB.Connection
  Dim DAORS As ADODB.Recordset
  Dim sSQL As String
  Dim aRETU As Variant
  Dim cARQ As String
  Dim cARQICO As String
  Dim ctmpline As String

  CenterFormToScreen Me
  cmdline = Trim(Command())
  
  
  'comand line properties project
  '$admin%CU#VIA
  ' user  grupo sub grupo
  'lstview
  '.txt .rtf jpg preview
  
  ' exemplo charmar rpt
  ' cTIPO = "R"
  '  zgrp = "PF"
  '  ZGRPSUB = "IE"
  '  eLOCALIZA = "ITA00148"
  '  escRPT.Show vbModal, Me
  
 ' ePASS01 = "C:\TEMP\TESTE.TXT"
 ' formConvertToPDF.Show vbModal, Me
 ' End
   
  
  If InStr(UCase(cmdline), "LSTVIEW") Then
     FrmLstView.Show vbModal, Me
     End
  End If
  
   If InStr(UCase(cmdline), ".TXT") _
     Or InStr(UCase(cmdline), ".RTF") _
     Or InStr(UCase(cmdline), ".ZPL") _
     Or InStr(UCase(cmdline), ".JPG") Then
     cARQRTF = cmdline
     escRPT.Show vbModal, Me
     End
  End If
  
  ' aDIREITOS = Array(True, True, True, True, True, True, True, True)
  ' FrmRTf.Show vbModal

  zgrp = ""
  ZGRPSUB = ""
  zUSER = ""
  eLOCALIZA = ""
  cARQRTF = ""

  '$admin%CU#VIA_RPT0001
  
  'pegando o usuario
  ctmpline = cmdline
  nPOS = InStr(ctmpline, "$")
  nPOS2 = InStr(ctmpline, "%")
  If nPOS > 0 And nPOS2 = 0 Then
    zUSER = Mid(ctmpline, 2)
    ctmpline = ""
  End If
  If nPOS > 0 And nPOS2 > 0 Then
    zUSER = Mid(ctmpline, 2, nPOS2 - 2)
    ctmpline = Mid(ctmpline, nPOS2)
  End If
  'pegando o grupo
  nPOS = InStr(ctmpline, "%")
  nPOS2 = InStr(ctmpline, "#")
  If nPOS > 0 And nPOS2 = 0 Then
    zgrp = Mid(ctmpline, 2)
    ctmpline = ""
  End If
  If nPOS > 0 And nPOS2 > 0 Then
    zgrp = Mid(ctmpline, 2, nPOS2 - 2)
    ctmpline = Mid(ctmpline, nPOS2)
  End If
   'pegando o subgrupo
  nPOS = InStr(ctmpline, "#")
  nPOS2 = InStr(ctmpline, "_")
  If nPOS > 0 And nPOS2 = 0 Then
    ZGRPSUB = Mid(ctmpline, 2)
    ctmpline = ""
  End If
  If nPOS > 0 And nPOS2 > 0 Then
    ZGRPSUB = Mid(ctmpline, 2, nPOS2 - 2)
    ctmpline = Mid(ctmpline, nPOS2)
  End If
  nPOS2 = InStr(ctmpline, "_")
  If nPOS2 > 0 Then
     eLOCALIZA = Mid(ctmpline, 2)
  End If
  

  zUSER = UCase(zUSER)
  zgrp = UCase(zgrp)
  ZGRPSUB = UCase(ZGRPSUB)

  If zUSER = "SUPERVISOR" Then
    zUSER = "ADMIN"
  End If

  cTIPORPT = "R"
  zRPTARQ = PegPath("PATH", "RPT")
  If Trim(UCase(App.EXEName)) = "WRPTF" Then
    cTIPORPT = "F"
    zRPTARQ = PegPath("PATH", "FOL")
  End If
  If Trim(UCase(App.EXEName)) = "WRPTD" Then
    cTIPORPT = "D"
    zRPTARQ = PegPath("PATH", "DOC")
  End If
  If Trim(UCase(App.EXEName)) = "WRPTI" Then
    cTIPORPT = "I"
    zRPTARQ = PegPath("PATH", "INT")
  End If
  cTIPO = cTIPORPT

  'Caminhos
  pICONES = PegPath("PATH", "ICONES")
  zAppPath = PegPath("PATH", "APP")
  zMANA5EMP = PegPath("PATH", "MANA5EMP")
  ZMANA5CON = PegPath("PATH", "MANA5CON")
  ZMANA5IMP = PegPath("PATH", "IMPREL")
  ZRELAT = PegPath("PATH", "RELAT")
  DBWRPT = PegPath("PATH", "WRPT")
  Dbname = PegPath("PATH", "MDB") ' agora no ini + "WRPT.MDB"
  zusalx = PegPath("PATH", "USALX")
  zusamc = PegPath("PATH", "USAMC")

  If Not FileExist(Dbname, True) Then
    Alert ("Erro Ini [PATH] MDB=")
    End
  End If

  'Cria o Arquivo log do mes
  cARQERRO = PegPath("PATH", "LOGCAM") & "LOG" & Format(Now, "MMYYYY") & ".sqlite"
  nRETU = CopyFileWindowsWay(PegPath("PATH", "LOG"), cARQERRO, False)
  If nRETU <> 0 Then
      Alert "Erro Criando Arquivo Log do Mes"
      End
  End If


  'competencia empresa
  zEMPRESA = 1
  zMES = Month(Date)
  zANO = Year(Date)

  'dados do usuario
  zWRPTID = 0
  zUSERID = 0
  zIDFOLHA = 0
  zIDUNI = 0
  zNOMEFOLHA = ""
  sSQL = "SELECT IDUSUARIO,IDFOLHA,NOMEFOLHA,ID FROM USUARIO WHERE USUARIO='" & zUSER & "'"
  aRETU = PegSQL(DBWRPT, sSQL, 4, Array("IDUSUARIO", "IDFOLHA", "NOMEFOLHA", "ID"), _
                 Array("NI", "NI", "C", "NI"), _
                 Array(0, 0, "", 0))
  If lRETU Then
    zWRPTID = aRETU(0)
    zUSERID = zWRPTID
    zIDFOLHA = aRETU(1)
    zNOMEFOLHA = aRETU(2)
    zIDUNI = aRETU(3)
  End If


  If zIDFOLHA = 0 Then
    Alert ("No Folha de Pagamento WRPT Nao Preenchido " & zUSER)
    End
  End If

  If Len(zNOMEFOLHA) = 0 Then
    Alert ("Nome Folha de Pagamento WRPT Nao Preenchido" & zUSER)
    End
  End If
  If zIDFOLHA > 0 Then
    If demitido(zIDFOLHA) Then
      End
    End If
  End If


  StatusBar1.Panels(6).tEXT = zUSER

  If zWRPTID = 0 Then
    MsgBox "Usuario Nao Cadastrado", vbOKOnly, "Bloqueio de Acesso"
    End 'unload me
  End If
  
  'grava ultima data de acesso
 'cARQ = PegPath("PATH", "SYSUSER")
 ' sSQL = "select DATAULT from USUARIO WHERE USUARIO='" & zUSER & "'"
 'sSQL = "select DATAULT from USUARIO WHERE ID=" & zIDUNI
'
'checar gravacao sqlite tentar por select id ou current date
'GrvSQL cARQ, sSQL, 1, Array("DATAULT"), Array(Today()), Array("DH")
 '
 '
 
 cARQ = PegPath("PATH", "SYSUSER")
 sSQL = "UPDATE USUARIO  SET DATAULT = CURRENTDATETIME   WHERE  USUARIO = '" & zUSER & "'"
 ADOComando cARQ, sSQL
  
  ' Carrega imagens para o ImageList
  Set DAODB = New ADODB.Connection
  DAODB.CursorLocation = adUseClient
  DAODB.ConnectionTimeout = 120
  DAODB.Open GeracArq(Dbname, , False)
  Set DAORS = New ADODB.Recordset
  DAORS.Open "select * from controle where form='frmMDIPRINCIPAL'", DAODB, adOpenForwardOnly, adLockReadOnly
  With DAORS
    Do While Not DAORS.EOF
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
  

End Sub
