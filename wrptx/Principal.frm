VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{379157C5-E9BD-43F1-9F83-B037496BED42}#1.1#0"; "vbccr18.ocx"
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
   Begin VBCCR18.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      Top             =   7575
      Width           =   12180
      _ExtentX        =   21484
      _ExtentY        =   661
      InitPanels      =   "Principal.frx":058A
   End
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
 ' Dim nPOS3 As Integer
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
  
  'cmdline = Trim(Command())
  
  
  
  'comand line properties project
  
  ' $admin%CU#VIA_fol0001.rpt
  ' $user  %grupo #subgrupo _relatorio
  
  ' exemplo charmar rpt
  '  cTIPO = "R" o tipo e pego abaixo conforme o exe wrptx wrptf wrpti wrptd
  '  zgrp = "PF"
  '  ZGRPSUB = "IE"
  '  eLOCALIZA = "ITA00148"
  '  escRPT.Show vbModal, Me
  
    
  'C:\TEMP\TESTE.TXT
  ' abre as extensoes com o preview
  '.txt .rtf .jpg .zpl
  
  ' teste convert pdf
 ' ePASS01 = "C:\TEMP\TESTE.TXT"
 ' formConvertToPDF.Show vbModal, Me
 ' End
   
  ' teste frmrtf
  ' aDIREITOS = Array(True, True, True, True, True, True, True, True)
  ' FrmRTf.Show vbModal
  ' end
  
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


  zgrp = ""
  ZGRPSUB = ""
  zUSER = ""
  eLOCALIZA = ""
  cARQRTF = ""
  
  'aRETU = pegue2delimitado("where mes=MONTH(%1%) AND ano=YEAR(%1%)", "YEAR(", ")")

  'aRETU = pegue2delimitado("where mes=MONTH(%1%) AND ano=YEAR(%1%)", "MONTH(", ")")

'  ctmpline = SQLDIAPAR("where mes=MONTH(%1%) AND ano=YEAR(%1%)", "YEAR(", ")", "TOTEXT(", ",'YYYY')")
 ' ctmpline = SQLDIAPAR(ctmpline, "MONTH(", ")", "TOTEXT(", ",'MM')")
  'End
 
  'pegando o usuario
  ctmpline = cmdline
  aRETU = pegue2delimitado(Command(), "$", "%")
  zUSER = aRETU(0)
  ctmpline = aRETU(1)

  'pegando o grupo
  aRETU = pegue2delimitado(Command(), "%", "#")
  zgrp = aRETU(0)
  ctmpline = aRETU(1)
  
   'pegando o subgrupo
  aRETU = pegue2delimitado(Command(), "#", "_")
  ZGRPSUB = aRETU(0)
  ctmpline = aRETU(1)
   
   
  'pegando o relatorio
  aRETU = pegue2delimitado(Command(), "_", "")
  eLOCALIZA = aRETU(0)
  ctmpline = aRETU(1)
  

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

  If Not FileConnExist(Dbname, True) Then
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


  StatusBar1.Panels(6).Text = zUSER

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
