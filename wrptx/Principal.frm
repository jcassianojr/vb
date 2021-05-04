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
            TextSave        =   "11:14"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            Object.Width           =   2302
            MinWidth        =   2293
            Picture         =   "Principal.frx":0B24
            TextSave        =   "04/05/2021"
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
   
    CenterFormToScreen Me
    cmdline = Trim(Command())
    
'    aDIREITOS = Array(True, True, True, True, True, True, True, True)
'    FrmRTf.Show
   
    zgrp = ""
    ZGRPSUB = ""
    zUSER = ""
      
    nPOS = InStr(cmdline, "%")
    zUSER = Mid(cmdline, 2, nPOS - 2)
   
    nPOS2 = InStr(cmdline, "#")
    If nPOS2 <> Len(cmdline) Then
        nPOS3 = nPOS2 - nPOS - 1
        zgrp = Mid(cmdline, nPOS + 1, nPOS3)
        ZGRPSUB = Mid(cmdline, nPOS2 + 1)
    Else
        zgrp = Mid(cmdline, nPOS + 1, Len(cmdline) - nPOS - 1)
    End If
   
    zUSER = UCase(zUSER)
    zgrp = UCase(zgrp)
    ZGRPSUB = UCase(ZGRPSUB)
   
    If zUSER = "SUPERVISOR" Then
        zUSER = "ADMIN"
    End If
   
   
    'Caminhos
    pICONES = PegPath("PATH", "ICONES")
    zAppPath = PegPath("PATH", "APP")
    zMANA5EMP = PegPath("PATH", "MANA5EMP")
    ZMANA5CON = PegPath("PATH", "MANA5CON")
    ZMANA5IMP = PegPath("PATH", "IMPREL")
    ZRELAT = PegPath("PATH", "RELAT")
    DBWRPT = PegPath("PATH", "WRPT")
    Dbname = PegPath("PATH", "MDB") + "WRPT.MDB"
    zusalx = PegPath("PATH", "USALX")
    zusamc = PegPath("PATH", "USAMC")
    
    cTIPORPT = "R"
    zRPTARQ = PegPath("PATH", "RPT") + "RPT.MDB"
    

    If Trim(UCase(App.EXEName)) = "WRPTF" Then
       cTIPORPT = "F"
       zRPTARQ = PegPath("PATH", "FOL") + "RPTFOL.MDB"
    End If
    If Trim(UCase(App.EXEName)) = "WRPTD" Then
       cTIPORPT = "D"
       zRPTARQ = PegPath("PATH", "RTF") + "RTF.MDB"
    End If
    If Trim(UCase(App.EXEName)) = "WRPTI" Then
       cTIPORPT = "I"
       zRPTARQ = PegPath("PATH", "INT") + "RPTINT.MDB"
    End If
    
    
    
    
    If Not FileExist(Dbname, True) Then
        Alert ("Erro Ini [PATH] MDB=")
        End
    End If
    
   
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
   
   
   
   

   
   
    ' Carrega imagens para o ImageList
    Set DAODB = New ADODB.Connection
    DAODB.CursorLocation = adUseClient
    DAODB.ConnectionTimeout = 120
    DAODB.Open GeracArq(Dbname, , False)

    
    Set DAORS = New ADODB.Recordset
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

    zEMPRESA = 1
    zMES = Month(Date)
    zANO = Year(Date)

    zWRPTID = 0
    zUSERID = 0
    zIDFOLHA = 0
    zNOMEFOLHA = ""
    
    
    
    
    
    sSQL = "SELECT IDUSUARIO,USUARIO,IDFOLHA,NOMEFOLHA FROM USUARIO WHERE USUARIO='" & zUSER & "'"
    
   aRETU = PegSQL(DBWRPT, sSQL, 3, Array("IDUSUARIO", "IDFOLHA", "NOMEFOLHA"), _
                              Array("NI", "NI", "C"), _
                              Array(0, 0, ""))
    If lRETU Then
        zWRPTID = aRETU(0)
        zUSERID = zWRPTID
        zIDFOLHA = aRETU(1)
        zNOMEFOLHA = aRETU(2)
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
        MsgBox "Usuário Não Cadastrado", vbOKOnly, "Bloqueio de Acesso"
        Unload Me
    Else
        If zgrp = "" Then
            escRPTGRP.Show vbModal
            End
        Else
            escRPT.Show vbModal
            End
        End If
    End If



   cARQ = PegPath("PATH", "SYSUSER")
   sSQL = "select * from USUARIO WHERE USUARIO='" & zUSER & "'"
   GrvSQL cARQ, sSQL, 1, Array("DATAULT"), Array(Today()), Array("D")


   

End Sub
