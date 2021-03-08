VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form escUser 
   Caption         =   "Escolhendo Usuarios"
   ClientHeight    =   7605
   ClientLeft      =   300
   ClientTop       =   210
   ClientWidth     =   11715
   HelpContextID   =   11
   Icon            =   "escUSER.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   11715
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   7605
      Left            =   10215
      TabIndex        =   1
      ToolTipText     =   "Botoes de Açoes"
      Top             =   0
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   13414
      ButtonWidth     =   1138
      ButtonHeight    =   582
      Appearance      =   1
      TextAlignment   =   1
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   7455
      Left            =   90
      TabIndex        =   0
      ToolTipText     =   "Lista de Registros"
      Top             =   90
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   13150
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "escUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const nFORMID = 1011
Const cFORMID = "Usuarios do Sistema"
Dim cARQ As String
Dim cSQL As String
Private Sub Apagar()
Dim nUSER As Long
    Grid.Col = 1
    nUSER = FixNum(Grid)
    cSQL = "select * from USUARIO WHERE IDUSUARIO=" & nUSER
    lRETU = ApagaSQLP(cARQ, cSQL, "Apagar")
    If lRETU Then
        cARQ = PegPath("PATH", "LOGCAM") & "LOG" & Format(Now, "MMYYYY") & ".MDB"
        cSQL = "select * from LOG WHERE USER=" & nUSER
        lRETU = ApagaSQLP(cARQ, cSQL, "Apagar Log do Sistema")
        Filgrid
    End If
End Sub

Private Sub Editar()
Dim nUSER As Long
    Grid.Col = 1
    nUSER = FixNum(Grid)
    ePASS01 = "select * from USUARIO WHERE IDUSUARIO=" & nUSER
    ePASS02 = nUSER
    frmUSER.Show vbModal
    Filgrid
  '  localizagrid Grid, nUSER, 2, False
End Sub
Private Sub senha()
    Grid.Col = 1
    zIDTEMP = FixNum(Grid)
    frmUSUSENHA.Show vbModal
End Sub


Private Sub Escolher()
    lRETU = True
    Grid.Col = 1   ' IDUSUARIO
    eRETU01 = Grid
    Grid.Col = 2   ' NOME USUARIO
    eRETU02 = Grid
    
    Unload Me
End Sub

Private Sub Filgrid()
cSQL = "select * from USUARIO order by USUARIO"
MontaGridFast Grid, 5, Array(1, 400, 3000, 2500, 2500), _
        Array("ID", "Nº", "Nome", "Ultimo Acesso", "Trocar a Senha Em"), _
        Array("ID", "R$IDUSUARIO", "L$USUARIO", "C$DATAULT", "C$TROCAR"), cARQ, cSQL
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    ''If KeyAscii > 31 And KeyAscii < 123 Then
        ''LocalizaGrid Grid, Chr(KeyAscii), 1, False
    ''End If
End Sub

Private Sub Form_Load()
   cARQ = PegPath("PATH", "SYSUSER")
    If AcessaForm(nFORMID) Then
        MontaToolbar Me.Toolbar1, nFORMID
        Filgrid
        If FixStr(eLOCALIZA) <> "" Then LocalizaGrid Grid, eLOCALIZA, 1
    Else
        MontaToolbar Me.Toolbar1, 999
    End If
   
    ''Configura Help
    lRETU = False
    HelpContextID = nFORMID
    Me.Caption = cFORMID
   
End Sub

Private Sub Grid_DblClick()
''Editar
End Sub

Private Sub Grid_KeyPress(KeyAscii As Integer)
If KeyAscii > 31 And KeyAscii < 123 Then
   LocalizaGrid Grid, Chr(KeyAscii), 1, False
End If
End Sub

Private Sub Grid_SelChange()
    With Grid
        If .Rows > 2 Then
            .Col = .Cols - 1
            .ColSel = 0
            .TopRow = .Row
        End If
    End With
End Sub

Private Sub incluir()
Dim nUSER As Long
Dim cSQL As String
   

    cSQL = "select IDUSUARIO from USUARIO ORDER BY IDUSUARIO"
    nUSER = FixNum(PegUltSQL(cARQ, cSQL, "IDUSUARIO", 0)) + 1
    ePASS01 = "select * from USUARIO WHERE IDUSUARIO=" & nUSER
    
    IncluiSQL cARQ, ePASS01, 2, Array("IDUSUARIO", "USUARIO"), Array(nUSER, " "), True
    
    ePASS02 = nUSER
    frmUSER.Show vbModal
    Filgrid
    LocalizaGrid Grid, nUSER, 1, False
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim sButton As String
Dim nINDICE As Integer
    
    sButton = Button
    sButton = Left(UCase(Replace(sButton, "&", "")), 3)
    
    If sButton = "SAI" Then
        Unload Me
    End If
    
    nINDICE = FixNum(Button.Index)
    If Not AcessaBotao(nFORMID, nINDICE) Then
        Exit Sub
    End If
    GravaLog nFORMID, nINDICE, sButton
    
    Select Case sButton
        Case "NOV"
            incluir
        Case "IMP"
            ''            cTIPO = "R"
            ''            zGRP = "CT"
            ''            escRPT.Show vbModal, Me
        Case "EDI"
            Editar
        Case "SEN"
            senha
        Case "ESC"
            Escolher
        Case "EXC"
            Apagar
        Case "LOC"
            LocalizaGrid Grid
    End Select
End Sub
