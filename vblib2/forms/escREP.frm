VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form escREP 
   Caption         =   "Escolha o Codigo da Imagem"
   ClientHeight    =   5550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9615
   HelpContextID   =   1001
   Icon            =   "escREP.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   9615
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5385
      Left            =   180
      TabIndex        =   0
      ToolTipText     =   "Listas de Imagens"
      Top             =   120
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   9499
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   5550
      Left            =   8025
      TabIndex        =   1
      ToolTipText     =   "Botoes de Açoes"
      Top             =   0
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   9790
      ButtonWidth     =   714
      ButtonHeight    =   714
      Appearance      =   1
      TextAlignment   =   1
      _Version        =   393216
   End
End
Attribute VB_Name = "escREP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const nFORMID = 1017
Const cFORMID = "Cadastro de Arquivos"

Private Sub Apagar()
Dim cARQ, cSQL, cARQUIVO As String
    Grid.Col = 0
    cARQUIVO = Grid
    cARQ = PegPath("PATH", "REPARQ")
    cSQL = "select * from REPARAR WHERE ARQUIVO='" & cARQUIVO & "'"
    If ApagaSQLP(cARQ, cSQL) Then
        FilRelat
    End If
End Sub

Private Sub Documentar()
    Grid.Col = 2
    cARQRTF = Caminex(Grid)
    FrmTableStru.Show vbModal, Me
End Sub

Private Sub Editar()
Dim cARQUIVO As String
    Grid.Col = 0
    cARQUIVO = Grid
    ePASS01 = "select * from REPARAR WHERE ARQUIVO='" & cARQUIVO & "'"
    FrmReparar.Show vbModal
    FilRelat
    LocalizaGrid Grid, cARQUIVO, 1, False
End Sub

Private Sub Escolher()
    lRETU = True
    Grid.Col = 0
    eRETU01 = Grid
    Unload Me
End Sub

Private Sub FilRelat()
Dim cARQ As String
Dim cSQL As String
    cARQ = PegPath("PATH", "REPARQ")
    cSQL = "SELECT ARQUIVO,DIZER,CAMINHO FROM REPARAR ORDER BY ARQUIVO"
    MontaGridFast Grid, 3, Array(2000, 4000, 1), Array("Arquivo", "Especificaçao", "C"), _
            Array("ARQUIVO", "DIZER", "CAMINHO"), cARQ, cSQL
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    ''If KeyAscii > 31 And KeyAscii < 123 Then
    ''    LocalizaGrid Grid, Chr(KeyAscii), 1, False
    ''End If
End Sub

Private Sub Form_Load()
    If AcessaForm(nFORMID) Then
        MontaToolbar Me.Toolbar1, nFORMID
        FilRelat
        If FixStr(eLOCALIZA) <> "" Then LocalizaGrid Grid, eLOCALIZA, 1
    Else
        MontaToolbar Me.Toolbar1, 999
    End If
    lRETU = False
    Me.Caption = cFORMID
    Me.HelpContextID = nFORMID
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
Dim cARQ As String
Dim cARQDB As String
Dim cCODNEW As String
Dim sFileName As String
Dim sPath As String
Dim sRECENTFILE As String
Dim sFILTER As String
sPath = App.path
sFILTER = "Microsoft Database" & vbNullChar & "*.mdb" & vbNullChar
sFileName = FileOpen(Me, sFILTER, 1, sRECENTFILE, "mdb", sPath, "Abrir Arquivo Access MDB")
If Len(sFileName) = 0 Then
   Alert "Arquivo nao Escolhido"
   Exit Sub
End If
cARQ = sFileName

    cCODNEW = NomeArq(cARQ)
    
    cARQDB = PegPath("PATH", "REPARQ")
    ePASS01 = "select * from reparar WHERE ARQUIVO='" & cCODNEW & "'"
    
    IncluiSQL cARQDB, ePASS01, 2, Array("ARQUIVO", "CAMINHO"), Array(cCODNEW, cARQ), True
    
    FrmReparar.Show vbModal
    FilRelat

    LocalizaGrid Grid, cCODNEW, 1, False
End Sub

Private Sub Reparar()
Dim cARQUIVO As String
    Grid.Col = 2
    cARQUIVO = Caminex(Grid)
    If MDG("Reparar" & cARQUIVO) Then
        If MDG("Compactar") Then
            Reparacao cARQUIVO, True
        Else
            Reparacao cARQUIVO
        End If
    End If
End Sub

Private Sub Sair()
    Screen.MousePointer = vbDefault
    Unload Me
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
            ZRPTCOD = "SYS"
            zRPTGRP = "RPT"
            escRpt.Show vbModal, Me
        Case "REP"
            Reparar
        Case "EDI"
            Editar
        Case "ESC"
            Escolher
        Case "EXC"
            Apagar
        Case "DOC"
            Documentar
        Case "LOC"
            LocalizaGrid Grid
    End Select

End Sub
