VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F22668DE-E08D-467B-8E41-13900013BD5F}#2.7#0"; "VBextra2.OCX"
Begin VB.Form EscRptExec 
   Caption         =   "Selecione o Desejado"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9105
   HelpContextID   =   6
   Icon            =   "escRptExec.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   9105
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   5775
      Left            =   7605
      TabIndex        =   1
      ToolTipText     =   "Botoes de Açoes"
      Top             =   0
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   10186
      ButtonWidth     =   714
      ButtonHeight    =   714
      Appearance      =   1
      TextAlignment   =   1
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Lista de Registros"
      Top             =   450
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   9128
      _Version        =   393216
   End
   Begin vbExtra.FlexFn FlexFn1 
      Height          =   405
      Left            =   240
      TabIndex        =   2
      Top             =   0
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   714
      PageNumbersFormat=   ""
      PageNumbersFormatIndex=   0
      BeginProperty PageNumbersFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "EscRptExec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const nFORMID = 1063
Const cFORMID = "Gerenciados Relatorios/Documentos Tipos Execuçao"
Dim cSQLREF As String

Private Sub Apagar()
    Dim nNUMERO As Long
    Dim cARQ, cSQL  As String
    Grid.Col = 0
    nNUMERO = FixNum(Grid)
    cARQ = PegPath("PATH", "SYSCONF")
    cSQL = "select * from RPTEXEC WHERE NUMERO=" & nNUMERO
    lRETU = ApagaSQLP(cARQ, cSQL, "Apagar")
    If lRETU Then
        Filgrid
    End If
End Sub

Private Sub Editar()
    Dim nNUMERO As Long
    Grid.Col = 0
    nNUMERO = FixNum(Grid)
    ePASS01 = "select * from RPTEXEC WHERE NUMERO=" & nNUMERO
    frmRptExec.Show vbModal
    Filgrid
End Sub

Private Sub Escolher()
    lRETU = True
    Grid.Col = 0                                 'Numero
    eRETU01 = Grid
    Grid.Col = 2                                 'Executar
    eRETU02 = Grid
    Unload Me
End Sub

Private Sub Filgrid()
    Dim cARQ As String
    Dim cSQL As String
    cARQ = PegPath("PATH", "SYSCONF")
    cSQL = cSQLREF
    '    MontaGridFast Grid, 4, Array(400, 500, 2000, 3000), Array("Nº", "Ext", "Executar", "NOME"), _
    Array("NUMERO", "EXTENSAO", "EXECUTAR", "NOME"), cARQ, cSQL
    MontaGridFast Grid, 4, Array(400, 500, 2000, 3000), Array("Nº", "Ext", "Executar", "NOME"), _
        Array("NUMERO", "EXTENSAO", "EXECUTAR", "NOME"), cARQ, cSQL
End Sub

Private Sub Form_Load()
    CenterFormToScreen Me
    If AcessaForm(nFORMID) Then
        MontaToolbar Me.Toolbar1, nFORMID
        
        If FixStr(ePASS01) <> "" Then
            cSQLREF = "select NUMERO,EXTENSAO,EXECUTAR,NOME from RPTEXEC where EXTENSAO='" & ePASS01 & "' order by EXTENSAO"
        Else
            cSQLREF = "select NUMERO,EXTENSAO,EXECUTAR,NOME from RPTEXEC order by EXTENSAO"
        End If
        Filgrid
        If FixStr(eLOCALIZA) <> "" Then LocalizaGrid Grid, eLOCALIZA, 3
    Else
        MontaToolbar Me.Toolbar1, 999
    End If
    ''Configura Help
    lRETU = False
    HelpContextID = nFORMID
    Me.Caption = cFORMID
End Sub

Private Sub Grid_KeyPress(KeyAscii As Integer)
    If KeyAscii > 31 And KeyAscii < 123 Then
        LocalizaGrid Grid, Chr(KeyAscii), 1, False
    End If

End Sub

'Private Sub Grid_SelChange()
'    With Grid
'        If .Rows > 2 Then
'            .Col = .Cols - 1
'            .ColSel = 0
'            .TopRow = .Row
'        End If
'    End With
'End Sub

Private Sub incluir()
    Dim nNUMERO As Long
    Dim cARQ As String
    
    cARQ = PegPath("PATH", "SYSCONF")
    nNUMERO = FixInt(PegMAXSQL(cARQ, "RPTEXEC", "NUMERO", 0)) + 1
    nNUMERO = FixNum(Val(Busca("Nº NUMERO", "Incluindo NUMERO", Str(nNUMERO), 3)))
    ePASS01 = "select * from RPTEXEC WHERE NUMERO=" & nNUMERO
    IncluiSQL cARQ, ePASS01, 1, Array("NUMERO"), Array(nNUMERO), True
    frmRptExec.Show vbModal
  
    Filgrid
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim sButton As String
    Dim nINDICE As Integer
    
    sButton = Button
    sButton = Left(UCase(Replace(sButton, "&", "")), 3)
    
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
    Case "ESC"
        Escolher
    Case "EXC"
        Apagar
    Case "LOC"
        LocalizaGrid Grid
    Case "SAI"
        Screen.MousePointer = vbDefault
        Unload Me
    End Select
End Sub


