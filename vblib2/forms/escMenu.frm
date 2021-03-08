VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form EscMenu 
   Caption         =   "Selecione o Desejado"
   ClientHeight    =   8040
   ClientLeft      =   60
   ClientTop       =   210
   ClientWidth     =   11880
   HelpContextID   =   6
   Icon            =   "escMenu.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8040
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   8040
      Left            =   10380
      TabIndex        =   1
      ToolTipText     =   "Botoes de Açoes"
      Top             =   0
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   14182
      ButtonWidth     =   714
      ButtonHeight    =   714
      Appearance      =   1
      TextAlignment   =   1
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   7845
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Lista de Registros"
      Top             =   90
      Width           =   10110
      _ExtentX        =   17833
      _ExtentY        =   13838
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
Attribute VB_Name = "EscMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const nFORMID = 1006
Const cFORMID = "Menus do Sistema"
Dim nMODULOID As Long
Dim lMODULO As Boolean
Private Sub Apagar()
Dim nINDICE As Long
Dim cARQ, cSQL, cMENU As String
    Grid.Col = 0
    cMENU = Grid
    Grid.Col = 1
    nINDICE = FixNum(Grid)
    If Not lMODULO Then
      Grid.Col = 2
      nMODULOID = FixNum(Grid)
    End If
    cARQ = PegPath("PATH", "SYSCONF")
    cSQL = "select * from MENUS WHERE MENU='" & cMENU & "' AND INDICE=" & nINDICE & " AND MODULO=" & nMODULOID
    lRETU = ApagaSQLP(cARQ, cSQL, "Apagar")
    If lRETU Then
        Filgrid
    End If
End Sub
Private Sub Editar()
Dim nINDICE As Long
Dim cMENU As String
    Grid.Col = 0
    cMENU = Grid
    Grid.Col = 1
    nINDICE = FixNum(Grid)
    If Not lMODULO Then
      Grid.Col = 2
      nMODULOID = FixNum(Grid)
    End If
    ePASS01 = "select * from MENUS WHERE MENU='" & cMENU & "' AND INDICE=" & nINDICE & " AND MODULO=" & nMODULOID
    frmMenu.Show vbModal
    Filgrid
    LocalizaGrid Grid, cMENU, 1, False
End Sub
Private Sub Escolher()
    lRETU = True
    Grid.Col = 0
    eRETU01 = Grid
    Grid.Col = 1
    eRETU02 = Grid
    Grid.Col = 2
    eRETU03 = Grid
    Unload Me
End Sub

Private Sub Filgrid()
Dim cARQ As String
Dim cSQL As String
    cARQ = PegPath("PATH", "SYSCONF")
    If lMODULO Then
       cSQL = "select * from menus where modulo=" & nMODULOID & " order by menu,indice"
    Else
       cSQL = "select * from menus order by menu,indice"
    End If
    MontaGridFast Grid, 5, Array(1400, 800, 800, 800, 6000), Array("Menu", "Indice", "Modulo", "Form", "Dizer"), _
            Array("MENU", "INDICE", "MODULO", "FORM", "DIZER"), cARQ, cSQL
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    ''If KeyAscii > 31 And KeyAscii < 123 Then
    ''    LocalizaGrid Grid, Chr(KeyAscii), 1, False
    ''End If
End Sub
Private Sub Form_Load()
    
    nMODULOID = ePASS01
    lMODULO = True
    If nMODULOID = 0 Then
       lMODULO = False
    End If
    
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
Dim nINDICE As Long
Dim cARQ, cMENU As String
    cMENU = LCase(Busca("Codigo do Menu", "Incluindo Menu", Space(3), 10))
    nINDICE = FixNum(Val(Busca("Indice do Menu", "Incluindo Menu", "1", 5)))
    cMENU = UCase(cMENU)
    cARQ = PegPath("PATH", "SYSCONF")
    If Not lMODULO Then
      ePASS01 = "select * from MENUS WHERE MENU='" & cMENU & "' AND INDICE=" & nINDICE
      IncluiSQL cARQ, ePASS01, 2, Array("MENU", "INDICE"), Array(cMENU, nINDICE), True
    Else
      ePASS01 = "select * from MENUS WHERE MENU='" & cMENU & "' AND INDICE=" & nINDICE & " AND MODULO=" & nMODULOID
      IncluiSQL cARQ, ePASS01, 3, Array("MENU", "INDICE", "MODULO"), Array(cMENU, nINDICE, nMODULOID), True
    End If
    frmMenu.Show vbModal
    Filgrid
    LocalizaGrid Grid, cMENU, 1, False
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
        Case "ESC"
            Escolher
        Case "EXC"
            Apagar
        Case "LOC"
            LocalizaGrid Grid
    End Select
End Sub
