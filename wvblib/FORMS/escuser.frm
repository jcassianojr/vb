VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{F22668DE-E08D-467B-8E41-13900013BD5F}#2.7#0"; "VBextra2.OCX"
Begin VB.Form escuser 
   Caption         =   "Selecione o Usuario Desejado"
   ClientHeight    =   6075
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8475
   Icon            =   "escuser.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6075
   ScaleWidth      =   8475
   StartUpPosition =   2  'CenterScreen
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
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5535
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   9763
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   6075
      Left            =   7005
      TabIndex        =   0
      Top             =   0
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   10716
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      TextAlignment   =   1
      _Version        =   393216
   End
End
Attribute VB_Name = "escuser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ApagaNF_Click()
    Dim aRETU As Variant
    Dim sSQL As String
    If Not MDG("Apagar Usuario", "Confirme Exclusão") Then
        Exit Sub
    End If
    Grid.Col = 0
    zIDTEMP = Grid
    'USUARIO
    sSQL = "select * from USUARIO WHERE IDUSUARIO=" & zIDTEMP
    ApagaSQL dbuser, sSQL
    'USUCAD
    sSQL = "select * from USUCAD WHERE IDUSUARIO=" & zIDTEMP
    ApagaSQL dbuser, sSQL
    'MENU
    sSQL = "select * from MENUUSU WHERE IDUSUARIO=" & zIDTEMP
    ApagaSQL dbuser, sSQL
    ''  Apagar Direitos Relatorios
    'RTFUSR
    sSQL = "select * from RTFUSR WHERE IDUSUARIO=" & zIDTEMP
    ApagaSQL DBWRPT, sSQL
    'RPTUSR
    sSQL = "select * from RPTUSR WHERE IDUSUARIO=" & zIDTEMP
    ApagaSQL DBWRPT, sSQL
    'RPTFOLUSR
    sSQL = "select * from RPTFOLUSR WHERE IDUSUARIO=" & zIDTEMP
    ApagaSQL DBWRPT, sSQL
    'RPTINTUSR
    sSQL = "select * from RPTINTUSR WHERE IDUSUARIO=" & zIDTEMP
    ApagaSQL DBWRPT, sSQL
    FilRelat
End Sub

Private Sub CmdSair_Click()
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub EditPf_Click()
    Dim nFOLHATMP As Long
    Dim sSQL As String
    Dim cARQ As String
    Dim aRETU As Variant
    Dim X As Integer
 

    Grid.Col = 0
    zIDTEMP = FixInt(Grid)
    
    cARQ = dbuser
    sSQL = "select IDFOLHA from USUARIO WHERE IDUSUARIO=" & zIDTEMP
    nFOLHATMP = PegUltSQL(cARQ, sSQL, "IDFOLHA", 0)
    If demitido(nFOLHATMP) Then
        If MDG("Deseja exlui-lo") Then
            ApagaNF_Click
            Exit Sub
        End If
    End If
    
    ePASS01 = "select * from USUARIO WHERE IDUSUARIO=" & zIDTEMP
    frmUSER.Show vbModal
    FilRelat
    LocalizaGrid Grid, zIDTEMP, 1, False


End Sub

Private Sub FilRelat()
    Dim cARQ As String
    Dim cSQL As String
    cARQ = dbuser
    cSQL = "SELECT IDUSUARIO,USUARIO,IDFOLHA,DATAULT FROM USUARIO ORDER BY USUARIO"
    MontaGridUltra Grid, 4, Array(400, 1600, 800, 1200), Array("ID", "NOme", "Folha", "Ultimo"), _
        Array("IDUSUARIO", "USUARIO", "IDFOLHA", "DATAULT"), cARQ, cSQL
End Sub

Private Sub Form_Load()
    Center Me
    FilRelat
    If FixStr(eLOCALIZA) <> "" Then LocalizaGri1 Grid, eLOCALIZA, 1
    xmontatoolbar Me.Toolbar1, "escUSER", True
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

Private Sub NovoPF_Click()
    zIDTEMP = FixInt(PegMAXSQL(dbuser, "usuario", "idusuario", 0)) + 1
    IncluiSQL dbuser, "SELECT * FROM USUARIO WHERE IDUSUARIO=" * zIDTEMP, 2, Array("IDUSUARIO", "USUARIO"), _
        Array(zIDTEMP, CStr(zIDTEMP)), False, False
    FilRelat
End Sub

Private Sub SENHA_Click()

   If Grid.Row > 0 Then 'evitar quando clicado na linha de cabecario
     Grid.Col = 0
     zIDTEMP = Grid
     frmUSUSENHA.Show vbModal
     FilRelat
  End If

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Dim sButton          As String
    sButton = Button
    sButton = Left(UCase(Replace(sButton, "&", "")), 3)
    
    If Not AcessaBtnOld("escUSER", Button.Index) Then
        Exit Sub
    End If
    GravaLog 0, Button.Index, sButton, "escUSER"
    
    
    Grid.Col = 0
    If Grid.Row > 0 Then
        zIDTEMP = Grid
        ePASS02 = zIDTEMP
    End If
    

    Select Case sButton

    Case "NOV"
        NovoPF_Click

    Case "EDI"
        EditPf_Click

    Case "EXC"
        ApagaNF_Click

    Case "SEN"
        SENHA_Click

    Case "LOC"
        LocalizaGri1 Grid

    Case "SAI"
        CmdSair_Click

    End Select

End Sub


