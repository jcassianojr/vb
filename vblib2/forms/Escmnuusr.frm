VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{F22668DE-E08D-467B-8E41-13900013BD5F}#2.7#0"; "VBextra2.OCX"
Object = "{2DA70529-3366-414A-B408-46083BCD481B}#1.8#0"; "VBFLXGRD18.OCX"
Begin VB.Form escmnuusr 
   Caption         =   "Liberação Usuario"
   ClientHeight    =   5040
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7800
   Icon            =   "Escmnuusr.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5040
   ScaleWidth      =   7800
   StartUpPosition =   2  'CenterScreen
   Begin VBFLXGRD18.VBFlexGrid gridrptusr 
      Height          =   3135
      Left            =   4800
      TabIndex        =   7
      Top             =   720
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   5530
   End
   Begin VBFLXGRD18.VBFlexGrid gridusr 
      Height          =   4335
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   7646
   End
   Begin XPControls.XPButton cmdSair 
      Height          =   435
      Left            =   3120
      TabIndex        =   2
      Top             =   4320
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "Escmnuusr.frx":038A
      Caption         =   "Retornar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton incusrrpt 
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   720
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Picture         =   "Escmnuusr.frx":0924
      Caption         =   "incluir"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton excrptusr 
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   1200
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Picture         =   "Escmnuusr.frx":0CBE
      Caption         =   "excluir"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin vbExtra.FlexFn FlexFn1 
      Height          =   405
      Left            =   3000
      TabIndex        =   5
      Top             =   120
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
   Begin VB.Label Label2 
      Caption         =   "Usuario Liberados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Usuario Disponiveis"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "escmnuusr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdSair_Click()

  Unload Me

End Sub

Private Sub excrptusr_Click()
  Dim eRETU As Variant
  Dim cdb As String
  Dim sSQL As String
  If Not MDG("Eliminar Acesso ao menu", "Confirme Exclusão") Then
    Exit Sub
  End If

  gridrptusr.Col = 0
  zIDRPTUSR = FixInt(gridrptusr)
  cdb = Dbname

  '******************************************************************
  'RPT
  sSQL = "select * from MENUUSU WHERE MENU='" & cMENU & "' AND INDICE=" & iMENU & " AND IDUSUARIO=" & zIDRPTUSR
  '******************************************************************

  eRETU = ApagaSQL(cdb, sSQL)
  filrptusr

End Sub

Private Sub filrptusr()
  Dim cNOME As String
  Dim DAODB As ADODB.Connection
  Dim daodb2 As ADODB.Connection
  Dim daors2 As New ADODB.Recordset
  Dim DAORS As New ADODB.Recordset
  Dim sSQL As String


  Set DAODB = New ADODB.Connection
  DAODB.CursorLocation = adUseClient
  DAODB.ConnectionTimeout = 120
  DAODB.Open GeracArq(dbuser, , False)

  Set daodb2 = New ADODB.Connection
  daodb2.CursorLocation = adUseClient
  daodb2.ConnectionTimeout = 120
  daodb2.Open GeracArq(Dbname, , False)

  sSQL = "select * from MENUUSU WHERE MENU='" & cMENU & "' AND INDICE=" & iMENU & " "
  daors2.Open sSQL, daodb2, adOpenForwardOnly, adLockReadOnly

  With gridrptusr
    .cols = 2
    .FixedCols = 0
    .Rows = 1
    .Row = 0
    .Col = 0
    .tEXT = "ID"
    .ColWidth(0) = 500
    .Col = 1
    .tEXT = "Nome"
    .ColWidth(1) = 2000
    If Not daors2.EOF Then
      daors2.MoveFirst
      While Not daors2.EOF
        cNOME = ""
        sSQL = "select USUARIO from USUARIO WHERE IDUSUARIO=" & daors2("IDUSUARIO")
        DAORS.Open sSQL, DAODB, adOpenForwardOnly, adLockReadOnly
        If Not DAORS.EOF Then
          cNOME = DAORS("USUARIO")
        End If
        DAORS.Close
        .AddItem daors2("IDUSUARIO") & vbTab & cNOME
        daors2.MoveNext
      Wend
    End If
  End With
  daors2.Close
  DAODB.Close
  daodb2.Close


End Sub

Private Sub filusr()

  Dim cARQ As String
  Dim cSQL As String
  cARQ = dbuser
  cSQL = "SELECT IDUSUARIO,USUARIO FROM USUARIO ORDER BY USUARIO"
  MontaGridUltra gridusr, 2, Array(500, 2000), Array("ID", "Nome"), _
                 Array("IDUSUARIO", "L$USUARIO"), cARQ, cSQL

End Sub

Private Sub Form_Load()
  CenterFormToScreen Me
  filusr
  filrptusr

End Sub

Private Sub gridrptusr_KeyPress(KeyAscii As Integer)
  If KeyAscii > 31 And KeyAscii < 123 Then
    LocalizaGrid gridrptusr, Chr(KeyAscii), 1, False
  End If


End Sub

'Private Sub gridrptusr_SelChange()
'    With gridrptusr
'        If .Rows > 2 Then
'            .Col = .Cols - 1
'            .ColSel = 0
'            .TopRow = .Row
'        End If
'    End With
'End Sub

Private Sub gridusr_KeyPress(KeyAscii As Integer)
  If KeyAscii > 31 And KeyAscii < 123 Then
    LocalizaGrid gridusr, Chr(KeyAscii), 1, False
  End If


End Sub

'Private Sub gridusr_SelChange()
'    With gridusr
'        If .Rows > 2 Then
'            .Col = .Cols - 1
'            .ColSel = 0
'            .TopRow = .Row
'        End If
'    End With
'End Sub

Private Sub incusrrpt_Click()
  Dim sSQL As String
  gridusr.Col = 0
  zIDRPTUSR = FixInt(gridusr)
  sSQL = "select * from MENUUSU WHERE MENU='" & cMENU & "' AND INDICE=" & iMENU & " AND IDUSUARIO=" & zIDRPTUSR
  IncluiSQL Dbname, sSQL, 5, Array("MENU", "INDICE", "IDUSUARIO", "LIGADO", "ATUALIZADO"), _
            Array(cMENU, iMENU, zIDRPTUSR, True, True), True, True
  filrptusr
End Sub


