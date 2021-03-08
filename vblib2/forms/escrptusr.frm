VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form escrptusr 
   Caption         =   "Liberação Usuario"
   ClientHeight    =   6030
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8700
   Icon            =   "escrptusr.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6030
   ScaleWidth      =   8700
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   855
      Left            =   3840
      Picture         =   "escrptusr.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton edirptusr 
      Caption         =   "Direitos"
      Height          =   795
      Left            =   3840
      Picture         =   "escrptusr.frx":0914
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton excrptusr 
      Caption         =   "Excluir"
      Height          =   735
      Left            =   3840
      Picture         =   "escrptusr.frx":0C9E
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton incusrrpt 
      Caption         =   "Incluir"
      Height          =   795
      Left            =   3840
      Picture         =   "escrptusr.frx":1028
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid gridusr 
      Height          =   5475
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   9657
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid gridrptusr 
      Height          =   5415
      Left            =   5160
      TabIndex        =   4
      Top             =   480
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   9551
      _Version        =   393216
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
      Left            =   4920
      TabIndex        =   3
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
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "escrptusr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cARQUSR As String
Private Sub CmdSair_Click()
    Unload Me
End Sub
Private Sub edirptusr_Click()
Dim nid As Integer
Dim cNOME As String
    gridrptusr.Col = 0
    nid = FixInt(gridrptusr)
    gridrptusr.Col = 1
    cNOME = FixStr(gridrptusr)
    ePASS01 = Array(nid, cNOME, ZRPTCOD, zRPTGRP, zRPT)
    frmRPTUSR.Show vbModal, Me
End Sub
Private Sub excrptusr_Click()
   Dim cSQL As String
   Dim USRID As Integer
   gridrptusr.Col = 0
   USRID = FixInt(gridrptusr)
   cSQL = "SELECT * FROM USURPT" & _
    " WHERE CODIGO='" & ZRPTCOD & "' AND GRP='" & zRPTGRP & "' AND RPT='" & zRPT & "' AND IDUSUARIO=" & USRID
    ApagaSQLP cARQUSR, cSQL
    filrptusr
End Sub

Private Sub filrptusr()
   Dim cSQL As String
    cSQL = "SELECT IDUSUARIO,USUARIO FROM USURPT" & _
    " WHERE CODIGO='" & ZRPTCOD & "' AND GRP='" & zRPTGRP & "' AND RPT='" & zRPT & "'" & _
    " ORDER BY USUARIO"
    MontaGridFast gridrptusr, 2, Array(500, 2000), Array("ID", "Usuario"), _
            Array("IDUSUARIO", "USUARIO"), cARQUSR, cSQL
End Sub
Private Sub filusr()
Dim cSQL As String
    cSQL = "SELECT IDUSUARIO,USUARIO FROM USUARIO ORDER BY USUARIO"
    MontaGridFast gridusr, 2, Array(500, 2000), Array("ID", "Nome"), _
            Array("IDUSUARIO", "USUARIO"), cARQUSR, cSQL
End Sub
Private Sub Form_Load()
    cARQUSR = PegPath("PATH", "SYSUSER")
    filusr
    filrptusr
End Sub
Private Sub gridrptusr_KeyPress(KeyAscii As Integer)
If KeyAscii > 31 And KeyAscii < 123 Then
   LocalizaGrid gridrptusr, Chr(KeyAscii), 1, False
End If
    If KeyAscii = vbKeyReturn Then
        edirptusr_Click
    End If
End Sub
Private Sub gridrptusr_SelChange()
    With gridrptusr
        If .Rows > 2 Then
            .Col = .Cols - 1
            .ColSel = 0
            .TopRow = .Row
        End If
    End With
End Sub
Private Sub gridusr_KeyPress(KeyAscii As Integer)
If KeyAscii > 31 And KeyAscii < 123 Then
   LocalizaGrid gridusr, Chr(KeyAscii), 1, False
End If
If KeyAscii = vbKeyReturn Then
   incusrrpt_Click
End If
End Sub
Private Sub gridusr_SelChange()
    With gridusr
        If .Rows > 2 Then
            .Col = .Cols - 1
            .ColSel = 0
            .TopRow = .Row
        End If
    End With
End Sub
Private Sub incusrrpt_Click()
   Dim cSQL As String
   Dim USRID As Integer
   Dim USRNOME As String
   gridusr.Col = 0
   USRID = FixInt(gridusr)
   gridusr.Col = 1
   USRNOME = FixStr(gridusr)
   cSQL = "SELECT * FROM USURPT" & _
    " WHERE CODIGO='" & ZRPTCOD & "' AND GRP='" & zRPTGRP & "' AND RPT='" & zRPT & "' AND IDUSUARIO=" & USRID
    IncluiSQL cARQUSR, cSQL, 5, Array("CODIGO", "GRP", "RPT", "IDUSUARIO", "USUARIO"), _
                            Array(ZRPTCOD, zRPTGRP, zRPT, USRID, USRNOME), _
                            True, False
   filrptusr
End Sub
