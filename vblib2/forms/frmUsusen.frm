VERSION 5.00
Begin VB.Form frmUSUSENHA 
   Caption         =   "Senha de Usuario"
   ClientHeight    =   2670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4590
   HelpContextID   =   13
   Icon            =   "frmUsusen.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2670
   ScaleWidth      =   4590
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCANCELAR 
      BackColor       =   &H00C0C0FF&
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   1020
      Picture         =   "frmUsusen.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Cancelar Troca de Senha"
      Top             =   1590
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   2640
      Picture         =   "frmUsusen.frx":0C54
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Confirme Troca Senha"
      Top             =   1590
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   285
      Picture         =   "frmUsusen.frx":0F5E
      ScaleHeight     =   375
      ScaleWidth      =   615
      TabIndex        =   6
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox txtFields 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   2340
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   1
      ToolTipText     =   "Confirme Nova Senha"
      Top             =   1050
      Width           =   1875
   End
   Begin VB.TextBox txtFields 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   2340
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   0
      ToolTipText     =   "Nova Senha"
      Top             =   630
      Width           =   1875
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1410
      TabIndex        =   8
      ToolTipText     =   "Nome do Usuario"
      Top             =   210
      Width           =   2685
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Usuário:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   285
      TabIndex        =   7
      Top             =   210
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Confirme a nova senha:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1095
      Width           =   2145
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Nova senha:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   945
      TabIndex        =   4
      Top             =   675
      Width           =   1290
   End
End
Attribute VB_Name = "frmUSUSENHA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const nFORMID = 1013
Const cFORMID = "Usuario Trocar Senha"
Private Sub CmdCancelar_Click()
    lRETU = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
Dim cSENHA, cCRYPTO, cARQ, cSQL As String
Dim aRETU As Variant
    lRETU = False
   

    txtFields(1) = UCase(Trim(txtFields(1)))
    txtFields(2) = UCase(Trim(txtFields(2)))

    If Len(txtFields(1)) > 8 Then

        Alert "Máximo 8 digitos para a Senha", "Senha Invalida"
        txtFields(1).Text = ""
        txtFields(1).SetFocus
        Exit Sub
    End If

    If Len(txtFields(2)) > 8 Then

        Alert "Máximo 8 digitos para a Senha", "Senha Invalida"
        txtFields(2).Text = ""
        txtFields(2).SetFocus
        Exit Sub

    End If

    If Len(txtFields(1)) < 8 Then

        Alert "Necessário 8 digitos para a Senha", "Senha Invalida"
        txtFields(1).Text = ""
        txtFields(1).SetFocus
        Exit Sub

    End If

    If Len(txtFields(2)) < 8 Then

        Alert "Necessário 8 digitos para a Senha", "Senha Invalida"
        txtFields(2).Text = ""
        txtFields(2).SetFocus
        Exit Sub

    End If

   
    cARQ = PegPath("PATH", "SYSUSER")
    cSQL = "select * from USUARIO WHERE IDUSUARIO=" & zIDTEMP
    aRETU = PegSQL(cARQ, cSQL, 1, Array("SENHA"), Array("C"), Array(Space(8)))
    cSENHA = aRETU(0)
    cCRYPTO = XOREncryption(SysCodeKey, txtFields(1))
   
    If cSENHA = cCRYPTO Then
        Alert "Senha Precisa ser Diferente da Anterior", "Senha Invalida"
        txtFields(1).Text = ""
        txtFields(2).Text = ""
        txtFields(1).SetFocus
        Exit Sub
    End If
   
    GrvSQL cARQ, cSQL, 2, Array("SENHA", "TROCAR"), Array(cCRYPTO, Date + 90), Array("C", "D")

    lRETU = True
    eRETU01 = cCRYPTO
    eRETU02 = Date + 90
   
    Unload Me
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   TeclaEnter KeyCode
End Sub

Private Sub Form_Load()
    frmUSUSENHA.Left = (Screen.Width - frmUSUSENHA.Width) / 2
    frmUSUSENHA.Top = (Screen.Height - frmUSUSENHA.Height) / 2

    cARQ = PegPath("PATH", "SYSUSER")
    cSQL = "select * from USUARIO WHERE IDUSUARIO=" & zIDTEMP
    aRETU = PegSQL(cARQ, cSQL, 1, Array("USUARIO"), Array("C"), Array(Space(8)))
    Label4.Caption = aRETU(0)

    lRETU = False
    ''Configura Help
    HelpContextID = nFORMID
    Me.Caption = cFORMID
   
End Sub


Private Sub txtFields_GotFocus(Index As Integer)
FocusMe
End Sub

Private Sub txtFields_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Or KeyCode = 40 Then
        SendKeys Chr(9)
    End If
    If KeyCode = 38 Then
        SendKeys "+" + Chr(9)
    End If
End Sub
