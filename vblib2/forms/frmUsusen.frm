VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{F22668DE-E08D-467B-8E41-13900013BD5F}#2.7#0"; "VBextra2.OCX"
Object = "{379157C5-E9BD-43F1-9F83-B037496BED42}#1.3#0"; "VBCCR18.OCX"
Begin VB.Form frmUSUSENHA 
   Caption         =   "Troca de Senha"
   ClientHeight    =   2268
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   5652
   Icon            =   "frmUsusen.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2268
   ScaleWidth      =   5652
   StartUpPosition =   2  'CenterScreen
   Begin vbExtra.SizeGrip SizeGrip1 
      Height          =   285
      Left            =   5424
      Top             =   2040
      Width           =   285
      _ExtentX        =   402
      _ExtentY        =   402
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   240
      Picture         =   "frmUsusen.frx":038A
      ScaleHeight     =   372
      ScaleWidth      =   612
      TabIndex        =   7
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   480
      Width           =   1095
   End
   Begin XPControls.XPButton cmdCANCELAR 
      Height          =   432
      Left            =   3840
      TabIndex        =   3
      Top             =   1200
      Width           =   1512
      _ExtentX        =   2667
      _ExtentY        =   762
      Picture         =   "frmUsusen.frx":07CC
      Caption         =   "Cancelar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VBCCR18.CommandButtonW cmdOK 
      Height          =   612
      Left            =   3840
      TabIndex        =   11
      Top             =   480
      Width           =   972
      _ExtentX        =   1715
      _ExtentY        =   1080
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ImageListAlignment=   1
      Caption         =   "Confimar"
      Alignment       =   0
      VerticalAlignment=   0
      Picture         =   "frmUsusen.frx":0D66
      PictureAndCaption=   -1  'True
      WordWrap        =   0   'False
      Style           =   1
   End
   Begin VB.Label Label4 
      Height          =   255
      Left            =   1440
      TabIndex        =   9
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Usuário:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Confirme a nova senha:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Novo senha:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Senha:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   4
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "frmUSUSENHA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CmdCancelar_Click()
  Unload Me
End Sub
Private Sub cmdOK_Click()
  Dim DB As ADODB.Connection
  Dim RSSENHA As ADODB.Recordset
  Dim strEncryptedText As String
  Dim sSQL As String
  eRETU01 = ""
  eRETU02 = ""
  eRETU03 = ""

  lRETU = True
  txtFields(0) = Trim(txtFields(0))
  txtFields(1) = Trim(txtFields(1))
  txtFields(2) = Trim(txtFields(2))


  If txtFields(0) = txtFields(1) Then
    Alert "Senha Precisa ser Diferente da Anterior"
    txtFields(1).Text = ""
    txtFields(2).Text = ""
    txtFields(1).SetFocus
    lRETU = False
    Exit Sub
  End If

  If Len(txtFields(1)) <> 8 Then
    Alert "8 digitos para a Senha"
    txtFields(1).Text = ""
    txtFields(2).Text = ""
    txtFields(1).SetFocus
    lRETU = False
    Exit Sub
  End If

  If Len(txtFields(1)) <> Len(txtFields(2)) Then
    Alert "Senha diferente da de confirmacao"
    txtFields(1).Text = ""
    txtFields(2).Text = ""
    txtFields(1).SetFocus
    lRETU = False
    Exit Sub
  End If

  'faze inicial acostumar os usuarios
  If Not CheckPass(txtFields(1)) Then
    If Not MDG("Senha Fraca , gravar mesmo assim") Then
      txtFields(1).Text = ""
      txtFields(2).Text = ""
      txtFields(1).SetFocus
      lRETU = False
      Exit Sub
    End If
  End If

  Set DB = New ADODB.Connection
  Set RSSENHA = New ADODB.Recordset

  DB.ConnectionTimeout = 120
  DB.Open GeracArq(dbuser)

  sSQL = "select * from USUARIO WHERE IDUSUARIO=" & zIDTEMP
  RSSENHA.Open sSQL, DB, adOpenKeyset, adLockOptimistic  'adOpenStatic

  With RSSENHA
    If Not .EOF Then
      strEncryptedText = XOREncryption(strCodeKey, txtFields(0))
      If strEncryptedText = "" & !Senha Then
        If txtFields(1) = txtFields(2) Then
          strEncryptedText = XOREncryption(strCodeKey, txtFields(1))
          RSSENHA("SENHA") = strEncryptedText
          RSSENHA("TROCAR") = Date + 90


          RSSENHA("CHAVEV") = UCase(CreateSHA256HashString(UCase(Trim(RSSENHA("usuario"))) + Trim(txtFields(1))))

          .Update
          eRETU01 = strEncryptedText
          eRETU02 = txtFields(1)
        Else
          Alert "Confirmação nao confere! A Senha nao foi alterada.", "Alteracao de senha"
        End If
      Else
        Alert "Senha nao confere! A Senha nao foi alterada.", "Alteracao de senha"
      End If
    End If
    .Close
  End With
  DB.Close
  Unload Me
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  TeclaEnter KeyCode
End Sub
Private Sub Form_Load()
  CenterFormToScreen Me
  Label4 = zIDTEMP
End Sub



