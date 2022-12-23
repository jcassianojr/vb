VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{F22668DE-E08D-467B-8E41-13900013BD5F}#2.7#0"; "VBextra2.OCX"
Begin VB.Form frmUSUSENHA 
   Caption         =   "Troca de Senha"
   ClientHeight    =   1800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4980
   Icon            =   "frmUsusen.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1800
   ScaleWidth      =   4980
   StartUpPosition =   2  'CenterScreen
   Begin vbExtra.SizeGrip SizeGrip1 
      Height          =   285
      Left            =   4695
      Top             =   1515
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   240
      Picture         =   "frmUsusen.frx":038A
      ScaleHeight     =   375
      ScaleWidth      =   615
      TabIndex        =   6
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
   Begin XPControls.XPButton cmdCANCELAR 
      Height          =   435
      Left            =   3120
      TabIndex        =   9
      Top             =   960
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "frmUsusen.frx":07CC
      Caption         =   "Cancelar"
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
   Begin XPControls.XPButton cmdOK 
      Height          =   435
      Left            =   3120
      TabIndex        =   10
      Top             =   360
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "frmUsusen.frx":0D66
      Caption         =   "Confirmar"
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
   Begin VB.Label Label4 
      Height          =   255
      Left            =   1200
      TabIndex        =   8
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Usuário:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Confirme a nova senha:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Novo senha:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Senha:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   3
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

    If Len(txtFields(1)) > 8 Then
        Alert "Máximo 8 digitos para a Senha", "Senha Invalida"
        txtFields(1).tEXT = ""
        txtFields(1).SetFocus
        lRETU = False
        Exit Sub
    End If

    If Len(txtFields(2)) > 8 Then
        Alert "Máximo 8 digitos para a Senha", "Senha Invalida"
        txtFields(2).tEXT = ""
        txtFields(2).SetFocus
        lRETU = False
        Exit Sub
    End If

    If Len(txtFields(1)) < 8 Then
        Alert "Necessário 8 digitos para a Senha", "Senha Invalida"
        txtFields(1).tEXT = ""
        txtFields(1).SetFocus
        lRETU = False
        Exit Sub
    End If

    If Len(txtFields(2)) < 8 Then
        Alert "Necessário 8 digitos para a Senha", "Senha Invalida"
        txtFields(2).tEXT = ""
        txtFields(2).SetFocus
        lRETU = False
        Exit Sub
    End If

    If txtFields(0) = txtFields(1) Then
        Alert "Senha Precisa ser Diferente da Anterior", "Senha Invalida"
        txtFields(1).tEXT = ""
        txtFields(2).tEXT = ""
        txtFields(1).SetFocus
        lRETU = False
        Exit Sub
    End If
    
    Set DB = New ADODB.Connection
    Set RSSENHA = New ADODB.Recordset

    DB.ConnectionTimeout = 120
    DB.Open GeracArq(dbuser)
    
    sSQL = "select * from USUARIO WHERE IDUSUARIO=" & zIDTEMP
    RSSENHA.Open sSQL, DB, adOpenKeyset, adLockOptimistic 'adOpenStatic

    With RSSENHA
        If Not .EOF Then
            strEncryptedText = XOREncryption(strCodeKey, txtFields(0))
            If strEncryptedText = "" & !senha Then
                If txtFields(1) = txtFields(2) Then
                    strEncryptedText = XOREncryption(strCodeKey, txtFields(1))
                    RSSENHA("SENHA") = strEncryptedText
                    RSSENHA("TROCAR") = Date + 60
                    .Update
                    eRETU01 = strEncryptedText
                    eRETU02 = txtFields(1)
                Else
                    Alert "Confirmação não confere! A Senha não foi alterada.", "Alteração de senha"
                End If
            Else
                Alert "Senha não confere! A Senha não foi alterada.", "Alteração de senha"
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

