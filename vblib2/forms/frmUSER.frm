VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmUSER 
   Caption         =   "Cadastro de Usuário"
   ClientHeight    =   5370
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   10515
   HelpContextID   =   12
   Icon            =   "frmUSER.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5370
   ScaleWidth      =   10515
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   5085
      Left            =   90
      TabIndex        =   17
      Top             =   90
      Width           =   10275
      Begin VB.Frame Frame12 
         Caption         =   "Buscar Id Envolvidos"
         Height          =   795
         Left            =   300
         TabIndex        =   30
         Top             =   3210
         Visible         =   0   'False
         Width           =   1935
         Begin VB.TextBox IdForm 
            Height          =   285
            Left            =   90
            TabIndex        =   31
            Top             =   300
            Width           =   1035
         End
         Begin VB.Label Label3 
            Caption         =   "do Form"
            Height          =   315
            Left            =   1200
            TabIndex        =   32
            Top             =   330
            Width           =   1695
         End
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00FF8080&
         Caption         =   "&Gravar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   7440
         Picture         =   "frmUSER.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Encerra e Grava"
         Top             =   3840
         Width           =   1575
      End
      Begin VB.TextBox tEXT 
         Enabled         =   0   'False
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
         Index           =   3
         Left            =   3420
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Equivalente"
         Top             =   2880
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox tEXT 
         Alignment       =   1  'Right Justify
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
         Index           =   5
         Left            =   1410
         TabIndex        =   8
         ToolTipText     =   "Nº Folha Pagamento"
         Top             =   2100
         Width           =   735
      End
      Begin VB.TextBox tEXT 
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
         Index           =   6
         Left            =   2970
         MaxLength       =   50
         TabIndex        =   9
         ToolTipText     =   "Nome do Usuario"
         Top             =   2100
         Width           =   3375
      End
      Begin VB.CommandButton escidfolha 
         Caption         =   "..."
         Height          =   255
         Left            =   2250
         TabIndex        =   26
         TabStop         =   0   'False
         ToolTipText     =   "Escolher Nº Folha"
         Top             =   2100
         Width           =   615
      End
      Begin VB.ComboBox cboEQUIVALENTE 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1950
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Equivalente"
         Top             =   2850
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "-->"
         Height          =   255
         Left            =   2250
         TabIndex        =   25
         TabStop         =   0   'False
         ToolTipText     =   "Retorna o Nome conforme o Nº"
         Top             =   2340
         Width           =   615
      End
      Begin VB.CommandButton cmdZeraSenha 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Zerar Senha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   4215
         Picture         =   "frmUSER.frx":0914
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Zerar Senha"
         Top             =   3840
         Width           =   1575
      End
      Begin VB.TextBox tEXT 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
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
         Index           =   2
         Left            =   8205
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Senha"
         Top             =   2850
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdTroca 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Trocar Senha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   5835
         Picture         =   "frmUSER.frx":0C1E
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Trocar Senha"
         Top             =   3840
         Width           =   1575
      End
      Begin VB.CheckBox chkweekend 
         Caption         =   "Acesso no fim-de-semana"
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
         Height          =   645
         Left            =   8130
         TabIndex        =   7
         ToolTipText     =   "Se o Usuario acesso final de semana"
         Top             =   1080
         Width           =   1950
      End
      Begin VB.Frame Frame2 
         Caption         =   "Horarios Permitidos de Uso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   1035
         Left            =   5250
         TabIndex        =   22
         Top             =   780
         Width           =   2715
         Begin MSComCtl2.DTPicker DTPicker3 
            Height          =   330
            Left            =   1485
            TabIndex        =   5
            Top             =   255
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   582
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
            CustomFormat    =   "00:00"
            Format          =   21233666
            CurrentDate     =   38397
         End
         Begin MSComCtl2.DTPicker DTPicker4 
            Height          =   330
            Left            =   1485
            TabIndex        =   6
            Top             =   645
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   582
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
            CustomFormat    =   "00:00"
            Format          =   21233666
            CurrentDate     =   0.999988425925926
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Hora inicial :"
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
            Index           =   12
            Left            =   180
            TabIndex        =   24
            Top             =   300
            Width           =   1215
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Hora final :"
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
            Index           =   13
            Left            =   270
            TabIndex        =   23
            Top             =   690
            Width           =   1125
         End
      End
      Begin VB.CheckBox chkAtivo 
         Caption         =   "Ativo"
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
         Height          =   285
         Left            =   300
         TabIndex        =   3
         ToolTipText     =   "Ativo - Acessa/Bloqueia Sistema"
         Top             =   1245
         Width           =   2250
      End
      Begin VB.TextBox tEXT 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
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
         Index           =   0
         Left            =   1290
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         ToolTipText     =   "Nº Usuario"
         Top             =   420
         Width           =   495
      End
      Begin VB.TextBox tEXT 
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
         Index           =   1
         Left            =   2610
         MaxLength       =   10
         TabIndex        =   1
         ToolTipText     =   "Nome Usuario"
         Top             =   420
         Width           =   2295
      End
      Begin VB.TextBox tEXT 
         Alignment       =   2  'Center
         Enabled         =   0   'False
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
         Index           =   4
         Left            =   5910
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Ultimo Acesso"
         Top             =   435
         Width           =   1365
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2790
         TabIndex        =   4
         Tag             =   "Expiraçao"
         ToolTipText     =   "Expiraçao"
         Top             =   1200
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   661
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
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   21233665
         CurrentDate     =   36516
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   6000
         TabIndex        =   12
         ToolTipText     =   "Data Troca Senha"
         Top             =   2790
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
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
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   21233665
         CurrentDate     =   36516
      End
      Begin VB.Label lblLabels 
         Caption         =   "Id Folha"
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
         Index           =   10
         Left            =   570
         TabIndex        =   29
         Top             =   2100
         Width           =   855
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Usuário Equivalente"
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
         Index           =   5
         Left            =   30
         TabIndex        =   28
         Top             =   2910
         Width           =   1845
      End
      Begin VB.Label lblLabels 
         Caption         =   "Data Para Troca Senha"
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
         Height          =   315
         Index           =   6
         Left            =   3900
         TabIndex        =   27
         Top             =   2850
         Width           =   2385
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Caption         =   "Data Limite Para Acessar o Sistema"
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
         Index           =   4
         Left            =   1800
         TabIndex        =   21
         Top             =   960
         Width           =   3165
      End
      Begin VB.Label lblLabels 
         Caption         =   "Nrº Usuário"
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
         Index           =   0
         Left            =   1110
         TabIndex        =   20
         Top             =   210
         Width           =   1035
      End
      Begin VB.Label lblLabels 
         Caption         =   "Cognome do Usuário"
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
         Index           =   1
         Left            =   2670
         TabIndex        =   19
         Top             =   210
         Width           =   2265
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Caption         =   "Último Acesso em :"
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
         Index           =   2
         Left            =   5700
         TabIndex        =   18
         Top             =   180
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmUSER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const nFORMID = 1012
Const cFORMID = "Cadastro de Usuários"
Dim aCAM, aFOR, aVAL, aPAD As Variant
Dim cARQ, cSQL As String
Dim nITEM, nCAMPOS As Long

Private Sub cboEQUIVALENTE_GotFocus()
FocusMe
End Sub

Private Sub cboEQUIVALENTE_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Or KeyCode = 40 Then
        Visual
        SendKeys Chr(9)
    End If
    If KeyCode = 38 Then
        Visual
        SendKeys "+" + Chr(9)
    End If

End Sub

Private Sub chkAtivo_Click()
Visual
End Sub

Private Sub chkAtivo_GotFocus()
FocusMe
End Sub

Private Sub chkAtivo_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Or KeyCode = 40 Then
        Visual
        SendKeys Chr(9)
    End If
    If KeyCode = 38 Then
        Visual
        SendKeys "+" + Chr(9)
    End If

End Sub

Private Sub chkweekend_Click()
Visual
End Sub

Private Sub chkweekend_GotFocus()
FocusMe
End Sub

Private Sub chkweekend_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Or KeyCode = 40 Then
        Visual
        SendKeys Chr(9)
    End If
    If KeyCode = 38 Then
        Visual
        SendKeys "+" + Chr(9)
    End If

End Sub

Private Sub cmdClose_Click()
 Unload Me
End Sub

Private Sub DTPicker1_GotFocus()
FocusMe
End Sub

Private Sub DTPicker1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Or KeyCode = 40 Then
        Visual
        SendKeys Chr(9)
    End If
    If KeyCode = 38 Then
        Visual
        SendKeys "+" + Chr(9)
    End If

End Sub

Private Sub DTPicker2_GotFocus()
FocusMe
End Sub

Private Sub DTPicker2_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Or KeyCode = 40 Then
        Visual
        SendKeys Chr(9)
    End If
    If KeyCode = 38 Then
        Visual
        SendKeys "+" + Chr(9)
    End If

End Sub

Private Sub DTPicker3_GotFocus()
FocusMe
End Sub

Private Sub DTPicker3_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Or KeyCode = 40 Then
        Visual
        SendKeys Chr(9)
    End If
    If KeyCode = 38 Then
        Visual
        SendKeys "+" + Chr(9)
    End If

End Sub

Private Sub DTPicker4_GotFocus()
FocusMe
End Sub

Private Sub DTPicker4_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Or KeyCode = 40 Then
        Visual
        SendKeys Chr(9)
    End If
    If KeyCode = 38 Then
        Visual
        SendKeys "+" + Chr(9)
    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If CDbl(Text(5)) <> 0 Then
    If Len(Trim(Text(2))) = 0 Then
      If MDG("Senha nao Colocada deseja Colocar ? ", "Gravar Senha") Then
          zIDTEMP = Text(0)
          frmUSUSENHA.Show vbModal, Me
          If lRETU Then
             Text(2) = eRETU01
             DTPicker2.Value = eRETU02
          End If
      End If
    End If

       
    If MDG("Gravar e Sair", "Gravando Usuarios") Then
        For nITEM = 0 To 6    'Array comeca 0
            aVAL(nITEM) = Text(nITEM)
        Next nITEM
        aVAL(9) = FixNumBol(chkAtivo)
        aVAL(10) = FixNumBol(chkweekend)
        aVAL(11) = DTPicker1
        aVAL(12) = DTPicker2
        aVAL(7) = DTPicker3.Hour & ":" & DTPicker3.Minute
        aVAL(8) = DTPicker4.Hour & ":" & DTPicker4.Minute
      
        GrvSQL cARQ, cSQL, nCAMPOS, aCAM, aVAL, aFOR
        GravaLog nFORMID, 0, "Gravou Usuario"
    Else
        GravaLog nFORMID, 0, "Nao Gravou Usuario"
    End If
Else
    Alert "Coloque o numero do seu Funcionario que esta cadastrado no Aplifolha"
    Cancel = 1
End If
    
    Screen.MousePointer = vbDefault
End Sub
Private Sub Command1_Click()
    Alert "Implementaçao Futura"
End Sub
Private Sub cmdTroca_Click()
    zIDTEMP = Text(0)
    frmUSUSENHA.Show vbModal, Me
    If lRETU Then
       Text(2) = eRETU01
       DTPicker2.Value = eRETU02
    End If
    
End Sub
Private Sub cmdZeraSenha_Click()
    Text(2) = " "
    cmdClose_Click
End Sub

Private Sub escidfolha_Click()
    Alert "Implementaçao Futura"
    ''   escMP04.Show vbModal, Me

    ''   If lRETUESC Then

    ''      frmUSER.txtFields(10) = eRETU01
    ''      frmUSER.txtFields(11) = eRETU02

    ''   End If

End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
     TeclaEnter (KeyCode)
End Sub
Private Sub Form_Load()
    'Configura Help
    Me.Caption = cFORMID
    HelpContextID = nFORMID
   
    ''Arquivo
    cARQ = PegPath("PATH", "SYSUSER")
    If Not FileExist(cARQ, True) Then
        End
    End If
   
    'Variaveis sql
    cSQL = ePASS01
    idForm = ePASS02
    zIDTEMP = ePASS02

    nCAMPOS = 13
    aCAM = Array("IDUSUARIO", "USUARIO", "SENHA", "EQUIVALENTE", _
            "DATAULT", "IDFOLHA", "NOMEFOLHA", "HORAINI", "HORAFIM", _
            "ATIVO", "WEEKEND", "DATAVAL", "TROCAR")
    aFOR = Array("N", "C", "C", "C", "D", "N", "C", "", "", "BN", "BN", "DC", "DC")
    aPAD = Array(0, "", "", "", "", 0, "", "07:00", "18:00", False, False, Today() + 30, Today() + 60)
    aVAL = PegSQL(cARQ, cSQL, nCAMPOS, aCAM, aFOR, aPAD)
    For nITEM = 0 To 6    '' array comeca 0
        Text(nITEM) = aVAL(nITEM)
    Next nITEM
    If aVAL(7) <> "" Then DTPicker3 = aVAL(7)
    If aVAL(8) <> "" Then DTPicker4 = aVAL(8)
     
    chkAtivo = FixBolNum(aVAL(9))
    chkweekend = FixBolNum(aVAL(10))
    If IsDate(aVAL(11)) Then
       DTPicker1 = aVAL(11)
    End If
    If IsDate(aVAL(12)) Then
       DTPicker2 = aVAL(12)
    End If
    Visual
End Sub
Private Sub Text_GotFocus(Index As Integer)
FocusMe
End Sub
Private Sub Text_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Or KeyCode = 40 Then
        Visual
        SendKeys Chr(9)
    End If
    If KeyCode = 38 Then
        Visual
        SendKeys "+" + Chr(9)
    End If
End Sub
Public Sub Visual()
If chkAtivo.Value = 1 Then
   chkAtivo.Caption = "Usuário Ativo"
Else
   chkAtivo.Caption = "Usuário Desativado"
End If
If chkweekend.Value = 1 Then
   chkweekend.Caption = "Pode Acessar no fim-de-semana"
Else
   chkweekend.Caption = "Não Pode Acessar no fim-de-semana"
End If
End Sub
