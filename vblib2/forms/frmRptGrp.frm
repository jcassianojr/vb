VERSION 5.00
Begin VB.Form FrmRptGrp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Editando Relatorios/Documentos"
   ClientHeight    =   2280
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   8970
   HelpContextID   =   7
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   8970
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox Liberar 
      Caption         =   "Liberar Todos Usuários"
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
      Left            =   240
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Marque Para que todos usuarios acessem"
      Top             =   1440
      Width           =   3315
   End
   Begin VB.CommandButton CmdEscCam 
      Height          =   330
      Left            =   6120
      Picture         =   "frmRptGrp.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Escolher Nº Formulario Lista"
      Top             =   960
      Width           =   405
   End
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   2
      Left            =   240
      MaxLength       =   120
      TabIndex        =   6
      ToolTipText     =   "Dizer/Nome/Descriçao"
      Top             =   960
      Width           =   5895
   End
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   1
      Left            =   1800
      MaxLength       =   120
      TabIndex        =   1
      ToolTipText     =   "Dizer/Nome/Descriçao"
      Top             =   360
      Width           =   5895
   End
   Begin VB.TextBox Text 
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Codigo "
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Gravar"
      Height          =   990
      Left            =   7800
      Picture         =   "frmRptGrp.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Encerra e Grava"
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Caminho"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Nome"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Codigo"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   0
      Left            =   270
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "FrmRptGrp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const nFORMID = 1060
Const cFORMID = "Grupos Relatórios / Documentos"
Dim aCAM, aFOR, aVAL, aPAD As Variant
Dim cARQ, cSQL As String
Dim nITEM, nCAMPOS As Long

Private Sub cmdClose_Click()
 Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If MDG("Gravar e Sair", "Gravando ") Then
        For nITEM = 0 To 2
            aVAL(nITEM) = Text(nITEM)
        Next nITEM
        aVAL(3) = Liberar.Value
        GrvSQL cARQ, cSQL, nCAMPOS, aCAM, aVAL, aFOR
        GravaLog nFORMID, 0, "Gravou "
    Else
        GravaLog nFORMID, 0, "Nao Gravou "
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub CmdEscCam_Click()
Dim DLG As New clsCommonDialogs
DLG.BrowseFolder "Escolha uma Pasta", Pastas, False
If DLG.Cancel = True Then
   Exit Sub
End If
Text(2).Text = DLG.Path & "\"
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    TeclaEnter KeyCode
End Sub
Private Sub Form_Load()

   
   
   
    cARQ = PegPath("PATH", "SYSRPT")
    cSQL = ePASS01
    nCAMPOS = 4
   
    If FileExist(cARQ, True, , cSQL) Then


   
        ''Configura Help
        Me.Caption = cFORMID
        HelpContextID = nFORMID
   
        aCAM = Array("GRP", "NOME", "CAMINHO", "LIBERAR")
        aFOR = Array("CU", "C", "C", "BN")
        aPAD = Array("", "", "", False)
        aVAL = PegSQL(cARQ, cSQL, nCAMPOS, aCAM, aFOR, aPAD)
        For nITEM = 0 To 2
            Text(nITEM) = aVAL(nITEM)
        Next nITEM
        Liberar.Value = aVAL(3)
    End If
End Sub

