VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Begin VB.Form FrmReparar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Arquivo Para Reparacao"
   ClientHeight    =   3150
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   9630
   HelpContextID   =   8
   Icon            =   "frmReparar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   9630
   ShowInTaskbar   =   0   'False
   Begin XPControls.XPText Text 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   503
      Text            =   ""
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
   Begin XPControls.XPButton cmdReparar 
      Height          =   495
      Index           =   0
      Left            =   1680
      TabIndex        =   4
      Top             =   2040
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Picture         =   "frmReparar.frx":058A
      Caption         =   "&Reparar"
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
   Begin XPControls.XPButton cmdReparar 
      Height          =   495
      Index           =   1
      Left            =   3240
      TabIndex        =   5
      Top             =   2040
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      Picture         =   "frmReparar.frx":0B24
      Caption         =   "Reparar e Compactar"
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
   Begin XPControls.XPButton CmdEscCam 
      Height          =   375
      Left            =   7440
      TabIndex        =   3
      Top             =   1080
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      Picture         =   "frmReparar.frx":10BE
      Caption         =   ""
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
   Begin XPControls.XPButton cmdclose 
      Height          =   495
      Left            =   8040
      TabIndex        =   6
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Picture         =   "frmReparar.frx":1658
      Caption         =   "&Gravar"
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
   Begin XPControls.XPButton cmdcancelar 
      Height          =   495
      Left            =   8040
      TabIndex        =   7
      Top             =   840
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Picture         =   "frmReparar.frx":1BF2
      Caption         =   "&Retornar"
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
   Begin XPControls.XPText Text 
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   503
      Text            =   ""
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
   Begin XPControls.XPText Text 
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   360
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   503
      Text            =   ""
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
   Begin VB.Label Label1 
      Caption         =   "Especificação"
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
      Left            =   120
      TabIndex        =   2
      Top             =   1425
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Arquivo Nome"
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
      Index           =   3
      Left            =   120
      TabIndex        =   1
      Top             =   45
      Width           =   1365
   End
   Begin VB.Label Label1 
      Caption         =   "Caminho"
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
      Left            =   120
      TabIndex        =   0
      Top             =   765
      Width           =   1095
   End
End
Attribute VB_Name = "FrmReparar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const nFORMID = 1018
Const cFORMID = "Cadastro de Arquivos"
Dim aCAM As Variant
Dim aFOR As Variant
Dim aVAL As Variant
Dim aPAD As Variant
Dim cARQ As String
Dim cSQL As String
Dim nITEM As Long
Dim nCAMPOS As Long
Private Sub cmdcancelar_Click()
Unload Me
End Sub
Private Sub cmdClose_Click()
 Unload Me
End Sub
Private Sub CmdEscCam_Click()
Dim sFileName As String
Dim sPath As String
Dim sRECENTFILE As String
Dim sFILTER As String
   If Len(Text(1).Text) > 0 Then
      sRECENTFILE = Text(1).Text
   Else
      sPath = App.path
   End If
   sFILTER = "Arquivo Access MDB" & vbNullChar & "*.mdb" & vbNullChar
   sFileName = FileOpen(Me, sFILTER, 1, sRECENTFILE, "mdb", sPath, "Escolher Arquivo Access MDB")
   If Len(sFileName) = 0 Then
      Exit Sub
   End If
   Text(1).Text = sFileName
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If MDG("Gravar e Sair", "Gravando Reparar") Then
        For nITEM = 0 To nCAMPOS - 1
            aVAL(nITEM) = Text(nITEM).Text
        Next nITEM
        GrvSQL cARQ, cSQL, nCAMPOS, aCAM, aVAL, aFOR
    End If
    Screen.MousePointer = vbDefault
End Sub
Private Sub CmdReparar_Click(Index As Integer)
    cARQRTF = Caminex(Text(1).Text)
    If MDG("Reparar" & cARQRTF) Then
        If Index = 0 Then
            Reparacao cARQRTF
        Else
            Reparacao cARQRTF, True
            
        End If
    End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   TeclaEnter KeyCode
End Sub
Private Sub Form_Load()
   
    ''Configura Help
    Me.Caption = cFORMID
    HelpContextID = nFORMID
    ''Arquivo
    cARQ = PegPath("PATH", "REPARQ")
    If Not FileExist(cARQ, True) Then
        End
    End If
   
    'Variaveis sql
    cSQL = ePASS01
    nCAMPOS = 3
    aCAM = Array("DIZER", "CAMINHO", "ARQUIVO")
    aFOR = Array("C", "CU", "CU")
    aPAD = Array("", "", "")
    aVAL = PegSQL(cARQ, cSQL, nCAMPOS, aCAM, aFOR, aPAD)
    For nITEM = 0 To nCAMPOS - 1
        Text(nITEM).Text = aVAL(nITEM)
    Next nITEM
      
   
   
End Sub

