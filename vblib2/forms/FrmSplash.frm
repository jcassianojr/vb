VERSION 5.00
Begin VB.Form FrmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4845
   ClientLeft      =   2250
   ClientTop       =   2235
   ClientWidth     =   7260
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   HelpContextID   =   2
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7275
      Begin VB.PictureBox Picture1 
         Height          =   1515
         Left            =   3600
         Picture         =   "FrmSplash.frx":0000
         ScaleHeight     =   1455
         ScaleWidth      =   3315
         TabIndex        =   9
         Top             =   990
         Width           =   3375
      End
      Begin VB.Timer Timer1 
         Interval        =   3000
         Left            =   0
         Top             =   0
      End
      Begin VB.Image Image1 
         Height          =   1455
         Left            =   300
         Stretch         =   -1  'True
         Top             =   1020
         Width           =   3105
      End
      Begin VB.Label lblversao 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Versao"
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
         Left            =   5625
         TabIndex        =   8
         Top             =   225
         Width           =   1455
      End
      Begin VB.Label lblsistema 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sistema"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   525
         Left            =   735
         TabIndex        =   7
         Top             =   375
         Width           =   5595
      End
      Begin VB.Label lblsysopr 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Windows 9x Acima"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   525
         Left            =   3300
         TabIndex        =   6
         Top             =   3420
         Width           =   3795
      End
      Begin VB.Label lblcopyright 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Copyright:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3630
         TabIndex        =   5
         Top             =   4050
         Width           =   3345
      End
      Begin VB.Label lbl6 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Atenção! Copia e Distribuição           Não Autorizada."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   4440
         Width           =   2895
      End
      Begin VB.Label lblinfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fone/Fax:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3660
         TabIndex        =   3
         Top             =   4560
         Width           =   3405
      End
      Begin VB.Label lblcompany 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Aplisoft"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   675
         Index           =   0
         Left            =   75
         TabIndex        =   2
         Top             =   2640
         Width           =   7095
      End
      Begin VB.Label lblcompany 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Aplisoft"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   735
         Index           =   1
         Left            =   90
         TabIndex        =   1
         Top             =   3600
         Width           =   3105
      End
   End
End
Attribute VB_Name = "FrmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const nFORMID = 501
Const cFORMID = "" '"Tela de Apresentaçao"

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Left = (Screen.Width - Me.Width) / 2
    lblversao = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblsistema = App.ProductName
    lblcompany(0) = App.FileDescription
    lblcompany(1) = App.CompanyName
    lblcopyright = App.LegalCopyright
    lblinfo = App.Comments
    Me.Caption = cFORMID
    HelpContextID = nFORMID

    'Carrega imagem background
    If FileExist(PegPath("PATH", "LOGO")) Then
        Me.Picture = LoadPicture(PegPath("PATH", "LOGO"))
        Image1.Picture = LoadPicture(PegPath("PATH", "LOGO"))
    End If



End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

Private Sub imgLogo_Click()
    Unload Me
End Sub

Private Sub lblcompany_Click(Index As Integer)
    Unload Me
End Sub

Private Sub lblcopyright_Click()
    Unload Me
End Sub

Private Sub lblinfo_Click()
    Unload Me
End Sub

Private Sub lblsistema_Click()
    Unload Me
End Sub

Private Sub lblsysopr_Click()
    Unload Me
End Sub

Private Sub Picture1_Click()
    Unload Me
End Sub

Private Sub Timer1_Timer()
    Unload Me
End Sub
