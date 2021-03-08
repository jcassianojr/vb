VERSION 5.00
Begin VB.Form FrmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5055
   ClientLeft      =   2250
   ClientTop       =   2235
   ClientWidth     =   7275
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   HelpContextID   =   2
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   7275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000000&
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Width           =   7275
      Begin VB.Timer Timer1 
         Interval        =   3000
         Left            =   0
         Top             =   0
      End
      Begin VB.Label lblversao 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   60
         TabIndex        =   5
         Top             =   780
         Width           =   7155
      End
      Begin VB.Label lblsistema 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sistema"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   60
         TabIndex        =   4
         Top             =   180
         Width           =   7155
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
         Height          =   495
         Left            =   3120
         TabIndex        =   3
         Top             =   4380
         Width           =   4065
      End
      Begin VB.Label lbl6 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Atenção! Copia e Distribuição     Não Autorizada."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   60
         TabIndex        =   2
         Top             =   4380
         Width           =   2985
      End
      Begin VB.Label lblcompany 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   3165
         Index           =   0
         Left            =   60
         TabIndex        =   1
         Top             =   1140
         Width           =   7125
      End
   End
End
Attribute VB_Name = "FrmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const nFORMID = 1002

Const cFORMID = "Tela de Apresentaçao"

Private Sub Form_KeyPress(KeyAscii As Integer)

    Unload Me

End Sub

Private Sub Form_Load()

    'CentralizaJanela Me
    Center Me
    lblversao = "Versao: " & App.Major & "." & App.Minor & "." & App.Revision
    lblsistema = App.ProductName
    lblcompany(0) = App.CompanyName
    '    lblcompany(1) = App.CompanyName
    lblcopyright = App.LegalCopyright & " " & Year(Date)
    'lblinfo = App.Comments
    Me.Caption = cFORMID
    HelpContextID = nFORMID

End Sub

Private Sub Frame1_Click()

    Unload Me

End Sub

Private Sub Timer1_Timer()

    Unload Me

End Sub

