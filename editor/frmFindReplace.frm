VERSION 5.00
Begin VB.Form frmFindReplace 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find and Replace"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command19 
      Height          =   315
      Left            =   3000
      Picture         =   "frmFindReplace.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Localizar"
      Top             =   240
      Width           =   495
   End
   Begin VB.CommandButton Command20 
      Height          =   315
      Left            =   3480
      Picture         =   "frmFindReplace.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Localizar Proxima"
      Top             =   240
      Width           =   495
   End
   Begin VB.CommandButton cmdReplaceAll 
      Caption         =   "Substituir Todas"
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton cmdReplace 
      Caption         =   "&Substituir"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   5160
      Picture         =   "frmFindReplace.frx":0714
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox txtReplaceWith 
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Top             =   960
      Width           =   1575
   End
   Begin VB.CheckBox chkCaseSensitive 
      Caption         =   "Diferenciar Maisculas Minusculas"
      Height          =   255
      Left            =   2040
      TabIndex        =   3
      Top             =   600
      Width           =   2655
   End
   Begin VB.CheckBox chkWholeWord 
      Caption         =   "Palavra Completa"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox txtfind 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Substiuir Por"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Localizar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   810
   End
   Begin VB.Menu MyMenu 
      Caption         =   "MyMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuFindNext 
         Caption         =   "Find Next"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu MyMenu1 
      Caption         =   "MyMenu1"
      Visible         =   0   'False
      Begin VB.Menu mnuFindNext1 
         Caption         =   "Find Next"
      End
      Begin VB.Menu mnuReplace 
         Caption         =   "Replace"
      End
      Begin VB.Menu mnuReplaceAll 
         Caption         =   "Replace All"
      End
      Begin VB.Menu mnuClose1 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "frmFindReplace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub cmdReplace_Click()
    findflags = chkWholeWord.Value * 2 + chkCaseSensitive.Value * 4
    FrmRTf.RichTextbox1.SelText = txtReplaceWith.Text
    Position = FrmRTf.RichTextbox1.Find(txtfind.Text, Position + 1, , findflags)
    
    If Position > 0 Then
        ''FrmRTf.RichTextbox1.SetFocus
        
    Else
''        MsgBox "Substituiçao Concluida " & Chr(34) & txtFind.Text & Chr(34) & " with " & _
''                Chr(34) & txtReplaceWith.Text & Chr(34)
    End If

End Sub
Private Sub cmdReplaceAll_Click()
    
    findflags = chkWholeWord.Value * 2 + chkCaseSensitive.Value * 4
    FrmRTf.RichTextbox1.SelText = txtReplaceWith.Text
    Position = FrmRTf.RichTextbox1.Find(txtfind.Text, Position + 1, , findflags)
    
    While Position > 0
        Screen.MousePointer = 11
        frmMain.StatusBar1.Panels(1).Text = "Replacing text, please wait..."
        FrmRTf.RichTextbox1.SelText = txtReplaceWith.Text
        Position = FrmRTf.RichTextbox1.Find(txtfind.Text, Position + 1, , findflags)
    Wend
    
    Screen.MousePointer = 0
    MsgBox "Substituicao Concluida " & Chr(34) & txtfind.Text & Chr(34) & " with " & _
            Chr(34) & txtReplaceWith.Text & Chr(34)
    
End Sub

Private Sub Command19_Click()
   FrmRTf.texto.Text = txtfind.Text
   achartf (0)
End Sub

Private Sub Command20_Click()
   FrmRTf.texto.Text = txtfind.Text
   achartf (FrmRTf.RichTextbox1.SelStart + 1)
End Sub

Private Sub Form_Load()
    txtfind.Text = FrmRTf.RichTextbox1.SelText
End Sub
Private Sub Form_Unload(Cancel As Integer)
    ''FrmRTf.RichTextbox1.SetFocus
End Sub
Public Sub achartf(nPOS)
Dim textfound As Long
    findflags = chkWholeWord.Value * 2 + chkCaseSensitive.Value * 4
''    FrmRTf.RichTextbox1.Find (FrmRTf.RichTextbox1.Text)
''    FrmRTf.RichTextbox1.SetFocus
    textfound = FrmRTf.RichTextbox1.Find(txtfind, nPOS, , findflags)
    If textfound = -1 Then
        MsgBox vbCr & "Texto Não Foi Encontrado", vbInformation
    End If
    ''FrmRTf.RichTextbox1.SetFocus
End Sub

