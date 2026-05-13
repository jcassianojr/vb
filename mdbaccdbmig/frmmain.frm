VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Comparador de Estrutura Access (MDB vs ACCDB)"
   ClientHeight    =   3405
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7500
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
   EndProperty
   Begin VB.CommandButton cmdProcessar 
      Caption         =   "GERAR SCRIPT DE CORREÇÃO"
      Height          =   615
      Left            =   2280
      TabIndex        =   9
      Top             =   2520
      Width           =   3015
   End
   Begin VB.CommandButton cmdSelSQL 
      Caption         =   "..."
      Height          =   315
      Left            =   6840
      TabIndex        =   8
      Top             =   1800
      Width           =   495
   End
   Begin VB.TextBox txtSQL 
      Height          =   315
      Left            =   240
      TabIndex        =   7
      Top             =   1800
      Width           =   6495
   End
   Begin VB.CommandButton cmdSelACCDB 
      Caption         =   "..."
      Height          =   315
      Left            =   6840
      TabIndex        =   5
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox txtACCDB 
      Height          =   315
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   6495
   End
   Begin VB.CommandButton cmdSelMDB 
      Caption         =   "..."
      Height          =   315
      Left            =   6840
      TabIndex        =   2
      Top             =   3600
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   6840
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSelMDB_Fix 
      Caption         =   "..."
      Height          =   315
      Left            =   6840
      TabIndex        =   1
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox txtMDB 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   6495
   End
   Begin VB.Label lblSQL 
      Caption         =   "Salvar Script SQL em:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label lblACCDB 
      Caption         =   "Arquivo Destino (.ACCDB):"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label lblMDB 
      Caption         =   "Arquivo Origem (.MDB):"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSelMDB_Fix_Click()
    On Error Resume Next
    dlg.Filter = "Access Antigo (*.mdb)|*.mdb"
    dlg.ShowOpen
    If dlg.FileName <> "" Then txtMDB.Text = dlg.FileName
End Sub

Private Sub cmdSelACCDB_Click()
    On Error Resume Next
    dlg.Filter = "Access Novo (*.accdb)|*.accdb"
    dlg.ShowOpen
    If dlg.FileName <> "" Then txtACCDB.Text = dlg.FileName
End Sub

Private Sub cmdSelSQL_Click()
    On Error Resume Next
    dlg.Filter = "Script SQL (*.sql)|*.sql"
    dlg.FileName = "correcao_indices.sql"
    dlg.ShowSave
    If dlg.FileName <> "" Then txtSQL.Text = dlg.FileName
End Sub

Private Sub cmdProcessar_Click()
    If txtMDB.Text = "" Or txtACCDB.Text = "" Or txtSQL.Text = "" Then
        MsgBox "Selecione todos os caminhos antes de prosseguir.", vbExclamation
        Exit Sub
    End If
    
    ' Chama a rotina do módulo .bas
    Call GerarScriptCorrecao(txtMDB.Text, txtACCDB.Text, txtSQL.Text)
End Sub