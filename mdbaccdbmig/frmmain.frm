VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Comparador de Estrutura Access (MDB vs ACCDB)"
   ClientHeight    =   4140
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7500
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   4140
   ScaleWidth      =   7500
   Begin VB.CommandButton CmdGeraPasta 
      Caption         =   "GERAR SCRIPT DE CORREÇÃO Pasta"
      Height          =   615
      Left            =   2880
      TabIndex        =   13
      Top             =   3240
      Width           =   3255
   End
   Begin VB.CommandButton CmdEscdir 
      Caption         =   "..."
      Height          =   372
      Left            =   6840
      TabIndex        =   12
      Top             =   2280
      Width           =   492
   End
   Begin VB.TextBox Diretorio 
      Height          =   372
      Left            =   240
      TabIndex        =   11
      Top             =   2280
      Width           =   5652
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Efetua  correcao"
      Height          =   372
      Left            =   240
      TabIndex        =   10
      Top             =   2760
      Width           =   1932
   End
   Begin VB.CommandButton cmdProcessar 
      Caption         =   "GERAR SCRIPT DE CORREÇÃO"
      Height          =   615
      Left            =   240
      TabIndex        =   8
      Top             =   3240
      Width           =   2415
   End
   Begin VB.CommandButton cmdSelSQL 
      Caption         =   "..."
      Height          =   315
      Left            =   6840
      TabIndex        =   7
      Top             =   1800
      Width           =   495
   End
   Begin VB.TextBox txtSQL 
      Height          =   315
      Left            =   240
      TabIndex        =   6
      Top             =   1800
      Width           =   6495
   End
   Begin VB.CommandButton cmdSelACCDB 
      Caption         =   "..."
      Height          =   315
      Left            =   6840
      TabIndex        =   4
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox txtACCDB 
      Height          =   315
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   6495
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   6840
      Top             =   3240
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
      TabIndex        =   5
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label lblACCDB 
      Caption         =   "Arquivo Destino (.ACCDB):"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label lblMDB 
      Caption         =   "Arquivo Origem (.MDB):"
      Height          =   255
      Left            =   240
      TabIndex        =   9
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

Private Sub cmdSelMDB_Click()

End Sub

Private Sub CmdEscdir_Click()
Dim sCaminho As String
    sCaminho = SelecionarPasta("Selecione a pasta contendo dos arquivos MDB:")
    
    If sCaminho <> "" Then
        Diretorio.Text = sCaminho
    End If
End Sub

Private Sub CmdGeraPasta_Click()
If Diretorio.Text = "" Then
        MsgBox "Por favor, selecione o diretório dos arquivos MDB.", vbExclamation
        Exit Sub
    End If
    
    ' Executa a nova rotina do módulo
    Call ProcessarPastaCompleta(Diretorio.Text, Check1.Value)
End Sub

Private Sub cmdSelMDB_Fix_Click()
    On Error Resume Next
    dlg.Filter = "Access Antigo (*.mdb)|*.mdb"
    dlg.ShowOpen
    If dlg.FileName <> "" Then
        txtMDB.Text = dlg.FileName
        
        ' Lógica para trocar a extensão para .accdb no destino
        Dim sCaminho As String
        sCaminho = dlg.FileName
        
        ' Verifica se termina com .mdb (case insensitive) e troca
        If LCase(Right(sCaminho, 4)) = ".mdb" Then
            txtACCDB.Text = Left(sCaminho, Len(sCaminho) - 4) & ".accdb"
            
            ' Opcional: Sugerir também o caminho do script SQL
            txtSQL.Text = Left(sCaminho, Len(sCaminho) - 4) & "_correcao.sql"
        End If
    End If
    
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
    Call GerarScriptCorrecao(txtMDB.Text, txtACCDB.Text, txtSQL.Text, Check1.Value)
End Sub

' Declaração para a Shell API no topo do módulo
Public Function SelecionarPasta(ByVal Titulo As String) As String
    Dim objShell As Object
    Dim objFolder As Object
    
    Set objShell = CreateObject("Shell.Application")
    ' 1 = BIF_RETURNONLYFSDIRS (Apenas pastas do sistema de arquivos)
    Set objFolder = objShell.BrowseForFolder(0, Titulo, 1)
    
    If Not objFolder Is Nothing Then
        SelecionarPasta = objFolder.Items.Item.Path
    Else
        SelecionarPasta = ""
    End If
    
    Set objFolder = Nothing
    Set objShell = Nothing
End Function
