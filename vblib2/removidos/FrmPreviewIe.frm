VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Begin VB.Form FrmPreview 
   Caption         =   "Preview"
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   8520
   Icon            =   "FrmPreviewIe.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6240
   ScaleWidth      =   8520
   StartUpPosition =   3  'Windows Default
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   5595
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   8205
      ExtentX         =   14473
      ExtentY         =   9869
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin XPControls.XPButton XPButton10 
      Height          =   435
      Left            =   600
      TabIndex        =   1
      Top             =   0
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   767
      Picture         =   "FrmPreviewIe.frx":058A
      Caption         =   "XPButton1"
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
   Begin XPControls.XPButton XPButton9 
      Height          =   435
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   767
      Picture         =   "FrmPreviewIe.frx":0B24
      Caption         =   "XPButton1"
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
   Begin XPControls.XPButton Salvar 
      Height          =   435
      Index           =   1
      Left            =   2040
      TabIndex        =   3
      Top             =   0
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "FrmPreviewIe.frx":10BE
      Caption         =   "Salvar &TXT"
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
   Begin XPControls.XPButton XPButton3 
      Height          =   435
      Left            =   1560
      TabIndex        =   4
      Top             =   0
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   767
      Picture         =   "FrmPreviewIe.frx":1458
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
      PictureAlignment=   4
   End
   Begin XPControls.XPButton XPButton5 
      Height          =   435
      Left            =   1080
      TabIndex        =   5
      Top             =   0
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   767
      Picture         =   "FrmPreviewIe.frx":19F2
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
      PictureAlignment=   4
   End
   Begin XPControls.XPButton Encerrar 
      Height          =   435
      Left            =   3720
      TabIndex        =   6
      Top             =   0
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "FrmPreviewIe.frx":1F8C
      Caption         =   "Retornar"
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
End
Attribute VB_Name = "FrmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Encerrar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'Center Me aqui ajusta a tela em vez de centralizar
    Me.Height = Screen.Height - 1200
    Me.Width = Screen.Width - 1200
    Me.Top = 100
    Me.Left = 100
' 0-Novo,1-Abrir,2-Salvar,3-Salvacomo
    ' 4-Editar,5-Visualizar,6-Imprimir,7-Exportar 8-Abrir Com
    If Not aDIREITOS(3) Then
       XPButton3.Enabled = False
       XPButton3.Visible = False
    End If
    If Not aDIREITOS(6) Then
       XPButton9.Enabled = False
       XPButton9.Visible = False
       XPButton5.Enabled = False
       XPButton5.Visible = False
       
    End If
     If Not aDIREITOS(5) Then
       XPButton10.Enabled = False
       XPButton10.Visible = False
    End If
      If Not aDIREITOS(2) Then
       salvar(1).Enabled = False
       salvar(1).Visible = False
    End If
    
    WebBrowser1.Navigate ePASS01
End Sub

Private Sub Salvar_Click(Index As Integer)
    Dim sFILTER, sFILENAME, CLINHA As String
    Dim nHANDLE
    eRETU01 = WebBrowser1.Document.Body.innerText
    CLINHA = WebBrowser1.Document.Body.innerText
    CLINHA = Html2Str(CStr(CLINHA))
    sFILTER = "TXT (*.TXT)" & vbNullChar & "*.TXT" & vbNullChar & "Todos Arquivo" & vbNullChar & "*.*"
    sFILENAME = FileSave(Me, sFILTER, 1, "TXT", "", "", "Salvar TXT Como")
    If Len(sFILENAME) = 0 Then
        Exit Sub
    End If
  
    nHANDLE = FreeFile
    Open sFILENAME For Output As #nHANDLE
    Print #nHANDLE, CLINHA
    Close (nHANDLE)

End Sub

Private Sub XPButton10_Click()
    WebBrowser1.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_PROMPTUSER
End Sub

Private Sub XPButton3_Click()
    WebBrowser1.ExecWB OLECMDID_SAVEAS, OLECMDEXECOPT_DODEFAULT
End Sub

Private Sub XPButton5_Click()
    WebBrowser1.ExecWB OLECMDID_PAGESETUP, OLECMDEXECOPT_DODEFAULT
End Sub

Private Sub XPButton9_Click()
    WebBrowser1.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_PROMPTUSER
End Sub

