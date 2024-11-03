VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{F74FE4B4-2AF4-4F6A-B323-05690FDB6730}#1.3#0"; "OrdoWebView2.OCX"
Begin VB.Form FrmPreview 
   Caption         =   "FrmPreview"
   ClientHeight    =   5130
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12855
   LinkTopic       =   "Form1"
   ScaleHeight     =   5130
   ScaleWidth      =   12855
   StartUpPosition =   3  'Windows Default
   Begin OrdoWebView2.OrdoWebView OrdoWebView1 
      Height          =   4695
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   8281
      Search_URL      =   ""
      IsScriptEnabled =   0   'False
      UseEdgeFixedVersion=   -1  'True
   End
   Begin XPControls.XPButton Encerrar 
      Height          =   435
      Left            =   11520
      TabIndex        =   1
      Top             =   720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   767
      Picture         =   "FrmPreview.frx":0000
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
   Begin XPControls.XPButton cmdimp 
      Height          =   435
      Left            =   11520
      TabIndex        =   2
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   767
      Picture         =   "FrmPreview.frx":059A
      Caption         =   "Imprimir"
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
   Begin XPControls.XPButton cmdSavehtml 
      Height          =   375
      Left            =   11520
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Picture         =   "FrmPreview.frx":0B34
      Caption         =   "Html"
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
   Begin XPControls.XPButton cmdSaveTXT 
      Height          =   375
      Left            =   11520
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Picture         =   "FrmPreview.frx":10CE
      Caption         =   "TXT"
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
Option Explicit

Private Sub cmdimp_Click()
OrdoWebView1.showPrintDialog
End Sub

Private Sub cmdSavehtml_Click()
Dim sFileName As String
Dim sFILTER As String
  sFILTER = "Arquivos de Textos (*.HTML)" & vbNullChar & "*.HTML" & vbNullChar & "Todos Arquivo" & vbNullChar & "*.*"
  sFileName = FileSave(Me, sFILTER, 1, "HTML", , , "Salvar HTML Como")
  FileWrite sFileName, OrdoWebView1.GetInnerHTML
End Sub

Private Sub cmdSaveTXT_Click()
Dim sFileName As String
Dim sFILTER As String
  sFILTER = "Arquivos de Textos (*.TXT)" & vbNullChar & "*.TXT" & vbNullChar & "Todos Arquivo" & vbNullChar & "*.*"
  sFileName = FileSave(Me, sFILTER, 1, "HTML", , , "Salvar TXT Como")
  FileWrite sFileName, OrdoWebView1.GetInnerText
End Sub

Private Sub Encerrar_Click()
Unload Me
End Sub

Private Sub Form_Load()
  OrdoWebView1.NavigateToFile cARQRTF
End Sub





 
