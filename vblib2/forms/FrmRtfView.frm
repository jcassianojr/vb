VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{451B73A5-1563-45D5-A6AC-7B2B7D30B778}#1.1#0"; "BSPrin10.ocx"
Object = "{379157C5-E9BD-43F1-9F83-B037496BED42}#1.1#0"; "vbccr18.ocx"
Begin VB.Form FrmRtfView 
   Caption         =   "FrmRtfView"
   ClientHeight    =   6315
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13170
   LinkTopic       =   "Form1"
   ScaleHeight     =   6315
   ScaleWidth      =   13170
   StartUpPosition =   3  'Windows Default
   Begin VBCCR18.RichTextBox RichTextbox1 
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   10610
      MultiLine       =   -1  'True
      ScrollBars      =   3
      TextRTF         =   "FrmRtfView.frx":0000
   End
   Begin BSPrinter.PrintPreview PrintPreview1 
      Left            =   12240
      Top             =   3240
      _ExtentX        =   1191
      _ExtentY        =   1191
   End
   Begin XPControls.XPButton cmdimp 
      Height          =   435
      Left            =   11760
      TabIndex        =   1
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   767
      Picture         =   "FrmRtfView.frx":0166
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
   Begin XPControls.XPButton Encerrar 
      Height          =   435
      Left            =   11760
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   767
      Picture         =   "FrmRtfView.frx":0700
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
   Begin XPControls.XPButton ImpDireto 
      Height          =   795
      Left            =   11760
      TabIndex        =   3
      Top             =   720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1402
      Picture         =   "FrmRtfView.frx":0C9A
      Caption         =   "Imprimir Direto (Porta)"
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
Attribute VB_Name = "FrmRtfView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim larqtxt As Boolean
Private Sub cmdimp_Click()
  PrintPreview1.ShowPreview
  RichTextbox1.SetFocus
End Sub


Private Sub Form_Load()
  larqtxt = False
  
  If IsExtensao(cARQRTF, "TXT") Or IsExtensao(cARQRTF, "ZPL ") Or IsExtensao(cARQRTF, "MAN") Then
    larqtxt = True
  End If
  If InStr(UCase(cARQRTF), ".LST") > 0 Then    ''.LST .LTS2 .LTS3... pode ser sequencial usando instr inves isextensao
    larqtxt = True
  End If

  ImpDireto.Visible = False
  
  If FileExists(cARQRTF) Then
    If IsExtensao(cARQRTF, "RTF") > 0 Then
      RichTextbox1.LoadFile cARQRTF, RtfLoadSaveFormatRTF
    Else
      If larqtxt Then
         ImpDireto.Visible = True
         RichTextbox1.LoadFile cARQRTF, RtfLoadSaveFormatText
      End If
    End If
  End If
  RichTextbox1.SetFocus

  PrintPreview1.AuxiliaryButtonVisible = PrintPreview1.PrinterExists("Microsoft Print to PDF")
  PrintPreview1.AuxiliaryButtonToolTipText = "Salvar como PDF"
  
End Sub

Private Sub ImpDireto_Click()
   FrmPicturePrinter.Show vbModal, Me
End Sub

Public Sub PrintPreview1_AuxiliaryButtonClick(UpdateReport As Boolean)
  PrintPreview1.ShowSaveToFile "Microsoft Print to PDF", "*.pdf"
  UpdateReport = False  ' we don't need to update the report in the Print preview window after this action (the default value of UpdateReport parameter is True)
End Sub

Public Sub PrintPreview1_PrepareReport(Cancel As Boolean)
  PrintPreview1.PrintRichTextBox RichTextbox1
End Sub
