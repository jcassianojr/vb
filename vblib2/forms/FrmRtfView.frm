VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{451B73A5-1563-45D5-A6AC-7B2B7D30B778}#1.1#0"; "BSPrin10.ocx"
Object = "{379157C5-E9BD-43F1-9F83-B037496BED42}#1.3#0"; "VBCCR18.OCX"
Begin VB.Form FrmRtfView 
   Caption         =   "FrmRtfView"
   ClientHeight    =   6312
   ClientLeft      =   60
   ClientTop       =   408
   ClientWidth     =   13176
   LinkTopic       =   "Form1"
   ScaleHeight     =   6312
   ScaleWidth      =   13176
   StartUpPosition =   3  'Windows Default
   Begin VBCCR18.RichTextBox RichTextbox1 
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11415
      _ExtentX        =   20130
      _ExtentY        =   10605
      MultiLine       =   -1  'True
      ScrollBars      =   3
      TextRTF         =   "FrmRtfView.frx":0000
   End
   Begin BSPrinter.PrintPreview PrintPreview1 
      Left            =   12480
      Top             =   5040
      _ExtentX        =   953
      _ExtentY        =   953
   End
   Begin XPControls.XPButton ImpDireto 
      Height          =   792
      Left            =   11760
      TabIndex        =   1
      Top             =   960
      Width           =   1212
      _ExtentX        =   2138
      _ExtentY        =   1397
      Picture         =   "FrmRtfView.frx":0166
      Caption         =   "Imprimir Direto (Porta)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton SalvarComoTxt 
      Height          =   372
      Left            =   11880
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3840
      Width           =   792
      _ExtentX        =   1397
      _ExtentY        =   656
      Picture         =   "FrmRtfView.frx":0700
      Caption         =   "TXT"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton SalvarComoHTML 
      Height          =   372
      Left            =   11880
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4320
      Width           =   852
      _ExtentX        =   1503
      _ExtentY        =   656
      Picture         =   "FrmRtfView.frx":0A9A
      Caption         =   "Html"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton CmdSalvarComoPdf 
      Height          =   372
      Left            =   11880
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4680
      Width           =   792
      _ExtentX        =   1397
      _ExtentY        =   656
      Picture         =   "FrmRtfView.frx":1034
      Caption         =   "PDF"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton SalvarComoRTF 
      Height          =   372
      Left            =   11880
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5640
      Width           =   792
      _ExtentX        =   1397
      _ExtentY        =   656
      Picture         =   "FrmRtfView.frx":13CE
      Caption         =   "RTF"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton Cmdsalvadoc 
      Height          =   372
      Left            =   11880
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5160
      Width           =   792
      _ExtentX        =   1397
      _ExtentY        =   656
      Picture         =   "FrmRtfView.frx":1768
      Caption         =   "doc"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VBCCR18.CommandButtonW retornar 
      Height          =   612
      Left            =   11760
      TabIndex        =   7
      Top             =   1680
      Width           =   972
      _ExtentX        =   1715
      _ExtentY        =   1080
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ImageListAlignment=   1
      Caption         =   "Retornar"
      Alignment       =   0
      VerticalAlignment=   0
      Picture         =   "FrmRtfView.frx":1B02
      PictureAndCaption=   -1  'True
      WordWrap        =   0   'False
      Style           =   1
   End
   Begin VBCCR18.CommandButtonW cmdPrint 
      Height          =   612
      Left            =   11760
      TabIndex        =   8
      Top             =   240
      Width           =   972
      _ExtentX        =   1715
      _ExtentY        =   1080
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ImageListAlignment=   1
      Caption         =   "Imprimir"
      Alignment       =   0
      VerticalAlignment=   0
      Picture         =   "FrmRtfView.frx":209C
      PictureAndCaption=   -1  'True
      WordWrap        =   0   'False
      Style           =   1
   End
End
Attribute VB_Name = "FrmRtfView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim larqtxt As Boolean


Private Sub cmdPrint_Click()
  PrintPreview1.ShowPreview
  RichTextbox1.SetFocus

End Sub

Private Sub Cmdsalvadoc_Click()
 Dim sFILTER As String
  Dim sFileName As String
  Dim stmp As String
  stmp = TrocaExt(cARQRTF, "DOC")
  sFILTER = "Arquivos de Textos (*.DOC)" & vbNullChar & "*.DOC" & vbNullChar & "Todos Arquivo" & vbNullChar & "*.*"
  sFileName = FileSave(Me, sFILTER, 1, "DOC", stmp, stmp, "Salvar DOC Como")
  If Not Len(sFileName) = 0 Then
    If FileConnExist(sFileName) Then
      If MDG("Arquivo ja Existe Sobrepor") Then
        DeleteFile sFileName  'Kill sFILENAME
        Exit Sub
      End If
    End If
    stmp = TrocaExt(sFileName, "TMP")
    RichTextbox1.SaveFile stmp, RtfLoadSaveFormatText
    ePASS01 = stmp
    txttodoc stmp, sFileName
    DeleteFile stmp  'Kill stmp
  End If
End Sub

Private Sub CmdSalvarComoPdf_Click()
  Dim sFILTER As String
  Dim sFileName As String
  Dim stmp As String
  stmp = TrocaExt(cARQRTF, "PDF")
  sFILTER = "Arquivos de Textos (*.PDF)" & vbNullChar & "*.PDF" & vbNullChar & "Todos Arquivo" & vbNullChar & "*.*"
  sFileName = FileSave(Me, sFILTER, 1, "PDF", stmp, stmp, "Salvar PDF Como")
  If Not Len(sFileName) = 0 Then
    stmp = TrocaExt(sFileName, "TMP")
    RichTextbox1.SaveFile stmp, RtfLoadSaveFormatText
    ePASS01 = stmp
    formConvertToPDF.Show vbModal, Me
    DeleteFile stmp  'Kill stmp
  End If
End Sub

Private Sub Form_Load()
  larqtxt = False
  
  If IsExtensao(cARQRTF, "TXT") Or IsExtensao(cARQRTF, "ZPL ") Or IsExtensao(cARQRTF, "MAN") _
     Or IsExtensao(cARQRTF, "INI") Then
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

Private Sub retornar_Click()
Unload Me
End Sub

Private Sub SalvarComoHTML_Click()
 Dim sFILTER As String
  Dim sFileName As String
  Dim stmp As String
  stmp = TrocaExt(cARQRTF, "HTML")
  sFILTER = "Arquivos de Textos (*.HTML)" & vbNullChar & "*.HTML" & vbNullChar & "Todos Arquivo" & vbNullChar & "*.*"
  sFileName = FileSave(Me, sFILTER, 1, "HTML", stmp, stmp, "Salvar HTML Como")
  If Not Len(sFileName) = 0 Then
    If FileConnExist(sFileName) Then
      If MDG("Arquivo ja Existe Sobrepor") Then
        DeleteFile sFileName  'Kill sFILENAME
        Exit Sub
      End If
    End If
    stmp = TrocaExt(sFileName, "TMP")
    RichTextbox1.SaveFile stmp, RtfLoadSaveFormatText
    ePASS01 = stmp
    txttohtml stmp, sFileName
    DeleteFile stmp  'Kill stmp
  End If
End Sub

Private Sub SalvarComoRTF_Click()
Dim sFILTER As String
  Dim sFileName As String
  sFILTER = "Arquivos de Textos (*.RTF)" & vbNullChar & "*.RTF" & vbNullChar & "Todos Arquivo" & vbNullChar & "*.*"
  sFileName = FileSave(Me, sFILTER, 1, "RTF", cARQRTF, cARQRTF, "Salvar RTF Como")
  If Not Len(sFileName) = 0 Then
    RichTextbox1.SaveFile sFileName, RtfLoadSaveFormatRTF
    RichTextbox1.SetFocus
  End If
End Sub

Private Sub SalvarComoTxt_Click()
Dim sFILTER As String
  Dim sFileName As String
  sFILTER = "Arquivos de Textos (*.TXT)" & vbNullChar & "*.TXT" & vbNullChar & "Todos Arquivo" & vbNullChar & "*.*"
  sFileName = FileSave(Me, sFILTER, 1, "TXT", cARQRTF, cARQRTF, "Salvar TXT Como")
  If Not Len(sFileName) = 0 Then
    RichTextbox1.SaveFile sFileName, RtfLoadSaveFormatText
    RichTextbox1.SetFocus
  End If
End Sub
