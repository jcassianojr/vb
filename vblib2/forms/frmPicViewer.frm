VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{451B73A5-1563-45D5-A6AC-7B2B7D30B778}#1.1#0"; "BSPrin10.ocx"
Begin VB.Form frmPicViewer 
   Caption         =   "Picture Viewer"
   ClientHeight    =   7920
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14715
   Icon            =   "frmPicViewer.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   7920
   ScaleWidth      =   14715
   StartUpPosition =   1  'CenterOwner
   Begin BSPrinter.PrintPreview PrintPreview1 
      Left            =   12240
      Top             =   5280
      _ExtentX        =   1191
      _ExtentY        =   1191
   End
   Begin VB.PictureBox picPrint 
      Height          =   3465
      Left            =   6840
      ScaleHeight     =   3405
      ScaleWidth      =   4905
      TabIndex        =   27
      Top             =   3570
      Visible         =   0   'False
      Width           =   4965
      Begin XPControls.XPButton cmdCloseDialog 
         Height          =   315
         Left            =   3750
         TabIndex        =   14
         Top             =   3060
         Width           =   1095
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "Close"
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
      Begin VB.Frame fraPrint 
         Caption         =   "Print Dialog"
         Height          =   2985
         Left            =   120
         TabIndex        =   28
         Top             =   0
         Width           =   4755
         Begin VB.PictureBox picFix 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   120
            ScaleHeight     =   495
            ScaleWidth      =   1065
            TabIndex        =   33
            Top             =   2370
            Width           =   1065
            Begin VB.OptionButton optOrientation 
               Caption         =   "Landscape"
               Enabled         =   0   'False
               Height          =   315
               Index           =   1
               Left            =   0
               TabIndex        =   12
               Top             =   240
               Width           =   1215
            End
            Begin VB.OptionButton optOrientation 
               Caption         =   "Portrait"
               Enabled         =   0   'False
               Height          =   315
               Index           =   0
               Left            =   0
               TabIndex        =   11
               Top             =   -30
               Value           =   -1  'True
               Width           =   1215
            End
         End
         Begin VB.ComboBox cmbPrinters 
            Height          =   315
            Left            =   120
            TabIndex        =   8
            Top             =   450
            Width           =   3585
         End
         Begin VB.ComboBox cmbPapers 
            Height          =   315
            Left            =   120
            TabIndex        =   9
            Top             =   1080
            Width           =   4545
         End
         Begin VB.ComboBox cmbResolutions 
            Height          =   315
            Left            =   120
            TabIndex        =   10
            Top             =   1710
            Width           =   1575
         End
         Begin VB.Label lblFix 
            AutoSize        =   -1  'True
            Caption         =   "Printer Names :"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   1080
         End
         Begin VB.Label lblFix 
            AutoSize        =   -1  'True
            Caption         =   "Paper IDs, Names, Sizes (x, y) in mm :"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   31
            Top             =   870
            Width           =   2670
         End
         Begin VB.Label lblFix 
            AutoSize        =   -1  'True
            Caption         =   "Resolutions (x, y) in DPI :"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   30
            Top             =   1500
            Width           =   1770
         End
         Begin VB.Label lblFix 
            AutoSize        =   -1  'True
            Caption         =   "Paper Orientation :"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   29
            Top             =   2130
            Width           =   1320
         End
      End
      Begin XPControls.XPButton cmdPrintDialog 
         Height          =   315
         Left            =   2610
         TabIndex        =   13
         Top             =   3060
         Width           =   1095
         _ExtentX        =   0
         _ExtentY        =   0
         Enabled         =   0   'False
         Caption         =   "Print"
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
   Begin XPControls.XPButton cmdPrint 
      Height          =   345
      Left            =   11010
      TabIndex        =   7
      Top             =   7065
      Width           =   765
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "Print..."
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
   Begin XPControls.XPButton cmdStopSlideShow 
      Height          =   345
      Left            =   7410
      TabIndex        =   1
      Top             =   7065
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "Stop Slide Show"
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
   Begin VB.Timer tmrSlideShow 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7800
      Top             =   7410
   End
   Begin VB.HScrollBar HScrollViewer 
      Height          =   240
      LargeChange     =   500
      Left            =   2610
      Max             =   5000
      SmallChange     =   50
      TabIndex        =   26
      Top             =   6750
      Width           =   9180
   End
   Begin VB.VScrollBar VScrollViewer 
      Height          =   6645
      LargeChange     =   500
      Left            =   11550
      Max             =   5000
      SmallChange     =   50
      TabIndex        =   25
      Top             =   120
      Width           =   240
   End
   Begin VB.PictureBox picInfo 
      Height          =   345
      Left            =   2610
      ScaleHeight     =   285
      ScaleWidth      =   4665
      TabIndex        =   19
      Top             =   7050
      Width           =   4725
      Begin VB.Label lblCoordinate 
         Alignment       =   1  'Right Justify
         Caption         =   "lblCoordinate"
         Height          =   225
         Left            =   3690
         TabIndex        =   22
         Top             =   60
         Width           =   900
      End
      Begin VB.Label lblDimension 
         Caption         =   "lblDimension"
         Height          =   225
         Left            =   1590
         TabIndex        =   21
         Top             =   60
         Width           =   1980
      End
      Begin VB.Label lblSize 
         Caption         =   "lblSize"
         Height          =   225
         Left            =   30
         TabIndex        =   20
         Top             =   60
         Width           =   1380
      End
   End
   Begin VB.TextBox txtAddress 
      Height          =   315
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   7470
      Width           =   11655
   End
   Begin VB.PictureBox picFolderFile 
      Height          =   7275
      Left            =   120
      ScaleHeight     =   7215
      ScaleWidth      =   2355
      TabIndex        =   15
      Top             =   120
      Width           =   2415
      Begin VB.FileListBox lstViewer 
         Height          =   6330
         Left            =   90
         TabIndex        =   16
         Top             =   555
         Width           =   2220
      End
      Begin XPControls.XPButton cmdBrowseFolder 
         Height          =   410
         Left            =   120
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   120
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   714
         Picture         =   "frmPicViewer.frx":5872
         Caption         =   "Buscar Pasta"
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
      Begin VB.Label lblTotalFile 
         Caption         =   "lblTotalFile"
         Height          =   225
         Left            =   90
         TabIndex        =   17
         Top             =   6930
         Width           =   2100
      End
   End
   Begin XPControls.XPButton cmdLast 
      Height          =   345
      Left            =   9870
      TabIndex        =   5
      Top             =   7065
      Width           =   345
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   ">>"
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
   Begin XPControls.XPButton cmdNext 
      Height          =   345
      Left            =   9510
      TabIndex        =   4
      Top             =   7065
      Width           =   345
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   ">"
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
   Begin XPControls.XPButton cmdPrevious 
      Height          =   345
      Left            =   9150
      TabIndex        =   3
      Top             =   7065
      Width           =   345
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "<"
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
   Begin XPControls.XPButton cmdFirst 
      Height          =   345
      Left            =   8790
      TabIndex        =   2
      Top             =   7065
      Width           =   345
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "<<"
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
   Begin XPControls.XPButton cmdMenu 
      Height          =   345
      Left            =   10245
      TabIndex        =   6
      Top             =   7065
      Width           =   735
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "Menu..."
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
   Begin VB.PictureBox picViewer 
      Height          =   6630
      Left            =   2610
      ScaleHeight     =   6570
      ScaleWidth      =   8880
      TabIndex        =   0
      Top             =   120
      Width           =   8940
      Begin VB.PictureBox picProperties 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   0
         ScaleHeight     =   945
         ScaleWidth      =   4470
         TabIndex        =   23
         Top             =   0
         Visible         =   0   'False
         Width           =   4500
         Begin VB.Timer tmrProperties 
            Enabled         =   0   'False
            Interval        =   50
            Left            =   0
            Top             =   0
         End
         Begin VB.TextBox txtProperties 
            BackColor       =   &H00C00000&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            ForeColor       =   &H00FFFFFF&
            Height          =   855
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   24
            Top             =   600
            Width           =   4515
         End
      End
      Begin VB.Image imgViewerUp 
         Height          =   1095
         Left            =   1140
         MouseIcon       =   "frmPicViewer.frx":5E0C
         Top             =   1140
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Image imgViewerDown 
         Height          =   1095
         Left            =   0
         MouseIcon       =   "frmPicViewer.frx":66D6
         Top             =   1140
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Image imgViewer 
         Height          =   1095
         Left            =   0
         Top             =   0
         Width           =   1095
      End
   End
   Begin XPControls.XPButton CmdFechar 
      Height          =   555
      Left            =   12000
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   720
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   979
      Picture         =   "frmPicViewer.frx":6FA0
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
   Begin VB.Menu mnuPopUp1 
      Caption         =   "PopUp1"
      Visible         =   0   'False
      Begin VB.Menu mnuPopUp1Zoom 
         Caption         =   "Zoom"
         Begin VB.Menu mnuPopUp1ZoomIn 
            Caption         =   "Zoom In"
         End
         Begin VB.Menu mnuPopUp1ZoomOut 
            Caption         =   "Zoom Out"
         End
         Begin VB.Menu mnuPopUp1ZoomActualSize 
            Caption         =   "Actual Size"
         End
         Begin VB.Menu mnuPopUp1Zoom43 
            Caption         =   "4 : 3"
         End
         Begin VB.Menu mnuPopUp1Zoom169 
            Caption         =   "16 : 9"
         End
         Begin VB.Menu mnuPopUp1ZoomFitImage 
            Caption         =   "Fit Image"
         End
         Begin VB.Menu mnuPopUp1ZoomFitWidth 
            Caption         =   "Fit Width"
         End
         Begin VB.Menu mnuPopUp1ZoomFitHeight 
            Caption         =   "Fit Height"
         End
         Begin VB.Menu mnuPopUp1ZoomTo 
            Caption         =   "Zoom To"
            Begin VB.Menu mnuPopUp1Zoom10 
               Caption         =   "10 %"
            End
            Begin VB.Menu mnuPopUp1Zoom25 
               Caption         =   "25 %"
            End
            Begin VB.Menu mnuPopUp1Zoom33 
               Caption         =   "33 %"
            End
            Begin VB.Menu mnuPopUp1Zoom50 
               Caption         =   "50 %"
            End
            Begin VB.Menu mnuPopUp1Zoom67 
               Caption         =   "67 %"
            End
            Begin VB.Menu mnuPopUp1Zoom75 
               Caption         =   "75 %"
            End
            Begin VB.Menu mnuPopUp1Zoom100 
               Caption         =   "100 %"
            End
            Begin VB.Menu mnuPopUp1Zoom150 
               Caption         =   "150 %"
            End
            Begin VB.Menu mnuPopUp1Zoom200 
               Caption         =   "200 %"
            End
            Begin VB.Menu mnuPopUp1Zoom300 
               Caption         =   "300 %"
            End
            Begin VB.Menu mnuPopUp1Zoom400 
               Caption         =   "400 %"
            End
            Begin VB.Menu mnuPopUp1Zoom500 
               Caption         =   "500 %"
            End
            Begin VB.Menu mnuPopUp1Zoom1000 
               Caption         =   "1000 %"
            End
         End
      End
      Begin VB.Menu mnuPopUp1First 
         Caption         =   "First"
      End
      Begin VB.Menu mnuPopUp1Previous 
         Caption         =   "Previous"
      End
      Begin VB.Menu mnuPopUp1Next 
         Caption         =   "Next"
      End
      Begin VB.Menu mnuPopUp1Last 
         Caption         =   "Last"
      End
      Begin VB.Menu mnuPopUp1SlideShow 
         Caption         =   "Slide Show..."
      End
      Begin VB.Menu PopUp1Garis1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopUp1CopyToClipboard 
         Caption         =   "Copy To Clipboard"
      End
      Begin VB.Menu PopUp1Garis2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopUp1Properties 
         Caption         =   "Properties..."
      End
   End
End
Attribute VB_Name = "frmPicViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const nFORMID = 1110
Const cFORMID = "frmpicviewer"
Option Explicit
'------------------------------------------------------------------------------
'Title:Picture Viewer
'Coder:Hidayat Suriapermana
'Date:August 4, 2011
'Last Update:November 10, 2018
'Email:hsuriapermana@yahoo.com
'English Comment:
'One more example of the image exploration with Visual Basic. Here we can
'see images along with detailed information about the size and dimensions
'also coordinates when the mouse pointer over the image. In addition, we can
'also zoom, navigation, slide show, copy to clipboard and shift the image being displayed.
'Information about the number of image files in a folder and its location is also shown.
'Print feature is also available.
'While a filtered image files are bmp, gif, jpg, ico, wmf, cur and png.
'Indonesian comment:
'Satu lagi contoh eksplorasi image dengan Visual Basic. Disini kita bisa
'melihat image beserta informasi detail mengenai size dan dimensinya
'juga koordinat pointer mouse saat melewati image. Selain itu, kita dapat
'pula melakukan zoom, navigasi, slide show, copy ke clipboard dan menggeser image yang sedang ditampilkan.
'Informasi mengenai jumlah file image dan lokasinya dalam sebuah folder juga ditampilkan.
'Fasilitas cetak juga tersedia.
'Sedangkan file image yang difilter adalah bmp, gif, jpg, ico, wmf, cur dan png.
'------------------------------------------------------------------------------

Dim KAX As Single
Dim KAY As Single
Dim VarFileGambar As String
Attribute VarFileGambar.VB_VarUserMemId = 1073938434
Dim FileTerpilih As String
Dim VarImage1Height As String
Dim VarImage1Width As String
Dim VarImage1HeightZoom As String
Dim VarImage1WidthZoom As String

'Untuk form resizer
Private objFormResizer As clsResizer
Attribute objFormResizer.VB_VarUserMemId = 1073938440
Dim MeWidthMin As Single
Attribute MeWidthMin.VB_VarUserMemId = 1073938441
Dim MeHeightMin As Single
Attribute MeHeightMin.VB_VarUserMemId = 1073938442

'Untuk slide show
Dim iFile As Long
Attribute iFile.VB_VarUserMemId = 1073938443

'Untuk menangani png file
Dim Token As Long
Attribute Token.VB_VarUserMemId = 1073938444

'Untuk fasilitas print
Private Const DC_PAPERS = 2
Private Const DC_PAPERSIZE = 3
Private Const DC_PAPERNAMES = 16
Private Const DC_ENUMRESOLUTIONS = 13
Private Const DC_ORIENTATION = 17
Private Type POINT
  x As Long
  y As Long
End Type

Private Declare Function DeviceCapabilities Lib "winspool.drv" Alias "DeviceCapabilitiesW" (ByVal lpDeviceName As Long, ByVal lpPort As Long, ByVal iIndex As Long, ByVal lpOutput As Long, ByVal lpDevMode As Long) As Long
Private intPaperIds() As Integer
Attribute intPaperIds.VB_VarUserMemId = 1073938445
Private ptPaperSizes() As POINT
Attribute ptPaperSizes.VB_VarUserMemId = 1073938446
Private lngResolutions() As Long
Attribute lngResolutions.VB_VarUserMemId = 1073938447
Private Pr As Printer
Attribute Pr.VB_VarUserMemId = 1073938448
Private Pic As StdPicture
Attribute Pic.VB_VarUserMemId = 1073938449

Sub LokasiFile()
  If Right(lstViewer.Path, 1) = "\" Then
    FileTerpilih = lstViewer.Path & lstViewer.FileName
  Else
    FileTerpilih = lstViewer.Path & "\" & lstViewer.FileName
  End If
End Sub

Sub TampilanBrowse()
  Me.WindowState = 0
End Sub

Sub TampilanView()
  Me.WindowState = 2
End Sub

Sub GambarKeTengah()
  If imgViewer.Width <= picViewer.Width Then
    imgViewer.Left = (picViewer.Width - imgViewer.Width) / 2
    HScrollViewer.Visible = False
  Else
    imgViewer.Left = 0
    HScrollViewer.Visible = True
    HScrollViewer.Value = 0
  End If
  If imgViewer.Height <= picViewer.Height Then
    imgViewer.Top = (picViewer.Height - imgViewer.Height) / 2
    VScrollViewer.Visible = False
  Else
    imgViewer.Top = 0
    VScrollViewer.Visible = True
    VScrollViewer.Value = 0
  End If
End Sub

Sub AturScrollbar()
'Horizontal
  If imgViewer.Width > picViewer.Width Then
    If imgViewer.Width - picViewer.Width <= 32767 Then
      HScrollViewer.Max = (imgViewer.Width - picViewer.Width)
    Else
      HScrollViewer.Max = 32767
    End If
    If imgViewer.Left < -(imgViewer.Width - picViewer.Width) Then
      HScrollViewer.Value = (imgViewer.Width - picViewer.Width)
    ElseIf imgViewer.Left > 0 Then
      HScrollViewer.Value = 0
    Else
      HScrollViewer.Value = -imgViewer.Left
    End If
  End If
  'Vertical
  If imgViewer.Height > picViewer.Height Then
    If imgViewer.Height - picViewer.Height <= 32767 Then
      VScrollViewer.Max = (imgViewer.Height - picViewer.Height)
    Else
      VScrollViewer.Max = 32767
    End If
    If imgViewer.Top < -(imgViewer.Height - picViewer.Height) Then
      VScrollViewer.Value = (imgViewer.Height - picViewer.Height)
    ElseIf imgViewer.Top > 0 Then
      VScrollViewer.Value = 0
    Else
      VScrollViewer.Value = -imgViewer.Top
    End If
  End If
End Sub

Private Sub CmdFEchar_Click()
  lRETU = False
  eRETU01 = ""
  If FileExists(txtAddress.Text) Then
    lRETU = True
    eRETU01 = txtAddress.Text
  End If
  Unload Me
End Sub

Private Sub cmdPrintDialog_Click()
  PrintPreview1.ShowPreview
End Sub



Private Sub Form_Load()
  lRETU = False
  'eRETU = ""
  Me.Caption = cFORMID
  HelpContextID = nFORMID

  'Untuk form resizer
  Set objFormResizer = New clsResizer
  With objFormResizer
    Set .Container = Me
    .SetAnchors "picFolderFile", avTop + avBottom + avLeft
    .SetAnchors "cmdBrowseFolder", avTop + avLeft
    .SetAnchors "lstViewer", avTop + avBottom + avLeft
    .SetAnchors "lblTotalFile", avBottom + avLeft
    .SetAnchors "txtAddress", avBottom + avLeft + avRight
    .SetAnchors "picViewer", avTop + avBottom + avLeft + avRight
    .SetAnchors "HScrollViewer", avBottom + avLeft + avRight
    .SetAnchors "VScrollViewer", avTop + avBottom + avRight
    .SetAnchors "picInfo", avBottom + avLeft
    .SetAnchors "cmdStopSlideShow", avBottom + avRight
    .SetAnchors "cmdFirst", avBottom + avRight
    .SetAnchors "cmdPrevious", avBottom + avRight
    .SetAnchors "cmdNext", avBottom + avRight
    .SetAnchors "cmdLast", avBottom + avRight
    .SetAnchors "cmdMenu", avBottom + avRight
    .SetAnchors "cmdPrint", avBottom + avRight
    .SetAnchors "picPrint", avBottom + avRight
    .SetAnchors "CmdFechar", avBottom + avRight
  End With
  MeWidthMin = 12000
  MeHeightMin = 8400

  lblSize.Caption = ""
  lblDimension.Caption = ""
  lblCoordinate.Caption = ""
  lstViewer.FileName = "*.bmp;*.gif;*.jpg;*.ico;*.wmf;*.cur;*.png"
  If lstViewer.ListCount <> 0 Then
    lstViewer.ListIndex = 0
  End If
  VarImage1HeightZoom = imgViewer.Height
  VarImage1WidthZoom = imgViewer.Width
  lblTotalFile.Caption = "Total : " & lstViewer.ListCount & " file(s)"

  PrintPreview1.AuxiliaryButtonVisible = PrintPreview1.PrinterExists("Microsoft Print to PDF")
  PrintPreview1.AuxiliaryButtonToolTipText = "Salvar como PDF"
End Sub
Public Sub PrintPreview1_AuxiliaryButtonClick(UpdateReport As Boolean)
  PrintPreview1.ShowSaveToFile "Microsoft Print to PDF", "*.pdf"
  UpdateReport = False  ' we don't need to update the report in the Print preview window after this action (the default value of UpdateReport parameter is True)
End Sub
Private Sub Form_Activate()
  lstViewer.SetFocus                           'So that the button can be used as previous and next (Agar tombol bisa digunakan sebagai previous dan next)
End Sub

Private Sub Form_Resize()
  If Me.Width >= MeWidthMin Then
    objFormResizer.Resize
  Else
    If Me.WindowState = 0 Then
      Me.Width = MeWidthMin
    End If
  End If
  If Me.Height >= MeHeightMin Then
    objFormResizer.Resize
  Else
    If Me.WindowState = 0 Then
      Me.Height = MeHeightMin
    End If
  End If
  GambarKeTengah
  picProperties.Width = picViewer.Width
  txtProperties.Width = picViewer.Width
  AturScrollbar
End Sub

Private Sub cmdBrowseFolder_Click()
Dim PICBROWSE
  PICBROWSE = BrowseFolders(Me.hWnd, "Choose a location picture", BrowseForFolders, CSIDL_RECENT)
  If PICBROWSE <> "" Then
    lstViewer.FileName = PathTujuan
    If lstViewer.ListCount <> 0 Then
      lstViewer.ListIndex = 0
    Else
      imgViewer.Picture = Nothing
      lblTotalFile.Caption = "Total : " & lstViewer.ListCount & " file(s)"
      lblDimension.Caption = ""
      lblSize.Caption = ""
      HScrollViewer.Visible = False
      VScrollViewer.Visible = False
    End If
    lstViewer.SetFocus                       'So that the button can be used as previous and next (Agar tombol bisa digunakan sebagai previous dan next)
  End If
End Sub

Private Sub imgViewer_DblClick()
  If Me.WindowState = 2 Then
    TampilanBrowse
  ElseIf Me.WindowState <> 2 Then
    TampilanView
  End If
End Sub

Private Sub imgViewer_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = 1 Then
    'KAX = Initial coordinates of the X axis (Koordinat Awal sb. X)
    'KAY = Initial coordinates of the Y axis (Koordinat Awal sb. Y)
    KAX = x
    KAY = y
    If imgViewer.Width > picViewer.Width Or imgViewer.Height > picViewer.Height Then
      'If Dir(App.Path + "\Mouse.ico") <> vbNullString Then
      '    MouseIcon = LoadPicture(App.Path + "\Mouse.ico")
      '    MousePointer = 99
      'Else
      '    MousePointer = 5
      'End If
      MouseIcon = imgViewerDown.MouseIcon
      MousePointer = 99
    End If
  End If
End Sub

Private Sub imgViewer_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  lstViewer.SetFocus                           'So that the button can be used as previous and next (Agar tombol bisa digunakan sebagai previous dan next)
  If Button = 1 Then
    AturScrollbar
    'X = VB by default stating the coordinates of the X axis (VB secara default menyatakan koordinat sb. X saat itu)
    'Y = VB by default stating the coordinates of the Y axis (VB secara default menyatakan koordinat sb. Y saat itu)
    lblCoordinate.Caption = "(" & x & "," & y & ")"
    'Horizontal
    If imgViewer.Width > picViewer.Width Then
      If imgViewer.Left >= 0 And x - KAX > 0 Then
        imgViewer.Left = 0
        Exit Sub
      ElseIf imgViewer.Left <= picViewer.Width - imgViewer.Width And x - KAX < 0 Then
        imgViewer.Left = picViewer.Width - imgViewer.Width
        Exit Sub
      Else
        imgViewer.Left = imgViewer.Left + x - KAX
      End If
    End If
    'Vertical
    If imgViewer.Height > picViewer.Height Or imgViewer.Height > picViewer.Height Then
      If imgViewer.Top >= 0 And y - KAY > 0 Then
        imgViewer.Top = 0
        Exit Sub
      ElseIf imgViewer.Top <= picViewer.Height - imgViewer.Height And y - KAY < 0 Then
        imgViewer.Top = picViewer.Height - imgViewer.Height
        Exit Sub
      Else
        imgViewer.Top = imgViewer.Top + y - KAY
      End If
    End If
  Else
    If imgViewer.Width > picViewer.Width Or imgViewer.Height > picViewer.Height Then
      MouseIcon = imgViewerUp.MouseIcon
      MousePointer = 99
    Else
      MousePointer = 0
    End If
  End If
End Sub

Private Sub imgViewer_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  If imgViewer.Width > picViewer.Width Or imgViewer.Height > picViewer.Height Then
    MouseIcon = imgViewerUp.MouseIcon
    MousePointer = 99
  End If
  If Button = 2 Then
    If Me.WindowState <> 2 Then
      PopupMenu mnuPopUp1
    ElseIf Me.WindowState = 2 Then
      PopupMenu mnuPopUp1
    End If
  End If
  lblCoordinate.Caption = ""
End Sub

Private Sub lstViewer_Click()
Dim FILESIZE As Integer
  On Error GoTo bawah
  LokasiFile
  VarFileGambar = True
  If UCase(Right(FileTerpilih, 4)) = ".PNG" Then
    Token = InitGDIPlus
    imgViewer.Picture = LoadPictureGDIPlus(FileTerpilih, , , vbWhite)
    FreeGDIPlus Token
  Else
    imgViewer.Picture = LoadPicture(FileTerpilih)
  End If
  imgViewer.Stretch = False
  GambarKeTengah
  lblDimension.Caption = "Dimension : " & imgViewer.Width & " x " & imgViewer.Height
  VarImage1Height = imgViewer.Height
  VarImage1Width = imgViewer.Width
  VarImage1HeightZoom = imgViewer.Height
  VarImage1WidthZoom = imgViewer.Width
  AturScrollbar
  txtAddress.Text = FileTerpilih
  lblTotalFile.Caption = "Total : " & lstViewer.ListCount & " file(s)"
  lblCoordinate.Caption = ""
  FILESIZE = Int(FileLen(FileTerpilih) / 1024 + 1)
  If FileLen(FileTerpilih) >= 1000000 Then
    lblSize.Caption = "Size : " & Round(FILESIZE / 1000, 3) & " MB"
  Else
    lblSize.Caption = "Size : " & FILESIZE & " KB"
  End If
  If imgViewer.Width > picViewer.Width Or imgViewer.Height > picViewer.Height Then
    MouseIcon = imgViewerUp.MouseIcon
    MousePointer = 99
  Else
    MousePointer = 0
  End If
  Exit Sub

bawah:
  If Err.Number = 53 Then
    lblDimension.Caption = ""
    lblSize.Caption = ""
    imgViewer.Picture = Nothing
    HScrollViewer.Visible = False
    VScrollViewer.Visible = False
    MsgBox "The file has been deleted, moved or renamed", vbCritical, ""  'File telah dihapus, dipindahkan, atau diganti nama
  ElseIf Err.Number = 76 Then
    lblDimension.Caption = ""
    lblSize.Caption = ""
    imgViewer.Picture = Nothing
    HScrollViewer.Visible = False
    VScrollViewer.Visible = False
    MsgBox "The folder has been deleted, moved or renamed", vbCritical, ""  'Folder telah dihapus, dipindahkan, atau diganti nama
  ElseIf Err.Number = 0 Then
    'Leave it alone (Biarkan saja)
  Else
    imgViewer.Picture = Nothing
    MsgBox "" & Err.Description, vbExclamation, ""
  End If
End Sub

Private Sub lstViewer_DblClick()
  If VarFileGambar = True Then
    imgViewer_DblClick
  End If
End Sub

Private Sub HScrollViewer_Change()
  imgViewer.Left = -HScrollViewer.Value
End Sub
Private Sub VScrollViewer_Change()
  imgViewer.Top = -VScrollViewer.Value
End Sub

Private Sub HScrollViewer_Scroll()
  imgViewer.Left = -HScrollViewer.Value
End Sub

Private Sub VScrollViewer_Scroll()
  imgViewer.Top = -VScrollViewer.Value
End Sub

Private Sub txtProperties_Click()
  picProperties.Visible = False
End Sub

Private Sub picProperties_Click()
  picProperties.Visible = False
End Sub

Private Sub cmdFirst_Click()
  mnuPopUp1First_Click
End Sub

Private Sub cmdPrevious_Click()
  mnuPopUp1Previous_Click
End Sub

Private Sub cmdNext_Click()
  mnuPopUp1Next_Click
End Sub

Private Sub cmdLast_Click()
  mnuPopUp1Last_Click
End Sub

Private Sub cmdMenu_Click()
  PopupMenu mnuPopUp1
End Sub

Private Sub mnuPopUp1ZoomIn_Click()
  VarImage1HeightZoom = imgViewer.Height * 1.25
  VarImage1WidthZoom = imgViewer.Width * 1.25
  imgViewer.Height = VarImage1HeightZoom * 1.25
  imgViewer.Width = VarImage1WidthZoom * 1.25
  GambarKeTengah
  imgViewer.Stretch = True
  AturScrollbar
End Sub

Private Sub mnuPopUp1ZoomOut_Click()
  VarImage1HeightZoom = imgViewer.Height * 0.8
  VarImage1WidthZoom = imgViewer.Width * 0.8
  imgViewer.Height = VarImage1HeightZoom * 0.8
  imgViewer.Width = VarImage1WidthZoom * 0.8
  GambarKeTengah
  imgViewer.Stretch = True
  AturScrollbar
End Sub

Private Sub mnuPopUp1Zoom43_Click()
  imgViewer.Height = 7000
  imgViewer.Width = imgViewer.Height * 1.33
  GambarKeTengah
  imgViewer.Stretch = True
  AturScrollbar
End Sub

Private Sub mnuPopUp1Zoom169_Click()
  imgViewer.Height = 7000
  imgViewer.Width = imgViewer.Height * 1.78
  GambarKeTengah
  imgViewer.Stretch = True
  AturScrollbar
End Sub

Private Sub mnuPopUp1ZoomFitImage_Click()
  imgViewer.Height = picViewer.Height
  imgViewer.Width = picViewer.Width
  GambarKeTengah
  imgViewer.Stretch = True
  AturScrollbar
End Sub

Private Sub mnuPopUp1ZoomFitWidth_Click()
  imgViewer.Width = picViewer.Width
  imgViewer.Height = VarImage1Height * (picViewer.Width / VarImage1Width)
  GambarKeTengah
  imgViewer.Stretch = True
  AturScrollbar
End Sub

Private Sub mnuPopUp1ZoomFitHeight_Click()
  imgViewer.Height = picViewer.Height
  imgViewer.Width = VarImage1Width * (picViewer.Height / VarImage1Height)
  GambarKeTengah
  imgViewer.Stretch = True
  AturScrollbar
End Sub

Private Sub mnuPopUp1Zoom10_Click()
  imgViewer.Height = VarImage1Height * 0.1
  imgViewer.Width = VarImage1Width * 0.1
  GambarKeTengah
  imgViewer.Stretch = True
  AturScrollbar
End Sub

Private Sub mnuPopUp1Zoom25_Click()
  imgViewer.Height = VarImage1Height * 0.25
  imgViewer.Width = VarImage1Width * 0.25
  GambarKeTengah
  imgViewer.Stretch = True
  AturScrollbar
End Sub

Private Sub mnuPopUp1Zoom33_Click()
  imgViewer.Height = VarImage1Height * 0.33
  imgViewer.Width = VarImage1Width * 0.33
  GambarKeTengah
  imgViewer.Stretch = True
  AturScrollbar
End Sub

Private Sub mnuPopUp1Zoom50_Click()
  imgViewer.Height = VarImage1Height * 0.5
  imgViewer.Width = VarImage1Width * 0.5
  GambarKeTengah
  imgViewer.Stretch = True
  AturScrollbar
End Sub

Private Sub mnuPopUp1Zoom67_Click()
  imgViewer.Height = VarImage1Height * 0.67
  imgViewer.Width = VarImage1Width * 0.67
  GambarKeTengah
  imgViewer.Stretch = True
  AturScrollbar
End Sub

Private Sub mnuPopUp1Zoom75_Click()
  imgViewer.Height = VarImage1Height * 0.75
  imgViewer.Width = VarImage1Width * 0.75
  GambarKeTengah
  imgViewer.Stretch = True
  AturScrollbar
End Sub

Private Sub mnuPopUp1Zoom100_Click()
  imgViewer.Height = VarImage1Height
  imgViewer.Width = VarImage1Width
  GambarKeTengah
  imgViewer.Stretch = True
  AturScrollbar
End Sub

Private Sub mnuPopUp1Zoom150_Click()
  imgViewer.Height = VarImage1Height * 1.5
  imgViewer.Width = VarImage1Width * 1.5
  GambarKeTengah
  imgViewer.Stretch = True
  AturScrollbar
End Sub

Private Sub mnuPopUp1Zoom200_Click()
  imgViewer.Height = VarImage1Height * 2
  imgViewer.Width = VarImage1Width * 2
  GambarKeTengah
  imgViewer.Stretch = True
  AturScrollbar
End Sub

Private Sub mnuPopUp1Zoom300_Click()
  imgViewer.Height = VarImage1Height * 3
  imgViewer.Width = VarImage1Width * 3
  GambarKeTengah
  imgViewer.Stretch = True
  AturScrollbar
End Sub

Private Sub mnuPopUp1Zoom400_Click()
  imgViewer.Height = VarImage1Height * 4
  imgViewer.Width = VarImage1Width * 4
  GambarKeTengah
  imgViewer.Stretch = True
  AturScrollbar
End Sub

Private Sub mnuPopUp1Zoom500_Click()
  imgViewer.Height = VarImage1Height * 5
  imgViewer.Width = VarImage1Width * 5
  GambarKeTengah
  imgViewer.Stretch = True
  AturScrollbar
End Sub

Private Sub mnuPopUp1Zoom1000_Click()
  imgViewer.Height = VarImage1Height * 10
  imgViewer.Width = VarImage1Width * 10
  GambarKeTengah
  imgViewer.Stretch = True
  AturScrollbar
End Sub

Private Sub mnuPopUp1ZoomActualSize_Click()
  mnuPopUp1Zoom100_Click
  AturScrollbar
End Sub

Private Sub mnuPopUp1First_Click()
  If lstViewer.ListCount <> 0 Then
    If lstViewer.ListIndex <> 0 Then
      lstViewer.ListIndex = 0
    End If
  End If
End Sub

Private Sub mnuPopUp1Previous_Click()
  If lstViewer.ListCount <> 0 Then
    If lstViewer.ListIndex <> 0 Then
      lstViewer.ListIndex = lstViewer.ListIndex - 1
    End If
  End If
End Sub

Private Sub mnuPopUp1Next_Click()
  If lstViewer.ListCount <> 0 Then
    If lstViewer.ListIndex < lstViewer.ListCount - 1 Then
      lstViewer.ListIndex = lstViewer.ListIndex + 1
    End If
  End If
End Sub

Private Sub mnuPopUp1Last_Click()
  If lstViewer.ListCount <> 0 Then
    If lstViewer.ListIndex < lstViewer.ListCount - 1 Then
      lstViewer.ListIndex = lstViewer.ListCount - 1
    End If
  End If
End Sub

Private Sub mnuPopUp1SlideShow_Click()
  iFile = 0
  cmdStopSlideShow.Visible = True
  tmrSlideShow.Enabled = True
End Sub

Private Sub tmrSlideShow_Timer()
  If lstViewer.ListCount > 0 Then
    If iFile < lstViewer.ListCount Then
      lstViewer.ListIndex = iFile
      iFile = iFile + 1
    Else
      tmrSlideShow.Enabled = False
      cmdStopSlideShow.Visible = False
    End If
  End If
End Sub

Private Sub cmdStopSlideShow_Click()
  tmrSlideShow.Enabled = False
  cmdStopSlideShow.Visible = False
End Sub

Private Sub mnuPopUp1Properties_Click()
  txtProperties.Top = 600
  tmrProperties.Enabled = True

  txtProperties.Text = "File : " & FileTerpilih + vbCrLf _
                     + lblDimension.Caption + vbCrLf _
                     + lblSize.Caption
End Sub

Private Sub tmrProperties_Timer()
  picProperties.Left = 0
  picProperties.Top = 0
  picProperties.Visible = True
  If txtProperties.Top > (picProperties.ScaleHeight - txtProperties.Height) Then
    txtProperties.Top = txtProperties.Top - 10
    If txtProperties.Top <= 90 Then
      picProperties.Visible = False
      tmrProperties.Enabled = False
    End If
  End If
End Sub

Private Sub mnuPopUp1CopyToClipboard_Click()
  On Error Resume Next
  Clipboard.Clear
  'If UCase(Right(FileTerpilih, 4)) = ".PNG" Then
  '    Token = InitGDIPlus
  '    Clipboard.SetData LoadPictureGDIPlus(FileTerpilih, , , vbWhite) 'Copy based on file name (Copy berdasarkan nama file)
  '    FreeGDIPlus Token
  'Else
  '    Clipboard.SetData LoadPicture(FileTerpilih) 'Copy based on file name (Copy berdasarkan nama file)
  'End If
  Clipboard.SetData imgViewer, 2               'Copy image direct from the image object (Copy gambar langsung dari object imgage)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  MousePointer = 0
End Sub

Private Sub cmdBrowseFolder_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  MousePointer = 0
End Sub

Private Sub lstViewer_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  MousePointer = 0
End Sub

Private Sub picPrint_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  MousePointer = 0
End Sub

Private Sub fraPrint_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  MousePointer = 0
End Sub

Private Sub picViewer_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  lstViewer.SetFocus                           'So that the button can be used as previous and next (Agar tombol bisa digunakan sebagai previous dan next)
End Sub

Private Sub picFolderFile_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  lstViewer.SetFocus                           'So that the button can be used as previous and next (Agar tombol bisa digunakan sebagai previous dan next)
End Sub

Private Sub picInfo_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  lstViewer.SetFocus                           'So that the button can be used as previous and next (Agar tombol bisa digunakan sebagai previous dan next)
End Sub

Private Sub cmdStopSlideShow_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  lstViewer.SetFocus                           'So that the button can be used as previous and next (Agar tombol bisa digunakan sebagai previous dan next)
End Sub

Private Sub cmdFirst_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  lstViewer.SetFocus                           'So that the button can be used as previous and next (Agar tombol bisa digunakan sebagai previous dan next)
End Sub

Private Sub cmdPrevious_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  lstViewer.SetFocus                           'So that the button can be used as previous and next (Agar tombol bisa digunakan sebagai previous dan next)
End Sub

Private Sub cmdNext_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  lstViewer.SetFocus                           'So that the button can be used as previous and next (Agar tombol bisa digunakan sebagai previous dan next)
End Sub

Private Sub cmdLast_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  lstViewer.SetFocus                           'So that the button can be used as previous and next (Agar tombol bisa digunakan sebagai previous dan next)
End Sub

Private Sub cmdMenu_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  lstViewer.SetFocus                           'So that the button can be used as previous and next (Agar tombol bisa digunakan sebagai previous dan next)
End Sub

Private Sub cmdPrint_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  lstViewer.SetFocus                           'So that the button can be used as previous and next (Agar tombol bisa digunakan sebagai previous dan next)
End Sub

Private Sub cmdCloseDialog_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  lstViewer.SetFocus                           'So that the button can be used as previous and next (Agar tombol bisa digunakan sebagai previous dan next)
End Sub

Private Sub txtAddress_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  lstViewer.SetFocus                           'So that the button can be used as previous and next (Agar tombol bisa digunakan sebagai previous dan next)
End Sub

Private Sub cmdPrint_Click()
  On Error GoTo ErrorHandle
  cmdPrint.Enabled = False
  cmdPrintDialog.Enabled = False
  picPrint.Visible = True
  cmbPrinters.Clear
  cmbPapers.Clear
  cmbResolutions.Clear
  optOrientation(0).Value = True
  optOrientation(0).Enabled = False
  optOrientation(1).Enabled = False
  For Each Pr In Printers
    cmbPrinters.AddItem Pr.DeviceName
  Next
  Exit Sub

ErrorHandle:
  MsgBox Err.Description, vbExclamation, ""
End Sub

Private Sub cmbPrinters_Click()
  On Error GoTo ErrorHandle

  Dim PaperCount As Long
  Dim strPaperNames As String
  Dim i As Long
  Dim strPaperName As String
  Dim ActualLength As Long
  Dim ResolutionCount As Long
  Dim LandscapeRotation As Long

  Set Pr = Printers(cmbPrinters.ListIndex)

  'Get paper IDs, names, sizes
  PaperCount = DeviceCapabilities(StrPtr(Pr.DeviceName), _
                                  StrPtr(Pr.Port), _
                                  DC_PAPERS, _
                                  0, _
                                  0)
  ReDim intPaperIds(PaperCount - 1)
  PaperCount = DeviceCapabilities(StrPtr(Pr.DeviceName), _
                                  StrPtr(Pr.Port), _
                                  DC_PAPERS, _
                                  VarPtr(intPaperIds(0)), _
                                  0)

  PaperCount = DeviceCapabilities(StrPtr(Pr.DeviceName), _
                                  StrPtr(Pr.Port), _
                                  DC_PAPERNAMES, _
                                  0, _
                                  0)
  strPaperNames = String$(PaperCount * 64, 0)
  PaperCount = DeviceCapabilities(StrPtr(Pr.DeviceName), _
                                  StrPtr(Pr.Port), _
                                  DC_PAPERNAMES, _
                                  StrPtr(strPaperNames), _
                                  0)

  PaperCount = DeviceCapabilities(StrPtr(Pr.DeviceName), _
                                  StrPtr(Pr.Port), _
                                  DC_PAPERSIZE, _
                                  0, _
                                  0)
  ReDim ptPaperSizes(PaperCount - 1)
  PaperCount = DeviceCapabilities(StrPtr(Pr.DeviceName), _
                                  StrPtr(Pr.Port), _
                                  DC_PAPERSIZE, _
                                  VarPtr(ptPaperSizes(0)), _
                                  0)
  cmbPapers.Clear
  For i = 0 To PaperCount - 1
    strPaperName = Mid$(strPaperNames, 64 * i + 1, 64)
    ActualLength = InStr(strPaperName, vbNullChar) - 1
    If ActualLength > 1 Then strPaperName = Left$(strPaperName, ActualLength)
    cmbPapers.AddItem CStr(intPaperIds(i)) _
                    & ": " & strPaperName _
                    & " (" & CStr(ptPaperSizes(i).x) / 10 _
                    & ", " & CStr(ptPaperSizes(i).y) / 10 & ")"
  Next
  cmbPapers.Enabled = True

  'Get resolutions
  ResolutionCount = DeviceCapabilities(StrPtr(Pr.DeviceName), _
                                       StrPtr(Pr.Port), _
                                       DC_ENUMRESOLUTIONS, _
                                       0, _
                                       0)
  ReDim lngResolutions(2 * ResolutionCount - 1)
  ResolutionCount = DeviceCapabilities(StrPtr(Pr.DeviceName), _
                                       StrPtr(Pr.Port), _
                                       DC_ENUMRESOLUTIONS, _
                                       VarPtr(lngResolutions(0)), _
                                       0)
  cmbResolutions.Clear
  For i = 0 To ResolutionCount - 1
    cmbResolutions.AddItem "(" & CStr(lngResolutions(2 * i)) _
                         & ", " & CStr(lngResolutions(2 * i + 1)) & ")"
  Next
  cmbResolutions.Enabled = True

  'Get orientations
  LandscapeRotation = DeviceCapabilities(StrPtr(Pr.DeviceName), _
                                         StrPtr(Pr.Port), _
                                         DC_ORIENTATION, _
                                         0, _
                                         0)
  optOrientation(0).Enabled = True
  If LandscapeRotation > 0 Then
    optOrientation(1).Enabled = True
  Else
    optOrientation(1).Enabled = False
    optOrientation(0).Value = True
  End If

  Exit Sub

ErrorHandle:
  cmbPapers.Clear
  cmbResolutions.Clear
  cmdPrintDialog.Enabled = False
  optOrientation(0).Enabled = False
  optOrientation(1).Enabled = False
End Sub

Private Sub cmbPapers_Click()
  If Len(cmbPapers.Text) > 0 And Len(cmbResolutions.Text) > 0 Then cmdPrintDialog.Enabled = True
End Sub

Private Sub cmbResolutions_Click()
  If Len(cmbPapers.Text) > 0 And Len(cmbResolutions.Text) > 0 Then cmdPrintDialog.Enabled = True
End Sub


Private Sub xcmdPrintDialog_Click()
  If UCase(Right(FileTerpilih, 4)) = ".PNG" Then
    Token = InitGDIPlus
    Set Pic = LoadPictureGDIPlus(FileTerpilih, , , vbWhite)
    FreeGDIPlus Token
  Else
    Set Pic = LoadPicture(FileTerpilih)
  End If

  'Print the StdPicture Pic centered on the selected rrinter (Pr)
  'with the selected paper (lstPapers) at the selected quality
  '(lngResolutions) within set margins
  Dim MarginsLR As Single
  Dim MarginsTB As Single
  Dim PrintableWidth As Single
  Dim PrintableHeight As Single
  Dim ScaleFactor As Double
  Dim ScaledWidth As Double
  Dim ScaledHeight As Double

  Set Printer = Pr
  With Printer
    'Set up paper
    .PaperSize = intPaperIds(cmbPapers.ListIndex)
    .PrintQuality = lngResolutions(2 * cmbResolutions.ListIndex)  'We can only set one
    'value DPI value, just
    'use X here
    If optOrientation(1).Value Then
      .Orientation = vbPRORLandscape
    Else
      .Orientation = vbPRORPortrait
    End If

    'Scale to paper, using 0.5" margins all around.Could also crop
    'the image here
    MarginsLR = .ScaleX(0.5, vbInches, .ScaleMode)
    MarginsTB = .ScaleY(0.5, vbInches, .ScaleMode)
    PrintableWidth = .Width - 2 * MarginsLR
    PrintableHeight = .Height - 2 * MarginsTB

    ScaleFactor = PrintableWidth / .ScaleX(Pic.Width, vbHimetric, .ScaleMode)
    If ScaleFactor * .ScaleY(Pic.Height, vbHimetric, .ScaleMode) > PrintableHeight Then
      ScaleFactor = PrintableHeight / .ScaleY(Pic.Height, vbHimetric, .ScaleMode)
    End If

    ScaledWidth = ScaleFactor * .ScaleX(Pic.Width, vbHimetric, .ScaleMode)
    ScaledHeight = ScaleFactor * .ScaleY(Pic.Height, vbHimetric, .ScaleMode)

    'Paint (print) the image, scaled.Could also do the actual cropping
    'here if any were desired by specifying additional arguments
    .PaintPicture Pic, _
                  (.Width - ScaledWidth) / 2, _
                  (.Height - ScaledHeight) / 2, _
                  ScaledWidth, _
                  ScaledHeight
    .NewPage
    '  .EndDoc
  End With
End Sub

Public Sub PrintPreview1_PrepareReport(Cancel As Boolean)
  xcmdPrintDialog_Click
End Sub

Private Sub cmdCloseDialog_Click()
  picPrint.Visible = False
  cmdPrint.Enabled = True
End Sub

Private Sub cmbPrinters_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub cmbPapers_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub cmbResolutions_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub


