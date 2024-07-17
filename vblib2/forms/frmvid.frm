VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{F22668DE-E08D-467B-8E41-13900013BD5F}#2.7#0"; "VBextra2.OCX"
Begin VB.Form FrmVid 
   Caption         =   "Media View"
   ClientHeight    =   4095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5655
   Icon            =   "frmvid.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   5655
   StartUpPosition =   3  'Windows Default
   Begin vbExtra.SizeGrip SizeGrip1 
      Height          =   285
      Left            =   5370
      Top             =   3810
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
   End
   Begin XPControls.XPButton encerrar 
      Height          =   495
      Left            =   4440
      TabIndex        =   0
      Top             =   240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      Picture         =   "frmvid.frx":038A
      Caption         =   "&Retornar"
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
   Begin WMPLibCtl.WindowsMediaPlayer Player 
      Height          =   3855
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4215
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   7435
      _cy             =   6800
   End
End
Attribute VB_Name = "FrmVid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Const nFORMID = 1071
Const cFORMID = "Visualizar de Videos"

Private Sub Encerrar_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  CenterFormToScreen Me
  If FileExist(cARQRTF, True) Then
    Player.URL = cARQRTF

    '        PlayerVideo.FileName = cARQRTF
  End If
End Sub

