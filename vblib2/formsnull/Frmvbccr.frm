VERSION 5.00
Object = "{7020C36F-09FC-41FE-B822-CDE6FBB321EB}#1.0#0"; "vbccr17.ocx"
Object = "{F22668DE-E08D-467B-8E41-13900013BD5F}#2.7#0"; "VBextra2.OCX"
Object = "{3E5D9624-07F7-4D22-90F8-1314327F7BAC}#1.0#0"; "VBFLXGRD15.OCX"
Begin VB.Form Frmvbccr 
   Caption         =   "Form1"
   ClientHeight    =   7335
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15675
   LinkTopic       =   "Form1"
   ScaleHeight     =   7335
   ScaleWidth      =   15675
   StartUpPosition =   3  'Windows Default
   Begin VBCCR17.SpinBox SpinBox1 
      Height          =   495
      Left            =   6840
      TabIndex        =   32
      Top             =   1440
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   873
   End
   Begin VBCCR17.TabStrip TabStrip1 
      Height          =   615
      Left            =   6600
      TabIndex        =   31
      Top             =   5640
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1085
      InitTabs        =   "Frmvbccr.frx":0000
   End
   Begin VBCCR17.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      Top             =   6960
      Width           =   15675
      _ExtentX        =   27649
      _ExtentY        =   661
      InitPanels      =   "Frmvbccr.frx":0054
   End
   Begin VBCCR17.SysInfo SysInfo1 
      Left            =   6600
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VBCCR17.TreeView TreeView1 
      Height          =   975
      Left            =   5400
      TabIndex        =   30
      Top             =   5880
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1720
   End
   Begin vbExtra.SSTabEx SSTabEx1 
      Height          =   1335
      Left            =   10560
      TabIndex        =   29
      Top             =   4440
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   2355
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Themed          =   -1  'True
      TabCaption(0)   =   "Tab 0"
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Tab 1"
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlCount=   0
   End
   Begin vbExtra.SizeGrip SizeGrip1 
      Height          =   285
      Left            =   15390
      Top             =   7050
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
   End
   Begin vbExtra.DateEnter DateEnter1 
      Height          =   315
      Left            =   9960
      TabIndex        =   28
      Top             =   3600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "__/__/____"
   End
   Begin vbExtra.ToolBarDA ToolBarDA1 
      Height          =   405
      Left            =   9000
      Top             =   3480
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   714
      ButtonsCount    =   1
      ButtonWidth1    =   438
   End
   Begin vbExtra.PopupList PopupList1 
      Height          =   375
      Left            =   11280
      TabIndex        =   27
      Top             =   2880
      Width           =   3030
      _ExtentX        =   5345
      _ExtentY        =   661
      Text            =   "PopupList1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin vbExtra.FontSizeChanger FontSizeChanger1 
      Height          =   375
      Left            =   9960
      TabIndex        =   26
      Top             =   2880
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   661
   End
   Begin vbExtra.History History1 
      Height          =   345
      Left            =   9120
      TabIndex        =   25
      Top             =   2880
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   609
      Enabled         =   0   'False
      Context         =   "Frmvbccr_History1"
      BackButtonToolTipText=   "Ir a ítem anterior (o click con el botón derecho para seleccionar)"
      ForwardButtonToolTipText=   "Ir a ítem siguiente (o click con el botón derecho para seleccionar)"
   End
   Begin vbExtra.ButtonExNoFocus ButtonExNoFocus1 
      Height          =   495
      Left            =   10680
      Top             =   2160
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      ShowFocusRect   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "ButtonExNoFocus1"
      UseMaskCOlor    =   -1  'True
   End
   Begin vbExtra.ButtonEx ButtonEx1 
      Height          =   495
      Left            =   9120
      TabIndex        =   24
      Top             =   2160
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      ShowFocusRect   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "ButtonEx1"
      UseMaskCOlor    =   -1  'True
   End
   Begin vbExtra.ScrollableContainer ScrollableContainer1 
      Height          =   1125
      Left            =   12480
      TabIndex        =   23
      Top             =   1080
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1984
      SavedVScrollMax =   73
      SavedVirtualHeight=   1095
      SavedHScrollMax =   87
      SavedVirtualWidth=   1305
   End
   Begin vbExtra.FontPicker FontPicker1 
      Height          =   375
      Left            =   12240
      TabIndex        =   22
      Top             =   600
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   661
   End
   Begin vbExtra.CommonDialogEx CommonDialogEx1 
      Left            =   11040
      Top             =   480
      _ExtentX        =   900
      _ExtentY        =   900
      MaxFileSize     =   255
      FontName        =   ""
   End
   Begin vbExtra.MouseWheelEnabler MouseWheelEnabler1 
      Left            =   10920
      Top             =   1320
      _ExtentX        =   900
      _ExtentY        =   900
   End
   Begin vbExtra.PrintFn PrintFn1 
      Left            =   9960
      Top             =   1200
      _ExtentX        =   900
      _ExtentY        =   900
      PrintPrevUseOneToolBar=   -1  'True
      PageNumbersFormat=   ""
      PageNumbersFormatIndex=   0
      BeginProperty PageNumbersFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin vbExtra.ComboFn ComboFn1 
      Height          =   510
      Left            =   8880
      TabIndex        =   21
      Top             =   1200
      Width           =   510
      _ExtentX        =   900
      _ExtentY        =   900
   End
   Begin vbExtra.FlexFn FlexFn1 
      Height          =   405
      Left            =   9000
      TabIndex        =   20
      Top             =   360
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   714
      PageNumbersFormat=   ""
      PageNumbersFormatIndex=   0
      BeginProperty PageNumbersFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VBFLXGRD15.VBFlexGrid VBFlexGrid1 
      Height          =   6855
      Left            =   7560
      TabIndex        =   19
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   12091
   End
   Begin VBCCR17.ToolBar ToolBar1 
      Align           =   3  'Align Left
      Height          =   6960
      Left            =   0
      Top             =   0
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   12277
      Orientation     =   1
      ButtonHeight    =   47
      ButtonWidth     =   33
      InitButtons     =   "Frmvbccr.frx":00E4
   End
   Begin VBCCR17.FontCombo FontCombo1 
      Height          =   315
      Left            =   840
      TabIndex        =   18
      Top             =   5760
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      Text            =   "Frmvbccr.frx":0104
   End
   Begin VBCCR17.CommandLink CommandLink1 
      Height          =   615
      Left            =   4080
      TabIndex        =   17
      Top             =   3840
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1085
      Caption         =   "Frmvbccr.frx":0138
   End
   Begin VBCCR17.LinkLabel LinkLabel1 
      Height          =   495
      Left            =   2640
      TabIndex        =   16
      Top             =   3840
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Caption         =   "Frmvbccr.frx":0170
   End
   Begin VBCCR17.Animation Animation1 
      Height          =   615
      Left            =   720
      TabIndex        =   15
      Top             =   4320
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1085
   End
   Begin VBCCR17.CheckBoxW CheckBoxW1 
      Height          =   495
      Left            =   840
      TabIndex        =   14
      Top             =   6120
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      Caption         =   "CheckBoxW1"
   End
   Begin VBCCR17.ComboBoxW ComboBoxW1 
      Height          =   315
      Left            =   720
      TabIndex        =   13
      Top             =   3840
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   556
      Text            =   "Frmvbccr.frx":01B2
   End
   Begin VBCCR17.CommandButtonW CommandButtonW1 
      Height          =   735
      Left            =   4800
      TabIndex        =   12
      Top             =   3000
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1296
      Caption         =   "CommandButtonW1"
   End
   Begin VBCCR17.CoolBar CoolBar1 
      Height          =   810
      Left            =   2640
      Top             =   2880
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1429
      InitBands       =   "Frmvbccr.frx":01E6
   End
   Begin VBCCR17.DTPicker DTPicker1 
      Height          =   375
      Left            =   720
      TabIndex        =   11
      Top             =   2880
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Value           =   43455
   End
   Begin VBCCR17.FrameW FrameW1 
      Height          =   615
      Left            =   840
      TabIndex        =   33
      Top             =   5040
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1085
      Caption         =   "FrameW1"
   End
   Begin VBCCR17.HotKey HotKey1 
      Height          =   615
      Left            =   5520
      TabIndex        =   10
      Top             =   2280
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1085
   End
   Begin VBCCR17.ImageCombo ImageCombo1 
      Height          =   330
      Left            =   2520
      TabIndex        =   9
      Top             =   2280
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      Text            =   "Frmvbccr.frx":03BA
   End
   Begin VBCCR17.ImageList ImageList1 
      Left            =   5760
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
      InitListImages  =   "Frmvbccr.frx":03F0
   End
   Begin VBCCR17.IPAddress IPAddress1 
      Height          =   495
      Left            =   720
      TabIndex        =   8
      Top             =   2280
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
   End
   Begin VBCCR17.ListBoxW ListBoxW1 
      Height          =   840
      Left            =   5040
      TabIndex        =   6
      Top             =   1320
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1482
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
   End
   Begin VBCCR17.ListView ListView1 
      Height          =   855
      Left            =   3240
      TabIndex        =   5
      Top             =   1320
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1508
   End
   Begin VBCCR17.MCIWnd MCIWnd1 
      Height          =   735
      Left            =   720
      TabIndex        =   4
      Top             =   1320
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1296
   End
   Begin VBCCR17.MonthView MonthView1 
      Height          =   2325
      Left            =   2640
      TabIndex        =   3
      Top             =   4560
      Width           =   2460
      _ExtentX        =   3757
      _ExtentY        =   4075
      Value           =   43455
   End
   Begin VBCCR17.OptionButtonW OptionButtonW1 
      Height          =   375
      Left            =   5760
      TabIndex        =   2
      Top             =   360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "OptionButtonW1"
   End
   Begin VBCCR17.Pager Pager1 
      Height          =   615
      Left            =   6840
      Top             =   600
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
   End
   Begin VBCCR17.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   3720
      Top             =   240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      Step            =   10
   End
   Begin VBCCR17.RichTextBox RichTextBox1 
      Height          =   975
      Left            =   2160
      TabIndex        =   1
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1720
      MultiLine       =   -1  'True
      Text            =   "Frmvbccr.frx":0410
      TextRTF         =   "Frmvbccr.frx":0448
   End
   Begin VBCCR17.Slider Slider1 
      Height          =   615
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1085
   End
   Begin VBCCR17.UpDown UpDown1 
      Height          =   375
      Left            =   5640
      Top             =   5400
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      Value           =   1
   End
   Begin VBCCR17.LabelW LabelW1 
      Height          =   375
      Left            =   3840
      TabIndex        =   7
      Top             =   840
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "LabelW1"
   End
End
Attribute VB_Name = "Frmvbccr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MCIWnd1_Click()
MCIWnd1.ShowOpen
Debug.Print MCIWnd1.FileName
End Sub
