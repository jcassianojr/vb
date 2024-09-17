VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{F22668DE-E08D-467B-8E41-13900013BD5F}#2.7#0"; "VBextra2.OCX"
Begin VB.Form FrmPicturePrinter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print a StdPicture centered within margins on a selected printer, paper, resolution, and orientation"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11250
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   11250
   StartUpPosition =   2  'CenterScreen
   Begin vbExtra.CommonDialogEx CommonDialogEx1 
      Left            =   9600
      Top             =   2280
      _ExtentX        =   900
      _ExtentY        =   900
      MaxFileSize     =   255
      FontName        =   ""
   End
   Begin VB.OptionButton optOrientation 
      Caption         =   "Paisagem"
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   9480
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.OptionButton optOrientation 
      Caption         =   "Retrato"
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   9480
      TabIndex        =   3
      Top             =   480
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.ListBox lstPrinters 
      CausesValidation=   0   'False
      Height          =   2370
      Left            =   60
      TabIndex        =   0
      Top             =   360
      Width           =   4095
   End
   Begin VB.ListBox lstPapers 
      CausesValidation=   0   'False
      Enabled         =   0   'False
      Height          =   2370
      Left            =   4320
      TabIndex        =   1
      Top             =   360
      Width           =   2760
   End
   Begin VB.ListBox lstResolutions 
      CausesValidation=   0   'False
      Enabled         =   0   'False
      Height          =   2370
      Left            =   7200
      TabIndex        =   2
      Top             =   360
      Width           =   2115
   End
   Begin XPControls.XPButton cmdPrint 
      Height          =   435
      Left            =   9480
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "FrmPicturePrinter.frx":0000
      Caption         =   "Imprimir "
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
      Left            =   9480
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1680
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "FrmPicturePrinter.frx":059A
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
   Begin VB.Label Label5 
      Caption         =   "Orientacao"
      Height          =   255
      Left            =   9480
      TabIndex        =   8
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Impressoras"
      Height          =   255
      Left            =   75
      TabIndex        =   7
      Top             =   60
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Tipo Papel Tamanho em 1/10 mm"
      Height          =   255
      Left            =   4080
      TabIndex        =   6
      Top             =   0
      Width           =   3015
   End
   Begin VB.Label Label4 
      Caption         =   "Resolucao (x, y) in DPI"
      Height          =   255
      Left            =   7200
      TabIndex        =   5
      Top             =   60
      Width           =   2115
   End
End
Attribute VB_Name = "FrmPicturePrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DC_PAPERS = 2
Private Const DC_PAPERSIZE = 3
Private Const DC_PAPERNAMES = 16
Private Const DC_ENUMRESOLUTIONS = 13
Private Const DC_ORIENTATION = 17

Private Type POINT
    x As Long
    y As Long
End Type

Private Declare Function DeviceCapabilities Lib "winspool.drv" _
    Alias "DeviceCapabilitiesW" ( _
    ByVal lpDeviceName As Long, _
    ByVal lpPort As Long, _
    ByVal iIndex As Long, _
    ByVal lpOutput As Long, _
    ByVal lpDevMode As Long) As Long

Private intPaperIds() As Integer
Private ptPaperSizes() As POINT
Private lngResolutions() As Long
Private Pr As Printer
Private Pic As StdPicture
Dim cEXTENSAO As String

Private Sub PRINTIMG()
    'Print the StdPicture Pic centered on the selected rrinter (Pr)
    'with the selected paper (lstPapers) at the selected quality
    '(lngResolutions) within set margins.
    Dim MarginsLR As Single
    Dim MarginsTB As Single
    Dim PrintableWidth As Single
    Dim PrintableHeight As Single
    Dim ScaleFactor As Double
    Dim ScaledWidth As Double
    Dim ScaledHeight As Double
    
    Set Printer = Pr
    With Printer
        'Set up paper:
        .PaperSize = intPaperIds(lstPapers.ListIndex)
        .PrintQuality = lngResolutions(2 * lstResolutions.ListIndex) 'We can only set one
                                                                     'value DPI value, just
                                                                     'use X here.
        If optOrientation(1).Value Then
            .Orientation = vbPRORLandscape
        Else
            .Orientation = vbPRORPortrait
        End If
        
        'Scale to paper, using 0.5" margins all around.  Could also crop
        'the image here:
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
        
        'Paint (print) the image, scaled.  Could also do the actual cropping
        'here if any were desired by specifying additional arguments:
        .PaintPicture Pic, _
                      (.Width - ScaledWidth) / 2, _
                      (.Height - ScaledHeight) / 2, _
                      ScaledWidth, _
                      ScaledHeight
        .NewPage
        .EndDoc
    End With
End Sub

Private Sub cmdPrint_Click()
  Select Case cEXTENSAO
         Case "TXT", "ZPL", "MAN"
              PrintTXT
         Case "JPG"
              PRINTIMG
  End Select
End Sub
Private Sub PrintTXT()
Dim fileFile As Integer
Dim STRBUFFER As String
  If cEXTENSAO <> "ZPL" Then 'zpl ja tem informacao papel fontes etc...
   CommonDialogEx1.Flags = cdlCFPrinterFonts
   CommonDialogEx1.ShowFont
    Printer.Font = CommonDialogEx1.Font 'todos os atributos
    
    Printer.PrintQuality = lngResolutions(2 * lstResolutions.ListIndex) 'We can only set one
                                                                     'value DPI value, just
                                                                     'use X here.

    
    If optOrientation(1).Value Then
       Printer.Orientation = vbPRORLandscape 'PrinterObjectConstants.vbPRORPortrait
    Else
      Printer.Orientation = vbPRORPortrait
    End If
    Printer.PaperSize = intPaperIds(lstPapers.ListIndex) 'PrinterObjectConstants.vbPRPSA4
    Printer.ColorMode = PrinterObjectConstants.vbPRCMMonochrome
 End If
    
  fileFile = FreeFile
  Open cARQRTF For Input As #fileFile
  Do While Not EOF(fileFile)
    'read line
    Input #fileFile, STRBUFFER
    Printer.Print STRBUFFER
  Loop
  Close fileFile
  Printer.EndDoc

End Sub
Private Sub Encerrar_Click()
  Unload Me
End Sub

Private Sub Form_Load()

    For Each Pr In Printers
        lstPrinters.AddItem Pr.DeviceName
    Next
  cEXTENSAO = EXTENSAO(cARQRTF)
    
  'Just load up a sample image.  Could just as easily be
  'from a WIA scan or something else.
  If cEXTENSAO = "JPG" Then
    Set Pic = LoadPicture(cARQRTF)
  End If
End Sub

Private Sub lstPapers_Click()
    cmdPrint.Enabled = (lstPapers.ListIndex > 0) And (lstResolutions.ListIndex > 0)
End Sub

Private Sub lstPrinters_Click()
    Dim PaperCount As Long
    Dim strPaperNames As String
    Dim i As Long
    Dim strPaperName As String
    Dim ActualLength As Long
    Dim ResolutionCount As Long
    Dim LandscapeRotation As Long
    
    Set Pr = Printers(lstPrinters.ListIndex)
    
    'Get paper IDs, names, sizes.
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
    lstPapers.Clear
    For i = 0 To PaperCount - 1
        strPaperName = Mid$(strPaperNames, 64 * i + 1, 64)
        ActualLength = InStr(strPaperName, vbNullChar) - 1
        If ActualLength > 1 Then strPaperName = Left$(strPaperName, ActualLength)
        lstPapers.AddItem CStr(intPaperIds(i)) _
                        & ": " & strPaperName _
                        & " (" & CStr(ptPaperSizes(i).x) _
                        & ", " & CStr(ptPaperSizes(i).y) & ")"
    Next
    lstPapers.Enabled = True
    
    'Get resolutions.
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
    lstResolutions.Clear
    For i = 0 To ResolutionCount - 1
        lstResolutions.AddItem "(" & CStr(lngResolutions(2 * i)) _
                             & ", " & CStr(lngResolutions(2 * i + 1)) & ")"
    Next
    lstResolutions.Enabled = True
    
    'Get orientations.
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
End Sub

Private Sub lstResolutions_Click()
    cmdPrint.Enabled = (lstPapers.ListIndex > -1) And (lstResolutions.ListIndex > -1)
End Sub
