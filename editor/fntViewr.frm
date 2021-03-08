VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.OCX"
Begin VB.Form FntViewer 
   Caption         =   "Visualizar de fontes"
   ClientHeight    =   7860
   ClientLeft      =   1545
   ClientTop       =   795
   ClientWidth     =   12405
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7860
   ScaleWidth      =   12405
   Begin VB.CommandButton CmdPreview 
      Height          =   315
      Left            =   1560
      Picture         =   "fntViewr.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Abrir Documento Existente"
      Top             =   0
      Width           =   360
   End
   Begin VB.CommandButton CmdSalvar 
      Height          =   315
      Left            =   120
      Picture         =   "fntViewr.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Salvar Documento Atual"
      Top             =   0
      Width           =   360
   End
   Begin VB.CommandButton CmdImprimir 
      Height          =   315
      Left            =   600
      Picture         =   "fntViewr.frx":0914
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Imprimir Documento Atual"
      Top             =   0
      Width           =   360
   End
   Begin VB.CommandButton Cmdconfimp 
      Height          =   315
      Left            =   1080
      Picture         =   "fntViewr.frx":0E9E
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Configurar Impressoras"
      Top             =   0
      Width           =   360
   End
   Begin VB.CommandButton CmdExit 
      Height          =   315
      Left            =   2520
      Picture         =   "fntViewr.frx":1228
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Sair"
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton CmdNextPage 
      Height          =   315
      Left            =   2040
      Picture         =   "fntViewr.frx":17B2
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Localizar Proxima"
      Top             =   0
      Width           =   360
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fntViewr.frx":1D3C
            Key             =   "s_Key1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fntViewr.frx":22D6
            Key             =   "s_Key2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fntViewr.frx":2870
            Key             =   "s_Key3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fntViewr.frx":2E0A
            Key             =   "s_Key4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fntViewr.frx":33A4
            Key             =   "s_Key5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fntViewr.frx":393E
            Key             =   "s_Key6"
         EndProperty
      EndProperty
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   6855
      Left            =   7800
      TabIndex        =   3
      Top             =   720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   7560
      Visible         =   0   'False
      Width           =   7695
   End
   Begin VB.PictureBox Container 
      AutoRedraw      =   -1  'True
      Height          =   6855
      Left            =   120
      ScaleHeight     =   6795
      ScaleWidth      =   7635
      TabIndex        =   0
      Top             =   720
      Width           =   7695
      Begin VB.PictureBox Document 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   7095
         Left            =   0
         ScaleHeight     =   7035
         ScaleWidth      =   7635
         TabIndex        =   1
         Top             =   0
         Width           =   7695
         Begin VB.Frame FraProgress 
            BorderStyle     =   0  'None
            Height          =   1575
            Left            =   600
            TabIndex        =   4
            Top             =   600
            Visible         =   0   'False
            Width           =   6375
            Begin VB.CommandButton CmdCancel 
               Caption         =   "Cancelar"
               Height          =   555
               Left            =   2880
               Picture         =   "fntViewr.frx":3ED8
               Style           =   1  'Graphical
               TabIndex        =   7
               Top             =   960
               Width           =   1335
            End
            Begin MSComctlLib.ProgressBar ProgressBar1 
               Height          =   375
               Left            =   120
               TabIndex        =   5
               Top             =   480
               Width           =   6135
               _ExtentX        =   10821
               _ExtentY        =   661
               _Version        =   393216
               Appearance      =   1
               Max             =   1
            End
            Begin VB.Label Label2 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3720
               TabIndex        =   8
               Top             =   120
               Width           =   2535
            End
            Begin VB.Label Label1 
               Caption         =   "Aguarde Comparando Fontes"
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   375
               Left            =   120
               TabIndex        =   6
               Top             =   0
               Width           =   3615
            End
         End
      End
   End
   Begin VB.Label lblstatus 
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   360
      Width           =   8055
   End
   Begin VB.Menu mnufile 
      Caption         =   "&Opcoes"
      Tag             =   "&Opcoes|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      Begin VB.Menu mnufilesub 
         Caption         =   "Selecionar Fontes"
         Index           =   0
         Tag             =   "Selecionar Fontes|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
         Begin VB.Menu mnufontssub 
            Caption         =   "Screen(Tela)"
            Index           =   0
            Tag             =   "Screen(Tela)|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
         End
         Begin VB.Menu mnufontssub 
            Caption         =   "Printer(Impressora)"
            Index           =   1
            Tag             =   "Printer(Impressora)|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
         End
         Begin VB.Menu mnufontssub 
            Caption         =   "Comparar"
            Index           =   2
            Tag             =   "Comparar|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
         End
      End
      Begin VB.Menu mnufilesub 
         Caption         =   "&Proxima Pagina"
         Enabled         =   0   'False
         Index           =   1
         Shortcut        =   ^N
         Tag             =   "&Proxima Pagina|#s_Key5|(Checked=0)(Enabled=0)(Visible=-1)(WindowList=0){^N}"
      End
      Begin VB.Menu mnufilesub 
         Caption         =   "&Imprimir"
         Index           =   2
         Tag             =   "&Imprimir|#s_Key2|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu mnufilesub 
         Caption         =   "Configurar Impressora"
         Index           =   3
         Tag             =   "Configurar Impressora|#s_Key4|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu mnufilesub 
         Caption         =   "Visualizar Impressao"
         Index           =   4
         Tag             =   "Visualizar Impressao|#s_Key3|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu mnufilesub 
         Caption         =   "Salvar"
         Index           =   5
         Tag             =   "Salvar|#s_Key6|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu mnufilesub 
         Caption         =   "Sair"
         Index           =   6
         Tag             =   "Sair|#s_Key1|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
   End
End
Attribute VB_Name = "FntViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Mode As String, X1 As Single, y1 As Single, L1 As Boolean
Public Col As Integer, Row As Integer, FontNo As Integer
Public DocLayOut As Integer, FontObj As Object, ViewMode As String
Private Page As Integer, OldFont As StdFont, Cancel As Boolean
Private WithEvents MenuEvents As CEvents
Attribute MenuEvents.VB_VarHelpID = -1

Private Sub CompareFonts()
    Dim PF As Integer, SF As Integer, NF As Integer, Flag As Integer
    Dim FCount As Long
    FraProgress.Visible = True
    Document.Cls
    DoEvents
    Document.Font = OldFont
   '' Document.Print "SCREEN FONTS NOT FOUND ON PRINTER"
    lblstatus.Caption = "Fontes de tela nao encontradas na impressora"
    ''Document.Print
    ProgressBar1.Max = Screen.FontCount - 1
    FCount = Screen.FontCount / 10
    F1Count = Printer.FontCount / 10
    FCount = ((FCount * F1Count) * 100)
    For SF = 0 To Screen.FontCount - 1
        Flag = 0
        For PF = 0 To Printer.FontCount - 1
            Flag = StrComp(Screen.Fonts(SF), Printer.Fonts(PF), vbTextCompare)
            If Flag = 0 Then
                Exit For
            End If
        Next PF
        ProgressBar1.Value = SF
        If Flag <> 0 Then
            Document.Print Screen.Fonts(SF)
            NF = NF + 1
        End If
        If SF Mod 2 = 0 Then
            Label2 = Format((SF * Printer.FontCount), "###,###,###") & " of " & Format(FCount, "###,###,###")
            DoEvents
             If Cancel Then
                ProgressBar1.Value = 0
                FraProgress.Visible = False
                Exit Sub
            End If
        End If
    Next SF
    Document.Print
    Document.Print "   " & NF
    FraProgress.Visible = False
End Sub

Private Sub KeyEvent(KeyCode As Integer)
    Select Case KeyCode
        Case 107
            mnuFileSub_Click 1
    End Select
End Sub
Private Sub SetFonts()
    Dim FontStr() As String
    ReDim FontStr(3)
    FontStr(0) = "abcdefghijklmnopqrstuvwxyz"
    FontStr(1) = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    FontStr(2) = "1234567890"
    With Document
        .FontUnderline = True
        .FontBold = True
        .FontItalic = False
        .Font = OldFont
        ''Document.Print "There are " & FontObj.FontCount & " " & Mode & " Fonts              Page No." & Page
        lblstatus.Caption = "Existem " & FontObj.FontCount & " " & Mode & " Fontes              Pag. No." & Page
        For i = FontNo To FontObj.FontCount - 1
            .Font = OldFont
            Document.Print
            If .CurrentY + (.TextHeight("A") * 5) < .Height Then
                .FontUnderline = True
                .ForeColor = RGB(255, 0, 0)
                Document.Print "(" & FontNo + 1 & ") " & FontObj.Fonts(i) & " Size is " & .FontSize
                .Font = FontObj.Fonts(i)
                .ForeColor = RGB(0, 0, 0)
                .FontUnderline = False
                Document.Print FontStr(0)
                Document.Print FontStr(1)
                Document.Print FontStr(2)
                If FontNo < FontObj.FontCount Then
                    FontNo = FontNo + 1
                End If
            Else
                If FontNo < FontObj.FontCount Then
                    mnufilesub(1).enabled = True
                Else
                    mnufilesub(1).enabled = False
                End If
                Exit Sub
            End If
        Next
    End With
End Sub
Public Sub PView()
    Dim PicHeightR As Double, PicWidthR As Double
    Mode = "Preview"
    PicHeightR = 1
    PicWidthR = 1
    Container.Cls
    If Document.ScaleHeight > Document.ScaleWidth Then
        SWidth = (Document.ScaleWidth / Document.ScaleHeight) * Container.ScaleHeight
    Else
        SWidth = (Document.ScaleHeight / Document.ScaleWidth) * Container.ScaleHeight
    End If
    SHeight = Container.ScaleHeight
    Container.PaintPicture Document.Image, 0, 0, SWidth, SHeight
End Sub
Public Sub SizeScrolls()
    With VScroll1
        .Left = Container.Left + Container.Width
        .Top = Container.Top
        .Max = Document.Height - Container.ScaleHeight '32,767
        .Min = MTOP
        .Value = .Min
        .Height = Container.Height
        .SmallChange = Container.Height / 10
        .LargeChange = Container.Height
    End With
    If Document.ScaleHeight > Container.ScaleHeight Then
        VScroll1.Visible = True
    Else
        VScroll1.Visible = False
    End If
    With HScroll1
        .Left = Container.Left
        .Top = Container.Top + Container.Height
        .Min = MLEFT
        .Width = Container.Width
        .Value = .Min
        .Max = Document.Width - Container.ScaleWidth
        .SmallChange = Container.ScaleWidth / 10
        .LargeChange = Container.Width
    End With
    If Document.ScaleWidth > Container.ScaleWidth Then
        HScroll1.Visible = True
    Else
        HScroll1.Visible = False
    End If
End Sub


Private Sub CmdCancel_Click()
    Cancel = True
End Sub

Private Sub CmdConfImp_Click()
mnuFileSub_Click (3)
End Sub

Private Sub CmdExit_Click()
mnuFileSub_Click (6)
End Sub

Private Sub CmdImprimir_Click()
  mnuFileSub_Click (2)
End Sub

Private Sub CmdNextPage_Click()
mnuFileSub_Click (1)
End Sub

Private Sub CmdPreview_Click()
mnuFileSub_Click (4)
End Sub

Private Sub CmdSalvar_Click()
salvarpict Me, Document.Image, , Replace(cARQRTF, NomeArq(cARQRTF, False), "")
End Sub
Private Sub Document_KeyDown(KeyCode As Integer, Shift As Integer)
    'Procedure to process all key events
    KeyEvent KeyCode
End Sub


Private Sub Form_Activate()
    VScroll1.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Procedure to process all key events
    KeyEvent KeyCode
End Sub

Private Sub Form_Load()


    '/ If gbSubClassMenu is False, the menu is not subclassed
    gbSubClassMenu = True

    If gbSubClassMenu Then SubClassMenuXP

    DocLayOut = Printer.Orientation
    Document.ScaleMode = vbTwips
    Select Case DocLayOut
        Case 1
            Document.Width = 8.5 * 1440
            Document.Height = 11 * 1440
        Case 2
            Document.Width = 11 * 1440
            Document.Height = 8.5 * 1440
    End Select
    Me.Show
    SizeScrolls
    Set FontObj = Screen
    Mode = "Screen"
    Page = 1
    Set OldFont = New StdFont
    OldFont = Document.Font
    SetFonts

End Sub

Private Sub Form_Resize()
    'This code will insure the viewer and scroll bars look the same regardless
    'the users screen resolution.
    'Resize if the form left is off the screen
    'If Me.Left + Me.Width > Screen.Width Then Me.Width = Screen.Width - Me.Left
    'Size the container to the screen size(always in TWIPS).
    Container.Width = Me.Width - ((Container.Left * 2) + VScroll1.Width)
    Container.Height = Me.ScaleHeight - ((Container.Top * 2) + HScroll1.Height)
    SizeScrolls
End Sub

Private Sub HScroll1_Change()
    Document.Left = -HScroll1.Value
End Sub

Private Sub HScroll1_KeyDown(KeyCode As Integer, Shift As Integer)
    'Procedure to process all key events
    KeyEvent KeyCode
End Sub

Private Sub HScroll1_Scroll()
    HScroll1_Change
End Sub

Private Sub mnuFileSub_Click(Index As Integer)
On Error GoTo erro
    Select Case Index ''//mnufilesub(Index).Caption
        Case 1 ''"&Next Page"
            If FontNo < FontObj.FontCount And Mode <> "Compare" Then
                Document.Cls
                VScroll1.Value = VScroll1.Min
                Page = Page + 1
                SetFonts
            Else
                Document.Cls
                VScroll1.Value = VScroll1.Min
                Page = 1
                FontNo = 0
                SetFonts
            End If
        Case 5
          CmdSalvar_Click
        Case 2 ''"&Print"
            MousePointer = 11
            ViewMode = "Print"
            FontNo = 0
            Do While FontNo < FontObj.FontCount
                Document.Cls
                DoEvents
                SetFonts
                Printer.PaintPicture Document.Image, 0, 0
                Printer.NewPage
            Loop
            Printer.EndDoc
            FontNo = 0
            Document.Cls
            SetFonts
            MousePointer = 0
        Case 4 ''"Print PreVie&w"
            If UCase(Left(mnufilesub(Index).Caption, 5)) = "VISUA" Then
               Document.Visible = False
               HScroll1.Visible = False
               VScroll1.Visible = False
               PView
               mnufilesub(Index).Caption = "Fechar Visualizacao"
            Else
               Container.Cls
               Document.Visible = True
               mnufilesub(Index).Caption = "Visualizar Impressao"
               HScroll1.Visible = True
               VScroll1.Visible = True
            End If
        Case 3 ''"Print Set&up"
            FrmPrintSetup.Show vbModal, Me
            DoEvents
        Case 6 ''"E&xit"
            Unload Me
    End Select
    Exit Sub
erro:
Select Case Err.Number
    Case Else
       ''SAYERRO
       Resume Next
End Select
End Sub

Private Sub mnuFontsSub_Click(Index As Integer)
    Document.Cls
    Select Case Index
        Case 0
            Set FontObj = Screen
            Mode = "Screen"
        Case 1
            Set FontObj = Printer
            Mode = "Printer"
        Case 2
            OldFont = Document.Font
            Mode = "Compare"
            CompareFonts
            Exit Sub
    End Select
    FontNo = 0
    Page = 1
    SetFonts
End Sub

Private Sub VScroll1_Change()
    Document.Top = -VScroll1.Value
End Sub

Private Sub VScroll1_KeyDown(KeyCode As Integer, Shift As Integer)
    'Procedure to process all key events
    KeyEvent KeyCode
End Sub

Private Sub VScroll1_Scroll()
    VScroll1_Change
End Sub

Private Sub MenuEvents_MenuHelp(ByVal MenuText As String, ByVal MenuHelp As String, ByVal enabled As Boolean)
    If gbSubClassMenu Then
        '/ this event show Decription menu-item to a StatusBar control
        '/ customize for your project, if you need this.
        '/ Note: MenuText contain the menu Caption.
        '/       MenuHelp contain the Description string.
        '/ example: sbStatusBar.Panels(1).text = MenuHelp
        '/          lblMenuHelp.Caption = MenuHelp
    End If
End Sub
Public Sub SubClassMenuXP()

    '/ this code is made by MenuCreator add-in

    '/ prepare the caption for subclassing. Warning! Don't remove this comment!!!
    mnufile.Caption = "&Opcoes"
          mnufilesub(0).Caption = "Selecionar Fontes"
          mnufontssub(0).Caption = "Screen(Tela)"
          mnufontssub(1).Caption = "Printer(Impressora)"
          mnufontssub(2).Caption = "Comparar"
          mnufilesub(1).Caption = "&Proxima Pagina|#s_Key5"
          mnufilesub(2).Caption = "&Imprimir|#s_Key2"
          mnufilesub(3).Caption = "Configurar Impressora|#s_Key4"
          mnufilesub(4).Caption = "Visualizar Impressao|#s_Key3"
          mnufilesub(5).Caption = "Salvar|#s_Key6"
          mnufilesub(6).Caption = "Sair|#s_Key1"

    '/ Subclassing menu. Warning! Don't remove this comment!!!

    Set MenuEvents = New CEvents
    Set objMenuEx = New cMenuEx
    Call objMenuEx.Install(Me.Hwnd, App.Path & "\" & Me.Name, ImageList1, 2, MenuEvents)

End Sub

Public Sub MenuDesigner()
    '/ Open Menu Designer tool
    objMenuEx.MenuDesigner Me.Hwnd
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '/ This Form_UnLoad is add by MenuCreator

    If gbSubClassMenu Then
        '/ prevent error if the menu is not subclassed
        On Error Resume Next
        '/ release object
        Call objMenuEx.Uninstall(Me.Hwnd, ImageList1, MenuEvents)
        Set MenuEvents = Nothing
        Set objMenuEx = Nothing
    End If

End Sub
