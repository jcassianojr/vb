VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Begin VB.Form frmIniEditor 
   Caption         =   "IniEditor"
   ClientHeight    =   7470
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11325
   Icon            =   "frmIniEditor.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7470
   ScaleWidth      =   11325
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   4
      Top             =   7155
      Width           =   11325
      _ExtentX        =   19976
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            Object.Width           =   1588
            MinWidth        =   1587
            Picture         =   "frmIniEditor.frx":058A
            TextSave        =   "14:02"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            Object.Width           =   2302
            MinWidth        =   2293
            Picture         =   "frmIniEditor.frx":0B24
            TextSave        =   "17/07/2024"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Picture         =   "frmIniEditor.frx":10BE
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9684
         EndProperty
      EndProperty
   End
   Begin VB.Frame frmFile 
      Caption         =   "Arquivo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   60
      TabIndex        =   3
      Top             =   0
      Width           =   10305
      Begin XPControls.XPButton cmdFile 
         Height          =   435
         Index           =   1
         Left            =   120
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   767
         Picture         =   "frmIniEditor.frx":1458
         Caption         =   "Novo"
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
      Begin XPControls.XPButton cmdFile 
         Height          =   435
         Index           =   0
         Left            =   1200
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   767
         Picture         =   "frmIniEditor.frx":19F2
         Caption         =   "Abrir"
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
      Begin XPControls.XPButton cmdFile 
         Height          =   435
         Index           =   2
         Left            =   2280
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   240
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   767
         Picture         =   "frmIniEditor.frx":1F8C
         Caption         =   "Salvar"
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
      Begin XPControls.XPButton cmdFile 
         Height          =   435
         Index           =   3
         Left            =   3360
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   240
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   767
         Picture         =   "frmIniEditor.frx":2526
         Caption         =   "Salvar Como"
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
      Begin XPControls.XPButton cmdFile 
         Height          =   435
         Index           =   6
         Left            =   7440
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   767
         Picture         =   "frmIniEditor.frx":2AC0
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
      Begin XPControls.XPButton cmdFile 
         Height          =   435
         Index           =   5
         Left            =   4800
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   767
         Picture         =   "frmIniEditor.frx":305A
         Caption         =   "Editar"
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
      Begin XPControls.XPButton cmdFile 
         Height          =   435
         Index           =   4
         Left            =   6120
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   767
         Picture         =   "frmIniEditor.frx":35F4
         Caption         =   "Desfazer"
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
   Begin VB.Frame frmSection 
      Caption         =   "Seção"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   60
      TabIndex        =   1
      Top             =   870
      Width           =   3825
      Begin XPControls.XPButton cmdSection 
         Height          =   435
         Index           =   0
         Left            =   120
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   767
         Picture         =   "frmIniEditor.frx":3B8E
         Caption         =   "Abrir"
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
      Begin XPControls.XPButton cmdSection 
         Height          =   435
         Index           =   1
         Left            =   1200
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   767
         Picture         =   "frmIniEditor.frx":4128
         Caption         =   "Nova"
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
      Begin XPControls.XPButton cmdSection 
         Height          =   435
         Index           =   2
         Left            =   2280
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   767
         Picture         =   "frmIniEditor.frx":46C2
         Caption         =   "apagar"
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
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   5355
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   9446
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Frame frmKey 
      Caption         =   "Chaves"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   4140
      TabIndex        =   2
      Top             =   870
      Width           =   6375
      Begin XPControls.XPButton cmdkey 
         Height          =   435
         Index           =   0
         Left            =   120
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   767
         Picture         =   "frmIniEditor.frx":4B5C
         Caption         =   "Abrir"
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
      Begin XPControls.XPButton cmdkey 
         Height          =   435
         Index           =   1
         Left            =   1200
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   767
         Picture         =   "frmIniEditor.frx":50F6
         Caption         =   "Nova"
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
      Begin XPControls.XPButton cmdkey 
         Height          =   435
         Index           =   2
         Left            =   2280
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   767
         Picture         =   "frmIniEditor.frx":5690
         Caption         =   "apagar"
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
      Begin XPControls.XPButton cmdkey 
         Height          =   435
         Index           =   3
         Left            =   3480
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   767
         Picture         =   "frmIniEditor.frx":5B2A
         Caption         =   "Editar"
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
      Begin XPControls.XPButton cmdkey 
         Height          =   435
         Index           =   4
         Left            =   4800
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   767
         Picture         =   "frmIniEditor.frx":60C4
         Caption         =   "arquivar"
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
End
Attribute VB_Name = "frmIniEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const nFORMID = 1056
Const cFORMID = "Editor de Arquivos INI"
Private oIni As New cIniSettings
''Private oReg As New cRegSettings

Private Sub AddKey(ParentNode As MSComctlLib.Node)

  Const sTitle = "Adcionado Chave"
  Dim A$
  Dim nodx As Node


  A$ = InputBox("Entre nome para a chave [" & ParentNode.tEXT & "]:", sTitle, "")
  If Len(A$) = 0 Then
    Exit Sub
  End If

  Set nodx = ParentNode.Child                  'First Child
  Do Until nodx Is Nothing
    If nodx.tEXT = A$ Then
      Alert "Ja existe a chave na seccao [" & ParentNode.tEXT & "].", sTitle
      Exit Sub
    End If

    Set nodx = nodx.Next
  Loop

  Set nodx = TreeView1.Nodes.Add(ParentNode.Index, tvwChild, , A$ & "=")
  nodx.Selected = True
  nodx.EnsureVisible

  ShowSave True

End Sub

Private Sub AddSection()

  Const sTitle = "Adcionando Seçao"
  Dim A$
  Dim nodx As Node


  A$ = InputBox("Entre o Nome da Nova Seçao:", sTitle, "")
  If Len(A$) = 0 Then
    Exit Sub
  End If

  For Each nodx In TreeView1.Nodes
    If nodx.Parent Is Nothing Then
      If nodx.tEXT = A$ Then
        Alert "Ja existe esta Seccao.", sTitle
        Exit Sub
      End If
    End If
  Next

  Set nodx = TreeView1.Nodes.Add(, , , A$)
  nodx.Selected = True
  nodx.EnsureVisible

  ShowSave True

End Sub

Private Sub ChangeKey(Node As MSComctlLib.Node)

  Const sTitle = "Alterando Chave"
  Dim A$
  Dim nodx As Node
  Dim Arr() As String

  Arr = Split(Node.tEXT, "=")

  A$ = InputBox("Digite o Nome da Chave:", sTitle, Arr(0))
  If Len(A$) = 0 Then
    Exit Sub
  End If

  Set nodx = Node.Parent.Child                 'First Child
  Do Until nodx Is Nothing
    If nodx.Index <> Node.Index Then
      If nodx.tEXT = A$ Then
        Alert "Chave ja existe na seccao [" & Node.Parent.tEXT & "]."
        Exit Sub
      End If
    End If
    Set nodx = nodx.Next
  Loop

  Arr(0) = A$

  If UBound(Arr) = 0 Then
    Arr(0) = Arr(0) & "="
  End If

  Node.tEXT = Join(Arr, "=")
  Node.Selected = True
  Node.EnsureVisible

  ShowSave True

End Sub

Private Sub ChangeSection(Node As MSComctlLib.Node)

  Const sTitle = "Mudando Seção"
  Dim A$
  Dim nodx As Node


  A$ = InputBox("Digite o nome da nova Seção:", sTitle, Node.tEXT)
  If Len(A$) = 0 Then
    Exit Sub
  End If

  For Each nodx In TreeView1.Nodes
    If nodx.Parent Is Nothing Then
      If nodx.Index <> Node.Index Then
        If nodx.tEXT = A$ Then
          Alert "Ja existe a seccao.", sTitle
          Exit Sub
        End If
      End If
    End If
  Next

  Node.tEXT = A$
  Node.Selected = True
  Node.EnsureVisible

  ShowSave True

End Sub

Private Sub cmdFile_Click(Index As Integer)
  Dim sFILENAME As String
  Dim sPath As String
  Dim sRECENTFILE As String
  Dim sFILTER As String

  On Error GoTo ERRTAG

  Select Case Index
  Case 0                                       'Open
    sPath = App.Path
    ''            sPATH = oReg.GetSetting("Settings", "RecentPath", App.Path)
    ''            sRECENTFILE = oReg.GetSetting("Settings", "RecentFile", "")
    sFILTER = "Ini Files" & vbNullChar & "*.ini" & vbNullChar & "All Files" & vbNullChar & "*.*"
    sFILENAME = FileOpen(Me, sFILTER, 1, sRECENTFILE, "ini", sPath, "Open Ini File")
    If Len(sFILENAME) = 0 Then
      Exit Sub
    End If
    ''          oReg.SaveSetting "Settings", "RecentFile", sFILENAME
    ''            oReg.SaveSetting "Settings", "RecentPath", GetFilePath(sFILENAME)

    oIni.Path = sFILENAME
    ShowIniData
    ShowSave False

  Case 1                                       'New
    oIni.Path = ""
    ShowIniData
    ShowSave False

  Case 2, 6                                    'Save e Sair
    If Len(oIni.Path) = 0 Then
      sPath = App.Path

      ''                sPATH = oReg.GetSetting("Settings", "RecentPath", App.Path)
      sFILTER = "Ini Files" & vbNullChar & "*.ini" & vbNullChar & "All Files" & vbNullChar & "*.*"
      sFILENAME = FileSave(Me, sFILTER, 1, , "ini", sPath, "Save Ini File")
      If Len(sFILENAME) = 0 Then
        Exit Sub
      End If
      ''                oReg.SaveSetting "Settings", "RecentFile", sFILENAME
      ''                oReg.SaveSetting "Settings", "RecentPath", GetFilePath(sFILENAME)
      oIni.Path = sFILENAME
    End If
    SaveIni

  Case 3                                       'Save As
    sPath = App.Path

    ''            sPATH = oReg.GetSetting("Settings", "RecentPath", App.Path)
    sFILTER = "Ini Files" & vbNullChar & "*.ini" & vbNullChar & "All Files" & vbNullChar & "*.*"
    sFILENAME = FileSave(Me, sFILTER, 1, , "ini", sPath, "Save Ini File As")
    If Len(sFILENAME) = 0 Then
      Exit Sub
    End If
    ''            oReg.SaveSetting "Settings", "RecentFile", sFILENAME
    ''            oReg.SaveSetting "Settings", "RecentPath", GetFilePath(sFILENAME)
    oIni.Path = sFILENAME
    SaveIni

  Case 4                                       'Undo
    ShowIniData
    ShowSave False

  Case 5                                       'Editar
    cARQRTF = oIni.Path
    FrmRTf.Show vbModal, Me



  End Select
  If Index = 6 Then
    Unload Me
  End If

  Exit Sub
ERREND:
  Exit Sub
ERRTAG:
  SayErro "FrmIniEditor"
  Resume ERREND

End Sub

Private Sub cmdKey_Click(Index As Integer)

  Dim nodx As Node
  Dim nPOS As Long


  Select Case Index
  Case 0                                       'Change
    Set nodx = TreeView1.selectedItem
    If nodx Is Nothing Then
      Beep
      Exit Sub
    End If

    If nodx.Parent Is Nothing Then           'Section
      Beep
      Exit Sub
    End If

    ChangeKey nodx

  Case 1                                       'Add
    Set nodx = TreeView1.selectedItem
    If nodx Is Nothing Then
      Beep
      Exit Sub
    End If

    If Not nodx.Parent Is Nothing Then
      Set nodx = nodx.Parent               'Section
    End If

    AddKey nodx

  Case 2                                       'Delete
    Set nodx = TreeView1.selectedItem
    If nodx Is Nothing Then
      Beep
      Exit Sub
    End If

    If nodx.Parent Is Nothing Then           'Section
      Beep
      Exit Sub
    End If
    DeleteKey nodx

  Case 3, 4                                    'Value
    Set nodx = TreeView1.selectedItem
    If nodx Is Nothing Then
      Beep
      Exit Sub
    End If

    If nodx.Parent Is Nothing Then           'Section
      Beep
      Exit Sub
    End If

    If Index = 3 Then
      KeyValue nodx
    Else
      nPOS = InStr(nodx, "=")
      ePASS01 = ""
      If nPOS > 0 Then
        ePASS01 = Mid(nodx, nPOS + 1)
      End If

      FrmPegdb.Show vbModal, Me
      If lRETU Then
        KeyValue nodx, eRETU01
      End If
    End If

  End Select

End Sub

Private Sub cmdKeyx_Click(Index As Integer)

End Sub

Private Sub cmdSection_Click(Index As Integer)

  Dim nodx As Node


  Select Case Index
  Case 0                                       'Change
    Set nodx = TreeView1.selectedItem
    If nodx Is Nothing Then
      Beep
      Exit Sub
    End If

    If Not nodx.Parent Is Nothing Then
      Set nodx = nodx.Parent
      nodx.Selected = True
    End If

    ChangeSection nodx

  Case 1                                       'Add
    AddSection

  Case 2                                       'Delete
    Set nodx = TreeView1.selectedItem
    If nodx Is Nothing Then
      Beep
      Exit Sub
    End If

    If Not nodx.Parent Is Nothing Then
      Set nodx = nodx.Parent
      nodx.Selected = True
    End If

    DeleteSection nodx


  End Select

End Sub

Private Sub DeleteKey(Node As MSComctlLib.Node)

  Const sTitle = "Key Delete"
  Dim nodx As Node


  If Not MDG("Apagar a Chave: " & Node.tEXT & "  in section [" & Node.Parent.tEXT & "]?") Then
    Exit Sub
  End If

  Set nodx = Node.Parent
  TreeView1.Nodes.Remove Node.Index

  nodx.Selected = True
  nodx.EnsureVisible

  ShowSave True

End Sub

Private Sub DeleteSection(Node As MSComctlLib.Node)

  Const sTitle = "Section Delete"
  Dim nodx As Node


  If Not MDG("Apagar a Seccao: " & Node.tEXT & "?") Then
    Exit Sub
  End If

  Do Until Node.children = 0
    TreeView1.Nodes.Remove Node.Child.Index
  Loop
  TreeView1.Nodes.Remove Node.Index

  If TreeView1.Nodes.Count > 0 Then
    Set nodx = TreeView1.Nodes(1)
    nodx.Selected = True
    nodx.EnsureVisible
  End If

  ShowSave True

End Sub

Private Sub Encerrar_Click()

End Sub

Private Sub Form_Load()
  CenterFormToScreen Me
  ''Configura Help
  Me.Caption = cFORMID
  HelpContextID = nFORMID

  ''    With oReg
  ''        .Company = App.CompanyName
  ''        .AppName = "IniEditor"
  ''        .RootHKey = hkrLocalMachine

  'Get Saved form size
  ''        Me.Width = .GetSetting("Pos", "Width", Me.Width)
  ''        Me.Height = .GetSetting("Pos", "Height", Me.Height)
  ''    End With


  ' 0-Novo,1-Abrir,2-Salvar,3-Salvacomo
  ' 4-Editar,5-Visualizar,6-Imprimir,7-Exportar

  ''Novo
  If Not aDIREITOS(0) Then
    cmdFile(1).Enabled = False
  End If
  ''Abrir
  If Not aDIREITOS(1) Then
    cmdFile(1).Enabled = False
  End If
  ''Salvar
  If Not aDIREITOS(2) Then
    cmdFile(2).Enabled = False
  End If
  ''Salvar Como
  If Not aDIREITOS(3) Then
    cmdFile(3).Enabled = False
  End If
  ''Ediçao
  If Not aDIREITOS(4) Then
    cmdSection(0).Enabled = False
    cmdSection(1).Enabled = False
    cmdSection(2).Enabled = False
    cmdkey(0).Enabled = False
    cmdkey(1).Enabled = False
    cmdkey(2).Enabled = False
    cmdkey(3).Enabled = False
    cmdkey(4).Enabled = False
    cmdFile(5).Enabled = False
  End If
  ''Visualizaçao Ainda nao Criada
  If Not aDIREITOS(5) Then
  End If
  ''Impressao e Configuraçao Impressao
  If Not aDIREITOS(6) Then
  End If
  ''Exportaçao Ainda Nao Criada
  If Not aDIREITOS(7) Then
  End If

  'Save handle to the form and
  'start subclassing to ensure minimum form size
  ''gHW = Me.hwnd
  ''Hook

  If Len(cARQRTF) > 0 Then
    oIni.Path = cARQRTF
    ShowIniData
    ShowSave False
  End If

  StatusBar1.Panels(6).tEXT = zUSER

End Sub

Private Sub Form_Resize()

  On Error Resume Next



  With TreeView1
    .Move .Left, .Top, Me.ScaleWidth - .Left - 90, Me.ScaleHeight - .Top - 90
  End With

End Sub

Private Function GetFilePath(sFILENAME As String) As String

  Dim Arr() As String
  ''Dim i As Long


  Arr = Split(sFILENAME, "\")

  ReDim Preserve Arr(UBound(Arr) - 1)
  GetFilePath = Join(Arr, "\")

End Function

Private Sub KeyValue(Node As MSComctlLib.Node, Optional ByVal A As String = "")

  Const sTitle = "Valor da Chave"

  ''Dim nodx As Node
  Dim Arr() As String
  Dim sTxt As String


  Arr = Split(Node.tEXT, "=")
  If UBound(Arr) > 0 Then
    sTxt = Arr(1)
  End If
  ''Se nao Passou pega
  If Len(A) = 0 Then
    A = InputBox("Digite Valor para Chave " & Arr(0) & ":" & vbCr & "Tecle 'DEL' para apagar este Valor.", sTitle, sTxt)
  End If
  ''Se vzio nao continua
  If Len(A) = 0 Then
    Exit Sub
  End If

  If A = "DEL" Then
    Arr(1) = ""
  Else
    Arr(1) = A
  End If

  Node.tEXT = Join(Arr, "=")
  Node.Selected = True
  Node.EnsureVisible

  ShowSave True

End Sub

Private Sub SaveIni()

  Const sTitle = "Save Ini File"
  Dim nodx As Node
  Dim Arr As Variant
  Dim i As Long
  Dim sSection As String
  Dim sKey As String
  Dim sValue As String

  If Not MDG("Salvar Alteracoes" & vbCr & "em: " & oIni.Path & " ?") Then
    Exit Sub
  End If

  Arr = oIni.GetAllSections()
  If Not IsEmpty(Arr) Then
    For i = 0 To UBound(Arr)
      sSection = Arr(i)

      oIni.DeleteSetting sSection
    Next
  End If


  For Each nodx In TreeView1.Nodes
    If nodx.Parent Is Nothing Then
      oIni.SaveSetting nodx.tEXT, "~TEMPORARY", "~TEMPORARY"
      oIni.DeleteSetting nodx.tEXT, "~TEMPORARY"  'Leave empty section
    Else
      Arr = Split(nodx.tEXT, "=")
      sKey = Arr(0)
      If UBound(Arr) > 0 Then
        sValue = Arr(1)
      Else
        sValue = ""
      End If

      oIni.SaveSetting nodx.Parent.tEXT, sKey, sValue
    End If
  Next

  ShowSave False

End Sub

Private Sub ShowIniData()

  Dim arrSections As Variant
  Dim arrKeys As Variant
  Dim i As Long
  Dim J As Long
  Dim sSection As String
  Dim sKey As String
  Dim sValue As String
  Dim nodx As Node

  TreeView1.Nodes.Clear
  Me.Caption = "AVE IniEditor"

  If Len(oIni.Path) = 0 Then
    Me.Caption = Me.Caption & ": " & "New File"
    Exit Sub
  End If
  Me.Caption = Me.Caption & ": " & oIni.Path

  arrSections = oIni.GetAllSections()
  If Not IsEmpty(arrSections) Then
    For i = 0 To UBound(arrSections)
      sSection = arrSections(i)
      Set nodx = TreeView1.Nodes.Add(, , , sSection)

      arrKeys = oIni.GetAllSettings(sSection)
      If Not IsEmpty(arrKeys) Then
        For J = 0 To UBound(arrKeys)
          sKey = arrKeys(J, 0)
          sValue = arrKeys(J, 1)

          TreeView1.Nodes.Add nodx.Index, tvwChild, , sKey & "=" & sValue
        Next
      End If
    Next
  End If

  If TreeView1.Nodes.Count > 0 Then
    Set nodx = TreeView1.Nodes(1)
    nodx.Selected = True
    nodx.EnsureVisible
  End If

End Sub

Private Sub ShowSave(bMode As Boolean)

  Dim nodx As Node

  cmdFile(2).Enabled = bMode
  cmdFile(3).Enabled = bMode
  cmdFile(4).Enabled = bMode

  cmdSection(0).Enabled = False
  cmdSection(1).Enabled = True
  cmdSection(2).Enabled = False

  cmdkey(0).Enabled = False
  cmdkey(1).Enabled = False
  cmdkey(2).Enabled = False
  cmdkey(3).Enabled = False
  cmdkey(4).Enabled = False

  Set nodx = TreeView1.selectedItem
  If Not nodx Is Nothing Then
    TreeView1_NodeClick nodx
  End If

End Sub

Private Sub TreeView1_DblClick()

  Dim nodx As Node

  Set nodx = TreeView1.selectedItem
  If nodx Is Nothing Then
    Exit Sub
  End If

  If nodx.Parent Is Nothing Then
    ChangeSection nodx
  Else
    KeyValue nodx
  End If


End Sub

Private Sub TreeView1_KeyPress(KeyAscii As Integer)

  If KeyAscii = 13 Then
    KeyAscii = 0
    TreeView1_DblClick
  End If

End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)

  If Node.Parent Is Nothing Then
    cmdSection(0).Enabled = True
    'cmdSection(1).Enabled = False
    cmdSection(2).Enabled = True

    cmdkey(0).Enabled = False
    cmdkey(1).Enabled = True
    cmdkey(2).Enabled = False
    cmdkey(3).Enabled = False
    cmdkey(4).Enabled = False
  Else
    cmdSection(0).Enabled = False
    'cmdSection(1).Enabled = False
    cmdSection(2).Enabled = False

    cmdkey(0).Enabled = True
    cmdkey(1).Enabled = True
    cmdkey(2).Enabled = True
    cmdkey(3).Enabled = True
    cmdkey(4).Enabled = True
  End If

End Sub

