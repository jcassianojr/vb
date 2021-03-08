VERSION 5.00
Begin VB.Form FrmVersao 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6075
   ClientLeft      =   2190
   ClientTop       =   1455
   ClientWidth     =   8595
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   HelpContextID   =   2
   Icon            =   "FrmVersao.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   8595
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H80000000&
      Height          =   5895
      Left            =   -60
      TabIndex        =   0
      Top             =   0
      Width           =   8655
      Begin VB.CommandButton CmdEquipamento 
         Caption         =   "Equipamento"
         Height          =   375
         Left            =   3600
         TabIndex        =   15
         Top             =   5400
         Width           =   1815
      End
      Begin VB.CommandButton CmdAddRemove 
         Caption         =   "Adcionar/Remover "
         Height          =   375
         Left            =   3600
         TabIndex        =   14
         Top             =   5040
         Width           =   1815
      End
      Begin VB.CommandButton cmdSysInfo 
         Caption         =   "&Informação do Sistema"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5640
         Picture         =   "FrmVersao.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   5040
         Width           =   1455
      End
      Begin VB.CommandButton cmdSalvar 
         Caption         =   "&Retornar"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   7200
         Picture         =   "FrmVersao.frx":0B14
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   5040
         Width           =   1215
      End
      Begin VB.PictureBox Picture1 
         Height          =   1515
         Left            =   150
         Picture         =   "FrmVersao.frx":109E
         ScaleHeight     =   1455
         ScaleWidth      =   3615
         TabIndex        =   7
         Top             =   240
         Width           =   3675
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         Caption         =   "www.aplisoft.com.br"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4500
         TabIndex        =   11
         Top             =   210
         Width           =   2985
      End
      Begin VB.Label Label1 
         Caption         =   "Este produto está licenciado para:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   270
         TabIndex        =   10
         Top             =   3345
         Width           =   4785
      End
      Begin VB.Label Label2 
         Caption         =   "...."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   270
         TabIndex        =   9
         Top             =   3705
         Width           =   6375
      End
      Begin VB.Label lblsistema 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         Caption         =   "Sistema"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   945
         Left            =   3960
         TabIndex        =   8
         Top             =   840
         Width           =   4425
      End
      Begin VB.Label lblversao 
         BackColor       =   &H8000000A&
         Caption         =   "Versão"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   345
         TabIndex        =   6
         Top             =   2925
         Width           =   1455
      End
      Begin VB.Label lblcopyright 
         BackColor       =   &H8000000A&
         Caption         =   "Copyright:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3990
         TabIndex        =   5
         Top             =   4050
         Width           =   4425
      End
      Begin VB.Label lbl6 
         BackColor       =   &H8000000A&
         Caption         =   "Atenção! Copia e Distribuição           Não Autorizada."
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   660
         TabIndex        =   4
         Top             =   5385
         Width           =   2565
      End
      Begin VB.Label lblinfo 
         BackColor       =   &H8000000A&
         Caption         =   "Fone/Fax:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4020
         TabIndex        =   3
         Top             =   4560
         Width           =   2595
      End
      Begin VB.Label lblcompany 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         Caption         =   "Aplisoft"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   675
         Index           =   0
         Left            =   705
         TabIndex        =   2
         Top             =   2220
         Width           =   7695
      End
      Begin VB.Label lblcompany 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         Caption         =   "Aplisoft Softwares Aplicativos"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   19.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1095
         Index           =   1
         Left            =   -240
         TabIndex        =   1
         Top             =   4185
         Width           =   4575
      End
   End
End
Attribute VB_Name = "FrmVersao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const nFORMID = 501
Const cFORMID = "" '"Tela de Apresentaçao"
' Reg Key Security Options...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long


Private Sub CmdAddRemove_Click()
 ExecutarDLL "appwiz.cpl"
End Sub

Private Sub CmdEquipamento_Click()
 ExecutarDLL "sysdm.cpl"
End Sub

Private Sub cmdSalvar_Click()
Unload Me
End Sub

Private Sub cmdSysInfo_Click()
  Call StartSysInfo
End Sub

Private Sub Form_Load()
    Me.Left = (Screen.Width - Me.Width) / 2
    lblversao = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblsistema = App.ProductName
    lblcompany(0) = App.FileDescription
    lblcompany(1) = App.CompanyName
    lblcopyright = App.LegalCopyright
    lblinfo = App.Comments
    Me.Caption = cFORMID
    HelpContextID = nFORMID
    Label2.Caption = mCognomeEmpresa
    'Carrega imagem background
''    If fileexist(pegpath("PATH", "LOGO")) Then
      ''  Picture1 = LoadPicture(pegpath("PATH", "LOGO"))
'        Picture2.Picture = LoadPicture(pegpath("PATH", "LOGO"))
        'Picture2.Height = Me.Height - 3395
        'Picture2.Width = Me.Width - 4500
''    End If



End Sub
Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' Try To Get System Info Program Path\Name From Registry...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' Try To Get System Info Program Path Only From Registry...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' Validate Existance Of Known 32 Bit File Version
        If (dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' Error - File Can Not Be Found...
        Else
            GoTo SysInfoErr
        End If
    ' Error - Registry Entry Can Not Be Found...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    Alert "Nao foi possivel obter informaco do sistema", vbOKOnly
End Sub

Public Function GetKeyValue(keyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' Loop Counter
    Dim rc As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(keyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
    tmpVal = String$(1024, 0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
    Else                                                    ' WinNT Does NOT Null Terminate String...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Search Data Types...
    Case REG_SZ                                             ' String Registry Key Data Type
        KeyVal = tmpVal                                     ' Copy String Value
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
    End Select
    
    GetKeyValue = True                                      ' Return Success
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:      ' Cleanup After An Error Has Occured...
    KeyVal = ""                                             ' Set Return Val To Empty String
    GetKeyValue = False                                     ' Return Failure
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function


