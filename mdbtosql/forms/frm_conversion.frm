VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Frm_Conversion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Advance Access to MySQL"
   ClientHeight    =   5370
   ClientLeft      =   3975
   ClientTop       =   2055
   ClientWidth     =   7440
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_conversion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   358
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   496
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fra_progress 
      BorderStyle     =   0  'None
      Height          =   3300
      Index           =   2
      Left            =   210
      TabIndex        =   15
      Top             =   1080
      Visible         =   0   'False
      Width           =   7200
      Begin VB.CheckBox ChkDropTables 
         Caption         =   "Drop Table(s)"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2700
         TabIndex        =   41
         Top             =   3000
         Width           =   1335
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "Select &All"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5490
         TabIndex        =   39
         Top             =   2970
         Width           =   1065
      End
      Begin VB.CheckBox ChkNoData 
         Caption         =   "Copy Only Structure"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   600
         TabIndex        =   36
         Top             =   3000
         Width           =   1815
      End
      Begin VB.ListBox LstTables 
         Appearance      =   0  'Flat
         Height          =   2280
         Left            =   600
         Style           =   1  'Checkbox
         TabIndex        =   35
         Top             =   480
         Width           =   5955
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select MS Access table(s) to convert into MySQL format"
         Height          =   195
         Left            =   600
         TabIndex        =   37
         Top             =   150
         Width           =   4005
      End
   End
   Begin VB.Frame fra_progress 
      BorderStyle     =   0  'None
      Height          =   3330
      Index           =   1
      Left            =   -30
      TabIndex        =   14
      Top             =   1290
      Visible         =   0   'False
      Width           =   7200
      Begin VB.Frame FraDumpFile 
         Caption         =   "Destination dump file path:"
         Height          =   885
         Left            =   120
         TabIndex        =   20
         Top             =   2340
         Width           =   6975
         Begin VB.CommandButton CmdSaveDumpFile 
            Caption         =   "&Browse"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   5640
            TabIndex        =   22
            Top             =   360
            Width           =   1125
         End
         Begin VB.TextBox TxtDestDumpFile 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   345
            Left            =   150
            TabIndex        =   21
            Text            =   "c:\test.sql"
            Top             =   360
            Width           =   5355
         End
      End
      Begin VB.Frame FraMySQL 
         Caption         =   "Provide necessary information to establish connection with MySQL server "
         Height          =   2085
         Left            =   120
         TabIndex        =   19
         Top             =   210
         Width           =   6975
         Begin VB.CheckBox UseMariadb 
            Caption         =   "Usar Mariadb Driver"
            Height          =   255
            Left            =   360
            TabIndex        =   43
            Top             =   1680
            Width           =   3015
         End
         Begin VB.ComboBox CboType 
            Height          =   315
            ItemData        =   "frm_conversion.frx":058A
            Left            =   4770
            List            =   "frm_conversion.frx":058C
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   1230
            Width           =   1725
         End
         Begin VB.TextBox TxtMySQLInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   0
            Left            =   1470
            TabIndex        =   27
            Text            =   "localhost"
            Top             =   390
            Width           =   1935
         End
         Begin VB.TextBox TxtMySQLInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   1
            Left            =   4770
            TabIndex        =   28
            Text            =   "3306"
            Top             =   390
            Width           =   1725
         End
         Begin VB.TextBox TxtMySQLInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   2
            Left            =   1470
            TabIndex        =   29
            Text            =   "root"
            Top             =   810
            Width           =   1935
         End
         Begin VB.TextBox TxtMySQLInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            IMEMode         =   3  'DISABLE
            Index           =   3
            Left            =   1470
            PasswordChar    =   "•"
            TabIndex        =   30
            Top             =   1230
            Width           =   1935
         End
         Begin VB.TextBox TxtMySQLInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   4
            Left            =   4770
            TabIndex        =   33
            Text            =   "test"
            Top             =   810
            Width           =   1725
         End
         Begin VB.Label LblMySQLInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Table Type:"
            Height          =   195
            Index           =   5
            Left            =   3840
            TabIndex        =   32
            Top             =   1290
            Width           =   855
         End
         Begin VB.Label LblMySQLInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Database:"
            Height          =   195
            Index           =   4
            Left            =   3840
            TabIndex        =   31
            Top             =   840
            Width           =   750
         End
         Begin VB.Label LblMySQLInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Password:"
            Height          =   195
            Index           =   3
            Left            =   330
            TabIndex        =   26
            Top             =   1290
            Width           =   750
         End
         Begin VB.Label LblMySQLInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Username:"
            Height          =   195
            Index           =   2
            Left            =   330
            TabIndex        =   25
            Top             =   840
            Width           =   780
         End
         Begin VB.Label LblMySQLInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Port:"
            Height          =   195
            Index           =   1
            Left            =   3840
            TabIndex        =   24
            Top             =   420
            Width           =   360
         End
         Begin VB.Label LblMySQLInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Server/Host:"
            Height          =   195
            Index           =   0
            Left            =   330
            TabIndex        =   23
            Top             =   450
            Width           =   930
         End
      End
   End
   Begin VB.Frame fra_progress 
      BorderStyle     =   0  'None
      Height          =   3345
      Index           =   3
      Left            =   180
      TabIndex        =   16
      Top             =   1200
      Visible         =   0   'False
      Width           =   7200
      Begin MSComctlLib.ProgressBar Pbar 
         Height          =   285
         Left            =   870
         TabIndex        =   42
         Top             =   1350
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   503
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.Label LblProgressStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LblProgressStatus"
         Height          =   195
         Left            =   870
         TabIndex        =   38
         Top             =   1110
         Width           =   1290
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Database conversion is in progress. Please wait until wizard completes converting your MS access database into MySQL data source."
         Height          =   435
         Left            =   240
         TabIndex        =   17
         Top             =   210
         Width           =   5985
      End
   End
   Begin MSComDlg.CommonDialog MyDlg 
      Left            =   750
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fra_progress 
      Caption         =   "Source Database (MS Access):"
      Height          =   2865
      Index           =   0
      Left            =   450
      TabIndex        =   7
      Top             =   1410
      Visible         =   0   'False
      Width           =   6765
      Begin VB.CommandButton CmdTestConnection 
         Caption         =   "&Test Connection"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4770
         TabIndex        =   13
         Top             =   2310
         Width           =   1785
      End
      Begin VB.TextBox TxtSrcPwd 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   1800
         PasswordChar    =   "*"
         TabIndex        =   12
         Top             =   1410
         Width           =   3585
      End
      Begin VB.CheckBox ChkSrcPwd 
         Caption         =   "Use Password:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   330
         TabIndex        =   11
         Top             =   1470
         Width           =   1365
      End
      Begin VB.CommandButton CmdBrowse 
         Caption         =   "&Browse"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5460
         TabIndex        =   10
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox TxtSourceDB 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   330
         TabIndex        =   9
         Top             =   960
         Width           =   5055
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter source database name. Click 'Browse' button to find the source database throught directiry tree."
         Height          =   465
         Left            =   360
         TabIndex        =   8
         Top             =   420
         Width           =   5775
      End
   End
   Begin VB.CommandButton CmdFinish 
      Caption         =   "&Finish"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3600
      TabIndex        =   18
      Top             =   4890
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.PictureBox PicWelcome 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   0
      Picture         =   "frm_conversion.frx":058E
      ScaleHeight     =   1
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   496
      TabIndex        =   0
      Top             =   0
      Width           =   7440
      Begin VB.OptionButton OptScenario 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OptScenario"
         Height          =   195
         Index           =   0
         Left            =   2760
         TabIndex        =   2
         Top             =   3960
         Width           =   1185
      End
   End
   Begin VB.CommandButton CmdHelp 
      Caption         =   "&Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6120
      TabIndex        =   6
      Top             =   4890
      Width           =   1125
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4860
      TabIndex        =   5
      Top             =   4890
      Width           =   1125
   End
   Begin VB.CommandButton CmdNext 
      Caption         =   "&Next >"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3600
      TabIndex        =   4
      Top             =   4890
      Width           =   1125
   End
   Begin VB.PictureBox PicHeader 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   900
      Left            =   0
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   496
      TabIndex        =   3
      Top             =   15
      Width           =   7440
   End
   Begin VB.CommandButton CmdBack 
      Caption         =   "< &Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2400
      TabIndex        =   1
      Top             =   4890
      Width           =   1125
   End
   Begin VB.TextBox TxtMessage 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3165
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   40
      Top             =   1020
      Visible         =   0   'False
      Width           =   6975
   End
End
Attribute VB_Name = "Frm_Conversion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents MyConversion As Cls_Convertion
Attribute MyConversion.VB_VarHelpID = -1
Private lProgress As Integer
Private bFinish As Boolean
Private MySQLStatus As Boolean
Private MyScenario As Integer
Private IsProgress As Boolean

' check mysql connection is allready or not
Private Function CheckMySQLConnection() As Boolean



    CheckMySQLConnection = MyConversion.OpenDestinationDB(TxtMySQLInfo(0).Text, _
        TxtMySQLInfo(2).Text, TxtMySQLInfo(3).Text, TxtMySQLInfo(1).Text, UseMariadb.Value)
        
End Function

' check source datbase connection
Private Function CheckMsAccessConnection() As Boolean
    CheckMsAccessConnection = MyConversion.OpenSourceDB(TxtSourceDB.Text, TxtSrcPwd.Text)
End Function


Private Sub CboType_Click()
    ' change mysql table type
    MyConversion.SetTableType CboType.Text
End Sub

' using / required password or not
Private Sub ChkSrcPwd_Click()
    TxtSrcPwd.Enabled = ChkSrcPwd.Value
    If ChkSrcPwd.Value Then TxtSrcPwd.SetFocus
End Sub


' back to the previous progress
Private Sub CmdBack_Click()

    ' back button progress
    If lProgress > -1 Then
        lProgress = lProgress - 1
        CreatePageProgress lProgress
        DisplayFrameProgress lProgress, lProgress + 1
    End If

End Sub

' open source databse filename (MS Access databse file)
Private Sub CmdBrowse_Click()
On Error Resume Next

    Dim sFile As String
    ' set filter
    MyDlg.Filter = "Microsoft Access Database|*.mdb|All Files|*.*"
    ' set window dialog title
    MyDlg.DialogTitle = "Open Source Database"
    ' initialize
    MyDlg.FileName = ""
    ' show open dialog
    MyDlg.ShowOpen
    ' return
    sFile = MyDlg.FileName
    
    ' check the result
    If sFile <> vbNullString Then
        TxtSourceDB.Text = sFile
        TxtDestDumpFile.Text = TrocaExt(TxtSourceDB.Text, "sql")
        TxtMySQLInfo(4).Text = parsefile(TxtSourceDB.Text, "N")
    End If
    
End Sub

Private Sub CmdCancel_Click()
    
    ' try to exit the wizard
       If MsgBox("Are you sure you want to exit the wizard?", _
            vbYesNo + vbQuestion) = vbYes Then
            ' cancel progress
            MyConversion.EndProgress = True
            bFinish = True
            If Not IsProgress Then
                Unload Me
            End If
       End If
    
End Sub

Private Sub CmdFinish_Click()

    ' close wizard
    ' exit app
    Unload Me
    
End Sub

Private Sub CmdHelp_Click()

    ' display help file
    MsgBox "The Advance Access to MySQl help does not exist", vbExclamation

End Sub

Private Sub CmdNext_Click()
On Error Resume Next
Dim bResult As Boolean
    ' next button progress
    If lProgress < 3 Then
    
        If lProgress = 0 Then
            bResult = CheckMsAccessConnection
            FraMySQL.Enabled = True
            FraDumpFile.Enabled = True
            If MyScenario = 0 Then
                FraDumpFile.Enabled = False
            ElseIf MyScenario = 1 Then
                FraMySQL.Enabled = False
            End If
        ElseIf lProgress = 1 Then
            If MyScenario = 0 Then
                bResult = CheckMySQLConnection
            ElseIf MyScenario = 1 Then
                bResult = IIf(TxtDestDumpFile.Text = "", False, True)
                If Not bResult Then
                    MsgBox "You should enter the destination dump file name", vbExclamation
                End If
            Else
                bResult = (CheckMySQLConnection And IIf(TxtDestDumpFile.Text = "", False, True))
                If bResult = False Then
                    MsgBox "Please check MySQL connection and destination dump file name", vbExclamation
                 End If
            End If
        Else
            bResult = True
        End If
        
        If bResult Then
    
            lProgress = lProgress + 1
            DisplayFrameProgress lProgress, lProgress - 1
            DoEvents
            CreatePageProgress lProgress
        
        End If
        
    End If
    
    ' conversion progress
    If lProgress = 3 Then
        ' disable back button and next button
        CmdBack.Enabled = False
        CmdNext.Enabled = False
        ' start conversion
        IsProgress = True
        ' init when start progress
        MyConversion.EndProgress = False
        ' conversion method here ....
        ConversionData
        ' conversion has been finished
        DisplayEndProgress
        DisplayFrameProgress -1, 3
        CreatePageProgress 4
    ElseIf lProgress = 2 Then
        GetSrcTables
    End If
    
End Sub

' set dump file path
Private Sub CmdSaveDumpFile_Click()
    On Error Resume Next
    
    Dim sFile As String
    ' set filter
    MyDlg.Filter = "SQL Script|*.sql"
    ' set window dialog title
    MyDlg.DialogTitle = "Save database into SQL script"
    ' initialize
    MyDlg.FileName = ""
    ' show open dialog
    MyDlg.ShowSave
    ' return
    sFile = MyDlg.FileName
    
    ' check the result
    If sFile <> vbNullString Then
        TxtDestDumpFile.Text = IIf(LCase(Right(sFile, 4)) = ".sql", sFile, sFile & ".sql")
        
    End If
    
End Sub

Private Sub cmdSelect_Click()
    Dim l As Long
    If LstTables.ListCount > 0 Then
        For l = 0 To LstTables.ListCount - 1
            LstTables.Selected(l) = True
        Next l
    End If
End Sub

Private Sub CmdTestConnection_Click()
On Error Resume Next
    Dim bConn As Boolean
    
    ' try to connect to the source database
    bConn = MyConversion.OpenSourceDB(TxtSourceDB.Text, TxtSrcPwd.Text)
    
    ' display result
    If bConn Then ' connection OK
        MsgBox "Test connection successed.", vbInformation
    Else ' Connection failed
        MsgBox "Test connection failed.", vbCritical
    End If
    
End Sub



Private Sub Form_Load()
On Error Resume Next

    ' init osenxpform
    
    
    ' init class conversion
    Set MyConversion = New Cls_Convertion
    
    CreatePageProgress -1 ' draw welcome page
    lProgress = -1 ' initialize
    DrawSeparator Me
    
    ' set frame progress position
    InitFramePosition
    
    ' insert item into cbotype
    With CboType
        .AddItem "InnoDB"
        .AddItem "MyISAM"
        .AddItem "ISAM"
        .AddItem "BDB"
        .AddItem "HEAP"
        .ListIndex = 0
    End With
    
    ' draw welcome page
    fDrawWelcome PicWelcome, OptScenario

End Sub

Private Sub Form_Unload(Cancel As Integer)

    ' confirm before exit
    If bFinish Then
        
        Set MyConversion = Nothing
    Else
        If MsgBox("Are you sure you want to exit the wizard?", _
            vbYesNo + vbQuestion) = vbYes Then
            
            Set MyConversion = Nothing
            
        Else ' cancel
            Cancel = 1
        End If
    End If
End Sub

'**************************************************************************
' Created date: 2004-11-01 07:42
' Purpose: Create/Set Page dialogs
'**************************************************************************
Private Sub CreatePageProgress(Index As Integer)
    ' create error handler
    On Error Resume Next
    Static pHeaderShow As Boolean
    
    If Index > -1 And pHeaderShow = False Then
        PicWelcome.Visible = False
        PicHeader.Visible = True
        pHeaderShow = True
        CmdBack.Enabled = True
    ElseIf Index = -1 Then
        PicHeader.Height = 0
        PicHeader.Visible = False
        PicWelcome.Visible = True
        pHeaderShow = False
        CmdBack.Enabled = False
    End If
    
    ' draw pages progress
    If Index <> -1 Then
        fDrawHeaderPage PicHeader, Index
    End If
    
    
    
End Sub


'************************************************************************
' Created date: 2004-11-01 08:08
' Purpose: Display FInish Button
'************************************************************************
Private Sub DisplayEndProgress()

    
    ' disable cancel button
    CmdCancel.Enabled = False
    
    ' disable help button
    CmdHelp.Enabled = False
    
    ' hide next button
    CmdNext.Visible = False
    
    ' display finish button
    CmdFinish.Visible = True
    
    TxtMessage.Visible = True
    
    ' set finished progress
    bFinish = True

End Sub


'*************************************************************************
' Created date: 2004-11-01 08:14
' Purpose: Display selected frame by page progress status
'*************************************************************************
Private Sub DisplayFrameProgress(Index As Integer, LastPosition As Integer)
    On Error Resume Next
    
    ' show / hide or select active frame
    
    ' hide last position
    If LastPosition >= 0 And LastPosition <= 3 Then
        fra_progress(LastPosition).Visible = False
    End If
    
    ' set active frame
    If Index <> -1 Then
        fra_progress(Index).Visible = True
    End If
    
End Sub


'************************************************************************
' Created date: 2004-11-01 08:22
' Purpose: Set frame position
'************************************************************************
Private Sub InitFramePosition()

Dim Ix As Integer

    ' set framei ndex0 position {MS Access Configuration}
    With fra_progress(0)
        .Left = 22
        .Top = 78
        .Width = 451
        .Height = 191
        
    End With
    
    For Ix = 1 To 3
        With fra_progress(Ix)
            .Left = 6
            .Top = 65
            .Width = 480
            .Height = 220
        End With
    Next Ix
    
End Sub

Private Sub MyConversion_ExecuteInfo(StrSQL As String)
    If MyScenario > 0 Then Print #1, StrSQL
End Sub

Private Sub MyConversion_Progress(ProgressStatus As Long)
    Pbar.Value = ProgressStatus
End Sub

' set scenario by selected user
' 0. Move to MySQL server directly
' 1. Store into dump file
' 2. Both
Private Sub OptScenario_Click(Index As Integer)
    '---------------
    MyScenario = Index
End Sub


' get tables from source database
Private Sub GetSrcTables()
On Error Resume Next
    Dim StrV() As String
    Dim StrArray
    Dim lT As Long
    Dim j As Long
    ' get tables
    lT = MyConversion.GetSourceTables(StrV)
    
    ' clear list
    LstTables.Clear
    For j = 1 To lT
        StrArray = Split(StrV(j), ":")
        LstTables.AddItem StrArray(0)
    Next j
    
End Sub

' get total records to be convert
Private Function getTotalRecords() As Long
On Error Resume Next
Dim lTotal As Long
Dim k As Integer

    If LstTables.ListCount > 0 Then
        For k = 1 To LstTables.ListCount
            If LstTables.Selected(k - 1) Then
                lTotal = lTotal + LstTables.ItemData(k - 1)
            End If
        Next k
        getTotalRecords = lTotal
    End If
End Function

'************************************************************************
' created date: 2004-11-01 12:16
' Purpose: Conversion Progress
'************************************************************************
Private Sub ConversionData()
    On Error Resume Next
    Dim i As Long
    Dim strTable As String
    Dim StrInfo As String
    Dim lTime As Long
    Dim lDATA As Boolean
    
    If LstTables.ListCount > 0 Then
    
        If MyScenario > 0 Then
            Open TxtDestDumpFile.Text For Output As #1
            Print #1, "Create database if not exists '" & TxtMySQLInfo(4).Text & "' ;"
            
        End If
        
            ' counter function
            lTime = GetTickCount ' gettickcount
            StrInfo = getmyTime & " Start conversion ...." & vbCrLf & vbCrLf
        
        
            If MyScenario <> 1 Then
                ' CREATE MYSQL DATABASE
                MyConversion.createDB TxtMySQLInfo(4).Text
                StrInfo = StrInfo & getmyTime & " Create database if not exists '" & TxtMySQLInfo(4).Text & "'" & vbCrLf & vbCrLf
            End If
            
            For i = 0 To LstTables.ListCount - 1
                If LstTables.Selected(i) Then
                    strTable = LstTables.List(i)
                    StrInfo = StrInfo & getmyTime & " Start processing table '" & strTable & "' " & vbCrLf
                    LblProgressStatus.Caption = "Processing table '" & strTable & "'"
                    lDATA = True
                    If ChkNoData.Value = True Then
                       lDATA = False
                    End If
                    MyConversion.ConvTable strTable, lDATA, ChkDropTables.Value, IIf(MyScenario = 1, 1, 0), TxtSourceDB.Text
                    If bFinish Then
                        StrInfo = StrInfo & vbCrLf & getmyTime & " Canceled by user ....." & vbCrLf
                        Exit For
                    End If
                    StrInfo = StrInfo & getmyTime & " Table '" & strTable & "' was successed converted" & vbCrLf & vbCrLf
                End If
            Next i
        
        If MyScenario > 0 Then
            Close #1
        End If
        
            StrInfo = StrInfo & vbCrLf & "(" & Format((GetTickCount - lTime) / 1000, "0.00") & " s taken)"
            TxtMessage.Text = StrInfo
            
    End If
End Sub

Private Function getmyTime() As String
    getmyTime = "[" & Time() & "]"
End Function

