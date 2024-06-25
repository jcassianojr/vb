VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{7020C36F-09FC-41FE-B822-CDE6FBB321EB}#1.2#0"; "vbccr18.ocx"
Object = "{451B73A5-1563-45D5-A6AC-7B2B7D30B778}#1.0#0"; "BSPrin10.ocx"
Begin VB.Form FrmSql 
   Caption         =   "Exportador de Sql"
   ClientHeight    =   6045
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   10365
   Icon            =   "Frmsql.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6045
   ScaleWidth      =   10365
   StartUpPosition =   3  'Windows Default
   Begin BSPrinter.PrintPreview PrintPreview1 
      Left            =   5640
      Top             =   5400
      _ExtentX        =   1191
      _ExtentY        =   1191
      LcK1b           =   "yefT59bnyufI583n1ufV59HnyefN58nn0+fK58nnzefR58rn1+fW5w=="
      LcK2b           =   "reeI58zn+ue059jn1Of656fni+eN543nleeM54HnlueR5w=="
      AmbientBb       =   $"Frmsql.frx":058A
   End
   Begin vbccr18.RichTextBox RichTextBox1 
      Height          =   255
      Left            =   3960
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   5520
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      MultiLine       =   -1  'True
      Text            =   "Frmsql.frx":062A
      TextRTF         =   "Frmsql.frx":0662
   End
   Begin VB.CheckBox CmdEmail 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Enviar Email"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Height          =   1155
      Left            =   7200
      TabIndex        =   5
      Top             =   720
      Width           =   3045
      Begin XPControls.XPButton CmdVisua 
         Height          =   435
         Left            =   120
         TabIndex        =   26
         Top             =   120
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   767
         Picture         =   "Frmsql.frx":07E2
         Caption         =   "Visualizar"
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
      Begin XPControls.XPButton cmdimp 
         Height          =   435
         Left            =   120
         TabIndex        =   27
         Top             =   600
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   767
         Picture         =   "Frmsql.frx":0D7C
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
         Left            =   1440
         TabIndex        =   28
         Top             =   120
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   767
         Picture         =   "Frmsql.frx":1316
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
      Begin XPControls.XPButton CmdConfImp 
         Height          =   435
         Left            =   1440
         TabIndex        =   29
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   767
         Picture         =   "Frmsql.frx":18B0
         Caption         =   "Propriedades"
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Exportar/Salvar Como"
      Height          =   1155
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   6975
      Begin XPControls.XPButton cmdexportar 
         Height          =   375
         Index           =   0
         Left            =   2520
         TabIndex        =   8
         Top             =   720
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         Picture         =   "Frmsql.frx":1E4A
         Caption         =   "xml"
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
      Begin XPControls.XPButton cmdexportar 
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Picture         =   "Frmsql.frx":23E4
         Caption         =   "xls(TAB)"
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
      Begin XPControls.XPButton cmdDeli 
         Height          =   375
         Index           =   0
         Left            =   3720
         TabIndex        =   10
         Top             =   240
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   661
         Caption         =   " "
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
      Begin XPControls.XPText TxtDeli 
         Height          =   375
         Left            =   3000
         TabIndex        =   11
         Top             =   240
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   661
         Text            =   "|"
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XPControls.XPButton cmdexportar 
         Height          =   375
         Index           =   2
         Left            =   3480
         TabIndex        =   12
         Top             =   720
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         Picture         =   "Frmsql.frx":297E
         Caption         =   "Html"
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
      Begin XPControls.XPButton cmdexportar 
         Height          =   375
         Index           =   3
         Left            =   4440
         TabIndex        =   13
         Top             =   720
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   661
         Picture         =   "Frmsql.frx":2F18
         Caption         =   "TXT"
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
      Begin XPControls.XPButton cmdexportar 
         Height          =   375
         Index           =   4
         Left            =   6120
         TabIndex        =   14
         Top             =   720
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   661
         Picture         =   "Frmsql.frx":32B2
         Caption         =   "PDF"
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
      Begin XPControls.XPButton cmdDeli 
         Height          =   375
         Index           =   1
         Left            =   3960
         TabIndex        =   15
         Top             =   240
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   661
         Caption         =   "|"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureAlignment=   4
      End
      Begin XPControls.XPButton cmdDeli 
         Height          =   375
         Index           =   2
         Left            =   4200
         TabIndex        =   16
         Top             =   240
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   661
         Caption         =   ","
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureAlignment=   4
      End
      Begin XPControls.XPButton cmdDeli 
         Height          =   375
         Index           =   3
         Left            =   4440
         TabIndex        =   17
         Top             =   240
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   661
         Caption         =   ";"
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
      Begin XPControls.XPButton cmdDeli 
         Height          =   375
         Index           =   4
         Left            =   4680
         TabIndex        =   18
         Top             =   240
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   661
         Caption         =   "#"
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
      Begin XPControls.XPButton cmdDeli 
         Height          =   375
         Index           =   5
         Left            =   4920
         TabIndex        =   19
         Top             =   240
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   661
         Caption         =   "~"
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
      Begin XPControls.XPButton XPButton1 
         Height          =   375
         Left            =   5880
         TabIndex        =   22
         Top             =   240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         Picture         =   "Frmsql.frx":364C
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureAlignment=   1
      End
      Begin XPControls.XPButton cmdexportar 
         Height          =   375
         Index           =   5
         Left            =   5280
         TabIndex        =   23
         Top             =   720
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   661
         Picture         =   "Frmsql.frx":39E6
         Caption         =   "RTF"
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
      Begin XPControls.XPButton cmdexportar 
         Height          =   375
         Index           =   6
         Left            =   1320
         TabIndex        =   25
         Top             =   720
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Picture         =   "Frmsql.frx":3D80
         Caption         =   "xls(TR)"
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
      Begin XPControls.XPButton cmdDeli 
         Height          =   375
         Index           =   6
         Left            =   5160
         TabIndex        =   35
         Top             =   240
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   661
         Caption         =   "<tab>"
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
      Begin VB.Label Label4 
         Caption         =   "Delimitador"
         Height          =   255
         Left            =   1920
         TabIndex        =   20
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
      Height          =   675
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   10155
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "Label1 - Nome do Relatorio que Estamos Imprimindo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   9375
      End
   End
   Begin VB.TextBox TxtArquivo 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "TxtArquivo"
      Top             =   1920
      Width           =   10105
   End
   Begin XPControls.XPText arquivo 
      Height          =   405
      Left            =   120
      TabIndex        =   7
      Top             =   4920
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   714
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MultiLine       =   -1  'True
   End
   Begin XPControls.XPText tabela 
      Height          =   405
      Left            =   120
      TabIndex        =   21
      Top             =   5280
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   714
      Text            =   ""
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
   Begin XPControls.XPText sql 
      Height          =   1245
      Left            =   120
      TabIndex        =   24
      Top             =   3720
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   2196
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MultiLine       =   -1  'True
   End
   Begin XPControls.XPButton CmdOrdem 
      Height          =   435
      Left            =   9000
      TabIndex        =   30
      Top             =   3120
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   767
      Picture         =   "Frmsql.frx":431A
      Caption         =   "ordernar"
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
   Begin XPControls.XPButton CmdFiltro 
      Height          =   435
      Left            =   9000
      TabIndex        =   31
      Top             =   2640
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   767
      Picture         =   "Frmsql.frx":48B4
      Caption         =   "filtrar"
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
   Begin XPControls.XPButton CmdEditar 
      Height          =   315
      Left            =   120
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   2280
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      Picture         =   "Frmsql.frx":4E4E
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
   Begin XPControls.XPButton CmdAbrirCom 
      Height          =   315
      Left            =   1200
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   2280
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Picture         =   "Frmsql.frx":53E8
      Caption         =   "Abrir Com"
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
   Begin XPControls.XPButton CmdShell 
      Height          =   315
      Left            =   2640
      TabIndex        =   34
      Top             =   2280
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   556
      Caption         =   "S"
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
   Begin VB.Label filtro 
      Caption         =   "filtro"
      Height          =   855
      Left            =   120
      TabIndex        =   6
      Top             =   2760
      Width           =   8655
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "FrmSql"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const nFORMID = 1180
Const cFORMID = "Imprimir Relatorios sql"
Private Sub CmdAbrirCom_Click()
   cARQRTF = TxtArquivo.Text
   If FileExist(cARQRTF, True) Then
       Call OpenWith(cARQRTF, OAIF_ALLOW_REGISTRATION Or OAIF_EXEC Or OAIF_FORCE_REGISTRATION, Me.hWnd)
   End If
End Sub

Private Sub CmdConfImp_Click()
    FrmPrintSetup.Show vbModal, Me
End Sub

Private Sub cmdDeli_Click(Index As Integer)
    Select Case Index
    Case 0
        TxtDeli.Text = " "
    Case 1
        TxtDeli.Text = "|"
    Case 2
        TxtDeli.Text = ","
    Case 3
        TxtDeli.Text = ";"
    Case 4
        TxtDeli.Text = "#"
    Case 5
        TxtDeli.Text = "~"
    Case 5
        TxtDeli.Text = "<tab>"
    
    End Select
End Sub

Private Sub CmdEditar_Click()
    Dim cEXTENSAO As String
    Dim nPOS As Long

    cEXTENSAO = "XXX"
    nPOS = InStrRev(TxtArquivo, ".")
    If nPOS > 0 Then
        cEXTENSAO = Mid(TxtArquivo, nPOS + 1)
    End If
    If cEXTENSAO = "RTF" Or cEXTENSAO = "TXT" Then
        cARQRTF = TxtArquivo.Text
        FrmRTf.Show vbModal, Me
    End If

End Sub

Private Sub cmdexportar_Click(Index As Integer)
    On Error Resume Next
    Dim cARQUIVO, cEXT, cSQL, tSQL, cCAMPOS, cARQTMP As String
    Dim RsExp As ADODB.Recordset
    Dim DBEXP As ADODB.Connection
    Dim RA, nPOS As Long
    Dim gsRecordSource
    Dim sFILTER As String
    Dim nDESTINO
    Dim cLIM As String
    cLIN = Chr(13) + Chr(10)


    cARQUIVO = tabela.Text
    gsRecordSource = tabela.Text
    cEXT = "TXT"
    If Len(cARQUIVO) = 0 Then
        Alert ("Escolha Uma Tabela Antes")
        Exit Sub
    End If
    cSQL = CStr(sql.Text)
    Select Case Index
            Case 0                                       'ok
                cEXT = "XML"
            Case 1, 6
                cEXT = "XLS"
            Case 2
                cEXT = "HTML"
            Case 3
                cEXT = "TXT"
            Case 4
                cEXT = "PDF"
            Case 5
                cEXT = "RTF"
    End Select


    sFILTER = "Formato (*." & cEXT & ")" & vbNullChar & "*." & cEXT
    cARQUIVO = FileSave(Me, sFILTER, 1, cEXT, , App.Path, "Salvar " & cEXT & " Como")
    TxtArquivo.Text = cARQUIVO

    If Len(cARQUIVO) = 0 Then
        Alert ("Nome Nao Definido")
        Exit Sub
    End If



    cCONN = GeracArq(arquivo.Text, , False)


    Set DBEXP = New ADODB.Connection
    Set RsExp = New ADODB.Recordset

    DBEXP.ConnectionTimeout = 120
    DBEXP.Open cCONN
    RsExp.Open cSQL, DBEXP, adOpenStatic, adLockReadOnly
    Select Case Index
            Case 0
                RsExp.Save cARQUIVO, adPersistXML
                'Case 1
                '     tSQL = "SELECT " & cCAMPOS & " INTO [" & gsRecordSource & "] IN '" & cARQUIVO & "' 'Excel 5.0;' FROM " & tabela.tEXT
                '     DBEXP.Execute tSQL, RA, adExecuteNoRecords
            Case 1, 2, 3, 4, 5, 6
                If Index = 1 Or Index = 3 Then
                    cARQTMP = cARQUIVO
                Else
                    cARQTMP = TrocaExt(cARQUIVO, "TMP")
                End If
                nDESTINO = FreeFile + 1
                Open cARQTMP For Output As #nDESTINO
                If Index = 6 Then
                    Print #nDESTINO, "<html>" & cLIN
                    Print #nDESTINO, "<body>" & cLIN
                    Print #nDESTINO, "<table border=" & Chr(34) & "1" & Chr(34) & ">" & cLIN
                    Print #nDESTINO, "<!-- cabecalho com os nomes dos campos da tabela -->" & cLIN
                    Print #nDESTINO, "<tr>" & cLIN
                    ''CAMPOS
                    Print #nDESTINO, "</tr>" & cLIN
                    While Not RsExp.EOF
                        
                        RsExp.MoveNext
                    Wend
                End If
                
                If Index = 1 Then
                    Print #nDESTINO, RsExp.GetString(adClipString, -1, Chr(9), Chr(13) + Chr(10), vbNullString)
                Else
                    If TxtDeli.Text = "<tab>" Then
                       Print #nDESTINO, RsExp.GetString(adClipString, -1, Chr(9), Chr(13) + Chr(10), vbNullString)
                    Else
                       ''o usuario pode digitar o campo txtdeli fixstr ajustas para espaco em caso de vazio
                       Print #nDESTINO, RsExp.GetString(adClipString, -1, FixStr(TxtDeli.Text, " "), Chr(13) + Chr(10), vbNullString)
                    End If
                End If
                
                If Index = 6 Then
                    Print #nDESTINO, "</table>" + cLIN
                    Print #nDESTINO, "</body>" + cLIN
                    Print #nDESTINO, "</html>" + cLIN
                End If
                
                Close nDESTINO
                If Index = 2 Then
                    txttohtml cARQTMP, cARQUIVO
                    DeleteFile cARQTMP 'Kill cARQTMP
                End If
                If Index = 4 Then
                    ePASS01 = cARQTMP
                    formConvertToPDF.Show
                End If
                If Index = 5 Then
                    txttoRTF cARQTMP, cARQUIVO
                End If
    End Select
    RsExp.Close
    DBEXP.Close
    Alert "Exportação Concluida"

    If CmdEmail.Value = 1 Then
        ePASS01 = Array("", _
                        "", _
                        "", _
                        "", _
                        "", _
                        cARQUIVO, _
                        "", _
                        False)
        FrmSendMail.Show vbModal, Me
    End If

End Sub

Private Sub CmdFiltro_Click()
    On Error GoTo errhandler
    Dim cFILTRO As String
    
    cFILTRO = FixStr(filtro)
    If aRELCFG(11) Then
        ePASS01 = ""
        FrmFiltro.Show vbModal, Me
        filtro = Replace(Replace(eRETU01, "{", ""), "}", "")
        If lRETU And Len(aRELCFG(15)) > 0 Then
            aRELCFG(15) = TrocaSqlWhere(aRELCFG(15), FixStr(filtro))
            sql.Text = aRELCFG(15)
        End If
    End If
    If Len(aRELCFG(15)) > 0 Then
        'aRELCFG(15) = TrocaSqlWhere(aRELCFG(15), FixStr(filtro))
        sql.Text = aRELCFG(15)
    End If
    
    Exit Sub
 
errhandler:
    Select Case Err.Number
    Case 20515
        SayErro "Filtro"
    Case 364
        Resume Next
    Case Else
        SayErro "Filtro"
            
    End Select
End Sub

Private Sub cmdimp_Click()
    If Extensao(TxtArquivo.Text, "TXT") Or Extensao(TxtArquivo.Text, "PDF") Or Extensao(TxtArquivo.Text, "HTML") Or Extensao(TxtArquivo.Text, "RTF") Then
        CmdVisua_Click
    Else
    
    End If
End Sub

Private Sub CmdOrdem_Click()
    ePASS01 = aARQUIVOS(0)
    ePASS02 = NomeTableSql(aRELCFG(15))
    EscArqOrdem.Show vbModal, Me
    If lRETU Then
        aRELCFG(15) = TrocaSQLOrder(aRELCFG(15), FixStr(eRETU01))
        sql.Text = aRELCFG(15)
    End If
End Sub

Private Sub CmdShell_Click()
    Dim cARQSHELL  As String
    cARQSHELL = FixStr(TxtArquivo)
    ShellEx cARQSHELL, essSW_SHOWDEFAULT, , , , Me.hWnd
End Sub

Private Sub CmdVisua_Click()
   cARQRTF = TxtArquivo.Text
   If Not FileExist(cARQRTF, True) Then
      Exit Sub
   End If
    If Extensao(cARQRTF, "TXT") Then
        ePASS03 = 1
        PrintPreview1.ShowPreview
    End If
    If Extensao(cARQRTF, "PDF") Then
        ShellEx cARQRTF, essSW_SHOWDEFAULT, , , , Me.hWnd
    End If
    If Extensao(cARQRTF, "HTML") Then
        'FrmPreview.Show vbModal, Me
        If MDG("Sim->Navegador Nao->Visualizador Interno") Then
           OpenUrl (cARQRTF)
        Else
            ePASS03 = 3
            PrintPreview1.ShowPreview
        End If
    End If
    If Extensao(cARQRTF, "RTF") Then
        RichTextBox1.LoadFile cARQRTF, rtfRTF
        ePASS02 = 3
        PrintPreview1.ShowPreview
        RichTextBox1.Text = ""
    End If
End Sub
Private Sub PrintPreview1_PrepareReport(Cancel As Boolean)
  If ePASS03 = 1 Then 'TXT
       MyPrintingTXT
   End If
   If ePASS03 = 2 Then 'rtf
      MyPrintingRTF
   End If
   If ePASS03 = 3 Then 'HTML
      MyPrintinghtml
   End If
End Sub
Public Sub MyPrintinghtml()
    Dim cTEXTO As String
    Dim cLINHA As String
    Dim LINES() As String
    Dim i As Integer

   ' If Not FileExist(cARQRTF, True) Then 'ja checado cmdvisualclick
   '     Exit Sub
   ' End If
    cTEXTO = FileText(cARQRTF)
    cTEXTO = HtmlToText(cTEXTO)

    LINES = Split(cLINHA, vbCrLf)

    For i = 0 To UBound(LINES)
        Printer.Print LINES(i)
    Next
End Sub
Public Sub MyPrintingRTF()
    PrinterEx.PrintRichTextBox RichTextBox1
End Sub
Public Sub MyPrintingTXT()
    Dim fileFile As Integer
    Dim STRBUFFER As String
  '  If Not FileExist(cARQRTF, True) Then 'ja checado cmdvisualclick
  '      Exit Sub
  '  End If
    fileFile = FreeFile
    Open cARQRTF For Input As #fileFile
    Do While Not EOF(fileFile)
        'read line
        Input #fileFile, STRBUFFER
        Printer.Print STRBUFFER
    Loop
    Close fileFile
End Sub
Private Sub Encerrar_Click()
    Unload Me
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    TeclaEnter KeyCode
End Sub

Private Sub Form_Load()


    Dim nARQUIVOS As Integer
    TxtArquivo = zUSER & Format(Now, "_DDMMYYYY_HHMMSS") & ".TXT"
    CenterFormToScreen Me
    ''Configura Help
    Me.Caption = cFORMID
    HelpContextID = nFORMID
    filtro = ""
      
    Label1 = aRELCFG(6)
    
    ' 0-Novo,1-Abrir,2-Salvar,3-Salvacomo
    ' 4-Editar,5-Visualizar,6-Imprimir,7-Exportar

    If Not aDIREITOS(6) Then cmdimp.Visible = False
    If Not aDIREITOS(6) Then CmdConfImp.Visible = False
    If Not aDIREITOS(5) Then CmdVisua.Visible = False
    'If Not aDIREITOS(7) Then Salvar(2).Visible = False
    'If Not aDIREITOS(7) Then Salvar(3).Visible = False
    'If Not aDIREITOS(7) Then Salvar(0).Visible = False
    'If Not aDIREITOS(7) Then Salvar(1).Visible = False
    If Not aDIREITOS(7) Then CmdEmail.Visible = False
    If Not aRELCFG(11) Then CmdFiltro.Visible = False
    If Not aRELCFG(11) Then filtro.Visible = False
    
    If Not aDIREITOS(4) Then CmdEditar.Visible = False
    If Not aDIREITOS(4) Then CmdShell.Visible = False
     
  
    nARQUIVOS = UBound(aARQUIVOS)
    If nARQUIVOS > 0 Then
        If Len(aARQUIVOS(0)) > 0 Then
            If FileExist(aARQUIVOS(0), True, True, aRELCFG(15)) Then
                arquivo.Text = aARQUIVOS(0)
            End If
        End If
    End If
    aRELCFG(14) = FixStr(aRELCFG(14))
    If Len(aRELCFG(14)) > 0 Then
        cFILTRO = aRELCFG(14)
        filtro = cFILTRO
    Else
        If aRELCFG(11) Then
            CmdFiltro_Click
        End If
    End If
    If Len(aRELCFG(15)) > 0 Then
        If Len(filtro) > 0 Then
            aRELCFG(15) = TrocaSqlWhere(aRELCFG(15), FixStr(filtro))
        End If
        sql.Text = aRELCFG(15)
    End If
    
    
    If aRELCFG(13) = "" Then                     'nome tabela nao preenchido pega do sql
        aRELCFG(13) = NomeTableSql(aRELCFG(15))
    End If
    tabela.Text = aRELCFG(13)
     PrintPreview1.AuxiliaryButtonVisible = PrintPreview1.PrinterExists("Microsoft Print to PDF")
    PrintPreview1.AuxiliaryButtonToolTipText = "Salvar como PDF"
     
End Sub
Public Sub PrintPreview1_AuxiliaryButtonClick(UpdateReport As Boolean)
    PrintPreview1.ShowSaveToFile "Microsoft Print to PDF", "*.pdf"
    UpdateReport = False ' we don't need to update the report in the Print preview window after this action (the default value of UpdateReport parameter is True)
End Sub
Private Sub XPButton1_Click()
    frmCharacters.Show vbModal, Me
    If lRETU Then
        TxtDeli.Text = eRETU01
    End If
End Sub


