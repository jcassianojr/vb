VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{451B73A5-1563-45D5-A6AC-7B2B7D30B778}#1.1#0"; "BSPrin10.ocx"
Object = "{379157C5-E9BD-43F1-9F83-B037496BED42}#1.3#0"; "VBCCR18.OCX"
Begin VB.Form FrmSql 
   Caption         =   "Exportador de Sql"
   ClientHeight    =   7140
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   12228
   Icon            =   "Frmsql.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7140
   ScaleWidth      =   12228
   StartUpPosition =   3  'Windows Default
   Begin BSPrinter.PrintPreview PrintPreview1 
      Left            =   5520
      Top             =   6240
      _ExtentX        =   953
      _ExtentY        =   953
   End
   Begin VBCCR18.RichTextBox RichTextBox1 
      Height          =   255
      Left            =   3960
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   6480
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2350
      _ExtentY        =   445
      MultiLine       =   -1  'True
      Text            =   "Frmsql.frx":058A
      TextRTF         =   "Frmsql.frx":05C2
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
      Height          =   1635
      Left            =   6240
      TabIndex        =   5
      Top             =   720
      Width           =   4212
      Begin XPControls.XPButton CmdConfImp 
         Height          =   432
         Left            =   360
         TabIndex        =   25
         Top             =   960
         Width           =   1452
         _ExtentX        =   2561
         _ExtentY        =   762
         Picture         =   "Frmsql.frx":0742
         Caption         =   "Propriedades"
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
      Begin VBCCR18.CommandButtonW Encerrar 
         Height          =   612
         Left            =   3120
         TabIndex        =   35
         Top             =   240
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
         Picture         =   "Frmsql.frx":0CDC
         PictureAndCaption=   -1  'True
         WordWrap        =   0   'False
         Style           =   1
      End
      Begin VBCCR18.CommandButtonW cmdimp 
         Height          =   612
         Left            =   2040
         TabIndex        =   36
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
         Picture         =   "Frmsql.frx":1276
         PictureAndCaption=   -1  'True
         WordWrap        =   0   'False
         Style           =   1
      End
      Begin VBCCR18.CommandButtonW CmdVisua 
         Height          =   612
         Left            =   960
         TabIndex        =   37
         Top             =   240
         Width           =   972
         _ExtentX        =   1715
         _ExtentY        =   1080
         Appearance      =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         ImageListAlignment=   1
         Caption         =   "Visualizar"
         Alignment       =   0
         VerticalAlignment=   0
         Picture         =   "Frmsql.frx":1810
         PictureAndCaption=   -1  'True
         WordWrap        =   0   'False
         Style           =   1
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Exportar/Salvar Como"
      Height          =   1635
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   6012
      Begin XPControls.XPButton cmdexportar 
         Height          =   375
         Index           =   0
         Left            =   2520
         TabIndex        =   8
         Top             =   720
         Width           =   855
         _ExtentX        =   1503
         _ExtentY        =   656
         Picture         =   "Frmsql.frx":1DAA
         Caption         =   "xml"
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
      Begin XPControls.XPButton cmdexportar 
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   1095
         _ExtentX        =   1926
         _ExtentY        =   656
         Picture         =   "Frmsql.frx":2344
         Caption         =   "xls(TAB)"
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
      Begin XPControls.XPButton cmdDeli 
         Height          =   375
         Index           =   0
         Left            =   3720
         TabIndex        =   10
         Top             =   240
         Width           =   195
         _ExtentX        =   339
         _ExtentY        =   656
         Caption         =   " "
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
      Begin XPControls.XPText TxtDeli 
         Height          =   375
         Left            =   3000
         TabIndex        =   11
         Top             =   240
         Width           =   675
         _ExtentX        =   1185
         _ExtentY        =   656
         Text            =   "|"
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.6
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
         Left            =   3960
         TabIndex        =   12
         Top             =   1200
         Width           =   855
         _ExtentX        =   1503
         _ExtentY        =   656
         Picture         =   "Frmsql.frx":28DE
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
      Begin XPControls.XPButton cmdexportar 
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   13
         Top             =   1200
         Width           =   915
         _ExtentX        =   1609
         _ExtentY        =   656
         Picture         =   "Frmsql.frx":2E78
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
      Begin XPControls.XPButton cmdexportar 
         Height          =   375
         Index           =   4
         Left            =   3000
         TabIndex        =   14
         Top             =   1200
         Width           =   915
         _ExtentX        =   1609
         _ExtentY        =   656
         Picture         =   "Frmsql.frx":3212
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
      Begin XPControls.XPButton cmdDeli 
         Height          =   375
         Index           =   1
         Left            =   3960
         TabIndex        =   15
         Top             =   240
         Width           =   195
         _ExtentX        =   339
         _ExtentY        =   656
         Caption         =   "|"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
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
         _ExtentX        =   339
         _ExtentY        =   656
         Caption         =   ","
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
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
         _ExtentX        =   339
         _ExtentY        =   656
         Caption         =   ";"
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
      Begin XPControls.XPButton cmdDeli 
         Height          =   375
         Index           =   4
         Left            =   4680
         TabIndex        =   18
         Top             =   240
         Width           =   195
         _ExtentX        =   339
         _ExtentY        =   656
         Caption         =   "#"
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
      Begin XPControls.XPButton cmdDeli 
         Height          =   375
         Index           =   5
         Left            =   4920
         TabIndex        =   19
         Top             =   240
         Width           =   195
         _ExtentX        =   339
         _ExtentY        =   656
         Caption         =   "~"
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
      Begin XPControls.XPButton cmdexportar 
         Height          =   375
         Index           =   5
         Left            =   1080
         TabIndex        =   22
         Top             =   1200
         Width           =   915
         _ExtentX        =   1609
         _ExtentY        =   656
         Picture         =   "Frmsql.frx":35AC
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
      Begin XPControls.XPButton cmdexportar 
         Height          =   375
         Index           =   6
         Left            =   1320
         TabIndex        =   24
         Top             =   720
         Width           =   1095
         _ExtentX        =   1926
         _ExtentY        =   656
         Picture         =   "Frmsql.frx":3946
         Caption         =   "xls(TR)"
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
      Begin XPControls.XPButton cmdDeli 
         Height          =   375
         Index           =   6
         Left            =   5160
         TabIndex        =   31
         Top             =   240
         Width           =   555
         _ExtentX        =   974
         _ExtentY        =   656
         Caption         =   "<tab>"
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
      Begin XPControls.XPButton cmdexportar 
         Height          =   375
         Index           =   7
         Left            =   2040
         TabIndex        =   33
         Top             =   1200
         Width           =   915
         _ExtentX        =   1609
         _ExtentY        =   656
         Picture         =   "Frmsql.frx":3EE0
         Caption         =   "DOC"
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
      Begin XPControls.XPButton cmdexportar 
         Height          =   375
         Index           =   8
         Left            =   3480
         TabIndex        =   34
         Top             =   720
         Width           =   1095
         _ExtentX        =   1926
         _ExtentY        =   656
         Picture         =   "Frmsql.frx":427A
         Caption         =   "xls(TD)"
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
      Width           =   10392
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "Label1 - Nome do Relatorio que Estamos Imprimindo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
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
      Top             =   2400
      Width           =   10105
   End
   Begin XPControls.XPText arquivo 
      Height          =   405
      Left            =   120
      TabIndex        =   7
      Top             =   5520
      Width           =   10095
      _ExtentX        =   17801
      _ExtentY        =   720
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
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
      Top             =   6000
      Width           =   2415
      _ExtentX        =   4255
      _ExtentY        =   720
      Text            =   ""
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
   Begin XPControls.XPText sql 
      Height          =   1245
      Left            =   120
      TabIndex        =   23
      Top             =   4200
      Width           =   10095
      _ExtentX        =   17801
      _ExtentY        =   2201
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
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
      TabIndex        =   26
      Top             =   3600
      Width           =   1275
      _ExtentX        =   2244
      _ExtentY        =   762
      Picture         =   "Frmsql.frx":4814
      Caption         =   "ordernar"
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
   Begin XPControls.XPButton CmdFiltro 
      Height          =   435
      Left            =   9000
      TabIndex        =   27
      Top             =   3120
      Width           =   1275
      _ExtentX        =   2244
      _ExtentY        =   762
      Picture         =   "Frmsql.frx":4DAE
      Caption         =   "filtrar"
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
   Begin XPControls.XPButton CmdEditar 
      Height          =   315
      Left            =   120
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   2760
      Width           =   975
      _ExtentX        =   1715
      _ExtentY        =   550
      Picture         =   "Frmsql.frx":5348
      Caption         =   "Editar"
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
   Begin XPControls.XPButton CmdAbrirCom 
      Height          =   315
      Left            =   1200
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   2760
      Width           =   1335
      _ExtentX        =   2350
      _ExtentY        =   550
      Picture         =   "Frmsql.frx":58E2
      Caption         =   "Abrir Com"
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
   Begin XPControls.XPButton CmdShell 
      Height          =   315
      Left            =   2640
      TabIndex        =   30
      Top             =   2760
      Width           =   375
      _ExtentX        =   656
      _ExtentY        =   550
      Caption         =   "S"
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
   Begin VB.Label filtro 
      Caption         =   "filtro"
      Height          =   855
      Left            =   120
      TabIndex        =   6
      Top             =   3240
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
Option Explicit
Private Sub CmdAbrirCom_Click()
  cARQRTF = TxtArquivo.Text
  If FileConnExist(cARQRTF, True) Then
    Call OpenWith(cARQRTF, OAIF_ALLOW_REGISTRATION Or OAIF_EXEC Or OAIF_FORCE_REGISTRATION, CLng(Me.hwnd))
  End If
End Sub

Private Sub CmdConfImp_Click()
 ' FrmPrintSetup.Show vbModal, Me
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
  If IsExtensao(TxtArquivo.Text, "RTF") Or IsExtensao(TxtArquivo.Text, "TXT") Then
    cARQRTF = TxtArquivo.Text
    FrmRtfView.Show vbModal, Me
  End If
End Sub
Private Sub cmdexportar_Click(Index As Integer)
    On Error GoTo TrataErro
    Dim cARQUIVO As String
    Dim cEXT As String
    Dim cSQL As String
    Dim RsExp As ADODB.Recordset
    Dim DBEXP As ADODB.Connection
    Dim sFILTER As String
    Dim cCONN As String
    ' Variáveis do motor FSO.Stream
    Dim fso As Object
    Dim streamOut As Object
    Dim i As Long
    Dim cLINHA As String
    Dim cDelim As String
    
    ' --- NOVAS VARIÁVEIS PARA A LOGICA SELETIVA DE ASPAS ---
    Dim cValorCampo As String
    Dim bDelimSensivel As Boolean
    ' -------------------------------------------------------

    ' Validação de tabela de origem
    If Len(Trim(tabela.Text)) = 0 Then
        Alert ("Escolha Uma Tabela Antes")
        Exit Sub
    End If
    
    cSQL = CStr(sql.Text)
    
    ' Define a extensão baseada no botão clicado (Sua lógica de Index original mantida)
    Select Case Index
        Case 0: cEXT = "XML"
        Case 1, 6, 8: cEXT = "XLS"
        Case 2: cEXT = "HTML"
        Case 3: cEXT = "TXT"
        Case 4: cEXT = "PDF"
        Case 5: cEXT = "RTF"
        Case 7: cEXT = "DOC"
    End Select
    
    ' Caixa de diálogo para salvar o arquivo
    sFILTER = "Formato (*." & cEXT & ")" & vbNullChar & "*." & cEXT
    cARQUIVO = FileSave(Me, sFILTER, 1, cEXT, , App.Path, "Salvar " & cEXT & " Como")
    TxtArquivo.Text = cARQUIVO
    
    If Len(cARQUIVO) = 0 Then
        Alert ("Nome Nao Definido")
        Exit Sub
    End If
    
    ' Conexão e abertura de dados via ADODB
    cCONN = GeracArq(arquivo.Text, , False)
    Set DBEXP = New ADODB.Connection
    Set RsExp = New ADODB.Recordset
    DBEXP.ConnectionTimeout = 120
    DBEXP.Open cCONN
    RsExp.Open cSQL, DBEXP, adOpenStatic, adLockReadOnly
    
    ' Se não houver dados, aborta de forma segura
    If RsExp.EOF And RsExp.BOF Then
        Alert "A consulta nao retornou nenhum registro para exportar."
        GoTo Fim
    End If
    
    ' Inicializa o FileSystemObject de forma global para os Cases que precisam dele
    If Index <> 0 Then
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set streamOut = fso.OpenTextFile(cARQUIVO, 2, True, False)
    End If
    
    ' Processamento dos Motores de Exportação
    Select Case Index
        Case 0
            ' Exportação nativa XML do ADO (não usa FSO)
            RsExp.Save cARQUIVO, adPersistXML
            
        Case 6
            ' ==================================================================
            ' MOTOR EXCEL AVANÇADO: HTML TABULAR DISFARÇADO DE .XLS (VIA FSO)
            ' ==================================================================
            streamOut.WriteLine "<html>"
            streamOut.WriteLine "<head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8""></head>"
            streamOut.WriteLine "<style>table { border-collapse: collapse; } th { background-color: #D3D3D3; font-weight: bold; } td, th { border: 1px solid #A0A0A0; font-family: Calibri, Arial; font-size: 11pt; }</style>"
            streamOut.WriteLine "<body>"
            streamOut.WriteLine "<table>"
            
            streamOut.WriteLine " <tr>"
            For i = 0 To RsExp.Fields.Count - 1
                streamOut.WriteLine " <th>" & RsExp.Fields(i).Name & "</th>"
            Next i
            streamOut.WriteLine " </tr>"
            
            RsExp.MoveFirst
            Do While Not RsExp.EOF
                streamOut.WriteLine " <tr>"
                For i = 0 To RsExp.Fields.Count - 1
                    cLINHA = Replace(Replace(CheckNullStr(RsExp.Fields(i).Value), vbCr, ""), vbLf, " ")
                    streamOut.WriteLine " <td>" & cLINHA & "</td>"
                Next i
                streamOut.WriteLine " </tr>"
                RsExp.MoveNext
            Loop
            streamOut.WriteLine "</table>"
            streamOut.WriteLine "</body>"
            streamOut.WriteLine "</html>"
            streamOut.Close
            
        Case 1, 2, 3, 4, 5, 7
            ' ==================================================================
            ' MOTOR CSV / DELIMITADORES / OUTROS FORMATOS (VIA FSO)
            ' ==================================================================
            ' Define o delimitador de colunas baseado na interface
            If Index = 1 Then
                cDelim = Chr(9) ' Tabulação para o xls(TAB)
            ElseIf TxtDeli.Text = "<tab>" Then
                cDelim = Chr(9)
            Else
                cDelim = FixStr(TxtDeli.Text, " ") ' Usa o caracter informado na caixa de texto
            End If
            
            ' --- VERIFICA SE O DELIMITADOR É SENSÍVEL (, ou ;) ---
            If cDelim = "," Or cDelim = ";" Then
                bDelimSensivel = True
            Else
                bDelimSensivel = False
            End If
            ' -----------------------------------------------------
            
            ' 1. Escreve a linha de cabeçalho com os nomes das colunas
            cLINHA = ""
            For i = 0 To RsExp.Fields.Count - 1
                cValorCampo = RsExp.Fields(i).Name
                
                ' Protege cabeçalhos caso contenham o delimitador sensível
                If bDelimSensivel Then
                    If InStr(cValorCampo, cDelim) > 0 Or InStr(cValorCampo, """") > 0 Then
                        cValorCampo = """" & Replace(cValorCampo, """", """""") & """"
                    End If
                End If
                
                cLINHA = cLINHA & cValorCampo & cDelim
            Next i
            
            If Len(cLINHA) > 0 Then
                cLINHA = Left(cLINHA, Len(cLINHA) - Len(cDelim))
                streamOut.WriteLine cLINHA
            End If
            
            ' 2. Escreve os registros em streaming de alta velocidade
            RsExp.MoveFirst
            Do While Not RsExp.EOF
                cLINHA = ""
                For i = 0 To RsExp.Fields.Count - 1
                    ' Captura o valor puro da célula blindado contra nulos
                    cValorCampo = CheckNullStr(RsExp.Fields(i).Value)
                    
                    ' --- NOVA RAMIFICAÇÃO TRATADA POR DELIMITADOR ---
                    If bDelimSensivel Then
                        ' FLUXO A: Se for Vírgula ou Ponto e Vírgula, escapa aspas internas e envolve com aspas duplas ("...")
                        ' Remove apenas quebras de linha para manter o registro linear
                        cValorCampo = Replace(Replace(cValorCampo, vbCr, ""), vbLf, " ")
                        
                        ' Se contiver o delimitador ou aspas na string, envelopa
                        If InStr(cValorCampo, cDelim) > 0 Or InStr(cValorCampo, """") > 0 Then
                            cValorCampo = """" & Replace(cValorCampo, """", """""") & """"
                        End If
                    Else
                        ' FLUXO B: Mantém a sua lógica original idêntica para Pipe, #, ~, Tab
                        ' Substitui o delimitador por espaço caso ele exista no texto para evitar desalinhamento de colunas
                        cValorCampo = Replace(Replace(CheckNullStr(RsExp.Fields(i).Value), cDelim, " "), vbCrLf, " ")
                    End If
                    ' ------------------------------------------------
                    
                    cLINHA = cLINHA & cValorCampo & cDelim
                Next i
                
                If Len(cLINHA) > 0 Then
                    cLINHA = Left(cLINHA, Len(cLINHA) - Len(cDelim))
                    streamOut.WriteLine cLINHA
                End If
                RsExp.MoveNext
            Loop
            
            ' Fecha o fluxo de escrita antes de chamar as funções legadas de conversão
            streamOut.Close
            
            ' Tratamentos pós-escrita originais mantidos por Index
            If Index = 2 Then txttohtml cARQUIVO, cARQUIVO
            If Index = 4 Then ePASS01 = cARQUIVO: formConvertToPDF.Show
            If Index = 5 Then txttortf cARQUIVO, cARQUIVO
            If Index = 7 Then txttodoc cARQUIVO, cARQUIVO
            If Index = 8 Then txttoxls cARQUIVO, cARQUIVO, , , TxtDeli
    End Select
    
    Alert "Exportacao Concluida com Sucesso!"
    
    ' Fluxo nativo de envio de e-mail se a checkbox estiver marcada
    If CmdEmail.Value = 1 Then
        ePASS01 = Array("", "", "", "", "", cARQUIVO, "", False)
        FrmSendMail.Show vbModal, Me
    End If
    
Fim:
    On Error Resume Next
    If Not streamOut Is Nothing Then streamOut.Close
    Set streamOut = Nothing
    Set fso = Nothing
    If Not RsExp Is Nothing Then
        If RsExp.State = adStateOpen Then RsExp.Close
    End If
    If Not DBEXP Is Nothing Then
        If DBEXP.State = adStateOpen Then DBEXP.Close
    End If
    Set RsExp = Nothing
    Set DBEXP = Nothing
    Exit Sub
    
TrataErro:
    MsgBox "Erro na exportacao: " & Err.Description, vbCritical, "Erro"
    Resume Fim
End Sub

' Função auxiliar interna para blindagem contra valores nulos de banco de dados
Private Function CheckNullStr(ByVal vValue As Variant) As String
    If IsNull(vValue) Then
        CheckNullStr = ""
    Else
        CheckNullStr = Trim(CStr(vValue))
    End If
End Function

Private Sub CmdFiltro_Click()
  On Error GoTo errhandler
  Dim cFILTRO As String

  cFILTRO = FixStr(FILTRO)
  If aRELCFG(11) Then
    ePASS01 = ""
    FrmFiltro.Show vbModal, Me
    FILTRO = Replace(Replace(eRETU01, "{", ""), "}", "")
    If lRETU And Len(aRELCFG(15)) > 0 Then
      aRELCFG(15) = TrocaSqlWhere(aRELCFG(15), FixStr(FILTRO))
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
  cARQRTF = TxtArquivo.Text
  If IsExtensao(TxtArquivo.Text, "TXT") Then
      ePASS01 = Array("Preview Interno", "Imprimir Direto Impressora", "Escolher Porta(Destino)", "Editor Interno")
      escOrdem.Show vbModal, Me
      eRETU01 = FixInt(eRETU01, 0)
      Select Case eRETU01
         Case 0
           CmdVisua_Click
         Case 1
           FrmPicturePrinter.Show vbModal, Me
         Case 2
           FrmTxl.Show vbModal, Me
         Case 3
           FrmRtfView.Show vbModal, Me
        End Select
  End If
 If IsExtensao(TxtArquivo.Text, "PDF") Then
    ePASS01 = Array("Externo", "Interno")
    escOrdem.Show vbModal, Me
    eRETU01 = FixInt(eRETU01, 0)
    Select Case eRETU01
      Case 0
          CmdVisua_Click
      Case 1
         FrmPreview.Show vbModal, Me
      End Select
  End If
  If IsExtensao(TxtArquivo.Text, "RTF") Then
     ePASS01 = Array("Editor Interno", "Preview Interno")
      escOrdem.Show vbModal, Me
      eRETU01 = FixInt(eRETU01, 0)
      Select Case eRETU01
      Case 0
         FrmRtfView.Show vbModal, Me
      Case 1
         CmdVisua_Click
      End Select
  End If
  If IsExtensao(TxtArquivo.Text, "HTML") Then
    CmdVisua_Click
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
  Dim cARQSHELL As String
  cARQSHELL = FixStr(TxtArquivo)
  ShellEx cARQSHELL, essSW_SHOWDEFAULT, , , , CLng(Me.hwnd)
End Sub

Private Sub CmdVisua_Click()
  cARQRTF = TxtArquivo.Text
  If Not FileConnExist(cARQRTF, True) Then
    Exit Sub
  End If
  If IsExtensao(cARQRTF, "TXT") Then
    ePASS03 = 1
    PrintPreview1.ShowPreview
  End If
  If IsExtensao(cARQRTF, "PDF") Then
    ShellEx cARQRTF, essSW_SHOWDEFAULT, , , , CLng(Me.hwnd)
  End If
  If IsExtensao(cARQRTF, "HTML") Then
    ePASS01 = Array("Navegador Externo", "Preview Interno", "Navegador Interno")
    escOrdem.Show vbModal, Me
    eRETU01 = FixInt(eRETU01, 0)
    Select Case eRETU01
    Case 0
      OpenUrl (cARQRTF)
    Case 1
      ePASS03 = 3
      PrintPreview1.ShowPreview
    Case 3
      FrmPreview.Show vbModal, Me
    End Select
  End If
  If IsExtensao(cARQRTF, "RTF") Then
    RichTextBox1.LoadFile cARQRTF, RtfLoadSaveFormatRTF
    ePASS02 = 3
    PrintPreview1.ShowPreview
    RichTextBox1.Text = ""
  End If
End Sub
Private Sub PrintPreview1_PrepareReport(Cancel As Boolean)
  If ePASS03 = 1 Then  'TXT
    MyPrintingTXT
  End If
  If ePASS03 = 2 Then  'rtf
    MyPrintingRTF
  End If
  If ePASS03 = 3 Then  'HTML
    MyPrintinghtml
  End If
End Sub
Public Sub MyPrintinghtml()
  Dim cTEXTO As String
  Dim cLINHA As String
  Dim LINES() As String
  Dim i As Integer

  ' If Not FileConnExist(cARQRTF, True) Then 'ja checado cmdvisualclick
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
  Dim strBuffer As String
  fileFile = FreeFile
  Open cARQRTF For Input As #fileFile
  Do While Not EOF(fileFile)
    'read line
    Input #fileFile, strBuffer
    Printer.Print strBuffer
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
  Dim cFILTRO As String

  Dim nARQUIVOS As Integer
  TxtArquivo = zUSER & Format(Now, "_DDMMYYYY_HHMMSS") & ".TXT"
  CenterFormToScreen Me
  ''Configura Help
  Me.Caption = cFORMID
  HelpContextID = nFORMID
  FILTRO = ""

  Label1 = aRELCFG(6)

  ' 0-Novo,1-Abrir,2-Salvar,3-Salvacomo
  ' 4-Editar,5-Visualizar,6-Imprimir,7-Exportar

  If Not aDIREITOS(6) Then cmdimp.Visible = False
  If Not aDIREITOS(6) Then CmdConfImp.Visible = False
  If Not aDIREITOS(5) Then CmdVisua.Visible = False
  If Not aDIREITOS(7) Then CmdEmail.Visible = False
  If Not aRELCFG(11) Then CmdFiltro.Visible = False
  If Not aRELCFG(11) Then FILTRO.Visible = False
  If Not aDIREITOS(4) Then CmdEditar.Visible = False
  If Not aDIREITOS(4) Then CmdShell.Visible = False


  nARQUIVOS = UBound(aARQUIVOS)
  If nARQUIVOS > 0 Then
    If Len(aARQUIVOS(0)) > 0 Then
      If FileConnExist(aARQUIVOS(0), True, True, aRELCFG(15)) Then
        arquivo.Text = aARQUIVOS(0)
      End If
    End If
  End If
  aRELCFG(14) = FixStr(aRELCFG(14))
  If Len(aRELCFG(14)) > 0 Then
    cFILTRO = aRELCFG(14)
    FILTRO = cFILTRO
  Else
    If aRELCFG(11) Then
      CmdFiltro_Click
    End If
  End If
  If Len(aRELCFG(15)) > 0 Then
    If Len(FILTRO) > 0 Then
      aRELCFG(15) = TrocaSqlWhere(aRELCFG(15), FixStr(FILTRO))
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
  UpdateReport = False  ' we don't need to update the report in the Print preview window after this action (the default value of UpdateReport parameter is True)
End Sub


