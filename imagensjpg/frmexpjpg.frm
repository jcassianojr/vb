VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "xpcontrols.ocx"
Object = "{7020C36F-09FC-41FE-B822-CDE6FBB321EB}#1.0#0"; "vbccr17.ocx"
Begin VB.Form Frmexpjpg 
   Caption         =   "Exportando "
   ClientHeight    =   5955
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8370
   Icon            =   "frmexpjpg.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5955
   ScaleWidth      =   8370
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   1935
      Left            =   120
      ScaleHeight     =   1875
      ScaleWidth      =   1515
      TabIndex        =   11
      Top             =   2880
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VBCCR17.SpinBox QtdeDigitos 
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   480
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      Min             =   1
      Max             =   20
      Value           =   8
      TextAlignment   =   1
   End
   Begin VBCCR17.FrameW FrameW1 
      Height          =   2055
      Left            =   240
      TabIndex        =   12
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   3625
      Caption         =   "Exportar Usando"
      Begin VB.CheckBox SepararEmpresa 
         Caption         =   "Codigo contem empresa"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   1320
         Width           =   2295
      End
      Begin VBCCR17.CheckBoxW TiraCaracter 
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         Value           =   1
         Caption         =   "Tira Caracters Pontuacao"
      End
      Begin VBCCR17.OptionButtonW PorNumero 
         Height          =   375
         Left            =   1560
         TabIndex        =   5
         Top             =   360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Caption         =   "Numero"
      End
      Begin VBCCR17.OptionButtonW PorCodigo 
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Caption         =   "Codigo"
      End
   End
   Begin XPControls.XPButton CmdEncerrar 
      Height          =   495
      Left            =   6720
      TabIndex        =   2
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Picture         =   "frmexpjpg.frx":038A
      Caption         =   "&Sair"
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
   Begin XPControls.XPButton Cmdiniciar 
      Height          =   495
      Left            =   5040
      TabIndex        =   3
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Picture         =   "frmexpjpg.frx":0924
      Caption         =   "&Exportar"
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
   Begin XPControls.XPButton CmdImportar 
      Height          =   495
      Left            =   5040
      TabIndex        =   9
      Top             =   840
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Picture         =   "frmexpjpg.frx":0CBE
      Caption         =   "&Importar"
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
   Begin XPControls.XPButton cmdBrowseFolder 
      Height          =   675
      Left            =   6720
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2160
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1191
      Picture         =   "frmexpjpg.frx":1058
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
   Begin VB.Label Caminho 
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   2280
      Width           =   6255
   End
   Begin VB.Label ccodigo 
      Height          =   315
      Left            =   3960
      TabIndex        =   1
      Top             =   5280
      Width           =   3675
   End
   Begin VB.Label Label1 
      Caption         =   "Aguarde Exportação"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   5040
      Width           =   3375
   End
End
Attribute VB_Name = "Frmexpjpg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cCAMJPG As String

Private Sub cmdBrowseFolder_Click()
   Dim picbrowse As String
   picbrowse = BrowseFolders(Me.hWnd, "Choose a location picture", BrowseForFolders, CSIDL_DESKTOP)
    If picbrowse <> "" Then
        Caminho.Caption = picbrowse & "\"
        cCAMJPG = picbrowse & "\"
    End If
End Sub


Private Sub CmdEncerrar_Click()
    Unload Me
End Sub

Private Sub CmdImportar_Click()
    Dim OBJCONN As ADODB.Connection
    Dim OBJRSGLOB As ADODB.Recordset
    Dim strSQL As String
    Dim DBCONNSTR As String
    Dim iFileNum As Integer
    Dim lFileLength As Long
    Dim abBytes() As Byte
    Dim sTEMPFILE As String
    Dim cCAMPO As String
    Set OBJCONN = New ADODB.Connection           'Create a new object
    Set OBJRSGLOB = New ADODB.Recordset


    cCAMPO = "IMAGEM"
    DBCONNSTR = GeracArq(cARQRTF)
    OBJCONN.Open DBCONNSTR
    If InStr(cARQRTF, "PF.MDB") > 0 Then
        strSQL = "SELECT CODINT AS NUMERO,FIG01 FROM PFS where (codint is not null) and codint<>''"
        cCAMPO = "FIG01"
    Else
        strSQL = "SELECT * FROM imagens where numero>0"
    End If
    OBJRSGLOB.Open strSQL, OBJCONN, adOpenForwardOnly, adLockOptimistic

    If Not OBJRSGLOB.RecordCount = 0 Then
        Do While (Not OBJRSGLOB.EOF)
            If IsNull(OBJRSGLOB(cCAMPO)) Then
                lFileLength = 0
            Else
                lFileLength = LenB(OBJRSGLOB(cCAMPO))
            End If
            If lFileLength > 1 Then              ' Ja carregado
            Else
                If PorCodigo.Value = True Then
                    If TiraCaracter.Value = vbChecked Then
                        sTEMPFILE = cCAMJPG & Trim(TiraOut(OBJRSGLOB.Fields("codigo"))) + ".JPG"
                    Else
                        sTEMPFILE = cCAMJPG & Trim(OBJRSGLOB.Fields("codigo")) + ".JPG"
                    End If
                Else
                    sTEMPFILE = cCAMJPG & Trim(StrZero(OBJRSGLOB.Fields("NUMERO"), QtdeDigitos.Value)) + ".JPG"
                End If
                If FileExist(sTEMPFILE) Then
                    ccodigo.Caption = OBJRSGLOB.Fields("CODIGO")
                    iFileNum = FreeFile
                    Open sTEMPFILE For Binary Access Read As #iFileNum
                    lFileLength = LOF(iFileNum)
                    ReDim abBytes(lFileLength)
                    Get #iFileNum, , abBytes()
                    Close #iFileNum
                    OBJRSGLOB.Fields(cCAMPO).AppendChunk abBytes()
                    OBJRSGLOB.Update
                End If
            End If
            OBJRSGLOB.MoveNext
        Loop
    End If
    ccodigo.Caption = ""
    Alert ("Importacao Concluida")

End Sub

Private Sub Cmdiniciar_Click()
    Dim OBJCONN As ADODB.Connection
    Dim OBJRSGLOB As ADODB.Recordset
    Dim strSQL As String
    Dim CSQLI As String
    Dim DBCONNSTR As String
    Dim iFileNum As Integer
    Dim lFileLength As Long
    Dim abBytes() As Byte
    Dim sTEMPFILE As String
    Dim cCAMPO As String
    Dim cCODIGOGRV As String
    Set OBJCONN = New ADODB.Connection           'Create a new object
    Set OBJRSGLOB = New ADODB.Recordset
    cCAMPO = "IMAGEM"
    DBCONNSTR = GeracArq(cARQRTF)
    OBJCONN.Open DBCONNSTR
    If InStr(cARQRTF, "PF.MDB") > 0 Then
        strSQL = "SELECT CODINT AS NUMERO,FIG01 FROM PFS where (codint is not null) and codint<>''"
        cCAMPO = "FIG01"
    Else
        strSQL = "SELECT * FROM imagens where numero>0"
    End If
    If InStr(UCase(cARQRTF), "OL_LOGIX") > 0 Then
       'strSQL = " SELECT"
       'strSQL = strSQL & " FIRST 10 CAST(EMPRESA || TRIM(STRZERO(MATRICULA,8)) AS CHAR(10)) AS CODIGO,"
       'strSQL = strSQL & " MATRICULA AS NUMERO,LEN(FOTO) AS TAMANHO,"
       'strSQL = strSQL & " FOTO"
       'strSQL = strSQL & " From rhu_funcio_foto"
       'strSQL = strSQL & " where foto is not null "
       'cCAMPO = "FOTO"
       strSQL = " SELECT CAST(EMPRESA || TRIM(STRZERO(MATRICULA,8)) AS CHAR(10)) AS CODIGO, empresa, matricula as numero from rhu_funcio_foto"
       
    End If
    If InStr(UCase(cARQRTF), "DATAMACE") > 0 Then
        strSQL = " SELECT"
        strSQL = strSQL & " RIGHT(TAB_CADFUN.FUN_COD_EMP,2) + RIGHT(FORMAT(TAB_CADFUN.FUN_REGISTRO,'00000000'),8) AS CODIGO"
        strSQL = strSQL & " ,FUN_REGISTRO as numero"
        strSQL = strSQL & " ,fotos.BN_FOTO as FOTO "
        strSQL = strSQL & " FROM gip.fotos AS gip"
        strSQL = strSQL & " LEFT JOIN Arquivos.FOTOS  AS fotos oN gip.cd_foto=FOTOS.cd_foto"
        strSQL = strSQL & " LEFT JOIN TAB_CADFUN ON gip.CD_FUN_KEY_NUMERO = TAB_CADFUN.FUN_KEY_NUMERO"
        strSQL = strSQL & " Where TAB_CADFUN.FUN_REGISTRO Is Not Null "
        cCAMPO = "FOTO"
    End If

    
    
    OBJRSGLOB.Open strSQL, OBJCONN, adOpenForwardOnly, adLockReadOnly
    If Not OBJRSGLOB.RecordCount = 0 Then        'not Have
        Do While (Not OBJRSGLOB.EOF)
            ccodigo.Caption = OBJRSGLOB.Fields("CODIGO")
            sTEMPFILE = cCAMJPG
            If SepararEmpresa.Value = True Or SepararEmpresa.Value = 1 Then
               sTEMPFILE = sTEMPFILE & Left(ccodigo, 2) & "\"
            End If
            If PorCodigo.Value = True Then
                If TiraCaracter.Value = vbChecked Then
                    cCODIGOGRV = Trim(TiraOut(OBJRSGLOB.Fields("codigo")))
                Else
                    cCODIGOGRV = Trim(OBJRSGLOB.Fields("codigo"))
                End If
                
                If SepararEmpresa.Value = True Or SepararEmpresa.Value = 1 Then
                    cCODIGOGRV = Mid(cCODIGOGRV, 3)
                End If
            Else
                cCODIGOGRV = Trim(StrZero(OBJRSGLOB.Fields("NUMERO"), QtdeDigitos.Value))
            End If
            sTEMPFILE = sTEMPFILE & cCODIGOGRV + ".JPG"
            If InStr(UCase(cARQRTF), "OL_LOGIX") > 0 Then
              CSQLI = "SELECT FOTO AS IMAGEM FROM rhu_funcio_foto  WHERE MATRICULA=" & OBJRSGLOB.Fields("NUMERO")
              CSQLI = CSQLI & " and empresa=" & OBJRSGLOB.Fields("EMPRESA")
              If ADOPegBlob(cARQRTF, CSQLI, Picture1) Then
                  PicSave.SavePicture Picture1.Picture, sTEMPFILE, fmtJPEG, 70
                 'salvarpict Me, Picture1, sTEMPFILE
              End If
            Else
                If Not IsNull(OBJRSGLOB.Fields(cCAMPO)) Then
                    lFileLength = FixInt(LenB(OBJRSGLOB(cCAMPO)))
                    If lFileLength = 0 And InStr(UCase(cARQRTF), "OL_LOGIX") Then
                       lFileLength = FixInt(OBJRSGLOB("TAMANHO"))
                    End If
                    If lFileLength > 0 Then
                        iFileNum = FreeFile
                        Open sTEMPFILE For Binary As #iFileNum
                        abBytes = OBJRSGLOB(cCAMPO).GetChunk(lFileLength)
                        Put #iFileNum, , abBytes()
                        Close #iFileNum
                    End If
                End If
            End If
            OBJRSGLOB.MoveNext
        Loop
    End If
    Alert ("Exportaçao Concluida")
End Sub
Private Sub Form_Load()
    Center Me
    cCAMJPG = PegPath("PATH", "IMGJPG")
    Caminho.Caption = cCAMJPG
End Sub

Private Sub XPButton5_Click()

End Sub
