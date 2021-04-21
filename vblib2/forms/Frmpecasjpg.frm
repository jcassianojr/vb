VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Begin VB.Form frmIMAGENS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "IMAGENS"
   ClientHeight    =   6435
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   6960
   HelpContextID   =   1002
   Icon            =   "Frmpecasjpg.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XPControls.XPButton ximgsave 
      Height          =   435
      Left            =   5160
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2520
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "Frmpecasjpg.frx":038A
      Caption         =   "Salvar Imagem"
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
   Begin XPControls.XPButton DelImg 
      Height          =   435
      Left            =   5160
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1920
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "Frmpecasjpg.frx":0824
      Caption         =   "Excluir Imagem"
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
   Begin XPControls.XPButton IncImg 
      Height          =   435
      Left            =   5160
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1320
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "Frmpecasjpg.frx":0CBE
      Caption         =   "Incluir Imagem"
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
   Begin XPControls.XPButton VerImg 
      Height          =   495
      Left            =   5160
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   5280
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Picture         =   "Frmpecasjpg.frx":1158
      Caption         =   "Navegar Imagens"
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
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      Height          =   3735
      Left            =   120
      ScaleHeight     =   3675
      ScaleWidth      =   4155
      TabIndex        =   5
      Top             =   960
      Width           =   4215
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   1335
      Left            =   120
      ScaleHeight     =   1275
      ScaleWidth      =   4155
      TabIndex        =   4
      Top             =   4800
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      Index           =   1
      Left            =   2880
      MaxLength       =   24
      TabIndex        =   2
      ToolTipText     =   "Numero da Imagem"
      Top             =   600
      Width           =   2175
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      Index           =   0
      Left            =   120
      MaxLength       =   24
      TabIndex        =   1
      ToolTipText     =   "Codigo da Imagem"
      Top             =   600
      Width           =   2655
   End
   Begin XPControls.XPButton CmdConfImp 
      Height          =   435
      Left            =   5160
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   4680
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "Frmpecasjpg.frx":15F2
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
   Begin XPControls.XPButton Encerrar 
      Height          =   435
      Left            =   5160
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "Frmpecasjpg.frx":1B8C
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
   Begin XPControls.XPButton cmdClose 
      Height          =   435
      Left            =   5160
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "Frmpecasjpg.frx":2126
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
   Begin XPControls.XPButton CmdPaste 
      Height          =   435
      Left            =   5160
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "Frmpecasjpg.frx":26C0
      Caption         =   "Copia"
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
   Begin XPControls.XPButton cmdimprimir 
      Height          =   435
      Left            =   5160
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "Frmpecasjpg.frx":2C5A
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
   Begin XPControls.XPButton Escolher 
      Height          =   375
      Index           =   0
      Left            =   840
      TabIndex        =   11
      Top             =   120
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Picture         =   "Frmpecasjpg.frx":31F4
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
   End
   Begin XPControls.XPButton Escolher 
      Height          =   375
      Index           =   1
      Left            =   1200
      TabIndex        =   12
      Top             =   120
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      Picture         =   "Frmpecasjpg.frx":378E
      Caption         =   "LX"
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
   Begin XPControls.XPButton Escolher 
      Height          =   375
      Index           =   2
      Left            =   1920
      TabIndex        =   13
      Top             =   120
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      Picture         =   "Frmpecasjpg.frx":3D28
      Caption         =   "MC"
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
      Height          =   435
      Left            =   5160
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   4080
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "Frmpecasjpg.frx":42C2
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
   Begin VB.Label lblLabels 
      Caption         =   "Numero"
      Height          =   255
      Index           =   2
      Left            =   2880
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblLabels 
      Caption         =   "Codigo"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmIMAGENS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const nFORMID = 1109
Const cFORMID = "frmimagens"

'Option Explicit
Dim cARQ As String
Dim cSQL As String
Dim CSQLI As String
Dim aVAL As Variant
Dim aFOR As Variant
Dim aCAM As Variant
Dim aPAD As Variant
Dim nCAMPOS As Integer
Dim iLOOP As Integer
Dim lABRE  As Boolean
Dim lTROCOU As Boolean
Dim cBASEDADOS As String

Private Sub CmdAbrirCom_Click()
salvarpict Me, Picture1, txtFields(0)
   If FileExist(txtFields(0), True) Then
       'Call OpenWith(cARQRTF, OAIF_ALLOW_REGISTRATION Or OAIF_EXEC Or OAIF_FORCE_REGISTRATION, Me.hWnd)
   End If
End Sub

Private Sub cmdClose_Click()
    On Error Resume Next
    If cBASEDADOS = "LOGIX" Or cBASEDADOS = "DATAMACE" Then  'InStr(UCase(cARQ), "OL_LOGIX") > 0 Then
        lABRE = False
    End If
    If lABRE Then
        If MDG("Gravar Alteracoes") Then
            txtFields(1) = FixNum(txtFields(1))
            If txtFields(1) = 0 Then txtFields(1) = Val(TiraOut(txtFields(0)))
            For iLOOP = 0 To nCAMPOS - 1
                aVAL(iLOOP) = txtFields(iLOOP)
            Next iLOOP
            GrvSQL cARQ, cSQL, nCAMPOS, aCAM, aVAL, aFOR
            If lTROCOU Then
                CSQLI = "select * from imagens  WHERE CODIGO='" & ZGRP & "'"
                ADOGrvBlob cARQ, CSQLI, Picture1
            End If
        End If
    End If
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub CmdConfImp_Click()
'picbrowse = BrowseFolders(Me.hWnd, "Choose a location picture", BrowseForPrinters, CSIDL_PRINTERS)
    FrmPrintSetup.Show vbModal, Me
End Sub

Private Sub CmdImprimir_Click()
    If Picture1.Height = 0 Then
        Alert ("Sem Imagem")
    Else
        PrinterEx.ShowPrintPreview Me, "MyPrintingRoutine"
    End If
End Sub

Private Sub CmdPaste_Click()
    If Picture1.Height = 0 Then
        Alert ("Sem Imagem")
    Else
        Clipboard.Clear
        Clipboard.SetData Picture1.Image, (vbCFBitmap)
    End If
End Sub

Private Sub DelImg_Click()
    Set Picture1.Picture = Nothing
    Set Picture2.Picture = Nothing
    lTROCOU = True
End Sub

Private Sub Encerrar_Click()
    If cBASEDADOS = "LOGIX" Or cBASEDADOS = "DATAMACE" Then   'InStr(UCase(cARQ), "OL_LOGIX") = 0 Then
        If MDG("Sair Sem Gravar") Then
            Screen.MousePointer = vbDefault
            Unload Me
        End If
    Else
        Screen.MousePointer = vbDefault
        Unload Me
    End If
End Sub

Private Sub Escolher_Click(Index As Integer)
    Dim cARQESC As String
    ePASS01 = ""
    If Index = 1 Then
        ePASS01 = "LOGIX"
    End If
    If Index = 2 Then
        ePASS01 = "MICRO"
    End If
    Select Case iImage

    Case 1                                       'Clientes
        cARQESC = "MA01"
        escNUMNOM.Show vbModal, Me

        If lRETU Then

            frmIMAGENS.txtFields(0) = CStr(eRETU01)
            frmIMAGENS.txtFields(1) = eRETU01

        End If

    Case 2                                       'Produtos
        ePASS01 = "MANA5"
        escms01.Show vbModal, Me
        If lRETU Then frmIMAGENS.txtFields(0) = eRETU01

    Case 3                                       'Ferramental
        escFER.Show vbModal, Me

        If lRETU Then frmIMAGENS.txtFields(0) = eRETU01

    Case 4                                       'Maquinas
        cARQESC = "ME01"
        escNUMNOM.Show vbModal, Me

        If lRETU Then frmIMAGENS.txtFields(0) = eRETU01

    Case 5                                       'Instrumentos
        escME04.Show vbModal, Me

        If lRETU Then frmIMAGENS.txtFields(0) = eRETU01

    Case 6                                       'Funcionarios

        escMP04.Show vbModal, Me

        If lRETU Then

            frmIMAGENS.txtFields(0) = CStr(eRETU01)
            frmIMAGENS.txtFields(1) = eRETU01

        End If

    Case 7                                       'Materia Prima mu01
        iMU01 = 1
        escmu01.Show vbModal, Me

        If lRETU Then frmIMAGENS.txtFields(0) = eRETU01

    Case 8                                       'cOMPONENTES mt01
        iMU01 = 2
        escmu01.Show vbModal, Me

        If lRETU Then frmIMAGENS.txtFields(0) = eRETU01

    Case 9                                       'cONSUMIVEIS mw05
        iMU01 = 3
        escmu01.Show vbModal, Me

        If lRETU Then frmIMAGENS.txtFields(0) = eRETU01

    Case 10                                      'embalagens mr01
        iMU01 = 4
        escmu01.Show vbModal, Me

        If lRETU Then frmIMAGENS.txtFields(0) = eRETU01

    Case 12                                      'iTENS MANUTENCAO
        iMU01 = 6
        escmu01.Show vbModal, Me

        If lRETU Then frmIMAGENS.txtFields(0) = eRETU01
    Case 13                                      'Sub Produtos
        iMU01 = 7
        escmu01.Show vbModal, Me

        If lRETU Then frmIMAGENS.txtFields(0) = eRETU01
            

    End Select

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    TeclaEnter KeyCode
End Sub

Private Sub Form_Load()
    Dim nTMPNUMERO
    Dim nPOS
    
    Center Me
    Me.Caption = cFORMID
    HelpContextID = nFORMID
    
    lTROCOU = False
    If iImage = 2 Then
        Escolher(1).Visible = True
    Else
        Escolher(1).Visible = False
    End If
    If iImage = 0 Then
        Escolher(0).Visible = False
        Escolher(1).Visible = False
        Escolher(2).Visible = False
    End If
    If iImage = 5 Then
        Escolher(2).Visible = False
    End If

    lABRE = True
    
    If InStr(UCase(cARQRTF), "OL_LOGIX") > 0 Then
        cBASEDADOS = "LOGIX"
    End If
    
    If InStr(UCase(cARQRTF), "DATAMACE") > 0 Then
        cBASEDADOS = "DATAMACE"
    End If
    

    ZGRP = FixStr(ZGRP, "")
    If Len(ZGRP) = 0 Then
        Alert ("Codigo em Branco")
        lABRE = False
    End If

    If lABRE Then
        cARQ = cARQRTF
        nPOS = InStr(ZGRP, "|")
        If nPOS > 0 Then
            nTMPNUMERO = Mid(ZGRP, nPOS + 1)
            ZGRP = Mid(ZGRP, 1, nPOS - 1)
        Else
            nTMPNUMERO = ZGRP
        End If
        nTMPNUMERO = funNumeroPuro(nTMPNUMERO)
        nTMPNUMERO = Val(nTMPNUMERO)


        cSQL = "select codigo,numero from IMAGENS WHERE CODIGO='" & ZGRP & "'"
        If cBASEDADOS = "LOGIX" Then 'InStr(UCase(cARQ), "OL_LOGIX") > 0 Then
            cSQL = "SELECT STRZERO(MATRICULA,8) AS CODIGO,MATRICULA AS NUMERO FROM rhu_funcio_foto  WHERE MATRICULA='" & ZGRP & "'"
            cSQL = cSQL & " and empresa=" & StrZero(zEMPRESA, 2)
        End If
        
        If cBASEDADOS = "DATAMACE" Then
            cSQL = " SELECT"
            cSQL = cSQL & " RIGHT(TAB_CADFUN.FUN_COD_EMP,2) + RIGHT(FORMAT(TAB_CADFUN.FUN_REGISTRO,'00000000'),8) AS CODIGO "
            cSQL = cSQL & " ,FUN_REGISTRO AS NUMERO"
            cSQL = cSQL & " FROM gip.fotos AS gip"
            cSQL = cSQL & " LEFT JOIN Arquivos.FOTOS  AS fotos oN gip.cd_foto=FOTOS.cd_foto"
            cSQL = cSQL & " LEFT JOIN TAB_CADFUN ON gip.CD_FUN_KEY_NUMERO = TAB_CADFUN.FUN_KEY_NUMERO"
            cSQL = cSQL & " WHERE TAB_CADFUN.FUN_COD_EMP='0" & Left(ZGRP, 2) & "' AND TAB_CADFUN.FUN_REGISTRO=" & Right(ZGRP, 8)
        End If

        
                
        If cBASEDADOS = "LOGIX" Or cBASEDADOS = "DATAMACE" Then  'InStr(UCase(cARQ), "OL_LOGIX") > 0 Then
        Else
            IncluiSQL cARQ, cSQL, 2, Array("CODIGO", "NUMERO"), Array(ZGRP, nTMPNUMERO), True, False
        End If
        
        
        nCAMPOS = 2
        aCAM = Array("CODIGO", "NUMERO")
        aFOR = Array("C", "NI")
        aPAD = Array("", 0)
        aVAL = PegSQL(cARQ, cSQL, nCAMPOS, aCAM, aFOR, aPAD)
        For iLOOP = 0 To nCAMPOS - 1
            txtFields(iLOOP) = aVAL(iLOOP)
        Next iLOOP
        If iImage = 2 And nPOS > 0 Then
            txtFields(1) = nTMPNUMERO
        End If
        CSQLI = "select IMAGEM from imagens  WHERE CODIGO='" & ZGRP & "'"
        If cBASEDADOS = "LOGIX" Then 'InStr(UCase(cARQ), "OL_LOGIX") > 0 Then
            CSQLI = "SELECT FOTO AS IMAGEM FROM rhu_funcio_foto  WHERE MATRICULA=" & ZGRP
            CSQLI = CSQLI & " and empresa=" & StrZero(zEMPRESA, 2)
        End If
        If cBASEDADOS = "DATAMACE" Then
            CSQLI = " SELECT"
            CSQLI = CSQLI & " fotos.BN_FOTO AS IMAGEM "
            CSQLI = CSQLI & " FROM gip.fotos AS gip"
            CSQLI = CSQLI & " LEFT JOIN Arquivos.FOTOS  AS fotos oN gip.cd_foto=FOTOS.cd_foto"
            CSQLI = CSQLI & " LEFT JOIN TAB_CADFUN ON gip.CD_FUN_KEY_NUMERO = TAB_CADFUN.FUN_KEY_NUMERO"
            CSQLI = CSQLI & " WHERE TAB_CADFUN.FUN_COD_EMP='0" & Left(ZGRP, 2) & "' AND TAB_CADFUN.FUN_REGISTRO=" & Right(ZGRP, 8)
        End If
        
        
        If ADOPegBlob(cARQ, CSQLI, Picture1) Then
            StretchSourcePictureFromPicture Picture1, Picture2
            If FixNum(eRETU01) > 500000 Then
                Alert ("Imagem Muito Grande,Ajuste o tamanho")
                salvarpict Me, Picture1, "Imagem_" & ZGRP
                Set Picture1.Picture = Nothing
                Set Picture2.Picture = Nothing
                lTROCOU = True
            End If
        Else
            Set Picture1.Picture = Nothing
            Set Picture2.Picture = Nothing
            If FileExist(zRPTCAM, False) Then
                If FixInt(FileLen(STMPFILE)) > 500000 Then
                    Alert ("Imagem Muito Grande,Ajuste o tamanho " & zRPTCAM)
                Else
                    Picture1.Picture = LoadPicture(zRPTCAM)
                    StretchSourcePictureFromPicture Picture1, Picture2
                    lTROCOU = True
                End If
            End If
        End If
    Else
        ''
    End If
End Sub

Private Sub IncImg_Click()
    Dim STMPFILE
    STMPFILE = OpenArqExt(Me, "", "JPG", "JPEG *.JPG")
    If lerarquivoimagem(STMPFILE, Picture1, Picture2) Then
        lTROCOU = True
    End If
End Sub

Private Sub Picture1_Click()
    salvarpict Me, Picture1
End Sub

Private Sub Picture2_Click()
    IncImg_Click
End Sub

Private Sub Picture2_DblClick()
    IncImg_Click
End Sub

Private Sub VerImg_Click()
    frmPicViewer.Show vbModal, Me
    If lRETU Then
        If lerarquivoimagem(eRETU01, Picture1, Picture2) Then
            lTROCOU = True
        End If
    End If
End Sub

Public Sub MyPrintingRoutine()
    On Error Resume Next
    Printer.Print
    Printer.PaintPicture Picture1, 0, 0
    Printer.EndDoc
End Sub

Public Property Get Printer() As Printer
    Set Printer = vbExtra.Printer2
End Property

Public Property Set Printer(nPrinter As Printer)
    Set vbExtra.Printer2 = nPrinter
End Property

Private Sub ximgsave_Click()
    If Picture1.Height = 0 Then
        Alert ("Sem Imagem")
    Else
        salvarpict Me, Picture1, txtFields(0)
    End If
End Sub

