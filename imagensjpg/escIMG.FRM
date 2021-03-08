VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "xpcontrols.ocx"
Object = "{F22668DE-E08D-467B-8E41-13900013BD5F}#2.7#0"; "VBextra2.OCX"
Begin VB.Form escIMG 
   Caption         =   "Escolha o Codigo da Imagem"
   ClientHeight    =   7230
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9945
   HelpContextID   =   1001
   Icon            =   "escIMG.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7230
   ScaleWidth      =   9945
   StartUpPosition =   3  'Windows Default
   Begin vbExtra.FlexFn FlexFn1 
      Height          =   405
      Left            =   8040
      TabIndex        =   11
      Top             =   120
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
   Begin VB.CommandButton Cmdexpall 
      Caption         =   "Exportar Todas Imagens jpg"
      Height          =   615
      Left            =   8040
      TabIndex        =   1
      Top             =   4680
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Listas de Imagens"
      Top             =   120
      Width           =   7755
      _ExtentX        =   13679
      _ExtentY        =   10186
      _Version        =   393216
   End
   Begin XPControls.XPText Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   6000
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   661
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
      Enabled         =   0   'False
   End
   Begin XPControls.XPButton cmdsair 
      Height          =   435
      Left            =   8040
      TabIndex        =   3
      Top             =   4080
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "escIMG.frx":058A
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
   Begin XPControls.XPButton cmdlocalizar 
      Height          =   435
      Left            =   8040
      TabIndex        =   4
      Top             =   2280
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "escIMG.frx":0B24
      Caption         =   "Localizar"
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
      Left            =   8040
      TabIndex        =   6
      Top             =   3480
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "escIMG.frx":10BE
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
   Begin XPControls.XPButton CmdOrdem 
      Height          =   435
      Left            =   8040
      TabIndex        =   7
      Top             =   2880
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "escIMG.frx":1658
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
   Begin XPControls.XPButton Command1 
      Height          =   435
      Left            =   7920
      TabIndex        =   8
      Top             =   6000
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "escIMG.frx":1BF2
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
   Begin XPControls.XPButton novo 
      Height          =   435
      Left            =   8040
      TabIndex        =   9
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "escIMG.frx":218C
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
   Begin XPControls.XPButton cmdEditar 
      Height          =   435
      Left            =   8040
      TabIndex        =   10
      Top             =   1080
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "escIMG.frx":2726
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
   Begin XPControls.XPButton apaga 
      Height          =   435
      Left            =   8040
      TabIndex        =   5
      Top             =   1560
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "escIMG.frx":2CC0
      Caption         =   "Apagar"
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
Attribute VB_Name = "escIMG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim aORDEM As Variant
Dim aORDES As Variant
Dim cORDEM As String
Dim cSUBWHERE As String
Dim cBASEDADOS As String

Private Sub Apaga_Click()
    Dim eRESP As Variant
    Dim SQL As String
    Dim ZGRP
    Grid.Col = 0
    ZGRP = Grid
    SQL = "select * from IMAGENS WHERE CODIGO='" & ZGRP & "'"
    If ApagaSQLP(cARQRTF, SQL) Then
        FilRelat
    End If
End Sub

Private Sub Cmdexpall_Click()
    Frmexpjpg.Show vbModal, Me
End Sub

Private Sub CmdFILTRO_Click()
    cSUBWHERE = ""
    If MDG("Usar Filtro Avancado") Then
        aARQUIVOS = Array(cARQRTF)
        ''Posicao 12 Nome da Tabela
        ''Posicao 13 Nome da Tabela
        aRELCFG = Array("", "", "", 0, False, _
                        False, "", "", "", "", "", _
                        False, "IMAGENS", "IMAGENS", "", "")
        FrmFiltro.Show vbModal, Me
        If lRETU Then
            cSUBWHERE = Replace(Replace(eRETU01, "{", ""), "}", "")
            FilRelat
        End If
    Else
        ePASS01 = aORDES
        frmLocalizaa.Show vbModal, Me
        If lRETU Then
            cSUBWHERE = MontaFiltro(aORDEM, Array("L%", "="), eRETU01, eRETU02)
            FilRelat
        End If
    End If

End Sub

Private Sub cmdlocalizar_Click()
    ePASS01 = aORDES
    frmLocalizaa.Show vbModal, Me
    If lRETU Then
        LocalizaGrid Grid, eRETU01, eRETU02
    End If
End Sub

Private Sub CmdOrdem_Click()
    ePASS01 = aORDES
    escOrdem.Show vbModal, Me
    If lRETU Then
        cORDEM = aORDEM(eRETU01)
        FilRelat
    End If
End Sub

Private Sub CmdSair_Click()
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub cmdEditar_Click()
    Grid.Col = 0
    ZGRP = Grid
    Load frmIMAGENS
    frmIMAGENS.txtFields(0).Enabled = False
    frmIMAGENS.txtFields(0).TabStop = False
    frmIMAGENS.txtFields(0).Locked = True
    frmIMAGENS.txtFields(1).Enabled = False
    frmIMAGENS.txtFields(1).TabStop = False
    frmIMAGENS.txtFields(1).Locked = True
    frmIMAGENS.Show vbModal
    FilRelat
    LocalizaGrid Grid, ZGRP
End Sub

Private Sub Command1_Click()
    Text1.Text = OpenArqExt(Me, "", "MDB", "Access *.MDB")
    cARQRTF = Text1.Text
End Sub

Private Sub Grid_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdEditar_Click
    Else
        LocalizaGrid Grid, Chr(KeyAscii), 1, False
    End If
End Sub

Private Sub Form_Load()
    Dim nPOS As Long
    Dim cARQHELP As String
    Dim SSQL As String
    
    Center Me
    
    aORDEM = Array("CODIGO", "NUMERO")
    aORDES = Array("CODIGO", "NUMERO")
    cORDEM = "CODIGO"
    
    'cARQRTF = UCase(Trim(Command())) 'usuario e senha e conn sao podem ser convertida para maisculas
    cARQRTF = Trim(Command())
    cARQRTF = Replace(cARQRTF, UCase(App.EXEName) & ".EXE", "")
    cARQRTF = Trim(cARQRTF)
    ZGRP = ""
    zRPTCAM = ""
    
  
    cARQHELP = PegPath("PATH", "HELP")
    If FileExist(cARQHELP) Then
       App.HelpFile = cARQHELP
    End If

    If Len(cARQRTF) = 0 Then
        cARQRTF = PegPath("PATH", "IMGARQ")
    End If
    
'    Alert cARQRTF
    
        nPOS = InStr(cARQRTF, "#")
        If nPOS > 0 Then
            ZGRP = Mid(cARQRTF, nPOS + 1)
            cARQRTF = Mid(cARQRTF, 1, nPOS - 1)
        End If
        nPOS = InStr(ZGRP, "$")
        If nPOS > 0 Then
            zRPTCAM = Mid(ZGRP, nPOS + 1)
            ZGRP = Mid(ZGRP, 1, nPOS - 1)
        End If
       
        nPOS = InStr(cARQRTF, "%")
        If nPOS > 0 Then
            zEMPRESA = Mid(cARQRTF, 1, nPOS - 1)
            cARQRTF = Mid(cARQRTF, nPOS + 1)
        End If
       
        If Len(cARQRTF) = 0 Then                 'Passado Somente Codigo
            cARQRTF = PegPath("PATH", "IMGARQ")
        End If
    
    cBASEDADOS = ""
    If InStr(UCase(cARQRTF), "OL_LOGIX") > 0 Then
        cBASEDADOS = "LOGIX"
    End If
    
    If InStr(UCase(cARQRTF), "DATAMACE") > 0 Then
        cBASEDADOS = "DATAMACE"
    End If
    
    
'    Alert cARQRTF
    'Alert cBASEDADOS
    
    cARQRTF = PegCamini(cARQRTF)
    
'    Alert cARQRTF
    
    If cBASEDADOS = "LOGIX" Or cBASEDADOS = "DATAMACE" Then
    '
    'nao busca mdb se for logix ou datamace
    Else
         If Not FileExist(cARQRTF, True) Then
'            Alert "comando1"
            Command1_Click
         End If
         If Not FileExist(cARQRTF, True) Then
            End
         End If
    End If
    
    Text1.Text = cARQRTF
    
    If Len(ZGRP) > 0 Then
    
        SSQL = "select codigo,numero from IMAGENS WHERE CODIGO='" & ZGRP & "'"
        IncluiSQL cARQRTF, SSQL, 2, Array("CODIGO", "NUMERO"), Array(ZGRP, Val(TiraOut(ZGRP))), True, False
       
 
      
    
        Load frmIMAGENS
        
        frmIMAGENS.txtFields(0).Enabled = False
        frmIMAGENS.txtFields(0).TabStop = False
        frmIMAGENS.txtFields(0).Locked = True
        frmIMAGENS.txtFields(1).Enabled = False
        frmIMAGENS.txtFields(1).TabStop = False
        frmIMAGENS.txtFields(1).Locked = True
        frmIMAGENS.Show vbModal
        End
    End If
    Grid.SelectionMode = flexSelectionByRow
    FilRelat
    
End Sub

Private Sub FilRelat()
    Dim cARQ As String
    Dim cSQL As String
    cARQ = cARQRTF
    cSQL = "SELECT CODIGO,NUMERO FROM IMAGENS "
    If InStr(cARQ, "PF.MDB") > 0 And InStr(cARQ, "CPF.MDB") = 0 Then
        cSQL = "SELECT PF AS CODIGO,CODINT AS NUMERO FROM PFS "
    End If
    If cBASEDADOS = "LOGIX" Then
        cSQL = "SELECT STRZERO(MATRICULA,8) AS CODIGO,MATRICULA AS NUMERO FROM rhu_funcio_foto"
        cSQL = cSQL & " where empresa=" & StrZero(zEMPRESA, 2)
    End If
    If cBASEDADOS = "DATAMACE" Then
        cSQL = " SELECT "
        cSQL = cSQL & " RIGHT(TAB_CADFUN.FUN_COD_EMP,2) + RIGHT(FORMAT(TAB_CADFUN.FUN_REGISTRO,'00000000'),8) AS CODIGO "
        cSQL = cSQL & " ,FUN_REGISTRO AS NUMERO "
        cSQL = cSQL & " FROM gip.fotos AS gip "
        cSQL = cSQL & " LEFT JOIN Arquivos.FOTOS  AS fotos oN gip.cd_foto=FOTOS.cd_foto "
        cSQL = cSQL & " LEFT JOIN TAB_CADFUN ON gip.CD_FUN_KEY_NUMERO = TAB_CADFUN.FUN_KEY_NUMERO "
    End If
    If Len(cSUBWHERE) > 0 Then
        cSQL = cSQL & " WHERE " & cSUBWHERE
    End If
    If InStr(cARQ, "PF.MDB") > 0 Then
    Else
        cSQL = cSQL & " ORDER BY " & cORDEM
    End If
    MontaGridUltra Grid, 2, Array(4000, 1500), Array("Codigo", "Numero"), _
        Array("L$CODIGO", "L$NUMERO"), cARQ, cSQL
End Sub

Private Sub Novo_Click()
    Dim cCODNEW As String
    Dim SSQL As String
    
    
    cCODNEW = Busca("Digite o Codigo", "Incluindo", "_", 24)
    
    SSQL = "select codigo,numero from IMAGENS WHERE CODIGO='" & cCODNEW & "'"
    
    If InStr(cARQRTF, "OL_LOGIX") > 0 Then
    Else
        IncluiSQL cARQRTF, SSQL, 2, Array("CODIGO", "NUMERO"), Array(cCODNEW, Val(TiraOut(cCODNEW))), True
    End If
    FilRelat
End Sub


