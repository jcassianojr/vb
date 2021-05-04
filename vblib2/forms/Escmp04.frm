VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F22668DE-E08D-467B-8E41-13900013BD5F}#2.7#0"; "VBextra2.OCX"
Begin VB.Form escMP04 
   Caption         =   "Escolha o Usuario Desejado"
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8265
   Icon            =   "Escmp04.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6090
   ScaleWidth      =   8265
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   9763
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   6090
      Left            =   6795
      TabIndex        =   1
      Top             =   0
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   10742
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      TextAlignment   =   1
      _Version        =   393216
   End
   Begin vbExtra.FlexFn FlexFn1 
      Height          =   405
      Left            =   0
      TabIndex        =   2
      Top             =   0
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
End
Attribute VB_Name = "escMP04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim aORDEM As Variant
Dim aORDES As Variant
Dim cORDEM As String
Dim cSUBWHERE As String

Private Sub CmdEscolher()
    Grid.Col = 0
    eRETU01 = Grid
    Grid.Col = 1
    eRETU02 = Grid
    lRETU = True
    Unload Me
End Sub

Private Sub cmdSair()
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub FilRelat()
    Dim cARQ As String
    Dim cSQL As String
    If ePASS01 = "LOGIX" Then
        cARQ = PegPath("PATH", "LOGIXODBC")
        cSQL = "SELECT funcionario.num_matricula as TECNICO,nom_completo as NOMTEC,dat_demis as DEMITIDO, fun_infor.num_cpf AS cpf"
        cSQL = cSQL & " From funcionario"
        cSQL = cSQL & " LEFT JOIN fun_infor ON funcionario.cod_empresa=fun_infor.cod_empresa AND funcionario.num_matricula=fun_infor.num_matricula"
        cSQL = cSQL & " WHERE funcionario.cod_empresa='01'  and dat_demis is null  AND funcionario.cod_vinculo<>80"
        If Len(cSUBWHERE) > 0 Then
           cSQL = cSQL & "  AND " & cSUBWHERE
        End If
        cSQL = cSQL & " ORDER BY " & cORDEM
    Else
        cARQ = PegPath("PATH", "CADMP04")
        cARQ = GeraConn(cARQ, "JETFOX")
        If Len(cSUBWHERE) > 0 Then
            cSQL = "SELECT TECNICO,NOMTEC,DEMITIDO,CPF FROM MP04 WHERE DAY(DEMITIDO)=0 AND " & cSUBWHERE & " ORDER BY " & cORDEM
        Else
            cSQL = "SELECT TECNICO,NOMTEC,DEMITIDO,CPF FROM MP04 WHERE DAY(DEMITIDO)=0 ORDER BY " & cORDEM
        End If
    End If
    MontaGridFast Grid, 3, Array(1200, 3500, 1), Array("Nº", "Nome", "X"), Array("TECNICO", "L$NOMTEC", "CPF"), cARQ, cSQL
End Sub

Private Sub Form_Load()
    CenterFormToScreen Me
    If ePASS01 = "LOGIX" Then
        aORDEM = Array("num_matricula", "nom_completo")
    Else
        aORDEM = Array("TECNICO", "NOMTEC")
    End If
    aORDES = Array("Nº", "Nome")
    cORDEM = "TECNICO"
    cSUBWHERE = ""
    xmontatoolbar Me.Toolbar1, "escMP04", True
    FilRelat
    If FixStr(eLOCALIZA) <> "" Then LocalizaGri1 Grid, eLOCALIZA, 1
End Sub

Private Sub Grid_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmdEscolher
    End If
    If KeyAscii > 31 And KeyAscii < 123 Then
        LocalizaGrid Grid, Chr(KeyAscii), 1, False
    End If

End Sub

'Private Sub Grid_SelChange()
'    With Grid
'        If .Rows > 2 Then
'            .Col = .Cols - 1
'            .ColSel = 0
'            .TopRow = .Row
'        End If
'    End With

'End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Dim sButton          As String
    sButton = Button
    sButton = Left(UCase(Replace(sButton, "&", "")), 3)
            
    
    If Not AcessaBtnOld("escMP04", Button.Index) Then
        Exit Sub
    End If

    GravaLog 0, Button.Index, sButton, "escMP04"
    

    Select Case sButton
    Case "ORD"
        ePASS01 = aORDES
        escOrdem.Show vbModal, Me
        If lRETU Then
            cORDEM = aORDEM(eRETU01)
            FilRelat
        End If
    Case "FIL"
        cSUBWHERE = ""
        If MDG("Usar Filtro Avancado") Then
            aARQUIVOS = Array(zMANA5EMP)
            ''Posicao 12 Nome da Tabela
            ''Posicao 13 Nome da Tabela
            aRELCFG = Array("", "", "", 0, False, _
                            False, "", "", "", "", "", _
                            False, "MP04", "MP04", "", "")
            FrmFiltro.Show vbModal, Me
            If lRETU Then
                cSUBWHERE = Replace(Replace(eRETU01, "{", ""), "}", "")
            End If
        Else
            ePASS01 = aORDES
            frmLocalizaa.Show vbModal, Me
            If lRETU Then
                cSUBWHERE = MontaFiltro(aORDEM, Array("=", "L%"), eRETU01, eRETU02)
            End If
        End If
        FilRelat
    Case "IMP"
        cTIPO = "R"
        zgrp = "MP"
        escRPT.Show vbModal, Me
    Case "FOT"
        'IMGMP04CPF
        'Grid.Col = 0
        'zgrp = Grid
        'zgrp = Trim(str(zgrp))
        'zgrp = String(8 - Len(zgrp), "0") & zgrp
        'iImage = 6
        'cARQRTF = Caminex(PegPath("PATH", "IMGMP04"), zANO, zMES, zEMPRESA)
        'Load frmIMAGENS
        'frmIMAGENS.TXTFIELDS(0).Enabled = False
        'frmIMAGENS.Escolher(0).Visible = False
        'frmIMAGENS.Show vbModal, Me
        
        Grid.Col = 2
        zgrp = Grid
        iImage = 6
        cARQRTF = PegPath("PATH", "IMGMP04CPF")
        Load frmIMAGENS
        frmIMAGENS.txtFields(0).Visible = False
        frmIMAGENS.txtFields(1).Visible = False
        frmIMAGENS.Escolher(0).Visible = False
        frmIMAGENS.Escolher(1).Visible = False
        frmIMAGENS.Escolher(2).Visible = False
        frmIMAGENS.cmdClose.Visible = False
        frmIMAGENS.IncImg.Visible = False
        frmIMAGENS.DelImg.Visible = False
        frmIMAGENS.Show vbModal, Me
    Case "ESC"
        CmdEscolher
    Case "LOC"
        ePASS01 = aORDES
        frmLocalizaa.Show vbModal, Me
        If lRETU Then
            LocalizaGrid Grid, eRETU01, eRETU02, , 1
        End If
    Case "SAI"
        cmdSair
    End Select

End Sub


