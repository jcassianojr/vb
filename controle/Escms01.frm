VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F22668DE-E08D-467B-8E41-13900013BD5F}#2.7#0"; "VBextra2.OCX"
Begin VB.Form escms01 
   Caption         =   "Selecione o Item Desejado"
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9375
   Icon            =   "Escms01.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6090
   ScaleWidth      =   9375
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   9763
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   6090
      Left            =   7905
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
      Left            =   240
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
Attribute VB_Name = "escms01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim aORDEM As Variant
Dim aORDES As Variant
Dim cORDEM As String
Dim cSUBWHERE As String
Dim cARQ As String
Dim cSQLBAS As String
Dim cARQORI As String

Private Sub CmdEscolher()
    Grid.Col = 0
    eRETU01 = Grid
    Grid.Col = 1
    eRETU02 = Grid
    Grid.Col = 2
    eRETU03 = Grid
    lRETU = True
    Unload Me
End Sub

Private Sub cmdSair()
    lRETU = False
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub FilRelat()
    Dim cSQL As String
    If cARQORI = "MANA5" Then
        cSQLBAS = "SELECT CODIGO,NOME,CODIGOINT FROM MS01 "
        cARQ = GeraConn(zMANA5EMP, "JETFOX")
    End If
    If cARQORI = "LOGIX" Then                    'item
        cARQ = PegPath("PATH", "LOGIXODBC")
        cSQLBAS = "SELECT DISTINCT item.cod_item as codigoint, item.den_item as nome ,cliente_item.cod_item_cliente as codigo"
        cSQLBAS = cSQLBAS & " From item INNER JOIN cliente_item ON item.cod_empresa = cliente_item.cod_empresa AND item.cod_item = cliente_item.cod_item "
        cSQLBAS = cSQLBAS & " where  item.cod_empresa='01' and len(item.cod_item)=12 and item.cod_item<>cliente_item.cod_item_cliente "
        cSQLBAS = cSQLBAS & " ORDER BY item.cod_item "
        cORDEM = ""
    End If
    If cARQORI = "MICRO" Then                    'sb1010
        cARQ = PegPath("PATH", "MICROSIGAODBC")
        cSQLBAS = "SELECT b1_cod as codigoint, b1_desc as nome , b1_cod as codigo"
        cSQLBAS = cSQLBAS & " From SB1010"
        cSQLBAS = cSQLBAS & " ORDER BY b1_cod "
        cORDEM = ""
    End If
    If cARQORI = "LOGIE" Then                    'cliente item
        cARQ = PegPath("PATH", "LOGIXODBC")
        cSQLBAS = " SELECT DISTINCT cliente_item.cod_item AS codigoint,item.den_item as nome , cliente_item.cod_item_cliente as codigo "
        cSQLBAS = cSQLBAS & " From cliente_item LEFT JOIN item ON cliente_item.cod_empresa=item.cod_empresa AND cliente_Item.cod_item=item.cod_item "
        cSQLBAS = cSQLBAS & "  WHERE cliente_item.cod_item[1,3]='EMB' "
        cSQLBAS = cSQLBAS & " ORDER BY cliente_Item.cod_item "
        cORDEM = ""
    End If
    If cARQORI = "LOGIM" Then                    'materia prima
        cARQ = PegPath("PATH", "LOGIXODBC")
        cSQLBAS = " SELECT DISTINCT item.cod_item AS codigoint, item.den_item AS nome, item.cod_item AS codigo"
        cSQLBAS = cSQLBAS & " From Item "
        cSQLBAS = cSQLBAS & " WHERE  ( cod_item[1,2]='BB' OR cod_item[1,2]='RL' OR cod_item[1,2]='CH' OR cod_item[1,3]='BBP' OR cod_item[1,3]='RLP'"
        cSQLBAS = cSQLBAS & "           OR cod_item[1,2]='TR'OR cod_item[1,3]='CHP' )"
        cSQLBAS = cSQLBAS & " ORDER BY item.cod_item "
        cORDEM = ""
    End If
    If cARQORI = "LOGIC" Then                    'componente
        cARQ = PegPath("PATH", "LOGIXODBC")
        cSQLBAS = " SELECT DISTINCT item.cod_item AS codigoint, item.den_item AS nome, item.cod_item AS codigo"
        cSQLBAS = cSQLBAS & " From Item "
        cSQLBAS = cSQLBAS & " WHERE   LEN(COD_iTEM)<> 12 AND LEN(COD_iTEM)<>14 AND "
        cSQLBAS = cSQLBAS & " cod_item[1,2]<>'BB' AND cod_item[1,2]<> 'RL' AND cod_item[1,2]<>'CH' AND cod_item[1,3]<>'BBP' "
        cSQLBAS = cSQLBAS & "         AND  cod_item[1,3]<>'RLP' AND  cod_item[1,2]<>'TR' AND cod_item[1,3]<>'CHP' "
        cSQLBAS = cSQLBAS & "  AND  cod_item[1,2]<>'77' AND  cod_item[1,2]<>'88' AND  cod_item[1,2]<>'99'"
        cSQLBAS = cSQLBAS & " ORDER BY item.cod_item "
        cORDEM = ""
    End If
    If cARQORI = "LOGIS" Then                    'subproduto
        cARQ = PegPath("PATH", "LOGIXODBC")
        cSQLBAS = " SELECT DISTINCT item.cod_item as codigoint, item.den_item as nome ,ciclo_peca_970.cod_item_cliente as codigo "
        cSQLBAS = cSQLBAS & " From Item "
        cSQLBAS = cSQLBAS & " INNER JOIN ciclo_peca_970 ON item.cod_empresa = ciclo_peca_970.cod_empresa "
        cSQLBAS = cSQLBAS & "      AND item.cod_item = ciclo_peca_970.cod_item "
        cSQLBAS = cSQLBAS & "  WHERE  item.cod_empresa='01' AND len(item.cod_item)=14 "
        cSQLBAS = cSQLBAS & "  and item.cod_item<>ciclo_peca_970.cod_item_cliente "
        cSQLBAS = cSQLBAS & " ORDER BY item.cod_item "
        cORDEM = ""
    End If
    If cARQORI = "LXEMB" Then                    'logix embalagem
        cARQ = PegPath("PATH", "LOGIXODBC")
        cSQLBAS = " SELECT embalagem.cod_embal as codigoint,embalagem.den_embal as nome ,de_para_embal.cod_embal_item  as codigo"
        cSQLBAS = cSQLBAS & " FROM embalagem  "
        cSQLBAS = cSQLBAS & " LEFT JOIN de_para_embal ON de_para_embal.cod_empresa='01' AND embalagem.cod_embal=de_para_embal.cod_embal_vdp "
        cSQLBAS = cSQLBAS & "  where de_para_embal.cod_embal_item IS NOT null "
        cSQLBAS = cSQLBAS & " ORDER BY de_para_embal.cod_embal_item "
        cORDEM = ""
    End If
    
    If cARQORI = "LXINSP" Then                   'logix embalagem
        cARQ = PegPath("PATH", "LOGIXODBC")
        cSQLBAS = " SELECT DISTINCT item.cod_item AS codigoint, item.den_item AS nome, item.cod_item AS codigo"
        cSQLBAS = cSQLBAS & " From Item "
        cSQLBAS = cSQLBAS & " WHERE   cod_empresa='01' "
        cSQLBAS = cSQLBAS & " AND ies_ctr_estoque='S' AND ies_tem_inspecao='S' AND ies_situacao='A' "
        cSQLBAS = cSQLBAS & " ORDER BY item.cod_item "
        cORDEM = ""
    End If
    
    
    
    cSQL = cSQLBAS
    If Len(cSUBWHERE) > 0 Then
        cSQL = cSQL & " WHERE " & cSUBWHERE
    End If
    If Len(cORDEM) > 0 Then
        cSQL = cSQL & " ORDER BY " & cORDEM
    End If
    
    MontaGridFast Grid, 3, Array(1400, 4000, 1400), Array("Codigo", "Nome", "Cod.Int"), _
        Array("L$CODIGO", "NOME", "L$CODIGOINT"), cARQ, cSQL
    If FixStr(eLOCALIZA) <> "" Then LocalizaGri1 Grid, eLOCALIZA, 1
End Sub

Private Sub Form_Load()
    CenterFormToScreen Me
    'Montatoolbar trata os tipos  LOGIX LOGIC LOGIE ... recebido epass01 para carqori
    cARQORI = ePASS01
    aORDEM = Array("CODIGO", "NOME", "CODIGOINT")
    aORDES = Array("CODIGO", "NOME", "CODIGOINT")
    cORDEM = "CODIGO"
    cSUBWHERE = ""
    If cARQORI = "LOGIX" Then
    End If
    If cARQORI = "MICRO" Then
        cSUBWHERE = " d_e_l_e_t_<>'*'"
        aORDEM = Array("B1_CODIGO", "B1_DESC", "B1_CODIGO")
    End If
    xmontatoolbar Me.Toolbar1, "escMS01", True
    lRETU = False
    eRETU01 = ""
    eRETU02 = ""
    eRETU03 = ""
    FilRelat
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
    Dim sButton As String
    sButton = Button
    sButton = Left(UCase(Replace(sButton, "&", "")), 3)
    If Not AcessaBtnOld("escMS01", Button.Index) Then
        Exit Sub
    End If

    
    GravaLog 0, Button.Index, sButton, "escMS01"

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
                            False, "MS01", "MS01", "", "")
            FrmFiltro.Show vbModal, Me
            If lRETU Then
                cSUBWHERE = Replace(Replace(eRETU01, "{", ""), "}", "")
            End If
        Else
            ePASS01 = aORDES
            frmLocalizaa.Show vbModal, Me
            If lRETU Then
                cSUBWHERE = MontaFiltro(aORDEM, Array("L%", "L%"), eRETU01, eRETU02)
            End If
        End If
        FilRelat

    Case "ENS"
        Grid.Col = 0
        nPF = 0
        nSEQ = 0
        nSSQ = 0
        Ccodcomp = Grid
        nARQPCS = 2
        FRMPCS.Show vbModal
        'Alert ("Ensaios")
    Case "IMP"
        cTIPO = "R"
        zgrp = "MS"
        escRPT.Show vbModal, Me

    Case "FOT"
        Grid.Col = 0
        zgrp = Grid
        iImage = 2
        cARQRTF = PegPath("PATH", "LOGOMS01") & "IMAGEM.MDB"
        Load frmIMAGENS
        frmIMAGENS.TXTFIELDS(0).Enabled = False
        frmIMAGENS.Escolher(0).Visible = False
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
        lRETU = False
        cmdSair

    End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault
End Sub


