VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F22668DE-E08D-467B-8E41-13900013BD5F}#2.7#0"; "VBextra2.OCX"
Begin VB.Form escNUMNOM 
   Caption         =   "Selecione item Desejado  desejado:"
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9270
   LinkTopic       =   "Form1"
   ScaleHeight     =   6090
   ScaleWidth      =   9270
   StartUpPosition =   2  'CenterScreen
   Begin vbExtra.FlexFn FlexFn1 
      Height          =   405
      Left            =   120
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
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   9763
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   6090
      Left            =   7800
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
End
Attribute VB_Name = "escNUMNOM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim aORDEM As Variant
Dim aORDES As Variant
Dim cORDEM As String
Dim cBANCOORI As String
Dim cSUBWHERE As String

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

    Screen.MousePointer = vbDefault
    lRETU = False
    Unload Me

End Sub

Private Sub FilRelat()
    Dim cARQ As String
    Dim sSQL As String
    Select Case cBANCOORI
    Case "LOGIX"
        cARQ = PegPath("PATH", "LOGIXODBC")
        Select Case cARQESC
        Case "MA01"
            sSQL = "SELECT COD_CLIENTE AS numero,NOM_CLIENTE AS NOME,NOM_REDUZIDO AS COGNOME FROM clientes WHERE ies_situacao='A'"
        Case "MB01"
            sSQL = "SELECT cod_fornecedor AS numero,raz_social AS nome,raz_social_reduz AS cognome FROM fornecedor WHERE ies_fornec_ativo='A'"
        Case "2DIG"
            sSQL = "SELECT codigo AS numero,descricao AS nome, descricao AS cognome FROM item_2dig_clientes_970 "
        End Select
    Case "MICRO"
        cARQ = PegPath("PATH", "MICROSIGAODBC")
    Case Else
        cARQ = GeraConn(zMANA5EMP, "JETFOX")
        sSQL = "SELECT NUMERO,NOME,COGNOME FROM " & cARQESC
    End Select
    If Len(cSUBWHERE) > 0 Then
        sSQL = sSQL & " WHERE " & cSUBWHERE & " ORDER BY " & cORDEM
    Else
        sSQL = sSQL & " ORDER BY " & cORDEM
    End If
    MontaGridFast Grid, 3, Array(1600, 3000, 2000), Array("Codigo", "Razao Social", "Cognome"), _
        Array("NUMERO", "L$NOME", "L$COGNOME"), cARQ, sSQL

End Sub

Private Sub Form_Load()
    Center Me
    aORDEM = Array("NUMERO", "NOME", "COGNOME")
    aORDES = Array("Codigo", "Razao Social", "Cognome")
    cORDEM = "NUMERO"
    cSUBWHERE = ""
    xmontatoolbar Me.Toolbar1, "esc" & cARQESC, True
    cBANCOORI = GeraConn(zMANA5EMP, "JETFOX")
    Select Case ePASS01
    Case "LOGIX"
        cBANCOORI = "LOGIX"                      ''PegPath("PATH", "LOGIXODBC") abaixo filrelat
        aORDEM = Array("1", "2", "3")
    Case "MICRO"
        cBANCOORI = "MICRO"                      ''PegPath("PATH", "MICROSIGAODBC") abaixo filralt
        aORDEM = Array("1", "2", "3")
    End Select
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
    Dim cCODINC As String
    Dim cNOMEINC As String
    Dim sButton As String
    Dim cARQTMP As String
    Dim sSQL As String


    sButton = Button
    sButton = Left(UCase(Replace(sButton, "&", "")), 3)
    
    If Not AcessaBtnOld("esc" & cARQESC, Button.Index) Then
        Exit Sub
    End If
    
    GravaLog 0, Button.Index, sButton, "escNUMNOM"

    Select Case sButton
    
    
    Case "EXC", "DEL"
        If cARQESC = "ME01" Then
            cARQTMP = PegPath("PATH", "MANA5")
            cARQTMP = GeraConn(cARQTMP, "SDECDX")
            Grid.Col = 0
            sSQL = "select numero from ME01 WHERE numero='" & Grid & "'"
            If ApagaSQLP(cARQTMP, sSQL) Then
                FilRelat
            End If
        End If
    Case "NOV", "INC"
        If cARQESC = "ME01" Then
            cCODINC = InputBox("Digite o Codigo", "Inclusão ", "__")
            cNOMEINC = InputBox("Digite a Descricao", "Descricao do codigo", "__")
            cARQTMP = PegPath("PATH", "MANA5")
            cARQTMP = GeraConn(cARQTMP, "SDECDX")
            sSQL = "select numero,nome from ME01 WHERE numero='" & cCODINC & "'"
            If IncluiSQL(cARQTMP, sSQL, 2, Array("CODIGO", "NOME"), Array(cCODINC, cNOMEINC), True, True) Then
                ePASS01 = cCODINC
                ePASS02 = cNOMEINC
                frmME01.Show vbModal, Me
            End If
        End If
    
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
                            False, cARQESC, cARQESC, "", "")
            FrmFiltro.Show vbModal, Me
            If lRETU Then
                cSUBWHERE = Replace(Replace(eRETU01, "{", ""), "}", "")
            End If
        Else
            ''frmLocalizar.Show vbModal, Me
            ePASS01 = aORDES
            frmLocalizaa.Show vbModal, Me
            If lRETU Then
                If cARQESC = "MC02" Then
                    cSUBWHERE = MontaFiltro(aORDEM, Array("L%", "L%", "L%"), eRETU01, eRETU02)
                Else
                    cSUBWHERE = MontaFiltro(aORDEM, Array("=", "L%", "L%"), eRETU01, eRETU02)
                End If
            End If
        End If
        FilRelat


    Case "IMP"
        cTIPO = "R"

        Select Case cARQESC

        Case "MA01"
            zgrp = "MA"

        Case "MB01"
            zgrp = "MB"

        Case "ME01"
            zgrp = "ME"

        Case "MC02"
            zgrp = "MC"
        Case Else
            Exit Sub
        End Select

        escRPT.Show vbModal, Me

    Case "EDI"

        Select Case cARQESC

        Case "ME01"
            Grid.Col = 0
            ePASS01 = Grid
            Grid.Col = 1
            ePASS02 = Grid
            frmME01.Show vbModal, Me

        End Select

    Case "FOT"
        Grid.Col = 0
        zgrp = Grid

        Select Case cARQESC

        Case "MA01"
            zgrp = Trim(str(zgrp))
            zgrp = String(8 - Len(zgrp), "0") & zgrp
            iImage = 1
            cARQRTF = PegPath("PATH", "LOGOMA01") & "IMGMA01.MDB"

        Case "ME01"
            iImage = 4
            cARQRTF = PegPath("PATH", "LOGOME01") & "IMGME01.MDB"

        End Select

        If cARQESC <> "MC02" Then

            Load frmIMAGENS
            frmIMAGENS.txtFields(0).Enabled = False
            frmIMAGENS.Escolher(0).Visible = False
            frmIMAGENS.Show vbModal, Me

        End If

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

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault
End Sub


