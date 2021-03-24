VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F22668DE-E08D-467B-8E41-13900013BD5F}#2.7#0"; "VBextra2.OCX"
Begin VB.Form escIMG 
   Caption         =   "Escolha o Codigo da Imagem"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8760
   Icon            =   "escIMG.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   8760
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   5775
      Left            =   7290
      TabIndex        =   1
      Top             =   0
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   10186
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      TextAlignment   =   1
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   9128
      _Version        =   393216
   End
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

Private Sub Apaga_Click()
    Dim sql              As String
    Grid.Col = 0
    zgrp = Grid
    sql = "select * from IMAGENS WHERE CODIGO='" & zgrp & "'"
    If ApagaSQLP(cARQRTF, sql, "Apagar Codigo") Then
        FilRelat
    End If
End Sub

Private Sub CmdSair_Click()
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub Edit_Click()
    Grid.Col = 0
    zgrp = Grid
    frmIMAGENS.Show vbModal
    FilRelat
End Sub

Private Sub FilRelat()
    Dim cARQ As String
    Dim cSQL As String
    cARQ = cARQRTF
    If Len(cSUBWHERE) = 0 Then
        cSQL = "SELECT CODIGO,NUMERO FROM IMAGENS ORDER BY " & cORDEM
    Else
        cSQL = "SELECT CODIGO,NUMERO FROM IMAGENS WHERE " & cSUBWHERE & " ORDER BY " & cORDEM
    End If
    MontaGridUltra Grid, 2, Array(4000, 2000), Array("Codigo", "Numero"), _
        Array("L$CODIGO", "L$NUMERO"), cARQ, cSQL
End Sub

Private Sub Form_Load()
    Center Me
    aORDEM = Array("CODIGO", "NUMERO")
    aORDES = Array("CODIGO", "NUMERO")
    cORDEM = "CODIGO"
    cSUBWHERE = ""
    Select Case iImage
    Case 1
        cARQRTF = PegPath("PATH", "IMGMA01")
    Case 2
        cARQRTF = PegPath("PATH", "IMGMS01")
    Case 3
        cARQRTF = PegPath("PATH", "IMGFER")
    Case 4
        cARQRTF = PegPath("PATH", "IMGME01")
    Case 5
        cARQRTF = PegPath("PATH", "IMGME04")
    Case 6
        ''frmcomp.Show vbModal, Me
        ''If lRETU Then
        ''    cARQRTF = Caminex(PegPath("PATH", "IMGMP04"), zANO, zMES, zEMPRESA)
        ''End If
        ''cARQRTF = Caminex(PegPath("PATH", "IMGMP04"), zANO, zMES, zEMPRESA)
        cARQRTF = PegPath("PATH", "IMGMP04CPF")
    Case 7
        cARQRTF = PegPath("PATH", "IMGMU01")
    Case 8
        cARQRTF = PegPath("PATH", "IMGMT01")
    Case 9
        cARQRTF = PegPath("PATH", "IMGMW05")
    Case 10
        cARQRTF = PegPath("PATH", "IMGMR01")
    Case 11
        cARQRTF = PegPath("PATH", "IMGME06")
    Case 12
        cARQRTF = PegPath("PATH", "IMGMW07")
    Case 13
        cARQRTF = PegPath("PATH", "IMGMQ01")
    End Select
    xmontatoolbar Me.Toolbar1, "escIMG", True
    FilRelat
    If FixStr(eLOCALIZA) <> "" Then LocalizaGri1 Grid, eLOCALIZA, 1
End Sub

Private Sub Grid_KeyPress(KeyAscii As Integer)
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

Private Sub Novo_Click()
    Dim cCODNEW          As String
    cCODNEW = InputBox("Digite o Codigo", "Incluindo", "__")
    cCODNEW = Left(cCODNEW, 24)
    IncluiSQL cARQRTF, "SELECT * FROM IMAGENS WHERE CODIGO='" & cCODNEW & "' ", 1, Array("CODIGO"), Array(cCODNEW), True, False
    FilRelat
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Dim sButton          As String
    sButton = Button
    sButton = Left(UCase(Replace(sButton, "&", "")), 3)
    If Not AcessaBtnOld("escIMG", Button.Index) Then
        Exit Sub
    End If
    
    GravaLog 0, Button.Index, sButton, "escIMG"

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
                cSUBWHERE = MontaFiltro(aORDEM, Array("L%", "-"), eRETU01, eRETU02)
            End If
        End If
        FilRelat


    Case "IMP"
        cTIPO = "R"
        Select Case iImage
        Case 1
            zgrp = "MA"
        Case 2
            zgrp = "MS"
        Case 3
            zgrp = "FR"
        Case 4
            zgrp = "ME"
        Case 5, 11
            zgrp = "M4"
        Case 6
            zgrp = "MP"
        Case 7
            zgrp = "MU"
        Case 8
            zgrp = "MT"
        Case 9, 12
            zgrp = "CO"
        Case 10
            zgrp = "MR"
        End Select

        escRPT.Show vbModal, Me

    Case "NOV"
        Novo_Click

    Case "EDI"
        Edit_Click

    Case "EXC"
        Apaga_Click

    Case "LOC"
        ePASS01 = aORDES
        frmLocalizaa.Show vbModal, Me
        If lRETU Then
            LocalizaGrid Grid, eRETU01, eRETU02, , 1
        End If

    Case "SAI"
        CmdSair_Click

    End Select

End Sub


