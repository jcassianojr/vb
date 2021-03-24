VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F22668DE-E08D-467B-8E41-13900013BD5F}#2.7#0"; "VBextra2.OCX"
Begin VB.Form escME04 
   Caption         =   "Selecione o Item  desejado:"
   ClientHeight    =   6090
   ClientLeft      =   2850
   ClientTop       =   1605
   ClientWidth     =   9105
   LinkTopic       =   "Form1"
   ScaleHeight     =   6090
   ScaleWidth      =   9105
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   6090
      Left            =   7635
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
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   9763
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
Attribute VB_Name = "escME04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim aORDEM As Variant
Dim aORDES As Variant
Dim cORDEM As String
Dim cSUBWHERE As String
Dim cARQME As String

Private Sub CmdEscolher()
    Grid.Col = 6
    If Grid <> "A" Then
        Alert ("Instrumento nao Ativo")
        Exit Sub
    End If
    Grid.Col = 0
    eRETU01 = Grid
    Grid.Col = 2
    eRETU02 = Grid
    Grid.Col = 3
    eRETU03 = Grid
    lRETU = True
    Unload Me
End Sub

Private Sub cmdSair()
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub FilRelat()
    Dim cSQL As String
    If Len(cSUBWHERE) = 0 Then
        cSQL = "SELECT CODIGO,TIPO,CODTIPO,NOMTIPO,APLICACAO,PF,SITUACAO FROM ME04 ORDER BY " & cORDEM
    Else
        cSQL = "SELECT CODIGO,TIPO,CODTIPO,NOMTIPO,APLICACAO,PF,SITUACAO FROM ME04 " & cSUBWHERE & " ORDER BY " & cORDEM
    End If
    MontaGridFast Grid, 7, Array(1200, 2000, 800, 1, 2000, 400, 400), Array("Código", "Descrição", "Tipo", "_", "Aplicação", "PF", "Sit"), _
        Array("L$CODIGO", "L$TIPO", "L$CODTIPO", "L$NOMTIPO", "L$APLICACAO", "PF", "C$SITUACAO"), cARQME, cSQL
End Sub

Private Sub Form_Load()
    Center Me
    cARQME = PegPath("PATH", "MANA5INS")
    cARQME = GeraConn(cARQME, "JETFOX")
    aORDEM = Array("CODIGO", "TIPO", "CODTIPO", "APLICACAO", "PF")
    aORDES = Array("Código", "Descrição", "Tipo", "Aplicação", "PF")
    cORDEM = "CODIGO"
    cSUBWHERE = ePASS01
    FilRelat
    xmontatoolbar Me.Toolbar1, "escME04", True
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
    'Dim CLINHA As String
    Dim cARQTMP As String
    Dim sSQL As String
    On Error Resume Next
    sButton = Button
    sButton = Left(UCase(Replace(sButton, "&", "")), 3)
    If Not AcessaBtnOld("escME04", Button.Index) Then
        Exit Sub
    End If
    
    GravaLog 0, Button.Index, sButton, "escME04"
    
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
            aARQUIVOS = Array(cARQME)
            aRELCFG = Array("", "", "", 0, False, _
                            False, "", "", "", "", "", _
                            False, "ME04", "ME04", "", "")
            FrmFiltro.Show vbModal, Me
            If lRETU Then
                cSUBWHERE = Replace(Replace(eRETU01, "{", ""), "}", "")
            End If
        Else
            ePASS01 = aORDES
            frmLocalizaa.Show vbModal, Me
            If lRETU Then
                cSUBWHERE = MontaFiltro(aORDEM, Array("L%", "L%", "L%", "L%"), eRETU01, eRETU02)
            End If
        End If
        FilRelat
    Case "EXC", "DEL"
        cARQTMP = PegPath("PATH", "MANA5INS")
        cARQTMP = GeraConn(cARQTMP, "SDECDX")
        Grid.Col = 0
        sSQL = "select codigo from ME04 WHERE CODIGO='" & Grid & "'"
        If ApagaSQLP(cARQTMP, sSQL) Then
            FilRelat
        End If
    Case "NOV", "INC"
        cCONJUNTO = InputBox("Digite o Codigo", "Inclusão Instr/Calib", "__")
        cARQTMP = PegPath("PATH", "MANA5INS")
        cARQTMP = GeraConn(cARQTMP, "SDECDX")
        sSQL = "select codigo from ME04 WHERE CODIGO='" & cCONJUNTO & "'"
        If IncluiSQL(cARQTMP, sSQL, 1, Array("CODIGO"), Array(cCONJUNTO), True, True) Then
            ePASS01 = cCONJUNTO
            frmMe04.Show vbModal, Me
        End If
    Case "EDI"
        Grid.Col = 0
        ePASS01 = Grid
        frmMe04.Show vbModal, Me
    Case "IMP"
        cTIPO = "R"
        zgrp = "M4"
        escRPT.Show vbModal, Me
    Case "FOT"
        Grid.Col = 0
        zgrp = Grid
        iImage = 5
        cARQRTF = PegPath("PATH", "IMGME04")
        Load frmIMAGENS
        frmIMAGENS.txtFields(0).Enabled = False
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
        cmdSair
            
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault
End Sub


