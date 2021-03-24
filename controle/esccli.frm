VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F22668DE-E08D-467B-8E41-13900013BD5F}#2.7#0"; "VBextra2.OCX"
Begin VB.Form escCLI 
   Caption         =   "Selecione o Cliente Desejado"
   ClientHeight    =   6390
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   8880
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   6390
      Left            =   7410
      TabIndex        =   1
      Top             =   0
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   11271
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      TextAlignment   =   1
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   10186
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
Attribute VB_Name = "escCLI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cARQDES As String
Dim aORDEM As Variant
Dim aORDES As Variant
Dim cORDEM As String
Dim cSUBWHERE As String

Private Sub Apagar_Click()
    Dim sSQL As String
    Grid.Col = 0
    nPF = FixInt(Grid)
    'Cliente
    sSQL = "select * from cliente WHERE cliente=" & nPF
    If ApagaSQLP(cARQDES, sSQL) Then
        'Produtos
        sSQL = "select * from produto WHERE cliente=" & nPF
        ApagaSQL cARQDES, sSQL
        FilRelat
    End If
End Sub

Private Sub CmdSair_Click()
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub Editar_Click()
    
    Grid.Col = 0
    nPF = FixInt(Grid)
    
    Grid.Col = 1
    cCLINOME = Grid
    
    If ZENGTIP = "DES" Then
        frmCLI.Show vbModal
    End If
    
    If ZENGTIP = "PRO" Then
        frmprot.Show vbModal
    End If
    
    
    
End Sub

Private Sub FilRelat()
    Dim cSQL As String
    If Len(cSUBWHERE) > 0 Then
        cSQL = "SELECT CLIENTE,CLINOME FROM CLIENTE WHERE " & cSUBWHERE & " ORDER BY " & cORDEM
    Else
        cSQL = "SELECT CLIENTE,CLINOME FROM CLIENTE ORDER BY " & cORDEM
    End If
    MontaGridFast Grid, 2, Array(1200, 4000), Array("Cliente", "Nome"), _
        Array("Cliente", "L$Clinome"), cARQDES, cSQL

End Sub

Private Sub Form_Load()
    Center Me
    aORDEM = Array("CLIENTE", "CLINOME")
    aORDES = Array("Nº CLIENTE", "NOME")
    cORDEM = "CLIENTE"
    cSUBWHERE = ""
    cARQDES = PegPath("PATH", "DESENHO")
    xmontatoolbar Me.Toolbar1, "escCLI", True
    FilRelat
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

Private Sub incluir_Click()
    Dim cNOME As String
    cNOME = InputBox("Digite o Nome", "Inclusão Cliente", " ")
    cNOME = Left$(cNOME, 50)
    nPF = FixInt(PegMAXSQL(cARQDES, "cliente", "cliente", 0)) + 1
    If IncluiSQL(cARQDES, "SELECT * FROM CLieNTE WHERE CLIENTE=" & nPF, 2, Array("CLIENTE", "CLINOME"), Array(nPF, cNOME), True, False) Then
        FilRelat
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Dim sButton As String
    sButton = Button
    sButton = Left(UCase(Replace(sButton, "&", "")), 3)
    
    If Not AcessaBtnOld("escCLI", Button.Index) Then
        Exit Sub
    End If
    
    GravaLog 0, Button.Index, sButton, "escCLI"
    

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
            aARQUIVOS = Array(cARQDES)
            ''Posicao 12 Nome da Tabela
            ''Posicao 13 Nome da Tabela
            aRELCFG = Array("", "", "", 0, False, _
                            False, "", "", "", "", "", _
                            False, "CLIENTE", "CLIENTE", "", "")
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
        zgrp = "EN"
        escRPT.Show vbModal, Me

    Case "NOV"
        incluir_Click

    Case "EDI"
        Editar_Click

    Case "EXC"
        Apagar_Click

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

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault

End Sub


