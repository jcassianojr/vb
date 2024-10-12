VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F22668DE-E08D-467B-8E41-13900013BD5F}#2.7#0"; "VBextra2.OCX"
Object = "{2DA70529-3366-414A-B408-46083BCD481B}#1.8#0"; "VBFLXGRD18.OCX"
Begin VB.Form escfemea 
   Caption         =   "Selecione o Fmea  desejado:"
   ClientHeight    =   5775
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11535
   Icon            =   "Escfemea.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   11535
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   5775
      Left            =   9945
      TabIndex        =   1
      Top             =   0
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   10186
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      TextAlignment   =   1
      _Version        =   393216
   End
   Begin VBFLXGRD18.VBFlexGrid Grid 
      Height          =   5295
      Left            =   60
      TabIndex        =   0
      Top             =   420
      Width           =   9795
      _ExtentX        =   17277
      _ExtentY        =   9340
   End
End
Attribute VB_Name = "escfemea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim aORDEM As Variant
Dim aORDES As Variant
Dim cORDEM As String
Dim cSUBWHERE As String

Private Sub cdmrevperiodica()
  Dim dDATA As Date
  Dim lNOVA As Boolean
  Dim DB As New ADODB.Connection
  Dim RSTAB As New ADODB.Recordset
  Dim RSTA2 As New ADODB.Recordset
  Dim RSTA3 As New ADODB.Recordset

  If Not MDG("Revisoes Periodicas") Then
    Exit Sub
  End If

  DB.ConnectionTimeout = 120

  DB.Open GeracArq(Sdb)
  '*    sSQL = "SELECT RESNUM,bloqueado,PF,STPFFE,CODIGO,DEscr,resdat,FEMEAD,FEMEAPRO,FEMEAREVD,FEMEAED FROM PF ORDER BY PF"
  RSTAB.Open "PF", DB, adOpenDynamic, adLockOptimistic
  RSTA2.Open "REV", DB, adOpenDynamic, adLockOptimistic
  RSTA3.Open "REVI", DB, adOpenDynamic, adLockOptimistic
  If Not RSTAB.EOF Then RSTAB.MoveFirst
  While Not RSTAB.EOF
    If (Not IsNull(RSTAB!RESNUM)) And (Not RSTAB!bloqueado) And RSTAB!RESNUM > 0 Then
      lNOVA = False
      dDATA = Today()
      If Not IsDate(RSTAB("FEMEAPRO")) Then
        'RSTAB.Edit
        If IsDate(RSTAB("FEMEAREVD")) Then
          RSTAB("FEMEAPRO") = RSTAB("FEMEAREVD") + 730
        Else
          If IsDate(RSTAB("FEMEAED")) Then
            RSTAB("FEMEAPRO") = RSTAB("FEMEAED") + 730
          End If
        End If
        RSTAB.Update
      End If
      If IsDate(RSTAB("FEMEAREVD")) Then
        dDATA = RSTAB("FEMEAREVD") + 730
        If dDATA < Today() Then
          lNOVA = True
        End If
      Else
        If IsDate(RSTAB("FEMEAED")) Then
          dDATA = RSTAB("fEMEAED") + 730
          If dDATA < Today() Then
            lNOVA = True
          End If
        End If
      End If
      If lNOVA Then
        nREV = 1
        If IsNumeric(RSTAB("FEMEAREV")) Then
          nREV = RSTAB("FEMEAREV") + 1
        End If
        RSTA2.AddNew
        RSTA2!PF = RSTAB("PF")
        RSTA2!REVISAO = nREV
        RSTA2!TIPO = "FE"
        RSTA2!Data = dDATA
        RSTA2.Update

        RSTA3.AddNew
        RSTA3!PF = RSTAB("PF")
        RSTA3!REVISAO = nREV
        RSTA3!TIPO = "FE"
        RSTA3!item = 1
        RSTA3!motivo = "Revisão Periodica"
        RSTA3.Update


        'RSTAB.Edit
        RSTAB("FEMEAREV") = nREV
        RSTAB("FEMEAREVD") = dDATA
        RSTAB.Update


      End If
    End If
    RSTAB.MoveNext
  Wend

  RSTAB.Close
  RSTA2.Close
  RSTA3.Close

End Sub

Private Sub CmdLiberar_Click()
  Dim nTMPELA
  Grid.Col = 6
  If IsDate(Grid) Then
    If Not MDG("Ja Liberada - Liberar Novamente", "Confirme Liberacao") Then
      Exit Sub
    End If
  End If
  Grid.Col = 0
  nPF = Grid

  'nTMPELA = PegUltSQL(Sdb, "SELECT FEMEAEF FROM PF WHERE PF=" & nPF, "FEMEAEF", 0)
  nTMPELA = PegCampoSQLADO(Sdb, "SELECT FEMEAEF FROM PF WHERE PF=" & nPF, "", 0)
  If nTMPELA = zIDFOLHA Then
    Alert ("Voce ja e o Elaborador")
    Exit Sub
  End If

  ePASS01 = "FEMEA"
  frmPFLIB.Show vbModal, Me
  FilRelat
End Sub

Private Sub CmdSair_Click()

  Screen.MousePointer = vbDefault
  Unload Me

End Sub

Private Sub EditPf_Click()

  Grid.Col = 0
  nPF = Grid
  frmfemea.Show vbModal

End Sub

Private Sub FilRelat()
  Dim cDIZ As Variant
  Dim dFEMEA As Date
  Dim DB As New ADODB.Connection
  Dim RSTAB As New ADODB.Recordset
  Dim sSQL As String


  DB.ConnectionTimeout = 120

  DB.Open GeracArq(Sdb)
  If Len(cSUBWHERE) > 0 Then
    sSQL = "SELECT CODIGOINT,FEMMUDPAD,FEMEAF,RESNUM,bloqueado,PF,STPFFE,CODIGO,DEscr,resdat,FEMEAD,FEMEAPRO,FEMEAREVD,FEMEAED FROM PF " & cSUBWHERE & " ORDER BY " & cORDEM
  Else
    sSQL = "SELECT CODIGOINT,FEMMUDPAD,FEMEAF,RESNUM,bloqueado,PF,STPFFE,CODIGO,DEscr,resdat,FEMEAD,FEMEAPRO,FEMEAREVD,FEMEAED FROM PF ORDER BY " & cORDEM
  End If
  RSTAB.Open sSQL, DB, adOpenDynamic, adLockOptimistic

  With Grid

    .cols = 8
    .FixedCols = 0
    .Rows = 1
    .Row = 0
    .Col = 0
    .tEXT = "PF"
    .Col = 1
    .tEXT = "PF"
    .Col = 2
    .tEXT = "PAD"
    .Col = 3
    .tEXT = "Produto"
    .Col = 4
    .tEXT = "Descrição"
    .Col = 5
    .tEXT = "Lib.PF"
    .Col = 6
    .tEXT = "Lib.Femea"
    .Col = 7
    .tEXT = "Codigo Int"
    .ColWidth(0) = 600
    .ColWidth(1) = 300
    .ColWidth(2) = 300
    .ColWidth(3) = 1600
    .ColWidth(4) = 3000
    .ColWidth(5) = 1000
    .ColWidth(6) = 1000
    .ColWidth(7) = 1600
    .ColAlignment(1) = FlexAlignmentCenterCenter
    .ColAlignment(2) = FlexAlignmentCenterCenter
    .ColAlignment(3) = FlexAlignmentLeftCenter
    .ColAlignment(4) = FlexAlignmentLeftCenter
    .ColAlignment(7) = FlexAlignmentLeftCenter

    If Not RSTAB.EOF Then RSTAB.MoveFirst

    While Not RSTAB.EOF

      If (Not IsNull(RSTAB!RESNUM)) And (Not RSTAB!bloqueado) And RSTAB!RESNUM > 0 Then


        If RSTAB("FEMEAF") > 0 And IsDate(RSTAB("FEMEAD")) Then
          dFEMEA = RSTAB("FEMEAD")
          cDIZ = Fdata(dFEMEA)
        Else
          cDIZ = Space(8)
        End If


        .AddItem RSTAB!PF & vbTab & IIf(RSTAB("STPFFE"), "*", " ") & vbTab & _
                 IIf(RSTAB("FEMMUDPAD"), "*", " ") & vbTab & _
                 RSTAB!CODIGO & vbTab & RSTAB!DEscr & vbTab & RSTAB!resdat & _
                 vbTab & cDIZ & vbTab & RSTAB("CODIGOINT")


        If Not IsDate(RSTAB("FEMEAPRO")) Then
          'RSTAB.Edit
          If IsDate(RSTAB("FEMEAREVD")) Then
            RSTAB("FEMEAPRO") = RSTAB("FEMEAREVD") + 730
          Else
            If IsDate(RSTAB("FEMEAED")) Then
              RSTAB("FEMEAPRO") = RSTAB("FEMEAED") + 730
            End If
          End If
          RSTAB.Update
        End If


      End If

      RSTAB.MoveNext

    Wend

  End With

  RSTAB.Close

End Sub

Private Sub Form_Load()
  CenterFormToScreen Me
  aORDEM = Array("PF", "STPFFE", "FEMUDPAD", "CODIGO", "DESCR", "RESDAT", "FEMEAD", "CODIGOINT")
  aORDES = Array("Nº PF", "Alterado PF", "Alterado Femea Padrao", "Codigo", "Descricao", "Data Liberacao Processo", "Data Liberacao Femea", "Codigo INT")
  cORDEM = "PF"
  cSUBWHERE = ""
  xmontatoolbar Me.Toolbar1, "escFEMEA", True
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

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
  Dim sButton As String
  sButton = Button
  sButton = Left(UCase(Replace(sButton, "&", "")), 3)

  If Not AcessaBtnOld("escFEMEA", Button.Index) Then
    Exit Sub
  End If

  GravaLog 0, Button.Index, sButton, "escFEMEA"


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
      aARQUIVOS = Array(Sdb)
      ''Posicao 12 Nome da Tabela
      ''Posicao 13 Nome da Tabela
      aRELCFG = Array("", "", "", 0, False, _
                      False, "", "", "", "", "", _
                      False, "PF", "PF", "", "")
      FrmFiltro.Show vbModal, Me
      If lRETU Then
        cSUBWHERE = Replace(Replace(eRETU01, "{", ""), "}", "")
      End If
    Else
      ePASS01 = aORDES
      frmLocalizaa.Show vbModal, Me
      If lRETU Then
        cSUBWHERE = MontaFiltro(aORDEM, Array("=", "B", "B", "L*", "L*", "L*", "L*"), eRETU01, eRETU02)
      End If
    End If
    FilRelat
  Case "APU"
    Alert ("Opcao No Modulo Femea Padrao")
    'escFemeaApu.Show vbModal, Me
  Case "IMP"
    cTIPO = "R"
    zgrp = "FEME"
    escRPT.Show vbModal, Me
  Case "EDI"
    EditPf_Click
  Case "LIB"
    CmdLiberar_Click
  Case "REV"
    cdmrevperiodica
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


