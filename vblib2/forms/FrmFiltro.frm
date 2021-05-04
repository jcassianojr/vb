VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Begin VB.Form FrmFiltro 
   Caption         =   "Filtro"
   ClientHeight    =   5895
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9360
   Icon            =   "FrmFiltro.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5895
   ScaleWidth      =   9360
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton oper 
      Caption         =   ">="
      Height          =   375
      Index           =   0
      Left            =   1080
      TabIndex        =   28
      Top             =   2880
      Width           =   495
   End
   Begin VB.CommandButton oper 
      Caption         =   "LIKE"
      Height          =   375
      Index           =   6
      Left            =   3960
      TabIndex        =   27
      Top             =   2880
      Width           =   495
   End
   Begin VB.TextBox Alias 
      Enabled         =   0   'False
      Height          =   300
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   25
      TabStop         =   0   'False
      ToolTipText     =   "Tabela"
      Top             =   540
      Width           =   1455
   End
   Begin VB.CommandButton TIPOFIL 
      Caption         =   "OU"
      Height          =   285
      Index           =   2
      Left            =   5760
      TabIndex        =   24
      Top             =   2940
      Width           =   420
   End
   Begin VB.CommandButton TIPOFIL 
      Caption         =   "E"
      Height          =   285
      Index           =   3
      Left            =   5280
      TabIndex        =   23
      Top             =   2940
      Width           =   420
   End
   Begin VB.TextBox cTIPO 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   22
      TabStop         =   0   'False
      Text            =   "E"
      Top             =   2940
      Width           =   420
   End
   Begin VB.CommandButton oper 
      Caption         =   "<>"
      Height          =   375
      Index           =   5
      Left            =   3480
      TabIndex        =   21
      Top             =   2880
      Width           =   495
   End
   Begin VB.CommandButton oper 
      Caption         =   "="
      Height          =   375
      Index           =   4
      Left            =   3000
      TabIndex        =   20
      Top             =   2880
      Width           =   495
   End
   Begin VB.CommandButton oper 
      Caption         =   "<"
      Height          =   375
      Index           =   3
      Left            =   2520
      TabIndex        =   19
      Top             =   2880
      Width           =   495
   End
   Begin VB.CommandButton oper 
      Caption         =   ">"
      Height          =   375
      Index           =   2
      Left            =   2040
      TabIndex        =   18
      Top             =   2880
      Width           =   495
   End
   Begin VB.CommandButton oper 
      Caption         =   "<="
      Height          =   375
      Index           =   1
      Left            =   1560
      TabIndex        =   17
      Top             =   2880
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Caption         =   "Operaçao"
      Height          =   645
      Left            =   120
      TabIndex        =   15
      Top             =   2700
      Width           =   4515
      Begin VB.TextBox coper 
         Enabled         =   0   'False
         Height          =   285
         Left            =   90
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Text            =   ">="
         Top             =   225
         Width           =   585
      End
   End
   Begin VB.CheckBox Vazio 
      Caption         =   "Vazio/Zerado"
      Height          =   285
      Left            =   6360
      TabIndex        =   14
      Top             =   2700
      Width           =   1545
   End
   Begin VB.Frame Frame 
      Caption         =   "Tipo"
      Height          =   645
      Left            =   4740
      TabIndex        =   13
      Top             =   2700
      Width           =   1545
   End
   Begin VB.TextBox camtip 
      Enabled         =   0   'False
      Height          =   300
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "Tipo do Campo Escolhido"
      Top             =   3840
      Width           =   1455
   End
   Begin VB.TextBox campo 
      Enabled         =   0   'False
      Height          =   300
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "Campo Escolhido"
      Top             =   3480
      Width           =   1455
   End
   Begin VB.TextBox ValFim 
      Height          =   300
      Left            =   495
      TabIndex        =   8
      ToolTipText     =   "Valor Final"
      Top             =   2340
      Width           =   6960
   End
   Begin VB.TextBox ValINI 
      Height          =   300
      Left            =   495
      TabIndex        =   7
      ToolTipText     =   "Valor Inicial do Campo"
      Top             =   1980
      Width           =   6960
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   2400
      Left            =   45
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Lista de Campos (Clique Para Escolher)"
      Top             =   3420
      Width           =   6630
      _ExtentX        =   11695
      _ExtentY        =   4233
      _Version        =   393216
   End
   Begin VB.TextBox FILTRO 
      Enabled         =   0   'False
      Height          =   870
      Left            =   135
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Filtro"
      Top             =   960
      Width           =   7380
   End
   Begin VB.TextBox TIPO 
      Enabled         =   0   'False
      Height          =   300
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Tipo de Arquivo"
      Top             =   600
      Width           =   1470
   End
   Begin VB.TextBox tabela 
      Enabled         =   0   'False
      Height          =   300
      Left            =   780
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Tabela"
      Top             =   585
      Width           =   1875
   End
   Begin VB.TextBox arquivo 
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Base de Dados"
      Top             =   120
      Width           =   7440
   End
   Begin XPControls.XPButton CmdAbrir 
      Height          =   375
      Left            =   6840
      TabIndex        =   30
      Top             =   4200
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Picture         =   "FrmFiltro.frx":058A
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
   Begin XPControls.XPButton Encerrar 
      Height          =   435
      Left            =   7680
      TabIndex        =   31
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   767
      Picture         =   "FrmFiltro.frx":0B24
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
   Begin XPControls.XPButton ok 
      Height          =   435
      Left            =   7560
      TabIndex        =   32
      Top             =   1920
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "FrmFiltro.frx":10BE
      Caption         =   "Confirmar"
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
   Begin XPControls.XPButton CmdZerar 
      Height          =   435
      Left            =   7680
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   960
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "FrmFiltro.frx":1658
      Caption         =   "Apagar Filtro"
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
   Begin VB.Label Label6 
      Caption         =   "Lista Use (;) Ex: SP;RJ 0;20"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   6360
      TabIndex        =   29
      Top             =   3060
      Width           =   2115
   End
   Begin VB.Label Label5 
      Caption         =   "Alias"
      ForeColor       =   &H00C00000&
      Height          =   270
      Left            =   4920
      TabIndex        =   26
      Top             =   585
      Width           =   420
   End
   Begin VB.Label Label4 
      Caption         =   "Ao:"
      ForeColor       =   &H00C00000&
      Height          =   330
      Left            =   135
      TabIndex        =   12
      Top             =   2340
      Width           =   510
   End
   Begin VB.Label Label3 
      Caption         =   "Do:"
      ForeColor       =   &H00C00000&
      Height          =   330
      Left            =   135
      TabIndex        =   11
      Top             =   1980
      Width           =   510
   End
   Begin VB.Label Label2 
      Caption         =   "TIPO"
      ForeColor       =   &H00C00000&
      Height          =   270
      Left            =   2760
      TabIndex        =   3
      Top             =   585
      Width           =   510
   End
   Begin VB.Label Label1 
      Caption         =   "Tabela"
      ForeColor       =   &H00C00000&
      Height          =   270
      Left            =   180
      TabIndex        =   2
      Top             =   600
      Width           =   660
   End
End
Attribute VB_Name = "FrmFiltro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const nFORMID = 1067
Const cFORMID = "Filtro Para Relatórios"
Dim lCONTINUA As Boolean
Dim cTIPOCON As String

Private Sub CmdAbrir_Click()
    Dim nPOS As Integer
    Dim cTMP As String
    Dim aRETU As Variant
    
    FILTRO = ""
    cTIPOCON = ""
    If Len(aRELCFG(14)) > 0 Then FILTRO = aRELCFG(14)
    
    If InStr(UCase(arquivo), ".DBF") > 0 Then
        nPOS = InStrRev(arquivo, "\")
        If nPOS > 0 Then
            arquivo = Mid(arquivo, 1, nPOS)
        End If
        tabela = NomeArq(aARQUIVOS(0), True)
        TIPO = "DBF"
        cTIPOCON = "JETFOX"
    Else
        If Len(aRELCFG(13)) = 0 Then
            Alert ("Tabela Não Preenchida")
            lCONTINUA = False
        Else
            tabela = aRELCFG(13)
        End If
        If InStr(UCase(arquivo), ".MDB") > 0 Then
            TIPO = "MDB"
            cTIPOCON = "JETMDB"
        End If
    End If
    If lCONTINUA Then
        CmdGera_Click
    End If
    If ePASS01 = "CRYSTAL" Then
        TIPO = "CRYSTAL"
    Else
        cTMP = GeraConn(arquivo, cTIPO)
        aRETU = TipoConn(cTMP)
        TIPO = aRETU(2)
    End If
End Sub

Private Sub CmdGera_Click()
    Dim DB As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim cCAM As String
    Dim cSQL As String
    'Dim cCAMPO As String
    Dim STRCAMPO As String
    Dim strtipo As String
    Dim strsize As String
    Dim strdec As String
    Dim x As Integer
    Dim RstCampos

    If Len(FixStr(arquivo)) = 0 Then
        Alert ("Arquivo Nao Definido")
        Exit Sub
    End If

    Screen.MousePointer = vbHourglass
    
    cCAM = GeracArq(arquivo, cTIPOCON, False)
    cSQL = tabela
    Set DB = New ADODB.Connection
    DB.ConnectionTimeout = 120
    DB.CursorLocation = adUseClient
    DB.Open cCAM


    cSQL = tabela.Text
    grid.Clear
    
    
    
    With grid
        .Clear
        .Cols = 4
        .FixedCols = 0
        .Rows = 1
        .Row = 0
        .Col = 0
        .Text = "Campo"
        .ColWidth(0) = 2000
        
        .Col = 1
        .Text = "Tipo"
        .ColWidth(1) = 400
        
        .Col = 2
        .Text = "Tipo"
        .ColWidth(2) = 1400
        
        .Col = 3
        .Text = "Tamanho"
        .ColWidth(3) = 800
        
        
        
    End With
        

    'Adicionar campos das tabelas
    Set RstCampos = DB.OpenSchema(adSchemaColumns, Array(Empty, Empty, tabela.Text, Empty))
    Do While Not RstCampos.EOF
        STRCAMPO = RstCampos!Column_Name
        strtipo = FixStr((TipoDados(RstCampos!DATA_TYPE)))
        strsize = FixStr(RstCampos!CHARACTER_MAXIMUM_LENGTH)
        If Len(strsize) = 0 Then
            strdec = FixStr(RstCampos!NUMERIC_SCALE)
            If Len(strdec) = 0 Then
                strsize = FixStr(RstCampos!NUMERIC_PRECISION)
            Else
                strsize = FixStr(RstCampos!NUMERIC_PRECISION) & "." & strdec
            End If
        End If
        grid.AddItem STRCAMPO & vbTab & TipoDado2(RstCampos!DATA_TYPE) & vbTab & strtipo & vbTab & strsize
        RstCampos.MoveNext
    Loop
    DB.Close
    Screen.MousePointer = vbDefault
    Exit Sub

    
trataerro:
    Select Case Err.Number
    Case -2147217887
        GoTo monta02
    Case Else
        Resume Next
    End Select
monta02:
    ''     RstCampos.Close
    Set rs = New ADODB.Recordset
    rs.Open cSQL, DB, adOpenForwardOnly, adLockReadOnly
    For x = 0 To rs.Fields.Count - 1
        STRCAMPO = rs.Fields(x).Name
        grid.AddItem STRCAMPO & vbTab & TipoDado2(rs.Fields(x).Type) & vbTab & TipoDados(rs.Fields(x).Type) & vbTab & rs.Fields(x).DefinedSize
    Next x
    rs.Close
    DB.Close
    Screen.MousePointer = vbDefault
    Set DB = Nothing
    Set rs = Nothing
    Exit Sub


End Sub

Private Sub CmdSair_Click()
    Unload Me
End Sub

Private Sub CmdZerar_Click()
    FILTRO.Text = ""
End Sub

Private Sub Encerrar_Click()
    lRETU = True
    Unload Me
End Sub

Private Sub Form_Load()
    ''Dim nPOS As Integer
    On Error GoTo trataerro
  
    'CentralizaJanela Me

    CenterFormToScreen Me
    ''Configura Help
    Me.Caption = cFORMID
    HelpContextID = nFORMID
    
    lCONTINUA = True
    lRETU = False
    ValINI.Enabled = False
    ValFim.Enabled = False
    Vazio.Enabled = False
    ok.Enabled = False
    arquivo = aARQUIVOS(0)
    CmdAbrir_Click
    TIPOFIL(3).Enabled = False
    TIPOFIL(2).Enabled = False
    oper(0).Enabled = False
    oper(1).Enabled = False
    oper(2).Enabled = False
    oper(3).Enabled = False
    oper(4).Enabled = False
    oper(5).Enabled = False
    oper(6).Enabled = False
   
    If Len(aRELCFG(12)) > 0 Then
        Alias = aRELCFG(12)
    Else
        Alias = aRELCFG(13)                      ''Nome da tabela
        If Len(Alias) = 0 Then                   ''DBF E o nome do arquivo
            Alias = NomeArq(aARQUIVOS(0), True)
        End If
    End If
    If Not lCONTINUA Then
        lRETU = False
        Unload Me
    End If
    
    Exit Sub
trataerro:
    Select Case Err.Number
    Case Else
        SayErro "Filtro Load"
    End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    'lRETU = True
    eRETU01 = FILTRO
End Sub

Private Sub grid_Click()
    ValINI.Enabled = True
    ValFim.Enabled = True
    ok.Enabled = True
    Encerrar.Enabled = False
    ValINI.SetFocus
    Vazio.Enabled = True
    TIPOFIL(3).Enabled = True
    TIPOFIL(2).Enabled = True
    oper(0).Enabled = True
    oper(1).Enabled = True
    oper(2).Enabled = True
    oper(3).Enabled = True
    oper(4).Enabled = True
    oper(5).Enabled = True
    oper(6).Enabled = True
  
  
    TIPOFIL(2) = False
    TIPOFIL(3) = True
 
    oper(0) = True
    oper(1) = False
    oper(2) = False
    oper(3) = False
    oper(4) = False
    oper(5) = False
    oper(6) = False
    
    Vazio.Value = vbUnchecked
 
   
    grid.Col = 0
    campo = grid
    grid.Col = 1
    camtip = grid
    Select Case camtip
    Case "N"
        ValINI = "0.00000"
        ValFim = "0.00000"
    Case "N"
        ValINI = "0"
        ValFim = "0"
    Case "S"
        ValINI = ""
        ValFim = ""
    Case "D"
        ValINI = Format(Date, "dd/mm/yyyy")
        ValFim = Format(Date, "dd/mm/yyyy")
    End Select
End Sub

Private Sub ok_Click()
    Dim x As Long
    Dim nVEZ As Long
    Dim aUSO As Variant
    If InStr(ValINI, ";") > 0 Then
        aUSO = Split(ValINI, ";")
        ValFim = ""
        nVEZ = UBound(aUSO)
        For x = 0 To nVEZ
            ValINI = aUSO(x)
            If x = nVEZ Then
                ok_ClickX (True)
            Else
                ok_ClickX (False)
            End If
        Next x
    Else
        ok_ClickX (True)
    End If
End Sub

Private Sub ok_ClickX(lZERA As Boolean)
    Dim cVALINI, cVALFIM, cDIZ As String
    Dim nVALINI, nVALFIM As Variant
    Dim cTIPDIZ As String
    Dim cOPE, cOP2 As String


    ValINI.Enabled = False
    ValFim.Enabled = False
    ok.Enabled = False
    Encerrar.Enabled = True
    Vazio.Enabled = False
    TIPOFIL(3).Enabled = False
    TIPOFIL(2).Enabled = False
    oper(0).Enabled = False
    oper(1).Enabled = False
    oper(2).Enabled = False
    oper(3).Enabled = False
    oper(4).Enabled = False
    oper(5).Enabled = False
    oper(6).Enabled = False
   
    cVALINI = ""
    cVALFIM = ""
   
    '******************************************************************
    ''cDIZ = "{" & Trim(UCase(TABELA)) & "." & Trim(UCase(campo)) & "}"
    ''As Tabelas Normais aLIAS Contem o Nome do Arquivo
    ''Os Fechados tem ALIAS COMO PADRAO ARQUIVO RPT ALIAS DEVE SER ARQUIVO
    '******************************************************************

    cDIZ = "{" & Trim(UCase(CStr(Alias))) & "." & Trim(UCase(campo)) & "}"
   
    cTIPDIZ = " AND "
    If cTIPO.Text = "OU" Then cTIPDIZ = " OR "
   
    Select Case camtip
    Case "N"
        nVALINI = Val(ValINI)
        nVALFIM = Val(ValFim)
        If nVALINI > 0 Then
            cVALINI = Str(nVALINI)
        End If
        If nVALFIM > 0 Then
            cVALFIM = Str(nVALFIM)
            If nVALINI = 0 Then cVALINI = "0"
        End If
    Case "S"
        If Len(ValINI) > 0 Then
            If coper = "LIKE" Then
                cVALINI = Chr(39) & "%" & ValINI & "%" & Chr(39)
                cVALFIM = ""
            Else
                cVALINI = Chr(39) & ValINI & Chr(39)
                If Len(ValFim) > 0 Then
                    cVALFIM = Chr(39) & ValFim & Chr(39)
                End If
            End If
        End If
    Case "D"
        cVALINI = DataToLit(ValINI, TIPO)
        ''"CDATE(" & Mid(ValINI, 7, 4) & "," & Mid(ValINI, 4, 2) & "," & Mid(ValINI, 1, 2) & ")"
        If Len(ValFim) > 0 Then
            cVALFIM = DataToLit(ValFim, TIPO)
            '' cVALFIM = "CDATE(" & Mid(ValFim, 7, 4) & "," & Mid(ValFim, 4, 2) & "," & Mid(ValFim, 1, 2) & ")"
        End If
    End Select
   
    cOPE = coper

  
    If Vazio.Value = vbChecked Then
        If Len(FILTRO) > 0 Then FILTRO = FILTRO & cTIPDIZ
        Select Case camtip
        Case "N"
            FILTRO = FILTRO & "(" & cDIZ & cOPE & "0" & ")"
        Case "S"
            FILTRO = FILTRO & "(" & "LEN(" & cDIZ & ")" & cOPE & "0" & ")"
        Case "D"
            FILTRO = FILTRO & "(" & "DAY(" & cDIZ & ")" & coper & "0" & ")"
        End Select
    Else
        If Len(cVALINI) > 0 Or Len(cVALFIM) > 0 Then
            If Len(FILTRO) > 0 Then FILTRO = FILTRO & cTIPDIZ
            If Len(cVALFIM) > 0 Then
                If Len(cVALINI) > 0 Then
                    cOPE = ">="                  '0
                    cOP2 = "<="                  '0
                    If coper = "<" Or coper = ">" Then
                        cOPE = ">"
                        cOP2 = "<"
                    End If
                    FILTRO = FILTRO & "(" & cDIZ & cOPE & cVALINI & " AND " & cDIZ & cOP2 & cVALFIM & ")"
                Else
                    cOP2 = "<="                  '0
                    If coper = "<" Or coper = ">" Then
                        cOP2 = "<"
                    End If
                    FILTRO = FILTRO & "(" & cDIZ & cOP2 & cVALFIM & ")"
                End If
            Else
                If coper <> "LIKE" Then
                    FILTRO = FILTRO & "(" & cDIZ & cOPE & cVALINI & ")"
                Else
                    FILTRO = FILTRO & "(" & cDIZ & " " & cOPE & " " & cVALINI & ")"
                End If
            End If
        End If
    End If
    If lZERA Then
        campo = ""
        camtip = ""
    End If
 
End Sub

Private Sub oper_Click(Index As Integer)
    If Index = 6 Then
        If camtip <> "S" Then
            Alert ("Somente Para Campos Caracteres")
            Exit Sub
        End If
    End If
    coper.Text = oper(Index).Caption
End Sub

Private Sub TIPOFIL_Click(Index As Integer)
    cTIPO.Text = TIPOFIL(Index).Caption
End Sub

Private Sub ValFim_GotFocus()
    FocusMe
End Sub

Private Sub ValINI_GotFocus()
    FocusMe
End Sub

Private Sub ValINI_KeyPress(KeyAscii As Integer)
    Select Case camtip
    Case "N"
        KeyAscii = ValiText(KeyAscii, "#N")
    Case "D"
        KeyAscii = ValidaDados(KeyAscii, "DL")
    End Select
End Sub

Private Sub Vazio_Click()
    If Vazio.Value = vbChecked Then
        ValINI.Enabled = False
        ValFim.Enabled = False
    Else
        ValINI.Enabled = True
        ValFim.Enabled = True
    End If
    Select Case camtip
    Case "7"
        ValINI = "0.00000"
        ValFim = "0.00000"
    Case "3"
        ValINI = "0"
        ValFim = "0"
    Case "10", "12"
        ValINI = ""
        ValFim = ""
    Case "8"
        ValINI = Format(Date, "dd/mm/yyyy")
        ValFim = Format(Date, "dd/mm/yyyy")
    End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    TeclaEnter KeyCode
End Sub

