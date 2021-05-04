VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{7020C36F-09FC-41FE-B822-CDE6FBB321EB}#1.0#0"; "vbccr17.ocx"
Object = "{F22668DE-E08D-467B-8E41-13900013BD5F}#2.7#0"; "VBextra2.OCX"
Begin VB.Form imgsin 
   Caption         =   "Sincronizar Arquivos de Imagens"
   ClientHeight    =   7395
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8370
   Icon            =   "imgsin.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7395
   ScaleWidth      =   8370
   StartUpPosition =   2  'CenterScreen
   Begin vbExtra.SizeGrip SizeGrip1 
      Height          =   285
      Left            =   8085
      Top             =   7110
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
   End
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   3840
      ScaleHeight     =   435
      ScaleWidth      =   195
      TabIndex        =   20
      Top             =   2520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture2 
      Height          =   3375
      Left            =   120
      ScaleHeight     =   3315
      ScaleWidth      =   3555
      TabIndex        =   19
      Top             =   2520
      Width           =   3615
   End
   Begin XPControls.XPButton XPButton1 
      Height          =   375
      Left            =   6000
      TabIndex        =   18
      Top             =   1080
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   661
      Caption         =   "Ajustar imagem grandes"
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
   Begin VBCCR17.SpinBox TxtEmpresa 
      Height          =   375
      Left            =   5040
      TabIndex        =   17
      Top             =   1320
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      Min             =   1
      Max             =   99
      Value           =   1
   End
   Begin VB.CheckBox ChkMQ01 
      Caption         =   "Sub Produtos"
      Height          =   375
      Left            =   1920
      TabIndex        =   14
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CheckBox Chkmw07 
      Caption         =   "Itens Manutençao"
      Height          =   375
      Left            =   3600
      TabIndex        =   13
      Top             =   840
      Width           =   1335
   End
   Begin VB.CheckBox ChKME06 
      Caption         =   "Utilitarios"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   1335
   End
   Begin VB.CheckBox ChKMR01 
      Caption         =   "Embalagens"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   1455
   End
   Begin VB.CheckBox Chkma01 
      Caption         =   "Clientes"
      Height          =   315
      Left            =   3600
      TabIndex        =   10
      Top             =   120
      Width           =   1095
   End
   Begin VB.CheckBox Chkme01 
      Caption         =   "Maq/Equipa"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   1575
   End
   Begin VB.CheckBox Chkmw05 
      Caption         =   "Consumiveis"
      Height          =   315
      Left            =   3600
      TabIndex        =   8
      Top             =   480
      Width           =   1335
   End
   Begin VB.CheckBox Chkmp04 
      Caption         =   "Funcionarios"
      Height          =   375
      Left            =   3600
      TabIndex        =   7
      Top             =   1260
      Width           =   1275
   End
   Begin VB.CheckBox ChKME04 
      Caption         =   "Instrumentos"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CheckBox Chkmu01 
      Caption         =   "Mat.Prima"
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CheckBox ChKMS01 
      Caption         =   "Produtos/PCP"
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   480
      Width           =   1455
   End
   Begin VB.CheckBox chkfer 
      Caption         =   "Ferramentas"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CheckBox Chkmt01 
      Caption         =   "Componentes"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   840
      Width           =   1455
   End
   Begin VB.CheckBox ChKPF 
      Caption         =   "Produtos/PF"
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin XPControls.XPButton CmdTransfer 
      Height          =   375
      Left            =   6000
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   600
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   661
      Picture         =   "imgsin.frx":058A
      Caption         =   "Sincronizar Arquivos"
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
   Begin XPControls.XPButton CmdCancel 
      Height          =   375
      Left            =   6000
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   180
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   661
      Picture         =   "imgsin.frx":0B24
      Caption         =   "Cancelar"
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
   Begin VB.Label LABEL 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   2100
      Width           =   7215
   End
End
Attribute VB_Name = "imgsin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub CmdTransfer_Click()

    If chkfer.Value = Checked Then

        transfer "IMGFER", "MANA5FER", "SELECT FERRAM,NUMERO FROM FERRAM ", "FERRAM", False, "NUMERO"

    End If

    If Chkmt01.Value = Checked Then

        transfer "IMGMT01", "MANA5EMP", "SELECT CODIGO FROM MT01 ", "CODIGO", False

    End If

    If ChkMQ01.Value = Checked Then

        transfer "IMGMQ01", "MANA5EMP", "SELECT CODIGO FROM MQ01 ", "CODIGO", False

    End If

    If ChKMR01.Value = Checked Then

        transfer "IMGMR01", "MANA5EMP", "SELECT CODIGO FROM MR01 ", "CODIGO", False

    End If

    If Chkmu01.Value = Checked Then

        transfer "IMGMU01", "MANA5EMP", "SELECT CODIGO FROM MU01 ", "CODIGO", False

    End If

    If Chkmw05.Value = Checked Then

        transfer "IMGMW05", "MANA5EMP", "SELECT CODIGO FROM MW05 ", "CODIGO", False

    End If

    If Chkmw07.Value = Checked Then

        transfer "IMGMW07", "MANA5EMP", "SELECT CODIGO FROM MW07 ", "CODIGO", False

    End If


    If ChKMS01.Value = Checked Then

        transfer "IMGMS01", "MANA5EMP", "SELECT CODIGO FROM MS01 ", "CODIGO", False
        transfer "IMGMS01", "MANA5EMP", "SELECT CODIGO FROM MS01X ", "CODIGO", False
        fazcodiMG
    End If

    If ChKPF.Value = Checked Then

        transfer "IMGMS01", "APP", "SELECT CODIGO,CODIGOINT FROM PF ", "CODIGO", False, "CODIGOINT"

    End If

    If Chkme01.Value = Checked Then

        transfer "IMGME01", "MANA5EMP", "SELECT NUMERO FROM ME01 ", "NUMERO", False

    End If

    If Chkma01.Value = Checked Then

        transfer "IMGMA01", "MANA5EMP", "SELECT NUMERO FROM MA01 ", "NUMERO", True

    End If

    If Chkmp04.Value = Checked Then

        transfer "IMGMP04" + StrZero(TxtEmpresa.Value, 1), "MANA5EMZ", "SELECT TECNICO FROM MP04 ", "TECNICO", True

    End If

    If ChKME04.Value = Checked Then

        transfer "IMGME04", "MANA5INS", "SELECT CODIGO FROM ME04 ", "CODIGO", False

    End If
    
    If ChKME06.Value = Checked Then

        transfer "IMGME06", "MANA5INS", "SELECT CODIGO FROM ME06 ", "CODIGO", False

    End If


    Unload Me

End Sub

Private Sub transfer(cLOGOARQ As String, cARQCAM As String, cCONSQL As String, cVARNAME As String, lCONV As Boolean, Optional cNUM As String)
    Dim DB As New ADODB.Connection
    Dim RSTAB As New ADODB.Recordset
    Dim cCODIGO, cCAMINHO As String
    Dim nNUMERO As Variant
    Dim lCVARNUM As Boolean
    Dim cARQTMP As String
    Dim aRETU As Variant
    Dim DB2 As New ADODB.Connection
    Dim RSTA2 As New ADODB.Recordset

    
    cNUM = CStr(cNUM)
    If Len(cNUM) > 0 Then lCVARNUM = True
    
    

    cCAMINHO = PegPath("PATH", cLOGOARQ)
    If Not FileExist(cCAMINHO, True) Then
        Exit Sub
    End If
    
    
    DB2.Open GeracArq(cCAMINHO, "A12MDB")
    
    
    With RSTA2
        .ActiveConnection = DB2
        .Source = "IMAGENS"
        .CursorLocation = adUseServer
        .CursorType = adOpenDynamic              ' adOpenKeyset
        .LockType = adLockOptimistic
        .Open Options:=adCmdTableDirect
        .Index = "PrimaryKey"
    End With
    

    cCAMINHO = PegPath("PATH", cARQCAM)

    If cARQCAM = "APP" Then

        DB.Open GeracArq(cCAMINHO & "PF.MDB")
        RSTAB.Open cCONSQL, DB, adOpenForwardOnly, adLockReadOnly

    Else
        If cARQCAM = "MANA5EMZ" Then
            cCAMINHO = Caminex(cCAMINHO, 0, 0, TxtEmpresa.Value)
        End If
        cARQTMP = GeraConn(cCAMINHO, "JETFOX")
        aRETU = TipoConn(cARQTMP)
        cARQTMP = aRETU(1)
        DB.Open cARQTMP
        RSTAB.Open cCONSQL, DB, adOpenStatic, adLockOptimistic

    End If


    Do While Not RSTAB.EOF

        If lCONV Then

            nNUMERO = RSTAB(cVARNAME)
            cCODIGO = Trim(Str(nNUMERO))
            cCODIGO = String(8 - Len(cCODIGO), "0") & cCODIGO

        Else

            cCODIGO = Trim(RSTAB(cVARNAME))
            If lCVARNUM Then
                nNUMERO = funNumeroPuro(RSTAB(cNUM))
            Else
                nNUMERO = 0
            End If
            If Not IsNumeric(nNUMERO) Then
                nNUMERO = 0
            End If
        End If
        
        cCODIGO = FixStr(cCODIGO, "")
        DizerBarra cCODIGO
        If Len(cCODIGO) > 0 Then

            Label.Caption = cCODIGO
            Label.Refresh
            
            If RSTA2.RecordCount > 0 Then
                RSTA2.MoveFirst
                RSTA2.Seek cCODIGO, adSeekFirstEQ
            End If

            If RSTA2.EOF Then

                RSTA2.AddNew
                RSTA2("CODIGO") = cCODIGO
                RSTA2("NUMERO") = nNUMERO
                RSTA2.Update
            Else
                If (RSTA2("NUMERO") = 0 Or cLOGOARQ = "IMGFER") And nNUMERO > 0 Then
                    RSTA2("NUMERO") = nNUMERO
                    RSTA2.Update
                End If
            End If
            
            
        End If
        RSTAB.MoveNext
    Loop
    RSTAB.Close
    RSTA2.Close
    
    
    Set DB = Nothing
    Set RSTAB = Nothing
    Set DB2 = Nothing
    Set RSTA2 = Nothing

End Sub

Private Sub Form_Load()
    CenterFormToScreen Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault

End Sub

Private Sub fazcodiMG()
    Dim oCON As New ADODB.Connection
    Dim oRS As New ADODB.Recordset
    Dim cARQ As String
    Dim sSQL As String
    Dim cCODIGO As String
    'Dim nCODINT As String
    Dim aRETU As Variant
    Dim cARQPF As String
    cARQPF = GeracArq(PegPath("PATH", "IMGMS01"))
    cARQ = GeraConn(zMANA5EMP, "JETFOX")
    oCON.Open cARQPF
    oRS.Open "SELECT CODIGO,NUMERO FROM IMAGENS WHERE NUMERO=0", oCON, adOpenDynamic, adLockOptimistic
    While Not oRS.EOF
        cCODIGO = FixStr(oRS("CODIGO"), "", "TRIM")
        DizerBarra cCODIGO
        sSQL = "SELECT NOME,CODIGOINT FROM MS01 WHERE CODIGO='" & cCODIGO & "'"
        aRETU = PegSQL(cARQ, sSQL, 2, Array("NOME", "CODIGOINT"), Array("C", "C"), Array("", ""))
        If lRETU Then
            oRS("Numero") = Val(aRETU(1))
        Else
            sSQL = "SELECT NOME,CODIGOINT FROM MS01X WHERE CODIGO='" & cCODIGO & "'"
            aRETU = PegSQL(cARQ, sSQL, 2, Array("NOME", "CODIGOINT"), Array("C", "C"), Array("", ""))
            If lRETU Then
                oRS("Numero") = Val(aRETU(1))
                '' Else
                ''    oRS("Numero") = 0
            End If
        End If
        oRS.Update
        oRS.MoveNext
    Wend
    oRS.Close
    oRS.Open "SELECT CODIGO,NUMERO FROM IMAGENS", oCON, adOpenDynamic, adLockOptimistic
    While Not oRS.EOF
        DizerBarra cCODIGO
        If FixStr(oRS("NUMERO")) = FixStr(oRS("CODIGO")) Then
            oRS("Numero") = 0
            oRS.Update
            '' oRS.DELETE
        End If
        oRS.MoveNext
    Wend
End Sub

Private Sub XPButton1_Click()
    limpafoto "select * from pfs       where fig01     is not null  and len(fig01)>500000", "fig01", "PFS"
    limpafoto "select * from pfs       where fig02     is not null  and len(fig02)>500000", "fig02", "PFS"
    limpafoto "select * from pfs       where figembal  is not null  and len(figembal)>500000", "figembal", "PFS"
    limpafoto "select * from pfs       where figseq01  is not null  and len(figseq01)>500000", "figseq01", "PFS"
    limpafoto "select * from pfs       where figseq02  is not null  and len(figseq02)>500000", "figseq02", "PFS"
    limpafoto "select * from pfs       where imgcroqui is not null  and len(imgcroqui)>500000", "imgcroqui", "PFS"
    limpafoto "select * from pfs       where imgalerta is not null  and len(imgalerta)>500000", "imgalerta", "PFS"
    limpafoto "select * from pf        where figmu011  is not null  and len(figmu011)>500000", "figmu011", "PF"
    limpafoto "select * from pf        where figmu012  is not null  and len(figmu012)>500000", "figmu012", "PF"
    limpafoto "select * from pf        where figmu013  is not null  and len(figmu013)>500000", "figmu013", "PF"
    limpafoto "select * from pfc       where imagem    is not null  and len(imagem)>500000", "imagem", "PFC"
    limpafoto "select * from pfco      where imagem    is not null  and len(imagem)>500000", "imagem", "PFCO"
    limpafoto "select * from pfms03    where figseq01  is not null  and len(figseq01)>500000", "figseq01", "PFMS03"
    limpafoto "select * from pfms03    where figseq02  is not null  and len(figseq02)>500000", "figseq02", "PFMS03"
    limpafoto "select * from pfi       where imagem    is not null  and len(imagem)>500000 ", "imagem", "PFI"
    'limpafoto "", "imagem", "IMGMS01"
    limpafoto "select * from imagens   where imagem    is not null  and len(imagem)>500000 ", "imagem", "IMGMS01"
    Alert ("concluida")
End Sub

Private Function limpafoto(cSQL As String, cCAMPO As String, cTabela)
    Dim oDB              As ADODB.Connection
    Dim oRS              As ADODB.Recordset
    Dim cARQPF, cARQ, sTEMPFILE, cDIZARQ   As String
    Dim iFileNum, X, nFIM As Integer
    Dim lFileLength As Long
    Dim abBytes() As Byte

    nFIM = 3
    If cTabela = "IMGMS01" Then
        nFIM = 1
    End If
    If Len(cSQL) = 0 Then
        cSQL = "select * from " & cTabela & " where " & cCAMPO & " is not null and len(" & cCAMPO & ")>500000 "
    End If

    'On Error Resume Next
    For X = 1 To nFIM
        Select Case X
        Case 1
            Select Case cTabela
            Case "IMGMS01"
                cARQPF = PegPath("PATH", "IMGMS01")
                cDIZARQ = "IMGMS01"
            Case Else
                cARQPF = PegPath("PATH", "PF")
                cDIZARQ = "PF"
            End Select
        Case 2
            cARQPF = PegPath("PATH", "PFG")
            cDIZARQ = "PFG"
        Case 3
            cARQPF = PegPath("PATH", "PFP")
            cDIZARQ = "PFP"
        End Select
        Label.Caption = cDIZARQ & "_" & cCAMPO
        Label.Refresh
        cARQ = GeracArq(cARQPF)
        Set oDB = New ADODB.Connection
        oDB.CursorLocation = adUseClient
        oDB.ConnectionTimeout = 120
        oDB.Open cARQ
        Set oRS = New ADODB.Recordset
        oRS.Open cSQL, oDB, adOpenStatic, adLockOptimistic
        While Not oRS.EOF
            sTEMPFILE = ""
            Select Case cTabela
            Case "IMGMS01"
                sTEMPFILE = zAppPath & "\" & cDIZARQ & "_PF_" & oRS("codigo") & "_" & cCAMPO & ".JPG"
            Case "PF"
                sTEMPFILE = zAppPath & "\" & cDIZARQ & "_PF_" & StrZero(oRS("PF"), 6) & "_" & cCAMPO & ".JPG"
            Case "PFS"
                sTEMPFILE = zAppPath & "\" & cDIZARQ & "_PFS_" & StrZero(oRS("PF"), 6) & "_SEQ_" & StrZero(oRS("SEQ"), 3) & "_SSQ_" & StrZero(oRS("SEQ"), 3) & "_item_" & cCAMPO & ".JPG"
            Case "PFMS03"
                sTEMPFILE = zAppPath & "\" & cDIZARQ & "_PFI_" & StrZero(oRS("PF"), 6) & "_SEQ_" & StrZero(oRS("SEQ"), 3) & "_SSQ_" & StrZero(oRS("SEQ"), 3) & "_item_" & StrZero(oRS("ITEM"), 3) & "_" & cCAMPO & ".JPG"
            Case "PFCO"
                sTEMPFILE = zAppPath & "\" & cDIZARQ & "_PFI_" & StrZero(oRS("PF"), 6) & "_SEQ_" & StrZero(oRS("SEQ"), 3) & "_SSQ_" & StrZero(oRS("SEQ"), 3) & "_item_" & StrZero(oRS("ITEM"), 3) + ".JPG"
            Case "PFC"
                sTEMPFILE = zAppPath & "\" & cDIZARQ & "_PFI_" & StrZero(oRS("PF"), 6) & "_SEQ_" & StrZero(oRS("SEQ"), 3) & "_SSQ_" & StrZero(oRS("SEQ"), 3) & "_item_" & StrZero(oRS("ITEM"), 3) + ".JPG"
            Case "PFI"
                sTEMPFILE = zAppPath & "\" & cDIZARQ & "_PFI_" & StrZero(oRS("PF"), 6) & "_SEQ_" & StrZero(oRS("SEQ"), 3) & "_SSQ_" & StrZero(oRS("SEQ"), 3) & "_item_" & StrZero(oRS("ITEM"), 3) + ".JPG"
            End Select
            Label.Caption = cDIZARQ & "_" & cCAMPO
            Label.Refresh
            If sTEMPFILE <> "" Then
                Label.Caption = sTEMPFILE
                Label.Refresh
                lFileLength = LenB(oRS(cCAMPO))
                iFileNum = FreeFile
                Open sTEMPFILE For Binary As #iFileNum
                abBytes = oRS(cCAMPO).GetChunk(lFileLength)
                Put #iFileNum, , abBytes()
                Close #iFileNum
                Picture1.Picture = LoadPicture(sTEMPFILE)
                StretchSourcePictureFromPicture Picture1, Picture2
            End If
            oRS.Fields(cCAMPO) = ""
            oRS.Update
            oRS.MoveNext
        Wend
        oRS.Close
        oDB.Close
    Next X
    Set oRS = Nothing
    Set oDB = Nothing
End Function

