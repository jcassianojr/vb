VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{7020C36F-09FC-41FE-B822-CDE6FBB321EB}#1.0#0"; "vbccr17.ocx"
Begin VB.Form frmPPAPI 
   Caption         =   "Itens PPAP"
   ClientHeight    =   3390
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9105
   Icon            =   "frmPPAPI.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3390
   ScaleWidth      =   9105
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TXTFIELDS 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   375
      Index           =   4
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox TXTFIELDS 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   375
      Index           =   5
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   720
      Width           =   1275
   End
   Begin VB.TextBox TXTFIELDS 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   375
      Index           =   6
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   720
      Width           =   3675
   End
   Begin VB.CommandButton Setor 
      Caption         =   "Interina"
      Height          =   315
      Index           =   3
      Left            =   5220
      TabIndex        =   14
      Top             =   60
      Width           =   1575
   End
   Begin VB.CommandButton Setor 
      Caption         =   "Aprovado"
      Height          =   315
      Index           =   0
      Left            =   3600
      TabIndex        =   13
      Top             =   60
      Width           =   1575
   End
   Begin VB.CommandButton Setor 
      Caption         =   "Reprovado"
      Height          =   315
      Index           =   1
      Left            =   3600
      TabIndex        =   12
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton Setor 
      Caption         =   "Outro"
      Height          =   315
      Index           =   2
      Left            =   5220
      TabIndex        =   11
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox TXTFIELDS 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   2460
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   300
      Width           =   975
   End
   Begin VB.TextBox TXTFIELDS 
      Height          =   1005
      Index           =   3
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   2040
      Width           =   8835
   End
   Begin VB.TextBox TXTFIELDS 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   1380
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   300
      Width           =   975
   End
   Begin VB.TextBox TXTFIELDS 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   300
      Width           =   975
   End
   Begin XPControls.XPButton Encerrar 
      Height          =   435
      Left            =   7320
      TabIndex        =   18
      Top             =   600
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "frmPPAPI.frx":058A
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
   Begin XPControls.XPButton cmdClose 
      Height          =   435
      Left            =   7320
      TabIndex        =   19
      Top             =   120
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "frmPPAPI.frx":0B24
      Caption         =   "Salvar"
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
   Begin VBCCR17.DTPicker DTPicker1 
      Height          =   375
      Left            =   1080
      TabIndex        =   20
      Top             =   1320
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Value           =   44285
      CheckBox        =   -1  'True
      AllowUserInput  =   -1  'True
   End
   Begin VBCCR17.DTPicker DTPicker2 
      Height          =   375
      Left            =   3840
      TabIndex        =   21
      Top             =   1320
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Value           =   44285
      CheckBox        =   -1  'True
      AllowUserInput  =   -1  'True
   End
   Begin VBCCR17.DTPicker DTPicker3 
      Height          =   375
      Left            =   6360
      TabIndex        =   22
      Top             =   1320
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Value           =   44285
      CheckBox        =   -1  'True
      AllowUserInput  =   -1  'True
   End
   Begin VB.Label lblLabels 
      Caption         =   "Prazo"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   6
      Left            =   5760
      TabIndex        =   17
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "Aprovaçäo"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   5
      Left            =   2880
      TabIndex        =   16
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "Abertura"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   15
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label lblLabels 
      Caption         =   "Disposição"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   3
      Left            =   2520
      TabIndex        =   10
      Top             =   60
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "Observação:"
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "Item"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   6
      Top             =   60
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "PPAP:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   60
      Width           =   975
   End
End
Attribute VB_Name = "frmPPAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cARQPF As String
Dim cSQL As String
Dim aVAL As Variant
Dim aFOR As Variant
Dim aCAM As Variant
Dim aPAD As Variant
Dim nCAMPOS As Integer
Dim iLOOP As Integer

Private Sub cmdClose_Click()
    On Error Resume Next
    If MDG("Gravar alteraçôes") Then
        For iLOOP = 0 To nCAMPOS - 4
            aVAL(iLOOP) = TXTFIELDS(iLOOP)
        Next iLOOP
        '         aVAL(4) = DTPicker1
        '         aVAL(5) = DTPicker2
        '         aVAL(6) = DTPicker3
        aVAL(7) = DTPicker1
        aVAL(8) = DTPicker2
        aVAL(9) = DTPicker3
         
        GrvSQL cARQPF, cSQL, nCAMPOS, aCAM, aVAL, aFOR
    End If
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub Encerrar_Click()
    If Not MDG("Sair sem gravar") Then
        Exit Sub
    End If
    Screen.MousePointer = vbDefault
    Unload Me
 
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    TeclaEnter KeyCode
End Sub

Private Sub Form_Load()

    CenterFormToScreen Me
    cARQPF = PegPath("PATH", "PPAP")
    
    
    Select Case iPPAP                            '
    Case 1, 11
        cSQL = "select * from PPAPI WHERE PPAP=" & nPPAP & " AND ITEM=" & nREVI
    Case 2
        cSQL = "select * from PPAFI WHERE PPAP=" & nPPAP & " AND ITEM=" & nREVI
        Me.Caption = "Item PPAP Fornecedores"
    Case 3
        cSQL = "select * from PPAGI WHERE PPAP=" & nPPAP & " AND ITEM=" & nREVI
        Me.Caption = "Item GP11"
    End Select
    
    nCAMPOS = 10                                 '7
    aCAM = Array("PPAP", "ITEM", "DISPO", "OBS", "DISPNUM", "DISPDATE", "DISPNOME", "DATA", "APROVADO", "PRAZO")
    aFOR = Array("NI", "NI", "C", "C", "NI", "D", "C", "D", "D", "D")
    aPAD = Array(0, 0, "", "", 0, "", "", Date, Date, Date)
    aVAL = PegSQL(cARQPF, cSQL, nCAMPOS, aCAM, aFOR, aPAD)
    For iLOOP = 0 To nCAMPOS - 4
        TXTFIELDS(iLOOP) = aVAL(iLOOP)
    Next iLOOP

    If IsDate(aVAL(7)) Then
        DTPicker1.Value = aVAL(7)
    End If
    If IsDate(aVAL(8)) Then
        DTPicker2.Value = aVAL(8)
    End If
    If IsDate(aVAL(9)) Then
        DTPicker3.Value = aVAL(9)
    End If
    
    
End Sub

Private Sub Setor_Click(Index As Integer)
    Dim DB As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim RSTAB As New ADODB.Recordset

    Dim cSQL As String
    Dim sSQL  As String

    
    On Error Resume Next
    
    Select Case Index
    Case 0
        TXTFIELDS(2) = "A"
        eRETU03 = "Aprovado"
    Case 1
        TXTFIELDS(2) = "R"
        eRETU03 = "Reprovado"
    Case 2
        TXTFIELDS(2) = "O"
        eRETU03 = "Outros"
    Case 3
        TXTFIELDS(2) = "I"
        eRETU03 = "Interina"
    End Select
    
    
    TXTFIELDS(4) = zIDFOLHA
    TXTFIELDS(5) = zNOMEFOLHA
    TXTFIELDS(6) = Date
    
    
    If Index = 0 Then
        If MDG("Gravar Cronograma", "Gravar Cronograma") Then
            Select Case iPPAP
            Case 1, 11
                cSQL = "select * from PPAPP WHERE PPAP=" & nPPAP & " ORDER BY CODIGO"
            Case 2
                cSQL = "select * FROM PPAFP WHERE PPAP=" & nPPAP & " ORDER BY CODIGO"
            Case 3
                cSQL = "select * FROM PPAGP WHERE PPAP=" & nPPAP & " ORDER BY CODIGO"
            End Select
            Select Case iPPAP
            Case 1, 11
                sSQL = "PPAPC"
            Case 2
                sSQL = "PPAFC"
            Case 3
                sSQL = "PPAGC"
            End Select
            
            DB.ConnectionTimeout = 120
            DB.Open GeracArq(cARQPF)
            RSTAB.Open cSQL, DB, adOpenForwardOnly, adLockReadOnly
            rs.Open sSQL, DB, adOpenDynamic, adLockOptimistic
            If Not RSTAB.EOF Then
                RSTAB.MoveFirst
                While Not RSTAB.EOF
                    rs.AddNew
                    rs("codigo") = RSTAB("CODIGO")
                    rs("PPAP") = nPPAP
                    rs("DATA") = Today()
                    rs("PREVISTO") = Today() + 182
                    rs.Update
                    RSTAB.MoveNext
                Wend
            End If
            rs.Close
            RSTAB.Close
            DB.Close

       
        End If
    End If
    
    eRETU02 = ""
    Select Case iPPAP
    Case 1, 11
        eRETU02 = "PPAP:" & Trim(Str(nPPAP)) & Chr(13) & Chr(10)
    Case 2
        eRETU02 = "PPAP Fornecedor:" & Trim(Str(nPPAP)) & Chr(13) & Chr(10)
    Case 3
        eRETU02 = "GP11:" & Trim(Str(nPPAP)) & Chr(13) & Chr(10)
    End Select
    

    eRETU02 = eRETU02 & "PF:" & frmPPAP.TXTFIELDS(32) & Chr(13) & Chr(10)
    eRETU02 = eRETU02 & "Código:" & frmPPAP.TXTFIELDS(5) & Chr(13) & Chr(10)
    If IsDate(DTPicker1) Then
        eRETU02 = eRETU02 & " Abertura:" & Fdata(DTPicker1) & Chr(13) & Chr(10)
    End If
    If IsDate(DTPicker2) Then
        eRETU02 = eRETU02 & " Data:" & Fdata(DTPicker2) & Chr(13) & Chr(10)
    End If
    If IsDate(DTPicker3) Then
        eRETU02 = eRETU02 & " Prazo:" & Fdata(DTPicker3) & Chr(13) & Chr(10)
    End If
    eRETU02 = eRETU02 & " Tipo:" & eRETU03 & Chr(13) & Chr(10)
      
    MAILENV "PPAP0001", eRETU02
      
       
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault
End Sub

