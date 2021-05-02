VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{7020C36F-09FC-41FE-B822-CDE6FBB321EB}#1.0#0"; "vbccr17.ocx"
Begin VB.Form NewSeq 
   Caption         =   "Incluindo Nova Sequenica"
   ClientHeight    =   1815
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6060
   LinkTopic       =   "Form1"
   ScaleHeight     =   1815
   ScaleWidth      =   6060
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox PF 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   600
      Width           =   1095
   End
   Begin XPControls.XPButton Gravar 
      Height          =   495
      Left            =   4320
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   600
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   873
      Picture         =   "NewSeq.frx":0000
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
   Begin VBCCR17.SpinBox seq 
      Height          =   495
      Left            =   1440
      TabIndex        =   7
      Top             =   600
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Min             =   1
      Max             =   999
      Value           =   10
      Increment       =   5
   End
   Begin VBCCR17.SpinBox Ssq 
      Height          =   495
      Left            =   2400
      TabIndex        =   8
      Top             =   600
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Min             =   1
      Max             =   999
      Value           =   10
      Increment       =   5
   End
   Begin VBCCR17.SpinBox item 
      Height          =   495
      Left            =   3360
      TabIndex        =   9
      Top             =   600
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Min             =   1
      Max             =   999
      Value           =   1
   End
   Begin VB.Label Label4 
      Caption         =   "Item"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   3360
      TabIndex        =   5
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "PF"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "SEQ=999 para operacao de controles"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "SSQ"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2400
      TabIndex        =   1
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Seq"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "NewSeq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    TeclaEnter KeyCode


End Sub

Private Sub Form_Load()
    Center Me
    PF.tEXT = nPF
    If FixInt(nPF) = 0 Then
        PF.Enabled = True
        PF.Locked = False
        PF.BackColor = &H80000005
    End If
    If iMU01 <> 3 And iMU01 <> 4 Then
        item.Enabled = False
        item.Visible = False
        item.Locked = True
    End If
    If iMU01 = 4 Then
        NewSeq.SEQ.Value = nSEQ
        NewSeq.SSQ.Value = nSSQ
        NewSeq.item.Value = nORD
    End If
   
End Sub

Private Sub Gravar_Click()
    Dim sSQL As String
    If iMU01 = 1 Then
        nSEQ = FixInt(NewSeq.SEQ)
        nSSQ = FixInt(NewSeq.SSQ)
        sSQL = "select pf,seq,ssq from PFS WHERE PF=" & nPF & " AND SEQ=" & nSEQ & " AND SSQ=" & nSSQ
        IncluiSQL Sdb, sSQL, 3, Array("PF", "SEQ", "SSQ"), _
        Array(nPF, nSEQ, nSSQ), True, True
    End If
    If iMU01 = 2 Then
        'PFS
        sSQL = "select pf,seq,ssq from PFS WHERE PF=" & nPF & " AND SEQ=" & nSEQ & " AND SSQ=" & nSSQ
        REORDER sSQL, "PF"
        REORDER sSQL, "PFP"
        REORDER sSQL, "PFG"
        'PFI
        sSQL = "select pf,seq,ssq from PFI WHERE PF=" & nPF & " AND SEQ=" & nSEQ & " AND SSQ=" & nSSQ
        REORDER sSQL, "PF"
        REORDER sSQL, "PFP"
        REORDER sSQL, "PFG"
        'PFC
        sSQL = "select pf,seq,ssq from PFC WHERE PF=" & nPF & " AND SEQ=" & nSEQ & " AND SSQ=" & nSSQ
        REORDER sSQL, "PF"
        REORDER sSQL, "PFP"
        REORDER sSQL, "PFG"
        'pfqsqblep (fluxo 2)
        sSQL = "select pf,seq,ssq from PFQSBLEP WHERE PF=" & nPF & " AND SEQ=" & nSEQ & " AND SSQ=" & nSSQ
        REORDER sSQL, "PF"
        REORDER sSQL, "PFP"
        REORDER sSQL, "PFG"
    End If
    
    If iMU01 = 3 Then
        nSEQ = FixInt(NewSeq.SEQ.Value)
        nSSQ = FixInt(NewSeq.SSQ.Value)
        nPF = FixInt(NewSeq.PF.tEXT)
        nORD = FixInt(NewSeq.item.Value)
        If nPF = 0 Then
            Alert ("PF Nao preenchido")
            Exit Sub
        End If
        If nORD = 0 Then
            Alert ("item Nao preenchido")
            Exit Sub
        End If
    End If
    
    If iMU01 = 4 Then
        nSEQ = FixInt(NewSeq.SEQ.Value)
        nSSQ = FixInt(NewSeq.SSQ.Value)
        nORD = FixInt(NewSeq.item.Value)
        sSQL = "select * from poKa WHERE numero=" & nPPAP
        REORDER sSQL, "POKA"
    End If
    

    
    If iMU01 = 1 Then
        eRETU02 = "PF:" & nPF & Chr(13) & Chr(10)
        eRETU02 = eRETU02 & " Desenho:" & frmPF.TXTFIELDS(2) & Chr(13) & Chr(10)
        eRETU02 = eRETU02 & " Data:" & Fdata(Date) & Chr(13) & Chr(10)
        eRETU02 = eRETU02 & " SEQ=" & nSEQ & " SSQ=" & nSSQ
        MAILENV "PFS00001", eRETU02
    End If
   
    Unload Me
    
End Sub

Private Sub REORDER(sSQL As Variant, Optional ByVal cARQ As String = "")
    Dim DB As New ADODB.Connection
    Dim RSTAB As New ADODB.Recordset
    Select Case cARQ
    Case "PF"
        cARQ = PegPath("PATH", "PF")
    Case "PFP"
        cARQ = PegPath("PATH", "PFP")
    Case "PFG"
        cARQ = PegPath("PATH", "PFG")
    Case "POKA"
        cARQ = PegPath("PATH", "POKA")
    Case Else
        cARQ = Sdb
    End Select
    cARQ = GeracArq(cARQ)
    DB.ConnectionTimeout = 120
    DB.Open cARQ
    RSTAB.Open sSQL, DB, adOpenDynamic, adLockOptimistic
    Do While Not RSTAB.EOF
        RSTAB("SEQ") = FixInt(NewSeq.SEQ)
        RSTAB("SSQ") = FixInt(NewSeq.SSQ)
        If iMU01 = 4 Then
            RSTAB("ITEM") = FixInt(NewSeq.item)
        End If
        RSTAB.Update
        RSTAB.MoveNext
    Loop
    RSTAB.Close
End Sub

Private Sub SEQ_GotFocus()
    FocusMe
End Sub

Private Sub SEQ_KeyPress(KeyAscii As Integer)
    KeyAscii = ValiText(KeyAscii, "#NI")
End Sub

Private Sub SSQ_GotFocus()
    FocusMe
End Sub

Private Sub SSQ_KeyPress(KeyAscii As Integer)
    KeyAscii = ValiText(KeyAscii, "#NI")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault
End Sub


