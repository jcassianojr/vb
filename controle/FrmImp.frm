VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{7020C36F-09FC-41FE-B822-CDE6FBB321EB}#1.0#0"; "vbccr17.ocx"
Begin VB.Form FrmImp 
   Caption         =   "Importando Dados"
   ClientHeight    =   3945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   ScaleHeight     =   3945
   ScaleWidth      =   6870
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox FemeaM 
      Caption         =   "Maquina"
      Height          =   255
      Left            =   3240
      TabIndex        =   15
      Top             =   1680
      Width           =   1212
   End
   Begin VB.CheckBox FemeaA 
      Caption         =   "Adicional"
      Height          =   255
      Left            =   3240
      TabIndex        =   14
      Top             =   1320
      Width           =   1212
   End
   Begin VB.CheckBox Final 
      Caption         =   "Plano Final"
      Height          =   372
      Left            =   120
      TabIndex        =   13
      Top             =   2760
      Width           =   2172
   End
   Begin VB.CheckBox laboratorio 
      Caption         =   "Plano Laboratorio"
      Height          =   372
      Left            =   120
      TabIndex        =   12
      Top             =   2280
      Width           =   2172
   End
   Begin VB.CheckBox Recebimento 
      Caption         =   "Plano de Recebimento"
      Height          =   372
      Left            =   120
      TabIndex        =   11
      Top             =   1920
      Width           =   2172
   End
   Begin VB.CheckBox FemeaG 
      Caption         =   "GP12"
      Height          =   255
      Left            =   3240
      TabIndex        =   10
      Top             =   960
      Width           =   1212
   End
   Begin VB.CheckBox FemeaR 
      Caption         =   "Pre"
      Height          =   255
      Left            =   3240
      TabIndex        =   9
      Top             =   600
      Width           =   1092
   End
   Begin VB.CheckBox FemeaP 
      Caption         =   "Padrao"
      Height          =   255
      Left            =   1680
      TabIndex        =   8
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CheckBox FemeaC 
      Caption         =   "Complementar"
      Height          =   255
      Left            =   1680
      TabIndex        =   7
      Top             =   960
      Width           =   1335
   End
   Begin VB.CheckBox FemeaN 
      Caption         =   "Producao"
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   600
      Width           =   1335
   End
   Begin VB.CheckBox Normal 
      Caption         =   "Producao"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CheckBox gp12 
      Caption         =   "GP12"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1335
   End
   Begin VB.CheckBox preliminar 
      Caption         =   " Pre"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox Text 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Text            =   "1"
      Top             =   120
      Width           =   975
   End
   Begin XPControls.XPButton Cancle 
      Height          =   435
      Left            =   4800
      TabIndex        =   16
      Top             =   720
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "FrmImp.frx":0000
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
   Begin XPControls.XPButton ok 
      Height          =   435
      Left            =   4800
      TabIndex        =   17
      Top             =   120
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "FrmImp.frx":059A
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
   Begin XPControls.XPButton CmdescPF 
      Height          =   375
      Left            =   4080
      TabIndex        =   18
      Top             =   120
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Picture         =   "FrmImp.frx":0B34
      Caption         =   ""
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
   Begin VBCCR17.SpinBox mes 
      Height          =   615
      Left            =   2520
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   2280
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   12648447
      Enabled         =   0   'False
      Min             =   1
      Max             =   3
      Value           =   1
   End
   Begin VB.Label LABEL 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   372
      Left            =   120
      TabIndex        =   2
      Top             =   3360
      Width           =   4692
   End
   Begin VB.Label Label1 
      Caption         =   "Digite o Numero do Processo Origem"
      ForeColor       =   &H00C00000&
      Height          =   252
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2772
   End
End
Attribute VB_Name = "FrmImp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cARQPF As String
Dim cARQPFP As String
Dim cARQPFG As String
Dim cARQFEMEA As String
Dim cARQFEMEAPF As String
Dim lFEMEA As Boolean
Dim cARQDES As String

Private Sub Cancle_Click()
    lRETU = False
    eRETU01 = 0
    Unload Me
End Sub

Private Sub impX()
    Dim nPFORI As Long
    On Error Resume Next
    Dim cORIGEM As String
    Dim cSQLORI As String
    Dim cDUPSQL As String
    Dim cDESARQ As String
    Dim cDESSQL As String

    nPFORI = Val(Text)
    lRETU = True
    eRETU01 = nPFORI
    
    'DUPLICACAO
    cDUPSQL = "select * from DUPLICAR WHERE TABELA='" & cARQIMP & "'"
    If cARQIMP = "FEMAVU" Or cARQIMP = "FEMADC" Or cARQIMP = "FEMMAQ" Or cARQIMP = "FEMAX" Or cARQIMP = "FEMEAPRE" Or cARQIMP = "FEMEAGP12" Then
        cDUPSQL = "select * from DUPLICAR WHERE TABELA='FEMEA'"
    End If
    
    'DESTINO
    cDESARQ = Sdb
    Select Case cARQIMP
    Case "FEMAVU", "FEMAX", "FEMADC", "FEMMAQ"
        cDESARQ = cARQFEMEA
    Case "FEMEAPRE"
        cDESARQ = cARQPFP
    Case "FEMEAGP12"
        cDESARQ = cARQPFG
    End Select
    cDESSQL = cARQIMP
    If cARQIMP = "FEMAX" Or cARQIMP = "FEMEAPRE" Or cARQIMP = "FEMEAGP12" Then
        cDESSQL = "FEMEA"
    End If

    
    
    
    'origem duplicaçao
    cORIGEM = cARQPF
    cSQLORI = "select * from " & cARQIMP & " WHERE pf=" & nPFORI
    If lFEMEA Then
        If FemeaC.Value = vbChecked Then
            cSQLORI = "select * from FEMAVU WHERE pf=" & nPFORI
            cORIGEM = cARQFEMEA
        End If
        If FemeaA.Value = vbChecked Then
            cSQLORI = "select * from FEMADC WHERE pf=" & nPFORI
            cORIGEM = cARQFEMEA
        End If
        If FemeaM.Value = vbChecked Then
            cSQLORI = "select * from FEMMAQ WHERE pf=" & nPFORI
            cORIGEM = cARQFEMEA
        End If
        If FemeaP.Value = vbChecked Then
            cORIGEM = cARQFEMEA
            cSQLORI = "select * from FEMEAPAD WHERE pf=" & nPFORI 'FEMEA
        End If
        If FemeaN.Value = vbChecked Then
            cORIGEM = cARQFEMEAPF                'cARQPF
            cSQLORI = "select * from FEMEA WHERE pf=" & nPFORI
        End If
        If FemeaR.Value = vbChecked Then
            cORIGEM = cARQPFP
            cSQLORI = "select * from FEMEA WHERE pf=" & nPFORI
        End If
        If FemeaG.Value = vbChecked Then
            cORIGEM = cARQPFG
            cSQLORI = "select * from FEMEA WHERE pf=" & nPFORI
        End If
    Else
        If preliminar.Value = vbChecked Then
            cORIGEM = cARQPFP
        End If
        If gp12.Value = vbChecked Then
            cORIGEM = cARQPFG
        End If
        If Normal.Value = vbChecked Then
            cORIGEM = cARQPF
        End If
        If Recebimento.Value = Checked Then
            Select Case mes
            Case 1
                cSQLORI = "select * from PFCO WHERE PF=" & nPFORI & " AND SEQ=10 AND SSQ=10"
            Case 2
                cSQLORI = "select * from PFCO WHERE PF=" & nPFORI & " AND SEQ=20 AND SSQ=10"
            Case 3
                cSQLORI = "select * from PFCO WHERE PF=" & nPFORI & " AND SEQ=30 AND SSQ=10"
            End Select
        End If
        If laboratorio.Value = Checked Then
            Select Case mes
            Case 1
                cSQLORI = "select * from PFCO WHERE PF=" & nPFORI & " AND SEQ=10 AND SSQ=20"
            Case 2
                cSQLORI = "select * from PFCO WHERE PF=" & nPFORI & " AND SEQ=20 AND SSQ=20"
            Case 3
                cSQLORI = "select * from PFCO WHERE PF=" & nPFORI & " AND SEQ=30 AND SSQ=20"
            End Select
        End If
        If Final.Value = Checked Then
            cSQLORI = "select * from PFCO WHERE PF=" & nPFORI & " AND SEQ=99 AND SSQ=99"
        End If
    End If
    
    If cARQIMP = "FEMAX" Then
        importa2 cARQPF, cDUPSQL, cORIGEM, cSQLORI, cDESARQ, cDESSQL, "FEMAX"
    Else
        importa2 cARQPF, cDUPSQL, cORIGEM, cSQLORI, cDESARQ, cDESSQL, ""
    End If
     
   

End Sub

Private Sub CmdescPF_Click()
    escpf.Show vbModal, Me
    If lRETU Then
        Text.Text = eRETU01
    End If
End Sub

Private Sub FemeaC_Click()
    FemeaN.Value = False
    FemeaP.Value = False
    FemeaR.Value = False
    FemeaG.Value = False
End Sub

Private Sub FemeaG_Click()
    FemeaC.Value = False
    FemeaP.Value = False
    FemeaR.Value = False
    FemeaN.Value = False
End Sub

Private Sub FemeaN_Click()
    FemeaC.Value = False
    FemeaP.Value = False
    FemeaR.Value = False
    FemeaG.Value = False
End Sub

Private Sub FemeaP_Click()
    FemeaN.Value = False
    FemeaC.Value = False
    FemeaR.Value = False
    FemeaG.Value = False
End Sub

Private Sub FemeaR_Click()
    FemeaN.Value = False
    FemeaC.Value = False
    FemeaP.Value = False
    FemeaG.Value = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    TeclaEnter KeyCode
End Sub

Private Sub Form_Load()
    Center Me
    cARQFEMEA = PegPath("PATH", "FEMEA")
    cARQFEMEAPF = PegPath("PATH", "FEMEAPF")
    cARQPF = PegPath("PATH", "PF")
    cARQPFP = PegPath("PATH", "PFP")
    cARQPFG = PegPath("PATH", "PFG")
    If cARQIMP = "FEMADC" Or cARQIMP = "FEMMAQ" Or cARQIMP = "FEMEA" Or cARQIMP = "FEMAVU" Or cARQIMP = "FEMAX" Or cARQIMP = "FEMEAPRE" Or cARQIMP = "FEMEAGP12" Then
        lFEMEA = True
        Select Case cARQIMP
        Case "FEMEA"
            cARQDES = cARQFEMEA                  'cARQPF
        Case "FEMEAPRE"
            cARQDES = cARQPFP
        Case "FEMEAGP12"
            cARQDES = cARQPFG
        Case "FEMAVU", "FEMMAQ"
            cARQDES = cARQFEMEA
        Case "FEMADC"
            cARQDES = cARQFEMEA
        Case "FEMAX"
            cARQDES = cARQFEMEA
        End Select
    End If
    
    If cARQIMP = "PFS" Then
        'preliminar.Visible = False
        'gp12.Visible = False
        'Normal.Visible = False
        Recebimento.Visible = False
        laboratorio.Visible = False
        Final.Visible = False
        mes.Visible = False
       ' maismes.Visible = False
       ' menosmes.Visible = False
    End If

    
    If lFEMEA Then
        preliminar.Visible = False
        gp12.Visible = False
        Normal.Visible = False
        Recebimento.Visible = False
        laboratorio.Visible = False
        Final.Visible = False
        mes.Visible = False
       ' maismes.Visible = False
       ' menosmes.Visible = False
    Else
        FemeaA.Visible = False
        FemeaP.Visible = False
        FemeaN.Visible = False
        FemeaC.Visible = False
        FemeaG.Visible = False
        FemeaR.Visible = False
        FemeaM.Visible = False
    End If
End Sub

Private Sub gp12_Click()
    Normal.Value = False
    preliminar.Value = False
End Sub

'Private Sub maismes_Click()
 '   mes = CStr(IncDec(mes, 1, 1, 3))
'End Sub

'Private Sub menosmes_Click()
'    mes = CStr(IncDec(mes, -1, 1, 3))
'End Sub

Private Sub Normal_Click()
    gp12.Value = False
    preliminar.Value = False
End Sub

Private Sub ok_Click()
    impX
    If cARQIMP = "PFS" Then
        cARQIMP = "PFI"
        impX
        cARQIMP = "PFQSBLEP"
        impX
    End If
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault
End Sub

Private Sub preliminar_Click()
    gp12.Value = False
    Normal.Value = False
End Sub

Private Sub Text_GotFocus()
    FocusMe
End Sub

Private Sub tEXT_KeyPress(KeyAscii As Integer)
    KeyAscii = ValiText(KeyAscii, "#NI")
End Sub

