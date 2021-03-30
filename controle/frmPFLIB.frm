VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{7020C36F-09FC-41FE-B822-CDE6FBB321EB}#1.0#0"; "vbccr17.ocx"
Begin VB.Form frmPFLIB 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PF"
   ClientHeight    =   2175
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   7080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   7080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VBCCR17.DTPicker DTPicker1 
      Height          =   495
      Left            =   2760
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1320
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Enabled         =   0   'False
      Value           =   44285
   End
   Begin VB.TextBox TxtRev 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   372
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
      Width           =   732
   End
   Begin VB.TextBox txtPF 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   372
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Width           =   732
   End
   Begin VB.TextBox txtFields 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   1080
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   720
      Width           =   3615
   End
   Begin VB.TextBox txtFields 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   720
      Width           =   735
   End
   Begin XPControls.XPButton Encerrar 
      Height          =   435
      Left            =   5040
      TabIndex        =   6
      Top             =   720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "frmPFLIB.frx":0000
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
      Left            =   5040
      TabIndex        =   7
      Top             =   240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "frmPFLIB.frx":059A
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
   Begin VB.Label Label1 
      Caption         =   " PF:"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
   Begin VB.Label CABECARIO 
      BackColor       =   &H00FFFFFF&
      Caption         =   "NORMAL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   2415
   End
End
Attribute VB_Name = "frmPFLIB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cTIPO

Private Sub cmdClose_Click()
    Dim cSQL
    Dim aFOR
    Dim aVAL
    Dim aCAM
    aFOR = Array("NI", "C", "D")
    aVAL = Array(txtFields(0), txtFields(1), DTPicker1)
    cSQL = "select PF,CODIGO,FEMEAF,FEMEAR,FEMEAD,PCLIBNUM,PCLIBNOM,PCLIBDATE,RESNUM,RESNOM,RESDAT,ELANUM,FEMEAEF,PCELANUM,PRDNUM,PRDNOM,PRDDAT,SEGNUM,SEGNOM,SEGDAT from PF WHERE PF=" & nPF
    If cTIPO = "PF" Then
        cSQL = "select PF,RESNUM,RESNOM,RESDAT from PF WHERE PF=" & nPF
        aCAM = Array("RESNUM", "RESNOM", "RESDAT")
    End If
    
    If cTIPO = "PFSEG" Then                      ' aprovador seguranca do trabalho
        cSQL = "select PF,SEGNUM,SEGNOM,SEGDAT from PF WHERE PF=" & nPF
        aCAM = Array("RESNUM", "RESNOM", "RESDAT")
    End If
    If cTIPO = "PFPRO" Then                      ' aprovador producao
        cSQL = "select PF,PRDNUM,PRDNOM,PRDDAT from PF WHERE PF=" & nPF
        aCAM = Array("RESNUM", "RESNOM", "RESDAT")
    End If
    
    If cTIPO = "FEMEA" Then
        cSQL = "select PF,FEMEAF,FEMEAR,FEMEAD from PF WHERE PF=" & nPF
        aCAM = Array("FEMEAF", "FEMEAR", "FEMEAD")
    End If
    If cTIPO = "PC" Then
        cSQL = "select PF,PCLIBNUM,PCLIBNOM,PCLIBDATE from PF WHERE PF=" & nPF
        aCAM = Array("PCLIBNUM", "PCLIBNOM", "PCLIBDATE")
    End If
    
    If cTIPO = "POA" Then
        cSQL = "select NUMERO,RESNUM,RESNOM,RESDAT from POA WHERE NUMERO=" & nPPAP
        aCAM = Array("RESNUM", "RESNOM", "RESDAT")
    End If
    
    If cTIPO = "POKA" Then
        cSQL = "select NUMERO,RESNUM,RESNOM,RESDAT from POKA WHERE NUMERO=" & nPPAP
        aCAM = Array("RESNUM", "RESNOM", "RESDAT")
    End If

    If cTIPO = "DISPO" Then
        cSQL = "select NUMERO,RESNUM,RESNOM,RESDAT from DISPO WHERE NUMERO=" & nPPAP
        aCAM = Array("RESNUM", "RESNOM", "RESDAT")
    End If

    
    GrvSQL Sdb, cSQL, 3, aCAM, aVAL, aFOR
    Unload Me
End Sub

Private Sub Encerrar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Center Me
    cTIPO = ePASS01
    CABECARIO.Caption = ""
    If InStr(Sdb, "PFP.MDB") Then
        ''Cabeçario
        CABECARIO.Caption = "PRELIMINAR"
    End If
    If InStr(Sdb, "PFG.MDB") Then
        ''Cabeçario
        CABECARIO.Caption = "GP12"
    End If
    frmPFLIB.txtFields(0) = zIDFOLHA
    frmPFLIB.txtFields(1) = zNOMEFOLHA
    DTPicker1.Value = Date
    TXTPF.Text = nPF
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault
End Sub

