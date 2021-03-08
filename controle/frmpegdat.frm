VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmPegDat 
   Caption         =   "Confirme Data"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   960
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker Data 
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   24510465
      CurrentDate     =   37334
   End
End
Attribute VB_Name = "FrmPegDat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
    eRETU01 = Data
    Unload Me
End Sub

Private Sub Form_Load()
    Data = date
End Sub
Private Sub Form_Unload(Cancel As Integer)
 Screen.MousePointer = vbDefault

End Sub
