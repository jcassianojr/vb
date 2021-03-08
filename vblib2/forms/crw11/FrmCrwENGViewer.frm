VERSION 5.00
Object = "{F62B9FA4-455F-4FE3-8A2D-205E4F0BCAFB}#11.5#0"; "CRViewer.dll"
Begin VB.Form FrmCrwENGViewer 
   Caption         =   "Visualizador de Relatorios"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14700
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   14700
   StartUpPosition =   3  'Windows Default
   Begin CrystalActiveXReportViewerLib11_5Ctl.CrystalActiveXReportViewer CRViewer1 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   14295
      _cx             =   25215
      _cy             =   1085
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   0   'False
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
      LocaleID        =   1046
      EnableInteractiveParameterPrompting=   0   'False
   End
End
Attribute VB_Name = "FrmCrwENGViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim MyApp As New CRAXDRT.Application
    Dim MyRpt As New CRAXDRT.Report

    'Center Me
    Me.Height = Screen.Height - 1200
    Me.Width = Screen.Width - 1200
    Me.Top = 100
    Me.Left = 100



    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight - 200
    CRViewer1.Width = ScaleWidth - 200
    CRViewer1.Visible = True

    CRViewer1.EnablePrintButton = aDIREITOS(6)
    CRViewer1.EnableGroupTree = aRELCFG(4)
    CRViewer1.EnableSearchControl = aRELCFG(5)
    CRViewer1.EnableExportButton = aDIREITOS(7)



    Set MyRpt = MyApp.OpenReport(cARQRTF, 1)
 

    CRViewer1.ReportSource = MyRpt
    CRViewer1.ViewReport


End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0

    CRViewer1.Left = 0

    CRViewer1.Height = ScaleHeight

    CRViewer1.Width = ScaleWidth
End Sub

