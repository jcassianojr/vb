VERSION 5.00
Object = "{F62B9FA4-455F-4FE3-8A2D-205E4F0BCAFB}#11.5#0"; "CRViewer.dll"
Begin VB.Form FrmCrwENGViewer 
   Caption         =   "Visualizador de Relatorios"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14700
   Icon            =   "FrmCrwENGViewer.frx":0000
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
Option Explicit
' Declaração dos objetos no escopo global do Form para que fiquem acessíveis
' caso você crie botões de exportação customizados posteriormente.
Private MyApp As CRAXDRT.Application
Private MyRpt As CRAXDRT.Report

Private Sub Form_Load()
    On Error GoTo ErroLoad11
    
    ' ==========================================================================
    ' 1. DIMENSIONAMENTO E POSICIONAMENTO DA TELA
    ' ==========================================================================
    Me.Height = Screen.Height - 1200
    Me.Width = Screen.Width - 1200
    Me.Top = 100
    Me.Left = 100
    
    ' ==========================================================================
    ' 2. INICIALIZAÇÃO DA ENGINE DO CRYSTAL 11.5 (RDC)
    ' ==========================================================================
    ' Instancia o aplicativo do Crystal em tempo de execução
    Set MyApp = New CRAXDRT.Application
    
    ' Abre o arquivo .rpt contido na sua variável global de transporte (cARQRTF)
    ' O parâmetro 1 indica abertura normal de leitura
    Set MyRpt = MyApp.OpenReport(cARQRTF, 1)
    
    ' ==========================================================================
    ' PASSO A: APLICA O FILTRO (SELECTION FORMULA) -> PRIMEIRO!
    ' ==========================================================================
    If Len(Trim(cFILTRO)) > 0 Then
        MyRpt.RecordSelectionFormula = cFILTRO
    End If
    
  ' --------------------------------------------------------------------------
    ' [NOVA LOGICA CRYSTAL 11]: SUBSTITUIÇÃO DE ARQUIVOS LOCAIS (.MDB, .DBF, ETC)
    ' --------------------------------------------------------------------------
    Dim nARQUIVOS As Long
    Dim x         As Long
    Dim oTable    As CRAXDRT.DatabaseTable
    
    nARQUIVOS = UBound(aARQUIVOS)
    
    ' Loop pelos arquivos passados pelo seu sistema
    For x = 0 To nARQUIVOS
        If Len(aARQUIVOS(x)) > 0 Then
            ' [BLINDAGEM] Mantém a sua checagem de existência original
            If Not FileConnExist(aARQUIVOS(x)) Then
                Alert "Arquivo nao Encontrado: " + aARQUIVOS(x)
                Exit Sub
            End If
            
            ' No Crystal 11, nós precisamos mapear o índice (x) para a tabela correspondente do RPT.
            ' Adicionamos + 1 porque a coleção de tabelas do Crystal inicia em 1 (base 1)
            If (x + 1) <= MyRpt.Database.Tables.Count Then
                Set oTable = MyRpt.Database.Tables(x + 1)
                
                ' Atualiza a propriedade Location com o novo caminho físico do banco/tabela
                oTable.location = aARQUIVOS(x)
            End If
        End If
    Next x
    
    
    ' Associa o relatório processado como a fonte oficial do visualizador de tela
    CRViewer1.ReportSource = MyRpt

    ' ==========================================================================
    ' 3. AJUSTE DE REDIMENSIONAMENTO DO COMPONENTE EM TELA
    ' ==========================================================================
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Width = ScaleWidth
    CRViewer1.Height = ScaleHeight
    
    ' ==========================================================================
    ' 4. CONTROLE SEGURO DA BARRA DE FERRAMENTAS (DIREITOS E CONFIGURAÇÕES)
    ' ==========================================================================
    With CRViewer1
        ' --- Recursos Fixos Operacionais (Sempre Ativos) ---
        .DisplayToolbar = -1                ' True (Exibe a barra superior)
        .EnableStopButton = -1              ' True (Permite parar carregamento travado)
        .EnableNavigationControls = -1      ' True (Setas de avançar/voltar página)
        .EnableZoomControl = -1             ' True (Controle de zoom da folha)
        .EnableRefreshButton = -1           ' True (Botão de recarregar dados)
        .EnablePopupMenu = -1               ' True (Menu de clique direito sobre as células)
        .LaunchHTTPHyperlinksInNewBrowser = -1 ' True (Links web abrem fora do sistema)
        .EnableLogonPrompts = -1            ' True (Permite tela de login caso o banco exija)
        .EnableInteractiveParameterPrompting = 0 ' False (Evita pedir parâmetros repetidos se já enviados)
        
        ' --- Segurança de Direitos por Usuário (aDIREITOS) ---
        ' Índice 6 controla se o operador pode imprimir relatórios no sistema
        .EnablePrintButton = aDIREITOS(6)
        
        ' Índice 7 controla se o operador pode exportar para PDF/Excel/Word
        .EnableExportButton = aDIREITOS(7)
        
        ' --- Configurações Específicas do Layout do Relatório (aRELCFG) ---
        ' Índice 4 define se exibe a árvore de navegação lateral (Group Tree)
        .EnableGroupTree = aRELCFG(4)
        .DisplayGroupTree = aRELCFG(4)
        
        ' Índice 5 define se ativa o botão de busca interna (Binóculo / Search Expert)
        .EnableSearchExpertButton = aRELCFG(5)
        
        ' Desativa recursos irrelevantes para o seu fluxo atual (Evita poluição visual)
        .EnableHelpButton = 0               ' False
    End With
    
    ' ==========================================================================
    ' 5. DISPARO FINAL DE RENDERIZAÇÃO
    ' ==========================================================================
    ' Dá a ordem para o componente processar e desenhar as páginas na tela do usuário
    CRViewer1.ViewReport
    
    Exit Sub

ErroLoad11:
    MsgBox "Falha crítica ao inicializar visualizador Crystal XI: " & Err.Description, vbCritical, "Erro de Inicialização"
End Sub

Private Sub Form_Resize()
    ' Bloco de segurança para que o componente acompanhe o redimensionamento
    ' da janela caso o usuário maximize ou arraste as bordas do formulário
    On Error Resume Next
    CRViewer1.Width = ScaleWidth
    CRViewer1.Height = ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Limpeza estrita da memória ao fechar a tela para evitar travar processos
    ' do Crystal Reports em background no Gerenciador de Tarefas do Windows
    On Error Resume Next
    Set MyRpt = Nothing
    Set MyApp = Nothing
End Sub


