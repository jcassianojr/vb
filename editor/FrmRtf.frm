VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{7020C36F-09FC-41FE-B822-CDE6FBB321EB}#1.2#0"; "vbccr17.ocx"
Object = "{451B73A5-1563-45D5-A6AC-7B2B7D30B778}#1.0#0"; "BSPrin10.ocx"
Begin VB.Form FrmRTf 
   Caption         =   "Editor Textos"
   ClientHeight    =   7245
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11715
   Icon            =   "FrmRtf.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7245
   ScaleWidth      =   11715
   StartUpPosition =   2  'CenterScreen
   Begin BSPrinter.PrintPreview PrintPreview1 
      Left            =   11160
      Top             =   2040
      _ExtentX        =   1191
      _ExtentY        =   1191
      LcK1b           =   "yefT59bnyufI583n1ufV59HnyefN58nn0+fK58nnzefR58rn1+fW5w=="
      LcK2b           =   "reeI58zn+ue059jn1Of656fni+eN543nleeM54HnlueR5w=="
      AmbientBb       =   $"FrmRtf.frx":058A
   End
   Begin VBCCR17.RichTextBox RichTextbox1 
      Height          =   6015
      Left            =   0
      TabIndex        =   42
      Top             =   840
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   10610
      MultiLine       =   -1  'True
      TextRTF         =   "FrmRtf.frx":062A
   End
   Begin VBCCR17.FontCombo FontCombo1 
      Height          =   315
      Left            =   2400
      TabIndex        =   41
      Top             =   360
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   556
      Text            =   "FrmRtf.frx":0790
      ExtendedUI      =   -1  'True
   End
   Begin VB.CommandButton CmdExpHtml 
      Height          =   315
      Left            =   1620
      Picture         =   "FrmRtf.frx":07C4
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   "Exportar Documento como HTML"
      Top             =   0
      Width           =   360
   End
   Begin VB.CommandButton CmdExpPdf 
      Height          =   315
      Left            =   1320
      Picture         =   "FrmRtf.frx":0D4E
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "Exportar Documento como PDF"
      Top             =   0
      Width           =   360
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   38
      Top             =   6930
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            Object.Width           =   1588
            MinWidth        =   1587
            Picture         =   "FrmRtf.frx":12D8
            TextSave        =   "21:28"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            Object.Width           =   2302
            MinWidth        =   2293
            Picture         =   "FrmRtf.frx":1872
            TextSave        =   "24/01/2023"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Picture         =   "FrmRtf.frx":1E0C
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10372
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgMenu 
      Left            =   11190
      Top             =   1020
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   42
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":21A6
            Key             =   "center"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":2540
            Key             =   "justificado"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":28DA
            Key             =   "left"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":2C74
            Key             =   "right"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":300E
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":33A8
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":3942
            Key             =   "cut"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":3CDC
            Key             =   "diahora"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":4076
            Key             =   "editor"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":4610
            Key             =   "find"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":49AA
            Key             =   "marcado"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":4B04
            Key             =   "letra01"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":4E9E
            Key             =   "tamanho"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":5238
            Key             =   "paragro"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":55D2
            Key             =   "decrese"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":596C
            Key             =   "incrise"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":5D06
            Key             =   "imagem"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":60A0
            Key             =   "italico"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":643A
            Key             =   "s_Key19"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":67D4
            Key             =   "mail"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":6D6E
            Key             =   "new"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":7108
            Key             =   "open"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":76A2
            Key             =   "xxx"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":7A3C
            Key             =   "imagem2"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":7DD6
            Key             =   "colar"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":8370
            Key             =   "printer"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":890A
            Key             =   "printcfg"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":8CA4
            Key             =   "sair"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":903E
            Key             =   "salvar"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":93D8
            Key             =   "busca"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":9772
            Key             =   "proximabusca"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":9B0C
            Key             =   "letra"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":9EA6
            Key             =   "abc"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":A240
            Key             =   "riscado"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":A5DA
            Key             =   "sub"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":A974
            Key             =   "sob"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":AD0E
            Key             =   "minusculas"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":B0A8
            Key             =   "maisculas"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":B442
            Key             =   "underline"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":B7DC
            Key             =   "refazer"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":BB76
            Key             =   "info"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":C110
            Key             =   "exit"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command34 
      Height          =   315
      Left            =   8760
      Picture         =   "FrmRtf.frx":C6AA
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Capitalizar"
      Top             =   360
      Width           =   360
   End
   Begin VB.CommandButton Command33 
      Height          =   315
      Left            =   3120
      Picture         =   "FrmRtf.frx":CA34
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Enviar Email  Documento Atual"
      Top             =   0
      Width           =   360
   End
   Begin VB.CommandButton Command32 
      Height          =   315
      Left            =   10320
      Picture         =   "FrmRtf.frx":CFBE
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Inserir Data"
      Top             =   360
      Width           =   360
   End
   Begin VB.CommandButton Command31 
      Height          =   315
      Left            =   9960
      Picture         =   "FrmRtf.frx":D348
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Incluir Imagens"
      Top             =   360
      Width           =   360
   End
   Begin VB.CommandButton Command30 
      Height          =   315
      Left            =   4560
      Picture         =   "FrmRtf.frx":D6D2
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Selecionar Todo o Texto"
      Top             =   0
      Width           =   360
   End
   Begin VB.CommandButton Command29 
      Height          =   315
      Left            =   9120
      Picture         =   "FrmRtf.frx":DA5C
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Mudar Letra Escrita Inferiores"
      Top             =   360
      Width           =   360
   End
   Begin VB.CommandButton Command28 
      Height          =   315
      Left            =   9600
      Picture         =   "FrmRtf.frx":DDE6
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Mudar Letra Escrita Superiores"
      Top             =   360
      Width           =   360
   End
   Begin VB.CommandButton Command27 
      Height          =   315
      Left            =   8400
      Picture         =   "FrmRtf.frx":E170
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Converter Maisculas"
      Top             =   360
      Width           =   360
   End
   Begin VB.CommandButton Command26 
      Height          =   315
      Left            =   8040
      Picture         =   "FrmRtf.frx":E4FA
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Converter Minusculas"
      Top             =   360
      Width           =   360
   End
   Begin VB.CommandButton Command25 
      Height          =   315
      Left            =   9120
      Picture         =   "FrmRtf.frx":E884
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Diminuir Recuo"
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton Command24 
      Height          =   315
      Left            =   8760
      Picture         =   "FrmRtf.frx":EC0E
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Aumentar Recuo"
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton Command23 
      Height          =   315
      Left            =   8400
      Picture         =   "FrmRtf.frx":EF98
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Marcadores"
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton Command22 
      Height          =   315
      Left            =   1920
      Picture         =   "FrmRtf.frx":F322
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Selecionar Fonte"
      Top             =   360
      Width           =   360
   End
   Begin VB.CommandButton Command21 
      Height          =   315
      Left            =   1560
      Picture         =   "FrmRtf.frx":F46C
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Cortado"
      Top             =   360
      Width           =   360
   End
   Begin VB.CommandButton Command20 
      Height          =   315
      Left            =   7800
      Picture         =   "FrmRtf.frx":F7F6
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Localizar Proxima"
      Top             =   0
      Width           =   360
   End
   Begin VB.TextBox texto 
      Height          =   285
      Left            =   5400
      TabIndex        =   22
      Tag             =   "Texto"
      ToolTipText     =   "Texto Para Localizar"
      Top             =   0
      Width           =   1995
   End
   Begin VB.CommandButton Command19 
      Height          =   315
      Left            =   7440
      Picture         =   "FrmRtf.frx":FB80
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Localizar"
      Top             =   0
      Width           =   360
   End
   Begin VB.CommandButton Command18 
      Height          =   315
      Left            =   3840
      Picture         =   "FrmRtf.frx":FF0A
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Copiar Texto Selecionado"
      Top             =   0
      Width           =   360
   End
   Begin VB.CommandButton Command17 
      Height          =   315
      Left            =   3480
      Picture         =   "FrmRtf.frx":10494
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Recortar Texto Selecionado"
      Top             =   0
      Width           =   360
   End
   Begin VB.CommandButton Command16 
      Height          =   315
      Left            =   4200
      Picture         =   "FrmRtf.frx":1081E
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Colar Texto "
      Top             =   0
      Width           =   360
   End
   Begin VB.CommandButton Command15 
      Height          =   315
      Left            =   7560
      Picture         =   "FrmRtf.frx":10DA8
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Escolher Cor"
      Top             =   360
      Width           =   495
   End
   Begin VB.CommandButton Command14 
      Height          =   315
      Left            =   11160
      Picture         =   "FrmRtf.frx":11132
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Sair"
      Top             =   360
      Width           =   495
   End
   Begin VB.CommandButton Command13 
      Height          =   315
      Left            =   2760
      Picture         =   "FrmRtf.frx":116BC
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Configurar Impressoras"
      Top             =   0
      Width           =   360
   End
   Begin VB.CommandButton Command12 
      Height          =   315
      Left            =   10320
      Picture         =   "FrmRtf.frx":11A46
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Alinhar a Direita"
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton Command11 
      Height          =   315
      Left            =   9960
      Picture         =   "FrmRtf.frx":11DD0
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Centralizar"
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton Command10 
      Height          =   315
      Left            =   9600
      Picture         =   "FrmRtf.frx":1215A
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Alinhar a Esquerda"
      Top             =   0
      Width           =   360
   End
   Begin VB.CommandButton Command9 
      Height          =   315
      Left            =   1200
      Picture         =   "FrmRtf.frx":124E4
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Sublinhado"
      Top             =   360
      Width           =   360
   End
   Begin VB.CommandButton Command8 
      Height          =   315
      Left            =   840
      Picture         =   "FrmRtf.frx":1286E
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Italico"
      Top             =   360
      Width           =   360
   End
   Begin VB.CommandButton Command7 
      Height          =   315
      Left            =   480
      Picture         =   "FrmRtf.frx":12BF8
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Negrito/Bold"
      Top             =   360
      Width           =   360
   End
   Begin VB.CommandButton Command6 
      Height          =   315
      Left            =   0
      Picture         =   "FrmRtf.frx":12F82
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Checar Ortografia"
      Top             =   360
      Width           =   360
   End
   Begin VB.CommandButton Command5 
      Height          =   315
      Left            =   4920
      Picture         =   "FrmRtf.frx":1330C
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Desfazer"
      Top             =   0
      Width           =   360
   End
   Begin VB.CommandButton Command4 
      Height          =   315
      Left            =   2100
      Picture         =   "FrmRtf.frx":13696
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Imprimir Documento Atual"
      Top             =   0
      Width           =   360
   End
   Begin VB.CommandButton Command3 
      Height          =   315
      Left            =   960
      Picture         =   "FrmRtf.frx":13C20
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salvar Documento Atual"
      Top             =   0
      Width           =   360
   End
   Begin VB.CommandButton Command2 
      Height          =   315
      Left            =   480
      Picture         =   "FrmRtf.frx":13FAA
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Abrir Documento Existente"
      Top             =   0
      Width           =   360
   End
   Begin VB.CommandButton Command1 
      Height          =   315
      Left            =   0
      Picture         =   "FrmRtf.frx":14534
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Iniciar Documento Em Branco"
      Top             =   0
      Width           =   360
   End
   Begin VB.ComboBox cmbfontcolor 
      Height          =   315
      Left            =   5760
      TabIndex        =   2
      Text            =   "Combo1"
      ToolTipText     =   "Escolher Cor"
      Top             =   360
      Width           =   1695
   End
   Begin VB.ComboBox cmbfontsize 
      Height          =   315
      Left            =   4800
      TabIndex        =   1
      Text            =   "Combo1"
      ToolTipText     =   "Tamanho da Letra"
      Top             =   360
      Width           =   915
   End
   Begin VB.ComboBox cmbfontname 
      Height          =   315
      Left            =   10920
      TabIndex        =   0
      Text            =   "Fonte de Letra"
      ToolTipText     =   "Fonte de Letra"
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   10680
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Menu Menu1 
      Caption         =   "&Arquivo"
      Tag             =   "&Arquivo|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      Begin VB.Menu mnunew 
         Caption         =   "&Novo"
         Shortcut        =   ^N
         Tag             =   "&Novo|#new|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0){^N}"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Abrir"
         Shortcut        =   ^A
         Tag             =   "&Abrir|#open|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0){^A}"
      End
      Begin VB.Menu mnusave 
         Caption         =   "&Salvar"
         Shortcut        =   ^S
         Tag             =   "&Salvar|#salvar|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0){^S}"
      End
      Begin VB.Menu mnusaveastxt 
         Caption         =   "Salvar Como &TXT"
         Tag             =   "Salvar Como &TXT|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu mnusaveasrtf 
         Caption         =   "Salvar Como &RTF"
         Tag             =   "Salvar Como &RTF|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu mnusaveasPDF 
         Caption         =   "Exportar PDF"
      End
      Begin VB.Menu mnusaveasHTML 
         Caption         =   "Exportar HTML"
      End
      Begin VB.Menu mnuprint 
         Caption         =   "&Imprimir"
         Shortcut        =   ^P
         Tag             =   "&Imprimir|#printer|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0){^P}"
      End
      Begin VB.Menu mnuprintcfg 
         Caption         =   "&Configurar Impressora"
         Tag             =   "&Configurar Impressora|#printcfg|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu mnusendmail 
         Caption         =   "&Enviar Email"
         Tag             =   "&Enviar Email|#mail|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "Sa&ir"
         Tag             =   "Sa&ir|#exit|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
   End
   Begin VB.Menu Menu2 
      Caption         =   "&Editar"
      Tag             =   "&Editar|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      Begin VB.Menu mnuundo 
         Caption         =   "&Desfazer"
         Shortcut        =   ^Z
         Tag             =   "&Desfazer|#refazer|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0){^Z}"
      End
      Begin VB.Menu tr01 
         Caption         =   "-"
         Tag             =   "-|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu mnuselectall 
         Caption         =   "Selecionar &Tudo"
         Shortcut        =   ^T
         Tag             =   "Selecionar &Tudo|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0){^T}"
      End
      Begin VB.Menu mnuselectline 
         Caption         =   "Selecionar Linha"
         Tag             =   "Selecionar Linha|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu mnuselword 
         Caption         =   "Selecionar Palavra"
         Tag             =   "Selecionar Palavra|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu mnuselsen 
         Caption         =   "Selecionar Sentençao"
         Tag             =   "Selecionar Sentençao|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu mnuDelSel 
         Caption         =   "Deletar Seleçao"
         Tag             =   "Deletar Seleçao|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu tr02 
         Caption         =   "-"
         Tag             =   "-|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu mnucopy 
         Caption         =   "&Copiar"
         Shortcut        =   ^C
         Tag             =   "&Copiar|#copy|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0){^C}"
      End
      Begin VB.Menu mnupaste 
         Caption         =   "Co&lar"
         Shortcut        =   ^V
         Tag             =   "Co&lar|#colar|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0){^V}"
      End
      Begin VB.Menu mnucut 
         Caption         =   "&Recortar"
         Shortcut        =   ^X
         Tag             =   "&Recortar|#cut|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0){^X}"
      End
      Begin VB.Menu tr03 
         Caption         =   "-"
         Tag             =   "-|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu mnufind 
         Caption         =   "Locali&zar"
         Shortcut        =   ^F
         Tag             =   "Locali&zar|#busca|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0){^F}"
      End
      Begin VB.Menu mnufindNext 
         Caption         =   "Localizar Pro&xima"
         Shortcut        =   {F3}
         Tag             =   "Localizar Pro&xima|#proximabusca|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0){{F3}}"
      End
      Begin VB.Menu mnufindrep 
         Caption         =   "Localizar e Substituir"
         Tag             =   "Localizar e Substituir|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu tr04 
         Caption         =   "-"
         Tag             =   "-|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu mnuinsdate 
         Caption         =   "Inserir Data"
         Tag             =   "Inserir Data|#diahora|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu mnuinsimg 
         Caption         =   "Inserir Imagem"
         Tag             =   "Inserir Imagem|#imagem|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu mnuEditarInserirCaracter 
         Caption         =   "Inserir Caracter"
         Tag             =   "Inserir Caracter|(Checked=0)(Enabled=0)(Visible=0)(WindowList=0)"
      End
   End
   Begin VB.Menu Menu3 
      Caption         =   "Format"
      Tag             =   "Format|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      Begin VB.Menu mnutext 
         Caption         =   "Alinhamento"
         Tag             =   "Alinhamento|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
         Begin VB.Menu mnuleft 
            Caption         =   "Esquerada"
            Tag             =   "Esquerada|#left|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
         End
         Begin VB.Menu mnuright 
            Caption         =   "Direita"
            Tag             =   "Direita|#right|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
         End
         Begin VB.Menu mnucenter 
            Caption         =   "Centralizado"
            Tag             =   "Centralizado|#center|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
         End
      End
      Begin VB.Menu mnucolor 
         Caption         =   "Cores"
         Tag             =   "Cores|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
         Begin VB.Menu mnucustomcolors 
            Caption         =   "Escolher Cor"
            Tag             =   "Escolher Cor|#imagem2|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
         End
      End
      Begin VB.Menu mnufonte 
         Caption         =   "Fontes Letra"
         Tag             =   "Fontes Letra|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
         Begin VB.Menu mnuFntesc 
            Caption         =   "Fonte de Letra"
            Tag             =   "Fonte de Letra|#letra|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
         End
         Begin VB.Menu mnubold 
            Caption         =   "Negrito"
            Tag             =   "Negrito|#Bold|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
         End
         Begin VB.Menu mnuitalico 
            Caption         =   "Italico"
            Tag             =   "Italico|#italico|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
         End
         Begin VB.Menu mnuunder 
            Caption         =   "Sublilhado"
            Tag             =   "Sublilhado|#underline|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
         End
         Begin VB.Menu mnustrik 
            Caption         =   "Riscado"
            Tag             =   "Riscado|#riscado|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
         End
      End
      Begin VB.Menu mnuVerFontes 
         Caption         =   "Ver Fontes"
         Tag             =   "Ver Fontes|#find|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu xxx 
         Caption         =   "Identicaçao"
         Tag             =   "Identicaçao|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
         Begin VB.Menu mnuIncInd 
            Caption         =   "Aumentar Posiçao"
            Tag             =   "Aumentar Posiçao|#imagem|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
         End
         Begin VB.Menu mnudecInd 
            Caption         =   "Diminuir Posiçao"
            Tag             =   "Diminuir Posiçao|#decrese|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
         End
      End
      Begin VB.Menu yyy 
         Caption         =   "Escrita"
         Tag             =   "Escrita|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
         Begin VB.Menu mnuscrsup 
            Caption         =   "Escrita Superior"
            Tag             =   "Escrita Superior|#sob|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
         End
         Begin VB.Menu mnuscrinf 
            Caption         =   "Escrita Inferior"
            Tag             =   "Escrita Inferior|#sub|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
         End
         Begin VB.Menu mnuscrnor 
            Caption         =   "Escrita Normal"
            Tag             =   "Escrita Normal|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
         End
      End
      Begin VB.Menu mnubullet 
         Caption         =   "Marcador"
         Tag             =   "Marcador|#xxx|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu mnucnv 
         Caption         =   "Converter "
         Tag             =   "Converter |(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
         Begin VB.Menu mnuUpper 
            Caption         =   "Maisculas"
            Tag             =   "Maisculas|#maisculas|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
         End
         Begin VB.Menu mnuDown 
            Caption         =   "Minusculas"
            Tag             =   "Minusculas|#minusculas|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
         End
         Begin VB.Menu mnucapitalizar 
            Caption         =   "Capitalizar"
            Tag             =   "Capitalizar|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
         End
         Begin VB.Menu mnutirace 
            Caption         =   "Tira Acentuaçao"
            Tag             =   "Tira Acentuaçao|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
         End
         Begin VB.Menu mnucnvoem 
            Caption         =   "Converter Ansi(Windows)-> Oem(DOS)"
            Tag             =   "Converter Ansi(Windows)-> Oem(DOS)|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
         End
         Begin VB.Menu mnucnvoem2 
            Caption         =   "Converter Ansi(Windows)-> Oem(DOS)Padrao2"
            Tag             =   "Converter Ansi(Windows)-> Oem(DOS)Padrao2|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
         End
         Begin VB.Menu mnucnvansi 
            Caption         =   "Converter Oem(DOS)->Ansi(Windows)"
            Tag             =   "Converter Oem(DOS)->Ansi(Windows)|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
         End
         Begin VB.Menu mnucnvansi2 
            Caption         =   "Converter Oem(DOS)->Ansi(Windows)Padrao2"
            Tag             =   "Converter Oem(DOS)->Ansi(Windows)Padrao2|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
         End
      End
   End
   Begin VB.Menu Menu4 
      Caption         =   "Aj&uda"
      Tag             =   "Aj&uda|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      Begin VB.Menu mnuabout 
         Caption         =   "&Sobre"
         Tag             =   "&Sobre|#info|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
   End
End
Attribute VB_Name = "FrmRTf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cARQTIPO As String
Const nFORMID = 1055
Const cFORMID = "Editor de Textos"
Private WithEvents MenuEvents As CEvents
Attribute MenuEvents.VB_VarHelpID = -1

'***************************************************************************
'*  Name         : ABRIR
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       :
'*  Date         : 20 dez 2003
'***************************************************************************

Private Sub ABRIR()
    On Error Resume Next
    If Len(cARQRTF) = 0 Then
        Exit Sub
    End If
    If Not FileExist(cARQRTF, True) Then
        Exit Sub
    End If
    If InStr(cARQRTF, "RTF") > 0 Then
        cARQTIPO = "RTF"
        RichTextBox1.LoadFile cARQRTF, rtfRTF
    Else
        cARQTIPO = "TXT"
        RichTextBox1.LoadFile cARQRTF, rtfText
    End If
    RichTextBox1.SetFocus
End Sub

'***************************************************************************
'*  Name         : achartf
'*  Description  :
'*  Parameters   : nPOS
'*  Returns      : Nothing
'*  Author       :
'*  Date         : 20 dez 2003
'***************************************************************************

Private Sub achartf(nPOS)
    Dim textfound As Long
    ''    RichTextbox1.Find (RichTextbox1.Text)
    RichTextBox1.SetFocus
    textfound = RichTextBox1.Find(texto, nPOS, , rtfWholeWord)
    If textfound = -1 Then
        Alert vbCr & "Texto Não Foi Encontrado"
    End If
    RichTextBox1.SetFocus
End Sub

'***************************************************************************
'*  Name         : cmbFontColor_Click
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       :
'*  Date         : 20 dez 2003
'***************************************************************************

Private Sub cmbFontColor_Click()

    '******************************************************************
    '* Change font colors of text using the
    '* Select Case statement with the ListIndex of the
    '* ComboBox control. Set the colors with
    '* the intrinsic constants for color.
    ''*   Me.Show
    '******************************************************************

    With RichTextBox1
        Select Case cmbfontcolor.ListIndex
        Case 0
            .SelColor = vbBlack
        Case 1
            .SelColor = vbBlue
        Case 2
            .SelColor = vbRed
        Case 3
            .SelColor = vbGreen
        End Select
    End With

    '******************************************************************
    '* Return focus to the RichTextbox control.
    ''*   RichTextbox1.SetFocus
    '******************************************************************

End Sub

'***************************************************************************
'*  Name         : cmbFontName_Click
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       :
'*  Date         : 20 dez 2003
'***************************************************************************

'Private Sub cmbfontname_Click()
'    RichTextbox1.SelFontName = cmbfontname
    ''RichTextbox1.SetFocus
'End Sub

'***************************************************************************
'*  Name         : cmbFontSize_Click
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       :
'*  Date         : 20 dez 2003
'***************************************************************************

Private Sub cmbFontSize_Click()
 '  RichTextbox1.SelFontSize = FontCombo1
  '  RichTextbox1.SetFocus
    
End Sub

Private Sub CmdExpHtml_Click()
    mnusaveasHTML_Click
End Sub

Private Sub CmdExpPdf_Click()
    mnusaveasPDF_Click
End Sub

'***************************************************************************
'*  Name         : Command10_Click
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       :
'*  Date         : 20 dez 2003
'***************************************************************************

Private Sub Command10_Click()
    mnuleft_Click
End Sub

'***************************************************************************
'*  Name         : Command11_Click
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       :
'*  Date         : 20 dez 2003
'***************************************************************************

Private Sub Command11_Click()
    mnucenter_Click
End Sub

'***************************************************************************
'*  Name         : Command12_Click
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       :
'*  Date         : 20 dez 2003
'***************************************************************************

Private Sub Command12_Click()
    mnuright_Click
End Sub

'***************************************************************************
'*  Name         : Command13_Click
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       :
'*  Date         : 20 dez 2003
'***************************************************************************

Private Sub Command13_Click()
    mnuprintcfg_Click
End Sub

'***************************************************************************
'*  Name         : Command14_Click
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       :
'*  Date         : 20 dez 2003
'***************************************************************************

Private Sub Command14_Click()
    mnuexit_Click
End Sub

'***************************************************************************
'*  Name         : Command15_Click
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       :
'*  Date         : 20 dez 2003
'***************************************************************************

Private Sub Command15_Click()
    mnucustomcolors_Click
End Sub

'***************************************************************************
'*  Name         : Command16_Click
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       :
'*  Date         : 20 dez 2003
'***************************************************************************

Private Sub Command16_Click()
    mnupaste_Click
End Sub

'***************************************************************************
'*  Name         : Command17_Click
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       :
'*  Date         : 20 dez 2003
'***************************************************************************

Private Sub Command17_Click()
    mnucut_Click
End Sub

'***************************************************************************
'*  Name         : Command18_Click
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       :
'*  Date         : 20 dez 2003
'***************************************************************************

Private Sub Command18_Click()
    mnucopy_Click
End Sub

'***************************************************************************
'*  Name         : Command19_Click
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       :
'*  Date         : 20 dez 2003
'***************************************************************************

Private Sub Command19_Click()
    achartf (0)
End Sub

'***************************************************************************
'*  Name         : Command1_Click
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       :
'*  Date         : 20 dez 2003
'***************************************************************************

Private Sub Command1_Click()
    mnunew_Click
End Sub

'***************************************************************************
'*  Name         : Command20_Click
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       :
'*  Date         : 20 dez 2003
'***************************************************************************

Private Sub Command20_Click()
    RichTextBox1.SetFocus
    achartf (RichTextBox1.SelStart + 1)
End Sub

'***************************************************************************
'*  Name         : Command21_Click
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       :
'*  Date         : 20 dez 2003
'***************************************************************************

Private Sub Command21_Click()
    RichTextBox1.SelStrikethru = Not RichTextBox1.SelStrikethru
    If RichTextBox1.SelStrikethru Then
        Command21.Appearance = 0
    Else
        Command21.Appearance = 1
    End If
End Sub

'***************************************************************************
'*  Name         : Command22_Click
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       :
'*  Date         : 20 dez 2003
'***************************************************************************

Private Sub Command22_Click()
    Dim oDLG As New clsCommonDialogs
    oDLG.FontFlags = BOTH + EFFECTS
    
    oDLG.Font = richtextbox1SelFontName
    oDLG.FontSize = richtextbox1SelFontSize
    ''oDLG.FontStrikeThru = richtextbox1SelStrikeThru
    ''oDLG.FontUnderline = richtextbox1SelUnderline
    oDLG.Color = richtextbox1SelColor
    oDLG.FontBold = richtextbox1SelBold
    oDLG.FontItalic = richtextbox1SelItalic
    oDLG.DialogFont
    
    
    
    With RichTextBox1
        .SelFontName = oDLG.Font
        .SelFontSize = oDLG.FontSize
        .SelBold = oDLG.FontBold
        .SelItalic = oDLG.FontItalic
        .SelStrikethru = oDLG.FontStrikeThru
        .SelUnderline = oDLG.FontUnderline
        .SelColor = oDLG.Color
    End With

End Sub

Private Sub Command23_Click()
    mnubullet_Click
End Sub

Private Sub Command24_Click()
    RichTextBox1.SelIndent = IncDec(RichTextBox1.SelIndent, 1, 0, 99)
End Sub

Private Sub Command25_Click()
    RichTextBox1.SelIndent = IncDec(RichTextBox1.SelIndent, -1, 0, 99)
End Sub

'***************************************************************************
'*  Name         : Command2_Click
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       :
'*  Date         : 20 dez 2003
'***************************************************************************

Private Sub Command2_Click()
    mnuOpen_Click
End Sub

Private Sub Command26_Click()
    RichTextBox1.SelText = LCase(RichTextBox1.SelText)
End Sub

Private Sub Command27_Click()
    RichTextBox1.SelText = UCase(RichTextBox1.SelText)
End Sub

Private Sub Command28_Click()
    mnuscrsup_Click
End Sub

Private Sub Command29_Click()
    mnuscrinf_Click
End Sub

'***************************************************************************
'*  Name         : Command3_Click
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       :
'*  Date         : 20 dez 2003
'***************************************************************************

Private Sub Command3_Click()
    mnusave_Click
End Sub

Private Sub Command30_Click()
    mnuselectall_Click
End Sub

Private Sub Command31_Click()
    mnuinsimg_Click
End Sub

Private Sub Command32_Click()
    mnuinsdate_Click
End Sub

Private Sub Command33_Click()
    mnusendmail_Click
End Sub

Private Sub Command34_Click()
    mnucapitalizar_Click
End Sub

'***************************************************************************
'*  Name         : Command4_Click
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       :
'*  Date         : 20 dez 2003
'***************************************************************************

Private Sub Command4_Click()
    mnuprint_Click
End Sub

'***************************************************************************
'*  Name         : Command5_Click
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       :
'*  Date         : 20 dez 2003
'***************************************************************************

Private Sub Command5_Click()
    mnuundo_Click
End Sub

'***************************************************************************
'*  Name         : Command7_Click
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       :
'*  Date         : 20 dez 2003
'***************************************************************************

Private Sub Command7_Click()
    RichTextBox1.SelBold = Not RichTextBox1.SelBold
    If RichTextBox1.SelBold Then
        Command7.Appearance = 0
    Else
        Command7.Appearance = 1
    End If
    RichTextBox1.SetFocus
End Sub

'***************************************************************************
'*  Name         : Command8_Click
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       :
'*  Date         : 20 dez 2003
'***************************************************************************

Private Sub Command8_Click()
    RichTextBox1.SelItalic = Not RichTextBox1.SelItalic
    If RichTextBox1.SelItalic Then
        Command8.Appearance = 0
    Else
        Command8.Appearance = 1
    End If
End Sub

'***************************************************************************
'*  Name         : Command9_Click
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       :
'*  Date         : 20 dez 2003
'***************************************************************************

Private Sub Command9_Click()
    RichTextBox1.SelUnderline = Not RichTextBox1.SelUnderline
    If RichTextBox1.SelUnderline = False Then
        Command9.Appearance = 0
    Else
        Command9.Appearance = 1
    End If
End Sub

Private Sub FontCombo1_Click()
    RichTextBox1.SelFontName = FontCombo1.Text
End Sub


'***************************************************************************
'*  Name         : Form_Load
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       :
'*  Date         : 20 dez 2003
'***************************************************************************


Private Sub Form_Load()
    '     CentralizaJanela Me
    CenterFormToScreen Me

'CommonDialogEx1.ShowColor

'CommonDialogEx1.ShowFolder

'CommonDialogEx1.ShowFont

'CommonDialogEx1.ShowHelp

'CommonDialogEx1.ShowOpen

'CommonDialogEx1.ShowPageSetup

'CommonDialogEx1.ShowPrinter'

'CommonDialogEx1.ShowSave


    '/ If gbSubClassMenu is False, the menu is not subclassed
    gbSubClassMenu = True

    If gbSubClassMenu Then SubClassMenuXP

    ''Configura Help
    Me.Caption = cFORMID
    HelpContextID = nFORMID
   
    '* Add colors to cmbFontColor.
    cARQTIPO = "RTF"
    With cmbfontcolor
        .AddItem "Black"
        .AddItem "Blue"
        .AddItem "Red"
        .AddItem "Green"
        .ListIndex = 0
    End With
   
  '  Dim i As Integer
  '  With cmbfontname
  '      For i = 0 To Screen.FontCount - 1
  '          .AddItem Screen.Fonts(i)
            '.Font Screen.Fonts(i)
            '.FontName = Screen.Fonts(i)
            '.FontSize = 18
            '.Width = 150
            '.Height = 150
   '     Next i
        '* Set ListIndex to 0.
    '    .ListIndex = 0
   ' End With

    With cmbFontSize

        '******************************************************************
        '* Populate the combo with sizes in
        '* increments of 2.
        '******************************************************************

        For i = 8 To 72 Step 2
            .AddItem i
        Next i

        '******************************************************************
        '* Set ListIndex to 0
        .ListIndex = 1                           '* size 10.
        '******************************************************************

    End With
    RichTextBox1.Width = Me.Width - 30           ''10695
    RichTextBox1.Height = Me.Height - 2225       ''6855
   
    '******************************************************************
    '* 0-Novo,1-Abrir,2-Salvar,3-Salvacomo
    '* 4-Editar,5-Visualizar,6-Imprimir,7-Exportar
    '******************************************************************
   
    ''Novo
    If Not aDIREITOS(0) Then
        mnunew.Enabled = False
        Command1.Enabled = False
    End If
    ''Abrir
    If Not aDIREITOS(1) Then
        mnuOpen.Enabled = False
        Command2.Enabled = False
    End If
    ''Salvar
    If Not aDIREITOS(2) Then
        mnusave.Enabled = False
        Command3.Enabled = False
    End If
    ''Salvar Como
    If Not aDIREITOS(3) Then
        mnusaveastxt.Enabled = False
        mnusaveasrtf.Enabled = False
    End If
    ''Editor Precisa estar Abilitado Barra Rolagen Funcionar
    If Not aDIREITOS(4) Then
    End If
    ''Visualizaçao Ainda nao Criada
    If Not aDIREITOS(5) Then
    End If
    ''Impressao e Configuraçao Impressao
    If Not aDIREITOS(6) Then
        mnuprint.Enabled = False
        mnuprintcfg.Enabled = False
        Command4.Enabled = False
        Command13.Enabled = False
    End If
    ''Exportaçao Ainda Nao Criada
    If Not aDIREITOS(7) Then
    End If
             
    ''Abre o Arquivo se passado cARQRTF
    ABRIR
    
    StatusBar1.Panels(6).Text = zUSER
    
   ' cmbfontname.tEXT = RichTextbox1.SelFontName
    
  '  FontCombo1.tEXT = RichTextbox1.SelFontName
  '  FontCombo1.SelFontName = RichTextbox1.SelFontName
    
  PrintPreview1.AuxiliaryButtonVisible = PrintPreview1.PrinterExists("Microsoft Print to PDF")
    PrintPreview1.AuxiliaryButtonToolTipText = "Salvar como PDF"
   
End Sub
Public Sub PrintPreview1_AuxiliaryButtonClick(UpdateReport As Boolean)
    PrintPreview1.ShowSaveToFile "Microsoft Print to PDF", "*.pdf"
    UpdateReport = False ' we don't need to update the report in the Print preview window after this action (the default value of UpdateReport parameter is True)
End Sub
'***************************************************************************
'*  Name         : Form_Resize
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       :
'*  Date         : 20 dez 2003
'***************************************************************************

Private Sub Form_Resize()
    Dim w As Integer
    Dim H As Integer

    w = 10695
    H = 6855

    If FrmRTf.WindowState = vbMaximized Then
        RichTextBox1.Width = (FrmRTf.Width - 300)
        RichTextBox1.Height = (FrmRTf.Height - 800)

    End If

    If FrmRTf.WindowState = vbNormal Then
        RichTextBox1.Width = w
        RichTextBox1.Height = H

    End If

End Sub

'***************************************************************************
'*  Name         : mnuabout_Click
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       :
'*  Date         : 20 dez 2003
'***************************************************************************

Private Sub mnuabout_Click()
    frmAbout.Show vbModal, Me
End Sub

'***************************************************************************
'*  Name         : mnubold_Click
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       :
'*  Date         : 20 dez 2003
'***************************************************************************

Private Sub mnubold_Click()
    Command7_Click
End Sub

Private Sub mnubullet_Click()
    RichTextBox1.SelBullet = Not RichTextBox1.SelBullet
    If RichTextBox1.SelBullet Then
        Command23.Appearance = 0
    Else
        Command23.Appearance = 1
    End If
    RichTextBox1.SetFocus

End Sub

Private Sub mnucapitalizar_Click()
    RichTextBox1.SelText = MMCase(RichTextBox1.SelText)
End Sub

'***************************************************************************
'*  Name         : mnucenter_Click
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       :
'*  Date         : 20 dez 2003
'***************************************************************************

Private Sub mnucenter_Click()
    If RichTextBox1.SelLength > 0 Then
        RichTextBox1.SelAlignment = 2
    End If
    RichTextBox1.SetFocus
End Sub

Private Sub mnucnvansi_Click()
    RichTextBox1.SelText = ConvAnsi(RichTextBox1.SelText)
End Sub

Private Sub mnucnvoem_Click()
    RichTextBox1.SelText = ConvOEM(RichTextBox1.SelText)
End Sub

Private Sub mnucnvansi2_Click()
    RichTextBox1.SelText = ConvAnsi2(RichTextBox1.SelText)
End Sub

Private Sub mnucnvoem2_Click()
    RichTextBox1.SelText = ConvOEM2(RichTextBox1.SelText)
End Sub

'***************************************************************************
'*  Name         : mnucopy_Click
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       :
'*  Date         : 20 dez 2003
'***************************************************************************

Private Sub mnucopy_Click()
    Clipboard.SetText RichTextBox1.SelText
    RichTextBox1.SetFocus
End Sub

'***************************************************************************
'*  Name         : mnucustomcolors_Click
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       :
'*  Date         : 20 dez 2003
'***************************************************************************

Private Sub mnucustomcolors_Click()
    Dim DLG As New clsCommonDialogs
    DLG.DialogColor Complete
    If DLG.Cancel = True Then
        Exit Sub
    End If
    RichTextBox1.SelColor = DLG.Color
    RichTextBox1.SetFocus
End Sub

'***************************************************************************
'*  Name         : mnucut_Click
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       :
'*  Date         : 20 dez 2003
'***************************************************************************

Private Sub mnucut_Click()
    Clipboard.SetText RichTextBox1.SelText
    RichTextBox1.SelText = ""
    RichTextBox1.SetFocus
End Sub

Private Sub mnudecInd_Click()
    Command25_Click
End Sub

Private Sub mnuDelSel_Click()
    If MDG("Apagar Seleçao") Then
        RichTextBox1.SelText = ""
    End If
 
End Sub

Private Sub mnuDown_Click()
    RichTextBox1.SelText = LCase(RichTextBox1.SelText)
End Sub

'***************************************************************************
'*  Name         : mnuexit_Click
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       :
'*  Date         : 20 dez 2003
'***************************************************************************
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If aDIREITOS(2) Then
        If MDG("Salvar Documento") Then
            mnusave_Click
        End If
    End If
End Sub

Private Sub mnuEditarInserirCaracter_Click()
    ePASS01 = FontCombo1.Text     'cmbfontname.text
    ePASS02 = True
    frmCharacters.Show vbModal, Me
    'FrmAscII.Show vbModal, Me
    If lRETU Then
        '      FrmRTf.RichTextbox1.SelFontName = eRETU02
        FrmRTf.RichTextBox1.SelText = eRETU01
        '      cmbfontname = eRETU02
    End If
End Sub

Private Sub mnuexit_Click()
    Unload Me
End Sub

'***************************************************************************
'*  Name         : mnufindNext_Click
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       :
'*  Date         : 20 dez 2003
'***************************************************************************

Private Sub mnuFindNext_Click()
    RichTextBox1.SetFocus
    achartf (RichTextBox1.SelStart + 1)
End Sub

Private Sub mnufindrep_Click()
    'frmFindReplace.Show vbModal, Me
End Sub

'***************************************************************************
'*  Name         : mnufind_Click
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       :
'*  Date         : 20 dez 2003
'***************************************************************************

Private Sub mnufind_Click()
    achartf (0)
End Sub

'***************************************************************************
'*  Name         : mnuFntesc_Click
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       :
'*  Date         : 20 dez 2003
'***************************************************************************

Private Sub mnuFntesc_Click()
    Command22_Click
End Sub

Private Sub mnuIncInd_Click()
    Command24_Click
End Sub

Private Sub mnuinsdate_Click()
    frmInsertDate.Show vbModal
    If lRETU Then
        FrmRTf.RichTextBox1.SelText = eRETU02
    End If
End Sub

Private Sub mnuinsimg_Click()
    Dim sFILTER, sFILENAME As String
    On Error Resume Next
    Clipboard.Clear
    sFILTER = ImgFILTER()
    sFILENAME = FileOpen(Me, sFILTER, 1, , , , "Abrir arquivo Imagens")
    If Not Len(sFILENAME) = 0 Then
        'Image1.Picture = LoadPictureEx(sFILENAME)
        Image1.Picture = LoadPicture(sFILENAME)
        Clipboard.SetData Image1.Picture
        SendMessage RichTextBox1.hWnd, WM_PASTE, 0, 0
    End If
End Sub

'***************************************************************************
'*  Name         : mnuitalico_Click
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       :
'*  Date         : 20 dez 2003
'***************************************************************************

Private Sub mnuitalico_Click()
    Command8_Click
End Sub

'***************************************************************************
'*  Name         : mnuleft_Click
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       :
'*  Date         : 20 dez 2003
'***************************************************************************

Private Sub mnuleft_Click()
    If RichTextBox1.SelLength > 0 Then
        RichTextBox1.SelAlignment = 0
    End If
    RichTextBox1.SetFocus
End Sub

'***************************************************************************
'*  Name         : mnunew_Click
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       :
'*  Date         : 20 dez 2003
'***************************************************************************

Private Sub mnunew_Click()
    RichTextBox1.Text = ""
    RichTextBox1.SetFocus
End Sub

'***************************************************************************
'*  Name         : mnuOpen_Click
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       :
'*  Date         : 20 dez 2003
'***************************************************************************

Private Sub mnuOpen_Click()
    Dim sRECENTFILE, sPath, sFILTER, sFILENAME As String
    sRECENTFILE = origem
    sPath = origem
    sFILTER = "Arquivos de Textos (*.TXT)" & vbNullChar & "*.TXT" & vbNullChar _
            & "Rich Text Arquivos (*.RTF)" & vbNullChar & "*.RTF" & vbNullChar _
            & "Todos os Arquivos" & vbNullChar & "*.*"
    sFILENAME = FileOpen(Me, sFILTER, 1, sRECENTFILE, "prn", sPath, "Abrir arquivo TXT RTF")
    If Not Len(sFILENAME) = 0 Then
        cARQRTF = NomeArq(sFILENAME)
        ABRIR
    End If


    '******************************************************************
    'Read in the text file cARQRTF
    'into the txtNotepad text box
    '******************************************************************

    '******************************************************************
    ''Open cARQRTF For Input As #1
    ''RichTextbox1 = Input(LOF(1), 1)
    ''Close #1
    '******************************************************************

End Sub

'***************************************************************************
'*  Name         : mnupaste_Click
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       :
'*  Date         : 20 dez 2003
'***************************************************************************


Private Sub mnupaste_Click()
    RichTextBox1.SelText = Clipboard.GetText
    RichTextBox1.SetFocus
End Sub

'***************************************************************************
'*  Name         : mnuprintcfg_Click
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       :
'*  Date         : 20 dez 2003
'***************************************************************************

Private Sub mnuprintcfg_Click()
    FrmPrintSetup.Show vbModal, Me
    RichTextBox1.SetFocus
End Sub

'***************************************************************************
'*  Name         : mnuprint_Click
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       :
'*  Date         : 20 dez 2003
'***************************************************************************

Private Sub mnuprint_Click()
    PrintPreview1.ShowPreview
    RichTextBox1.SetFocus
End Sub
Public Sub PrintPreview1_PrepareReport(Cancel As Boolean)
    PrintPreview1.PrintRichTextBox RichTextBox1
End Sub
Private Sub mnuright_Click()
    If RichTextBox1.SelLength > 0 Then
        RichTextBox1.SelAlignment = 1
    End If
    RichTextBox1.SetFocus
End Sub

Private Sub mnusaveasHTML_Click()
    Dim sFILTER, sFILENAME, stmp As String
    stmp = TrocaExt(cARQRTF, "HTML")
    sFILTER = "Arquivos de Textos (*.HTML)" & vbNullChar & "*.HTML" & vbNullChar & "Todos Arquivo" & vbNullChar & "*.*"
    sFILENAME = FileSave(Me, sFILTER, 1, "HTML", stmp, stmp, "Salvar HTML Como")
    If Not Len(sFILENAME) = 0 Then
        If FileExist(sFILENAME) Then
            If MDG("Arquivo ja Existe Sobrepor") Then
                DeleteFile sFILENAME 'Kill sFILENAME
                Exit Sub
            End If
        End If
        stmp = TrocaExt(sFILENAME, "TMP")
        RichTextBox1.SaveFile stmp, rtfText
        ePASS01 = stmp
        txttohtml stmp, sFILENAME
        DeleteFile stmp 'Kill stmp
    End If
End Sub

Private Sub mnusaveasPDF_Click()
    Dim sFILTER, sFILENAME, stmp As String
    stmp = TrocaExt(cARQRTF, "PDF")
    sFILTER = "Arquivos de Textos (*.PDF)" & vbNullChar & "*.PDF" & vbNullChar & "Todos Arquivo" & vbNullChar & "*.*"
    sFILENAME = FileSave(Me, sFILTER, 1, "PDF", stmp, stmp, "Salvar PDF Como")
    If Not Len(sFILENAME) = 0 Then
        stmp = TrocaExt(sFILENAME, "TMP")
        RichTextBox1.SaveFile stmp, rtfText
        ePASS01 = stmp
        formConvertToPDF.Show vbModal, Me
        DeleteFile stmp 'Kill stmp
    End If
End Sub

Private Sub mnusaveasrtf_Click()
    '***************************************************************************
    '*  Name         : mnusaveasrtf_Click
    '*  Description  :
    '*  Parameters   : None
    '*  Returns      : Nothing
    '*  Author       :
    '*  Date         : 20 dez 2003
    '***************************************************************************

    Dim sFILTER, sFILENAME As String
    sFILTER = "Arquivos de Textos (*.RTF)" & vbNullChar & "*.RTF" & vbNullChar & "Todos Arquivo" & vbNullChar & "*.*"
    sFILENAME = FileSave(Me, sFILTER, 1, "RTF", cARQRTF, cARQRTF, "Salvar RTF Como")
    If Not Len(sFILENAME) = 0 Then
        cARQRTF = sFILENAME
        cARQTIPO = "RTF"
        RichTextBox1.SaveFile sFILENAME, rtfRTF
        RichTextBox1.SetFocus
    End If
End Sub

'***************************************************************************
'*  Name         : mnusaveastxt_Click
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       :
'*  Date         : 20 dez 2003
'***************************************************************************

Private Sub mnusaveastxt_Click()
    Dim sFILTER, sFILENAME As String
    sFILTER = "Arquivos de Textos (*.TXT)" & vbNullChar & "*.TXT" & vbNullChar & "Todos Arquivo" & vbNullChar & "*.*"
    sFILENAME = FileSave(Me, sFILTER, 1, "TXT", cARQRTF, cARQRTF, "Salvar TXT Como")
    If Not Len(sFILENAME) = 0 Then
        cARQRTF = sFILENAME
        cARQTIPO = "TXT"
        RichTextBox1.SaveFile sFILENAME, rtfText
        RichTextBox1.SetFocus
    End If
    
End Sub

'***************************************************************************
'*  Name         : mnusave_Click
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       :
'*  Date         : 20 dez 2003
'***************************************************************************

Private Sub mnusave_Click()
    If cARQTIPO = "RTF" Then
        RichTextBox1.SaveFile cARQRTF, rtfRTF
    Else
        RichTextBox1.SaveFile cARQRTF, rtfText
    End If
    RichTextBox1.SetFocus
End Sub

Private Sub mnuscrinf_Click()
    RichTextBox1.SelCharOffset = -55
End Sub

Private Sub mnuscrnor_Click()
    RichTextBox1.SelCharOffset = 0
End Sub

Private Sub mnuscrsup_Click()
    RichTextBox1.SelCharOffset = 55
End Sub

'***************************************************************************
'*  Name         : mnuselectall_Click
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       :
'*  Date         : 20 dez 2003
'***************************************************************************


Private Sub mnuselectall_Click()
    RichTextBox1.SelStart = 0
    RichTextBox1.SelLength = Len(RichTextBox1.Text)
    RichTextBox1.SetFocus
End Sub

Private Sub mnuselectline_Click()
    SendKeys "{HOME}"
    SendKeys "+{END}"
End Sub

Private Sub mnuselsen_Click()
    With RichTextBox1
     '   .Span ".?!:", True, True
        .SelLength = .SelLength + 1
    End With
End Sub

Private Sub mnuselword_Click()
   'RichTextbox1.S
   ' RichTextbox1.Span " ,;:.?!", True, True
End Sub

Private Sub mnusendmail_Click()
    mnusave_Click                                ''salva antes de enviar
    ePASS01 = Array("", _
                    "", _
                    "", _
                    "", _
                    "", _
                    cARQRTF, _
                    "", _
                    False)
    FrmSendMail.Show vbModal, Me
End Sub

'***************************************************************************
'*  Name         : mnustrik_Click
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       :
'*  Date         : 20 dez 2003
'***************************************************************************

Private Sub mnustrik_Click()
    Command21_Click
End Sub

Private Sub mnutirace_Click()
    RichTextBox1.SelText = Tirace(RichTextBox1.SelText)
End Sub

'***************************************************************************
'*  Name         : mnuunder_Click
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       :
'*  Date         : 20 dez 2003
'***************************************************************************

Private Sub mnuunder_Click()
    Command9_Click
End Sub

'***************************************************************************
'*  Name         : mnuundo_Click
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       :
'*  Date         : 20 dez 2003
'***************************************************************************

Private Sub mnuundo_Click()
    Dim Success
    Success = SendMessage(RichTextBox1.hWnd, WM_UNDO, 0&, 0&)
    RichTextBox1.SetFocus
End Sub

Private Sub mnuUpper_Click()
    RichTextBox1.SelText = UCase(RichTextBox1.SelText)
End Sub


Private Sub MenuEvents_MenuHelp(ByVal MenuText As String, ByVal MenuHelp As String, ByVal Enabled As Boolean)
    If gbSubClassMenu Then
        '/ this event show Decription menu-item to a StatusBar control
        '/ customize for your project, if you need this.
        '/ Note: MenuText contain the menu Caption.
        '/       MenuHelp contain the Description string.
        '/ example: sbStatusBar.Panels(1).text = MenuHelp
        '/          lblMenuHelp.Caption = MenuHelp
    End If
End Sub

Public Sub SubClassMenuXP()

    '/ this code is made by MenuCreator add-in

    '/ prepare the caption for subclassing. Warning! Don't remove this comment!!!
    Menu1.Caption = "&Arquivo"
    mnunew.Caption = "&Novo|#new"
    mnuOpen.Caption = "&Abrir|#open"
    mnusave.Caption = "&Salvar|#salvar"
    mnusaveastxt.Caption = "Salvar Como &TXT"
    mnusaveasrtf.Caption = "Salvar Como &RTF"
    mnuprint.Caption = "&Imprimir|#printer"
    mnuprintcfg.Caption = "&Configurar Impressora|#printcfg"
    mnusendmail.Caption = "&Enviar Email|#mail"
    mnuexit.Caption = "Sa&ir|#exit"
    Menu2.Caption = "&Editar"
    mnuundo.Caption = "&Desfazer|#refazer"
    tr01.Caption = "-"
    mnuselectall.Caption = "Selecionar &Tudo"
    mnuselectline.Caption = "Selecionar Linha"
    mnuselword.Caption = "Selecionar Palavra"
    mnuselsen.Caption = "Selecionar Sentençao"
    mnuDelSel.Caption = "Deletar Seleçao"
    tr02.Caption = "-"
    mnucopy.Caption = "&Copiar|#copy"
    mnupaste.Caption = "Co&lar|#colar"
    mnucut.Caption = "&Recortar|#cut"
    tr03.Caption = "-"
    mnufind.Caption = "Locali&zar|#busca"
    mnufindNext.Caption = "Localizar Pro&xima|#proximabusca"
    mnufindrep.Caption = "Localizar e Substituir"
    tr04.Caption = "-"
    mnuinsdate.Caption = "Inserir Data|#diahora"
    mnuinsimg.Caption = "Inserir Imagem|#imagem"
    mnuEditarInserirCaracter.Caption = "Inserir Caracter"
    Menu3.Caption = "Format"
    mnutext.Caption = "Alinhamento"
    mnuleft.Caption = "Esquerada|#left"
    mnuright.Caption = "Direita|#right"
    mnucenter.Caption = "Centralizado|#center"
    mnufonte.Caption = "Fontes Letra"
    mnucustomcolors.Caption = "Escolher Cor|#imagem2"
    mnucolor.Caption = "Cores"
    mnuFntesc.Caption = "Fonte de Letra|#letra"
    mnubold.Caption = "Negrito|#Bold"
    mnuitalico.Caption = "Italico|#italico"
    mnuunder.Caption = "Sublilhado|#underline"
    mnustrik.Caption = "Riscado|#riscado"
    mnuVerFontes.Caption = "Ver Fontes|#find"
    xxx.Caption = "Identicaçao"
    mnuIncInd.Caption = "Aumentar Posiçao|#imagem"
    mnudecInd.Caption = "Diminuir Posiçao|#decrese"
    yyy.Caption = "Escrita"
    mnuscrsup.Caption = "Escrita Superior|#sob"
    mnuscrinf.Caption = "Escrita Inferior|#sub"
    mnuscrnor.Caption = "Escrita Normal"
    mnubullet.Caption = "Marcador|#xxx"
    mnucnv.Caption = "Converter "
    mnuUpper.Caption = "Maisculas|#maisculas"
    mnuDown.Caption = "Minusculas|#minusculas"
    mnucapitalizar.Caption = "Capitalizar"
    mnutirace.Caption = "Tira Acentuaçao"
    mnucnvoem.Caption = "Converter Ansi(Windows)-> Oem(DOS)"
    mnucnvoem2.Caption = "Converter Ansi(Windows)-> Oem(DOS)Padrao2"
    mnucnvansi.Caption = "Converter Oem(DOS)->Ansi(Windows)"
    mnucnvansi2.Caption = "Converter Oem(DOS)->Ansi(Windows)Padrao2"
    Menu4.Caption = "Aj&uda"
    mnuabout.Caption = "&Sobre|#info"

    '/ Subclassing menu. Warning! Don't remove this comment!!!

    Set MenuEvents = New CEvents
    Set objMenuEx = New cMenuEx
    Call objMenuEx.Install(Me.hWnd, App.Path & "\" & Me.Name, ImgMenu, 2, MenuEvents)

End Sub

Public Sub MenuDesigner()
    '/ Open Menu Designer tool
    objMenuEx.MenuDesigner Me.hWnd
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '/ This Form_UnLoad is add by MenuCreator

    If gbSubClassMenu Then
        '/ prevent error if the menu is not subclassed
        On Error Resume Next
        '/ release object
        Call objMenuEx.Uninstall(Me.hWnd, ImgMenu, MenuEvents)
        Set MenuEvents = Nothing
        Set objMenuEx = Nothing
    End If

End Sub


