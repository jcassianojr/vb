VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{7020C36F-09FC-41FE-B822-CDE6FBB321EB}#1.0#0"; "vbccr17.ocx"
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
   Begin VBCCR17.FontCombo FontCombo1 
      Height          =   315
      Left            =   2400
      TabIndex        =   42
      Top             =   360
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   556
      Text            =   "FrmRtf.frx":058A
      ExtendedUI      =   -1  'True
   End
   Begin VB.CommandButton CmdExpHtml 
      Height          =   315
      Left            =   1620
      Picture         =   "FrmRtf.frx":05BE
      Style           =   1  'Graphical
      TabIndex        =   41
      ToolTipText     =   "Exportar Documento como HTML"
      Top             =   0
      Width           =   360
   End
   Begin VB.CommandButton CmdExpPdf 
      Height          =   315
      Left            =   1320
      Picture         =   "FrmRtf.frx":0B48
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   "Exportar Documento como PDF"
      Top             =   0
      Width           =   360
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   39
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
            Picture         =   "FrmRtf.frx":10D2
            TextSave        =   "14:37"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            Object.Width           =   2302
            MinWidth        =   2293
            Picture         =   "FrmRtf.frx":166C
            TextSave        =   "05/05/2021"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Picture         =   "FrmRtf.frx":1C06
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
            Picture         =   "FrmRtf.frx":1FA0
            Key             =   "center"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":233A
            Key             =   "justificado"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":26D4
            Key             =   "left"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":2A6E
            Key             =   "right"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":2E08
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":31A2
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":373C
            Key             =   "cut"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":3AD6
            Key             =   "diahora"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":3E70
            Key             =   "editor"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":440A
            Key             =   "find"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":47A4
            Key             =   "marcado"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":48FE
            Key             =   "letra01"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":4C98
            Key             =   "tamanho"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":5032
            Key             =   "paragro"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":53CC
            Key             =   "decrese"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":5766
            Key             =   "incrise"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":5B00
            Key             =   "imagem"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":5E9A
            Key             =   "italico"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":6234
            Key             =   "s_Key19"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":65CE
            Key             =   "mail"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":6B68
            Key             =   "new"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":6F02
            Key             =   "open"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":749C
            Key             =   "xxx"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":7836
            Key             =   "imagem2"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":7BD0
            Key             =   "colar"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":816A
            Key             =   "printer"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":8704
            Key             =   "printcfg"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":8A9E
            Key             =   "sair"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":8E38
            Key             =   "salvar"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":91D2
            Key             =   "busca"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":956C
            Key             =   "proximabusca"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":9906
            Key             =   "letra"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":9CA0
            Key             =   "abc"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":A03A
            Key             =   "riscado"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":A3D4
            Key             =   "sub"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":A76E
            Key             =   "sob"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":AB08
            Key             =   "minusculas"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":AEA2
            Key             =   "maisculas"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":B23C
            Key             =   "underline"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":B5D6
            Key             =   "refazer"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":B970
            Key             =   "info"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRtf.frx":BF0A
            Key             =   "exit"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command34 
      Height          =   315
      Left            =   8760
      Picture         =   "FrmRtf.frx":C4A4
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "Capitalizar"
      Top             =   360
      Width           =   360
   End
   Begin VB.CommandButton Command33 
      Height          =   315
      Left            =   3120
      Picture         =   "FrmRtf.frx":C82E
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Enviar Email  Documento Atual"
      Top             =   0
      Width           =   360
   End
   Begin VB.CommandButton Command32 
      Height          =   315
      Left            =   10320
      Picture         =   "FrmRtf.frx":CDB8
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Inserir Data"
      Top             =   360
      Width           =   360
   End
   Begin VB.CommandButton Command31 
      Height          =   315
      Left            =   9960
      Picture         =   "FrmRtf.frx":D142
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Incluir Imagens"
      Top             =   360
      Width           =   360
   End
   Begin VB.CommandButton Command30 
      Height          =   315
      Left            =   4560
      Picture         =   "FrmRtf.frx":D4CC
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Selecionar Todo o Texto"
      Top             =   0
      Width           =   360
   End
   Begin VB.CommandButton Command29 
      Height          =   315
      Left            =   9120
      Picture         =   "FrmRtf.frx":D856
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Mudar Letra Escrita Inferiores"
      Top             =   360
      Width           =   360
   End
   Begin VB.CommandButton Command28 
      Height          =   315
      Left            =   9600
      Picture         =   "FrmRtf.frx":DBE0
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Mudar Letra Escrita Superiores"
      Top             =   360
      Width           =   360
   End
   Begin VB.CommandButton Command27 
      Height          =   315
      Left            =   8400
      Picture         =   "FrmRtf.frx":DF6A
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Converter Maisculas"
      Top             =   360
      Width           =   360
   End
   Begin VB.CommandButton Command26 
      Height          =   315
      Left            =   8040
      Picture         =   "FrmRtf.frx":E2F4
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Converter Minusculas"
      Top             =   360
      Width           =   360
   End
   Begin VB.CommandButton Command25 
      Height          =   315
      Left            =   9120
      Picture         =   "FrmRtf.frx":E67E
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Diminuir Recuo"
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton Command24 
      Height          =   315
      Left            =   8760
      Picture         =   "FrmRtf.frx":EA08
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Aumentar Recuo"
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton Command23 
      Height          =   315
      Left            =   8400
      Picture         =   "FrmRtf.frx":ED92
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Marcadores"
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton Command22 
      Height          =   315
      Left            =   1920
      Picture         =   "FrmRtf.frx":F11C
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Selecionar Fonte"
      Top             =   360
      Width           =   360
   End
   Begin VB.CommandButton Command21 
      Height          =   315
      Left            =   1560
      Picture         =   "FrmRtf.frx":F266
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Cortado"
      Top             =   360
      Width           =   360
   End
   Begin VB.CommandButton Command20 
      Height          =   315
      Left            =   7800
      Picture         =   "FrmRtf.frx":F5F0
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Localizar Proxima"
      Top             =   0
      Width           =   360
   End
   Begin VB.TextBox texto 
      Height          =   285
      Left            =   5400
      TabIndex        =   23
      Tag             =   "Texto"
      ToolTipText     =   "Texto Para Localizar"
      Top             =   0
      Width           =   1995
   End
   Begin VB.CommandButton Command19 
      Height          =   315
      Left            =   7440
      Picture         =   "FrmRtf.frx":F97A
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Localizar"
      Top             =   0
      Width           =   360
   End
   Begin VB.CommandButton Command18 
      Height          =   315
      Left            =   3840
      Picture         =   "FrmRtf.frx":FD04
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Copiar Texto Selecionado"
      Top             =   0
      Width           =   360
   End
   Begin VB.CommandButton Command17 
      Height          =   315
      Left            =   3480
      Picture         =   "FrmRtf.frx":1028E
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Recortar Texto Selecionado"
      Top             =   0
      Width           =   360
   End
   Begin VB.CommandButton Command16 
      Height          =   315
      Left            =   4200
      Picture         =   "FrmRtf.frx":10618
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Colar Texto "
      Top             =   0
      Width           =   360
   End
   Begin VB.CommandButton Command15 
      Height          =   315
      Left            =   7560
      Picture         =   "FrmRtf.frx":10BA2
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Escolher Cor"
      Top             =   360
      Width           =   495
   End
   Begin VB.CommandButton Command14 
      Height          =   315
      Left            =   11160
      Picture         =   "FrmRtf.frx":10F2C
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Sair"
      Top             =   360
      Width           =   495
   End
   Begin VB.CommandButton Command13 
      Height          =   315
      Left            =   2760
      Picture         =   "FrmRtf.frx":114B6
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Configurar Impressoras"
      Top             =   0
      Width           =   360
   End
   Begin VB.CommandButton Command12 
      Height          =   315
      Left            =   10320
      Picture         =   "FrmRtf.frx":11840
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Alinhar a Direita"
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton Command11 
      Height          =   315
      Left            =   9960
      Picture         =   "FrmRtf.frx":11BCA
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Centralizar"
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton Command10 
      Height          =   315
      Left            =   9600
      Picture         =   "FrmRtf.frx":11F54
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Alinhar a Esquerda"
      Top             =   0
      Width           =   360
   End
   Begin VB.CommandButton Command9 
      Height          =   315
      Left            =   1200
      Picture         =   "FrmRtf.frx":122DE
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Sublinhado"
      Top             =   360
      Width           =   360
   End
   Begin VB.CommandButton Command8 
      Height          =   315
      Left            =   840
      Picture         =   "FrmRtf.frx":12668
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Italico"
      Top             =   360
      Width           =   360
   End
   Begin VB.CommandButton Command7 
      Height          =   315
      Left            =   480
      Picture         =   "FrmRtf.frx":129F2
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Negrito/Bold"
      Top             =   360
      Width           =   360
   End
   Begin VB.CommandButton Command6 
      Height          =   315
      Left            =   0
      Picture         =   "FrmRtf.frx":12D7C
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Checar Ortografia"
      Top             =   360
      Width           =   360
   End
   Begin VB.CommandButton Command5 
      Height          =   315
      Left            =   4920
      Picture         =   "FrmRtf.frx":13106
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Desfazer"
      Top             =   0
      Width           =   360
   End
   Begin VB.CommandButton Command4 
      Height          =   315
      Left            =   2100
      Picture         =   "FrmRtf.frx":13490
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Imprimir Documento Atual"
      Top             =   0
      Width           =   360
   End
   Begin VB.CommandButton Command3 
      Height          =   315
      Left            =   960
      Picture         =   "FrmRtf.frx":13A1A
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Salvar Documento Atual"
      Top             =   0
      Width           =   360
   End
   Begin VB.CommandButton Command2 
      Height          =   315
      Left            =   480
      Picture         =   "FrmRtf.frx":13DA4
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Abrir Documento Existente"
      Top             =   0
      Width           =   360
   End
   Begin VB.CommandButton Command1 
      Height          =   315
      Left            =   0
      Picture         =   "FrmRtf.frx":1432E
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Iniciar Documento Em Branco"
      Top             =   0
      Width           =   360
   End
   Begin VB.ComboBox cmbfontcolor 
      Height          =   315
      Left            =   5760
      TabIndex        =   3
      Text            =   "Combo1"
      ToolTipText     =   "Escolher Cor"
      Top             =   360
      Width           =   1695
   End
   Begin VB.ComboBox cmbfontsize 
      Height          =   315
      Left            =   4800
      TabIndex        =   2
      Text            =   "Combo1"
      ToolTipText     =   "Tamanho da Letra"
      Top             =   360
      Width           =   915
   End
   Begin VB.ComboBox cmbfontname 
      Height          =   315
      Left            =   10920
      TabIndex        =   1
      Text            =   "Fonte de Letra"
      ToolTipText     =   "Fonte de Letra"
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin RichTextLib.RichTextBox RichTextbox1 
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   10610
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"FrmRtf.frx":146B8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
        RichTextbox1.LoadFile cARQRTF, rtfRTF
    Else
        cARQTIPO = "TXT"
        RichTextbox1.LoadFile cARQRTF, rtfText
    End If
    RichTextbox1.SetFocus
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
    RichTextbox1.SetFocus
    textfound = RichTextbox1.Find(texto, nPOS, , rtfWholeWord)
    If textfound = -1 Then
        Alert vbCr & "Texto Não Foi Encontrado"
    End If
    RichTextbox1.SetFocus
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

    With RichTextbox1
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
    RichTextbox1.SelFontSize = cmbFontSize.tEXT
    ''RichTextbox1.SetFocus
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
    RichTextbox1.SetFocus
    achartf (RichTextbox1.SelStart + 1)
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
    RichTextbox1.SelStrikeThru = Not RichTextbox1.SelStrikeThru
    If RichTextbox1.SelStrikeThru Then
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
    
    
    
    With RichTextbox1
        .SelFontName = oDLG.Font
        .SelFontSize = oDLG.FontSize
        .SelBold = oDLG.FontBold
        .SelItalic = oDLG.FontItalic
        .SelStrikeThru = oDLG.FontStrikeThru
        .SelUnderline = oDLG.FontUnderline
        .SelColor = oDLG.Color
    End With

End Sub

Private Sub Command23_Click()
    mnubullet_Click
End Sub

Private Sub Command24_Click()
    RichTextbox1.SelIndent = IncDec(RichTextbox1.SelIndent, 1, 0, 99)
End Sub

Private Sub Command25_Click()
    RichTextbox1.SelIndent = IncDec(RichTextbox1.SelIndent, -1, 0, 99)
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
    RichTextbox1.SelText = LCase(RichTextbox1.SelText)
End Sub

Private Sub Command27_Click()
    RichTextbox1.SelText = UCase(RichTextbox1.SelText)
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
    RichTextbox1.SelBold = Not RichTextbox1.SelBold
    If RichTextbox1.SelBold Then
        Command7.Appearance = 0
    Else
        Command7.Appearance = 1
    End If
    RichTextbox1.SetFocus
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
    RichTextbox1.SelItalic = Not RichTextbox1.SelItalic
    If RichTextbox1.SelItalic Then
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
    RichTextbox1.SelUnderline = Not RichTextbox1.SelUnderline
    If RichTextbox1.SelUnderline = False Then
        Command9.Appearance = 0
    Else
        Command9.Appearance = 1
    End If
End Sub

Private Sub FontCombo1_Click()
    RichTextbox1.SelFontName = FontCombo1.tEXT
    RichTextbox1.SetFocus
    
    
    

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
    RichTextbox1.Width = Me.Width - 30           ''10695
    RichTextbox1.Height = Me.Height - 2225       ''6855
   
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
    
    StatusBar1.Panels(6).tEXT = zUSER
    
   ' cmbfontname.tEXT = RichTextbox1.SelFontName
    
  '  FontCombo1.tEXT = RichTextbox1.SelFontName
  '  FontCombo1.SelFontName = RichTextbox1.SelFontName
    

   
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
    Dim h As Integer

    w = 10695
    h = 6855

    If FrmRTf.WindowState = vbMaximized Then
        RichTextbox1.Width = (FrmRTf.Width - 300)
        RichTextbox1.Height = (FrmRTf.Height - 800)

    End If

    If FrmRTf.WindowState = vbNormal Then
        RichTextbox1.Width = w
        RichTextbox1.Height = h

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
    RichTextbox1.SelBullet = Not RichTextbox1.SelBullet
    If RichTextbox1.SelBullet Then
        Command23.Appearance = 0
    Else
        Command23.Appearance = 1
    End If
    RichTextbox1.SetFocus

End Sub

Private Sub mnucapitalizar_Click()
    RichTextbox1.SelText = MMCase(RichTextbox1.SelText)
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
    If RichTextbox1.SelLength > 0 Then
        RichTextbox1.SelAlignment = 2
    End If
    RichTextbox1.SetFocus
End Sub

Private Sub mnucnvansi_Click()
    RichTextbox1.SelText = ConvAnsi(RichTextbox1.SelText)
End Sub

Private Sub mnucnvoem_Click()
    RichTextbox1.SelText = ConvOEM(RichTextbox1.SelText)
End Sub

Private Sub mnucnvansi2_Click()
    RichTextbox1.SelText = ConvAnsi2(RichTextbox1.SelText)
End Sub

Private Sub mnucnvoem2_Click()
    RichTextbox1.SelText = ConvOEM2(RichTextbox1.SelText)
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
    Clipboard.SetText RichTextbox1.SelText
    RichTextbox1.SetFocus
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
    RichTextbox1.SelColor = DLG.Color
    RichTextbox1.SetFocus
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
    Clipboard.SetText RichTextbox1.SelText
    RichTextbox1.SelText = ""
    RichTextbox1.SetFocus
End Sub

Private Sub mnudecInd_Click()
    Command25_Click
End Sub

Private Sub mnuDelSel_Click()
    If MDG("Apagar Seleçao") Then
        RichTextbox1.SelText = ""
    End If
 
End Sub

Private Sub mnuDown_Click()
    RichTextbox1.SelText = LCase(RichTextbox1.SelText)
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
    ePASS01 = FontCombo1.tEXT     'cmbfontname.text
    ePASS02 = True
    frmCharacters.Show vbModal, Me
    'FrmAscII.Show vbModal, Me
    If lRETU Then
        '      FrmRTf.RichTextbox1.SelFontName = eRETU02
        FrmRTf.RichTextbox1.SelText = eRETU01
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
    RichTextbox1.SetFocus
    achartf (RichTextbox1.SelStart + 1)
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
        FrmRTf.RichTextbox1.SelText = eRETU02
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
        SendMessage RichTextbox1.hWnd, WM_PASTE, 0, 0
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
    If RichTextbox1.SelLength > 0 Then
        RichTextbox1.SelAlignment = 0
    End If
    RichTextbox1.SetFocus
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
    RichTextbox1.tEXT = ""
    RichTextbox1.SetFocus
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
    RichTextbox1.SelText = Clipboard.GetText
    RichTextbox1.SetFocus
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
    RichTextbox1.SetFocus
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
    'Dim DLG As New clsCommonDialogs
    'DLG.PrinterFlags = CdlPDReturnDC + CdlPDNoPageNums
    'If RichTextbox1.SelLength = 0 Then
    '    DLG.PrinterFlags = DLG.PrinterFlags + CdlPDAllPages
    'Else
    '    DLG.PrinterFlags = DLG.PrinterFlags + CdlPDSelection
    'End If
    'DLG.DialogPrint
    'Printer.Print ""                             'Inicializa Impressora Necessario
    'RichTextbox1.SelPrint Printer.hDC
    PrinterEx.ShowPrintPreview Me, "MyPrintingRTF"
    RichTextbox1.SetFocus
End Sub
Public Property Get Printer() As Printer
    Set Printer = vbExtra.Printer2
End Property
Public Property Set Printer(nPrinter As Printer)
    Set vbExtra.Printer2 = nPrinter
End Property
Public Sub MyPrintingRTF()
'    Printer.FontName = "Arial"
'    Printer.FontSize = 12
    PrinterEx.PrintRichTextBox RichTextbox1
    Printer.EndDoc
End Sub


'***************************************************************************
'*  Name         : mnuright_Click
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       :
'*  Date         : 20 dez 2003
'***************************************************************************

Private Sub mnuright_Click()
    If RichTextbox1.SelLength > 0 Then
        RichTextbox1.SelAlignment = 1
    End If
    RichTextbox1.SetFocus
End Sub

Private Sub mnusaveasHTML_Click()
    Dim sFILTER, sFILENAME, stmp As String
    stmp = TrocaExt(cARQRTF, "HTML")
    sFILTER = "Arquivos de Textos (*.HTML)" & vbNullChar & "*.HTML" & vbNullChar & "Todos Arquivo" & vbNullChar & "*.*"
    sFILENAME = FileSave(Me, sFILTER, 1, "HTML", stmp, stmp, "Salvar HTML Como")
    If Not Len(sFILENAME) = 0 Then
        If FileExist(sFILENAME) Then
            If MDG("Arquivo ja Existe Sobrepor") Then
                Kill sFILENAME
                Exit Sub
            End If
        End If
        stmp = TrocaExt(sFILENAME, "TMP")
        RichTextbox1.SaveFile stmp, rtfText
        ePASS01 = stmp
        txttohtml stmp, sFILENAME
        Kill stmp
    End If
End Sub

Private Sub mnusaveasPDF_Click()
    Dim sFILTER, sFILENAME, stmp As String
    stmp = TrocaExt(cARQRTF, "PDF")
    sFILTER = "Arquivos de Textos (*.PDF)" & vbNullChar & "*.PDF" & vbNullChar & "Todos Arquivo" & vbNullChar & "*.*"
    sFILENAME = FileSave(Me, sFILTER, 1, "PDF", stmp, stmp, "Salvar PDF Como")
    If Not Len(sFILENAME) = 0 Then
        stmp = TrocaExt(sFILENAME, "TMP")
        RichTextbox1.SaveFile stmp, rtfText
        ePASS01 = stmp
        formConvertToPDF.Show vbModal, Me
        Kill stmp
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
        RichTextbox1.SaveFile sFILENAME, rtfRTF
        RichTextbox1.SetFocus
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
        RichTextbox1.SaveFile sFILENAME, rtfText
        RichTextbox1.SetFocus
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
        RichTextbox1.SaveFile cARQRTF, rtfRTF
    Else
        RichTextbox1.SaveFile cARQRTF, rtfText
    End If
    RichTextbox1.SetFocus
End Sub

Private Sub mnuscrinf_Click()
    RichTextbox1.SelCharOffset = -55
End Sub

Private Sub mnuscrnor_Click()
    RichTextbox1.SelCharOffset = 0
End Sub

Private Sub mnuscrsup_Click()
    RichTextbox1.SelCharOffset = 55
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
    RichTextbox1.SelStart = 0
    RichTextbox1.SelLength = Len(RichTextbox1.tEXT)
    RichTextbox1.SetFocus
End Sub

Private Sub mnuselectline_Click()
    SendKeys "{HOME}"
    SendKeys "+{END}"
End Sub

Private Sub mnuselsen_Click()
    With RichTextbox1
        .Span ".?!:", True, True
        .SelLength = .SelLength + 1
    End With
End Sub

Private Sub mnuselword_Click()
    RichTextbox1.Span " ,;:.?!", True, True
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
    RichTextbox1.SelText = Tirace(RichTextbox1.SelText)
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
    Success = SendMessage(RichTextbox1.hWnd, WM_UNDO, 0&, 0&)
    RichTextbox1.SetFocus
End Sub

Private Sub mnuUpper_Click()
    RichTextbox1.SelText = UCase(RichTextbox1.SelText)
End Sub

'Private Sub mnuVerFontes_Click()
'    FntViewer.Show
'End Sub

Private Sub RichTextbox1_Change()
    ''        With RichTextbox1
    ''            LineCount = SendMessageLong(.hwnd, EM_GETLINECOUNT, 0&, 0&)
    ''            ChrsUpToLast = SendMessageLong(.hwnd, EM_LINEINDEX, LineCount - 1, 0&)
    ''            DocumentSize = SendMessageLong(.hwnd, EM_LINELENGTH, ChrsUpToLast, 0&) + ChrsUpToLast
    ''            CURRLINE = SendMessageLong(frmNewDocument.RichTextbox1.hwnd, EM_LINEFROMCHAR, -1&, 0&) + 1
    ''        End With

End Sub

Private Sub RichTextbox1_KeyDown(KeyCode As Integer, Shift As Integer)
    ''    Dim CurrLine
    
    ''    CurrLine = SendMessageLong(RichTextbox1.hwnd, EM_LINEFROMCHAR, -1&, 0&) + 1
    ''    StatusBar1.Panels(2).Text = "Linha " & CurrLine
    ''StatusBar1.Panels(4) = Format$(CurrLine, "##,###")
    ''Call ControlsEnabled

End Sub

'***************************************************************************
'*  Name         : RichTextBox1_SelChange
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       :
'*  Date         : 20 dez 2003
'***************************************************************************

Private Sub RichTextBox1_SelChange()
    On Error GoTo errhandler

    '******************************************************************
    '* SelFontSize returns the font size, or Null if
    '* it's mixed.
    '******************************************************************

    If Not IsNull(RichTextbox1.SelFontSize) Then
        cmbFontSize.tEXT = RichTextbox1.SelFontSize
    End If

    '* Show Font name in the ComboBox.
   ' cmbfontname.tEXT = RichTextbox1.SelFontName

    '******************************************************************
    '* Show color of text in the ComboBox. Use the
    '* intrinsic constants for color to determine
    '* the color of the text.
    '******************************************************************

    Select Case RichTextbox1.SelColor
    Case vbBlack
        cmbfontcolor.ListIndex = 0
    Case vbBlue
        cmbfontcolor.ListIndex = 1
    Case vbRed
        cmbfontcolor.ListIndex = 2
    Case vbGreen
        cmbfontcolor.ListIndex = 3
    End Select
    Exit Sub
errhandler:
   ' cmbfontname.tEXT = cmbfontname.tEXT
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


