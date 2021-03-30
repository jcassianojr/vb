VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{7020C36F-09FC-41FE-B822-CDE6FBB321EB}#1.0#0"; "vbccr17.ocx"
Begin VB.Form FrmFemeaPad 
   Caption         =   "Femea Padrao"
   ClientHeight    =   7065
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11100
   LinkTopic       =   "Form1"
   ScaleHeight     =   7065
   ScaleWidth      =   11100
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   7095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   12515
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Padrao"
      TabPicture(0)   =   "FrmFemeaPad.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "dizcau"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "dizefe"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "dizfal(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "dizpro(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label4"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblLabels(14)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "gridrevcau"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "gridcau"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "gridefe"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "GRIDFAL"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "GridPRO"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "CmdExcPro"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "CmdExcFal"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "CmdExcEfe"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "CmdExcCau"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Retornar(0)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Command1"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Cmdescefe"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "CmdEscFal"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "CmdEscPro"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "CmdNovo(3)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "CmdNovo(2)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "CmdNovo(1)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "CmdNovo(0)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "CmdExpCau(0)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "CmdRem(0)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "CmdRem(1)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "CmdRem(2)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Command4"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "CmdRem(3)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "CmdSev"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "CmdRenumerar(0)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "CmdRenumerar(1)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "CmdRenumerar(2)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "CmdRenumerar(3)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "CmdExpCau(1)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "CmdExpCau(2)"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "txtFields(24)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "CMDSIG(0)"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "CmdMacSig(0)"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "CmdMacSig(1)"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "CmdMacSig(2)"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "CmdREvi"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "Command5"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "CmdEdiRevi"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "DIZAPU"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "CmdExpRPN"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "Command6"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "CMDSIG(1)"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "barra"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).ControlCount=   53
      TabCaption(1)   =   "Pre Cadastro"
      TabPicture(1)   =   "FrmFemeaPad.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Gridfemprecad"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Retornar(1)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdedipre(2)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "CmddelPre"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "CmdPreinc"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Command3"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Processos"
      TabPicture(2)   =   "FrmFemeaPad.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "dizpro(1)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "dizfal(1)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "gridpf"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "GridPadPF"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Retornar(2)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "CmdExcPFPad(0)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Command2"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "CmdIncPf"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "CmdExcPFPad(1)"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "CmdExcPFtot"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).ControlCount=   10
      TabCaption(3)   =   "Apuracao"
      TabPicture(3)   =   "FrmFemeaPad.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Grid"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Toolbar1"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).ControlCount=   2
      Begin VBCCR17.ProgressBar barra 
         Height          =   255
         Left            =   8400
         Top             =   3120
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
      End
      Begin VB.CommandButton CMDSIG 
         Height          =   255
         Index           =   1
         Left            =   10440
         Picture         =   "FrmFemeaPad.frx":0070
         Style           =   1  'Graphical
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Revisar Todas Causas"
         Height          =   495
         Left            =   9600
         TabIndex        =   67
         Top             =   6000
         Width           =   1215
      End
      Begin VB.CommandButton CmdExpRPN 
         Caption         =   "RPN->Mes"
         Height          =   255
         Left            =   9840
         TabIndex        =   66
         ToolTipText     =   "Exporta os Indices RPN para Uma Competencia Ja Acumulada"
         Top             =   4560
         Width           =   975
      End
      Begin VB.TextBox DIZAPU 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   9840
         Locked          =   -1  'True
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   3120
         Width           =   975
      End
      Begin VB.CommandButton CmdEdiRevi 
         Caption         =   "Editar"
         Height          =   255
         Left            =   8400
         TabIndex        =   64
         Top             =   6720
         Width           =   1095
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Nova Rev."
         Height          =   255
         Left            =   8400
         TabIndex        =   63
         Top             =   6360
         Width           =   1095
      End
      Begin VB.CommandButton CmdREvi 
         Caption         =   "Ver Revisoes"
         Height          =   255
         Left            =   8400
         TabIndex        =   62
         Top             =   6000
         Width           =   1095
      End
      Begin VB.CommandButton CmdExcPFtot 
         Caption         =   "Excluir Todos PF"
         Height          =   615
         Left            =   -68640
         TabIndex        =   60
         Top             =   2040
         Width           =   855
      End
      Begin VB.CommandButton CmdExcPFPad 
         Caption         =   "Excluir Falha"
         Height          =   735
         Index           =   1
         Left            =   -65040
         TabIndex        =   57
         Top             =   1680
         Width           =   855
      End
      Begin VB.CommandButton CmdMacSig 
         Caption         =   "Sig"
         Height          =   255
         Index           =   2
         Left            =   9360
         TabIndex        =   56
         Top             =   4200
         Width           =   375
      End
      Begin VB.CommandButton CmdMacSig 
         Caption         =   "Sig"
         Height          =   255
         Index           =   1
         Left            =   9360
         TabIndex        =   55
         Top             =   2760
         Width           =   375
      End
      Begin VB.CommandButton CmdMacSig 
         Caption         =   "Sig"
         Height          =   255
         Index           =   0
         Left            =   9360
         TabIndex        =   54
         Top             =   1440
         Width           =   375
      End
      Begin VB.CommandButton CMDSIG 
         Height          =   255
         Index           =   0
         Left            =   10080
         Picture         =   "FrmFemeaPad.frx":05FA
         Style           =   1  'Graphical
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   24
         Left            =   9600
         MaxLength       =   1
         TabIndex        =   51
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton CmdExpCau 
         Caption         =   "Exportar"
         Height          =   255
         Index           =   2
         Left            =   8520
         TabIndex        =   50
         Top             =   4200
         Width           =   735
      End
      Begin VB.CommandButton CmdExpCau 
         Caption         =   "Exportar"
         Height          =   255
         Index           =   1
         Left            =   8520
         TabIndex        =   49
         Top             =   2760
         Width           =   735
      End
      Begin VB.CommandButton CmdRenumerar 
         Caption         =   "Renumerar"
         Height          =   255
         Index           =   3
         Left            =   9720
         TabIndex        =   48
         Top             =   5640
         Width           =   1095
      End
      Begin VB.CommandButton CmdRenumerar 
         Caption         =   "Renumerar"
         Height          =   255
         Index           =   2
         Left            =   9840
         TabIndex        =   47
         Top             =   4200
         Width           =   975
      End
      Begin VB.CommandButton CmdRenumerar 
         Caption         =   "Renumerar"
         Height          =   255
         Index           =   1
         Left            =   9840
         TabIndex        =   46
         Top             =   2760
         Width           =   975
      End
      Begin VB.CommandButton CmdRenumerar 
         Caption         =   "Renumerar"
         Height          =   255
         Index           =   0
         Left            =   9840
         TabIndex        =   45
         Top             =   1440
         Width           =   975
      End
      Begin VB.CommandButton CmdSev 
         Caption         =   "Troca Severidade"
         Height          =   255
         Left            =   8400
         TabIndex        =   44
         Top             =   4560
         Width           =   1455
      End
      Begin VB.CommandButton CmdRem 
         Caption         =   "Renomear"
         Height          =   255
         Index           =   3
         Left            =   9720
         TabIndex        =   43
         Top             =   5280
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Escolher"
         Height          =   255
         Left            =   8520
         TabIndex        =   42
         Top             =   5640
         Width           =   1095
      End
      Begin VB.CommandButton CmdRem 
         Caption         =   "Renomear"
         Height          =   255
         Index           =   2
         Left            =   9840
         TabIndex        =   41
         Top             =   3840
         Width           =   975
      End
      Begin VB.CommandButton CmdRem 
         Caption         =   "Renomear"
         Height          =   255
         Index           =   1
         Left            =   9720
         TabIndex        =   40
         Top             =   2400
         Width           =   1095
      End
      Begin VB.CommandButton CmdRem 
         Caption         =   "Renomear"
         Height          =   255
         Index           =   0
         Left            =   9720
         TabIndex        =   39
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton CmdExpCau 
         Caption         =   "Exportar"
         Height          =   255
         Index           =   0
         Left            =   8520
         TabIndex        =   38
         Top             =   1440
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Localizar"
         Height          =   375
         Left            =   -66480
         TabIndex        =   37
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CommandButton CmdIncPf 
         Caption         =   "-->"
         Height          =   375
         Left            =   -68640
         TabIndex        =   36
         Top             =   1560
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Localizar"
         Height          =   375
         Left            =   -68640
         TabIndex        =   35
         Top             =   840
         Width           =   855
      End
      Begin VB.CommandButton CmdExcPFPad 
         Caption         =   "Excluir Processo"
         Height          =   615
         Index           =   0
         Left            =   -65040
         TabIndex        =   32
         Top             =   960
         Width           =   855
      End
      Begin VB.CommandButton CmdPreinc 
         Caption         =   "Novo"
         Height          =   255
         Left            =   -66480
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton CmddelPre 
         Caption         =   "Excluir"
         Height          =   255
         Left            =   -66480
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton cmdedipre 
         Caption         =   "Editar"
         Height          =   255
         Index           =   2
         Left            =   -66480
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton Retornar 
         Caption         =   "Retornar"
         Height          =   375
         Index           =   2
         Left            =   -65640
         TabIndex        =   27
         Top             =   6240
         Width           =   1455
      End
      Begin VB.CommandButton Retornar 
         Caption         =   "Retornar"
         Height          =   375
         Index           =   1
         Left            =   -65640
         TabIndex        =   26
         Top             =   6240
         Width           =   1455
      End
      Begin VB.CommandButton CmdNovo 
         Caption         =   "Novo"
         Height          =   255
         Index           =   0
         Left            =   8520
         TabIndex        =   13
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton CmdNovo 
         Caption         =   "Novo"
         Height          =   255
         Index           =   1
         Left            =   8520
         TabIndex        =   12
         Top             =   2040
         Width           =   1095
      End
      Begin VB.CommandButton CmdNovo 
         Caption         =   "Novo"
         Height          =   255
         Index           =   2
         Left            =   8520
         TabIndex        =   11
         Top             =   3480
         Width           =   1095
      End
      Begin VB.CommandButton CmdNovo 
         Caption         =   "Novo"
         Height          =   255
         Index           =   3
         Left            =   8520
         TabIndex        =   10
         Top             =   4920
         Width           =   1095
      End
      Begin VB.CommandButton CmdEscPro 
         Caption         =   "Escolher"
         Height          =   255
         Left            =   8520
         TabIndex        =   9
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton CmdEscFal 
         Caption         =   "Escolher"
         Height          =   255
         Left            =   8520
         TabIndex        =   8
         Top             =   2400
         Width           =   1095
      End
      Begin VB.CommandButton Cmdescefe 
         Caption         =   "Escolher"
         Height          =   255
         Left            =   8520
         TabIndex        =   7
         Top             =   3840
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Editar"
         Height          =   255
         Left            =   8520
         TabIndex        =   6
         Top             =   5280
         Width           =   1095
      End
      Begin VB.CommandButton Retornar 
         Caption         =   "Retornar"
         Height          =   375
         Index           =   0
         Left            =   10080
         TabIndex        =   5
         Top             =   6600
         Width           =   735
      End
      Begin VB.CommandButton CmdExcCau 
         Caption         =   "Excluir"
         Height          =   255
         Left            =   9720
         TabIndex        =   4
         Top             =   4920
         Width           =   1095
      End
      Begin VB.CommandButton CmdExcEfe 
         Caption         =   "Excluir"
         Height          =   255
         Left            =   9840
         TabIndex        =   3
         Top             =   3480
         Width           =   975
      End
      Begin VB.CommandButton CmdExcFal 
         Caption         =   "Excluir"
         Height          =   255
         Left            =   9720
         TabIndex        =   2
         Top             =   2040
         Width           =   1095
      End
      Begin VB.CommandButton CmdExcPro 
         Caption         =   "Excluir"
         Height          =   255
         Left            =   9720
         TabIndex        =   1
         Top             =   720
         Width           =   1095
      End
      Begin MSFlexGridLib.MSFlexGrid GridPRO 
         Height          =   1155
         Left            =   360
         TabIndex        =   14
         Top             =   600
         Width           =   7995
         _ExtentX        =   14102
         _ExtentY        =   2037
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid GRIDFAL 
         Height          =   1155
         Left            =   360
         TabIndex        =   15
         Top             =   2040
         Width           =   7995
         _ExtentX        =   14102
         _ExtentY        =   2037
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid gridefe 
         Height          =   1155
         Left            =   360
         TabIndex        =   16
         Top             =   3480
         Width           =   7995
         _ExtentX        =   14102
         _ExtentY        =   2037
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid gridcau 
         Height          =   1035
         Left            =   360
         TabIndex        =   17
         Top             =   4920
         Width           =   7995
         _ExtentX        =   14102
         _ExtentY        =   1826
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid Gridfemprecad 
         Height          =   5895
         Left            =   -74880
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   600
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   10398
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid GridPadPF 
         Height          =   5235
         Left            =   -67680
         TabIndex        =   33
         Top             =   840
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   9234
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid gridpf 
         Height          =   5355
         Left            =   -74880
         TabIndex        =   34
         Top             =   840
         Width           =   6195
         _ExtentX        =   10927
         _ExtentY        =   9446
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid gridrevcau 
         Height          =   1035
         Left            =   360
         TabIndex        =   61
         Top             =   6000
         Width           =   7995
         _ExtentX        =   14102
         _ExtentY        =   1826
         _Version        =   393216
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   600
         Left            =   -65640
         TabIndex        =   68
         Top             =   480
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   1058
         ButtonWidth     =   609
         ButtonHeight    =   953
         Appearance      =   1
         TextAlignment   =   1
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid Grid 
         Height          =   5595
         Left            =   -74820
         TabIndex        =   69
         Top             =   840
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   9869
         _Version        =   393216
      End
      Begin VB.Label dizfal 
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   1
         Left            =   -69480
         TabIndex        =   59
         Top             =   480
         Width           =   5175
      End
      Begin VB.Label dizpro 
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   1
         Left            =   -74880
         TabIndex        =   58
         Top             =   480
         Width           =   5175
      End
      Begin VB.Label lblLabels 
         Caption         =   "Siginificativa"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   14
         Left            =   8520
         TabIndex        =   53
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Funcao do Processo Requisitos"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   360
         TabIndex        =   25
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Modo de Falha Potencial"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   360
         TabIndex        =   24
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "Efeito Potencial da Falha"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   360
         TabIndex        =   23
         Top             =   3240
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Causa(s) Mecanismo(s) Potencial(ais)"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   360
         TabIndex        =   22
         Top             =   4680
         Width           =   2775
      End
      Begin VB.Label dizpro 
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   0
         Left            =   2760
         TabIndex        =   21
         Top             =   360
         Width           =   5535
      End
      Begin VB.Label dizfal 
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   0
         Left            =   2640
         TabIndex        =   20
         Top             =   1800
         Width           =   5655
      End
      Begin VB.Label dizefe 
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2400
         TabIndex        =   19
         Top             =   3240
         Width           =   5895
      End
      Begin VB.Label dizcau 
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3240
         TabIndex        =   18
         Top             =   4680
         Width           =   5055
      End
   End
End
Attribute VB_Name = "FrmFemeaPad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private aNUM As Variant
Private aDIZ As Variant
Private cARQFEMEA As String
Private cARQPF As String
Private nBARPOS
Private nROWREC As Long
Dim cARQUSO As String
Dim aORDEM As Variant
Dim aORDES As Variant
Dim cORDEM As String
Dim cSUBWHERE As String

Private Sub CmddelPre_Click()
    Dim nSEQ As Integer
    Dim cSQL As String
    If Gridfemprecad.Row > 0 Then
        Gridfemprecad.Col = 0
        nSEQ = Gridfemprecad
        cSQL = "select * from FEMPRE WHERE SEGGRA=" & nSEQ
        If ApagaSQLP(cARQFEMEA, cSQL) Then
            FILGRIdfemprecad
        End If
    End If
End Sub

Private Sub cmdedipre_Click(Index As Integer)
    Dim nSEQ As Integer
    If Gridfemprecad.Row > 0 Then
        Gridfemprecad.Col = 0
        nSEQ = Gridfemprecad
        ePASS01 = "select * from FEMPRE WHERE SEGGRA=" & nSEQ
        ePASS02 = cARQFEMEA
        frmFEMEI.Show vbModal, Me
        FILGRIdfemprecad
    End If
End Sub

Private Sub CmdEdiRevi_Click()
    Dim nREVI As Long
    If gridrevcau.Rows < 2 Then
        Exit Sub
    End If
    gridrevcau.Col = 4
    nREVI = FixInt(gridrevcau)
    ePASS01 = "select * from FEMREVI WHERE TIPOAPU='P' AND PRONUM=" & aNUM(0) & " AND FALNUM=" & aNUM(1) & " AND EFENUM=" & aNUM(2) & " AND CAUNUM=" & aNUM(3) & " AND SEGGRA=" & nREVI
    ePASS02 = cARQFEMEA
    Load frmFEMEI
    frmFEMEI.TXTFIELDS(2).Enabled = False
    frmFEMEI.TXTFIELDS(3).Enabled = False
    frmFEMEI.TXTFIELDS(4).Enabled = False
    frmFEMEI.TXTFIELDS(5).Enabled = False
    frmFEMEI.cmdClose.Enabled = False
    frmFEMEI.CmdAltman.Enabled = False
    frmFEMEI.Show vbModal, Me
End Sub

Private Sub Cmdescefe_Click()
    Dim nNUM As Integer
    Dim cDIZ As String
    gridefe.Col = 2
    nNUM = FixInt(gridefe)
    aNUM(2) = nNUM
    aNUM(3) = 0
    gridefe.Col = 3
    cDIZ = FixStr(gridefe)
    aDIZ(2) = cDIZ
    aDIZ(3) = ""
    fillall (4)
    LocalizaGrid gridefe, aNUM(2), 3, False, 0
End Sub

Private Sub CmdEscFal_Click()
    Dim nNUM As Integer
    Dim cDIZ As String
    GRIDFAL.Col = 1
    nNUM = FixInt(GRIDFAL)
    aNUM(1) = nNUM
    aNUM(2) = 0
    aNUM(3) = 0
    GRIDFAL.Col = 2
    cDIZ = FixStr(GRIDFAL)
    aDIZ(1) = cDIZ
    aDIZ(2) = ""
    aDIZ(3) = ""
    fillall (3)
    LocalizaGrid GRIDFAL, aNUM(1), 2, False, 0
    filgridpadpf
End Sub

Private Sub CmdEscPro_Click()
    Dim nNUM As Integer
    Dim cDIZ As String
    GridPRO.Col = 0
    nNUM = FixInt(GridPRO)
    aNUM(0) = nNUM
    aNUM(1) = 0
    aNUM(2) = 0
    aNUM(3) = 0
    GridPRO.Col = 1
    cDIZ = FixStr(GridPRO)
    aDIZ(0) = cDIZ
    aDIZ(1) = ""
    aDIZ(2) = ""
    aDIZ(3) = ""
    fillall (2)
    LocalizaGrid GridPRO, aNUM(0), 1, False, 0
    filgridpadpf

End Sub

Private Sub fillall(Optional ByVal ninicio As Integer = 0)

    If ninicio < 1 Then FILGRIdPRO
    If ninicio < 3 Then FILGRIdFAL
    If ninicio < 4 Then FILGRIdEFE
    If ninicio < 5 Then FILGRIdCAU
    dizpro(0).Caption = CStr(aNUM(0)) & " - " & aDIZ(0)
    dizpro(1).Caption = CStr(aNUM(0)) & " - " & aDIZ(0)
    dizfal(0).Caption = CStr(aNUM(1)) & " - " & aDIZ(1)
    dizfal(1).Caption = CStr(aNUM(1)) & " - " & aDIZ(1)
    dizefe.Caption = CStr(aNUM(2)) & " - " & aDIZ(2)
    dizcau.Caption = CStr(aNUM(3)) & " - " & aDIZ(3)

End Sub

Private Sub CmdExcCau_Click()

    Dim aVAL As Variant
    Dim cSQL As String

    If Not MDG("Este Processo Excluira Femea e Itens Continuar") Then
        Exit Sub
    End If
    aVAL = Array(0, 0, 0, 0)
    gridcau.Col = 0
    aVAL(0) = gridcau
    gridcau.Col = 1
    aVAL(1) = gridcau
    gridcau.Col = 2
    aVAL(2) = gridcau
    gridcau.Col = 3
    aVAL(3) = gridcau
    cSQL = "select * from FEMCAU WHERE PRONUM=" & aVAL(0) & " AND FALNUM=" & aVAL(1) & " AND EFENUM=" & aVAL(2) & " AND CAUNUM=" & aVAL(3)
    If ApagaSQLP(cARQFEMEA, cSQL) Then
        'cSQL = "select * from FEMEA WHERE PRONUM=" & aVAL(0) & " AND FALNUM=" & aVAL(1) & " AND EFENUM=" & aVAL(2) & " AND CAUNUM=" & aVAL(3)
        cSQL = "select * from FEMEAPAD WHERE PRONUM=" & aVAL(0) & " AND FALNUM=" & aVAL(1) & " AND EFENUM=" & aVAL(2) & " AND CAUNUM=" & aVAL(3)
        ApagaSQL cARQFEMEA, cSQL
        cSQL = "select * from FEMRPNT WHERE PRONUM=" & aVAL(0) & " AND FALNUM=" & aVAL(1) & " AND EFENUM=" & aVAL(2) & " AND CAUNUM=" & aVAL(3)
        ApagaSQL cARQFEMEA, cSQL
        fillall (4)
    End If

End Sub

Private Sub CmdExcEfe_Click()

    Dim aVAL As Variant
    Dim cSQL As String

    If Not MDG("Este Processo Excluira Femea e Itens Continuar") Then
        Exit Sub
    End If
    aVAL = Array(0, 0, 0, 0)
    gridefe.Col = 0
    aVAL(0) = gridefe
    gridefe.Col = 1
    aVAL(1) = gridefe
    gridefe.Col = 2
    aVAL(2) = gridefe
    cSQL = "select * from FEMEFE WHERE PRONUM=" & aVAL(0) & " AND FALNUM=" & aVAL(1) & " AND EFENUM=" & aVAL(2)
    If ApagaSQLP(cARQFEMEA, cSQL) Then
        cSQL = "select * from FEMCAU WHERE PRONUM=" & aVAL(0) & " AND FALNUM=" & aVAL(1) & " AND EFENUM=" & aVAL(2)
        ApagaSQL cARQFEMEA, cSQL
        'cSQL = "select * from FEMEA WHERE PRONUM=" & aVAL(0) & " AND FALNUM=" & aVAL(1) & " AND EFENUM=" & aVAL(2)
        cSQL = "select * from FEMEAPAD WHERE PRONUM=" & aVAL(0) & " AND FALNUM=" & aVAL(1) & " AND EFENUM=" & aVAL(2)
        ApagaSQL cARQFEMEA, cSQL
        cSQL = "select * from FEMRPNT WHERE PRONUM=" & aVAL(0) & " AND FALNUM=" & aVAL(1) & " AND EFENUM=" & aVAL(2)
        ApagaSQL cARQFEMEA, cSQL
        fillall (3)
    End If

End Sub

Private Sub CmdExcFal_Click()

    Dim aVAL As Variant
    Dim cSQL As String

    If Not MDG("Este Processo Excluira Femea e Itens Continuar") Then
        Exit Sub
    End If
    aVAL = Array(0, 0, 0, 0)
    GRIDFAL.Col = 0
    aVAL(0) = GRIDFAL
    GRIDFAL.Col = 1
    aVAL(1) = GRIDFAL
    cSQL = "select * from FEMFAL WHERE PRONUM=" & aVAL(0) & " AND FALNUM=" & aVAL(1)
    If ApagaSQLP(cARQFEMEA, cSQL) Then
        cSQL = "select * from FEMEFE WHERE PRONUM=" & aVAL(0) & " AND FALNUM=" & aVAL(1)
        ApagaSQL cARQFEMEA, cSQL
        cSQL = "select * from FEMCAU WHERE PRONUM=" & aVAL(0) & " AND FALNUM=" & aVAL(1)
        ApagaSQL cARQFEMEA, cSQL
        cSQL = "select * from FEMPF WHERE PRONUM=" & aVAL(0) & " AND FALNUM=" & aVAL(1)
        ApagaSQL cARQFEMEA, cSQL
        'cSQL = "select * from FEMEA WHERE PRONUM=" & aVAL(0) & " AND FALNUM=" & aVAL(1)
        cSQL = "select * from FEMEAPAD WHERE PRONUM=" & aVAL(0) & " AND FALNUM=" & aVAL(1)
        ApagaSQL cARQFEMEA, cSQL
        cSQL = "select * from FEMRPNT WHERE PRONUM=" & aVAL(0) & " AND FALNUM=" & aVAL(1)
        ApagaSQL cARQFEMEA, cSQL
        fillall (2)
    End If

End Sub

Private Sub CmdExcPFPad_Click(Index As Integer)

    Dim aVAL As Variant
    Dim cSQL As String
    aVAL = Array(0, 0, 0)

    GridPadPF.Col = 0
    aVAL(0) = FixInt(GridPadPF)
    GridPadPF.Col = 1
    aVAL(1) = FixInt(GridPadPF)
    GridPadPF.Col = 2
    aVAL(2) = FixInt(GridPadPF)
    If Index = 0 Then
        cSQL = "select * from FEMPF WHERE PRONUM=" & aVAL(1) & " AND PF=" & aVAL(0)
    Else
        cSQL = "select * from FEMPF WHERE PRONUM=" & aVAL(1) & " AND FALNUM=" & aVAL(2) & " AND PF=" & aVAL(0)
    End If
    If ApagaSQLP(cARQFEMEA, cSQL) Then
        If Index = 0 Then
            ' cSQL = "select * from FEMEA WHERE PRONUM=" & aVAL(1) & " AND PF=" & aVAL(0)
            cSQL = "select * from FEMEAPAD WHERE PRONUM=" & aVAL(1) & " AND PF=" & aVAL(0)
            ApagaSQL cARQFEMEA, cSQL
            cSQL = "select * from FEMRPNO WHERE PRONUM=" & aVAL(1) & " AND PF=" & aVAL(0)
            ApagaSQL cARQFEMEA, cSQL
            cSQL = "select * from FEMRPNT WHERE PRONUM=" & aVAL(1) & " And PF = " & aVAL(0)
            ApagaSQL cARQFEMEA, cSQL
        Else
            'cSQL = "select * from FEMEA WHERE PRONUM=" & aVAL(1) & " AND FALNUM=" & aVAL(2) & " AND PF=" & aVAL(0)
            cSQL = "select * from FEMEAPAD WHERE PRONUM=" & aVAL(1) & " AND FALNUM=" & aVAL(2) & " AND PF=" & aVAL(0)
            ApagaSQL cARQFEMEA, cSQL
            cSQL = "select * from FEMRPNT WHERE PRONUM=" & aVAL(1) & " AND FALNUM=" & aVAL(2) & " AND PF=" & aVAL(0)
            ApagaSQL cARQFEMEA, cSQL
        End If
        filgridpadpf
        Alert ("Exclusao Concluida")
    End If

End Sub

Private Sub CmdExcPFtot_Click()

    Dim cSQL As String

    If Not MDG("Isto Ira Apagar todos os Femeas do Processo ") Then
        Exit Sub
    End If
    cSQL = "select * from FEMPF WHERE PRONUM=" & aNUM(0)
    If ApagaSQLP(cARQFEMEA, cSQL) Then
        'cSQL = "select * from FEMEA WHERE PRONUM=" & aNUM(0)
        cSQL = "select * from FEMEAPAD WHERE PRONUM=" & aNUM(0)
        ApagaSQL cARQFEMEA, cSQL
    End If

End Sub

Private Sub CmdExcPro_Click()

    Dim aVAL As Variant
    Dim cSQL As String

    If Not MDG("Este Processo Excluira Femea e Itens Continuar") Then
        Exit Sub
    End If
    aVAL = Array(0, 0, 0, 0)
    GridPRO.Col = 0
    aVAL(0) = GridPRO
    cSQL = "select * from FEMPRO WHERE PRONUM=" & aVAL(0)
    If ApagaSQLP(cARQFEMEA, cSQL) Then
        cSQL = "select * from FEMFAL WHERE PRONUM=" & aVAL(0)
        ApagaSQL cARQFEMEA, cSQL
        cSQL = "select * from FEMEFE WHERE PRONUM=" & aVAL(0)
        ApagaSQL cARQFEMEA, cSQL
        cSQL = "select * from FEMCAU WHERE PRONUM=" & aVAL(0)
        ApagaSQL cARQFEMEA, cSQL
        cSQL = "select * from FEMPF WHERE PRONUM=" & aVAL(0)
        ApagaSQL cARQFEMEA, cSQL
        'cSQL = "select * from FEMEA WHERE PRONUM=" & aVAL(0)
        cSQL = "select * from FEMEAPAD WHERE PRONUM=" & aVAL(0)
        ApagaSQL cARQFEMEA, cSQL
        cSQL = "select * from FEMRPNT WHERE PRONUM=" & aVAL(0)
        ApagaSQL cARQFEMEA, cSQL
        cSQL = "select * from FEMRPNO WHERE PRONUM=" & aVAL(0)
        ApagaSQL cARQFEMEA, cSQL
        fillall
    End If

End Sub

Private Sub CmdExpCau_Click(Index As Integer)
    Dim DB As New ADODB.Connection
    Dim DB2 As New ADODB.Connection
    Dim DB3 As New ADODB.Connection
    Dim RSTAB As ADODB.Recordset
    Dim RSTA2 As ADODB.Recordset
    Dim RSTA3 As ADODB.Recordset
    Dim nPRODES
    Dim nFALDES
    Dim nEFEDES
    Dim nCAUDES As Integer
    Dim cSQLORI
    Dim cSQL2
    Dim cSQL3
    Dim cDIZPRO
    Dim cDIZFAL
    Dim cDIZEFE
    Dim cDIZEXP As String
    Dim cSQLPFPAD
    Dim oRSPFPAD As New ADODB.Recordset
    On Error Resume Next

    If Index > 0 Then
        If Not checkval(Index - 1) Then          ''precisa somente o anterior
            Exit Sub
        End If
    End If
    nPRODES = FixNum(Val(Busca("Nº Processo Destino", "Exportando", CStr(aNUM(0)), 8)))
    cDIZPRO = PegUltSQL(cARQFEMEA, "select * from FEMPRO WHERE PRONUM=" & nPRODES, "PROCESSO", "")
    cDIZPRO = FixStr(InputBox("Digite a Descricao Processo", "Exportando", cDIZPRO), "")
    If Index > 0 Then
        nFALDES = FixNum(Val(Busca("Nº Falha Destino", "Exportando", CStr(aNUM(1)), 8)))
        cDIZFAL = PegUltSQL(cARQFEMEA, "select * from FEMFAL WHERE PRONUM=" & nPRODES & " AND FALNUM=" & nFALDES, "FALTIP", "")
        cDIZFAL = FixStr(InputBox("Digite a Descricao Falha", "Exportando", cDIZFAL), "")
    End If
    If Index > 1 Then
        nEFEDES = FixNum(Val(Busca("Nº Efeito Destino", "Exportando", CStr(aNUM(2) + 1), 8)))
        cDIZEFE = PegUltSQL(cARQFEMEA, "select * from FEMEFE WHERE PRONUM=" & nPRODES & " AND FALNUM=" & nFALDES & " AND EFENUM=" & nEFEDES, "FALEFE", "")
        cDIZEFE = FixStr(InputBox("Digite a Descricao Efeito", "Exportando", cDIZEFE), "")
    End If
    Select Case Index
    Case 0
        cDIZEXP = "Exporta Para " & CStr(nPRODES)
    Case 1
        cDIZEXP = "Exporta Para " & CStr(nPRODES) & " - " & CStr(nFALDES)
    Case 2
        cDIZEXP = "Exporta Para " & CStr(nPRODES) & " - " & CStr(nFALDES) & " - " & CStr(nEFEDES)
    End Select
    If Not MDG(cDIZEXP) Then
        Exit Sub
    End If
    ''Gerando Itens
    'sequencia duplicacao
    DB.ConnectionTimeout = 120
    DB.Open GeracArq(cARQPF)
    RSTAB.Open "select * from DUPLICAR WHERE TABELA='FEMEA'", DB, adOpenForwardOnly, adLockReadOnly
    'arquivo destino
    DB2.ConnectionTimeout = 120
    DB2.Open GeracArq(cARQFEMEA)
    'origem duplicaçao
    Select Case Index
    Case 0
        cSQLORI = "select * from FEMCAU WHERE PRONUM=" & aNUM(0)
    Case 1
        cSQLORI = "select * from FEMCAU WHERE PRONUM=" & aNUM(0) & " AND FALNUM=" & aNUM(1)
    Case 2
        cSQLORI = "select * from FEMCAU WHERE PRONUM=" & aNUM(0) & " AND FALNUM=" & aNUM(1) & " AND EFENUM=" & aNUM(2)
    End Select
    DB3.ConnectionTimeout = 120
    DB3.Open GeracArq(cARQFEMEA)
    RSTA3.Open cSQLORI, DB3, adOpenForwardOnly, adLockReadOnly
    If Not RSTA3.EOF Then
        Barra.Value = 0
        nBARPOS = 0
        nROWREC = RSTA3.RecordCount
    End If
    Do While Not RSTA3.EOF
        BARPOS
        DIZAPU = FixStr(RSTA3("CAUNUM"))
        If Index = 0 Then                        'Processo
            nFALDES = FixInt(RSTA3("FALNUM"))
            nEFEDES = FixInt(RSTA3("EFENUM"))
            cDIZFAL = RSTA3("FALTIP")
            cDIZEFE = RSTA3("FALEFE")
        End If
        If Index = 1 Then
            nEFEDES = FixInt(RSTA3("EFENUM"))
            cDIZEFE = RSTA3("FALEFE")
        End If
        nCAUDES = FixInt(RSTA3("CAUNUM"))
        cSQL2 = "select * from FEMCAU WHERE PRONUM=" & nPRODES & " AND FALNUM=" & nFALDES & " AND EFENUM=" & nEFEDES & " AND CAUNUM=" & nCAUDES
        RSTA2.Open cSQL2, DB2, adOpenDynamic, adLockOptimistic
        If RSTA2.EOF Then
            RSTA2.AddNew
            RSTA2("PRONUM") = nPRODES
            RSTA2("FALNUM") = nFALDES
            RSTA2("EFENUM") = nEFEDES
            RSTA2("CAUNUM") = nCAUDES
            RSTAB.MoveFirst
            Do While Not RSTAB.EOF
                ''LABEL.Caption = rsTAB("CAMPO")
                RSTA2(RSTAB("CAMPO")) = RSTA3(RSTAB("CAMPO"))
                RSTAB.MoveNext
            Loop
            ''Depois se nao grava antigo
            RSTA2("PROCESSO") = cDIZPRO
            If Index > 0 Then
                RSTA2("FALTIP") = cDIZFAL
            End If
            If Index > 1 Then
                RSTA2("FALEFE") = cDIZEFE
            End If
            RSTA2.Update
        End If
        If Index = 0 Then
            cSQL3 = "select * from FEMFAL WHERE PRONUM=" & nPRODES & " AND FALNUM=" & nFALDES
            IncluiSQL cARQFEMEA, cSQL3, 3, Array("FALNUM", "FALTIP", "PRONUM"), Array(nFALDES, cDIZFAL, nPRODES), True, False
        End If
        If Index = 0 Or Index = 1 Then
            cSQL3 = "select * from FEMEFE WHERE PRONUM=" & nPRODES & " AND FALNUM=" & nFALDES & " AND EFENUM=" & nEFEDES
            IncluiSQL cARQFEMEA, cSQL3, 4, Array("EFENUM", "FALEFE", "PRONUM", "FALNUM"), Array(nEFEDES, cDIZEFE, nPRODES, nFALDES), True, False
        End If
        RSTA3.MoveNext
    Loop
    If Index = 2 Then                            ''Necessita gerar os PF PADROES
        cSQLPFPAD = "SELECT * FROM FEMPF WHERE PRONUM=" & nPRODES & " AND FALNUM=" & nFALDES & " ORDER BY PF"
        oRSPFPAD.Open cSQLPFPAD, DB2, adOpenForwardOnly, adLockReadOnly
        If Not oRSPFPAD.EOF Then
            Barra.Value = 0
            nBARPOS = 0
            nROWREC = oRSPFPAD.RecordCount
        End If
        Do While Not oRSPFPAD.EOF
            nPF = FixInt(oRSPFPAD("PF"))
            RSTA3.MoveFirst
            Do While Not RSTA3.EOF
                BARPOS
                DIZAPU = FixStr(nPF) & "/" & FixStr(RSTA3("CAUNUM"))
                nCAUDES = FixInt(RSTA3("CAUNUM"))
                'cSQL2 = "select * from FEMEA WHERE PRONUM=" & nPRODES & " AND FALNUM=" & nFALDES & " AND EFENUM=" & nEFEDES & " AND CAUNUM=" & nCAUDES
                cSQL2 = "select * from FEMEAPAD WHERE PRONUM=" & nPRODES & " AND FALNUM=" & nFALDES & " AND EFENUM=" & nEFEDES & " AND CAUNUM=" & nCAUDES
                RSTA2.Open cSQL2, DB2, adOpenDynamic, adLockOptimistic
                If RSTA2.EOF Then
                    RSTA2.AddNew
                    RSTA2("PRONUM") = nPRODES
                    RSTA2("FALNUM") = nFALDES
                    RSTA2("EFENUM") = nEFEDES
                    RSTA2("CAUNUM") = nCAUDES
                    RSTA2("PF") = nPF
                    RSTAB.MoveFirst
                    Do While Not RSTAB.EOF
                        ''LABEL.Caption = rsTAB("CAMPO")
                        RSTA2(RSTAB("CAMPO")) = RSTA3(RSTAB("CAMPO"))
                        RSTAB.MoveNext
                    Loop
                    ''Depois se nao grava antigo
                    RSTA2("PROCESSO") = cDIZPRO
                    RSTA2("FALTIP") = cDIZFAL
                    RSTA2("FALEFE") = cDIZEFE
                    RSTA2.Update
                End If
                RSTA3.MoveNext
            Loop
            oRSPFPAD.MoveNext
        Loop
    End If
    RSTAB.Close
    RSTA2.Close
    RSTA3.Close
    Set DB = Nothing
    Set RSTAB = Nothing
    Set DB2 = Nothing
    Set RSTA2 = Nothing
    Set DB3 = Nothing
    Set RSTA3 = Nothing
    fillall (3)

End Sub

Private Sub CmdExpRPN_Click()
    Dim cSQL As String
    Dim aVAL As Variant
    Dim nANO As Integer
    Dim nMES As Integer
    Dim oDB As New ADODB.Connection
    Dim oRS As New ADODB.Recordset

    If Not checkval(3) Then
        Exit Sub
    End If
    frmcomp.Show vbModal, Me
    If Not MDG("Exportar para " & CStr(eRETU01) & "/" & CStr(eRETU02)) Then
        Exit Sub
    End If
    cSQL = "select * from FEMCAU WHERE PRONUM=" & aNUM(0) & " AND FALNUM=" & aNUM(1) & " AND EFENUM=" & aNUM(2) & " AND CAUNUM=" & aNUM(3)
    aVAL = PegSQL(cARQFEMEA, cSQL, 4, Array("RINDOCO", "RINDSER", "RINDDET", "RINDRIS"), Array("NI", "NI", "NI", "NI"), Array(0, 0, 0, 0))
    aVAL(0) = FixInt(aVAL(0))
    aVAL(1) = FixInt(aVAL(1))
    aVAL(2) = FixInt(aVAL(2))
    aVAL(3) = FixInt(aVAL(3))
    nMES = FixInt(eRETU02)
    nANO = FixInt(eRETU01)
    cSQL = "select * from FEMRPNT WHERE PRONUM=" & aNUM(0) & " AND FALNUM=" & aNUM(1) & " AND EFENUM=" & aNUM(2) & " AND CAUNUM=" & aNUM(3) & " AND MES=" & nMES & " AND ANO=" & nANO
    oDB.ConnectionTimeout = 120
    oDB.Open GeracArq(cARQFEMEA)
    oRS.Open cSQL, oDB, adOpenDynamic, adLockOptimistic
    Do While Not oRS.EOF
        oRS("RINDoco") = aVAL(0)
        oRS("RINDser") = aVAL(1)
        oRS("RINDdet") = aVAL(2)
        oRS("RINDris") = aVAL(3)
        oRS.Update
        oRS.MoveNext
    Loop
    oRS.Close
    oDB.Close
    Set oDB = Nothing
    Set oRS = Nothing

End Sub

Private Sub CmdIncPf_Click()

    Dim nPFTMP As Integer

    If aNUM(0) = 0 Then
        Alert ("Processo Nao Escolhido")
        Exit Sub
    End If
    If aNUM(1) = 0 Then
        Alert ("Falha Nao Escolhido")
        Exit Sub
    End If
    gridpf.Col = 0
    nPFTMP = FixInt(gridpf)
    GERAFEMEAPAD aNUM(0), aNUM(1), nPFTMP
    filgridpadpf
    Alert ("Inclusao Concluida")

End Sub

Private Sub CmdMacSig_Click(Index As Integer)

    Dim cSQL
    Dim cSQL2
    Dim cSIG As String
    cSIG = TXTFIELDS(24)

    If Len(cSIG) = 0 Then
        Alert ("Siginificativa Nao Escolhida")
    End If
    If Index > 0 Then
        If Not checkval(Index) Then
            Exit Sub
        End If
    End If
    If Not MDG("Mudar para " & cSIG) Then
        Exit Sub
    End If
    If Not MDG("Isto Altera Padrao e FEMEA Continuar") Then
        Exit Sub
    End If
    Select Case Index
    Case 0
        cSQL = "select SIGI from FEMCAU WHERE PRONUM=" & aNUM(0)
        'cSQL2 = "select SIGI from FEMEA WHERE PRONUM=" & aNUM(0)
        cSQL2 = "select SIGI from FEMEAPAD WHERE PRONUM=" & aNUM(0)
    Case 1
        cSQL = "select SIGI from FEMCAU WHERE PRONUM=" & aNUM(0) & " AND FALNUM=" & aNUM(1)
        'cSQL2 = "select SIGI from FEMEA WHERE PRONUM=" & aNUM(0) & " AND FALNUM=" & aNUM(1)
        cSQL2 = "select SIGI from FEMEAPAD WHERE PRONUM=" & aNUM(0) & " AND FALNUM=" & aNUM(1)
    Case 2
        cSQL = "select SIGI from FEMCAU WHERE PRONUM=" & aNUM(0) & " AND FALNUM=" & aNUM(1) & " AND EFENUM=" & aNUM(2)
        'cSQL2 = "select SIGI from FEMEA WHERE PRONUM=" & aNUM(0) & " AND FALNUM=" & aNUM(1) & " AND EFENUM=" & aNUM(2)
        cSQL2 = "select SIGI from FEMEAPAD WHERE PRONUM=" & aNUM(0) & " AND FALNUM=" & aNUM(1) & " AND EFENUM=" & aNUM(2)
    End Select
    GrvSQL cARQFEMEA, cSQL, 1, Array("SIGI"), Array(cSIG), Array("C")
    GrvSQL cARQFEMEA, cSQL2, 1, Array("SIGI"), Array(cSIG), Array("C")

End Sub

Private Sub CmdNovo_Click(Index As Integer)

    Dim nNUM As Integer
    Dim cDIZ As String
    Dim cSQL As String
    Dim cCAMPO As String

    If Index > 0 Then
        If Not checkval(Index - 1) Then          ''menos um por que tem q ter a anterior
            Exit Sub
        End If
    End If
    Select Case Index
    Case 0
        cSQL = "FEMPRO"
        cCAMPO = "PRONUM"
    Case 1
        cSQL = "FEMFAL WHERE PRONUM=" & aNUM(0)
        cCAMPO = "FALNUM"
    Case 2
        cSQL = "FEMEFE WHERE PRONUM=" & aNUM(0) & " AND FALNUM=" & aNUM(1)
        cCAMPO = "EFENUM"
    Case 3
        cSQL = "FEMCAU WHERE PRONUM=" & aNUM(0) & " AND FALNUM=" & aNUM(1) & " AND EFENUM=" & aNUM(2)
        cCAMPO = "CAUNUM"
    End Select
    nNUM = FixInt(PegMAXSQL(cARQFEMEA, cSQL, cCAMPO, 0)) + 1
    nNUM = FixNum(Val(Busca("Nº", "Incluindo", CStr(nNUM), 8)))
    cDIZ = FixStr(InputBox("Digite a Descricao", "Incluindo", ""), "")
    If nNUM = 0 Then
        Alert ("Nº em Branco")
        Exit Sub
    End If
    If Len(cDIZ) = 0 Then
        Alert ("Dizer em Branco")
        Exit Sub
    End If
    Select Case Index
    Case 0
        ePASS01 = "select * from FEMPRO WHERE PRONUM=" & nNUM
        IncluiSQL cARQFEMEA, ePASS01, 2, Array("PRONUM", "PROCESSO"), Array(nNUM, cDIZ), True
        aNUM(0) = nNUM
        CmdEscPro_Click
    Case 1
        ePASS01 = "select * from FEMFAL WHERE PRONUM=" & aNUM(0) & " AND FALNUM=" & nNUM
        IncluiSQL cARQFEMEA, ePASS01, 3, Array("FALNUM", "FALTIP", "PRONUM"), Array(nNUM, cDIZ, aNUM(0)), True
        aNUM(1) = nNUM
        CmdEscFal_Click
    Case 2
        ePASS01 = "select * from FEMEFE WHERE PRONUM=" & aNUM(0) & " AND FALNUM=" & aNUM(1) & " AND EFENUM=" & nNUM
        IncluiSQL cARQFEMEA, ePASS01, 4, Array("EFENUM", "FALEFE", "PRONUM", "FALNUM"), Array(nNUM, cDIZ, aNUM(0), aNUM(1)), True
        aNUM(2) = nNUM
        Cmdescefe_Click
    Case 3
        ePASS01 = "select * from FEMCAU WHERE PRONUM=" & aNUM(0) & " AND FALNUM=" & aNUM(1) & " AND EFENUM=" & aNUM(2) & " AND CAUNUM=" & nNUM
        IncluiSQL cARQFEMEA, ePASS01, 8, Array("CAUNUM", "FALCAU", "PRONUM", "FALNUM", "EFENUM", "PROCESSO", "FALTIP", "FALEFE"), Array(nNUM, cDIZ, aNUM(0), aNUM(1), aNUM(2), aDIZ(0), aDIZ(1), aDIZ(2)), True
        aNUM(3) = nNUM
    End Select
    If Index = 3 Then
        ePASS01 = "select * from FEMCAU WHERE PRONUM=" & aNUM(0) & " AND FALNUM=" & aNUM(1) & " AND EFENUM=" & aNUM(2) & " AND CAUNUM=" & aNUM(3)
        ePASS02 = cARQFEMEA
        Load frmFEMEI
        frmFEMEI.TXTFIELDS(2).Enabled = False
        frmFEMEI.TXTFIELDS(3).Enabled = False
        frmFEMEI.TXTFIELDS(4).Enabled = False
        frmFEMEI.TXTFIELDS(5).Enabled = False
        frmFEMEI.Show vbModal, Me
        gerapfs
        Alert ("Indices Transferidos")
    End If
    fillall

End Sub

Private Sub gerapfs()
    Dim cSQL2 As String
    Dim cSQL3 As String
    Dim csql4 As String
    Dim nPFTMP As Long
    Dim cTMPPRO As String
    Dim cCAMPO As String

    Dim DB As ADODB.Connection
    Dim DB2 As ADODB.Connection
    Dim RSTAB As ADODB.Recordset
    Dim RSTA2 As ADODB.Recordset
    Dim RSTA3 As ADODB.Recordset
    Dim RSTA4 As ADODB.Recordset

    Set DB = New ADODB.Connection
    Set DB2 = New ADODB.Connection
    Set RSTAB = New ADODB.Recordset
    Set RSTA2 = New ADODB.Recordset
    Set RSTA3 = New ADODB.Recordset
    Set RSTA4 = New ADODB.Recordset


    ''Gerando Itens
    'sequencia duplicacao
    DB.ConnectionTimeout = 120
    DB.Open GeracArq(cARQPF)
    RSTAB.Open "select * from DUPLICAR WHERE TABELA='FEMEA'", DB, adOpenDynamic, adLockReadOnly
    
    'arquivo destino
    DB2.ConnectionTimeout = 120
    DB2.Open GeracArq(cARQFEMEA)
    
    ''Registro Recem Editado
    cSQL3 = "select * from FEMCAU WHERE PRONUM=" & aNUM(0) & " AND FALNUM=" & aNUM(1) & " AND EFENUM=" & aNUM(2) & " AND CAUNUM=" & aNUM(3)
    RSTA3.Open cSQL3, DB2, adOpenDynamic, adLockReadOnly
    
    ''PF Padrao
    csql4 = "select * from FEMPF WHERE PRONUM=" & aNUM(0) & " AND FALNUM=" & aNUM(1) & " ORDER BY PF"
    RSTA4.Open csql4, DB2, adOpenForwardOnly, adLockReadOnly
    
    Barra.Value = 0
    nBARPOS = 0
    nROWREC = RSTA4.RecordCount
    RSTA4.MoveFirst

    Do While Not RSTA4.EOF
        nPFTMP = RSTA4("PF")
        cTMPPRO = ""
        DIZAPU = FixStr(nPFTMP)
        BARPOS
        
        'cSQL2 = "select * from FEMEA WHERE PRONUM=" & aNUM(0) & " AND FALNUM=" & aNUM(1) & " AND EFENUM=" & aNUM(2) & " AND CAUNUM=" & aNUM(3) & " AND PF=" & nPFTMP
        cSQL2 = "select * from FEMEAPAD WHERE PRONUM=" & aNUM(0) & " AND FALNUM=" & aNUM(1) & " AND EFENUM=" & aNUM(2) & " AND CAUNUM=" & aNUM(3) & " AND PF=" & nPFTMP
        RSTA2.Open cSQL2, DB2, adOpenDynamic, adLockOptimistic
        
        If RSTA2.EOF Then
            RSTA2.AddNew
            RSTA2("PF") = nPFTMP
            RSTA2("PRONUM") = aNUM(0)
            RSTA2("FALNUM") = aNUM(1)
            RSTA2("EFENUM") = aNUM(2)
            RSTA2("CAUNUM") = aNUM(3)
        Else
            'RSTA2.Edit
            cTMPPRO = FixStr(RSTA2("PROCESSO"))
        End If
        RSTAB.MoveFirst
        Do While Not RSTAB.EOF
            cCAMPO = RSTAB("CAMPO")
            RSTA2(cCAMPO) = RSTA3(cCAMPO)
            RSTAB.MoveNext
        Loop
        If cTMPPRO <> "" Then
            RSTA2("PROCESSO") = cTMPPRO
        End If
        RSTA2.Update
        RSTA2.Close
        RSTA4.MoveNext
    Loop
    RSTAB.Close
    'RSTA2.Close
    RSTA3.Close
    RSTA4.Close
    Set DB = Nothing
    Set RSTAB = Nothing
    Set DB2 = Nothing
    Set RSTA2 = Nothing
    Set RSTA3 = Nothing
    Set RSTA4 = Nothing
    Alert "Concluido Atualização PF/FEMEA/PADRAO"
    
End Sub

Private Sub CmdPreinc_Click()

    Dim cTITULO As String
    cTITULO = FixStr(Busca("Digite o Titulo", "Titulo Pre-Cadastro Femea", Space(50), 50), "")

    If Len(cTITULO) = 0 Then
        Alert ("Titulo Necessario")
        Exit Sub
    End If
    IncluiSQL cARQFEMEA, "SELECT * FEMPRE WHERE TITULO='" & cTITULO & "'", 1, Array("TITULO"), Array(cTITULO), True
    FILGRIdfemprecad

End Sub

Private Sub CmdRem_Click(Index As Integer)

    Dim cDIZ
    Dim cSQL As String

    If Not checkval(Index) Then
        Exit Sub
    End If
    cDIZ = aDIZ(Index)
    cDIZ = FixStr(InputBox("Digite a Descricao", "Incluindo", ""), "")
    If Len(cDIZ) = 0 Then
        Alert ("Dizer em Branco")
        Exit Sub
    End If
    If Not MDG("Trocar Dizer") Then
        Exit Sub
    End If
    If Not MDG("Isto Altera Causas e FEMEA Continuar") Then
        Exit Sub
    End If
    Select Case Index
    Case 0
        cSQL = "select PROCESSO from FEMPRO WHERE PRONUM=" & aNUM(0)
        GrvSQL cARQFEMEA, cSQL, 1, Array("PROCESSO"), Array(cDIZ), Array("C")
        cSQL = "select PROCESSO from FEMCAU WHERE PRONUM=" & aNUM(0)
        GrvSQL cARQFEMEA, cSQL, 1, Array("PROCESSO"), Array(cDIZ), Array("C")
        'cSQL = "select PROCESSO from FEMEA WHERE PRONUM=" & aNUM(0)
        cSQL = "select PROCESSO from FEMEAPAD WHERE PRONUM=" & aNUM(0)
        GrvSQL cARQFEMEA, cSQL, 1, Array("PROCESSO"), Array(cDIZ), Array("C")
    Case 1
        cSQL = "select FALTIP from FEMFAL WHERE PRONUM=" & aNUM(0) & " AND FALNUM=" & aNUM(1)
        GrvSQL cARQFEMEA, cSQL, 1, Array("FALTIP"), Array(cDIZ), Array("C")
        cSQL = "select FALTIP from FEMCAU WHERE PRONUM=" & aNUM(0) & " AND FALNUM=" & aNUM(1)
        GrvSQL cARQFEMEA, cSQL, 1, Array("FALTIP"), Array(cDIZ), Array("C")
        'cSQL = "select FALTIP from FEMEA WHERE PRONUM=" & aNUM(0) & " AND FALNUM=" & aNUM(1)
        cSQL = "select FALTIP from FEMEAPAD WHERE PRONUM=" & aNUM(0) & " AND FALNUM=" & aNUM(1)
        GrvSQL cARQFEMEA, cSQL, 1, Array("FALTIP"), Array(cDIZ), Array("C")
    Case 2
        cSQL = "select FALEFE from FEMEFE WHERE PRONUM=" & aNUM(0) & " AND FALNUM=" & aNUM(1) & " AND EFENUM=" & aNUM(2)
        GrvSQL cARQFEMEA, cSQL, 1, Array("FALEFE"), Array(cDIZ), Array("C")
        cSQL = "select FALEFE from FEMCAU WHERE PRONUM=" & aNUM(0) & " AND FALNUM=" & aNUM(1) & " AND EFENUM=" & aNUM(2)
        GrvSQL cARQFEMEA, cSQL, 1, Array("FALEFE"), Array(cDIZ), Array("C")
        'cSQL = "select FALEFE from FEMEA WHERE PRONUM=" & aNUM(0) & " AND FALNUM=" & aNUM(1) & " AND EFENUM=" & aNUM(2)
        cSQL = "select FALEFE from FEMEAPAD WHERE PRONUM=" & aNUM(0) & " AND FALNUM=" & aNUM(1) & " AND EFENUM=" & aNUM(2)
        GrvSQL cARQFEMEA, cSQL, 1, Array("FALEFE"), Array(cDIZ), Array("C")
    Case 3
        cSQL = "select FALCAU from FEMCAU WHERE PRONUM=" & aNUM(0) & " AND FALNUM=" & aNUM(1) & " AND EFENUM=" & aNUM(2) & " AND CAUNUM=" & aNUM(3)
        GrvSQL cARQFEMEA, cSQL, 1, Array("FALCAU"), Array(cDIZ), Array("C")
        'cSQL = "select FALCAU from FEMEA WHERE PRONUM=" & aNUM(0) & " AND FALNUM=" & aNUM(1) & " AND EFENUM=" & aNUM(2) & " AND CAUNUM=" & aNUM(3)
        cSQL = "select FALCAU from FEMEAPAD WHERE PRONUM=" & aNUM(0) & " AND FALNUM=" & aNUM(1) & " AND EFENUM=" & aNUM(2) & " AND CAUNUM=" & aNUM(3)
        GrvSQL cARQFEMEA, cSQL, 1, Array("FALCAU"), Array(cDIZ), Array("C")
    End Select
    aNUM = Array(0, 0, 0, 0)
    aDIZ = Array("", "", "", "")
    fillall

End Sub

Private Sub CmdRenumerar_Click(Index As Integer)

    Dim nNUM As Integer
    Dim cSQL As String

    If Not checkval(Index) Then
        Exit Sub
    End If
    nNUM = FixInt(Val(Busca("Novo Numero", "Renumerando", CStr(nNUM), 8)))
    If nNUM = 0 Then
        Alert ("Novo Numero Em Branco")
        Exit Sub
    End If
    If Not MDG("Renumerar para " & CStr(nNUM)) Then
        Exit Sub
    End If
    If Not MDG("Isto Altera Padrao e FEMEA Continuar") Then
        Exit Sub
    End If
    Select Case Index
    Case 0
        cSQL = "select PRONUM from FEMPRO WHERE PRONUM=" & aNUM(0)
        GrvSQL cARQFEMEA, cSQL, 1, Array("PRONUM"), Array(nNUM), Array("NI")
        cSQL = "select PRONUM from FEMFAL WHERE PRONUM=" & aNUM(0)
        GrvSQL cARQFEMEA, cSQL, 1, Array("PRONUM"), Array(nNUM), Array("NI")
        cSQL = "select PRONUM from FEMEFE WHERE PRONUM=" & aNUM(0)
        GrvSQL cARQFEMEA, cSQL, 1, Array("PRONUM"), Array(nNUM), Array("NI")
        cSQL = "select PRONUM from FEMCAU WHERE PRONUM=" & aNUM(0)
        GrvSQL cARQFEMEA, cSQL, 1, Array("PRONUM"), Array(nNUM), Array("NI")
        'cSQL = "select PRONUM from FEMEA WHERE PRONUM=" & aNUM(0)
        cSQL = "select PRONUM from FEMEAPAD WHERE PRONUM=" & aNUM(0)
        GrvSQL cARQFEMEA, cSQL, 1, Array("PRONUM"), Array(nNUM), Array("NI")
        cSQL = "select PRONUM from FEMPF WHERE PRONUM=" & aNUM(0)
        GrvSQL cARQFEMEA, cSQL, 1, Array("PRONUM"), Array(nNUM), Array("NI")
        cSQL = "select PRONUM from FEMRPNT WHERE PRONUM=" & aNUM(0)
        GrvSQL cARQFEMEA, cSQL, 1, Array("PRONUM"), Array(nNUM), Array("NI")
        cSQL = "select PRONUM from FEMRPNO WHERE PRONUM=" & aNUM(0)
        GrvSQL cARQFEMEA, cSQL, 1, Array("PRONUM"), Array(nNUM), Array("NI")
    Case 1
        cSQL = "select FALNUM from FEMFAL WHERE PRONUM=" & aNUM(0) & " AND FALNUM=" & aNUM(1)
        GrvSQL cARQFEMEA, cSQL, 1, Array("FALNUM"), Array(nNUM), Array("NI")
        cSQL = "select FALNUM from FEMEFE WHERE PRONUM=" & aNUM(0) & " AND FALNUM=" & aNUM(1)
        GrvSQL cARQFEMEA, cSQL, 1, Array("FALNUM"), Array(nNUM), Array("NI")
        cSQL = "select FALNUM from FEMCAU WHERE PRONUM=" & aNUM(0) & " AND FALNUM=" & aNUM(1)
        GrvSQL cARQFEMEA, cSQL, 1, Array("FALNUM"), Array(nNUM), Array("NI")
        'cSQL = "select FALNUM from FEMEA WHERE PRONUM=" & aNUM(0) & " AND FALNUM=" & aNUM(1)
        cSQL = "select FALNUM from FEMEAPAD WHERE PRONUM=" & aNUM(0) & " AND FALNUM=" & aNUM(1)
        GrvSQL cARQFEMEA, cSQL, 1, Array("FALNUM"), Array(nNUM), Array("NI")
        cSQL = "select FALNUM from FEMPF WHERE PRONUM=" & aNUM(0) & " AND FALNUM=" & aNUM(1)
        GrvSQL cARQFEMEA, cSQL, 1, Array("FALNUM"), Array(nNUM), Array("NI")
        cSQL = "select FALNUM from FEMRPNT WHERE PRONUM=" & aNUM(0) & " AND FALNUM=" & aNUM(1)
        GrvSQL cARQFEMEA, cSQL, 1, Array("FALNUM"), Array(nNUM), Array("NI")
    Case 2
        cSQL = "select EFENUM from FEMEFE WHERE PRONUM=" & aNUM(0) & " AND FALNUM=" & aNUM(1) & " AND EFENUM=" & aNUM(2)
        GrvSQL cARQFEMEA, cSQL, 1, Array("EFENUM"), Array(nNUM), Array("NI")
        cSQL = "select EFENUM from FEMCAU WHERE PRONUM=" & aNUM(0) & " AND FALNUM=" & aNUM(1) & " AND EFENUM=" & aNUM(2)
        GrvSQL cARQFEMEA, cSQL, 1, Array("EFENUM"), Array(nNUM), Array("NI")
        'cSQL = "select EFENUM from FEMEA WHERE PRONUM=" & aNUM(0) & " AND FALNUM=" & aNUM(1) & " AND EFENUM=" & aNUM(2)
        cSQL = "select EFENUM from FEMEAPAD WHERE PRONUM=" & aNUM(0) & " AND FALNUM=" & aNUM(1) & " AND EFENUM=" & aNUM(2)
        GrvSQL cARQFEMEA, cSQL, 1, Array("EFENUM"), Array(nNUM), Array("NI")
        cSQL = "select EFENUM from FEMRPNT WHERE PRONUM=" & aNUM(0) & " AND FALNUM=" & aNUM(1) & " AND EFENUM=" & aNUM(2)
        GrvSQL cARQFEMEA, cSQL, 1, Array("EFENUM"), Array(nNUM), Array("NI")
    Case 3
        cSQL = "select CAUNUM from FEMCAU WHERE PRONUM=" & aNUM(0) & " AND FALNUM=" & aNUM(1) & " AND EFENUM=" & aNUM(2) & " AND CAUNUM=" & aNUM(3)
        GrvSQL cARQFEMEA, cSQL, 1, Array("CAUNUM"), Array(nNUM), Array("NI")
        'cSQL = "select CAUNUM from FEMEA WHERE PRONUM=" & aNUM(0) & " AND FALNUM=" & aNUM(1) & " AND EFENUM=" & aNUM(2) & " AND CAUNUM=" & aNUM(3)
        cSQL = "select CAUNUM from FEMEAPAD WHERE PRONUM=" & aNUM(0) & " AND FALNUM=" & aNUM(1) & " AND EFENUM=" & aNUM(2) & " AND CAUNUM=" & aNUM(3)
        GrvSQL cARQFEMEA, cSQL, 1, Array("CAUNUM"), Array(nNUM), Array("NI")
        cSQL = "select CAUNUM from FEMRPNT WHERE PRONUM=" & aNUM(0) & " AND FALNUM=" & aNUM(1) & " AND EFENUM=" & aNUM(2) & " AND CAUNUM=" & aNUM(3)
        GrvSQL cARQFEMEA, cSQL, 1, Array("CAUNUM"), Array(nNUM), Array("NI")
    End Select
    fillall

End Sub

Private Sub CmdREvi_Click()

    If Not checkval(3) Then
        Exit Sub
    End If
    FILGRIDrevCAU

End Sub

Private Sub CmdSev_Click()

    Dim nNUM As Integer
    Dim cSQL As String

    If Not checkval(2) Then
        Exit Sub
    End If
    nNUM = FixInt(Val(Busca("Qtde Severidade", "Trocando", CStr(nNUM), 8)))
    If nNUM = 0 Then
        Alert ("Qtde em Branco")
        Exit Sub
    End If
    If Not MDG("Trocar Severidade") Then
        Exit Sub
    End If
    If Not MDG("Isto Altera Causas e FEMEA Continuar") Then
        Exit Sub
    End If
    cSQL = "select INDSEV,INDOCO,INDDET,INDRIS from FEMCAU WHERE PRONUM=" & aNUM(0) & " AND FALNUM=" & aNUM(1) & " AND EFENUM=" & aNUM(2)
    TROCASEV cSQL, nNUM
    'cSQL = "select INDSEV,INDOCO,INDDET,INDRIS from FEMEA WHERE PRONUM=" & aNUM(0) & " AND FALNUM=" & aNUM(1) & " AND EFENUM=" & aNUM(2)
    cSQL = "select INDSEV,INDOCO,INDDET,INDRIS from FEMEAPAD WHERE PRONUM=" & aNUM(0) & " AND FALNUM=" & aNUM(1) & " AND EFENUM=" & aNUM(2)
    TROCASEV cSQL, nNUM

End Sub

Private Function TROCASEV(ByVal cSQL, ByVal nSEV)
    Dim DB As New ADODB.Connection
    Dim RSTAB As New ADODB.Recordset
    On Error Resume Next
    DB.ConnectionTimeout = 120
    DB.Open GeracArq(cARQFEMEA)
    RSTAB.Open cSQL, DB, adOpenDynamic, adLockOptimistic
    Do While Not RSTAB.EOF
        RSTAB("INDSEV") = nSEV
        RSTAB("INDRIS") = nSEV * FixInt(RSTAB("INDOCO")) * FixInt(RSTAB("INDDET"))
        RSTAB.Update
        RSTAB.MoveNext
    Loop
    RSTAB.Close
    DB.Close
    Set RSTAB = Nothing
    Set DB = Nothing
End Function

Private Sub CMDSIG_Click(Index As Integer)

    lRETU = False
    eRETU01 = ""
    eRETU02 = ""
    If Index = 0 Then
        escFLX.Show vbModal, Me
    Else
        ePASS01 = "isoqsymbol"
        ePASS02 = True
        frmCharacters.Show vbModal, Me
    End If
    If lRETU Then
        Me.TXTFIELDS(24) = eRETU01
    End If
    
  

End Sub

Private Sub Command1_Click()
    Dim nNUM
    Dim cDIZ As String
    gridcau.Col = 3
    nNUM = FixInt(gridcau)
    aNUM(3) = nNUM
    gridcau.Col = 4
    cDIZ = FixStr(gridcau)
    aDIZ(3) = cDIZ
    dizcau.Caption = CStr(nNUM) & " - " & cDIZ
    LocalizaGrid gridcau, aNUM(3), 4, False, 0
    ePASS01 = "select * from FEMCAU WHERE PRONUM=" & aNUM(0) & " AND FALNUM=" & aNUM(1) & " AND EFENUM=" & aNUM(2) & " AND CAUNUM=" & aNUM(3)
    ePASS02 = cARQFEMEA
    Load frmFEMEI
    frmFEMEI.cmdpegpro(0).Enabled = False
    frmFEMEI.cmdpegpro(1).Enabled = False
    frmFEMEI.cmdpegpro(2).Enabled = False
    frmFEMEI.cmdpegpro(3).Enabled = False
    frmFEMEI.TXTFIELDS(2).Enabled = False
    frmFEMEI.TXTFIELDS(3).Enabled = False
    frmFEMEI.TXTFIELDS(4).Enabled = False
    frmFEMEI.TXTFIELDS(5).Enabled = False
    
    frmFEMEI.TXTFIELDS(7).Enabled = False
    frmFEMEI.TXTFIELDS(8).Enabled = False
    frmFEMEI.TXTFIELDS(9).Enabled = False
    frmFEMEI.TXTFIELDS(10).Enabled = False
    
    frmFEMEI.Show vbModal, Me
    ALTERAFEMEA
End Sub

Private Sub ALTERAFEMEA()
    Dim cSQL As String
    Dim nROWS As Integer
    Dim x As Integer
    Dim nPFTMP As Long

    Dim DB As ADODB.Connection
    Dim RSTAB As ADODB.Recordset
    Set DB = New ADODB.Connection
    Set RSTAB = New ADODB.Recordset


    If MDG("Fazer Alteracoes No Femea") Then
        gerapfs
        DB.ConnectionTimeout = 120
        DB.Open GeracArq(cARQFEMEA)
        ''marca as que sao manuais
        'cSQL = "select PF,MUDPAD from FEMEA WHERE PRONUM=" & aNUM(0) & " AND FALNUM=" & aNUM(1) & " AND EFENUM=" & aNUM(2) & " AND CAUNUM=" & aNUM(3) & " AND ALTMAN"
        cSQL = "select PF,MUDPAD from FEMEAPAD WHERE PRONUM=" & aNUM(0) & " AND FALNUM=" & aNUM(1) & " AND EFENUM=" & aNUM(2) & " AND CAUNUM=" & aNUM(3) & " AND ALTMAN"
        RSTAB.Open cSQL, DB, adOpenDynamic, adLockOptimistic
        If Not RSTAB.EOF Then
            While Not RSTAB.EOF
                DIZAPU = RSTAB("pf")
                DIZAPU.Refresh
                RSTAB("MUDPAD") = True
                RSTAB.Update
                RSTAB.MoveNext
            Wend
        End If
        RSTAB.Close
        DB.Close
        'GrvSQLado cARQFEMEA, cSQL, 1, Array("MUDPAD"), Array(True), Array("") um a um gerava lentidao
        ''marca os pf
        nROWS = GridPadPF.Rows - 1
        Barra.Value = 0
        nBARPOS = 0
        nROWREC = nROWS
        GridPadPF.Col = 0
        DB.ConnectionTimeout = 120
        DB.Open GeracArq(cARQPF)
        For x = 0 To nROWS
            GridPadPF.Row = x
            nPFTMP = FixInt(GridPadPF)
            DIZAPU = nPFTMP & "- " & x & "/" & nROWS
            DIZAPU.Refresh
            BARPOS
            cSQL = "select FEMMUDPAD from PF WHERE PF=" & nPFTMP
            If nPFTMP > 0 Then
                RSTAB.Open cSQL, DB, adOpenDynamic, adLockOptimistic
                If Not RSTAB.EOF Then
                    RSTAB("FEMMUDPAD") = True
                    RSTAB.Update
                End If
                '   GrvSQL cARQPF, cSQL, 1, Array("FEMMUDPAD"), Array(True), Array("") 'muito lento gravando um a um
                RSTAB.Close
            End If
        Next
        DB.Close
        Set DB = Nothing
        Set RSTAB = Nothing
        Alert ("Indices Transferidos")
    End If

End Sub

Private Sub FILGRIdfemprecad()
    Dim cSQL As String
    DizerBarra "Carretando Femea pre Cadastro"
    cSQL = "SELECT SEGGRA,TITULO FROM FEMPRE ORDER BY TITULO"
    MontaGridUltra Gridfemprecad, 2, Array(800, 4000), Array("Código", "Descrição"), Array("SEGGRA", "L$TITULO"), cARQFEMEA, cSQL
    DizerBarra ""
End Sub

Private Sub Command2_Click()
    LocalizaGri1 gridpf
End Sub

Private Sub Command3_Click()
    LocalizaGri1 Gridfemprecad
End Sub

Private Sub Command4_Click()
    Dim nNUM As Integer
    Dim cDIZ As String
    gridcau.Col = 3

    nNUM = FixInt(gridcau)
    If nNUM = 0 Then
        Exit Sub
    End If
    aNUM(3) = nNUM
    gridcau.Col = 4
    cDIZ = FixStr(gridcau)
    aDIZ(3) = cDIZ
    fillall (4)
    LocalizaGrid gridcau, aNUM(3), 4, False, 0
    CmdREvi_Click

End Sub

Private Sub Command5_Click()
    If Not checkval(3) Then
        Exit Sub
    End If
    gerarevisao aNUM(3)
    ALTERAFEMEA

End Sub

Private Sub gerarevisao(nCAUSA)

    Dim nREVI As Long
    Dim cSQL As String
    Dim aVAL As Variant
    Dim aCAM As Variant
    Dim aFOR As Variant
    Dim aPAD As Variant

    nCAUSA = FixInt(nCAUSA)
    
    ''Numero da Proxima revisao
    cSQL = "FEMrevi WHERE  TIPOAPU='P' AND PRONUM=" & FixInt(aNUM(0)) & " AND FALNUM=" & FixInt(aNUM(1)) & " AND EFEnum=" & FixInt(aNUM(2)) & " AND caunum=" & nCAUSA
    nREVI = FixInt(PegMAXSQL(cARQFEMEA, cSQL, "SEGGRA", 0))
    nREVI = nREVI + 1
    
    
    aCAM = Array("PF", "ITEM", "PROCESSO", "FALTIP", "FALEFE", _
                 "FALCAU", "CRTATU", "INDOCO", "INDSEV", "INDDET", _
                 "INDRIS", "ACAREC", "RESCOD", "RESNOM", "RESCOD2", _
                 "ACATOM", "RESCOD3", "RINDOCO", "RINDSER", "RINDDET", _
                 "RINDRIS", "RESNOM2", "RESNOM3", "SITUACAO", "SIGI", _
                 "CARAPREV", "RESDAT", "RESDAT2", "RESDAT3", "ACADAT", _
                 "ACAO", "EXCRPN", "ALTMAN", "MUDPAD", "TITULO", _
                 "PRONUM", "FALNUM", "EFENUM", "CAUNUM", "SEGGRA", "TIPOAPU")
    aFOR = Array("NI", "NI", "C", "C", "C", _
                 "C", "C", "NI", "NI", "NI", _
                 "NI", "C", "NI", "C", "NI", _
                 "C", "NI", "NI", "NI", "NI", _
                 "NI", "C", "C", "C", "C", "C", _
                 "DN", "DN", "DN", "DN", "BN", _
                 "BN", "BN", "BN", "C", _
                 "NI", "NI", "NI", "NI", "NI", "C")
    aPAD = Array(0, 0, "", "", "", _
                 "", "", 0, 0, 0, _
                 0, "", 0, "", 0, _
                 "", 0, 0, 0, 0, 0, _
                 "", "", "", "", "", _
                 "", "", "", "", False, _
                 False, False, False, "", _
                 0, 0, 0, 0, 0, "P")
                 
    
    cSQL = "select * from FEMCAU WHERE PRONUM=" & aNUM(0) & " AND FALNUM=" & aNUM(1) & " AND EFENUM=" & aNUM(2) & " AND CAUNUM=" & nCAUSA
    aVAL = PegSQL(cARQFEMEA, cSQL, 41, aCAM, aFOR, aPAD)
    aVAL(39) = nREVI
    aVAL(40) = "P"                               'Padrao
    
    If aVAL(17) = 0 Or aVAL(18) = 0 Or aVAL(19) = 0 Or aVAL(20) = 0 Then
        Alert ("Novos Indices Nao Preenchidos")
        Exit Sub
    End If
    
    
    
    cSQL = "select * from FEMREVI WHERE  TIPOAPU='P' AND PRONUM=" & aNUM(0) & " AND FALNUM=" & aNUM(1) & " AND EFENUM=" & aNUM(2) & " AND CAUNUM=" & nCAUSA & " AND SEGGRA=" & nREVI
    IncluiSQL cARQFEMEA, cSQL, 41, aCAM, aVAL, True, False
    
    cSQL = "select * from FEMCAU WHERE PRONUM=" & aNUM(0) & " AND FALNUM=" & aNUM(1) & " AND EFENUM=" & aNUM(2) & " AND CAUNUM=" & nCAUSA
    
    GrvSQL cARQFEMEA, cSQL, 17, Array("ACAREC", "RESCOD", "RESNOM", "RESCOD2", "ACATOM", _
                                      "RESCOD3", "RINDOCO", "RINDSER", "RINDDET", "RINDRIS", _
                                      "RESNOM2", "RESNOM3", "INDOCO", "INDSEV", "INDDET", _
                                      "INDRIS", "ACAO"), _
                                      Array("", 0, "", 0, "", 0, 0, 0, 0, 0, "", "", aVAL(17), aVAL(18), aVAL(19), aVAL(20), False), _
                                      Array("C", "NI", "C", "NI", "C", "NI", "NI", "NI", "NI", "NI", "C", "C", "NI", "NI", "NI", "NI", "")
    FILGRIDrevCAU

End Sub

Private Sub Command6_Click()

    Dim nROWS
    Dim x
    Dim nNUM As Integer

    If Not MDG("Revisar Todas as Causas") Then
        Exit Sub
    End If
    nROWS = gridcau.Rows
    For x = 1 To nROWS - 1
        gridcau.Row = x
        gridcau.Col = 3
        nNUM = FixInt(gridcau)
        If nNUM > 0 Then
            gerarevisao nNUM
        End If
    Next

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    TeclaEnter KeyCode
End Sub

Private Sub FilRelat()
    Dim cSQL As String
    If Len(cSUBWHERE) = 0 Then
        cSQL = "SELECT SEQ,MES,ANO FROM FEMEARPN ORDER BY " & cORDEM
    Else
        cSQL = "SELECT SEQ,MES,ANO FROM FEMEARPN WHERE " & cSUBWHERE & " ORDER BY " & cORDEM
    End If
    DizerBarra "Carregando Sequencias"
    MontaGridUltra Grid, 3, Array(400, 400, 600), Array("Seq ", "Mes", "Ano"), _
        Array("SEQ", "MES", "ANO"), cARQUSO, cSQL
    DizerBarra ""
End Sub

Private Sub Form_Load()
    Center Me
    aNUM = Array(0, 0, 0, 0)
    aDIZ = Array("", "", "", "")
    cARQFEMEA = PegPath("PATH", "FEMEA")
    cARQPF = PegPath("PATH", "PF")
    fillall
    FILGRIdfemprecad
    filgridpadpf
    filgridpf

    
    cARQUSO = PegPath("PATH", "FEMEA")
    aORDEM = Array("SEQ", "MES,ANO", "ANO,MES")
    aORDES = Array("Sequencia", "Mes/ANO", "Ano/Mes")
    cORDEM = "SEQ"
    cSUBWHERE = ""
    xmontatoolbar Me.Toolbar1, "escFemeaApu", True
    FilRelat
    
    TXTFIELDS(24).Font = "isoqsymbol"
    
    
End Sub

Private Sub FILGRIdPRO()

    Dim cSQL As String
    cSQL = "SELECT pronum,processo FROM FEMPRO ORDER BY PRONUM"
    DizerBarra "Carregando Processos"
    MontaGridUltra GridPRO, 2, Array(600, 6000), Array("Nº", "Descricao"), Array("PRONUM", "PROCESSO"), cARQFEMEA, cSQL
    DizerBarra ""

End Sub

Private Sub FILGRIdFAL()

    Dim cSQL As String
    cSQL = "SELECT PRONUM, FALNUM, FALTIP FROM FEMFAL WHERE PRONUM=" & FixInt(aNUM(0)) & " order by FALNUM"

    DizerBarra "Carregando Falhas"
    MontaGridUltra GRIDFAL, 3, Array(600, 600, 6000), Array("NºPro", "Nº", "Descricao"), Array("PRONUM", "FALNUM", "FALTIP"), cARQFEMEA, cSQL
    DizerBarra ""

End Sub

Private Sub FILGRIdEFE()

    Dim cSQL As String
    cSQL = "SELECT PRONUM, FALNUM, EFENUM, FALEFE FROM FEMEFE WHERE PRONUM=" & FixInt(aNUM(0)) & " AND FALNUM=" & FixInt(aNUM(1)) & " ORDER BY EFENUM"
   
    DizerBarra "Carregando Efeitos"
    MontaGridUltra gridefe, 4, Array(600, 600, 600, 6000), Array("NºPro", "NºFal", "NºEfe", "Descricao"), Array("PRONUM", "FALNUM", "EFENUM", "FALEFE"), cARQFEMEA, cSQL
    DizerBarra ""

End Sub

Private Sub FILGRIdCAU()

    Dim cSQL As String
    cSQL = "SELECT PRONUM,FALNUM,EFENUM,cAUNUM,FALCAU FROM FEMCAU WHERE PRONUM=" & FixInt(aNUM(0)) & " AND FALNUM=" & FixInt(aNUM(1)) & " AND EFEnum=" & FixInt(aNUM(2)) & " ORDER BY CAUNUM"

    DizerBarra "Carregando Causas"
    MontaGridUltra gridcau, 5, Array(600, 600, 600, 600, 5000), Array("NºPro", "NºFal", "NºEfe", "NºCau", "Descricao"), Array("PRONUM", "FALNUM", "EFENUM", "CAUNUM", "FALCAU"), cARQFEMEA, cSQL
    DizerBarra ""

End Sub

Private Sub FILGRIDrevCAU()

    Dim cSQL As String
    cSQL = "SELECT PRONUM,FALNUM,EFENUM,CAUNUM,SEGGRA,FALCAU FROM FEMrevi WHERE  TIPOAPU='P' AND PRONUM=" & FixInt(aNUM(0)) & " AND FALNUM=" & FixInt(aNUM(1)) & " AND EFEnum=" & FixInt(aNUM(2)) & " AND caunum=" & FixInt(aNUM(3)) & " order by SEGGRA"

    DizerBarra "Carregando Revisao Causas"
    MontaGridUltra gridrevcau, 6, Array(600, 600, 600, 600, 600, 5000), Array("NºPro", "NºFal", "NºEfe", "NºCau", "Rev", "Descricao"), Array("PRONUM", "FALNUM", "EFENUM", "CAUNUM", "SEGGRA", "FALCAU"), cARQFEMEA, cSQL
    DizerBarra ""

End Sub

Private Sub Retornar_Click(Index As Integer)

    Unload Me

End Sub

Private Sub filgridpadpf()
    Dim cSQL As String
    If aNUM(0) = 0 Or aNUM(1) = 0 Then
        Exit Sub
    End If
    cSQL = "SELECT PF,PRONUM,FALNUM FROM FEMPF WHERE PRONUM=" & FixInt(aNUM(0)) & " AND FALNUM=" & FixInt(aNUM(1)) & " ORDER BY PF"
    DizerBarra "carregando femea pf"
    MontaGridUltra GridPadPF, 3, Array(400, 400, 400), Array("PF", "NºPro", "NºFAL"), Array("PF", "PRONUM", "FALNUM"), cARQFEMEA, cSQL
    DizerBarra ""
End Sub

Private Sub filgridpf()
    Dim cSQL As String
    cSQL = "SELECT PF,codigo,DEscr FROM PF ORDER BY PF"
    DizerBarra "carregando processos de fabricacao"
    MontaGridUltra gridpf, 3, Array(400, 1600, 4000), Array("PF", "Produto", "Descricao"), Array("PF", "L$CODIGO", "DESCR"), cARQPF, cSQL
    DizerBarra ""
End Sub

Private Function checkval(Index As Integer)

    checkval = True
    If aNUM(0) = 0 Then
        Alert ("Processo Nao Escolhido")
        checkval = False
    End If
    If aNUM(1) = 0 And Index > 0 Then
        Alert ("Falha Nao Escolhido")
        checkval = False
    End If
    If aNUM(2) = 0 And Index > 1 Then
        Alert ("Efeito Nao Escolhido")
        checkval = False
    End If
    If aNUM(3) = 0 And Index > 2 Then
        Alert ("Causa Nao Escolhido")
        checkval = False
    End If

End Function

Private Sub BARPOS()

    Dim nPOSATU As Long
    nPOSATU = 0

    nBARPOS = nBARPOS + 1
    If nROWREC > 0 Then
        nPOSATU = Int(nBARPOS / nROWREC * 100)
    End If
    If nPOSATU > 0 And nPOSATU < 101 Then
        Debug.Print nPOSATU
        Barra.Value = nPOSATU
    End If
    DIZAPU.Refresh

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim sButton As String
    sButton = Button
    sButton = Left(UCase(Replace(sButton, "&", "")), 3)
    
    If Not AcessaBtnOld("escFemeaApu", Button.Index) Then
        Exit Sub
    End If
    
    GravaLog 0, Button.Index, sButton, "escFemeaApu"
    
    Select Case sButton
    Case "ORD"
        ePASS01 = aORDES
        escOrdem.Show vbModal, Me
        If lRETU Then
            cORDEM = aORDEM(eRETU01)
            FilRelat
        End If
    Case "FIL"
        cSUBWHERE = ""
        If MDG("Usar Filtro Avancado") Then
            aARQUIVOS = Array(cARQUSO)
            ''Posicao 12 Nome da Tabela
            ''Posicao 13 Nome da Tabela
            aRELCFG = Array("", "", "", 0, False, _
                            False, "", "", "", "", "", _
                            False, "FEMEARPN", "FEMEARPN", "", "")
            FrmFiltro.Show vbModal, Me
            If lRETU Then
                cSUBWHERE = Replace(Replace(eRETU01, "{", ""), "}", "")
            End If
        Else
            ePASS01 = aORDES
            frmLocalizaa.Show vbModal, Me
            If lRETU Then
                cSUBWHERE = MontaFiltro(aORDEM, Array("=", "=", "="), eRETU01, eRETU02)
            End If
        End If
        FilRelat
    Case "IMP"
        cTIPO = "R"
        zgrp = "FEME"
        escRPT.Show vbModal, Me
    Case "NOV"
        NovoPF_Click
    Case "EDI"
        EditPf_Click
    Case "EXC"
        ApagaNF_Click
    Case "LOC"
        ePASS01 = aORDES
        frmLocalizaa.Show vbModal, Me
        If lRETU Then
            LocalizaGrid Grid, eRETU01, eRETU02, , 1
        End If
            
    Case "SAI"
        Unload Me
    End Select
End Sub

Private Sub ApagaNF_Click()
    Dim nNUMERO As Long
    Dim sSQL As String
    Dim nMES As Long
    Dim nANO As Long
    Grid.Col = 0
    nNUMERO = Grid
    Grid.Col = 1
    nMES = Grid
    Grid.Col = 2
    nANO = Grid
    sSQL = "select * from FEMEARPN WHERE SEQ=" & nNUMERO
    If ApagaSQLP(cARQFEMEA, sSQL) Then
        sSQL = "select * from FEMEARPNI WHERE SEQ=" & nNUMERO
        ApagaSQL cARQFEMEA, sSQL
        sSQL = "select * from FEMRPNO WHERE SEQ=" & nNUMERO
        ApagaSQL cARQFEMEA, sSQL
        sSQL = "select * from FEMRPNT WHERE SEQ=" & nNUMERO
        ApagaSQL cARQFEMEA, sSQL
        sSQL = "select * from FEMEARPNG WHERE ANO=" & nANO & " AND MES=" & nMES
        FilRelat
    End If
End Sub

Private Sub NovoPF_Click()
    Dim nNUMERO As Long
    nNUMERO = FixInt(PegMAXSQL(cARQUSO, "FEMEARPN", "SEQ", 0)) + 1
    IncluiSQL cARQUSO, "SELECT * FROM FEMEARPN WHERE SEQ=" & nNUMERO, 1, Array("SEQ"), Array(nNUMERO)
    FilRelat
End Sub

Private Sub EditPf_Click()
    Dim nNUMERO As Long
    Grid.Col = 0
    nNUMERO = Grid
    ePASS01 = "select * from FEMEARPN WHERE SEQ=" & nNUMERO
    frmFemeaApu.Show vbModal
    FilRelat
End Sub

Private Sub Grid_KeyPress(KeyAscii As Integer)
    If KeyAscii > 31 And KeyAscii < 123 Then
        LocalizaGrid Grid, Chr(KeyAscii), 1, False
    End If
End Sub

'Private Sub Grid_SelChange()
  '  With Grid
'        If .Rows > 2 Then
 '           .Col = .Cols - 1
  '          .ColSel = 0
   '         .TopRow = .Row
    '    End If
   ' End With
'End Sub


