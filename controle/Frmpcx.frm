VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{7020C36F-09FC-41FE-B822-CDE6FBB321EB}#1.0#0"; "vbccr17.ocx"
Begin VB.Form frmPCX 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Plano de Controle"
   ClientHeight    =   9045
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   12750
   Icon            =   "Frmpcx.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9045
   ScaleWidth      =   12750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton NewMat 
      Caption         =   "F."
      Height          =   255
      Index           =   4
      Left            =   9120
      TabIndex        =   344
      Top             =   1560
      Width           =   255
   End
   Begin VB.CommandButton NewMat 
      Caption         =   "L."
      Height          =   255
      Index           =   23
      Left            =   8760
      TabIndex        =   268
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton NewMat 
      Caption         =   "L."
      Height          =   255
      Index           =   22
      Left            =   8760
      TabIndex        =   267
      Top             =   1860
      Width           =   255
   End
   Begin VB.CommandButton NewMat 
      Caption         =   "L."
      Height          =   255
      Index           =   21
      Left            =   8760
      TabIndex        =   266
      Top             =   1560
      Width           =   255
   End
   Begin VB.CommandButton NewMat 
      Caption         =   "R."
      Height          =   255
      Index           =   11
      Left            =   8400
      TabIndex        =   265
      Top             =   1560
      Width           =   315
   End
   Begin VB.CommandButton NewMat 
      Caption         =   "R."
      Height          =   255
      Index           =   12
      Left            =   8400
      TabIndex        =   264
      Top             =   1860
      Width           =   315
   End
   Begin VB.CommandButton NewMat 
      Caption         =   "R."
      Height          =   255
      Index           =   13
      Left            =   8400
      TabIndex        =   263
      Top             =   2160
      Width           =   315
   End
   Begin VB.TextBox TXTPF 
      BackColor       =   &H00C0FFFF&
      DataField       =   "NOMMU013"
      DataSource      =   "datPrimaryRS"
      Enabled         =   0   'False
      Height          =   285
      Index           =   11
      Left            =   1260
      Locked          =   -1  'True
      TabIndex        =   262
      Top             =   2160
      Width           =   7095
   End
   Begin VB.TextBox TXTPF 
      BackColor       =   &H00C0FFFF&
      DataField       =   "NOMMU012"
      DataSource      =   "datPrimaryRS"
      Enabled         =   0   'False
      Height          =   285
      Index           =   10
      Left            =   1260
      Locked          =   -1  'True
      TabIndex        =   261
      Top             =   1860
      Width           =   7095
   End
   Begin VB.TextBox TXTPF 
      BackColor       =   &H00C0FFFF&
      DataField       =   "NOMMU011"
      DataSource      =   "datPrimaryRS"
      Enabled         =   0   'False
      Height          =   285
      Index           =   9
      Left            =   1260
      Locked          =   -1  'True
      TabIndex        =   260
      Top             =   1560
      Width           =   7095
   End
   Begin VB.TextBox TXTPF 
      BackColor       =   &H00C0FFFF&
      DataField       =   "CODMU013"
      DataSource      =   "datPrimaryRS"
      Enabled         =   0   'False
      Height          =   285
      Index           =   8
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   259
      TabStop         =   0   'False
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox TXTPF 
      BackColor       =   &H00C0FFFF&
      DataField       =   "CODMU012"
      DataSource      =   "datPrimaryRS"
      Enabled         =   0   'False
      Height          =   285
      Index           =   7
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   258
      TabStop         =   0   'False
      Top             =   1860
      Width           =   1095
   End
   Begin VB.TextBox TXTPF 
      BackColor       =   &H00C0FFFF&
      DataField       =   "CODMU011"
      DataSource      =   "datPrimaryRS"
      Enabled         =   0   'False
      Height          =   285
      Index           =   6
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   257
      TabStop         =   0   'False
      Top             =   1560
      Width           =   1095
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6255
      Index           =   24
      Left            =   120
      TabIndex        =   15
      Top             =   2640
      Width           =   12195
      _ExtentX        =   21511
      _ExtentY        =   11033
      _Version        =   393216
      Tabs            =   23
      Tab             =   20
      TabsPerRow      =   8
      TabHeight       =   529
      TabCaption(0)   =   "Recebimento"
      TabPicture(0)   =   "Frmpcx.frx":058A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblLabels(9)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblLabels(10)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "gridrevi(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "gridrev(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "GRIDMAT(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "EditSeq(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "CmdNovaREV(1)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "escTIPPC(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "TXTpc(5)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "TXTpc(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "CmdMotRev(1)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "ComMotInc(1)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "ComMotExc(1)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "ComMotAlt(1)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "ComverRev(1)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "CmdTotMot(1)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "TXTpc(15)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Cmddistro(1)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).ControlCount=   18
      TabCaption(1)   =   "Laboratorio"
      TabPicture(1)   =   "Frmpcx.frx":05A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblLabels(11)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblLabels(12)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "gridrevi(2)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "gridrev(2)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "gridlab(0)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "EditSeq(2)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "CmdNovaREV(2)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "TXTpc(6)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "TXTpc(1)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "CmdMotRev(2)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "ComMotInc(2)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "ComMotExc(2)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "ComMotAlt(2)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "ComverRev(2)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "escTIPPC(2)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "CmdTotMot(2)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "TXTpc(16)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Cmddistro(2)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).ControlCount=   18
      TabCaption(2)   =   "Producao"
      TabPicture(2)   =   "Frmpcx.frx":05C2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblLabels(13)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "lblLabels(14)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "gridrevi(3)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "gridrev(3)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "GridSeq(0)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "EditSeq(3)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "CmdNovaREV(3)"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "escTIPPC(3)"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "TXTpc(7)"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "TXTpc(2)"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "CmdMotRev(3)"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "ComMotInc(3)"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "ComMotExc(3)"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "ComMotAlt(3)"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "ComverRev(3)"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "CmdTotMot(3)"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "TXTpc(17)"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "Cmddistro(3)"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).ControlCount=   18
      TabCaption(3)   =   "Final"
      TabPicture(3)   =   "Frmpcx.frx":05DE
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lblLabels(15)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "lblLabels(16)"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "lblLabels(34)"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "gridrevi(4)"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "gridrev(4)"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "GRIDFIM(0)"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "EditSeq(4)"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "CmdNovaREV(4)"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "escTIPPC(4)"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "TXTpc(8)"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "TXTpc(3)"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "CmdMotRev(4)"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).Control(12)=   "ComMotInc(4)"
      Tab(3).Control(12).Enabled=   0   'False
      Tab(3).Control(13)=   "ComMotExc(4)"
      Tab(3).Control(13).Enabled=   0   'False
      Tab(3).Control(14)=   "ComMotAlt(4)"
      Tab(3).Control(14).Enabled=   0   'False
      Tab(3).Control(15)=   "selmar(0)"
      Tab(3).Control(15).Enabled=   0   'False
      Tab(3).Control(16)=   "selmar(1)"
      Tab(3).Control(16).Enabled=   0   'False
      Tab(3).Control(17)=   "Check4"
      Tab(3).Control(17).Enabled=   0   'False
      Tab(3).Control(18)=   "TXTpc(22)"
      Tab(3).Control(18).Enabled=   0   'False
      Tab(3).Control(19)=   "ESCprofin(0)"
      Tab(3).Control(19).Enabled=   0   'False
      Tab(3).Control(20)=   "selmar(2)"
      Tab(3).Control(20).Enabled=   0   'False
      Tab(3).Control(21)=   "TXTpc(25)"
      Tab(3).Control(21).Enabled=   0   'False
      Tab(3).Control(22)=   "ComverRev(4)"
      Tab(3).Control(22).Enabled=   0   'False
      Tab(3).Control(23)=   "ESCprofin(1)"
      Tab(3).Control(23).Enabled=   0   'False
      Tab(3).Control(24)=   "CmdTotMot(4)"
      Tab(3).Control(24).Enabled=   0   'False
      Tab(3).Control(25)=   "TXTpc(18)"
      Tab(3).Control(25).Enabled=   0   'False
      Tab(3).Control(26)=   "Cmddistro(4)"
      Tab(3).Control(26).Enabled=   0   'False
      Tab(3).ControlCount=   27
      TabCaption(4)   =   "Elab./Rev."
      TabPicture(4)   =   "Frmpcx.frx":05FA
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "lblLabels(0)"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "lblLabels(8)"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "lblLabels(6)"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "gridrev(0)"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "gridrevi(0)"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "TXTpc(13)"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "escpfela(0)"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).Control(7)=   "TXTpc(10)"
      Tab(4).Control(7).Enabled=   0   'False
      Tab(4).Control(8)=   "CmdMotRev(0)"
      Tab(4).Control(8).Enabled=   0   'False
      Tab(4).Control(9)=   "ComMotAlt(0)"
      Tab(4).Control(9).Enabled=   0   'False
      Tab(4).Control(10)=   "ComMotExc(0)"
      Tab(4).Control(10).Enabled=   0   'False
      Tab(4).Control(11)=   "ComMotInc(0)"
      Tab(4).Control(11).Enabled=   0   'False
      Tab(4).Control(12)=   "Check1"
      Tab(4).Control(12).Enabled=   0   'False
      Tab(4).Control(13)=   "Check2"
      Tab(4).Control(13).Enabled=   0   'False
      Tab(4).Control(14)=   "TXTpc(11)"
      Tab(4).Control(14).Enabled=   0   'False
      Tab(4).Control(15)=   "TXTpc(14)"
      Tab(4).Control(15).Enabled=   0   'False
      Tab(4).Control(16)=   "Command1(0)"
      Tab(4).Control(16).Enabled=   0   'False
      Tab(4).Control(17)=   "Check3"
      Tab(4).Control(17).Enabled=   0   'False
      Tab(4).Control(18)=   "ComverRev(0)"
      Tab(4).Control(18).Enabled=   0   'False
      Tab(4).Control(19)=   "CmdTotMot(0)"
      Tab(4).Control(19).Enabled=   0   'False
      Tab(4).Control(20)=   "escidfolha(0)"
      Tab(4).Control(20).Enabled=   0   'False
      Tab(4).Control(21)=   "TXTpc(12)"
      Tab(4).Control(21).Enabled=   0   'False
      Tab(4).Control(22)=   "ComMotInc(26)"
      Tab(4).Control(22).Enabled=   0   'False
      Tab(4).Control(23)=   "TXTpc(32)"
      Tab(4).Control(23).Enabled=   0   'False
      Tab(4).Control(24)=   "Cmddistro(0)"
      Tab(4).Control(24).Enabled=   0   'False
      Tab(4).Control(25)=   "escidfolha(10)"
      Tab(4).Control(25).Enabled=   0   'False
      Tab(4).Control(26)=   "PCDinipc"
      Tab(4).Control(26).Enabled=   0   'False
      Tab(4).ControlCount=   27
      TabCaption(5)   =   "Composição"
      TabPicture(5)   =   "Frmpcx.frx":0616
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "lblLabels(7)"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "lblLabels(17)"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "gridrevi(5)"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "gridrev(5)"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).Control(4)=   "GRIDCOM(0)"
      Tab(5).Control(4).Enabled=   0   'False
      Tab(5).Control(5)=   "EditSeq(5)"
      Tab(5).Control(5).Enabled=   0   'False
      Tab(5).Control(6)=   "CmdNovaREV(5)"
      Tab(5).Control(6).Enabled=   0   'False
      Tab(5).Control(7)=   "escTIPPC(5)"
      Tab(5).Control(7).Enabled=   0   'False
      Tab(5).Control(8)=   "TXTpc(9)"
      Tab(5).Control(8).Enabled=   0   'False
      Tab(5).Control(9)=   "TXTpc(4)"
      Tab(5).Control(9).Enabled=   0   'False
      Tab(5).Control(10)=   "CmdMotRev(5)"
      Tab(5).Control(10).Enabled=   0   'False
      Tab(5).Control(11)=   "ComMotInc(5)"
      Tab(5).Control(11).Enabled=   0   'False
      Tab(5).Control(12)=   "ComMotExc(5)"
      Tab(5).Control(12).Enabled=   0   'False
      Tab(5).Control(13)=   "ComMotAlt(5)"
      Tab(5).Control(13).Enabled=   0   'False
      Tab(5).Control(14)=   "ComverRev(5)"
      Tab(5).Control(14).Enabled=   0   'False
      Tab(5).Control(15)=   "CmdTotMot(5)"
      Tab(5).Control(15).Enabled=   0   'False
      Tab(5).Control(16)=   "TXTpc(19)"
      Tab(5).Control(16).Enabled=   0   'False
      Tab(5).Control(17)=   "Cmddistro(5)"
      Tab(5).Control(17).Enabled=   0   'False
      Tab(5).ControlCount=   18
      TabCaption(6)   =   "Rev.PF"
      TabPicture(6)   =   "Frmpcx.frx":0632
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "GridPFrev"
      Tab(6).Control(1)=   "gridrevi(6)"
      Tab(6).Control(2)=   "ComMotAlt(6)"
      Tab(6).Control(3)=   "ComverRev(6)"
      Tab(6).Control(4)=   "TXTPF(18)"
      Tab(6).Control(4).Enabled=   0   'False
      Tab(6).Control(5)=   "TXTPF(17)"
      Tab(6).Control(5).Enabled=   0   'False
      Tab(6).Control(6)=   "CmdMarcaPC(0)"
      Tab(6).Control(7)=   "CmdMarcaPC(1)"
      Tab(6).Control(8)=   "CmdMarcaPC(2)"
      Tab(6).Control(9)=   "CmdMarcaPC(3)"
      Tab(6).Control(10)=   "CmdMarcaPC(4)"
      Tab(6).Control(11)=   "Cmddistro(6)"
      Tab(6).ControlCount=   12
      TabCaption(7)   =   "Rev Femea"
      TabPicture(7)   =   "Frmpcx.frx":064E
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "gridrevi(7)"
      Tab(7).Control(1)=   "ComMotAlt(7)"
      Tab(7).Control(2)=   "ComverRev(7)"
      Tab(7).Control(3)=   "TXTPF(16)"
      Tab(7).Control(3).Enabled=   0   'False
      Tab(7).Control(4)=   "TXTPF(15)"
      Tab(7).Control(4).Enabled=   0   'False
      Tab(7).Control(5)=   "TXTPF(19)"
      Tab(7).Control(5).Enabled=   0   'False
      Tab(7).Control(6)=   "Cmddistro(7)"
      Tab(7).ControlCount=   7
      TabCaption(8)   =   "Rec.  Pre"
      TabPicture(8)   =   "Frmpcx.frx":066A
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Cmddistro(11)"
      Tab(8).Control(1)=   "TXTpre(15)"
      Tab(8).Control(1).Enabled=   0   'False
      Tab(8).Control(2)=   "CmdTotMot(11)"
      Tab(8).Control(3)=   "ComverRev(11)"
      Tab(8).Control(4)=   "EditSeq(11)"
      Tab(8).Control(5)=   "CmdNovaREV(11)"
      Tab(8).Control(6)=   "escTIPPC(11)"
      Tab(8).Control(7)=   "TXTpre(5)"
      Tab(8).Control(7).Enabled=   0   'False
      Tab(8).Control(8)=   "TXTpre(0)"
      Tab(8).Control(8).Enabled=   0   'False
      Tab(8).Control(9)=   "CmdMotRev(11)"
      Tab(8).Control(10)=   "ComMotInc(11)"
      Tab(8).Control(11)=   "ComMotExc(11)"
      Tab(8).Control(12)=   "ComMotAlt(11)"
      Tab(8).Control(13)=   "GRIDMAT(1)"
      Tab(8).Control(14)=   "gridrev(11)"
      Tab(8).Control(15)=   "gridrevi(11)"
      Tab(8).Control(16)=   "lblLabels(36)"
      Tab(8).Control(17)=   "lblLabels(25)"
      Tab(8).ControlCount=   18
      TabCaption(9)   =   "Lab Pre"
      TabPicture(9)   =   "Frmpcx.frx":0686
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "Cmddistro(12)"
      Tab(9).Control(1)=   "TXTpre(16)"
      Tab(9).Control(1).Enabled=   0   'False
      Tab(9).Control(2)=   "CmdTotMot(12)"
      Tab(9).Control(3)=   "ComverRev(12)"
      Tab(9).Control(4)=   "EditSeq(12)"
      Tab(9).Control(5)=   "CmdNovaREV(12)"
      Tab(9).Control(6)=   "escTIPPC(12)"
      Tab(9).Control(7)=   "TXTpre(6)"
      Tab(9).Control(7).Enabled=   0   'False
      Tab(9).Control(8)=   "TXTpre(1)"
      Tab(9).Control(8).Enabled=   0   'False
      Tab(9).Control(9)=   "CmdMotRev(12)"
      Tab(9).Control(10)=   "ComMotInc(12)"
      Tab(9).Control(11)=   "ComMotExc(12)"
      Tab(9).Control(12)=   "ComMotAlt(12)"
      Tab(9).Control(13)=   "gridlab(1)"
      Tab(9).Control(14)=   "gridrev(9)"
      Tab(9).Control(15)=   "gridrevi(9)"
      Tab(9).Control(16)=   "lblLabels(27)"
      Tab(9).Control(17)=   "lblLabels(26)"
      Tab(9).ControlCount=   18
      TabCaption(10)  =   "Prod Pre"
      TabPicture(10)  =   "Frmpcx.frx":06A2
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "Cmddistro(13)"
      Tab(10).Control(1)=   "TXTpre(17)"
      Tab(10).Control(1).Enabled=   0   'False
      Tab(10).Control(2)=   "CmdTotMot(13)"
      Tab(10).Control(3)=   "ComverRev(13)"
      Tab(10).Control(4)=   "EditSeq(13)"
      Tab(10).Control(5)=   "CmdNovaREV(13)"
      Tab(10).Control(6)=   "escTIPPC(13)"
      Tab(10).Control(7)=   "TXTpre(7)"
      Tab(10).Control(7).Enabled=   0   'False
      Tab(10).Control(8)=   "TXTpre(2)"
      Tab(10).Control(8).Enabled=   0   'False
      Tab(10).Control(9)=   "CmdMotRev(13)"
      Tab(10).Control(10)=   "ComMotInc(13)"
      Tab(10).Control(11)=   "ComMotExc(13)"
      Tab(10).Control(12)=   "ComMotAlt(13)"
      Tab(10).Control(13)=   "GridSeq(1)"
      Tab(10).Control(14)=   "gridrev(13)"
      Tab(10).Control(15)=   "gridrevi(13)"
      Tab(10).Control(16)=   "lblLabels(29)"
      Tab(10).Control(17)=   "lblLabels(28)"
      Tab(10).ControlCount=   18
      TabCaption(11)  =   "Final Pre"
      TabPicture(11)  =   "Frmpcx.frx":06BE
      Tab(11).ControlEnabled=   0   'False
      Tab(11).Control(0)=   "Cmddistro(14)"
      Tab(11).Control(1)=   "TXTpre(18)"
      Tab(11).Control(1).Enabled=   0   'False
      Tab(11).Control(2)=   "CmdTotMot(14)"
      Tab(11).Control(3)=   "ComverRev(14)"
      Tab(11).Control(4)=   "EditSeq(14)"
      Tab(11).Control(5)=   "CmdNovaREV(14)"
      Tab(11).Control(6)=   "escTIPPC(14)"
      Tab(11).Control(7)=   "TXTpre(8)"
      Tab(11).Control(7).Enabled=   0   'False
      Tab(11).Control(8)=   "TXTpre(3)"
      Tab(11).Control(8).Enabled=   0   'False
      Tab(11).Control(9)=   "CmdMotRev(14)"
      Tab(11).Control(10)=   "ComMotInc(14)"
      Tab(11).Control(11)=   "ComMotExc(14)"
      Tab(11).Control(12)=   "ComMotAlt(14)"
      Tab(11).Control(13)=   "GRIDFIM(1)"
      Tab(11).Control(14)=   "gridrev(14)"
      Tab(11).Control(15)=   "gridrevi(14)"
      Tab(11).Control(16)=   "lblLabels(31)"
      Tab(11).Control(17)=   "lblLabels(30)"
      Tab(11).ControlCount=   18
      TabCaption(12)  =   "Elab/Rev Pre"
      TabPicture(12)  =   "Frmpcx.frx":06DA
      Tab(12).ControlEnabled=   0   'False
      Tab(12).Control(0)=   "lblLabels(22)"
      Tab(12).Control(0).Enabled=   0   'False
      Tab(12).Control(1)=   "lblLabels(23)"
      Tab(12).Control(1).Enabled=   0   'False
      Tab(12).Control(2)=   "lblLabels(24)"
      Tab(12).Control(2).Enabled=   0   'False
      Tab(12).Control(3)=   "Command1(1)"
      Tab(12).Control(3).Enabled=   0   'False
      Tab(12).Control(4)=   "TXTpre(14)"
      Tab(12).Control(4).Enabled=   0   'False
      Tab(12).Control(5)=   "TXTpre(11)"
      Tab(12).Control(5).Enabled=   0   'False
      Tab(12).Control(6)=   "TXTpre(12)"
      Tab(12).Control(6).Enabled=   0   'False
      Tab(12).Control(7)=   "escpfela(1)"
      Tab(12).Control(7).Enabled=   0   'False
      Tab(12).Control(8)=   "TXTpre(13)"
      Tab(12).Control(8).Enabled=   0   'False
      Tab(12).Control(9)=   "escidfolha(1)"
      Tab(12).Control(9).Enabled=   0   'False
      Tab(12).Control(10)=   "TXTpre(10)"
      Tab(12).Control(10).Enabled=   0   'False
      Tab(12).Control(11)=   "escidfolha(11)"
      Tab(12).Control(11).Enabled=   0   'False
      Tab(12).Control(12)=   "PCDinipre"
      Tab(12).Control(12).Enabled=   0   'False
      Tab(12).ControlCount=   13
      TabCaption(13)  =   "Composiçao Pre"
      TabPicture(13)  =   "Frmpcx.frx":06F6
      Tab(13).ControlEnabled=   0   'False
      Tab(13).Control(0)=   "Cmddistro(15)"
      Tab(13).Control(1)=   "TXTpre(19)"
      Tab(13).Control(1).Enabled=   0   'False
      Tab(13).Control(2)=   "CmdTotMot(15)"
      Tab(13).Control(3)=   "ComverRev(15)"
      Tab(13).Control(4)=   "EditSeq(15)"
      Tab(13).Control(5)=   "CmdNovaREV(15)"
      Tab(13).Control(6)=   "escTIPPC(15)"
      Tab(13).Control(7)=   "TXTpre(9)"
      Tab(13).Control(7).Enabled=   0   'False
      Tab(13).Control(8)=   "TXTpre(4)"
      Tab(13).Control(8).Enabled=   0   'False
      Tab(13).Control(9)=   "CmdMotRev(15)"
      Tab(13).Control(10)=   "ComMotInc(15)"
      Tab(13).Control(11)=   "ComMotExc(15)"
      Tab(13).Control(12)=   "ComMotAlt(15)"
      Tab(13).Control(13)=   "GRIDCOM(1)"
      Tab(13).Control(14)=   "gridrev(15)"
      Tab(13).Control(15)=   "gridrevi(15)"
      Tab(13).Control(16)=   "lblLabels(33)"
      Tab(13).Control(17)=   "lblLabels(32)"
      Tab(13).ControlCount=   18
      TabCaption(14)  =   "Peça Critica "
      TabPicture(14)  =   "Frmpcx.frx":0712
      Tab(14).ControlEnabled=   0   'False
      Tab(14).Control(0)=   "TXTpc(23)"
      Tab(14).Control(1)=   "Check5"
      Tab(14).ControlCount=   2
      TabCaption(15)  =   "Rel Inspecao"
      TabPicture(15)  =   "Frmpcx.frx":072E
      Tab(15).ControlEnabled=   0   'False
      Tab(15).Control(0)=   "CmdNovoRI(1)"
      Tab(15).Control(1)=   "CmdNovoRI(0)"
      Tab(15).Control(2)=   "CmdGrdRI(1)"
      Tab(15).Control(3)=   "CmdGrdRI(0)"
      Tab(15).Control(4)=   "grdRI"
      Tab(15).ControlCount=   5
      TabCaption(16)  =   "Rec GP12"
      TabPicture(16)  =   "Frmpcx.frx":074A
      Tab(16).ControlEnabled=   0   'False
      Tab(16).Control(0)=   "Cmddistro(21)"
      Tab(16).Control(1)=   "TXTgp12(15)"
      Tab(16).Control(1).Enabled=   0   'False
      Tab(16).Control(2)=   "CmdTotMot(21)"
      Tab(16).Control(3)=   "ComMotAlt(21)"
      Tab(16).Control(4)=   "ComMotExc(21)"
      Tab(16).Control(5)=   "ComMotInc(21)"
      Tab(16).Control(6)=   "CmdMotRev(21)"
      Tab(16).Control(7)=   "TXTgp12(0)"
      Tab(16).Control(7).Enabled=   0   'False
      Tab(16).Control(8)=   "TXTgp12(5)"
      Tab(16).Control(8).Enabled=   0   'False
      Tab(16).Control(9)=   "escTIPPC(21)"
      Tab(16).Control(10)=   "CmdNovaREV(21)"
      Tab(16).Control(11)=   "EditSeq(21)"
      Tab(16).Control(12)=   "ComverRev(21)"
      Tab(16).Control(13)=   "GRIDMAT(2)"
      Tab(16).Control(14)=   "gridrev(21)"
      Tab(16).Control(15)=   "gridrevi(21)"
      Tab(16).Control(16)=   "lblLabels(46)"
      Tab(16).Control(17)=   "lblLabels(45)"
      Tab(16).ControlCount=   18
      TabCaption(17)  =   "Lab GP12"
      TabPicture(17)  =   "Frmpcx.frx":0766
      Tab(17).ControlEnabled=   0   'False
      Tab(17).Control(0)=   "Cmddistro(22)"
      Tab(17).Control(1)=   "TXTgp12(16)"
      Tab(17).Control(1).Enabled=   0   'False
      Tab(17).Control(2)=   "CmdTotMot(22)"
      Tab(17).Control(3)=   "ComMotAlt(22)"
      Tab(17).Control(4)=   "ComMotExc(22)"
      Tab(17).Control(5)=   "ComMotInc(22)"
      Tab(17).Control(6)=   "CmdMotRev(22)"
      Tab(17).Control(7)=   "TXTgp12(1)"
      Tab(17).Control(7).Enabled=   0   'False
      Tab(17).Control(8)=   "TXTgp12(6)"
      Tab(17).Control(8).Enabled=   0   'False
      Tab(17).Control(9)=   "escTIPPC(0)"
      Tab(17).Control(10)=   "CmdNovaREV(22)"
      Tab(17).Control(11)=   "EditSeq(22)"
      Tab(17).Control(12)=   "ComverRev(22)"
      Tab(17).Control(13)=   "gridlab(2)"
      Tab(17).Control(14)=   "gridrev(22)"
      Tab(17).Control(15)=   "gridrevi(22)"
      Tab(17).Control(16)=   "lblLabels(48)"
      Tab(17).Control(17)=   "lblLabels(47)"
      Tab(17).ControlCount=   18
      TabCaption(18)  =   "Prod GP12"
      TabPicture(18)  =   "Frmpcx.frx":0782
      Tab(18).ControlEnabled=   0   'False
      Tab(18).Control(0)=   "Cmddistro(23)"
      Tab(18).Control(1)=   "TXTgp12(17)"
      Tab(18).Control(1).Enabled=   0   'False
      Tab(18).Control(2)=   "CmdTotMot(23)"
      Tab(18).Control(3)=   "ComMotAlt(23)"
      Tab(18).Control(4)=   "ComMotExc(23)"
      Tab(18).Control(5)=   "ComMotInc(23)"
      Tab(18).Control(6)=   "CmdMotRev(23)"
      Tab(18).Control(7)=   "TXTgp12(2)"
      Tab(18).Control(7).Enabled=   0   'False
      Tab(18).Control(8)=   "TXTgp12(7)"
      Tab(18).Control(8).Enabled=   0   'False
      Tab(18).Control(9)=   "escTIPPC(23)"
      Tab(18).Control(10)=   "CmdNovaREV(23)"
      Tab(18).Control(11)=   "EditSeq(23)"
      Tab(18).Control(12)=   "ComverRev(23)"
      Tab(18).Control(13)=   "GridSeq(2)"
      Tab(18).Control(14)=   "gridrev(23)"
      Tab(18).Control(15)=   "gridrevi(23)"
      Tab(18).Control(16)=   "lblLabels(42)"
      Tab(18).Control(17)=   "lblLabels(41)"
      Tab(18).ControlCount=   18
      TabCaption(19)  =   "Final GP12"
      TabPicture(19)  =   "Frmpcx.frx":079E
      Tab(19).ControlEnabled=   0   'False
      Tab(19).Control(0)=   "Cmddistro(24)"
      Tab(19).Control(1)=   "TXTgp12(18)"
      Tab(19).Control(1).Enabled=   0   'False
      Tab(19).Control(2)=   "CmdTotMot(24)"
      Tab(19).Control(3)=   "ComMotAlt(24)"
      Tab(19).Control(4)=   "ComMotExc(24)"
      Tab(19).Control(5)=   "ComMotInc(24)"
      Tab(19).Control(6)=   "CmdMotRev(24)"
      Tab(19).Control(7)=   "TXTgp12(3)"
      Tab(19).Control(7).Enabled=   0   'False
      Tab(19).Control(8)=   "TXTgp12(8)"
      Tab(19).Control(8).Enabled=   0   'False
      Tab(19).Control(9)=   "escTIPPC(24)"
      Tab(19).Control(10)=   "CmdNovaREV(24)"
      Tab(19).Control(11)=   "EditSeq(24)"
      Tab(19).Control(12)=   "ComverRev(24)"
      Tab(19).Control(13)=   "GRIDFIM(2)"
      Tab(19).Control(14)=   "gridrev(24)"
      Tab(19).Control(15)=   "gridrevi(24)"
      Tab(19).Control(16)=   "lblLabels(39)"
      Tab(19).Control(17)=   "lblLabels(38)"
      Tab(19).ControlCount=   18
      TabCaption(20)  =   "ElaRev GP12"
      TabPicture(20)  =   "Frmpcx.frx":07BA
      Tab(20).ControlEnabled=   -1  'True
      Tab(20).Control(0)=   "lblLabels(49)"
      Tab(20).Control(0).Enabled=   0   'False
      Tab(20).Control(1)=   "lblLabels(50)"
      Tab(20).Control(1).Enabled=   0   'False
      Tab(20).Control(2)=   "lblLabels(51)"
      Tab(20).Control(2).Enabled=   0   'False
      Tab(20).Control(3)=   "TXTgp12(13)"
      Tab(20).Control(3).Enabled=   0   'False
      Tab(20).Control(4)=   "escpfela(2)"
      Tab(20).Control(4).Enabled=   0   'False
      Tab(20).Control(5)=   "TXTgp12(12)"
      Tab(20).Control(5).Enabled=   0   'False
      Tab(20).Control(6)=   "TXTgp12(11)"
      Tab(20).Control(6).Enabled=   0   'False
      Tab(20).Control(7)=   "TXTgp12(14)"
      Tab(20).Control(7).Enabled=   0   'False
      Tab(20).Control(8)=   "Command1(2)"
      Tab(20).Control(8).Enabled=   0   'False
      Tab(20).Control(9)=   "escidfolha(2)"
      Tab(20).Control(9).Enabled=   0   'False
      Tab(20).Control(10)=   "TXTgp12(10)"
      Tab(20).Control(10).Enabled=   0   'False
      Tab(20).Control(11)=   "escidfolha(12)"
      Tab(20).Control(11).Enabled=   0   'False
      Tab(20).Control(12)=   "PCDinigp12"
      Tab(20).Control(12).Enabled=   0   'False
      Tab(20).ControlCount=   13
      TabCaption(21)  =   "Comp GP12"
      TabPicture(21)  =   "Frmpcx.frx":07D6
      Tab(21).ControlEnabled=   0   'False
      Tab(21).Control(0)=   "Cmddistro(25)"
      Tab(21).Control(1)=   "TXTgp12(19)"
      Tab(21).Control(1).Enabled=   0   'False
      Tab(21).Control(2)=   "CmdTotMot(25)"
      Tab(21).Control(3)=   "ComMotAlt(25)"
      Tab(21).Control(4)=   "ComMotExc(25)"
      Tab(21).Control(5)=   "ComMotInc(25)"
      Tab(21).Control(6)=   "CmdMotRev(25)"
      Tab(21).Control(7)=   "TXTgp12(4)"
      Tab(21).Control(7).Enabled=   0   'False
      Tab(21).Control(8)=   "TXTgp12(9)"
      Tab(21).Control(8).Enabled=   0   'False
      Tab(21).Control(9)=   "escTIPPC(25)"
      Tab(21).Control(10)=   "CmdNovaREV(25)"
      Tab(21).Control(11)=   "EditSeq(25)"
      Tab(21).Control(12)=   "ComverRev(25)"
      Tab(21).Control(13)=   "GRIDCOM(2)"
      Tab(21).Control(14)=   "gridrev(25)"
      Tab(21).Control(15)=   "gridrevi(25)"
      Tab(21).Control(16)=   "lblLabels(44)"
      Tab(21).Control(17)=   "lblLabels(43)"
      Tab(21).ControlCount=   18
      TabCaption(22)  =   "Outras Contato"
      TabPicture(22)  =   "Frmpcx.frx":07F2
      Tab(22).ControlEnabled=   0   'False
      Tab(22).Control(0)=   "escidfolha(13)"
      Tab(22).Control(0).Enabled=   0   'False
      Tab(22).Control(1)=   "escidfolha(3)"
      Tab(22).Control(1).Enabled=   0   'False
      Tab(22).Control(2)=   "TXTpc(34)"
      Tab(22).Control(2).Enabled=   0   'False
      Tab(22).Control(3)=   "TXTpc(33)"
      Tab(22).Control(3).Enabled=   0   'False
      Tab(22).Control(4)=   "TXTpc(26)"
      Tab(22).Control(5)=   "TXTpc(21)"
      Tab(22).Control(6)=   "lblLabels(20)"
      Tab(22).Control(7)=   "lblLabels(35)"
      Tab(22).ControlCount=   8
      Begin VBCCR17.DTPicker PCDinigp12 
         Height          =   375
         Left            =   1320
         TabIndex        =   404
         Top             =   1560
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         Value           =   44285
         CheckBox        =   -1  'True
         AllowUserInput  =   -1  'True
      End
      Begin VBCCR17.DTPicker PCDinipre 
         Height          =   375
         Left            =   -73680
         TabIndex        =   403
         Top             =   1440
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         Value           =   44285
      End
      Begin VBCCR17.DTPicker PCDinipc 
         Height          =   375
         Left            =   -73680
         TabIndex        =   402
         Top             =   1440
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Value           =   44285
         CheckBox        =   -1  'True
         AllowUserInput  =   -1  'True
      End
      Begin VB.CommandButton escidfolha 
         Caption         =   "LX"
         Height          =   375
         Index           =   13
         Left            =   -72840
         Picture         =   "Frmpcx.frx":080E
         Style           =   1  'Graphical
         TabIndex        =   399
         TabStop         =   0   'False
         ToolTipText     =   "Escolher Nº Folha"
         Top             =   1560
         Width           =   375
      End
      Begin VB.CommandButton escidfolha 
         Caption         =   "LX"
         Height          =   375
         Index           =   12
         Left            =   10200
         Picture         =   "Frmpcx.frx":0D98
         Style           =   1  'Graphical
         TabIndex        =   398
         TabStop         =   0   'False
         ToolTipText     =   "Escolher Nº Folha"
         Top             =   2040
         Width           =   375
      End
      Begin VB.CommandButton escidfolha 
         Caption         =   "LX"
         Height          =   375
         Index           =   11
         Left            =   -64680
         Picture         =   "Frmpcx.frx":1322
         Style           =   1  'Graphical
         TabIndex        =   397
         TabStop         =   0   'False
         ToolTipText     =   "Escolher Nº Folha"
         Top             =   1920
         Width           =   375
      End
      Begin VB.CommandButton escidfolha 
         Caption         =   "LX"
         Height          =   375
         Index           =   10
         Left            =   -64680
         Picture         =   "Frmpcx.frx":18AC
         Style           =   1  'Graphical
         TabIndex        =   396
         TabStop         =   0   'False
         ToolTipText     =   "Escolher Nº Folha"
         Top             =   1800
         Width           =   375
      End
      Begin VB.CommandButton escidfolha 
         Height          =   375
         Index           =   3
         Left            =   -73440
         Picture         =   "Frmpcx.frx":1E36
         Style           =   1  'Graphical
         TabIndex        =   395
         TabStop         =   0   'False
         ToolTipText     =   "Escolher Nº Folha"
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox TXTpc 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   34
         Left            =   -72000
         Locked          =   -1  'True
         TabIndex        =   394
         TabStop         =   0   'False
         Top             =   1560
         Width           =   4215
      End
      Begin VB.TextBox TXTpc 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   33
         Left            =   -74760
         Locked          =   -1  'True
         TabIndex        =   393
         TabStop         =   0   'False
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton Cmddistro 
         Caption         =   "Distribuicao"
         Height          =   375
         Index           =   25
         Left            =   -66480
         TabIndex        =   392
         Top             =   3360
         Width           =   1335
      End
      Begin VB.CommandButton Cmddistro 
         Caption         =   "Distribuicao"
         Height          =   375
         Index           =   24
         Left            =   -66360
         TabIndex        =   391
         Top             =   3480
         Width           =   1335
      End
      Begin VB.CommandButton Cmddistro 
         Caption         =   "Distribuicao"
         Height          =   375
         Index           =   23
         Left            =   -66240
         TabIndex        =   390
         Top             =   3240
         Width           =   1335
      End
      Begin VB.CommandButton Cmddistro 
         Caption         =   "Distribuicao"
         Height          =   375
         Index           =   22
         Left            =   -66360
         TabIndex        =   389
         Top             =   3240
         Width           =   1335
      End
      Begin VB.CommandButton Cmddistro 
         Caption         =   "Distribuicao"
         Height          =   375
         Index           =   21
         Left            =   -66480
         TabIndex        =   388
         Top             =   3120
         Width           =   1335
      End
      Begin VB.CommandButton Cmddistro 
         Caption         =   "Distribuicao"
         Height          =   375
         Index           =   15
         Left            =   -66480
         TabIndex        =   387
         Top             =   3360
         Width           =   1335
      End
      Begin VB.CommandButton Cmddistro 
         Caption         =   "Distribuicao"
         Height          =   375
         Index           =   14
         Left            =   -66480
         TabIndex        =   386
         Top             =   3480
         Width           =   1455
      End
      Begin VB.CommandButton Cmddistro 
         Caption         =   "Distribuicao"
         Height          =   375
         Index           =   13
         Left            =   -66360
         TabIndex        =   385
         Top             =   3240
         Width           =   1335
      End
      Begin VB.CommandButton Cmddistro 
         Caption         =   "Distribuicao"
         Height          =   375
         Index           =   12
         Left            =   -66360
         TabIndex        =   384
         Top             =   3120
         Width           =   1335
      End
      Begin VB.CommandButton Cmddistro 
         Caption         =   "Distribuicao"
         Height          =   375
         Index           =   11
         Left            =   -66360
         TabIndex        =   383
         Top             =   3120
         Width           =   1335
      End
      Begin VB.CommandButton Cmddistro 
         Caption         =   "Distribuicao"
         Height          =   375
         Index           =   6
         Left            =   -66000
         TabIndex        =   382
         Top             =   2520
         Width           =   975
      End
      Begin VB.CommandButton Cmddistro 
         Caption         =   "Distribuicao"
         Height          =   375
         Index           =   5
         Left            =   -66480
         TabIndex        =   381
         Top             =   3240
         Width           =   1335
      End
      Begin VB.CommandButton Cmddistro 
         Caption         =   "Distribuicao"
         Height          =   375
         Index           =   7
         Left            =   -66000
         TabIndex        =   380
         Top             =   3000
         Width           =   1095
      End
      Begin VB.CommandButton Cmddistro 
         Caption         =   "Distribuicao"
         Height          =   255
         Index           =   4
         Left            =   -66480
         TabIndex        =   379
         Top             =   3840
         Width           =   1335
      End
      Begin VB.CommandButton Cmddistro 
         Caption         =   "Distribuicao"
         Height          =   375
         Index           =   3
         Left            =   -66480
         TabIndex        =   378
         Top             =   4440
         Width           =   1335
      End
      Begin VB.CommandButton Cmddistro 
         Caption         =   "Distribuicao"
         Height          =   375
         Index           =   2
         Left            =   -66360
         TabIndex        =   377
         Top             =   4920
         Width           =   1335
      End
      Begin VB.CommandButton Cmddistro 
         Caption         =   "Distribuicao"
         Height          =   375
         Index           =   1
         Left            =   -66360
         TabIndex        =   376
         Top             =   4800
         Width           =   1335
      End
      Begin VB.CommandButton Cmddistro 
         Caption         =   "Distribuicao"
         Height          =   375
         Index           =   0
         Left            =   -66480
         TabIndex        =   375
         Top             =   3840
         Width           =   1335
      End
      Begin VB.TextBox TXTpc 
         BackColor       =   &H00C0FFFF&
         DataField       =   "PCELANUM"
         DataSource      =   "datPrimaryRS"
         Enabled         =   0   'False
         Height          =   285
         Index           =   32
         Left            =   -66120
         Locked          =   -1  'True
         TabIndex        =   374
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton ComMotInc 
         Caption         =   "PF->Revisao"
         Height          =   255
         Index           =   26
         Left            =   -66120
         TabIndex        =   373
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox TXTPF 
         BackColor       =   &H00C0FFFF&
         DataField       =   "FEMEAREVD"
         DataSource      =   "datPrimaryRS"
         Enabled         =   0   'False
         Height          =   285
         Index           =   19
         Left            =   -65940
         Locked          =   -1  'True
         TabIndex        =   372
         TabStop         =   0   'False
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox TXTpc 
         BackColor       =   &H00C0FFFF&
         DataField       =   "PCELANUM"
         DataSource      =   "datPrimaryRS"
         Enabled         =   0   'False
         Height          =   285
         Index           =   19
         Left            =   -71640
         Locked          =   -1  'True
         TabIndex        =   371
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox TXTpc 
         BackColor       =   &H00C0FFFF&
         DataField       =   "PCELANUM"
         DataSource      =   "datPrimaryRS"
         Enabled         =   0   'False
         Height          =   285
         Index           =   18
         Left            =   -71400
         Locked          =   -1  'True
         TabIndex        =   370
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox TXTpc 
         BackColor       =   &H00C0FFFF&
         DataField       =   "PCELANUM"
         DataSource      =   "datPrimaryRS"
         Enabled         =   0   'False
         Height          =   285
         Index           =   17
         Left            =   -72840
         Locked          =   -1  'True
         TabIndex        =   369
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox TXTpc 
         BackColor       =   &H00C0FFFF&
         DataField       =   "PCELANUM"
         DataSource      =   "datPrimaryRS"
         Enabled         =   0   'False
         Height          =   285
         Index           =   16
         Left            =   -72600
         Locked          =   -1  'True
         TabIndex        =   368
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox TXTpc 
         BackColor       =   &H00C0FFFF&
         DataField       =   "PCELANUM"
         DataSource      =   "datPrimaryRS"
         Enabled         =   0   'False
         Height          =   285
         Index           =   15
         Left            =   -71520
         Locked          =   -1  'True
         TabIndex        =   367
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox TXTpc 
         BackColor       =   &H00C0FFFF&
         DataField       =   "PCELANUM"
         DataSource      =   "datPrimaryRS"
         Enabled         =   0   'False
         Height          =   285
         Index           =   12
         Left            =   -73800
         Locked          =   -1  'True
         TabIndex        =   366
         TabStop         =   0   'False
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox TXTgp12 
         BackColor       =   &H00C0FFFF&
         DataField       =   "PCRREV"
         DataSource      =   "DataGP12"
         Enabled         =   0   'False
         Height          =   285
         Index           =   19
         Left            =   -71640
         Locked          =   -1  'True
         TabIndex        =   365
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox TXTgp12 
         BackColor       =   &H00C0FFFF&
         DataField       =   "PCRREV"
         DataSource      =   "DataGP12"
         Enabled         =   0   'False
         Height          =   285
         Index           =   18
         Left            =   -71400
         Locked          =   -1  'True
         TabIndex        =   364
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox TXTgp12 
         BackColor       =   &H00C0FFFF&
         DataField       =   "PCRREV"
         DataSource      =   "DataGP12"
         Enabled         =   0   'False
         Height          =   285
         Index           =   17
         Left            =   -72720
         Locked          =   -1  'True
         TabIndex        =   363
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox TXTgp12 
         BackColor       =   &H00C0FFFF&
         DataField       =   "PCRREV"
         DataSource      =   "DataGP12"
         Enabled         =   0   'False
         Height          =   285
         Index           =   16
         Left            =   -72840
         Locked          =   -1  'True
         TabIndex        =   362
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox TXTgp12 
         BackColor       =   &H00C0FFFF&
         DataField       =   "PCRREV"
         DataSource      =   "DataGP12"
         Enabled         =   0   'False
         Height          =   285
         Index           =   15
         Left            =   -71520
         Locked          =   -1  'True
         TabIndex        =   361
         TabStop         =   0   'False
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox TXTgp12 
         BackColor       =   &H00C0FFFF&
         DataField       =   "PCELANUM"
         DataSource      =   "DataGP12"
         Enabled         =   0   'False
         Height          =   285
         Index           =   10
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   360
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox TXTpre 
         BackColor       =   &H00C0FFFF&
         DataField       =   "PCRREV"
         DataSource      =   "DataPreliminar"
         Enabled         =   0   'False
         Height          =   285
         Index           =   19
         Left            =   -71520
         Locked          =   -1  'True
         TabIndex        =   359
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox TXTpre 
         BackColor       =   &H00C0FFFF&
         DataField       =   "PCRREV"
         DataSource      =   "DataPreliminar"
         Enabled         =   0   'False
         Height          =   285
         Index           =   18
         Left            =   -71280
         Locked          =   -1  'True
         TabIndex        =   358
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox TXTpre 
         BackColor       =   &H00C0FFFF&
         DataField       =   "PCRREV"
         DataSource      =   "DataPreliminar"
         Enabled         =   0   'False
         Height          =   285
         Index           =   17
         Left            =   -72720
         Locked          =   -1  'True
         TabIndex        =   357
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox TXTpre 
         BackColor       =   &H00C0FFFF&
         DataField       =   "PCRREV"
         DataSource      =   "DataPreliminar"
         Enabled         =   0   'False
         Height          =   285
         Index           =   16
         Left            =   -72720
         Locked          =   -1  'True
         TabIndex        =   356
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox TXTpre 
         BackColor       =   &H00C0FFFF&
         DataField       =   "PCRREV"
         DataSource      =   "DataPreliminar"
         Enabled         =   0   'False
         Height          =   285
         Index           =   15
         Left            =   -71520
         Locked          =   -1  'True
         TabIndex        =   355
         TabStop         =   0   'False
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox TXTpre 
         BackColor       =   &H00C0FFFF&
         DataField       =   "PCELANUM"
         DataSource      =   "DataPreliminar"
         Enabled         =   0   'False
         Height          =   285
         Index           =   10
         Left            =   -67440
         Locked          =   -1  'True
         TabIndex        =   354
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton CmdMarcaPC 
         Caption         =   "OK-Prod"
         Height          =   375
         Index           =   4
         Left            =   -66000
         TabIndex        =   353
         Top             =   4440
         Width           =   975
      End
      Begin VB.CommandButton CmdMarcaPC 
         Caption         =   "OK-COM"
         Height          =   375
         Index           =   3
         Left            =   -66000
         TabIndex        =   352
         Top             =   4080
         Width           =   975
      End
      Begin VB.CommandButton CmdMarcaPC 
         Caption         =   "OK-FIN"
         Height          =   375
         Index           =   2
         Left            =   -66000
         TabIndex        =   351
         Top             =   3720
         Width           =   975
      End
      Begin VB.CommandButton CmdMarcaPC 
         Caption         =   "OK-LAB"
         Height          =   375
         Index           =   1
         Left            =   -66000
         TabIndex        =   350
         Top             =   3360
         Width           =   975
      End
      Begin VB.CommandButton CmdMarcaPC 
         Caption         =   "OK-Rec"
         Height          =   375
         Index           =   0
         Left            =   -66000
         TabIndex        =   348
         Top             =   3000
         Width           =   975
      End
      Begin VB.CommandButton escidfolha 
         Height          =   375
         Index           =   2
         Left            =   9720
         Picture         =   "Frmpcx.frx":23C0
         Style           =   1  'Graphical
         TabIndex        =   347
         TabStop         =   0   'False
         ToolTipText     =   "Escolher Nº Folha"
         Top             =   2040
         Width           =   375
      End
      Begin VB.CommandButton escidfolha 
         Height          =   375
         Index           =   1
         Left            =   -65280
         Picture         =   "Frmpcx.frx":294A
         Style           =   1  'Graphical
         TabIndex        =   346
         TabStop         =   0   'False
         ToolTipText     =   "Escolher Nº Folha"
         Top             =   1920
         Width           =   375
      End
      Begin VB.CommandButton escidfolha 
         Height          =   375
         Index           =   0
         Left            =   -65160
         Picture         =   "Frmpcx.frx":2ED4
         Style           =   1  'Graphical
         TabIndex        =   345
         TabStop         =   0   'False
         ToolTipText     =   "Escolher Nº Folha"
         Top             =   1800
         Width           =   375
      End
      Begin VB.CommandButton CmdTotMot 
         Caption         =   "Todos Motivos"
         Height          =   255
         Index           =   0
         Left            =   -66480
         TabIndex        =   343
         Top             =   3600
         Width           =   1335
      End
      Begin VB.CommandButton CmdTotMot 
         Caption         =   "Todos Motivos"
         Height          =   255
         Index           =   25
         Left            =   -66540
         TabIndex        =   342
         Top             =   3120
         Width           =   1335
      End
      Begin VB.CommandButton CmdTotMot 
         Caption         =   "Todos Motivos"
         Height          =   255
         Index           =   24
         Left            =   -66360
         TabIndex        =   341
         Top             =   3240
         Width           =   1335
      End
      Begin VB.CommandButton CmdTotMot 
         Caption         =   "Todos Motivos"
         Height          =   255
         Index           =   23
         Left            =   -66240
         TabIndex        =   340
         Top             =   3000
         Width           =   1335
      End
      Begin VB.CommandButton CmdTotMot 
         Caption         =   "Todos Motivos"
         Height          =   255
         Index           =   15
         Left            =   -66480
         TabIndex        =   339
         Top             =   3120
         Width           =   1335
      End
      Begin VB.CommandButton CmdTotMot 
         Caption         =   "Todos Motivos"
         Height          =   255
         Index           =   5
         Left            =   -66480
         TabIndex        =   338
         Top             =   3000
         Width           =   1335
      End
      Begin VB.CommandButton CmdTotMot 
         Caption         =   "Todos Motivos"
         Height          =   255
         Index           =   14
         Left            =   -66420
         TabIndex        =   337
         Top             =   3240
         Width           =   1335
      End
      Begin VB.CommandButton CmdTotMot 
         Caption         =   "Todos Motivos"
         Height          =   255
         Index           =   4
         Left            =   -66540
         TabIndex        =   336
         Top             =   3600
         Width           =   1335
      End
      Begin VB.CommandButton CmdTotMot 
         Caption         =   "Todos Motivos"
         Height          =   255
         Index           =   22
         Left            =   -66360
         TabIndex        =   335
         Top             =   3000
         Width           =   1335
      End
      Begin VB.CommandButton CmdTotMot 
         Caption         =   "Todos Motivos"
         Height          =   255
         Index           =   13
         Left            =   -66360
         TabIndex        =   334
         Top             =   3000
         Width           =   1335
      End
      Begin VB.CommandButton CmdTotMot 
         Caption         =   "Todos Motivos"
         Height          =   255
         Index           =   12
         Left            =   -66360
         TabIndex        =   333
         Top             =   2880
         Width           =   1335
      End
      Begin VB.CommandButton CmdTotMot 
         Caption         =   "Todos Motivos"
         Height          =   255
         Index           =   2
         Left            =   -66420
         TabIndex        =   332
         Top             =   3180
         Width           =   1335
      End
      Begin VB.CommandButton CmdTotMot 
         Caption         =   "Todos Motivos"
         Height          =   255
         Index           =   21
         Left            =   -66480
         TabIndex        =   331
         Top             =   2880
         Width           =   1335
      End
      Begin VB.CommandButton CmdTotMot 
         Caption         =   "Todos Motivos"
         Height          =   255
         Index           =   11
         Left            =   -66360
         TabIndex        =   330
         Top             =   2880
         Width           =   1335
      End
      Begin VB.CommandButton CmdTotMot 
         Caption         =   "Todos Motivos"
         Height          =   255
         Index           =   1
         Left            =   -66420
         TabIndex        =   329
         Top             =   3180
         Width           =   1335
      End
      Begin VB.CommandButton CmdTotMot 
         Caption         =   "Todos Motivos"
         Height          =   255
         Index           =   3
         Left            =   -66480
         TabIndex        =   328
         Top             =   4200
         Width           =   1335
      End
      Begin VB.TextBox TXTpc 
         DataField       =   "CLIESP"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   26
         Left            =   -68940
         TabIndex        =   325
         Top             =   1080
         Width           =   2775
      End
      Begin VB.TextBox TXTpc 
         DataField       =   "OUTRAAPR"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   21
         Left            =   -73080
         TabIndex        =   323
         Top             =   1080
         Width           =   2775
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Grupo Femea"
         Height          =   435
         Index           =   2
         Left            =   120
         TabIndex        =   319
         Top             =   2760
         Width           =   735
      End
      Begin VB.TextBox TXTgp12 
         BackColor       =   &H00C0FFFF&
         DataField       =   "FEMEAG"
         DataSource      =   "DataPreliminar"
         Enabled         =   0   'False
         Height          =   525
         Index           =   14
         Left            =   840
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   318
         TabStop         =   0   'False
         Top             =   2700
         Width           =   8895
      End
      Begin VB.TextBox TXTgp12 
         DataField       =   "PCGRUPO"
         DataSource      =   "DataGP12"
         Height          =   585
         Index           =   11
         Left            =   840
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   317
         Top             =   2040
         Width           =   8895
      End
      Begin VB.TextBox TXTgp12 
         BackColor       =   &H00C0FFFF&
         DataField       =   "PCELANUM"
         DataSource      =   "DataGP12"
         Enabled         =   0   'False
         Height          =   285
         Index           =   12
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   316
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton escpfela 
         Height          =   315
         Index           =   2
         Left            =   2460
         Picture         =   "Frmpcx.frx":345E
         Style           =   1  'Graphical
         TabIndex        =   315
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox TXTgp12 
         BackColor       =   &H00C0FFFF&
         DataField       =   "PCELANOM"
         DataSource      =   "DataGP12"
         Enabled         =   0   'False
         Height          =   285
         Index           =   13
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   314
         TabStop         =   0   'False
         Top             =   1080
         Width           =   4215
      End
      Begin VB.CommandButton ComMotAlt 
         Caption         =   "Alterar"
         Height          =   375
         Index           =   22
         Left            =   -66300
         TabIndex        =   308
         Top             =   3720
         Width           =   1335
      End
      Begin VB.CommandButton ComMotExc 
         Caption         =   "Excluir"
         Height          =   375
         Index           =   22
         Left            =   -66300
         TabIndex        =   307
         Top             =   4380
         Width           =   1335
      End
      Begin VB.CommandButton ComMotInc 
         Caption         =   "Incluir"
         Height          =   375
         Index           =   22
         Left            =   -66300
         TabIndex        =   306
         Top             =   4020
         Width           =   1335
      End
      Begin VB.CommandButton CmdMotRev 
         Caption         =   "Motivos"
         Height          =   255
         Index           =   22
         Left            =   -66360
         TabIndex        =   305
         Top             =   2760
         Width           =   1335
      End
      Begin VB.TextBox TXTgp12 
         BackColor       =   &H00C0FFFF&
         DataField       =   "PCLREV"
         DataSource      =   "DataGP12"
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   -74040
         Locked          =   -1  'True
         TabIndex        =   304
         TabStop         =   0   'False
         Top             =   1140
         Width           =   1095
      End
      Begin VB.TextBox TXTgp12 
         BackColor       =   &H00C0FFFF&
         DataField       =   "PCLTIPO"
         DataSource      =   "DataGP12"
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   -69240
         Locked          =   -1  'True
         TabIndex        =   303
         TabStop         =   0   'False
         Top             =   1200
         Width           =   255
      End
      Begin VB.CommandButton escTIPPC 
         Height          =   315
         Index           =   0
         Left            =   -68820
         Picture         =   "Frmpcx.frx":39E8
         Style           =   1  'Graphical
         TabIndex        =   302
         Top             =   1200
         Width           =   435
      End
      Begin VB.CommandButton CmdNovaREV 
         Caption         =   "Nova Revisão"
         Height          =   375
         Index           =   22
         Left            =   -71160
         TabIndex        =   301
         Top             =   1140
         Width           =   1335
      End
      Begin VB.CommandButton EditSeq 
         Caption         =   "Editar"
         Height          =   255
         Index           =   22
         Left            =   -66360
         TabIndex        =   300
         Top             =   1620
         Width           =   1335
      End
      Begin VB.CommandButton ComverRev 
         Caption         =   "Ver Revisoes"
         Height          =   315
         Index           =   22
         Left            =   -66360
         TabIndex        =   299
         Top             =   2400
         Width           =   1335
      End
      Begin VB.CommandButton ComMotAlt 
         Caption         =   "Alterar"
         Height          =   375
         Index           =   21
         Left            =   -66420
         TabIndex        =   293
         Top             =   3540
         Width           =   1335
      End
      Begin VB.CommandButton ComMotExc 
         Caption         =   "Excluir"
         Height          =   375
         Index           =   21
         Left            =   -66420
         TabIndex        =   292
         Top             =   4320
         Width           =   1335
      End
      Begin VB.CommandButton ComMotInc 
         Caption         =   "Incluir"
         Height          =   375
         Index           =   21
         Left            =   -66420
         TabIndex        =   291
         Top             =   3900
         Width           =   1335
      End
      Begin VB.CommandButton CmdMotRev 
         Caption         =   "Motivos"
         Height          =   255
         Index           =   21
         Left            =   -66480
         TabIndex        =   290
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox TXTgp12 
         BackColor       =   &H00C0FFFF&
         DataField       =   "PCRREV"
         DataSource      =   "DataGP12"
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   -74100
         Locked          =   -1  'True
         TabIndex        =   289
         TabStop         =   0   'False
         Top             =   1020
         Width           =   1095
      End
      Begin VB.TextBox TXTgp12 
         BackColor       =   &H00C0FFFF&
         DataField       =   "PCRTIPO"
         DataSource      =   "DataGP12"
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   -69420
         Locked          =   -1  'True
         TabIndex        =   288
         TabStop         =   0   'False
         Top             =   1020
         Width           =   255
      End
      Begin VB.CommandButton escTIPPC 
         Height          =   315
         Index           =   21
         Left            =   -69060
         Picture         =   "Frmpcx.frx":3F72
         Style           =   1  'Graphical
         TabIndex        =   287
         Top             =   1020
         Width           =   435
      End
      Begin VB.CommandButton CmdNovaREV 
         Caption         =   "Nova Revisão"
         Height          =   255
         Index           =   21
         Left            =   -72960
         TabIndex        =   286
         Top             =   1020
         Width           =   1335
      End
      Begin VB.CommandButton EditSeq 
         Caption         =   "Editar"
         Height          =   255
         Index           =   21
         Left            =   -66480
         TabIndex        =   285
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton ComverRev 
         Caption         =   "Ver Revisoes"
         Height          =   315
         Index           =   21
         Left            =   -66480
         TabIndex        =   284
         Top             =   2280
         Width           =   1335
      End
      Begin VB.CommandButton ComMotAlt 
         Caption         =   "Alterar"
         Height          =   375
         Index           =   25
         Left            =   -66480
         TabIndex        =   278
         Top             =   3900
         Width           =   1335
      End
      Begin VB.CommandButton ComMotExc 
         Caption         =   "Excluir"
         Height          =   375
         Index           =   25
         Left            =   -66480
         TabIndex        =   277
         Top             =   4620
         Width           =   1335
      End
      Begin VB.CommandButton ComMotInc 
         Caption         =   "Incluir"
         Height          =   375
         Index           =   25
         Left            =   -66480
         TabIndex        =   276
         Top             =   4260
         Width           =   1335
      End
      Begin VB.CommandButton CmdMotRev 
         Caption         =   "Motivos"
         Height          =   255
         Index           =   25
         Left            =   -66540
         TabIndex        =   275
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox TXTgp12 
         BackColor       =   &H00C0FFFF&
         DataField       =   "PCCREV"
         DataSource      =   "DataGP12"
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   -74040
         Locked          =   -1  'True
         TabIndex        =   274
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox TXTgp12 
         BackColor       =   &H00C0FFFF&
         DataField       =   "PCCTIPO"
         DataSource      =   "DataGP12"
         Enabled         =   0   'False
         Height          =   285
         Index           =   9
         Left            =   -69360
         Locked          =   -1  'True
         TabIndex        =   273
         TabStop         =   0   'False
         Top             =   1140
         Width           =   255
      End
      Begin VB.CommandButton escTIPPC 
         Height          =   315
         Index           =   25
         Left            =   -69060
         Picture         =   "Frmpcx.frx":44FC
         Style           =   1  'Graphical
         TabIndex        =   272
         Top             =   1140
         Width           =   435
      End
      Begin VB.CommandButton CmdNovaREV 
         Caption         =   "Nova Revisão"
         Height          =   315
         Index           =   25
         Left            =   -72900
         TabIndex        =   271
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton EditSeq 
         Caption         =   "Editar"
         Height          =   255
         Index           =   25
         Left            =   -66600
         TabIndex        =   270
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CommandButton ComverRev 
         Caption         =   "Ver Revisoes"
         Height          =   315
         Index           =   25
         Left            =   -66540
         TabIndex        =   269
         Top             =   2520
         Width           =   1335
      End
      Begin VB.CommandButton ComMotAlt 
         Caption         =   "Alterar"
         Height          =   375
         Index           =   23
         Left            =   -66240
         TabIndex        =   251
         Top             =   3720
         Width           =   1335
      End
      Begin VB.CommandButton ComMotExc 
         Caption         =   "Excluir"
         Height          =   375
         Index           =   23
         Left            =   -66240
         TabIndex        =   250
         Top             =   4440
         Width           =   1335
      End
      Begin VB.CommandButton ComMotInc 
         Caption         =   "Incluir"
         Height          =   375
         Index           =   23
         Left            =   -66240
         TabIndex        =   249
         Top             =   4080
         Width           =   1335
      End
      Begin VB.CommandButton CmdMotRev 
         Caption         =   "Motivos"
         Height          =   255
         Index           =   23
         Left            =   -66240
         TabIndex        =   248
         Top             =   2760
         Width           =   1335
      End
      Begin VB.TextBox TXTgp12 
         BackColor       =   &H00C0FFFF&
         DataField       =   "PCPREV"
         DataSource      =   "DataGP12"
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   -73980
         Locked          =   -1  'True
         TabIndex        =   247
         TabStop         =   0   'False
         Top             =   1020
         Width           =   1095
      End
      Begin VB.TextBox TXTgp12 
         BackColor       =   &H00C0FFFF&
         DataField       =   "PCPTIPO"
         DataSource      =   "DataPreliminar"
         Enabled         =   0   'False
         Height          =   285
         Index           =   7
         Left            =   -69300
         Locked          =   -1  'True
         TabIndex        =   246
         TabStop         =   0   'False
         Top             =   1020
         Width           =   255
      End
      Begin VB.CommandButton escTIPPC 
         Height          =   315
         Index           =   23
         Left            =   -68940
         Picture         =   "Frmpcx.frx":4A86
         Style           =   1  'Graphical
         TabIndex        =   245
         Top             =   1020
         Width           =   495
      End
      Begin VB.CommandButton CmdNovaREV 
         Caption         =   "Nova Revisão"
         Height          =   375
         Index           =   23
         Left            =   -71220
         TabIndex        =   244
         Top             =   1020
         Width           =   1335
      End
      Begin VB.CommandButton EditSeq 
         Caption         =   "Editar"
         Height          =   255
         Index           =   23
         Left            =   -66300
         TabIndex        =   243
         Top             =   1500
         Width           =   1335
      End
      Begin VB.CommandButton ComverRev 
         Caption         =   "Ver Revisoes"
         Height          =   315
         Index           =   23
         Left            =   -66240
         TabIndex        =   242
         Top             =   2400
         Width           =   1335
      End
      Begin VB.CommandButton ComMotAlt 
         Caption         =   "Alterar"
         Height          =   375
         Index           =   24
         Left            =   -66360
         TabIndex        =   236
         Top             =   4020
         Width           =   1335
      End
      Begin VB.CommandButton ComMotExc 
         Caption         =   "Excluir"
         Height          =   375
         Index           =   24
         Left            =   -66360
         TabIndex        =   235
         Top             =   4740
         Width           =   1335
      End
      Begin VB.CommandButton ComMotInc 
         Caption         =   "Incluir"
         Height          =   375
         Index           =   24
         Left            =   -66360
         TabIndex        =   234
         Top             =   4380
         Width           =   1335
      End
      Begin VB.CommandButton CmdMotRev 
         Caption         =   "Motivos"
         Height          =   255
         Index           =   24
         Left            =   -66360
         TabIndex        =   233
         Top             =   3000
         Width           =   1335
      End
      Begin VB.TextBox TXTgp12 
         BackColor       =   &H00C0FFFF&
         DataField       =   "PCFREV"
         DataSource      =   "DataGP12"
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   -73920
         Locked          =   -1  'True
         TabIndex        =   232
         TabStop         =   0   'False
         Top             =   1020
         Width           =   1095
      End
      Begin VB.TextBox TXTgp12 
         BackColor       =   &H00C0FFFF&
         DataField       =   "PCFTIPO"
         DataSource      =   "DataGP12"
         Enabled         =   0   'False
         Height          =   285
         Index           =   8
         Left            =   -69120
         Locked          =   -1  'True
         TabIndex        =   231
         TabStop         =   0   'False
         Top             =   1080
         Width           =   255
      End
      Begin VB.CommandButton escTIPPC 
         Height          =   315
         Index           =   24
         Left            =   -68760
         Picture         =   "Frmpcx.frx":5010
         Style           =   1  'Graphical
         TabIndex        =   230
         Top             =   1080
         Width           =   495
      End
      Begin VB.CommandButton CmdNovaREV 
         Caption         =   "Nova Revisão"
         Height          =   375
         Index           =   24
         Left            =   -72720
         TabIndex        =   229
         Top             =   1020
         Width           =   1215
      End
      Begin VB.CommandButton EditSeq 
         Caption         =   "Editar"
         Height          =   255
         Index           =   24
         Left            =   -66420
         TabIndex        =   228
         Top             =   1860
         Width           =   1335
      End
      Begin VB.CommandButton ComverRev 
         Caption         =   "Ver Revisoes"
         Height          =   255
         Index           =   24
         Left            =   -66360
         TabIndex        =   227
         Top             =   2760
         Width           =   1335
      End
      Begin VB.CommandButton escTIPPC 
         Height          =   315
         Index           =   2
         Left            =   -68460
         Picture         =   "Frmpcx.frx":559A
         Style           =   1  'Graphical
         TabIndex        =   226
         Top             =   1020
         Width           =   555
      End
      Begin VB.TextBox TXTPF 
         BackColor       =   &H00C0FFFF&
         DataField       =   "FEMEAREV"
         DataSource      =   "datPrimaryRS"
         Enabled         =   0   'False
         Height          =   285
         Index           =   15
         Left            =   -65940
         Locked          =   -1  'True
         TabIndex        =   223
         TabStop         =   0   'False
         Top             =   1080
         Width           =   675
      End
      Begin VB.TextBox TXTPF 
         BackColor       =   &H00C0FFFF&
         DataField       =   "FEMEAREVD"
         DataSource      =   "datPrimaryRS"
         Enabled         =   0   'False
         Height          =   285
         Index           =   16
         Left            =   -65940
         Locked          =   -1  'True
         TabIndex        =   222
         TabStop         =   0   'False
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox TXTPF 
         BackColor       =   &H00C0FFFF&
         DataField       =   "REVPRO"
         DataSource      =   "datPrimaryRS"
         Enabled         =   0   'False
         Height          =   285
         Index           =   17
         Left            =   -66060
         Locked          =   -1  'True
         TabIndex        =   221
         TabStop         =   0   'False
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox TXTPF 
         BackColor       =   &H00C0FFFF&
         DataField       =   "REVDAT"
         DataSource      =   "datPrimaryRS"
         Enabled         =   0   'False
         Height          =   285
         Index           =   18
         Left            =   -66060
         Locked          =   -1  'True
         TabIndex        =   220
         TabStop         =   0   'False
         Top             =   1440
         Width           =   975
      End
      Begin VB.CommandButton ESCprofin 
         Caption         =   "LX"
         Height          =   315
         Index           =   1
         Left            =   -65400
         TabIndex        =   219
         TabStop         =   0   'False
         Top             =   1020
         Width           =   375
      End
      Begin VB.CommandButton ComverRev 
         Caption         =   "Ver Revisoes"
         Height          =   435
         Index           =   7
         Left            =   -65940
         TabIndex        =   216
         Top             =   2160
         Width           =   1035
      End
      Begin VB.CommandButton ComverRev 
         Caption         =   "Ver Revisoes"
         Height          =   435
         Index           =   6
         Left            =   -66000
         TabIndex        =   215
         Top             =   1800
         Width           =   915
      End
      Begin VB.CommandButton ComverRev 
         Caption         =   "Ver Revisoes"
         Height          =   315
         Index           =   15
         Left            =   -66480
         TabIndex        =   214
         Top             =   2520
         Width           =   1335
      End
      Begin VB.CommandButton ComverRev 
         Caption         =   "Ver Revisoes"
         Height          =   315
         Index           =   5
         Left            =   -66480
         TabIndex        =   213
         Top             =   2280
         Width           =   1335
      End
      Begin VB.CommandButton ComverRev 
         Caption         =   "Ver Revisoes"
         Height          =   315
         Index           =   0
         Left            =   -66480
         TabIndex        =   212
         Top             =   2940
         Width           =   1335
      End
      Begin VB.CommandButton ComverRev 
         Caption         =   "Ver Revisoes"
         Height          =   315
         Index           =   4
         Left            =   -66540
         TabIndex        =   211
         Top             =   3000
         Width           =   1335
      End
      Begin VB.CommandButton ComverRev 
         Caption         =   "Ver Revisoes"
         Height          =   315
         Index           =   14
         Left            =   -66420
         TabIndex        =   210
         Top             =   2640
         Width           =   1335
      End
      Begin VB.CommandButton ComverRev 
         Caption         =   "Ver Revisoes"
         Height          =   315
         Index           =   13
         Left            =   -66360
         TabIndex        =   209
         Top             =   2400
         Width           =   1335
      End
      Begin VB.CommandButton ComverRev 
         Caption         =   "Ver Revisoes"
         Height          =   315
         Index           =   3
         Left            =   -66480
         TabIndex        =   208
         Top             =   3600
         Width           =   1335
      End
      Begin VB.CommandButton ComverRev 
         Caption         =   "Ver Revisoes"
         Height          =   315
         Index           =   2
         Left            =   -66420
         TabIndex        =   207
         Top             =   2520
         Width           =   1335
      End
      Begin VB.CommandButton ComverRev 
         Caption         =   "Ver Revisoes"
         Height          =   255
         Index           =   12
         Left            =   -66360
         TabIndex        =   206
         Top             =   2280
         Width           =   1335
      End
      Begin VB.CommandButton ComverRev 
         Caption         =   "Ver Revisoes"
         Height          =   315
         Index           =   11
         Left            =   -66360
         TabIndex        =   205
         Top             =   2280
         Width           =   1335
      End
      Begin VB.CommandButton ComverRev 
         Caption         =   "Ver Revisoes"
         Height          =   315
         Index           =   1
         Left            =   -66420
         TabIndex        =   204
         Top             =   2580
         Width           =   1335
      End
      Begin VB.CommandButton CmdNovoRI 
         Appearance      =   0  'Flat
         Caption         =   "Novo RI PC-Producao Preliminar"
         Height          =   495
         Index           =   1
         Left            =   -67440
         TabIndex        =   202
         Top             =   2580
         Width           =   2175
      End
      Begin VB.CommandButton CmdNovoRI 
         Appearance      =   0  'Flat
         Caption         =   "Novo RI PC-Producao"
         Height          =   495
         Index           =   0
         Left            =   -67440
         TabIndex        =   201
         Top             =   2040
         Width           =   2175
      End
      Begin VB.CommandButton CmdGrdRI 
         Caption         =   "Ver RI Codigo Final"
         Height          =   375
         Index           =   1
         Left            =   -67440
         TabIndex        =   200
         Top             =   1620
         Width           =   1575
      End
      Begin VB.CommandButton CmdGrdRI 
         Caption         =   "Ver RI Codigo"
         Height          =   375
         Index           =   0
         Left            =   -67440
         TabIndex        =   199
         Top             =   1200
         Width           =   1575
      End
      Begin MSFlexGridLib.MSFlexGrid grdRI 
         Height          =   3855
         Left            =   -74820
         TabIndex        =   198
         Top             =   1020
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   6800
         _Version        =   393216
      End
      Begin VB.TextBox TXTpc 
         DataField       =   "CRITIOBS"
         DataSource      =   "datPrimaryRS"
         Height          =   1455
         Index           =   23
         Left            =   -74760
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   197
         Top             =   2010
         Width           =   9495
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Produto Critico"
         DataField       =   "CRITICO"
         DataSource      =   "datPrimaryRS"
         Height          =   315
         Left            =   -74700
         TabIndex        =   196
         Top             =   1650
         Width           =   1695
      End
      Begin VB.TextBox TXTpc 
         BackColor       =   &H00C0FFFF&
         DataField       =   "seldata"
         DataSource      =   "datPrimaryRS"
         Enabled         =   0   'False
         Height          =   285
         Index           =   25
         Left            =   -70740
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   195
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CommandButton selmar 
         Caption         =   "Checar"
         Height          =   375
         Index           =   2
         Left            =   -73080
         TabIndex        =   194
         Top             =   1380
         Width           =   735
      End
      Begin VB.CommandButton ESCprofin 
         Height          =   315
         Index           =   0
         Left            =   -65820
         Picture         =   "Frmpcx.frx":5B24
         Style           =   1  'Graphical
         TabIndex        =   193
         TabStop         =   0   'False
         Top             =   1020
         Width           =   375
      End
      Begin VB.TextBox TXTpc 
         DataField       =   "CODFISCAL"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   22
         Left            =   -67320
         MaxLength       =   50
         TabIndex        =   191
         Top             =   1020
         Width           =   1455
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Seleçao 100 %"
         DataField       =   "sel100"
         DataSource      =   "datPrimaryRS"
         Enabled         =   0   'False
         Height          =   255
         Left            =   -72240
         TabIndex        =   190
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CommandButton selmar 
         Caption         =   "DesMarcar"
         Height          =   375
         Index           =   1
         Left            =   -74100
         TabIndex        =   188
         Top             =   1380
         Width           =   975
      End
      Begin VB.CommandButton selmar 
         Caption         =   "Marcar"
         Height          =   375
         Index           =   0
         Left            =   -74880
         TabIndex        =   187
         Top             =   1380
         Width           =   735
      End
      Begin VB.CommandButton EditSeq 
         Caption         =   "Editar"
         Height          =   255
         Index           =   15
         Left            =   -66600
         TabIndex        =   181
         Top             =   1500
         Width           =   1335
      End
      Begin VB.CommandButton CmdNovaREV 
         Caption         =   "Nova Revisão"
         Height          =   315
         Index           =   15
         Left            =   -72900
         TabIndex        =   180
         Top             =   1020
         Width           =   1215
      End
      Begin VB.CommandButton escTIPPC 
         Height          =   315
         Index           =   15
         Left            =   -69060
         Picture         =   "Frmpcx.frx":60AE
         Style           =   1  'Graphical
         TabIndex        =   179
         Top             =   1080
         Width           =   435
      End
      Begin VB.TextBox TXTpre 
         BackColor       =   &H00C0FFFF&
         DataField       =   "PCCTIPO"
         DataSource      =   "DataPreliminar"
         Enabled         =   0   'False
         Height          =   285
         Index           =   9
         Left            =   -69360
         Locked          =   -1  'True
         TabIndex        =   178
         TabStop         =   0   'False
         Top             =   1080
         Width           =   255
      End
      Begin VB.TextBox TXTpre 
         BackColor       =   &H00C0FFFF&
         DataField       =   "PCCREV"
         DataSource      =   "DataPreliminar"
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   -74040
         Locked          =   -1  'True
         TabIndex        =   177
         TabStop         =   0   'False
         Top             =   1020
         Width           =   1095
      End
      Begin VB.CommandButton CmdMotRev 
         Caption         =   "Motivos"
         Height          =   255
         Index           =   15
         Left            =   -66480
         TabIndex        =   176
         Top             =   2880
         Width           =   1335
      End
      Begin VB.CommandButton ComMotInc 
         Caption         =   "Incluir"
         Height          =   375
         Index           =   15
         Left            =   -66480
         TabIndex        =   175
         Top             =   4200
         Width           =   1335
      End
      Begin VB.CommandButton ComMotExc 
         Caption         =   "Excluir"
         Height          =   375
         Index           =   15
         Left            =   -66480
         TabIndex        =   174
         Top             =   4560
         Width           =   1335
      End
      Begin VB.CommandButton ComMotAlt 
         Caption         =   "Alterar"
         Height          =   375
         Index           =   15
         Left            =   -66480
         TabIndex        =   173
         Top             =   3840
         Width           =   1335
      End
      Begin VB.CommandButton EditSeq 
         Caption         =   "Editar"
         Height          =   255
         Index           =   14
         Left            =   -66540
         TabIndex        =   167
         Top             =   1860
         Width           =   1335
      End
      Begin VB.CommandButton CmdNovaREV 
         Caption         =   "Nova Revisão"
         Height          =   375
         Index           =   14
         Left            =   -72720
         TabIndex        =   166
         Top             =   1020
         Width           =   1215
      End
      Begin VB.CommandButton escTIPPC 
         Height          =   315
         Index           =   14
         Left            =   -68880
         Picture         =   "Frmpcx.frx":6638
         Style           =   1  'Graphical
         TabIndex        =   165
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox TXTpre 
         BackColor       =   &H00C0FFFF&
         DataField       =   "PCFTIPO"
         DataSource      =   "DataPreliminar"
         Enabled         =   0   'False
         Height          =   285
         Index           =   8
         Left            =   -69180
         Locked          =   -1  'True
         TabIndex        =   164
         TabStop         =   0   'False
         Top             =   1080
         Width           =   255
      End
      Begin VB.TextBox TXTpre 
         BackColor       =   &H00C0FFFF&
         DataField       =   "PCFREV"
         DataSource      =   "DataPreliminar"
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   -73920
         Locked          =   -1  'True
         TabIndex        =   163
         TabStop         =   0   'False
         Top             =   1020
         Width           =   1095
      End
      Begin VB.CommandButton CmdMotRev 
         Caption         =   "Motivos"
         Height          =   255
         Index           =   14
         Left            =   -66420
         TabIndex        =   162
         Top             =   3000
         Width           =   1335
      End
      Begin VB.CommandButton ComMotInc 
         Caption         =   "Incluir"
         Height          =   375
         Index           =   14
         Left            =   -66420
         TabIndex        =   161
         Top             =   4380
         Width           =   1335
      End
      Begin VB.CommandButton ComMotExc 
         Caption         =   "Excluir"
         Height          =   375
         Index           =   14
         Left            =   -66420
         TabIndex        =   160
         Top             =   4740
         Width           =   1335
      End
      Begin VB.CommandButton ComMotAlt 
         Caption         =   "Alterar"
         Height          =   375
         Index           =   14
         Left            =   -66420
         TabIndex        =   159
         Top             =   4020
         Width           =   1335
      End
      Begin VB.CommandButton EditSeq 
         Caption         =   "Editar"
         Height          =   255
         Index           =   13
         Left            =   -66360
         TabIndex        =   153
         Top             =   1500
         Width           =   1335
      End
      Begin VB.CommandButton CmdNovaREV 
         Caption         =   "Nova Revisão"
         Height          =   375
         Index           =   13
         Left            =   -71280
         TabIndex        =   152
         Top             =   1020
         Width           =   1335
      End
      Begin VB.CommandButton escTIPPC 
         Height          =   315
         Index           =   13
         Left            =   -69000
         Picture         =   "Frmpcx.frx":6BC2
         Style           =   1  'Graphical
         TabIndex        =   151
         Top             =   1020
         Width           =   495
      End
      Begin VB.TextBox TXTpre 
         BackColor       =   &H00C0FFFF&
         DataField       =   "PCPTIPO"
         DataSource      =   "DataPreliminar"
         Enabled         =   0   'False
         Height          =   285
         Index           =   7
         Left            =   -69360
         Locked          =   -1  'True
         TabIndex        =   150
         TabStop         =   0   'False
         Top             =   1020
         Width           =   255
      End
      Begin VB.TextBox TXTpre 
         BackColor       =   &H00C0FFFF&
         DataField       =   "PCPREV"
         DataSource      =   "DataPreliminar"
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   -74040
         Locked          =   -1  'True
         TabIndex        =   149
         TabStop         =   0   'False
         Top             =   1020
         Width           =   1095
      End
      Begin VB.CommandButton CmdMotRev 
         Caption         =   "Motivos"
         Height          =   255
         Index           =   13
         Left            =   -66360
         TabIndex        =   148
         Top             =   2760
         Width           =   1335
      End
      Begin VB.CommandButton ComMotInc 
         Caption         =   "Incluir"
         Height          =   375
         Index           =   13
         Left            =   -66300
         TabIndex        =   147
         Top             =   4080
         Width           =   1335
      End
      Begin VB.CommandButton ComMotExc 
         Caption         =   "Excluir"
         Height          =   375
         Index           =   13
         Left            =   -66300
         TabIndex        =   146
         Top             =   4440
         Width           =   1335
      End
      Begin VB.CommandButton ComMotAlt 
         Caption         =   "Alterar"
         Height          =   375
         Index           =   13
         Left            =   -66300
         TabIndex        =   145
         Top             =   3720
         Width           =   1335
      End
      Begin VB.CommandButton EditSeq 
         Caption         =   "Editar"
         Height          =   255
         Index           =   12
         Left            =   -66420
         TabIndex        =   139
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CommandButton CmdNovaREV 
         Caption         =   "Nova Revisão"
         Height          =   375
         Index           =   12
         Left            =   -71220
         TabIndex        =   138
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton escTIPPC 
         Height          =   315
         Index           =   12
         Left            =   -68880
         Picture         =   "Frmpcx.frx":714C
         Style           =   1  'Graphical
         TabIndex        =   137
         Top             =   1140
         Width           =   435
      End
      Begin VB.TextBox TXTpre 
         BackColor       =   &H00C0FFFF&
         DataField       =   "PCLTIPO"
         DataSource      =   "DataPreliminar"
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   -69300
         Locked          =   -1  'True
         TabIndex        =   136
         TabStop         =   0   'False
         Top             =   1140
         Width           =   255
      End
      Begin VB.TextBox TXTpre 
         BackColor       =   &H00C0FFFF&
         DataField       =   "PCLREV"
         DataSource      =   "DataPreliminar"
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   -74100
         Locked          =   -1  'True
         TabIndex        =   135
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton CmdMotRev 
         Caption         =   "Motivos"
         Height          =   255
         Index           =   12
         Left            =   -66360
         TabIndex        =   134
         Top             =   2640
         Width           =   1335
      End
      Begin VB.CommandButton ComMotInc 
         Caption         =   "Incluir"
         Height          =   375
         Index           =   12
         Left            =   -66360
         TabIndex        =   133
         Top             =   3960
         Width           =   1335
      End
      Begin VB.CommandButton ComMotExc 
         Caption         =   "Excluir"
         Height          =   375
         Index           =   12
         Left            =   -66360
         TabIndex        =   132
         Top             =   4320
         Width           =   1335
      End
      Begin VB.CommandButton ComMotAlt 
         Caption         =   "Alterar"
         Height          =   375
         Index           =   12
         Left            =   -66360
         TabIndex        =   131
         Top             =   3600
         Width           =   1335
      End
      Begin VB.CommandButton EditSeq 
         Caption         =   "Editar"
         Height          =   255
         Index           =   11
         Left            =   -66480
         TabIndex        =   126
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton CmdNovaREV 
         Caption         =   "Nova Revisão"
         Height          =   255
         Index           =   11
         Left            =   -72960
         TabIndex        =   125
         Top             =   1020
         Width           =   1335
      End
      Begin VB.CommandButton escTIPPC 
         Height          =   315
         Index           =   11
         Left            =   -69060
         Picture         =   "Frmpcx.frx":76D6
         Style           =   1  'Graphical
         TabIndex        =   124
         Top             =   1020
         Width           =   435
      End
      Begin VB.TextBox TXTpre 
         BackColor       =   &H00C0FFFF&
         DataField       =   "PCRTIPO"
         DataSource      =   "DataPreliminar"
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   -69420
         Locked          =   -1  'True
         TabIndex        =   123
         TabStop         =   0   'False
         Top             =   1020
         Width           =   255
      End
      Begin VB.TextBox TXTpre 
         BackColor       =   &H00C0FFFF&
         DataField       =   "PCRREV"
         DataSource      =   "DataPreliminar"
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   -74100
         Locked          =   -1  'True
         TabIndex        =   122
         TabStop         =   0   'False
         Top             =   1020
         Width           =   1095
      End
      Begin VB.CommandButton CmdMotRev 
         Caption         =   "Motivos"
         Height          =   255
         Index           =   11
         Left            =   -66360
         TabIndex        =   121
         Top             =   2640
         Width           =   1335
      End
      Begin VB.CommandButton ComMotInc 
         Caption         =   "Incluir"
         Height          =   375
         Index           =   11
         Left            =   -66420
         TabIndex        =   120
         Top             =   3900
         Width           =   1335
      End
      Begin VB.CommandButton ComMotExc 
         Caption         =   "Excluir"
         Height          =   375
         Index           =   11
         Left            =   -66420
         TabIndex        =   119
         Top             =   4320
         Width           =   1335
      End
      Begin VB.CommandButton ComMotAlt 
         Caption         =   "Alterar"
         Height          =   375
         Index           =   11
         Left            =   -66420
         TabIndex        =   118
         Top             =   3540
         Width           =   1335
      End
      Begin VB.TextBox TXTpre 
         BackColor       =   &H00C0FFFF&
         DataField       =   "PCELANOM"
         DataSource      =   "DataPreliminar"
         Enabled         =   0   'False
         Height          =   285
         Index           =   13
         Left            =   -71880
         Locked          =   -1  'True
         TabIndex        =   114
         TabStop         =   0   'False
         Top             =   1080
         Width           =   4215
      End
      Begin VB.CommandButton escpfela 
         Height          =   315
         Index           =   1
         Left            =   -72420
         Picture         =   "Frmpcx.frx":7C60
         Style           =   1  'Graphical
         TabIndex        =   113
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox TXTpre 
         BackColor       =   &H00C0FFFF&
         DataField       =   "PCELANUM"
         DataSource      =   "DataPreliminar"
         Enabled         =   0   'False
         Height          =   285
         Index           =   12
         Left            =   -73680
         Locked          =   -1  'True
         TabIndex        =   112
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox TXTpre 
         DataField       =   "PCGRUPO"
         DataSource      =   "DataPreliminar"
         Height          =   585
         Index           =   11
         Left            =   -74160
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   111
         Top             =   1980
         Width           =   8895
      End
      Begin VB.TextBox TXTpre 
         BackColor       =   &H00C0FFFF&
         DataField       =   "FEMEAG"
         DataSource      =   "datPrimaryRS"
         Enabled         =   0   'False
         Height          =   525
         Index           =   14
         Left            =   -74160
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   110
         TabStop         =   0   'False
         Top             =   2580
         Width           =   8895
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Grupo Femea"
         Height          =   435
         Index           =   1
         Left            =   -74880
         TabIndex        =   109
         Top             =   2580
         Width           =   735
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Bloqueado"
         DataField       =   "PCBLOQ"
         DataSource      =   "datPrimaryRS"
         Height          =   255
         Left            =   -67320
         TabIndex        =   107
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Grupo Femea"
         Height          =   435
         Index           =   0
         Left            =   -74940
         TabIndex        =   106
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox TXTpc 
         BackColor       =   &H00C0FFFF&
         DataField       =   "FEMEAG"
         DataSource      =   "datPrimaryRS"
         Enabled         =   0   'False
         Height          =   405
         Index           =   14
         Left            =   -74160
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   105
         TabStop         =   0   'False
         Top             =   2400
         Width           =   8895
      End
      Begin VB.CommandButton ComMotAlt 
         Caption         =   "Motivo"
         Height          =   315
         Index           =   7
         Left            =   -65940
         TabIndex        =   102
         Top             =   2640
         Width           =   1035
      End
      Begin VB.CommandButton ComMotAlt 
         Caption         =   "Motivo"
         Height          =   315
         Index           =   6
         Left            =   -66000
         TabIndex        =   101
         Top             =   2280
         Width           =   915
      End
      Begin VB.CommandButton ComMotAlt 
         Caption         =   "Alterar"
         Height          =   375
         Index           =   5
         Left            =   -66480
         TabIndex        =   98
         Top             =   3660
         Width           =   1335
      End
      Begin VB.CommandButton ComMotExc 
         Caption         =   "Excluir"
         Height          =   375
         Index           =   5
         Left            =   -66480
         TabIndex        =   97
         Top             =   4380
         Width           =   1335
      End
      Begin VB.CommandButton ComMotInc 
         Caption         =   "Incluir"
         Height          =   375
         Index           =   5
         Left            =   -66480
         TabIndex        =   96
         Top             =   4020
         Width           =   1335
      End
      Begin VB.CommandButton ComMotAlt 
         Caption         =   "Alterar"
         Height          =   375
         Index           =   4
         Left            =   -66540
         TabIndex        =   94
         Top             =   4140
         Width           =   1335
      End
      Begin VB.CommandButton ComMotExc 
         Caption         =   "Excluir"
         Height          =   375
         Index           =   4
         Left            =   -66540
         TabIndex        =   93
         Top             =   4860
         Width           =   1335
      End
      Begin VB.CommandButton ComMotInc 
         Caption         =   "Incluir"
         Height          =   375
         Index           =   4
         Left            =   -66540
         TabIndex        =   92
         Top             =   4500
         Width           =   1335
      End
      Begin VB.CommandButton ComMotAlt 
         Caption         =   "Alterar"
         Height          =   375
         Index           =   3
         Left            =   -66480
         TabIndex        =   90
         Top             =   4920
         Width           =   1335
      End
      Begin VB.CommandButton ComMotExc 
         Caption         =   "Excluir"
         Height          =   375
         Index           =   3
         Left            =   -66480
         TabIndex        =   89
         Top             =   5640
         Width           =   1335
      End
      Begin VB.CommandButton ComMotInc 
         Caption         =   "Incluir"
         Height          =   375
         Index           =   3
         Left            =   -66480
         TabIndex        =   88
         Top             =   5280
         Width           =   1335
      End
      Begin VB.CommandButton ComMotAlt 
         Caption         =   "Alterar"
         Height          =   375
         Index           =   2
         Left            =   -66360
         TabIndex        =   86
         Top             =   3600
         Width           =   1335
      End
      Begin VB.CommandButton ComMotExc 
         Caption         =   "Excluir"
         Height          =   375
         Index           =   2
         Left            =   -66360
         TabIndex        =   85
         Top             =   4440
         Width           =   1335
      End
      Begin VB.CommandButton ComMotInc 
         Caption         =   "Incluir"
         Height          =   375
         Index           =   2
         Left            =   -66360
         TabIndex        =   84
         Top             =   4020
         Width           =   1335
      End
      Begin VB.CommandButton ComMotAlt 
         Caption         =   "Alterar"
         Height          =   375
         Index           =   1
         Left            =   -66420
         TabIndex        =   82
         Top             =   3660
         Width           =   1335
      End
      Begin VB.CommandButton ComMotExc 
         Caption         =   "Excluir"
         Height          =   375
         Index           =   1
         Left            =   -66420
         TabIndex        =   81
         Top             =   4380
         Width           =   1335
      End
      Begin VB.CommandButton ComMotInc 
         Caption         =   "Incluir"
         Height          =   375
         Index           =   1
         Left            =   -66420
         TabIndex        =   80
         Top             =   4020
         Width           =   1335
      End
      Begin VB.CommandButton CmdMotRev 
         Caption         =   "Motivos"
         Height          =   255
         Index           =   5
         Left            =   -66480
         TabIndex        =   79
         Top             =   2640
         Width           =   1335
      End
      Begin VB.CommandButton CmdMotRev 
         Caption         =   "Motivos"
         Height          =   255
         Index           =   4
         Left            =   -66540
         TabIndex        =   78
         Top             =   3360
         Width           =   1335
      End
      Begin VB.CommandButton CmdMotRev 
         Caption         =   "Motivos"
         Height          =   255
         Index           =   3
         Left            =   -66480
         TabIndex        =   77
         Top             =   3960
         Width           =   1335
      End
      Begin VB.CommandButton CmdMotRev 
         Caption         =   "Motivos"
         Height          =   255
         Index           =   2
         Left            =   -66420
         TabIndex        =   76
         Top             =   2880
         Width           =   1335
      End
      Begin VB.CommandButton CmdMotRev 
         Caption         =   "Motivos"
         Height          =   255
         Index           =   1
         Left            =   -66420
         TabIndex        =   75
         Top             =   2940
         Width           =   1335
      End
      Begin VB.TextBox TXTpc 
         BackColor       =   &H00C0FFFF&
         DataField       =   "PCCREV"
         DataSource      =   "datPrimaryRS"
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   -74040
         Locked          =   -1  'True
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   1020
         Width           =   1095
      End
      Begin VB.TextBox TXTpc 
         BackColor       =   &H00C0FFFF&
         DataField       =   "PCCTIPO"
         DataSource      =   "datPrimaryRS"
         Enabled         =   0   'False
         Height          =   285
         Index           =   9
         Left            =   -69420
         Locked          =   -1  'True
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   1080
         Width           =   255
      End
      Begin VB.CommandButton escTIPPC 
         Height          =   315
         Index           =   5
         Left            =   -69120
         Picture         =   "Frmpcx.frx":81EA
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   1080
         Width           =   615
      End
      Begin VB.CommandButton CmdNovaREV 
         Caption         =   "Nova Revisão"
         Height          =   375
         Index           =   5
         Left            =   -72900
         TabIndex        =   63
         Top             =   1020
         Width           =   1215
      End
      Begin VB.TextBox TXTpc 
         BackColor       =   &H00C0FFFF&
         DataField       =   "PCFREV"
         DataSource      =   "datPrimaryRS"
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   -73980
         Locked          =   -1  'True
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   1020
         Width           =   1095
      End
      Begin VB.TextBox TXTpc 
         BackColor       =   &H00C0FFFF&
         DataField       =   "PCFTIPO"
         DataSource      =   "datPrimaryRS"
         Enabled         =   0   'False
         Height          =   285
         Index           =   8
         Left            =   -69300
         Locked          =   -1  'True
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   1020
         Width           =   255
      End
      Begin VB.CommandButton escTIPPC 
         Height          =   315
         Index           =   4
         Left            =   -69000
         Picture         =   "Frmpcx.frx":8774
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   1020
         Width           =   375
      End
      Begin VB.CommandButton CmdNovaREV 
         Caption         =   "Nova Revisão"
         Height          =   255
         Index           =   4
         Left            =   -72840
         TabIndex        =   56
         Top             =   1020
         Width           =   1215
      End
      Begin VB.TextBox TXTpc 
         BackColor       =   &H00C0FFFF&
         DataField       =   "PCPREV"
         DataSource      =   "datPrimaryRS"
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   -74040
         Locked          =   -1  'True
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox TXTpc 
         BackColor       =   &H00C0FFFF&
         DataField       =   "PCPTIPO"
         DataSource      =   "datPrimaryRS"
         Enabled         =   0   'False
         Height          =   285
         Index           =   7
         Left            =   -69600
         Locked          =   -1  'True
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   1080
         Width           =   255
      End
      Begin VB.CommandButton escTIPPC 
         Height          =   315
         Index           =   3
         Left            =   -69240
         Picture         =   "Frmpcx.frx":8CFE
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   1080
         Width           =   495
      End
      Begin VB.CommandButton CmdNovaREV 
         Caption         =   "Nova Revisão"
         Height          =   375
         Index           =   3
         Left            =   -71640
         TabIndex        =   50
         Top             =   1020
         Width           =   1335
      End
      Begin VB.TextBox TXTpc 
         BackColor       =   &H00C0FFFF&
         DataField       =   "PCLREV"
         DataSource      =   "datPrimaryRS"
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   -73800
         Locked          =   -1  'True
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   1020
         Width           =   1095
      End
      Begin VB.TextBox TXTpc 
         BackColor       =   &H00C0FFFF&
         DataField       =   "PCLTIPO"
         DataSource      =   "datPrimaryRS"
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   -68760
         Locked          =   -1  'True
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   1020
         Width           =   255
      End
      Begin VB.CommandButton CmdNovaREV 
         Caption         =   "Nova Revisão"
         Height          =   375
         Index           =   2
         Left            =   -70800
         TabIndex        =   45
         Top             =   1020
         Width           =   1575
      End
      Begin VB.TextBox TXTpc 
         BackColor       =   &H00C0FFFF&
         DataField       =   "PCRREV"
         DataSource      =   "datPrimaryRS"
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   -74100
         Locked          =   -1  'True
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   1020
         Width           =   1095
      End
      Begin VB.TextBox TXTpc 
         BackColor       =   &H00C0FFFF&
         DataField       =   "PCRTIPO"
         DataSource      =   "datPrimaryRS"
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   -69300
         Locked          =   -1  'True
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   1080
         Width           =   255
      End
      Begin VB.CommandButton escTIPPC 
         Height          =   315
         Index           =   1
         Left            =   -68880
         Picture         =   "Frmpcx.frx":9288
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   1080
         Width           =   435
      End
      Begin VB.CommandButton CmdNovaREV 
         Caption         =   "Nova Revisão"
         Height          =   255
         Index           =   1
         Left            =   -72960
         TabIndex        =   39
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox TXTpc 
         DataField       =   "PCGRUPO"
         DataSource      =   "datPrimaryRS"
         Height          =   585
         Index           =   11
         Left            =   -74160
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   38
         Top             =   1800
         Width           =   8895
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Indicador Revisão Femea"
         DataField       =   "STFEPC"
         DataSource      =   "datPrimaryRS"
         Height          =   255
         Left            =   -69840
         TabIndex        =   37
         Top             =   1440
         Width           =   2295
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Indicador Revisão PF"
         Height          =   255
         Left            =   -71940
         TabIndex        =   36
         Top             =   1440
         Width           =   1995
      End
      Begin VB.CommandButton ComMotInc 
         Caption         =   "Incluir"
         Height          =   375
         Index           =   0
         Left            =   -66540
         TabIndex        =   34
         Top             =   4560
         Width           =   1335
      End
      Begin VB.CommandButton ComMotExc 
         Caption         =   "Excluir"
         Height          =   375
         Index           =   0
         Left            =   -66540
         TabIndex        =   33
         Top             =   4920
         Width           =   1335
      End
      Begin VB.CommandButton ComMotAlt 
         Caption         =   "Alterar"
         Height          =   375
         Index           =   0
         Left            =   -66540
         TabIndex        =   32
         Top             =   4200
         Width           =   1335
      End
      Begin VB.CommandButton CmdMotRev 
         Caption         =   "Motivos"
         Height          =   255
         Index           =   0
         Left            =   -66480
         TabIndex        =   30
         Top             =   3300
         Width           =   1335
      End
      Begin VB.CommandButton EditSeq 
         Caption         =   "Editar"
         Height          =   255
         Index           =   5
         Left            =   -66480
         TabIndex        =   29
         Top             =   1500
         Width           =   1335
      End
      Begin VB.CommandButton EditSeq 
         Caption         =   "Editar"
         Height          =   255
         Index           =   4
         Left            =   -66540
         TabIndex        =   26
         Top             =   2220
         Width           =   1335
      End
      Begin VB.CommandButton EditSeq 
         Caption         =   "Editar"
         Height          =   255
         Index           =   2
         Left            =   -66480
         TabIndex        =   24
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CommandButton EditSeq 
         Caption         =   "Editar"
         Height          =   255
         Index           =   1
         Left            =   -66480
         TabIndex        =   22
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton EditSeq 
         Caption         =   "Editar"
         Height          =   255
         Index           =   3
         Left            =   -66480
         TabIndex        =   20
         Top             =   1500
         Width           =   1335
      End
      Begin VB.TextBox TXTpc 
         BackColor       =   &H00C0FFFF&
         DataField       =   "PCELANUM"
         DataSource      =   "datPrimaryRS"
         Enabled         =   0   'False
         Height          =   285
         Index           =   10
         Left            =   -67560
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1020
         Width           =   1095
      End
      Begin VB.CommandButton escpfela 
         Height          =   315
         Index           =   0
         Left            =   -72480
         Picture         =   "Frmpcx.frx":9812
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1020
         Width           =   315
      End
      Begin VB.TextBox TXTpc 
         BackColor       =   &H00C0FFFF&
         DataField       =   "PCELANOM"
         DataSource      =   "datPrimaryRS"
         Enabled         =   0   'False
         Height          =   285
         Index           =   13
         Left            =   -71880
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1020
         Width           =   4215
      End
      Begin MSFlexGridLib.MSFlexGrid GridSeq 
         Height          =   2115
         Index           =   0
         Left            =   -74880
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1500
         Width           =   8355
         _ExtentX        =   14737
         _ExtentY        =   3731
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid GRIDMAT 
         Height          =   1095
         Index           =   0
         Left            =   -74820
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   1440
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   1931
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid gridlab 
         Height          =   1035
         Index           =   0
         Left            =   -74880
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   1500
         Width           =   8355
         _ExtentX        =   14737
         _ExtentY        =   1826
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid GRIDFIM 
         Height          =   1275
         Index           =   0
         Left            =   -74880
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1800
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   2249
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid GRIDCOM 
         Height          =   1155
         Index           =   0
         Left            =   -74880
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   1440
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   2037
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid gridrevi 
         Height          =   1335
         Index           =   0
         Left            =   -74760
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   4020
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   2355
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid gridrev 
         Height          =   1095
         Index           =   1
         Left            =   -74820
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   2580
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   1931
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid gridrev 
         Height          =   1095
         Index           =   2
         Left            =   -74880
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   2520
         Width           =   8355
         _ExtentX        =   14737
         _ExtentY        =   1931
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid gridrev 
         Height          =   1215
         Index           =   3
         Left            =   -74880
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   3600
         Width           =   8355
         _ExtentX        =   14737
         _ExtentY        =   2143
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid gridrev 
         Height          =   1095
         Index           =   4
         Left            =   -74880
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   3060
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   1931
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid gridrev 
         Height          =   1095
         Index           =   5
         Left            =   -74880
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   2580
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   1931
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid gridrev 
         Height          =   1095
         Index           =   0
         Left            =   -74760
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   2940
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   1931
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid gridrevi 
         Height          =   1335
         Index           =   1
         Left            =   -74820
         TabIndex        =   83
         TabStop         =   0   'False
         Top             =   3660
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   2355
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid gridrevi 
         Height          =   1335
         Index           =   2
         Left            =   -74880
         TabIndex        =   87
         TabStop         =   0   'False
         Top             =   3600
         Width           =   8355
         _ExtentX        =   14737
         _ExtentY        =   2355
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid gridrevi 
         Height          =   1215
         Index           =   3
         Left            =   -74880
         TabIndex        =   91
         TabStop         =   0   'False
         Top             =   4800
         Width           =   8355
         _ExtentX        =   14737
         _ExtentY        =   2143
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid gridrevi 
         Height          =   1215
         Index           =   4
         Left            =   -74880
         TabIndex        =   95
         TabStop         =   0   'False
         Top             =   4140
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   2143
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid gridrevi 
         Height          =   1215
         Index           =   5
         Left            =   -74880
         TabIndex        =   99
         TabStop         =   0   'False
         Top             =   3660
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   2143
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid gridrevi 
         Height          =   2055
         Index           =   6
         Left            =   -74880
         TabIndex        =   100
         TabStop         =   0   'False
         Top             =   1020
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   3625
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid gridrevi 
         Height          =   3255
         Index           =   7
         Left            =   -74880
         TabIndex        =   103
         TabStop         =   0   'False
         Top             =   1020
         Width           =   8835
         _ExtentX        =   15584
         _ExtentY        =   5741
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid GRIDMAT 
         Height          =   1095
         Index           =   1
         Left            =   -74820
         TabIndex        =   127
         TabStop         =   0   'False
         Top             =   1380
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   1931
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid gridrev 
         Height          =   1095
         Index           =   11
         Left            =   -74820
         TabIndex        =   128
         TabStop         =   0   'False
         Top             =   2460
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   1931
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid gridrevi 
         Height          =   1335
         Index           =   11
         Left            =   -74820
         TabIndex        =   129
         TabStop         =   0   'False
         Top             =   3540
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   2355
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid gridlab 
         Height          =   1035
         Index           =   1
         Left            =   -74880
         TabIndex        =   140
         TabStop         =   0   'False
         Top             =   1500
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   1826
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid gridrev 
         Height          =   1095
         Index           =   9
         Left            =   -74880
         TabIndex        =   141
         TabStop         =   0   'False
         Top             =   2520
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   1931
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid gridrevi 
         Height          =   1215
         Index           =   9
         Left            =   -74880
         TabIndex        =   142
         TabStop         =   0   'False
         Top             =   3600
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   2143
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid GridSeq 
         Height          =   1095
         Index           =   1
         Left            =   -74880
         TabIndex        =   154
         TabStop         =   0   'False
         Top             =   1440
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   1931
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid gridrev 
         Height          =   1095
         Index           =   13
         Left            =   -74880
         TabIndex        =   155
         TabStop         =   0   'False
         Top             =   2580
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   1931
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid gridrevi 
         Height          =   1275
         Index           =   13
         Left            =   -74880
         TabIndex        =   156
         TabStop         =   0   'False
         Top             =   3660
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   2249
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid GRIDFIM 
         Height          =   1515
         Index           =   1
         Left            =   -74880
         TabIndex        =   168
         TabStop         =   0   'False
         Top             =   1440
         Width           =   8355
         _ExtentX        =   14737
         _ExtentY        =   2672
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid gridrev 
         Height          =   1095
         Index           =   14
         Left            =   -74880
         TabIndex        =   169
         TabStop         =   0   'False
         Top             =   2940
         Width           =   8355
         _ExtentX        =   14737
         _ExtentY        =   1931
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid gridrevi 
         Height          =   1155
         Index           =   14
         Left            =   -74880
         TabIndex        =   170
         TabStop         =   0   'False
         Top             =   4020
         Width           =   8355
         _ExtentX        =   14737
         _ExtentY        =   2037
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid GRIDCOM 
         Height          =   1155
         Index           =   1
         Left            =   -74880
         TabIndex        =   182
         TabStop         =   0   'False
         Top             =   1500
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   2037
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid gridrev 
         Height          =   1095
         Index           =   15
         Left            =   -74880
         TabIndex        =   183
         TabStop         =   0   'False
         Top             =   2700
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   1931
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid gridrevi 
         Height          =   1215
         Index           =   15
         Left            =   -74880
         TabIndex        =   184
         TabStop         =   0   'False
         Top             =   3780
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   2143
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid GRIDFIM 
         Height          =   1515
         Index           =   2
         Left            =   -74820
         TabIndex        =   237
         TabStop         =   0   'False
         Top             =   1440
         Width           =   8355
         _ExtentX        =   14737
         _ExtentY        =   2672
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid gridrev 
         Height          =   1095
         Index           =   24
         Left            =   -74820
         TabIndex        =   238
         TabStop         =   0   'False
         Top             =   2940
         Width           =   8355
         _ExtentX        =   14737
         _ExtentY        =   1931
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid gridrevi 
         Height          =   1155
         Index           =   24
         Left            =   -74820
         TabIndex        =   239
         TabStop         =   0   'False
         Top             =   4020
         Width           =   8355
         _ExtentX        =   14737
         _ExtentY        =   2037
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid GridSeq 
         Height          =   1095
         Index           =   2
         Left            =   -74820
         TabIndex        =   252
         TabStop         =   0   'False
         Top             =   1440
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   1931
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid gridrev 
         Height          =   1095
         Index           =   23
         Left            =   -74820
         TabIndex        =   253
         TabStop         =   0   'False
         Top             =   2580
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   1931
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid gridrevi 
         Height          =   1275
         Index           =   23
         Left            =   -74820
         TabIndex        =   254
         TabStop         =   0   'False
         Top             =   3660
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   2249
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid GRIDCOM 
         Height          =   1155
         Index           =   2
         Left            =   -74880
         TabIndex        =   279
         TabStop         =   0   'False
         Top             =   1560
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   2037
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid gridrev 
         Height          =   1095
         Index           =   25
         Left            =   -74880
         TabIndex        =   280
         TabStop         =   0   'False
         Top             =   2760
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   1931
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid gridrevi 
         Height          =   1215
         Index           =   25
         Left            =   -74880
         TabIndex        =   281
         TabStop         =   0   'False
         Top             =   3840
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   2143
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid GRIDMAT 
         Height          =   1095
         Index           =   2
         Left            =   -74820
         TabIndex        =   294
         TabStop         =   0   'False
         Top             =   1380
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   1931
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid gridrev 
         Height          =   1095
         Index           =   21
         Left            =   -74820
         TabIndex        =   295
         TabStop         =   0   'False
         Top             =   2460
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   1931
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid gridrevi 
         Height          =   1335
         Index           =   21
         Left            =   -74820
         TabIndex        =   296
         TabStop         =   0   'False
         Top             =   3540
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   2355
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid gridlab 
         Height          =   1035
         Index           =   2
         Left            =   -74820
         TabIndex        =   309
         TabStop         =   0   'False
         Top             =   1560
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   1826
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid gridrev 
         Height          =   1095
         Index           =   22
         Left            =   -74820
         TabIndex        =   310
         TabStop         =   0   'False
         Top             =   2580
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   1931
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid gridrevi 
         Height          =   1215
         Index           =   22
         Left            =   -74820
         TabIndex        =   311
         TabStop         =   0   'False
         Top             =   3660
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   2143
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid GridPFrev 
         Height          =   1695
         Left            =   -74880
         TabIndex        =   349
         Top             =   3120
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   2990
         _Version        =   393216
      End
      Begin VB.Label lblLabels 
         Caption         =   "Inf Req.Cliente"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   20
         Left            =   -70200
         TabIndex        =   326
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Outra Aprovaçao Data"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   35
         Left            =   -74880
         TabIndex        =   324
         Top             =   1140
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Grupo"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   51
         Left            =   180
         TabIndex        =   322
         Top             =   2100
         Width           =   615
      End
      Begin VB.Label lblLabels 
         Caption         =   "Data Abertura"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   50
         Left            =   180
         TabIndex        =   321
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         Caption         =   "Elaborador"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   49
         Left            =   180
         TabIndex        =   320
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblLabels 
         Caption         =   "Revisão"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   48
         Left            =   -74820
         TabIndex        =   313
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblLabels 
         Caption         =   "Tipo"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   47
         Left            =   -69720
         TabIndex        =   312
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label lblLabels 
         Caption         =   "Tipo"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   46
         Left            =   -69840
         TabIndex        =   298
         Top             =   1020
         Width           =   375
      End
      Begin VB.Label lblLabels 
         Caption         =   "Revisão"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   45
         Left            =   -74880
         TabIndex        =   297
         Top             =   1020
         Width           =   735
      End
      Begin VB.Label lblLabels 
         Caption         =   "Revisão"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   44
         Left            =   -74880
         TabIndex        =   283
         Top             =   1080
         Width           =   675
      End
      Begin VB.Label lblLabels 
         Caption         =   "Tipo"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   43
         Left            =   -69840
         TabIndex        =   282
         Top             =   1140
         Width           =   375
      End
      Begin VB.Label lblLabels 
         Caption         =   "Revisão"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   42
         Left            =   -74760
         TabIndex        =   256
         Top             =   1020
         Width           =   615
      End
      Begin VB.Label lblLabels 
         Caption         =   "Tipo"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   41
         Left            =   -69780
         TabIndex        =   255
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label lblLabels 
         Caption         =   "Revisão"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   39
         Left            =   -74820
         TabIndex        =   241
         Top             =   1020
         Width           =   855
      End
      Begin VB.Label lblLabels 
         Caption         =   "Tipo"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   38
         Left            =   -69660
         TabIndex        =   240
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label lblLabels 
         Caption         =   "Revisão"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   36
         Left            =   -74880
         TabIndex        =   203
         Top             =   1020
         Width           =   735
      End
      Begin VB.Label lblLabels 
         Caption         =   "Cod.Final Fatur."
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   34
         Left            =   -68580
         TabIndex        =   192
         Top             =   1020
         Width           =   1275
      End
      Begin VB.Label lblLabels 
         Caption         =   "Tipo"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   33
         Left            =   -69840
         TabIndex        =   186
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label lblLabels 
         Caption         =   "Revisão"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   32
         Left            =   -74880
         TabIndex        =   185
         Top             =   1020
         Width           =   675
      End
      Begin VB.Label lblLabels 
         Caption         =   "Tipo"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   31
         Left            =   -69720
         TabIndex        =   172
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label lblLabels 
         Caption         =   "Revisão"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   30
         Left            =   -74820
         TabIndex        =   171
         Top             =   1020
         Width           =   855
      End
      Begin VB.Label lblLabels 
         Caption         =   "Tipo"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   29
         Left            =   -69840
         TabIndex        =   158
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label lblLabels 
         Caption         =   "Revisão"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   28
         Left            =   -74820
         TabIndex        =   157
         Top             =   1020
         Width           =   615
      End
      Begin VB.Label lblLabels 
         Caption         =   "Tipo"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   27
         Left            =   -69780
         TabIndex        =   144
         Top             =   1140
         Width           =   375
      End
      Begin VB.Label lblLabels 
         Caption         =   "Revisão"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   26
         Left            =   -74880
         TabIndex        =   143
         Top             =   1020
         Width           =   735
      End
      Begin VB.Label lblLabels 
         Caption         =   "Tipo"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   25
         Left            =   -69840
         TabIndex        =   130
         Top             =   1020
         Width           =   375
      End
      Begin VB.Label lblLabels 
         Caption         =   "Elaborador"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   24
         Left            =   -74820
         TabIndex        =   117
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblLabels 
         Caption         =   "Data Abertura"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   23
         Left            =   -74820
         TabIndex        =   116
         Top             =   1500
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         Caption         =   "Grupo"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   22
         Left            =   -74820
         TabIndex        =   115
         Top             =   1980
         Width           =   615
      End
      Begin VB.Label lblLabels 
         Caption         =   "Revisão"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   17
         Left            =   -74880
         TabIndex        =   68
         Top             =   1020
         Width           =   795
      End
      Begin VB.Label lblLabels 
         Caption         =   "Tipo"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   7
         Left            =   -69900
         TabIndex        =   67
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label lblLabels 
         Caption         =   "Grupo"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   6
         Left            =   -74880
         TabIndex        =   62
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label lblLabels 
         Caption         =   "Revisão"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   16
         Left            =   -74880
         TabIndex        =   61
         Top             =   1020
         Width           =   855
      End
      Begin VB.Label lblLabels 
         Caption         =   "Tipo"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   15
         Left            =   -69780
         TabIndex        =   60
         Top             =   1020
         Width           =   375
      End
      Begin VB.Label lblLabels 
         Caption         =   "Revisão"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   14
         Left            =   -74880
         TabIndex        =   55
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label lblLabels 
         Caption         =   "Tipo"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   13
         Left            =   -70080
         TabIndex        =   54
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label lblLabels 
         Caption         =   "Revisão"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   12
         Left            =   -74880
         TabIndex        =   49
         Top             =   1020
         Width           =   975
      End
      Begin VB.Label lblLabels 
         Caption         =   "Tipo"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   11
         Left            =   -69180
         TabIndex        =   48
         Top             =   1020
         Width           =   375
      End
      Begin VB.Label lblLabels 
         Caption         =   "Revisão"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   10
         Left            =   -74880
         TabIndex        =   44
         Top             =   1020
         Width           =   795
      End
      Begin VB.Label lblLabels 
         Caption         =   "Tipo"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   9
         Left            =   -69720
         TabIndex        =   43
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label lblLabels 
         Caption         =   "Data Abertura"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   8
         Left            =   -74880
         TabIndex        =   35
         Top             =   1380
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         Caption         =   "Elaborador"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   0
         Left            =   -74880
         TabIndex        =   19
         Top             =   1020
         Width           =   975
      End
   End
   Begin VB.TextBox TXTPF 
      BackColor       =   &H00C0FFFF&
      DataField       =   "CODCLIENTE"
      DataSource      =   "datPrimaryRS"
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   8160
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   224
      TabStop         =   0   'False
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox TXTPF 
      BackColor       =   &H00C0FFFF&
      DataField       =   "CODIGOINT"
      DataSource      =   "datPrimaryRS"
      Enabled         =   0   'False
      Height          =   285
      Index           =   13
      Left            =   6480
      Locked          =   -1  'True
      MaxLength       =   24
      TabIndex        =   218
      TabStop         =   0   'False
      Top             =   960
      Width           =   2415
   End
   Begin VB.TextBox TXTPF 
      BackColor       =   &H00C0FFFF&
      DataField       =   "CODFINAL"
      DataSource      =   "datPrimaryRS"
      Enabled         =   0   'False
      Height          =   285
      Index           =   12
      Left            =   9000
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   189
      TabStop         =   0   'False
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox TXTpc 
      DataField       =   "FILIAL"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   24
      Left            =   9720
      TabIndex        =   5
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox TXTPF 
      BackColor       =   &H00C0FFFF&
      DataField       =   "CLINOME"
      DataSource      =   "datPrimaryRS"
      Enabled         =   0   'False
      Height          =   285
      Index           =   5
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   360
      Width           =   4335
   End
   Begin VB.TextBox TXTPF 
      BackColor       =   &H00C0FFFF&
      DataField       =   "OPCAO"
      DataSource      =   "datPrimaryRS"
      Enabled         =   0   'False
      Height          =   285
      Index           =   14
      Left            =   1920
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox TXTPF 
      BackColor       =   &H00C0FFFF&
      DataField       =   "CLIENTE"
      DataSource      =   "datPrimaryRS"
      Enabled         =   0   'False
      Height          =   285
      Index           =   4
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox TXTPF 
      BackColor       =   &H00C0FFFF&
      DataField       =   "TXPF"
      DataSource      =   "datPrimaryRS"
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   2640
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   7
      Top             =   960
      Width           =   3735
   End
   Begin VB.TextBox TXTPF 
      BackColor       =   &H00C0FFFF&
      DataField       =   "CODIGO"
      DataSource      =   "datPrimaryRS"
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   120
      Locked          =   -1  'True
      MaxLength       =   24
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   960
      Width           =   2415
   End
   Begin VB.TextBox TXTPF 
      BackColor       =   &H00C0FFFF&
      DataField       =   "CPF"
      DataSource      =   "datPrimaryRS"
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   1320
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox TXTNPF 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   360
      Width           =   1095
   End
   Begin XPControls.XPButton Encerrar 
      Height          =   435
      Left            =   9480
      TabIndex        =   400
      Top             =   2040
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "Frmpcx.frx":9D9C
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
      Left            =   9480
      TabIndex        =   401
      Top             =   1440
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "Frmpcx.frx":A336
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
   Begin VB.Label lblLabels 
      Caption         =   "Codigo Final PF"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   52
      Left            =   9000
      TabIndex        =   327
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lblLabels 
      Caption         =   "TIPO:  (C)ontenção Avancanda ( )Normal (I)tem Seguranca (X)Cont.Avancada/Item Seg. (G)P12 agressivo GP12 Plan(N)"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   37
      Left            =   120
      TabIndex        =   225
      Top             =   1320
      Width           =   9195
   End
   Begin VB.Label lblLabels 
      Caption         =   "Codigo Integrado"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   40
      Left            =   6480
      TabIndex        =   217
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label lblLabels 
      Caption         =   "Codigo Cliente"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   21
      Left            =   8220
      TabIndex        =   108
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblLabels 
      Caption         =   "Fabrica"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   18
      Left            =   9720
      TabIndex        =   104
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label11 
      Caption         =   "PF"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblLabels 
      Caption         =   "Opção"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   19
      Left            =   1920
      TabIndex        =   13
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblLabels 
      Caption         =   "Nome:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   5
      Left            =   3840
      TabIndex        =   12
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "CLIENTE:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   4
      Left            =   2520
      TabIndex        =   11
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "Descrição"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   3
      Left            =   2640
      TabIndex        =   10
      Top             =   720
      Width           =   1035
   End
   Begin VB.Label lblLabels 
      Caption         =   "CODIGO:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   9
      Top             =   705
      Width           =   915
   End
   Begin VB.Label lblLabels 
      Caption         =   "CPF:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   1
      Left            =   1320
      TabIndex        =   8
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmPCX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cARQPF As String
Dim cARQPFP As String
Dim cARQPFG As String
Dim aVAL As Variant
Dim aFOR As Variant
Dim aCAM As Variant
Dim aPAD As Variant
Dim iLOOP As Integer
Dim aVALPC As Variant
Dim aFORPC As Variant
Dim aCAMPC As Variant
Dim aPADPC As Variant
Dim aVALPRE As Variant
Dim aVALGP12 As Variant

Private Sub cmdClose_Click()
    On Error GoTo errhandler
    Dim sSQL As String
    Dim dFEMEA As Date
    Dim dPF As Date
    Dim dMAIOR As Date
    Dim dPC As Date
    Dim dPC2 As Date
    Dim cSQL As String

    'cSQL = "select * from PF WHERE PF=" & nPF
    If IsDate(TXTPF(18)) And IsDate(TXTPF(16)) Then
        dFEMEA = TXTPF(16)
        dPF = TXTPF(18)
        If IsDate(TXTPF(19)) Then
            dFEMEA = TXTPF(19)
        End If
        If dFEMEA < dPF Then
            Alert ("Data Revisao Femea Menor que Processo Fabricacao")
        End If
    End If
    If IsDate(TXTPF(18)) Or IsDate(TXTPF(16)) Then
        dMAIOR = TXTPF(18)
        If IsDate(TXTPF(16)) Then
            dFEMEA = TXTPF(16)
            If IsDate(TXTPF(19)) Then
                dFEMEA = TXTPF(19)
            End If
            If dFEMEA > dMAIOR Then
                dMAIOR = dFEMEA
            End If
        End If
        If IsDate(TXTpc(32)) Then
            dPC2 = TXTpc(32)
            If dPC2 > dMAIOR Then
                dMAIOR = dPC
            End If
        End If
        If IsDate(TXTpc(19)) Then
            dPC = TXTpc(19)
            If IsDate(dPC2) Then
                dPC = dPC2
            End If
            If dPC < dMAIOR Then
                Alert ("Data Revisao Plano Controle Componentes Menor que Processo Fabricacao OU Femea")
            End If
        End If
        If IsDate(TXTpc(18)) Then
            dPC = TXTpc(18)
            If IsDate(dPC2) Then
                dPC = dPC2
            End If
            If dPC < dMAIOR Then
                Alert ("Data Revisao Plano Controle Final Menor que Processo Fabricacao OU Femea")
            End If
        End If
        If IsDate(TXTpc(17)) Then
            dPC = TXTpc(17)
            If IsDate(dPC2) Then
                dPC = dPC2
            End If
            If dPC < dMAIOR Then
                Alert ("Data Revisao Plano Controle Producao Menor que Processo Fabricacao OU Femea")
            End If
        End If
        If IsDate(TXTpc(16)) Then
            dPC = TXTpc(16)
            If IsDate(dPC2) Then
                dPC = dPC2
            End If
            If dPC < dMAIOR Then
                Alert ("Data Revisao Plano Controle Laboratorio Menor que Processo Fabricacao OU Femea")
            End If
        End If
        If IsDate(TXTpc(15)) Then
            dPC = TXTpc(15)
            If IsDate(dPC2) Then
                dPC = dPC2
            End If
            If dPC < dMAIOR Then
                Alert ("Data Revisao Plano Controle Recebimento Menor que Processo Fabricacao OU Femea")
            End If
        End If
      
      
    End If
   
   
    ' Sempre grava por causa das revisoes
        
    'valpc
    For iLOOP = 0 To 19
        aVALPC(iLOOP) = TXTpc(iLOOP)
    Next iLOOP
    aVALPC(20) = PCDinipc.Value
    For iLOOP = 21 To 26
        aVALPC(iLOOP) = TXTpc(iLOOP)
    Next iLOOP
    aVALPC(27) = Check1.Value
    aVALPC(28) = Check2.Value
    aVALPC(29) = Check3.Value
    aVALPC(30) = Check3.Value
    aVALPC(31) = Check5.Value
    aVALPC(32) = TXTpc(32)
    aVALPC(33) = TXTpc(33)
    aVALPC(34) = TXTpc(34)
    
     cSQL = "select PF,PCRREV,PCLREV,PCPREV,PCFREV,PCCREV,PCRTIPO,PCLTIPO,PCPTIPO,PCFTIPO,PCCTIPO,"
     cSQL = cSQL & "PCELADAT,PCGRUPO,PCELANUM,PCELANOM,FEMEAG,PCRDAT,PCLDAT,PCPDAT,PCFDAT,PCCDAT,"
     cSQL = cSQL & "PCDINI,OUTRAAPR,CODFISCAL,CRITIOBS,FILIAL,SELDATA,CLIESP,STPFPC,STFEPC,PCBLOQ,"
     cSQL = cSQL & "SEL100,CRITICO,PCREVD2,CONTATONUM,CONTATONOM"
     cSQL = cSQL & " from PF WHERE PF=" & nPF
    
    
    
    GrvSQL cARQPF, cSQL, 35, aCAMPC, aVALPC, aFORPC
        
    'pre
    For iLOOP = 0 To 19
        aVALPRE(iLOOP) = TXTpre(iLOOP)
    Next iLOOP
    aVALPRE(20) = PCDinipre.Value

    GrvSQL cARQPFP, cSQL, 21, aCAMPC, aVALPRE, aFORPC
        
    'gp12
    For iLOOP = 0 To 19
        aVALGP12(iLOOP) = TXTgp12(iLOOP)
    Next iLOOP
    aVALGP12(20) = PCDinigp12.Value
    GrvSQL cARQPFG, cSQL, 21, aCAMPC, aVALGP12, aFORPC
         
      
    '  End If
   
   
    ATUPFPG

    Screen.MousePointer = vbDefault
    Unload Me
    Exit Sub
errhandler:
 
    Select Case Err.Number
    Case Else
        SayErro "frmPCX"
    End Select

End Sub

Private Sub Cmddistro_Click(Index As Integer)
    If nREV = 0 Then
        Alert "Revisão Nao Escolhida"
        Exit Sub
    End If
    cTIPO = CTIPOPLANO(Index)
    ePASS01 = "select * from rev WHERE PF=" & nPF & " AND TIPO='" & cTIPO & "' AND REVISAO=" & nREV
    FrmRev.Show vbModal, Me
End Sub

Private Sub CmdGrdRI_Click(Index As Integer)
    Dim cCODIGO As String
    Dim cCON As String
    Dim cSQL As String
    If Index = 0 Then
        cCODIGO = TXTPF(2)
    Else
        cCODIGO = TXTpc(22)
    End If
    If Len(cCODIGO) < 1 Then
        Alert ("Codigo em branco")
        Exit Sub
    End If
    cCON = PegPath("PATH", "MANA5QUA")
    cCON = GeracArq(cCON, "JETFOX")

    cSQL = "select ri,data,origem from ri where desenho='" & cCODIGO & "' order by RI"
    MontaGridUltra grdRI, 3, Array(800, 1200, 400), Array("RI", "Data", "Origem"), _
        Array("RI", "DATA", "origem"), cCON, cSQL

End Sub

Private Sub CmdMarcaPC_Click(Index As Integer)
    Dim cSQLTMP As String
    GridPFrev.Col = 3 + Index
    If Len(GridPFrev) > 0 Then
        Alert ("Ja Verificado")
        Exit Sub
    End If
    GridPFrev.Col = 1
    cSQLTMP = "select * from PFrev WHERE PF=" & nPF & " AND REV=" & FixInt(GridPFrev)
    Select Case Index
    Case 0
        GrvSQL cARQPF, cSQLTMP, 2, Array("PCRD", "PCRN"), Array(Date, zNOMEFOLHA), Array("D", "C")
    Case 1
        GrvSQL cARQPF, cSQLTMP, 2, Array("PCLD", "PCLN"), Array(Date, zNOMEFOLHA), Array("D", "C")
    Case 2
        GrvSQL cARQPF, cSQLTMP, 2, Array("PCFD", "PCFN"), Array(Date, zNOMEFOLHA), Array("D", "C")
    Case 3
        GrvSQL cARQPF, cSQLTMP, 2, Array("PCCD", "PCCN"), Array(Date, zNOMEFOLHA), Array("D", "C")
    Case 4
        GrvSQL cARQPF, cSQLTMP, 2, Array("PCPD", "PCPN"), Array(Date, zNOMEFOLHA), Array("D", "C")
    End Select
    FILGRIdpfrev
End Sub

Private Sub CmdMotRev_Click(Index As Integer)
    If gridrev(Index).Row > 0 Then
        gridrev(Index).Col = 1
        nREV = gridrev(Index)
        ComMotAlt(Index).Enabled = True
        ComMotExc(Index).Enabled = True
        ComMotInc(Index).Enabled = True
        Filgridrevi Index, nREV
    End If
End Sub

Private Sub CmdNovaREV_Click(Index As Integer)
    Dim nREV As Integer
    Dim cTIPO As String

    eRETU02 = Today()
    Select Case Index
    Case 1, 2, 3, 4, 5
        eRETU02 = TXTpc(Index + 15 - 1)
    Case 11, 12, 13, 14, 15
        eRETU02 = TXTpre(Index + 15 - 11)
    Case 21, 22, 23, 24, 25
        eRETU02 = TXTgp12(Index + 15 - 21)
    End Select


    'na importacao protheus algumas data vieram em branco
    If Len(eRETU02) > 0 Then
        If eRETU02 > Today() Then
            Alert ("Data Revisao Anterior Maior que Data do Sistema")
            Exit Sub
        End If
    End If
    
 
 
    'If eRETU02 > Today() Then
    '   Alert ("Data Revisao Anterior Maior que Data do Sistema")
    '   Exit Sub
    'End If
    
    
    If MDG("Fazer Nova Revisao", "Confirme Gravação") Then
        Encerrar.Visible = False
        eRETU01 = Date
        
        Select Case Index
        Case 1, 2, 3, 4, 5
            nREV = FixInt(TXTpc(Index - 1)) + 1
            TXTpc(Index - 1) = nREV
            TXTpc(Index + 15 - 1) = eRETU01
        Case 11, 12, 13, 14, 15
            nREV = FixInt(TXTpre(Index - 11)) + 1
            TXTpre(Index - 11) = nREV
            TXTpre(Index + 15 - 11) = eRETU01
        Case 21, 22, 23, 24, 25
            nREV = FixInt(TXTgp12(Index - 21)) + 1
            TXTgp12(Index - 21) = nREV
            TXTgp12(Index + 15 - 21) = eRETU01
        End Select

        
        
        cTIPO = CTIPOPLANO(Index)
        
        
        IncluiSQL cARQPF, "SELECT * FROM REV WHERE PF=" & nPF & "AND TIPO='" & cTIPO & "' AND REVISAO=" & nREV, 4, Array("PF", "REVISAO", "TIPO", "DATA"), _
        Array(nPF, nREV, cTIPO, eRETU01), True, False
                                 
        Filgridrev (Index)
        
        
        eRETU02 = "PF:" & nPF & " " & "Tipo Plano: " & cTIPO & Chr(13) & Chr(10)
        eRETU02 = eRETU02 & " Desenho:" & TXTPF(2) & Chr(13) & Chr(10)
        eRETU02 = eRETU02 & " Data:" & Fdata(Date) & Chr(13) & Chr(10)
        
        MAILENV "PC000001", eRETU02
           
        Select Case Index
        Case 1, 2, 3, 4, 5
            escpfela_Click (0)
        Case 11, 12, 13, 14, 15
            escpfela_Click (1)
        Case 21, 22, 23, 24, 25
            escpfela_Click (2)
        End Select
        
        cmdClose_Click
        
    End If

End Sub

Private Sub CmdNovoRI_Click(Index As Integer)
    Dim oRS As New ADODB.Recordset
    Dim oCONN As New ADODB.Connection
    Dim cARQ As String
    Dim cSQL As String
    Dim cData As Variant
    Dim nRIAREA As Variant
    Dim iRETVAL As Variant
    Dim nRI As Variant
    Dim nITEM As Variant
    Dim cCAM As String
    Dim nCLIENTE As Variant
    Dim cDESCRI As Variant
    Dim cDESENHO As Variant
    Dim cCLINOME As Variant
    On Error GoTo erro

    cData = Format(Date, "DD/MM/YY")
    cCAM = PegPath("PATH", "MANA5QUA")
    
    sx_SetEpoch (Year(Date) - 30)
    sx_SetDateFormat BRITISH
    sx_SetDeleted True
    sx_SetSoftSeek True
    
    nRIAREA = sx_Use(cCAM & "RI.Dbf", "RI", READWRITE, SDEFOX)
    iRETVAL = sx_SetOrder(1)
    sx_GoBottom
    nRI = FixInt(sx_GetVariant("RI"))
    nRI = nRI + 1
    cDESENHO = FixStr(TXTPF(2))
    cDESCRI = FixStr(TXTPF(3))
    nCLIENTE = FixInt(TXTPF(4))
    cCLINOME = FixStr(TXTPF(5))

    sx_AppendBlank
    If sx_Rlock(sx_RecNo()) Then
        sx_PutVariant "RI", nRI
        sx_PutVariant "DESENHO", cDESENHO
        sx_PutVariant "DESCRI", cDESCRI
        sx_PutVariant "CLIENTE", nCLIENTE
        sx_PutVariant "CLINOME", cCLINOME
        sx_PutVariant "DATA", cData
        sx_PutVariant "TIPOI", "P"
        sx_PutVariant "TIPO2", "D"
        If Index = 0 Then
            sx_PutVariant "ORIGEM", "PCP"
        Else
            sx_PutVariant "ORIGEM", "PPP"
        End If
        sx_Commit
    End If
    sx_Unlock sx_RecNo()
    sx_CloseAll

    nITEM = 1
    cARQ = cARQPF
    If Index = 1 Then
        cARQ = cARQPFP
    End If
    nRIAREA = sx_Use(cCAM & "RII.Dbf", "RII", READWRITE, SDEFOX)
    cARQ = GeracArq(cARQ)
    cSQL = "SELECT ESPE FROM PFC WHERE PF=" & nPF & " AND SAIRI ORDER BY SEQ,SSQ,ITEM"
    oCONN.ConnectionTimeout = 120
    oCONN.Open cARQ
    oRS.Open cSQL, oCONN, adOpenForwardOnly, adLockReadOnly
    While Not oRS.EOF
        sx_AppendBlank
        If sx_Rlock(sx_RecNo()) Then
            sx_PutVariant "RI", nRI
            sx_PutVariant "ITEM", nITEM
            sx_PutVariant "ESPE", FixStr(oRS("ESPE"))
            sx_Commit
        End If
        sx_Unlock sx_RecNo()
        nITEM = nITEM + 1
        oRS.MoveNext
    Wend
    sx_CloseAll
    oRS.Close
    oCONN.Close
    Alert "RI Gerado Nº" & nRI
    Exit Sub
erro:
    Select Case Err.Number
    Case Else
        SayErro
        Resume Next
    End Select
End Sub

Private Sub CmdTotMot_Click(Index As Integer)
    Filgridrev Index
    Filgridrevi Index, 0
    ComMotAlt(Index).Enabled = False
    ComMotExc(Index).Enabled = False
    ComMotInc(Index).Enabled = False
End Sub

Private Sub Command1_Click(Index As Integer)
    Select Case Index
    Case 0
        TXTpc(11) = TXTpc(14)
    Case 1
        TXTpre(11) = TXTpc(14)
    Case 2
        TXTgp12(11) = TXTpc(14)
    End Select
End Sub

Private Sub ComMotAlt_Click(Index As Integer)
    Dim cTIPO As String

    If gridrevi(Index).Row = 0 Then
        Exit Sub
    End If
        
    gridrevi(Index).Col = 1
    nREV = gridrevi(Index)
    
    gridrevi(Index).Col = 3
    nREVI = gridrevi(Index)
    
    lEDITAR = True
    If Index = 6 Or Index = 7 Then
        lEDITAR = False
    End If

    cTIPO = CTIPOPLANO(Index)
    ePASS01 = "select * from revi WHERE PF=" & nPF & " AND TIPO='" & cTIPO & "' AND REVISAO=" & nREV & " AND  ITEM=" & nREVI
    frmREVI.Show vbModal, Me
    Filgridrevi Index, nREV

End Sub

Private Sub ComMotExc_Click(Index As Integer)
    Dim cTIPO As String
    Dim sSQL As String
    If gridrevi(Index).Row = 0 Then
        Exit Sub
    End If
    
    gridrevi(Index).Col = 1
    nREV = gridrevi(Index)

    gridrevi(Index).Col = 3
    nREVI = gridrevi(Index)
    
    cTIPO = CTIPOPLANO(Index)
    
    sSQL = "select * from revi WHERE PF=" & nPF & " AND TIPO='" & cTIPO & "' AND REVISAO=" & nREV & " AND  ITEM=" & nREVI
    If ApagaSQLP(cARQPF, sSQL) Then
        Filgridrevi Index, nREV
    End If
End Sub

Private Sub ComMotInc_Click(Index As Integer)
    Dim nITEM As Long
    Dim cTIPO As String
    Dim cDIZ As String

    cDIZ = ""
    If Index = 26 Then
        ComverRev_Click (0)
        cDIZ = "PF: " & nPF & "Revisao: " & TXTPF(17) & " Nao necessita revisar plano de controle"
        Index = 0
        gridrev(Index).Row = gridrev(Index).Rows - 1
        TXTpc(32) = Today()
    End If
    If gridrev(Index).Row = 0 Then
        Exit Sub
    End If
    CmdMotRev_Click (Index)                      'Pega revisao  pela funcao
    cTIPO = CTIPOPLANO(Index)
    If gridrevi(Index).Rows > 1 Then
        nITEM = PegUltGrid(gridrevi(Index), 3)
    Else
        nITEM = 0
    End If
    nITEM = nITEM + 1
    IncluiSQL cARQPF, "SELECT * FROM REVI WHERE PF=" & nPF & " AND TIPO='" & cTIPO & "' AND REVISAO=" & nREV & " AND ITEM=" & nITEM, 5, Array("PF", "REVISAO", "TIPO", "ITEM", "MOTIVO"), _
        Array(nPF, nREV, cTIPO, nITEM, cDIZ)
    Filgridrevi Index, nREV
End Sub

Private Function CTIPOPLANO(Index As Integer) As String
    CTIPOPLANO = "XXX"
    Select Case Index
    Case 0
        CTIPOPLANO = "PC"
    Case 1
        CTIPOPLANO = "PCR"                       'Plano de Controle Recebimento
    Case 2
        CTIPOPLANO = "PCL"
    Case 3
        CTIPOPLANO = "PCP"
    Case 4
        CTIPOPLANO = "PCF"
    Case 5
        CTIPOPLANO = "PCC"
    Case 6
        CTIPOPLANO = "PF"
    Case 7
        CTIPOPLANO = "FE"
    Case 11
        CTIPOPLANO = "PPR"                       'Plano de Controle Preliminar Recebimento
    Case 12
        CTIPOPLANO = "PPL"                       'Plano de Controle Preliminar lABORATORIO
    Case 13
        CTIPOPLANO = "PPP"                       'Plano de Controle Preliminar Producao
    Case 14
        CTIPOPLANO = "PPF"                       'Plano de Controle Preliminar fINAL
    Case 15
        CTIPOPLANO = "PPC"                       'Plano de Controle Preliminar Componentes
    Case 21
        CTIPOPLANO = "PGR"                       'Plano de Controle GP12 Recebimento
    Case 22
        CTIPOPLANO = "PGL"                       'Plano de Controle GP12 lABORATORIO
    Case 23
        CTIPOPLANO = "PGP"                       'Plano de Controle GP12 Producao
    Case 24
        CTIPOPLANO = "PGF"                       'Plano de Controle GP12 fINAL
    Case 25
        CTIPOPLANO = "PGC"                       'Plano de Controle GP12 Componentes
    End Select
End Function

Private Sub ComverRev_Click(Index As Integer)
    If Index = 6 Or Index = 7 Then
        If gridrevi(Index).Row = 0 Then
            Exit Sub
        End If
        Filgridrevi Index, nREV                  'pf e femea todas revisoes direto
    Else
        If gridrev(Index).Row = 0 Then
            Exit Sub
        End If
        Filgridrev (Index)
    End If
End Sub

Private Sub EditSeq_Click(Index As Integer)
    nSEQ = 0
    nSSQ = 0



    Select Case Index

    Case 3                                       'Produto
        If GridSeq(0).Row > 0 Then
            GridSeq(0).Col = 1
            nSEQ = GridSeq(0)
            GridSeq(0).Col = 2
            nSSQ = GridSeq(0)
            nARQPCS = 0
            Sdb = cARQPF
        End If
            
    Case 13                                      'Produto Preliminar
        If GridSeq(1).Row > 0 Then
            GridSeq(1).Col = 1
            nSEQ = GridSeq(1)
            GridSeq(1).Col = 2
            nSSQ = GridSeq(1)
            nARQPCS = 0
            Sdb = cARQPFP
        End If
    Case 23                                      'Produto gp12
        If GridSeq(1).Row > 0 Then
            GridSeq(2).Col = 1
            nSEQ = GridSeq(2)
            GridSeq(2).Col = 2
            nSSQ = GridSeq(2)
            nARQPCS = 0
            Sdb = cARQPFG
        End If
            

    Case 1                                       'Recebimento
        If GRIDMAT(0).Row > 0 Then
            GRIDMAT(0).Col = 1
            nSEQ = GRIDMAT(0)
            GRIDMAT(0).Col = 2
            nSSQ = GRIDMAT(0)
            nARQPCS = 1
            Sdb = cARQPF
        End If
    Case 11                                      'Recebimento Preliminar
        If GRIDMAT(1).Row > 0 Then
            GRIDMAT(1).Col = 1
            nSEQ = GRIDMAT(1)
            GRIDMAT(1).Col = 2
            nSSQ = GRIDMAT(1)
            nARQPCS = 1
            Sdb = cARQPFP
        End If
          
    Case 21                                      'Recebimento GP12
        If GRIDMAT(2).Row > 0 Then
            GRIDMAT(2).Col = 1
            nSEQ = GRIDMAT(2)
            GRIDMAT(2).Col = 2
            nSSQ = GRIDMAT(2)
            nARQPCS = 1
            Sdb = cARQPFG
        End If
          
          
    Case 2                                       ''laboratorio
        If gridlab(0).Row > 0 Then
            gridlab(0).Col = 1
            nSEQ = gridlab(0)
            gridlab(0).Col = 2
            nSSQ = gridlab(0)
            nARQPCS = 1
            Sdb = cARQPF
        End If
    Case 12                                      ''laboratorio Preliminar
        If gridlab(1).Row > 0 Then
            gridlab(1).Col = 1
            nSEQ = gridlab(1)
            gridlab(1).Col = 2
            nSSQ = gridlab(1)
            nARQPCS = 1
            Sdb = cARQPFP
        End If
           
    Case 22                                      ''laboratorio Preliminar
        If gridlab(2).Row > 0 Then
            gridlab(2).Col = 1
            nSEQ = gridlab(2)
            gridlab(2).Col = 2
            nSSQ = gridlab(2)
            nARQPCS = 1
            Sdb = cARQPFG
        End If
           
           
    Case 4                                       'Final
        If GRIDFIM(0).Row > 0 Then
            GRIDFIM(0).Col = 1
            nSEQ = GRIDFIM(0)
            GRIDFIM(0).Col = 2
            nSSQ = GRIDFIM(0)
            nARQPCS = 1
            Sdb = cARQPF
        End If
    Case 14                                      'Final Preliminar
        If GRIDFIM(1).Row > 0 Then
            GRIDFIM(1).Col = 1
            nSEQ = GRIDFIM(1)
            GRIDFIM(1).Col = 2
            nSSQ = GRIDFIM(1)
            nARQPCS = 1
            Sdb = cARQPFP
        End If
    Case 24                                      'Final GP12
        If GRIDFIM(2).Row > 0 Then
            GRIDFIM(2).Col = 1
            nSEQ = GRIDFIM(2)
            GRIDFIM(2).Col = 2
            nSSQ = GRIDFIM(2)
            nARQPCS = 1
            Sdb = cARQPFG
        End If
    Case 5                                       'Componentes
          
        If GRIDCOM(0).Row > 0 Then
            GRIDCOM(0).Col = 0
            Ctipoent = GRIDCOM(0)
            GRIDCOM(0).Col = 1
            Ccodcomp = GRIDCOM(0)
            nARQPCS = 2
            Sdb = cARQPF
        End If
    Case 15                                      'Componentes Preliminar
        If GRIDCOM(1).Row > 0 Then
            GRIDCOM(1).Col = 0
            Ctipoent = GRIDCOM(1)
            GRIDCOM(1).Col = 1
            Ccodcomp = GRIDCOM(1)
            nARQPCS = 2
            Sdb = cARQPFP
        End If

    Case 25                                      'Componentes GP12
        If GRIDCOM(2).Row > 0 Then
            GRIDCOM(2).Col = 0
            Ctipoent = GRIDCOM(2)
            GRIDCOM(2).Col = 1
            Ccodcomp = GRIDCOM(2)
            nARQPCS = 2
            Sdb = cARQPFG
        End If

    End Select
    If Index <> 5 And Index <> 15 And Index <> 25 Then 'componentes nao tem operacao
        nSEQ = FixInt(nSEQ)
        nSSQ = FixInt(nSSQ)
        If nSEQ = 0 And nSSQ = 0 Then
            Alert "Escolha uma Sequencia"
            Exit Sub
        End If
    Else
        Ctipoent = FixStr(Ctipoent)
        Ccodcomp = FixStr(Ccodcomp)
        If Ctipoent = "" Or Ccodcomp = "" Then
            Alert "Escolha um Componente"
            Exit Sub
        End If
    End If

    FRMPCS.Show vbModal

End Sub

Private Sub Encerrar_Click()
    If Not MDG("Sair sem gravar") Then
        Exit Sub
    End If
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub escidfolha_Click(Index As Integer)
    ePASS01 = ""
    If Index > 9 Then
        ePASS01 = "LOGIX"
    End If
    escMP04.Show vbModal, Me
    If lRETU Then

        Select Case Index
        Case 0, 10
            TXTpc(11).tEXT = TXTpc(11).tEXT & " , " & Trim(MMCase(eRETU02))
        Case 1, 11
            TXTpre(11).tEXT = TXTpre(11).tEXT & " , " & Trim(MMCase(eRETU02))
        Case 2, 12
            TXTgp12(11).tEXT = TXTgp12(11).tEXT & " , " & Trim(MMCase(eRETU02))
        Case 3, 13
            TXTpc(33).tEXT = eRETU01
            TXTpc(34).tEXT = eRETU02
        End Select
    End If


End Sub

Private Sub escpfela_Click(Index As Integer)
    If Not MDG("Gravar Elaborador", "Confirme Gravação") Then
        Exit Sub
    End If

    If Index = 0 Then
        frmPCX.TXTpc(12) = zIDFOLHA
        frmPCX.TXTpc(13) = zNOMEFOLHA
        frmPCX.TXTpc(10) = Date
    End If
    If Index = 1 Then
        frmPCX.TXTpre(12) = zIDFOLHA
        frmPCX.TXTpre(13) = zNOMEFOLHA
        frmPCX.TXTpre(10) = Date
    End If
    If Index = 2 Then
        frmPCX.TXTgp12(12) = zIDFOLHA
        frmPCX.TXTgp12(13) = zNOMEFOLHA
        frmPCX.TXTgp12(10) = Date
    End If
    
End Sub

Private Sub ESCprofin_Click(Index As Integer)
    Dim cCHAVEBUS
    If Len(TXTpc(22)) = 0 Then
        cCHAVEBUS = TXTPF(12)
    Else
        cCHAVEBUS = TXTpc(22)
    End If
    ePASS01 = "MANA5"
    If Index = 1 Then
        ePASS01 = "LOGIX"
    End If
    escms01.Show vbModal, Me
    If lRETU Then
        TXTpc(22) = eRETU01
    End If
    cCHAVEBUS = ""
End Sub

Private Sub escTIPPC_Click(Index As Integer)
    Dim cTIPO As String
    Dim nPOS
    
    ePASS01 = Array("(C)ontencao Avancada", "( )Normal", "(I)tem Seguranca", "(X)Contencao Avancada+Item Seguranca", "(G)P12 Agressivo", "GP12 Pla(N)", "Processo Alternativo-(B)Y PASS", "MUDANCA (T)emporaria")
    escOrdem.Show vbModal, Me
    If Not lRETU Then
        Exit Sub
    End If
    Select Case eRETU01
        Case 0
            cTIPO = "C"
        Case 1
            cTIPO = " "
        Case 2
            cTIPO = "I"
        Case 3
            cTIPO = "X"
        Case 4
            cTIPO = "G"
        Case 5
            cTIPO = "N"
        Case 6
            cTIPO = "B"
        Case 7
            cTIPO = "T"
    End Select
    
    Select Case Index
    Case 1, 2, 3, 4, 5
        nPOS = Index + 5 - 1
        TXTpc(nPOS).tEXT = cTIPO
    Case 11, 12, 13, 14, 15
        nPOS = Index + 5 - 11
        TXTpre(nPOS).tEXT = cTIPO
    Case 21, 22, 23, 24, 25
        nPOS = Index + 5 - 21
        TXTgp12(nPOS).tEXT = cTIPO
    End Select
End Sub

Private Sub FilCom()
    Dim cARQ As String
    Dim cSQL As String
    DizerBarra "Montando Lista Componetes"
    cARQ = PegPath("PATH", "PF")
    cSQL = "select TIPOENT,CODCOMP,DESCRI from PFMS03 WHERE PF=" & nPF & " ORDER BY TIPOENT,CODCOMP"
    MontaGridUltra GRIDCOM(0), 3, Array(600, 1600, 6000), Array("Tipo", "Codigo", "Nome"), _
        Array("TIPOENT", "L$CODCOMP", "DESCRI"), cARQ, cSQL
    CloneGrid GRIDCOM(0), GRIDCOM(1)
    CloneGrid GRIDCOM(0), GRIDCOM(2)
    DizerBarra ""
End Sub

Private Sub FilgridFIM()
    Dim cARQ As String
    Dim cSQL As String
    DizerBarra "Montando Lista final"
    cARQ = PegPath("PATH", "PF")
    cSQL = "select PF,SEQ,SSQ,DESCRI from PFSO WHERE PF=" & nPF & " AND SSQ=99 ORDER BY SEQ,SSQ"
    MontaGridUltra GRIDFIM(0), 4, Array(600, 400, 400, 6000), Array("PF", "SEQ", "SSQ", "Descriçao"), _
        Array("PF", "SEQ", "SSQ", "Descri"), cARQ, cSQL
    CloneGrid GRIDFIM(0), GRIDFIM(1)
    CloneGrid GRIDFIM(0), GRIDFIM(2)
    DizerBarra ""
End Sub

Private Sub FilgridLAB()
    Dim cARQ As String
    Dim cSQL As String
    DizerBarra "Montando Lista Laboratorio"
    cARQ = PegPath("PATH", "PF")
    cSQL = "select PF,SEQ,SSQ,DESCRI from PFSO WHERE PF=" & nPF & " AND SSQ=20 ORDER BY SEQ,SSQ"
    MontaGridUltra gridlab(0), 4, Array(600, 400, 400, 6000), Array("PF", "SEQ", "SSQ", "Descriçao"), _
        Array("PF", "SEQ", "SSQ", "Descri"), cARQ, cSQL
    CloneGrid gridlab(0), gridlab(1)
    CloneGrid gridlab(0), gridlab(2)
    DizerBarra ""
End Sub

Private Sub Filgridmat()
    Dim cARQ As String
    Dim cSQL As String
    DizerBarra "Montando LIsta Materia Prima"
    cARQ = PegPath("PATH", "PF")
    cSQL = "select PF,SEQ,SSQ,DESCRI from PFSO WHERE PF=" & nPF & " AND SSQ=10 ORDER BY SEQ,SSQ"
    MontaGridUltra GRIDMAT(0), 4, Array(600, 400, 400, 4000), Array("PF", "SEQ", "SSQ", "Descriçao"), _
        Array("PF", "SEQ", "SSQ", "DESCRI"), cARQ, cSQL
    CloneGrid GRIDMAT(0), GRIDMAT(1)
    CloneGrid GRIDMAT(0), GRIDMAT(2)
    DizerBarra ""
End Sub

Private Sub Filgridrev(nINDICE As Integer)
    Dim cARQ As String
    Dim cSQL As String
    Dim cTIPO As String
    DizerBarra "Carregando Revisoes " & nINDICE
    cTIPO = CTIPOPLANO(nINDICE)
    cARQ = PegPath("PATH", "PF")
    cSQL = "select PF,REVISAO,TIPO,DATA from rev WHERE PF=" & nPF & " AND TIPO='" & cTIPO & "' ORDER BY REVISAO"
    MontaGridFast gridrev(nINDICE), 4, Array(600, 400, 400, 1200), Array("PF", "REV", "Tipo", "DATA"), _
        Array("PF", "REVISAO", "TIPO", "C$DATA"), cARQ, cSQL
    DizerBarra ""
End Sub

Private Sub Filgridrevi(ByVal nINDICE As Integer, ByVal nSUBREV As Integer)
    Dim cTIPO As String
    Dim cARQ As String
    Dim cSQL As String
    DizerBarra "Carregando Itens Revisoes " & nINDICE
    cTIPO = CTIPOPLANO(nINDICE)
    cARQ = PegPath("PATH", "PF")
    If nINDICE = 6 Or nINDICE = 7 Or nSUBREV = 0 Then
        cSQL = "select PF,REVISAO,TIPO,ITEM,MOTIVO from revI WHERE PF=" & nPF & " AND TIPO='" & cTIPO & "' ORDER BY revisao,ITEM"
    Else
        cSQL = "select PF,REVISAO,TIPO,ITEM,MOTIVO from revI WHERE PF=" & nPF & " AND REVISAO=" & FixInt(nREV) & " AND TIPO='" & cTIPO & "' ORDER BY revisao,ITEM"
    End If
    MontaGridFast gridrevi(nINDICE), 5, Array(600, 400, 400, 400, 6000), Array("PF", "REV", "Tipo", "Item", "Motivo"), _
        Array("PF", "REVISAO", "TIPO", "ITEM", "MOTIVO"), cARQ, cSQL
    DizerBarra ""
End Sub

Private Sub Filgridseq()
    Dim cARQ As String
    Dim cSQL As String
    DizerBarra "Carregando Sequencias de Operacoes"
    cARQ = PegPath("PATH", "PF")
    cSQL = "select PF,SEQ,SSQ,CODINT,DESCRI from PFS WHERE PF=" & nPF & " ORDER BY SEQ,SSQ"
    MontaGridUltra GridSeq(0), 5, Array(600, 400, 400, 1600, 5000), Array("PF", "SEQ", "SSQ", "CodInt", "Descriçao"), _
        Array("PF", "SEQ", "SSQ", "L$CODINT", "L$DESCRI"), cARQ, cSQL
    CloneGrid GridSeq(0), GridSeq(1)
    CloneGrid GridSeq(0), GridSeq(2)
    DizerBarra ""
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    TeclaEnter KeyCode
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim X As Integer
    Dim cSQL As String
    
    CenterFormToScreen Me
    cARQPF = PegPath("PATH", "PF")
    cARQPFP = PegPath("PATH", "PFP")
    cARQPFG = PegPath("PATH", "PFG")
    
    CHECKPFPG                                    'inclui embalagem preliminar gp12
    
    TXTNPF.tEXT = nPF
    
    
    cSQL = "select PF,CODCLIENTE,CPF,CODIGO,DESCR,CLIENTE,CLINOME,CODMU011,CODMU012,CODMU013,NOMMU011,NOMMU012,NOMMU013,CODFINAL,CODIGOINT,OPCAO,FEMEAREV,FEMEAREVD,REVPRO,REVDAT,FEMEAREVD2 from PF WHERE PF=" & nPF
    
    
    aCAM = Array("CODCLIENTE", "CPF", "CODIGO", "DESCR", "CLIENTE", "CLINOME", "CODMU011", "CODMU012", "CODMU013", "NOMMU011", _
                 "NOMMU012", "NOMMU013", "CODFINAL", "CODIGOINT", "OPCAO", "FEMEAREV", "FEMEAREVD", "REVPRO", "REVDAT", "FEMEAREVD2")
    aFOR = Array("C", "C", "C", "C", "NI", "C", "C", "C", "C", "C", _
                 "C", "C", "C", "C", "NI", "NI", "DZ", "NI", "DZ", "DZ")
    
    aPAD = Array("", "", "", "", 0, "", "", "", "", "", _
                 "", "", "", "", 0, 0, "", 0, "", "")
    
    aVAL = PegSQL(cARQPF, cSQL, 20, aCAM, aFOR, aPAD)
    For iLOOP = 0 To 19
        TXTPF(iLOOP) = aVAL(iLOOP)
    Next iLOOP
    
    
     cSQL = "select PF,PCRREV,PCLREV,PCPREV,PCFREV,PCCREV,PCRTIPO,PCLTIPO,PCPTIPO,PCFTIPO,PCCTIPO,"
     cSQL = cSQL & "PCELADAT,PCGRUPO,PCELANUM,PCELANOM,FEMEAG,PCRDAT,PCLDAT,PCPDAT,PCFDAT,PCCDAT,"
     cSQL = cSQL & "PCDINI,OUTRAAPR,CODFISCAL,CRITIOBS,FILIAL,SELDATA,CLIESP,STPFPC,STFEPC,PCBLOQ,"
     cSQL = cSQL & "SEL100,CRITICO,PCREVD2,CONTATONUM,CONTATONOM"
     cSQL = cSQL & " from PF WHERE PF=" & nPF
    
    
    aCAMPC = Array("PCRREV", "PCLREV", "PCPREV", "PCFREV", "PCCREV", "PCRTIPO", "PCLTIPO", "PCPTIPO", "PCFTIPO", "PCCTIPO", _
                   "PCELADAT", "PCGRUPO", "PCELANUM", "PCELANOM", "FEMEAG", "PCRDAT", "PCLDAT", "PCPDAT", "PCFDAT", "PCCDAT", _
                   "PCDINI", "OUTRAAPR", "CODFISCAL", "CRITIOBS", "FILIAL", "SELDATA", "CLIESP", "STPFPC", "STFEPC", "PCBLOQ", _
                   "SEL100", "CRITICO", "PCREVD2", "CONTATONUM", "CONTATONOM")
    aFORPC = Array("NI", "NI", "NI", "NI", "NI", "C", "C", "C", "C", "C", _
                   "DZ", "C", "NI", "C", "C", "DZ", "DZ", "DZ", "DZ", "DZ", _
                   "D", "C", "C", "C", "C", "D", "C", "BN", "BN", "BN", _
                   "BN", "BN", "DZ", "NI", "C")
    aPADPC = Array(0, 0, 0, 0, 0, 0, "", "", "", "", _
                   "", "", 0, "", "", "", "", "", "", "", _
                   "", "", "", "", "", "", "", "", "", "", "", _
                   "", "", "", 0, "")

    ''PC
    aVALPC = PegSQL(cARQPF, cSQL, 35, aCAMPC, aFORPC, aPADPC)
    For iLOOP = 0 To 19
        TXTpc(iLOOP) = aVALPC(iLOOP)
    Next iLOOP
    For iLOOP = 21 To 26
        TXTpc(iLOOP) = aVALPC(iLOOP)
    Next iLOOP
    If IsDate(aVALPC(20)) Then
        PCDinipc.Value = aVALPC(20)
    End If
    Check1.Value = aVALPC(27)
    Check2.Value = aVALPC(28)
    Check3.Value = aVALPC(29)
    Check4.Value = aVALPC(30)
    Check5.Value = aVALPC(31)
    TXTpc(32) = aVALPC(32)
    TXTpc(33) = aVALPC(33)
    TXTpc(34) = aVALPC(34)
    
    ''PRE
    aVALPRE = PegSQL(cARQPFP, cSQL, 21, aCAMPC, aFORPC, aPADPC)
    For iLOOP = 0 To 19
        TXTpre(iLOOP) = aVALPRE(iLOOP)
    Next iLOOP
    If IsDate(aVALPRE(20)) Then
        PCDinipre.Value = aVALPRE(20)
    End If
     
    ''GP12
    aVALGP12 = PegSQL(cARQPFG, cSQL, 21, aCAMPC, aFORPC, aPADPC)
    For iLOOP = 0 To 19
        TXTgp12(iLOOP) = aVALGP12(iLOOP)
    Next iLOOP
    If IsDate(aVALGP12(20)) Then
        PCDinigp12.Value = aVALGP12(20)
    End If
    
    Filgridseq
    Filgridmat
    FilgridLAB
    FilgridFIM
    FilCom
    FILGRIdpfrev
    
    For X = 0 To 5
        ComMotAlt(X).Enabled = False
        ComMotExc(X).Enabled = False
        ComMotInc(X).Enabled = False
    Next X
    For X = 11 To 15
        ComMotAlt(X).Enabled = False
        ComMotExc(X).Enabled = False
        ComMotInc(X).Enabled = False
    Next X
    For X = 21 To 25
        ComMotAlt(X).Enabled = False
        ComMotExc(X).Enabled = False
        ComMotInc(X).Enabled = False
    Next X
    

    If zUSER = "ADMIN" Then
        'Liberacoes ajustes
        For X = 0 To 4
            TXTpc(X).Enabled = True
            TXTpc(15 + X).Enabled = True
            TXTpre(X).Enabled = True
            TXTpre(15 + X).Enabled = True
            TXTgp12(X).Enabled = True
            TXTgp12(15 + X).Enabled = True
        Next X
    End If

End Sub

Private Sub NewMat_Click(Index As Integer)
    Dim cDESCRI As String
    Dim sSQL As String
   
    Select Case Index

    Case 11, 21
        nSEQ = 10

    Case 12, 22
        nSEQ = 20

    Case 13, 23
        nSEQ = 30

    Case 4
        nSEQ = 99

    End Select

    Select Case Index

    Case 11, 12, 13
        nSSQ = 10

    Case 21, 22, 23
        nSSQ = 20

    Case 4
        nSSQ = 99

    End Select
    
    Select Case Index

    Case 11
        cDESCRI = "Matéria Prima: " & TXTPF(6)

    Case 12
        cDESCRI = "Matéria Prima: " & TXTPF(7)

    Case 13
        cDESCRI = "Matéria Prima: " & TXTPF(8)

    Case 21
        cDESCRI = "Laboratório: " & TXTPF(6)

    Case 22
        cDESCRI = "Laboratório: " & TXTPF(7)

    Case 23
        cDESCRI = "Laboratório: " & TXTPF(8)

    Case 4
        cDESCRI = "Final "

    End Select


    sSQL = "select * from PFSO WHERE PF=" & nPF & " AND SEQ=" & nSEQ & " AND SSQ=" & nSSQ
    IncluiSQL cARQPF, sSQL, 4, Array("PF", "SEQ", "SSQ", "DESCRI"), _
        Array(nPF, nSEQ, nSSQ, cDESCRI), True, False
                            
    IncluiSQL cARQPFP, sSQL, 4, Array("PF", "SEQ", "SSQ", "DESCRI"), _
        Array(nPF, nSEQ, nSSQ, cDESCRI), True, False
                                
    IncluiSQL cARQPFG, sSQL, 4, Array("PF", "SEQ", "SSQ", "DESCRI"), _
        Array(nPF, nSEQ, nSSQ, cDESCRI), True, False
                                 
                            
    Select Case Index
    Case 11, 12, 13
        Filgridmat
    Case 21, 22, 23
        FilgridLAB
    Case 4
        FilgridFIM
    End Select

End Sub

Private Sub selmar_Click(Index As Integer)
    'Dim CLINHA As String
    Dim cCODIGO As String
    Dim cARQ As String
    Dim sSQL As String
    Dim aRETU As Variant
    Dim dDATA


    On Error Resume Next

    cARQ = PegPath("PATH", "MANA5FIN")
    cARQ = GeraConn(cARQ, "SDECDX")
    cCODIGO = Trim(TXTpc(22))
    sSQL = "SELECT SEL100,RIFSEL FROM Rifpr WHERE PRODUTO='" & cCODIGO & "'"
    If Len(cCODIGO) = 0 Then
        Alert "Codigo Final/Fiscal em Branco"
        Exit Sub
    End If
    Select Case Index
    Case 0
        ePASS01 = Date + 30
        frmInsertDate.Show vbModal, Me
        dDATA = Fdata(eRETU01)
        If Not IncluiSQL(cARQ, sSQL, 3, Array("PRODUTO", "SEL100", "RIFSEL"), Array(cCODIGO, dDATA, 0), True, False) Then
            GrvSQL cARQ, sSQL, 2, Array("SEL100", "RIFSEL"), Array(dDATA, 0), Array("", "")
        End If
    Case 1
        GrvSQL cARQ, sSQL, 2, Array("SEL100", "RIFSEL"), Array(Space(8), 0), Array("", "")
    End Select
    
    
    aRETU = PegSQL(cARQ, sSQL, 1, Array("SEL100"), Array("DC"), Array(""))
    eRETU01 = aRETU(0)
    Check4.Value = Unchecked
    If IsDate(eRETU01) Then
        If Day(eRETU01) > 0 Then
            Check4.Value = Checked
            TXTpc(25) = eRETU01
        Else
            TXTpc(25) = NullDate()
        End If
    Else
        TXTpc(25) = NullDate()
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault
End Sub

Private Sub FILGRIdpfrev()
    Dim cSQL As String
    DizerBarra "Carregando Revisoes Processo Fabricacao"
    cSQL = "SELECT PF,REV,PFDATA,PCRD,PCLD,PCFD,PCCD,PCPD FROM PFREV WHERE PF=" & nPF & " ORDER BY REV"
    MontaGridFast GridPFrev, 8, Array(600, 400, 1000, 1000, 1000, 1000, 1000, 1000), _
        Array("PF", "REV", "DATA", "Rec", "Lab", "Fin", "Com", "Pro"), _
        Array("PF", "REV", "PFDATA", "PCRD", "PCLD", "PCFD", "PCCD", "PCPD"), cARQPF, cSQL
    DizerBarra ""
End Sub


