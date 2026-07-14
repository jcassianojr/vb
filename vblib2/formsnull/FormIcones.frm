VERSION 5.00
Object = "{379157C5-E9BD-43F1-9F83-B037496BED42}#1.3#0"; "vbccr18.ocx"
Begin VB.Form FormIcones 
   Caption         =   "Form1"
   ClientHeight    =   4824
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   9984
   LinkTopic       =   "Form1"
   ScaleHeight     =   4824
   ScaleWidth      =   9984
   StartUpPosition =   3  'Windows Default
   Begin VBCCR18.CommandButtonW btnclose 
      Height          =   612
      Left            =   7680
      TabIndex        =   0
      Top             =   360
      Width           =   972
      _ExtentX        =   1715
      _ExtentY        =   1080
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ImageListAlignment=   1
      Caption         =   "Retornar"
      Alignment       =   0
      VerticalAlignment=   0
      Picture         =   "FormIcones.frx":0000
      PictureAndCaption=   -1  'True
      WordWrap        =   0   'False
      Style           =   1
   End
   Begin VBCCR18.CommandButtonW cmdClose 
      Height          =   612
      Left            =   7680
      TabIndex        =   1
      Top             =   1080
      Width           =   972
      _ExtentX        =   1715
      _ExtentY        =   1080
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ImageListAlignment=   1
      Caption         =   "Salvar"
      Alignment       =   0
      VerticalAlignment=   0
      Picture         =   "FormIcones.frx":059A
      PictureAndCaption=   -1  'True
      WordWrap        =   0   'False
      Style           =   1
   End
   Begin VBCCR18.CommandButtonW cmdOK 
      Height          =   612
      Left            =   6600
      TabIndex        =   2
      Top             =   360
      Width           =   972
      _ExtentX        =   1715
      _ExtentY        =   1080
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ImageListAlignment=   1
      Caption         =   "Confimar"
      Alignment       =   0
      VerticalAlignment=   0
      Picture         =   "FormIcones.frx":0B34
      PictureAndCaption=   -1  'True
      WordWrap        =   0   'False
      Style           =   1
   End
   Begin VBCCR18.CommandButtonW cmdimp 
      Height          =   612
      Left            =   6600
      TabIndex        =   3
      Top             =   1080
      Width           =   972
      _ExtentX        =   1715
      _ExtentY        =   1080
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ImageListAlignment=   1
      Caption         =   "Imprimir"
      Alignment       =   0
      VerticalAlignment=   0
      Picture         =   "FormIcones.frx":0FCE
      PictureAndCaption=   -1  'True
      WordWrap        =   0   'False
      Style           =   1
   End
   Begin VBCCR18.CommandButtonW cmdCANCELAR 
      Height          =   612
      Left            =   5520
      TabIndex        =   4
      Top             =   360
      Width           =   972
      _ExtentX        =   1715
      _ExtentY        =   1080
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ImageListAlignment=   1
      Caption         =   "Cancelar"
      Alignment       =   0
      VerticalAlignment=   0
      Picture         =   "FormIcones.frx":1568
      PictureAndCaption=   -1  'True
      WordWrap        =   0   'False
      Style           =   1
   End
   Begin VBCCR18.CommandButtonW Command1 
      Height          =   612
      Left            =   4320
      TabIndex        =   5
      Top             =   360
      Width           =   972
      _ExtentX        =   1715
      _ExtentY        =   1080
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ImageListAlignment=   1
      Caption         =   "Localizar"
      Alignment       =   0
      VerticalAlignment=   0
      Picture         =   "FormIcones.frx":1A02
      PictureAndCaption=   -1  'True
      WordWrap        =   0   'False
      Style           =   1
   End
   Begin VBCCR18.CommandButtonW CommandButtonW1 
      Height          =   252
      Left            =   4320
      TabIndex        =   6
      Top             =   1200
      Width           =   372
      _ExtentX        =   656
      _ExtentY        =   445
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ImageListAlignment=   1
      Alignment       =   0
      VerticalAlignment=   0
      Picture         =   "FormIcones.frx":1F9C
      WordWrap        =   0   'False
      Style           =   1
   End
   Begin VBCCR18.CommandButtonW escidfolha 
      Height          =   492
      Index           =   0
      Left            =   4320
      TabIndex        =   7
      Top             =   1560
      Width           =   372
      _ExtentX        =   656
      _ExtentY        =   868
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ImageListAlignment=   1
      Caption         =   "  "
      Alignment       =   0
      VerticalAlignment=   0
      Picture         =   "FormIcones.frx":2536
      WordWrap        =   0   'False
      Style           =   1
   End
   Begin VBCCR18.CommandButtonW CommandButtonW2 
      Height          =   612
      Left            =   5520
      TabIndex        =   8
      Top             =   1080
      Width           =   972
      _ExtentX        =   1715
      _ExtentY        =   1080
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ImageListAlignment=   1
      Caption         =   "Visualizar"
      Alignment       =   0
      VerticalAlignment=   0
      Picture         =   "FormIcones.frx":2AD0
      PictureAndCaption=   -1  'True
      WordWrap        =   0   'False
      Style           =   1
   End
   Begin VBCCR18.CommandButtonW CommandButtonW3 
      Height          =   252
      Left            =   3720
      TabIndex        =   9
      Top             =   1200
      Width           =   372
      _ExtentX        =   656
      _ExtentY        =   445
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ImageListAlignment=   1
      Alignment       =   0
      VerticalAlignment=   0
      Picture         =   "FormIcones.frx":306A
      WordWrap        =   0   'False
      Style           =   1
   End
   Begin VBCCR18.CommandButtonW CommandButtonW4 
      Height          =   252
      Left            =   3720
      TabIndex        =   10
      Top             =   1560
      Width           =   372
      _ExtentX        =   656
      _ExtentY        =   445
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ImageListAlignment=   1
      Alignment       =   0
      VerticalAlignment=   0
      Picture         =   "FormIcones.frx":3604
      WordWrap        =   0   'False
      Style           =   1
   End
   Begin VBCCR18.CommandButtonW CMDX 
      Height          =   612
      Left            =   5520
      TabIndex        =   11
      Top             =   1800
      Width           =   972
      _ExtentX        =   1715
      _ExtentY        =   1080
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ImageListAlignment=   1
      Caption         =   "Email"
      Alignment       =   0
      VerticalAlignment=   0
      Picture         =   "FormIcones.frx":399E
      PictureAndCaption=   -1  'True
      WordWrap        =   0   'False
      Style           =   1
   End
   Begin VBCCR18.CommandButtonW CommandButtonW5 
      Height          =   612
      Left            =   6600
      TabIndex        =   12
      Top             =   1800
      Width           =   972
      _ExtentX        =   1715
      _ExtentY        =   1080
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ImageListAlignment=   1
      Caption         =   "Configurar"
      Alignment       =   0
      VerticalAlignment=   0
      Picture         =   "FormIcones.frx":3F38
      PictureAndCaption=   -1  'True
      WordWrap        =   0   'False
      Style           =   1
   End
   Begin VBCCR18.CommandButtonW CommandButtonW6 
      Height          =   612
      Left            =   5520
      TabIndex        =   13
      Top             =   2520
      Width           =   972
      _ExtentX        =   1715
      _ExtentY        =   1080
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ImageListAlignment=   1
      Caption         =   "Filtrar"
      Alignment       =   0
      VerticalAlignment=   0
      Picture         =   "FormIcones.frx":44D2
      PictureAndCaption=   -1  'True
      WordWrap        =   0   'False
      Style           =   1
   End
   Begin VBCCR18.CommandButtonW Command8 
      Height          =   612
      Left            =   6600
      TabIndex        =   14
      Top             =   2520
      Width           =   972
      _ExtentX        =   1715
      _ExtentY        =   1080
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ImageListAlignment=   1
      Caption         =   "Limpar"
      Alignment       =   0
      VerticalAlignment=   0
      Picture         =   "FormIcones.frx":496C
      PictureAndCaption=   -1  'True
      WordWrap        =   0   'False
      Style           =   1
   End
   Begin VBCCR18.CommandButtonW CommandButtonW7 
      Height          =   252
      Left            =   960
      TabIndex        =   15
      Top             =   960
      Width           =   372
      _ExtentX        =   656
      _ExtentY        =   445
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ImageListAlignment=   1
      Alignment       =   0
      VerticalAlignment=   0
      Picture         =   "FormIcones.frx":4E06
      WordWrap        =   0   'False
      Style           =   1
   End
   Begin VBCCR18.CommandButtonW CmdPreviewGRID 
      Height          =   372
      Left            =   1560
      TabIndex        =   16
      Top             =   960
      Width           =   612
      _ExtentX        =   1080
      _ExtentY        =   656
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ImageListAlignment=   1
      Alignment       =   0
      VerticalAlignment=   0
      Picture         =   "FormIcones.frx":53A0
      WordWrap        =   0   'False
      Style           =   1
   End
End
Attribute VB_Name = "FormIcones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
