VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "xpcontrols.ocx"
Begin VB.Form frmInsertDate 
   Caption         =   "Inserir Datas"
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7305
   Icon            =   "frmInsertDate.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3180
   ScaleWidth      =   7305
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtData 
      Enabled         =   0   'False
      Height          =   375
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2640
      Width           =   5295
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2370
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   120324097
      CurrentDate     =   36522
   End
   Begin VB.ListBox lstDates 
      Height          =   2400
      Left            =   2760
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
   Begin XPControls.XPButton cmdOK 
      Height          =   435
      Left            =   5520
      TabIndex        =   3
      Top             =   240
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "frmInsertDate.frx":0442
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
   Begin XPControls.XPButton cmdCancel 
      Height          =   435
      Left            =   5520
      TabIndex        =   4
      Top             =   840
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "frmInsertDate.frx":09DC
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
End
Attribute VB_Name = "frmInsertDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Whe the form is loaded then Listbox is seeded with the current date
'When a user clicks on the Monthlist this reseeds the list with
'the date selected

Private Sub SetDateList(DateSelected As Date)
    'Purpose:  Seed the Listbox with the correct date formats

    lstDates.Clear
    
    'Month Year only
    lstDates.AddItem Format(DateSelected, "mmm yy")
    lstDates.AddItem Format(DateSelected, "mmm yyyy")
    lstDates.AddItem Format(DateSelected, "mmmm yy")
    lstDates.AddItem Format(DateSelected, "mmmm yyyy")
    
    'Short dates
    lstDates.AddItem Format(DateSelected, "m-d-yy")
    lstDates.AddItem Format(DateSelected, "m-d-yyyy")
    If Format(DateSelected, "mm-dd-yy") <> Format(DateSelected, "m-d-yy") Then
        lstDates.AddItem Format(DateSelected, "mm-dd-yy")
        lstDates.AddItem Format(DateSelected, "mm-dd-yyyy")
    End If
    lstDates.AddItem Format(DateSelected, "d-mmm-yy")
    lstDates.AddItem Format(DateSelected, "d-mmm-yyyy")
    If Format(DateSelected, "dd-mmm-yy") <> Format(DateSelected, "d-mmm-yy") Then
        lstDates.AddItem Format(DateSelected, "dd-mmm-yy")
        lstDates.AddItem Format(DateSelected, "dd-mmm-yyyy")
    End If
    lstDates.AddItem Format(DateSelected, "m/d/yy")
    lstDates.AddItem Format(DateSelected, "m/d/yyyy")
    If Format(DateSelected, "mm/dd/yy") <> Format(DateSelected, "m/d/yy") Then
        lstDates.AddItem Format(DateSelected, "mm/dd/yy")
        lstDates.AddItem Format(DateSelected, "mm/dd/yyyy")
    End If
    lstDates.AddItem Format(DateSelected, "d/mmm/yy")
    lstDates.AddItem Format(DateSelected, "d/mmm/yyyy")
    If Format(DateSelected, "dd/mmm/yy") <> Format(DateSelected, "d/mmm/yy") Then
        lstDates.AddItem Format(DateSelected, "dd/mmm/yy")
        lstDates.AddItem Format(DateSelected, "dd/mmm/yyyy")
    End If
    lstDates.AddItem Format(DateSelected, "d mmm yy")
    lstDates.AddItem Format(DateSelected, "d mmm yyyy")
    lstDates.AddItem Format(DateSelected, "mmm, d yy")
    lstDates.AddItem Format(DateSelected, "mmm, dd yy")
    If Format(DateSelected, "dd mmm yy") <> Format(DateSelected, "d mmm yy") Then
        lstDates.AddItem Format(DateSelected, "dd mmm yy")
        lstDates.AddItem Format(DateSelected, "dd mmm yyyy")
        lstDates.AddItem Format(DateSelected, "mmm, d yyyy")
        lstDates.AddItem Format(DateSelected, "mmm, dd yyyy")
    End If
    
    'Long dates
    lstDates.AddItem Format(DateSelected, "d mmmm yy")
    lstDates.AddItem Format(DateSelected, "d mmmm yy")
    lstDates.AddItem Format(DateSelected, "mmmm, d yy")
    lstDates.AddItem Format(DateSelected, "mmmm, d yyyy")
    If Format(DateSelected, "mmmm, d yy") <> Format(DateSelected, "mmmm, dd yy") Then
        lstDates.AddItem Format(DateSelected, "dd mmmm yy")
        lstDates.AddItem Format(DateSelected, "dd mmmm yy")
        lstDates.AddItem Format(DateSelected, "mmmm, dd yy")
        lstDates.AddItem Format(DateSelected, "mmmm, dd yyyy")
    End If
    lstDates.AddItem Format(DateSelected, "ddd d mmmm yy")
    lstDates.AddItem Format(DateSelected, "ddd d mmmm yy")
    lstDates.AddItem Format(DateSelected, "ddd mmmm, d yy")
    lstDates.AddItem Format(DateSelected, "ddd mmmm, d yyyy")
    If Format(DateSelected, "mmmm, d yy") <> Format(DateSelected, "mmmm, dd yy") Then
        lstDates.AddItem Format(DateSelected, "ddd dd mmmm yy")
        lstDates.AddItem Format(DateSelected, "ddd dd mmmm yy")
        lstDates.AddItem Format(DateSelected, "ddd mmmm, dd yy")
        lstDates.AddItem Format(DateSelected, "ddd mmmm, dd yyyy")
    End If
    lstDates.AddItem Format(DateSelected, "dddd d mmmm yy")
    lstDates.AddItem Format(DateSelected, "dddd d mmmm yy")
    lstDates.AddItem Format(DateSelected, "dddd mmmm, d yy")
    lstDates.AddItem Format(DateSelected, "dddd mmmm, d yyyy")
    If Format(DateSelected, "mmmm, d yy") <> Format(DateSelected, "mmmm, dd yy") Then
        lstDates.AddItem Format(DateSelected, "dddd dd mmmm yy")
        lstDates.AddItem Format(DateSelected, "dddd dd mmmm yy")
        lstDates.AddItem Format(DateSelected, "dddd mmmm, dd yy")
        lstDates.AddItem Format(DateSelected, "dddd mmmm, dd yyyy")
    End If

    lstDates.ListIndex = 15                      'Default selection
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    lRETU = True
    eRETU01 = MonthView1.Value
    If lstDates.ListIndex >= 0 Then
        eRETU02 = lstDates.List(lstDates.ListIndex)
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    Center Me
    lRETU = False
    eRETU01 = Date
    eRETU02 = ""
    If IsDate(ePASS01) Then
        MonthView1.Value = ePASS01
    Else
        MonthView1.Value = Date
    End If
    TxtData = Format(MonthView1.Value, "dddd, dd/mmmm/yyyy")
    SetDateList MonthView1.Value
    
End Sub

Private Sub lstDates_DblClick()
    ''SetDateList DateClicked
    TxtData = Format(MonthView1.Value, "dddd, dd/mmmm/yyyy")
    cmdOK_Click
End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
    SetDateList DateClicked
    TxtData = Format(MonthView1.Value, "dddd, dd/mmmm/yyyy")
    cmdOK.SetFocus
End Sub

Private Sub MonthView1_DateDblClick(ByVal DateDblClicked As Date)
    SetDateList DateDblClicked
    TxtData = Format(MonthView1.Value, "dddd, dd/mmmm/yyyy")
    cmdOK_Click
End Sub

