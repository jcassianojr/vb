VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{379157C5-E9BD-43F1-9F83-B037496BED42}#1.1#0"; "vbccr18.ocx"
Begin VB.Form dCompare 
   Caption         =   "Database Comparador/Corretor"
   ClientHeight    =   5805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8955
   Icon            =   "dcompare.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5805
   ScaleWidth      =   8955
   StartUpPosition =   3  'Windows Default
   Begin XPControls.XPButton CmdExportarSqlite 
      Height          =   375
      Left            =   5160
      TabIndex        =   12
      Top             =   1200
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Caption         =   "Exportar Sqlite"
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
   Begin VBCCR18.TextBoxW Text3 
      Height          =   2895
      Left            =   120
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2280
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   5106
      Text            =   "dcompare.frx":058A
      MultiLine       =   -1  'True
      ScrollBars      =   2
   End
   Begin VBCCR18.TextBoxW Text2 
      Height          =   375
      Left            =   960
      TabIndex        =   9
      Top             =   600
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   661
   End
   Begin VBCCR18.TextBoxW Text1 
      Height          =   375
      Left            =   960
      TabIndex        =   8
      Top             =   120
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   661
   End
   Begin XPControls.XPButton Command3 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      Caption         =   "Compare Arquivos"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorScheme     =   1
      ColorBegin      =   8454143
      ColorEnd        =   8454143
   End
   Begin XPControls.XPButton Command2 
      Height          =   375
      Left            =   8520
      TabIndex        =   3
      Top             =   600
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Picture         =   "dcompare.frx":05B0
      Caption         =   ""
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
   Begin XPControls.XPButton Command1 
      Height          =   375
      Left            =   8520
      TabIndex        =   2
      Top             =   120
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Picture         =   "dcompare.frx":0B4A
      Caption         =   ""
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
   Begin XPControls.XPButton Command4 
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   1200
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      Caption         =   "Comparar e Corrigir"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorScheme     =   1
      ColorBegin      =   8454143
      ColorEnd        =   8454143
   End
   Begin XPControls.XPButton Command5 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      Picture         =   "dcompare.frx":10E4
      Caption         =   "Criar Tabela de Destino"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorScheme     =   1
      ColorBegin      =   8454143
      ColorEnd        =   8454143
   End
   Begin XPControls.XPButton Encerrar 
      Height          =   375
      Left            =   7560
      TabIndex        =   7
      Top             =   1080
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Picture         =   "dcompare.frx":167E
      Caption         =   "&Encerrar"
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
   Begin XPControls.XPButton CmdCompactar 
      Height          =   375
      Left            =   2520
      TabIndex        =   11
      Top             =   1680
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      Picture         =   "dcompare.frx":1C18
      Caption         =   "Compactar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorScheme     =   1
      ColorBegin      =   8454143
      ColorEnd        =   8454143
   End
   Begin VB.Label lProgress 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   5160
      TabIndex        =   13
      Top             =   1680
      Width           =   3360
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Destino"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   795
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Origem"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   270
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   825
   End
End
Attribute VB_Name = "dCompare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nERRO
Private WithEvents C As cConverter
Attribute C.VB_VarHelpID = -1
'Attribute C.VB_VarHelpID = -1

Private Sub CmdCompactar_Click()
  Dim lngBefore, lngAfter As Long
  Dim strResult As String
  lngBefore = FileLen(Text1.tEXT)
  If CompactAccess(Text1.tEXT) Then

    lngAfter = FileLen(Text1.tEXT)
    strResult = "Resultados da Compactacao" & vbCrLf
    strResult = strResult & "-------------------------------------------" & vbCrLf
    strResult = strResult & "Tamanho Antes  da Compactação: " & lngBefore & " bytes" & vbCrLf
    strResult = strResult & "Tamanho Depois da Compactação: " & lngAfter & " bytes"
    MsgBox strResult, vbInformation
  Else
    MsgBox "Erro na compactacao do arquivo", vbExclamation
  End If
End Sub

Private Sub CmdExportarSqlite_Click()
  ' Text2.tEXT = TrocaExt(Text1.tEXT, "sqlite")
   If MDG("Converter Reescreve Destino") Then
      Convert Text1.tEXT, Text2.tEXT
   End If
   
End Sub
Private Sub Convert(NWindMDBFileName$, SQLiteFileName$)
Dim i&, aCnn As ADODB.Connection, sCnn As cConnection
  On Error Resume Next
  Dim cCONN As String
  
  Set aCnn = New ADODB.Connection
  aCnn.CursorLocation = adUseClient
  cCONN = GeracArq(NWindMDBFileName)
  'aCnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & NWindMDBFileName
  aCnn.Open cCONN
  If Err Then MsgBox Err.Description: Err.Clear: Exit Sub
 
  'If FileExist(SQLiteFileName, True) Then
  '   If MDG("Excluir Destino") Then
  '      Kill SQLiteFileName 'ensure, that we kill the older File first
  '   Else
  '      Exit Sub
  '   End If
     
  'End If
  Err.Clear
  
  Set sCnn = New_c.Connection
  'sCnn.CreateNewDB SQLiteFileName
  sCnn.OpenDB SQLiteFileName
  
  If Err Then MsgBox Err.Description: Err.Clear: Exit Sub
  
  Set C = New cConverter
  C.ConvertDatabase aCnn, sCnn
  If Err Then MsgBox Err.Description: Err.Clear: Exit Sub
  lProgress.Caption = "Table-Schemas created, Table-Data transferred!"
  
  C.ConvertIndexes aCnn, sCnn
  If Err Then MsgBox Err.Description
  lProgress.Caption = "Index-Import finished!"
  
  'Now we add the (prior) manually created SQLite-Views from their Def-Files
  'For i = 1 To 10
  '  Err.Clear
  '  sCnn.Execute GetTextFromFile(App.Path & "\NWind.db View-Defs\" & i & ".txt")
  '  If Err Then MsgBox Err.Description
  'Next i
  'lProgress.Caption = "Manually defined View-Pendants created!"

  'finally we create a backup-table from 'employees' inside the new
  'created SQLite-Database (used inside the DC/DataBinding-Demos)
  Err.Clear
  'sCnn.BeginTrans
  '  sCnn.Execute "Create Table EmployeesBackup As Select * from Employees"
  sCnn.CommitTrans
  If Err Then MsgBox Err.Description: Err.Clear

  Set C = Nothing
End Sub

Private Sub Command1_Click()
  Dim sFILENAME As String
  Dim sPath As String
  Dim sRECENTFILE As String
  Dim sFILTER As String

  sPath = App.Path
  sFILTER = "Access mdb" & vbNullChar & "*.mdb" & vbNullChar & "DBF" & vbNullChar & "*.DBF"
  sFILENAME = FileOpen(Me, sFILTER, 1, sRECENTFILE, "ini", sPath, "Open Ini File")
  If Len(sFILENAME) = 0 Then
    Text1.tEXT = ""
    Exit Sub
  End If
  Text1.tEXT = sFILENAME
End Sub
Private Sub Command2_Click()
  Dim sFILENAME As String
  Dim sPath As String
  Dim sRECENTFILE As String
  Dim sFILTER As String


  sPath = App.Path
  sFILTER = "Access mdb" & vbNullChar & "*.mdb" & vbNullChar & "Sqlite" & vbNullChar & "*.sqlite"
  sFILENAME = FileOpen(Me, sFILTER, 1, sRECENTFILE, "ini", sPath, "Open Ini File")
  If Len(sFILENAME) = 0 Then
    Text2.tEXT = ""
    Exit Sub
  End If
  Text2.tEXT = sFILENAME
  
End Sub
Private Sub Command3_Click()
  If Len(Text1.tEXT) = 0 Or Len(Text2.tEXT) = 0 Then
    Alert "Origem ou Destino Nao Preenchido"
    Exit Sub
  End If
  corrige Text1.tEXT, Text2.tEXT, False
End Sub
Private Sub Command4_Click()
  If Len(Text1.tEXT) = 0 Or Len(Text2.tEXT) = 0 Then
    Alert "Origem ou Destino Nao Preenchido"
    Exit Sub
  End If
  corrige Text1.tEXT, Text2.tEXT, True
End Sub
Private Sub corrige(ByVal cORIGEM As String, ByVal cDESTINO As String, Optional ByVal lGRAVA As Boolean = False)
  Dim WrkSpace As DAO.Workspace
  Dim Baza1 As DAO.Database
  Dim Baza2 As DAO.Database
  Dim rec1 As DAO.Recordset
  Dim rec2 As DAO.Recordset
  Dim NewTable As DAO.TableDef
  Dim oTableDef As DAO.TableDef
  Dim fldTemp As DAO.Field
  Dim NewFld As DAO.Field
  Dim N, m As Integer
  Dim cTabela As String
  Dim T_Ime, R_Ime, nekej, r_type, R_size As Variant
  Dim oFieldDef As Variant


  Dim i As Integer
  Dim x As Integer
  Dim cARQ As String

  On Error GoTo errhandler
  Set WrkSpace = DBEngine.CreateWorkspace("Compare", "Admin", "")
  Set Baza1 = WrkSpace.OpenDatabase(Text1.tEXT)
  Set Baza2 = WrkSpace.OpenDatabase(Text2.tEXT)
  ''On Error Resume Next
  Text3.tEXT = ""
  For N = 0 To Baza1.TableDefs.Count - 1
    If Baza1.TableDefs(N).Properties(5) = 0 Then
      nERRO = 0
      DoEvents
      T_Ime = Baza1.TableDefs(N).Name
      Set rec1 = Baza1.OpenRecordset(T_Ime)
      Set rec2 = Baza2.OpenRecordset(T_Ime)
      If nERRO = 3078 Then
        Text3.tEXT = Text3.tEXT & "Falta Tabela: " & T_Ime & vbNewLine
        If lGRAVA Then
          Set NewTable = Baza2.CreateTableDef(T_Ime)
          With NewTable
            For m = 0 To rec1.Fields.Count - 1
              R_Ime = rec1.Fields(m).Name
              r_type = rec1.Fields(m).Type
              R_size = rec1.Fields(m).Size
              Set fldTemp = NewTable.CreateField(R_Ime, r_type, R_size)
              If r_type = dbText Or r_type = dbMemo Then
                fldTemp.AllowZeroLength = True
              End If
              .Fields.Append fldTemp
            Next m
          End With
          Baza2.TableDefs.Append NewTable
        End If
      Else
        For m = 0 To rec1.Fields.Count - 1
          nERRO = 0
          R_Ime = rec1.Fields(m).Name
          nekej = rec2.Fields(R_Ime).Name
          If nERRO = 3265 Then
            Text3.tEXT = Text3.tEXT & "Falta Campo: " & T_Ime & "." & R_Ime & vbNewLine
            If lGRAVA Then
              Set NewTable = Baza2.TableDefs(T_Ime)
              rec2.Close
              With NewTable
                R_Ime = rec1.Fields(m).Name
                r_type = rec1.Fields(m).Type
                R_size = rec1.Fields(m).Size
                Set fldTemp = NewTable.CreateField(R_Ime, r_type, R_size)
                If r_type = dbText Or r_type = dbMemo Then
                  fldTemp.AllowZeroLength = True
                End If
                .Fields.Append fldTemp
              End With
              Set rec2 = Baza2.OpenRecordset(T_Ime)
            End If
          End If
        Next m
      End If
    End If
  Next N
  If lGRAVA Then
    rec2.Close
    rec1.Close
  End If
  Baza2.Close
  Baza1.Close
  WrkSpace.Close
  If Text3.tEXT = "" Then
    Text3.tEXT = "Arquivos sem diferencas"
  End If

  Exit Sub
errhandler:
  Select Case Err.Number
  Case 3422, 3078, 3265
    nERRO = FixInt(Err.Number)
    Resume Next
  Case 3219, 91
    Resume Next
  Case Else
    SayErro "Dcompare"
    Exit Sub
  End Select
End Sub

Private Sub Command5_Click()
  Dim sFILTER, sFILENAME As String
  sFILTER = "Arquivos Access (*.MDB)" & vbNullChar & "*.MDB" & vbNullChar & "Todos Arquivo" & vbNullChar & "*.*"
  sFILENAME = FileSave(Me, sFILTER, 1, "MDB", "", "", "Criar Access Como")
  If Not Len(sFILENAME) = 0 Then
    If FileExist(sFILENAME) Then
      ''nAO GRAVA EM CIMA
    Else
      AdoNewTable sFILENAME, False, 5
      Text2.tEXT = sFILENAME
    End If
  End If
End Sub

Private Sub Encerrar_Click()
  Unload Me
End Sub

Private Sub Form_Load()
'Center Me
  CenterFormToScreen Me
End Sub
