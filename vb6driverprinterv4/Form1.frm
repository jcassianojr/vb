VERSION 5.00
Object = "{EA6F8A1C-8599-4982-A762-DE32D2ED8173}#1.0#0"; "ucPrinterComboEx.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5592
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   8736
   LinkTopic       =   "Form1"
   ScaleHeight     =   5592
   ScaleWidth      =   8736
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Info printer ocx"
      Height          =   492
      Left            =   5640
      TabIndex        =   4
      Top             =   1800
      Width           =   2292
   End
   Begin ucPrinterComboEx.ucPrinterComboEx ucPrinterComboEx1 
      Height          =   348
      Left            =   480
      TabIndex        =   3
      Top             =   3360
      Width           =   5292
      _ExtentX        =   9334
      _ExtentY        =   614
   End
   Begin VB.ListBox List1 
      Height          =   2160
      Left            =   480
      TabIndex        =   2
      Top             =   840
      Width           =   4452
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Obter listas via profile"
      Height          =   252
      Left            =   5520
      TabIndex        =   1
      Top             =   960
      Width           =   1692
   End
   Begin VB.ComboBox Combo1 
      Height          =   288
      Left            =   480
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   360
      Width           =   4212
   End
   Begin VB.Label Label1 
      Caption         =   "Necessario servico spool printer rodando"
      Height          =   612
      Left            =   600
      TabIndex        =   5
      Top             =   4080
      Width           =   6612
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" _
    (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, _
    ByVal lpReturnedString As String, ByVal nSize As Long) As Long


Public Function ObterListaImpressoras() As String()
    Dim sBuffer As String
    Dim lSize As Long
    Dim lRet As Long
    Dim vPrinters() As String
    Dim aResultado() As String
    Dim i As Long, j As Long
    
    ' Buffer de 16KB
    sBuffer = String(16384, vbNullChar)
    lSize = Len(sBuffer)
    
    ' Busca os dispositivos no registro
    lRet = GetProfileString("Devices", vbNullString, "", sBuffer, lSize)
    
    If lRet <= 0 Then
       MsgBox "O serviço de impressão (Spooler) não está respondendo. Verifique se o serviço está rodando.", vbCritical
       'Exit Function
    End If
    
    If lRet > 0 Then
        ' Ajusta o buffer para o tamanho real retornado pela API
        sBuffer = Left$(sBuffer, lRet)
        
        ' Split retorna o array automaticamente
        vPrinters = Split(sBuffer, vbNullChar)
        
        ' Limpa itens vazios (o último item do Split geralmente é vazio por causa do nulo duplo)
        ReDim aResultado(UBound(vPrinters))
        j = 0
        For i = LBound(vPrinters) To UBound(vPrinters)
            If Len(Trim$(vPrinters(i))) > 0 Then
                aResultado(j) = vPrinters(i)
                j = j + 1
            End If
        Next i
        
        ' Ajusta o tamanho final do array
        If j > 0 Then
            ReDim Preserve aResultado(j - 1)
        Else
            ReDim aResultado(-1 To -1) ' Retorna array vazio
        End If
    Else
        ReDim aResultado(-1 To -1)
    End If
    
    ObterListaImpressoras = aResultado
End Function


Private Sub Command2_Click()
Dim aLista() As String
    Dim i As Long
    
    ' Chama a função que criamos
    aLista = ObterListaImpressoras()
    
    ' Limpa o ListBox antes de carregar
    List1.Clear
    
    ' Verifica se o array tem itens
    If LBound(aLista) <= UBound(aLista) Then
        For i = LBound(aLista) To UBound(aLista)
            List1.AddItem aLista(i)
        Next i
    End If
End Sub

Private Sub Command3_Click()
MsgBox ucPrinterComboEx1.SelectedPrinter

End Sub
