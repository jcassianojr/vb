Attribute VB_Name = "ImagensMain"
Public ZGRP As Variant
Public iImage As Long

Sub Main()
Dim CCOMANDO As String
Dim nNUMERO As Integer

    'InitXPStyles
    '01%[CONN]Provider=MSDASQL.1;Persist Security Info=False;Data Source=ol_logix
    EnabledVisualStyles
    zMES = Month(Date)
    zANO = Year(Date)
    zEMPRESA = 1
    
    ' AdoComandodbf GeracArq("c:\develop\modelos\banco\", "JETFOX"), "BA01", "SET DATE DD/MM/YY"
    
    
    'AdoComandodbf GeracArq("c:\develop\modelos\banco\", "JETFOX"), "BA01", "SET DATE BRITISH"
    
     'SET DATE BRITISH'
 
    
   ' CCOMANDO = "insert into mail (numero,erro,data,hora,de,destino,assunto,texto) values (1,'erro','','12:00','de','destino','assunto','texto')"

'     nNUMERO = FixInt(PegMAXSQL(GeracArq("c:\develop\modelos\mail\", "JETFOX"), "MAIL", "NUMERO", 0)) + 1

'nNUMERO = FixInt(PegMAXSQL(GeracArq("c:\develop\modelos\banco\", "JETFOX"), "BA01", "NUMERO", 0)) + 1


  '  CCOMANDO = "insert into mail (numero) values (" + Str(nNUMERO) + ")"

   ' Adodbfteste GeracArq("c:\develop\modelos\mail\", "JETFOX"), "mail", CCOMANDO
    
    
 '   IncluiSQL GeracArq("c:\develop\modelos\mail\", "JETFOX"), "SELECT * FROM MAIL WHERE NUMERO=" & nNUMERO, 3, Array("NUMERO", "DATA", "HORA"), Array(nNUMERO, Date, FixStr(Time)), False, False
    
 '    IncluiSQL GeracArq("c:\develop\modelos\banco\", "JETFOX"), "SELECT * FROM BA01 WHERE NUMERO=" & nNUMERO, 1, Array("NUMERO"), Array(nNUMERO), False, False
    
   
    escIMG.Show
End Sub

Public Function gravalog(Optional ByVal e01, Optional ByVal e02, Optional ByVal e03)
End Function

