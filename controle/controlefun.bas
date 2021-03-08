Attribute VB_Name = "controlefun"
Public iMU01      As Long
Public iPPAP            As Long
Public iITEM            As Integer
Public cCONJUNTO
Public cDESENHO
Public cARQESC As String
Public Sdb
Public nPF
Public nSEQ
Public nSSQ
Public nORD
Public nPPAP
Public nARQPCS
Public Zsetor           As String
Public Ctipoent As String
Public Ccodcomp As String
Public lEDITAR
Public cCLINOME As String
Public ZENGTIP
Public cARQIMP As String
Public nREV
Public nREVI
Public iMD02
Public nNOVAPOS As Long
Public zPFTIPO As String

Private Sub CapMenu(vForm As Form, bTYPE As Byte, sMnu As String, sCTR As String, i As Integer, bLigado As Boolean, sHK As String)

    With vForm
        Select Case sMnu
        Case "MENU"
            If bTYPE = 0 Then
                vForm.mnuMENU(i).Caption = sCTR
            ElseIf bTYPE = 1 Then
                vForm.mnuMENU(i).Visible = False
            Else
                Load vForm.mnuMENU(i)
            End If
            vForm.mnuMENU(i).Enabled = bLigado
        Case "SUBMENU"
            If bTYPE = 0 And i = 0 Then
                vForm.mnuSUBMENU(i).Caption = sCTR
            ElseIf bTYPE = 1 Then
                vForm.mnuSUBMENU(i).Visible = False
            Else
                If sCTR = "-" Then
                    If bLigado Then
                        Load vForm.mnuSUBMENU(i)
                        vForm.mnuSUBMENU(i).Caption = sCTR
                    End If
                Else
                    Load vForm.mnuSUBMENU(i)
                    vForm.mnuSUBMENU(i).Caption = sCTR
                End If
            End If
            If sCTR = "-" Then
                If bLigado Then
                    vForm.mnuSUBMENU(i).Enabled = bLigado
                End If
            Else
                vForm.mnuSUBMENU(i).Enabled = bLigado
                'vForm.mnuSUBMENU(I).SHORTCUT = vbKeyF1
            End If
        Case "SUBMENU2"
            If bTYPE = 0 And i = 0 Then
                vForm.mnuSUBMENU2(i).Caption = sCTR
            ElseIf bTYPE = 1 Then
                vForm.mnuSUBMENU2(i).Visible = False
            Else
                Load vForm.mnuSUBMENU2(i)
                vForm.mnuSUBMENU2(i).Caption = sCTR
            End If
            vForm.mnuSUBMENU2(i).Enabled = bLigado
        Case "SUBMENU3"
            If bTYPE = 0 And i = 0 Then
                vForm.mnuSUBMENU3(i).Caption = sCTR
            ElseIf bTYPE = 1 Then
                vForm.mnuSUBMENU3(i).Visible = False
            Else
                Load vForm.mnuSUBMENU3(i)
                vForm.mnuSUBMENU3(i).Caption = sCTR
            End If
            vForm.mnuSUBMENU3(i).Enabled = bLigado
        Case "SUBMENU4"
            If bTYPE = 0 And i = 0 Then
                vForm.mnuSUBMENU4(i).Caption = sCTR
            ElseIf bTYPE = 1 Then
                vForm.mnuSUBMENU4(i).Visible = False
            Else
                Load vForm.mnuSUBMENU4(i)
                vForm.mnuSUBMENU4(i).Caption = sCTR
            End If
            vForm.mnuSUBMENU4(i).Enabled = bLigado
        Case "SUBMENU5"
            If bTYPE = 0 And i = 0 Then
                vForm.mnuSUBMENU5(i).Caption = sCTR
            ElseIf bTYPE = 1 Then
                vForm.mnuSUBMENU5(i).Visible = False
            Else
                Load vForm.mnuSUBMENU5(i)
                vForm.mnuSUBMENU5(i).Caption = sCTR
            End If
            vForm.mnuSUBMENU5(i).Enabled = bLigado
        End Select
    End With

End Sub

Public Function FillMenu(sMnu As String) As Boolean
    Dim DB As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim i As Integer
    Dim bLigado As Boolean
    sMnu = UCase(sMnu)
    FillMenu = False
    DB.CommandTimeout = 120
    DB.Open GeracArq(Dbname, , False)
    rs.Open "SELECT * FROM MENU WHERE MENU='mnu" & sMnu & "'", DB, adOpenForwardOnly, adLockReadOnly
    i = 0
    While Not rs.EOF
        bLigado = True
        CapMenu frmPRINCIPAL, 0, sMnu, rs("descricao"), i, bLigado, "" & rs("Hotkey")
        FillMenu = True
        rs.MoveNext
        i = i + 1
    Wend
    rs.Close
    DB.Close
End Function

Public Function DizerBarra(Optional ByVal cMES As String = "")
    frmPRINCIPAL.StatusBar1.Panels(7).tEXT = cMES
End Function

