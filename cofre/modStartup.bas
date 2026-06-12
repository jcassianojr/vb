Attribute VB_Name = "modstartup"
Option Explicit

Public Sub Main()
    Dim Resposta As VbMsgBoxResult
    
    ' Pergunta ao usuário qual cofre ele deseja abrir
    Resposta = MsgBox("Qual cofre você deseja gerenciar?" & vbCrLf & vbCrLf & _
                      "[ Sim ] - Cofre de Bancos de Dados" & vbCrLf & _
                      "[ Não ] - Cofre de Contas de E-mail", _
                      vbQuestion + vbYesNoCancel + vbDefaultButton1, "Selecionar Módulo")
                      
    Select Case Resposta
        Case vbYes
            ' Carrega e exibe o formulário de Bancos de Dados
            Frmbancos.Show
            
        Case vbNo
            ' Carrega e exibe o formulário de E-mails
            frmEmail.Show
            
        Case vbCancel
            ' Encerra a aplicação sem abrir nenhum formulário
            End
    End Select
End Sub
