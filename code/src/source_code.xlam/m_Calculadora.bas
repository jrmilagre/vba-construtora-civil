Attribute VB_Name = "m_Calculadora"
Option Private Module
Option Explicit

Public ccurVisor            As Currency
Dim aBotoes()               As New c_Calculadora  ' Vetor que armazena todos os botões de dia do Calendário

Public Function GetCalculadora() As Double
    
    ' Declara variáveis
    Dim lTotalBotoes As Long   ' Total de rótulos
    Dim Ctrl As control
    Dim frm As New f_Calculadora      ' Formulário
    
    ' Atribui cada um dos botões em um elemento do vetor da classe
    For Each Ctrl In frm.Controls
        If Ctrl.name Like "btnUsar" Then
            lTotalBotoes = lTotalBotoes + 1
            ReDim Preserve aBotoes(1 To lTotalBotoes)
            Set aBotoes(lTotalBotoes).btnGrupo = Ctrl
        End If
    Next Ctrl
    
    frm.Show
    
    ' Se a data escolhida for nula ou inválida, retorna-se a data atual:
    If IsNumeric(CCur(frm.txbVisor.Text)) Then
        GetCalculadora = CCur(frm.txbVisor.Text)
    Else
        GetCalculadora = ccurVisor
    End If
    
    Unload frm
End Function
