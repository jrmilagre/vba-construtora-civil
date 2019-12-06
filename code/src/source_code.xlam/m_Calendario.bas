Attribute VB_Name = "m_Calendario"
Option Private Module
Option Explicit

'---Para obter uma data do calend�rio na Plan1 e na c�lula A1, basta escrever o seguinte c�digo:
'--- Plan1.Range("A1") = GetCalendario

Public Const sMascaraData   As String = "DD/MM/YYYY"   '---formata��o de datas
Public dtDate               As Date
Dim Botoes()                As New c_Calendario  '---vetor que armazena todos os Label de dia do Calend�rio

Function GetCalendario() As Date
    
    Dim iTotalBotoes As Integer ' Total de bot�es
    Dim Ctrl As control
    Dim frm As f_Calendario     ' Formul�rio
    
    Set frm = New f_Calendario  ' Cria novo objeto setando formul�rio nele
    
    ' Atribui cada um dos Label num elemento do vetor da classe
    For Each Ctrl In frm.Controls
        If Ctrl.name Like "l?c?" Then
            iTotalBotoes = iTotalBotoes + 1
            ReDim Preserve Botoes(1 To iTotalBotoes)
            Set Botoes(iTotalBotoes).btnGrupo = Ctrl
        End If
    Next Ctrl
    
    frm.Show
    
    ' Se a data escolhida for nula ou inv�lida, retorna-se a data atual:
    If IsDate(frm.Tag) Then
        GetCalendario = frm.Tag
    Else
        GetCalendario = dtDate
    End If
    
    Unload frm
    
End Function
    

