VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} f_Calendario 
   Caption         =   "Calendário"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3600
   OleObjectBlob   =   "f_Calendario.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "f_Calendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub UserForm_Initialize()
    
    '---A data inicial é a data que está na célula A1 da Plan1
    'If CInt(dtDate) = 0 Then dtDate = Date
    
    '---Escreve a data de hoje na Label no rodapé do formulário
    lblHoje.Caption = "Hoje: " & Format(Date, sMascaraData)
    
    '---Calcula a quantidade de dias desde o ano 0 (zero) até a data base
    '---e atribui esse valor ao tamanho do SpinButton
    'sb.Value = Year(dtDate) * 12 + Month(dtDate)
    
    
    txtAno.Text = Year(dtDate)
    scrMes.Value = Month(dtDate)
    
    With spbAno
        .Value = Year(dtDate)
        .Max = Year(dtDate) + 1
        .Min = Year(dtDate) - 1
    End With
    
    Select Case scrMes.Value
        Case 1: lblMes.Caption = "Janeiro"
        Case 2: lblMes.Caption = "Fereveiro"
        Case 3: lblMes.Caption = "Março"
        Case 4: lblMes.Caption = "Abril"
        Case 5: lblMes.Caption = "Maio"
        Case 6: lblMes.Caption = "Junho"
        Case 7: lblMes.Caption = "Julho"
        Case 8: lblMes.Caption = "Agosto"
        Case 9: lblMes.Caption = "Setembro"
        Case 10: lblMes.Caption = "Outubro"
        Case 11: lblMes.Caption = "Novembro"
        Case 12: lblMes.Caption = "Dezembro"
    End Select

End Sub
Private Sub scrMes_Change()

    Select Case scrMes.Value
        
        Case 0:
            txtAno.Text = txtAno.Text - 1
            scrMes.Value = 12
            
        Case 13:
            txtAno.Text = txtAno.Text + 1
            scrMes.Value = 1
            
        Case 1: lblMes.Caption = "Janeiro"
        Case 2: lblMes.Caption = "Fereveiro"
        Case 3: lblMes.Caption = "Março"
        Case 4: lblMes.Caption = "Abril"
        Case 5: lblMes.Caption = "Maio"
        Case 6: lblMes.Caption = "Junho"
        Case 7: lblMes.Caption = "Julho"
        Case 8: lblMes.Caption = "Agosto"
        Case 9: lblMes.Caption = "Setembro"
        Case 10: lblMes.Caption = "Outubro"
        Case 11: lblMes.Caption = "Novembro"
        Case 12: lblMes.Caption = "Dezembro"
    End Select
    AtualizarDias DateSerial(txtAno.Text, scrMes.Value, 1)
End Sub
Private Sub spbAno_Change()
    txtAno.Text = spbAno.Value
    spbAno.Max = spbAno.Value + 1
    spbAno.Min = spbAno.Value - 1
    AtualizarDias DateSerial(txtAno.Text, scrMes.Value, 1)
End Sub
Private Sub lblHoje_Click()
    '---Quando se clica no Label do dia atual, o calendário atualiza-se
    '---para o mês atual.
    '---O modo de cálculo do mês em questão é o número de meses.
    '---Como um ano possui 12 meses, o valor da ScrollBar é o número
    '---total de meses:
    dtDate = Date
    
    With spbAno
        .Max = Year(Date) + 1
        .Min = Year(Date) - 1
        .Value = Year(Date)
    End With

    scrMes.Value = Month(Date)
    
End Sub
Private Sub spb_Change()
    AtualizarDias DateSerial(spbAno.Value, scrMes.Value, 1)
End Sub

Private Sub AtualizarDias(dt As Date)
    '--------------------------------------------------------------'
    '---Rotina que atualiza todos os dias (botões) do calendário---'
    '--------------------------------------------------------------'
    
    ' Declara variáveis
    Dim l As Long ' linha
    Dim c As Long '---coluna
    Dim dtDia As Date
    Dim Ctrl As control
    
    For l = 1 To 6 '---Laço que percorre as Linhas do calendário
        For c = 1 To 7 '---Laço que percorre as Colunas do calendário
            
            '---Seta o botão que receberá o rótulo do dia correspondente
            Set Ctrl = Controls("l" & l & "c" & c)
            
            'O entendimento da linha abaixo é fundamental para entender como todos os
            'labels foram povoados:
            dtDia = DateSerial(Year(dt), Month(dt), (l - 1) * 7 + c - Weekday(dt) + 1)

            'Ctrl.Caption = Format(Day(dtDia), "00")
            Ctrl.Caption = Day(dtDia)
            Ctrl.Tag = dtDia
            If Ctrl.Tag = dtDate Then
                Ctrl.SetFocus
            End If
            
            
            'Dias de um mês diferente do mês visualizado ficarão na cor cinza claro:
            If Month(dtDia) <> Month(dt) Then
                Ctrl.ForeColor = &HC0C0C0
            Else
                Ctrl.ForeColor = &H800000
            End If
            
            'Realçar dia atual presente, caso esteja visível no calendário:
            If dtDia = Date Then
                Ctrl.ForeColor = &HFF&
            End If

        Next c
        Set Ctrl = Nothing
    Next l
End Sub


' Tratamento de navegação pelos botões com as setas
Private Sub l1c7_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 39 Then l2c1.SetFocus
End Sub
Private Sub l2c7_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 39 Then l3c1.SetFocus
End Sub
Private Sub l3c7_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 39 Then l4c1.SetFocus
End Sub
Private Sub l4c7_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 39 Then l5c1.SetFocus
End Sub
Private Sub l5c7_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 39 Then l6c1.SetFocus
End Sub
Private Sub l6c1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 37 Then l5c7.SetFocus
End Sub
Private Sub l5c1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 37 Then l4c7.SetFocus
End Sub
Private Sub l4c1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 37 Then l3c7.SetFocus
End Sub
Private Sub l3c1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 37 Then l2c7.SetFocus
End Sub
Private Sub l2c1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 37 Then l1c7.SetFocus
End Sub
Private Sub l1c1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 37 Then
        scrMes.Value = scrMes.Value - 1
        l6c7.SetFocus
    End If
End Sub
Private Sub l6c7_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 39 Then
        scrMes.Value = scrMes.Value + 1
        l1c1.SetFocus
    End If
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    '---Impede que se dê Unload no formulário, caso contrário a linha que testa
    '---frm.Tag na linha seguinte do módulo mCalendario dará erro, pois o objeto
    '---deixará de existir. Ao invés de dar Unload, usa-se Hide para o objeto
    '---continuar a existir na memória.
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Hide
    End If
End Sub
