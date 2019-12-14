Attribute VB_Name = "m_Ribbon"
'Option Private Module
Public Myribbon As IRibbonUI
Sub OnActionButton(control As IRibbonControl)

    If Conecta() = True Then
        Select Case control.ID
            Case "btnPedreiros": fPedreiros.Show
            Case "btnClientes": fClientes.Show
            Case "btnFornecedores": fFornecedores.Show
            Case "btnProdutos": fProdutos.Show
            Case "btnObras": fObras.Show
            Case "btnCompras": fCompras.Show
            Case "btnDicionarioDados": Call AtualizaBD
            'Case "btnOrcamentos": fOrcamentos.Show
            Case Else: MsgBox "Botão ainda não implementado", vbInformation
        End Select
    End If

End Sub

'Callback for customUI.onLoad
Sub ribbonLoaded(ribbon As IRibbonUI)
    Stop
    Set Myribbon = ribbon
End Sub


'Callback for DynamicMenu getContent
Sub dyMenuImportacoes(control As IRibbonControl, ByRef returnedVal)
'   This procedure is executed whenever a sheet is activated
'   (See the Worksheet_Activate procedure in ThisWorkbook)
    
    Dim XMLcode As String
    
'   Read the XML markup from the active sheet
    XMLcode = "<menu xmlns=" & Chr(34) & "http://schemas.microsoft.com/office/2006/01/customui" & Chr(34)
    XMLcode = XMLcode & " >"
    XMLcode = XMLcode & "<button id=" & Chr(34) & "bTransConta" & Chr(34) & " image=" & Chr(34) & "money99" & Chr(34)
    XMLcode = XMLcode & " label=" & Chr(34) & "Money99: Transações da conta" & Chr(34)
    XMLcode = XMLcode & " onAction=" & Chr(34) & "ActionDyMenuImportacoes" & Chr(34) & " />"
    XMLcode = XMLcode & "<button id=" & Chr(34) & "bSaldos" & Chr(34) & " image=" & Chr(34) & "money99" & Chr(34)
    XMLcode = XMLcode & " label=" & Chr(34) & "Money99: Saldo das contas" & Chr(34)
    XMLcode = XMLcode & " onAction=" & Chr(34) & "ActionDyMenuImportacoes" & Chr(34) & " />"
    XMLcode = XMLcode & "<button id=" & Chr(34) & "bBradescoCC" & Chr(34) & " image=" & Chr(34) & "bradesco" & Chr(34)
    XMLcode = XMLcode & " label=" & Chr(34) & "Bradesco: Extrato da conta corrente" & Chr(34)
    XMLcode = XMLcode & " onAction=" & Chr(34) & "ActionDyMenuImportacoes" & Chr(34) & " />"
    XMLcode = XMLcode & "<button id=" & Chr(34) & "bSantanderFatura" & Chr(34) & " image=" & Chr(34) & "santander" & Chr(34)
    XMLcode = XMLcode & " label=" & Chr(34) & "Santander: Fatura de cartão" & Chr(34)
    XMLcode = XMLcode & " onAction=" & Chr(34) & "ActionDyMenuImportacoes" & Chr(34) & " />"
    XMLcode = XMLcode & "</menu>"

    returnedVal = XMLcode
    
End Sub

Sub UpdateDynamicRibbon()
'   Invalidate the ribbon to force a call to dynamicMenuContent
    On Error Resume Next
    Myribbon.Invalidate
    If Err.Number <> 0 Then
        'MsgBox "Lost the Ribbon object. Save and reload."
    End If
End Sub

Sub ActionDyMenuImportacoes(control As IRibbonControl)
'   Executed when Sheet1 is active
    If Conecta() = True Then
        Select Case control.ID
            Case "bTransConta": fImportaTransacoesM99.Show
            'Case "bSaldos": f_import02.Show
            'Case "bBradescoCC": f_import03.Show
            'Case "bSantanderFatura": Call f_import04.Show
            Case Else: MsgBox "Botão ainda não implementado", vbInformation
        End Select
    End If
End Sub
Sub dyMenuOutrosCadastros(control As IRibbonControl, ByRef returnedVal)
'   This procedure is executed whenever a sheet is activated
'   (See the Worksheet_Activate procedure in ThisWorkbook)
    
    Dim XMLcode As String
    
'   Read the XML markup from the active sheet
    XMLcode = "<menu xmlns=" & Chr(34) & "http://schemas.microsoft.com/office/2006/01/customui" & Chr(34)
    XMLcode = XMLcode & " >"
    
'    XMLcode = XMLcode & "<button id=" & Chr(34) & "bBairros" & Chr(34) & " imageMso=" & Chr(34) & "OpenStartPage" & Chr(34)
'    XMLcode = XMLcode & " label=" & Chr(34) & "Bairros" & Chr(34)
'    XMLcode = XMLcode & " onAction=" & Chr(34) & "ActionDyMenuOutrosCadastros" & Chr(34) & " />"
    
    XMLcode = XMLcode & "<button id=" & Chr(34) & "bTiposObra" & Chr(34) & " imageMso=" & Chr(34) & "OpenStartPage" & Chr(34)
    XMLcode = XMLcode & " label=" & Chr(34) & "Tipos de obra" & Chr(34)
    XMLcode = XMLcode & " onAction=" & Chr(34) & "ActionDyMenuOutrosCadastros" & Chr(34) & " />"
    
    XMLcode = XMLcode & "<button id=" & Chr(34) & "bContas" & Chr(34) & " image=" & Chr(34) & "Contas" & Chr(34)
    XMLcode = XMLcode & " label=" & Chr(34) & "Contas" & Chr(34)
    XMLcode = XMLcode & " onAction=" & Chr(34) & "ActionDyMenuOutrosCadastros" & Chr(34) & " />"
    
    XMLcode = XMLcode & "</menu>"

    returnedVal = XMLcode
    
End Sub
Sub ActionDyMenuOutrosCadastros(control As IRibbonControl)
'   Executed when Sheet1 is active
    If Conecta() = True Then
        Select Case control.ID
            'Case "bBairros": fBa.Show
            Case "bTiposObra": fTiposObra.Show
            Case "bContas": fContas.Show
            Case Else: MsgBox "Botão ainda não implementado", vbInformation
        End Select
    End If
End Sub
