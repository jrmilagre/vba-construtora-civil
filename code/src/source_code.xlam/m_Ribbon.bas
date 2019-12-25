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
            Case "btnRequisicoes": fRequisicoes.Show
            Case "btnLancamentoRapido": fLancamentosRapidos.Show
            Case "btnDicionarioDados": Call AtualizaBD
            Case "btnBackup": fBackup.Show
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
Sub UpdateDynamicRibbon()
'   Invalidate the ribbon to force a call to dynamicMenuContent
    On Error Resume Next
    Myribbon.Invalidate
    If Err.Number <> 0 Then
        'MsgBox "Lost the Ribbon object. Save and reload."
    End If
End Sub

Sub dyMenuOutrosCadastros(control As IRibbonControl, ByRef returnedVal)
'   This procedure is executed whenever a sheet is activated
'   (See the Worksheet_Activate procedure in ThisWorkbook)
    
    Dim XMLcode As String
    
'   Read the XML markup from the active sheet
    XMLcode = "<menu xmlns=" & Chr(34) & "http://schemas.microsoft.com/office/2006/01/customui" & Chr(34)
    XMLcode = XMLcode & " >"
        
    XMLcode = XMLcode & "<button id=" & Chr(34) & "bContas" & Chr(34) & " image=" & Chr(34) & "Contas" & Chr(34)
    XMLcode = XMLcode & " label=" & Chr(34) & "Contas" & Chr(34)
    XMLcode = XMLcode & " onAction=" & Chr(34) & "ActionDyMenuOutrosCadastros" & Chr(34) & " />"
    
    XMLcode = XMLcode & "<button id=" & Chr(34) & "bCategorias" & Chr(34) & " image=" & Chr(34) & "Categorias" & Chr(34)
    XMLcode = XMLcode & " label=" & Chr(34) & "Categorias" & Chr(34)
    XMLcode = XMLcode & " onAction=" & Chr(34) & "ActionDyMenuOutrosCadastros" & Chr(34) & " />"
    
    XMLcode = XMLcode & "<button id=" & Chr(34) & "bEtapas" & Chr(34) & " imageMso=" & Chr(34) & "OpenStartPage" & Chr(34)
    XMLcode = XMLcode & " label=" & Chr(34) & "Etapas da obra" & Chr(34)
    XMLcode = XMLcode & " onAction=" & Chr(34) & "ActionDyMenuOutrosCadastros" & Chr(34) & " />"

    XMLcode = XMLcode & "<button id=" & Chr(34) & "bTiposObra" & Chr(34) & " imageMso=" & Chr(34) & "OpenStartPage" & Chr(34)
    XMLcode = XMLcode & " label=" & Chr(34) & "Tipos de obra" & Chr(34)
    XMLcode = XMLcode & " onAction=" & Chr(34) & "ActionDyMenuOutrosCadastros" & Chr(34) & " />"
    
    XMLcode = XMLcode & "<button id=" & Chr(34) & "bUnidadesMedida" & Chr(34) & " imageMso=" & Chr(34) & "OpenStartPage" & Chr(34)
    XMLcode = XMLcode & " label=" & Chr(34) & "Unidades de medida" & Chr(34)
    XMLcode = XMLcode & " onAction=" & Chr(34) & "ActionDyMenuOutrosCadastros" & Chr(34) & " />"
    
    XMLcode = XMLcode & "</menu>"

    returnedVal = XMLcode
    
End Sub
Sub ActionDyMenuOutrosCadastros(control As IRibbonControl)
'   Executed when Sheet1 is active
    If Conecta() = True Then
        Select Case control.ID
            'Case "bBairros": fBa.Show
            Case "bContas": fContas.Show
            Case "bCategorias": fCategorias.Show
            Case "bEtapas": fEtapas.Show
            Case "bTiposObra": fTiposObra.Show
            Case "bUnidadesMedida": fUnidadesMedida.Show
            
            Case Else: MsgBox "Botão ainda não implementado", vbInformation
        End Select
    End If
End Sub
Sub dyMenuContasReceber(control As IRibbonControl, ByRef returnedVal)
'   This procedure is executed whenever a sheet is activated
'   (See the Worksheet_Activate procedure in ThisWorkbook)
    
    Dim XMLcode As String
    
'   Read the XML markup from the active sheet
    XMLcode = "<menu xmlns=" & Chr(34) & "http://schemas.microsoft.com/office/2006/01/customui" & Chr(34)
    XMLcode = XMLcode & " >"
        
    XMLcode = XMLcode & "<button id=" & Chr(34) & "bContasReceber" & Chr(34) & " imageMso=" & Chr(34) & "AppointmentColor9" & Chr(34)
    XMLcode = XMLcode & " label=" & Chr(34) & "Contas à receber" & Chr(34)
    XMLcode = XMLcode & " onAction=" & Chr(34) & "ActionDyMenuContasReceber" & Chr(34) & " />"
    
    XMLcode = XMLcode & "<button id=" & Chr(34) & "bRecebimentos" & Chr(34) & " imageMso=" & Chr(34) & "AppointmentColor9" & Chr(34)
    XMLcode = XMLcode & " label=" & Chr(34) & "Recebimentos" & Chr(34)
    XMLcode = XMLcode & " onAction=" & Chr(34) & "ActionDyMenuContasReceber" & Chr(34) & " />"
    
    XMLcode = XMLcode & "</menu>"

    returnedVal = XMLcode
    
End Sub
Sub ActionDyMenuContasReceber(control As IRibbonControl)
'   Executed when Sheet1 is active
    If Conecta() = True Then
        Select Case control.ID
            Case "bContasReceber": fTitulosReceber.Show
            Case "bRecebimentos": fRecebimentos.Show
            'Case "bEtapas": fEtapas.Show
            'Case "bTiposObra": fTiposObra.Show
            'Case "bUnidadesMedida": fUnidadesMedida.Show
            
            Case Else: MsgBox "Botão ainda não implementado", vbInformation
        End Select
    End If
End Sub
Sub dyMenuContasPagar(control As IRibbonControl, ByRef returnedVal)
'   This procedure is executed whenever a sheet is activated
'   (See the Worksheet_Activate procedure in ThisWorkbook)
    
    Dim XMLcode As String
    
'   Read the XML markup from the active sheet
    XMLcode = "<menu xmlns=" & Chr(34) & "http://schemas.microsoft.com/office/2006/01/customui" & Chr(34)
    XMLcode = XMLcode & " >"
        
    XMLcode = XMLcode & "<button id=" & Chr(34) & "bContasPagar" & Chr(34) & " imageMso=" & Chr(34) & "AppointmentColor1" & Chr(34)
    XMLcode = XMLcode & " label=" & Chr(34) & "Contas à pagar" & Chr(34)
    XMLcode = XMLcode & " onAction=" & Chr(34) & "ActionDyMenuContasPagar" & Chr(34) & " />"
    
    XMLcode = XMLcode & "<button id=" & Chr(34) & "bPagamentos" & Chr(34) & " imageMso=" & Chr(34) & "AppointmentColor1" & Chr(34)
    XMLcode = XMLcode & " label=" & Chr(34) & "Pagamentos" & Chr(34)
    XMLcode = XMLcode & " onAction=" & Chr(34) & "ActionDyMenuContasPagar" & Chr(34) & " />"
    
    XMLcode = XMLcode & "</menu>"

    returnedVal = XMLcode
    
End Sub
Sub ActionDyMenuContasPagar(control As IRibbonControl)
'   Executed when Sheet1 is active
    If Conecta() = True Then
        Select Case control.ID
            'Case "bBairros": fBa.Show
            Case "bPagamentos": fPagamentos.Show
            'Case "bEtapas": fEtapas.Show
            'Case "bTiposObra": fTiposObra.Show
            'Case "bUnidadesMedida": fUnidadesMedida.Show
            
            Case Else: MsgBox "Botão ainda não implementado", vbInformation
        End Select
    End If
End Sub
