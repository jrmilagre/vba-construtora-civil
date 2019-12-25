VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fLancamentosMod2 
   Caption         =   ":: Lançamentos (modelo 2) ::"
   ClientHeight    =   9465
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13560
   OleObjectBlob   =   "fLancamentosMod2.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fLancamentosMod2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private oFuncionario        As New cFuncionario
Private oEmpresa            As New cEmpresa
Private oTabelaPreco        As New cTabelaPreco
Private oLancamento         As New cLancamento
Private oExame              As New cExame
Private oMedico             As New cMedico
Private myRst               As ADODB.RecordSet
Private lPagina             As Long
Private oControles          As New Collection ' Para eventos de campos
Private vbResposta          As VbMsgBoxResult

Private Sub UserForm_Initialize()
    
    Call cbbFltEmpresasPopular
    Call cbbFltFuncionarioPopular
    Call cbbFltStatusPopular
    Call ComboBoxPopularFuncionarios
    Call ComboBoxPopularEmpresas
    Call ComboBoxPopularTiposExame
    Call ComboBoxPopularStatus
    Call ComboBoxPopularMedicos
    Call ComboBoxPopularTabelas
    Call ComboBoxPopularExames
    
    Call Campos("Desabilitar")
    Call EventosCampos
    
    Set myRst = New ADODB.RecordSet
    Set myRst = oLancamento.RetornaLancamentos(cbbFltEmpresa.List(cbbFltEmpresa.ListIndex, 1), cbbFltFuncionario.List(cbbFltFuncionario.ListIndex, 1), cbbFltStatus.List(cbbFltStatus.ListIndex, 1))

    With scrPagina
        .Min = 1
        .Max = myRst.PageCount
    End With
    
    lPagina = myRst.PageCount
    
    myRst.AbsolutePage = myRst.PageCount
    
    scrPagina.Value = lPagina
    
    Call ListBoxPrincipalPopular(scrPagina.Value)
    
    btnCancelar.Visible = False: btnConfirmar.Visible = False
    btnAlterar.Enabled = False: btnExcluir.Enabled = False
    scrPagina.Enabled = False
    MultiPage1.Value = 0
    
    btnExInclui.Visible = False
    btnExAltera.Visible = False
    btnExExclui.Visible = False
    btnExConfirmar.Visible = False
    btnExCancelar.Visible = False
    
End Sub
Private Sub txbLiberacao_AfterUpdate()
    If txbLiberacao.Text = "" Then
        cbbStatus.Text = "LIBERAR"
    Else
        cbbStatus.Text = "LIBERADO"
    End If
End Sub
Private Sub cbbStatus_Change()
    If cbbStatus.Text = "LIBERAR" Then
        txbLiberacao.Text = ""
    Else
        txbLiberacao.Text = Date
    End If
End Sub
Private Sub btnPromoveLiberar_Click()

    Dim i As Integer
    Dim c As Integer
    Dim vbResposta As VbMsgBoxResult
    Dim sMensagem As String
    
    c = 0
    
    For i = 0 To lstPrincipal.ListCount - 1
        If lstPrincipal.Selected(i) = True Then
            c = c + 1
        End If
    Next i
    
    If c > 0 Then
        If c = 1 Then
            sMensagem = "Deseja realmente promover o lançamento para status 'LIBERADO'?"
        Else
            sMensagem = "Deseja realmente promover os lançamentos para status 'LIBERADO'?"
        End If
        
        vbResposta = MsgBox(sMensagem, vbQuestion + vbYesNo, "Pergunta")
        
        If vbResposta = vbYes Then
            For i = 0 To lstPrincipal.ListCount - 1
                If lstPrincipal.Selected(i) = True Then
                    oLancamento.PromoveParaLiberado CLng(lstPrincipal.List(i, 0))
                End If
            Next i
            
            MsgBox "Alteração realizada com sucesso!"
        End If
        
        chbVarios.Value = False: chbVarios.Visible = False
        btnPromoveLiberar.Visible = False: lblPromoveLiberar.Visible = False
        Call btnFiltrar_Click
    Else
        MsgBox "Nenhum lançamento foi selecionado", vbInformation
        chbVarios.Value = False: chbVarios.Visible = False
        btnPromoveLiberar.Visible = False: lblPromoveLiberar.Visible = False
        Call btnFiltrar_Click
    End If

End Sub
Private Sub btnPromoveLiberado_Click()
    
    Dim i As Integer
    Dim c As Integer
    Dim vbResposta As VbMsgBoxResult
    Dim sMensagem As String
    
    c = 0
    
    For i = 0 To lstPrincipal.ListCount - 1
        If lstPrincipal.Selected(i) = True Then
            c = c + 1
        End If
    Next i
    
    If c > 0 Then
        If c = 1 Then
            sMensagem = "Deseja realmente promover o lançamento para status 'EXTRATO'?"
        Else
            sMensagem = "Deseja realmente promover os lançamentos para status 'EXTRATO'?"
        End If
        
        vbResposta = MsgBox(sMensagem, vbQuestion + vbYesNo, "Pergunta")
        
        If vbResposta = vbYes Then
            For i = 0 To lstPrincipal.ListCount - 1
                If lstPrincipal.Selected(i) = True Then
                    oLancamento.PromoveParaExtrato CLng(lstPrincipal.List(i, 0))
                End If
            Next i
            
            MsgBox "Alteração realizada com sucesso!"
        End If
        
        chbVarios.Value = False: chbVarios.Visible = False
        btnPromoveLiberado.Visible = False: lblPromoveLiberado.Visible = False
        Call btnFiltrar_Click
    Else
        MsgBox "Nenhum lançamento foi selecionado", vbInformation
        chbVarios.Value = False: chbVarios.Visible = False
        btnPromoveLiberado.Visible = False: lblPromoveLiberado.Visible = False
        Call btnFiltrar_Click
    End If
    
End Sub
Private Sub btnPromoveExtrato_Click()
    
    Dim i As Integer
    Dim c As Integer
    Dim vbResposta As VbMsgBoxResult
    Dim sMensagem As String
    
    c = 0
    
    For i = 0 To lstPrincipal.ListCount - 1
        If lstPrincipal.Selected(i) = True Then
            c = c + 1
        End If
    Next i
    
    If c > 0 Then
        If c = 1 Then
            sMensagem = "Deseja realmente promover o lançamento para status 'FATURADO'?"
        Else
            sMensagem = "Deseja realmente promover os lançamentos para status 'FATURADO'?"
        End If
        
        vbResposta = MsgBox(sMensagem, vbQuestion + vbYesNo, "Pergunta")
        
        If vbResposta = vbYes Then
            For i = 0 To lstPrincipal.ListCount - 1
                If lstPrincipal.Selected(i) = True Then
                    oLancamento.PromoveParaFaturado CLng(lstPrincipal.List(i, 0))
                End If
            Next i
            
            MsgBox "Alteração realizada com sucesso!"
        End If
        
        chbVarios.Value = False: chbVarios.Visible = False
        btnPromoveExtrato.Visible = False: lblPromoveExtrato.Visible = False
        Call btnFiltrar_Click
    Else
        MsgBox "Nenhum lançamento foi selecionado", vbInformation
        chbVarios.Value = False: chbVarios.Visible = False
        btnPromoveExtrato.Visible = False: lblPromoveExtrato.Visible = False
        Call btnFiltrar_Click
    End If
    
End Sub
Private Sub btnConfirmar_Click()
    
    Dim vbResposta As VBA.VbMsgBoxResult
    Dim sDecisao As String
    Dim i As Integer
    
    sDecisao = Replace(btnConfirmar.Caption, "Confirmar ", "")
    
    If Valida = True Then
        vbResposta = MsgBox("Deseja realmente fazer a " & sDecisao & "?", vbYesNo + vbQuestion, "Pergunta")
        
        If vbResposta = vbYes Then
            If sDecisao = "Exclusão" Then
                oLancamento.ExcluiTotal
            Else
                If sDecisao = "Inclusão" Then
                    oLancamento.Id = oLancamento.ProximoId
                ElseIf sDecisao = "Alteração" Then
                    oLancamento.ExcluiTotal
                End If
                
                For i = 0 To lstExames.ListCount - 1
                    oLancamento.ExameId = CInt(lstExames.List(i, 1))
                    oLancamento.Preco = CCur(lstExames.List(i, 3))
                    oLancamento.TravaPreco = IIf(lstExames.List(i, 4) = "Sim", True, False)
                    oLancamento.Inclui
                Next i
              
            End If
            
            If sDecisao = "Inclusão" Then
                If lstPrincipal.ListCount < myRst.PageSize Then
                    lPagina = Trim(Mid(lblPaginaAtual.Caption, InStr(1, lblPaginaAtual.Caption, "de") + 3, Len(lblPaginaAtual.Caption)))
                Else
                    lPagina = Trim(Mid(lblPaginaAtual.Caption, InStr(1, lblPaginaAtual.Caption, "de") + 3, Len(lblPaginaAtual.Caption))) + 1
                End If
            Else
                lPagina = Trim(Mid(lblPaginaAtual.Caption, InStr(1, lblPaginaAtual.Caption, "de") + 3, Len(lblPaginaAtual.Caption)))
            End If
            
            Set myRst = New ADODB.RecordSet
            Set myRst = oLancamento.RetornaLancamentos(cbbFltEmpresa.List(cbbFltEmpresa.ListIndex, 1), cbbFltFuncionario.List(cbbFltFuncionario.ListIndex, 1), cbbFltStatus.List(cbbFltStatus.ListIndex, 1))
        
            With scrPagina
                .Min = 1
                .Max = myRst.PageCount
                '.Max = lPagina
            End With
            
            'lPagina = myRst.PageCount
            
            myRst.AbsolutePage = myRst.PageCount
            'myRst.AbsolutePage = lPagina
            scrPagina.Value = lPagina
            
            Call ListBoxPrincipalPopular(lPagina)
            
            ' Exibe mensagem de sucesso na decisão tomada (inclusão, alteração ou exclusão do registro).
            MsgBox sDecisao & " realizada com sucesso.", vbInformation, sDecisao & " de registro"
            
            MultiPage1.Value = 0
            
            Call btnCancelar_Click
            
        ElseIf vbResposta = vbNo Then
            Call btnCancelar_Click
        End If
    
    End If
    
    Set oLancamento = Nothing
    Set oLancamento = New cLancamento
    
End Sub
Private Function Valida() As Boolean
    
    Valida = False
    
    If txbData.Text = Empty Then
        MsgBox "Campo 'Data' é obrigatório", vbCritical
        MultiPage1.Value = 1: txbData.SetFocus: Exit Function
    ElseIf cbbFuncionario.ListIndex = -1 Then
        MsgBox "Campo 'Funcionário' é obrigatório", vbCritical
        MultiPage1.Value = 1: cbbFuncionario.SetFocus: Exit Function
    ElseIf cbbTipoExame.ListIndex = -1 Then
        MsgBox "Campo 'Tipo do exame' é obrigatório", vbCritical
        MultiPage1.Value = 1: cbbTipoExame.SetFocus: Exit Function
    ElseIf cbbStatus.ListIndex = -1 Then
        MsgBox "Campo 'Status' é obrigatório", vbCritical
        MultiPage1.Value = 1: cbbStatus.SetFocus: Exit Function
    Else
        If lstExames.ListCount = 0 Then
            MsgBox "Não há exames apontados nesse lançamento"
            MultiPage1.Value = 2: lstExames.SetFocus: Exit Function
            
        Else
            With oLancamento
                .Data = CDate(txbData.Text)
                .FuncionarioId = CLng(cbbFuncionario.List(cbbFuncionario.ListIndex, 1))
                .TipoExame = cbbTipoExame.Text
                .Pendencia = txbPendencia.Text
                .Comentario = txbComentario.Text
                .Status = cbbStatus.Text
                
                If cbbMedico.ListIndex > -1 Then
                    .MedicoId = CInt(cbbMedico.List(cbbMedico.ListIndex, 1))
                Else
                    .MedicoId = 0
                End If
                
                If cbbStatus.Text = "LIBERAR" Then
                    .Liberacao = 0
                Else
                    If txbLiberacao.Text = "" Then
                        MsgBox "Campo 'Data de liberação' é obrigatório", vbCritical
                        MultiPage1.Value = 1: txbLiberacao.SetFocus: Exit Function
                    Else
                        .Liberacao = CDate(txbLiberacao.Text)
                    End If
                End If
                
                If cbbEmpresa.ListIndex = -1 Then
                    MsgBox "Campo 'Empresa' é obrigatório", vbCritical
                    MultiPage1.Value = 1: cbbEmpresa.SetFocus: Exit Function
                Else
                    If cbbFilial.ListIndex = -1 Then
                        Call ComboBoxPopularFiliais
                        cbbFilial.Text = oEmpresa.Filial
                    End If
                    
                    oEmpresa.Carrega2 cbbEmpresa.Text, cbbFilial.Text
                    .EmpresaID = oEmpresa.Id
                End If
                                            
                If oEmpresa.Pagamento <> "P" Then
                    If cbbTabelaPreco.ListIndex = -1 Then
                        .TabelaId = 0
                        'MsgBox "Campo 'Tabela' é obrigatório", vbCritical
                        'MultiPage1.Value = 1: cbbEmpresa.SetFocus: Exit Function
                    Else
                        .TabelaId = CLng(cbbTabelaPreco.List(cbbTabelaPreco.ListIndex, 1))
                    End If
                Else
                    .TabelaId = 0
                End If
                
            End With
            
            Valida = True
        End If
    End If

End Function
Private Sub btnExConfirmar_Click()
    
    Dim sDecisaoLancamento As String
    Dim sDecisaoItem As String
    
    sDecisaoLancamento = Replace(btnConfirmar.Caption, "Confirmar ", "")
    sDecisaoItem = btnExConfirmar.Caption
    
    If sDecisaoItem = "Incluir" Then
    
        If ValidaItem = True Then
            
            With lstExames
                .ColumnCount = 6
                .ColumnWidths = "200pt; 0pt; 60pt; 60pt"
                .Font = "Consolas"
                .AddItem
                .List(.ListCount - 1, 0) = cbbExame.List(cbbExame.ListIndex, 0)
                .List(.ListCount - 1, 1) = cbbExame.List(cbbExame.ListIndex, 1)
                .List(.ListCount - 1, 2) = Space(9 - Len(Format(txbPrecoTabela.Text, "#,##0.00"))) & Format(txbPrecoTabela.Text, "#,##0.00")
                .List(.ListCount - 1, 3) = Space(9 - Len(Format(txbPrecoCobrado.Text, "#,##0.00"))) & Format(txbPrecoCobrado.Text, "#,##0.00")
                .List(.ListCount - 1, 4) = IIf(chbTravar.Value = True, "Sim", "Não")
            End With
            
            Call btnExCancelar_Click
        Else
            MsgBox "Campo(s) não preenchido(s)", vbCritical
        End If
    ElseIf sDecisaoItem = "Alterar" Then
        If ValidaItem = True Then
            With lstExames
                .List(.ListIndex, 0) = cbbExame.List(cbbExame.ListIndex, 0)
                .List(.ListIndex, 1) = cbbExame.List(cbbExame.ListIndex, 1)
                .List(.ListIndex, 2) = Space(9 - Len(Format(txbPrecoTabela.Text, "#,##0.00"))) & Format(txbPrecoTabela.Text, "#,##0.00")
                .List(.ListIndex, 3) = Space(9 - Len(Format(txbPrecoCobrado.Text, "#,##0.00"))) & Format(txbPrecoCobrado.Text, "#,##0.00")
                .List(.ListIndex, 4) = IIf(chbTravar.Value = True, "Sim", "Não")
            End With
            Call btnExCancelar_Click
        Else
            MsgBox "Campo(s) não preenchido(s)", vbCritical
        End If
    ElseIf sDecisaoItem = "Excluir" Then
        lstExames.RemoveItem (lstExames.ListIndex)
        Call btnExCancelar_Click
    End If

End Sub
Private Function ValidaItem() As Boolean
    ValidaItem = False
    If cbbExame.ListIndex = -1 Then
        Exit Function
    ElseIf txbPrecoCobrado.Text = Empty Then
        Exit Function
    Else
        ValidaItem = True
    End If
End Function
Private Sub txbPrecoCobrado_AfterUpdate()
    txbPrecoCobrado.Text = Format(txbPrecoCobrado.Text, "#,##0.00")
End Sub
Private Sub cbbExame_Change()
    
    Dim cPreco As Currency
    Dim idxTabela As Integer
    Dim idxExame As Integer
    
    If cbbExame.ListIndex > -1 Then
        If cbbTabelaPreco.ListIndex = -1 Then
            txbPrecoTabela.Text = "0,00"
            txbPrecoCobrado.Text = "0,00"
        Else
            idxTabela = CInt(cbbTabelaPreco.List(cbbTabelaPreco.ListIndex, 1))
            idxExame = CInt(cbbExame.List(cbbExame.ListIndex, 1))
            cPreco = oTabelaPreco.CarregaPreco(idxTabela, idxExame)
            txbPrecoTabela.Text = Format(cPreco, "#,##0.00")
            txbPrecoCobrado.Text = Format(cPreco, "#,##0.00")
        End If
    End If
    
End Sub
Private Sub btnExInclui_Click()
    btnExConfirmar.Caption = "Incluir"
    lstExames.ListIndex = -1
    cbbExame.Enabled = True: lblExame.Enabled = True: cbbExame.ListIndex = -1
    txbPrecoCobrado.Enabled = True: lblPrecoPraticado.Enabled = True: txbPrecoCobrado.Text = Empty
    txbPrecoTabela.Text = Empty
    chbTravar.Enabled = True
    btnExInclui.Visible = False
    btnExAltera.Visible = False
    btnExExclui.Visible = False
    btnExCancelar.Visible = True
    btnExConfirmar.Visible = True
    lstExames.Enabled = False: lstExames.ForeColor = &H80000010
    btnConfirmar.Enabled = False
    btnCancelar.Enabled = False
End Sub
Private Sub btnExAltera_Click()
    btnExConfirmar.Caption = "Alterar"
    cbbExame.Enabled = True: lblExame.Enabled = True
    txbPrecoCobrado.Enabled = True: lblPrecoPraticado.Enabled = True
    chbTravar.Enabled = True
    btnExInclui.Visible = False
    btnExAltera.Visible = False
    btnExExclui.Visible = False
    btnExCancelar.Visible = True
    btnExConfirmar.Visible = True
    lstExames.Enabled = False: lstExames.ForeColor = &H80000010
    btnConfirmar.Enabled = False
    btnCancelar.Enabled = False
End Sub
Private Sub btnExExclui_Click()
    btnExConfirmar.Caption = "Excluir"
    btnExInclui.Visible = False
    btnExAltera.Visible = False
    btnExExclui.Visible = False
    btnExCancelar.Visible = True
    btnExConfirmar.Visible = True
    lstExames.Enabled = False: lstExames.ForeColor = &H80000010
    btnConfirmar.Enabled = False
    btnCancelar.Enabled = False
End Sub
Private Sub btnExCancelar_Click()
    btnExConfirmar.Caption = "Confirmar"
    lstExames.ListIndex = -1
    cbbExame.ListIndex = -1: cbbExame.Enabled = False: lblExame.Enabled = False
    txbPrecoCobrado.Enabled = False: lblPrecoPraticado.Enabled = False: txbPrecoCobrado.Text = Empty
    txbPrecoTabela.Text = Empty
    chbTravar.Enabled = False
    btnExInclui.Visible = True
    btnExAltera.Visible = True: btnExAltera.Enabled = False
    btnExExclui.Visible = True: btnExExclui.Enabled = False
    btnExCancelar.Visible = False
    btnExConfirmar.Visible = False
    lstExames.Enabled = True: lstExames.ForeColor = &H80000008
    btnConfirmar.Enabled = True
    btnCancelar.Enabled = True
End Sub
Private Sub lstExames_Change()

    If lstExames.ListIndex > -1 And btnExConfirmar.Caption <> "Alterar" Then
        cbbExame.Text = lstExames.List(lstExames.ListIndex, 0)
        txbPrecoTabela.Text = lstExames.List(lstExames.ListIndex, 2)
        txbPrecoCobrado.Text = lstExames.List(lstExames.ListIndex, 3)
        chbTravar.Value = IIf(lstExames.List(lstExames.ListIndex, 4) = "Não", False, True)
        btnExAltera.Enabled = True
        btnExExclui.Enabled = True
    End If

End Sub
Private Sub lstPrincipal_Change()

    If lstPrincipal.ListIndex > -1 And chbVarios.Value = False Then
        btnAlterar.Enabled = True
        btnExcluir.Enabled = True
        
        ' Carrega informações do lançamento
        oLancamento.Carrega (CLng(lstPrincipal.List(lstPrincipal.ListIndex, 0)))
        
        ' Preenche cabeçalho
        lblCabID.Caption = IIf(oLancamento.Id = 0, "", Format(oLancamento.Id, "0000000000"))
        lblCabData.Caption = oLancamento.Data
        lblCabTipoExame.Caption = oLancamento.TipoExame
        
        oFuncionario.Carrega oLancamento.FuncionarioId
        oEmpresa.Carrega oLancamento.EmpresaID
        
        If oLancamento.TabelaId > 0 Then
            oTabelaPreco.Carrega oLancamento.TabelaId
            cbbTabelaPreco.Visible = True
            lblTabela.Visible = True
            cbbTabelaPreco.Text = oTabelaPreco.Tabela
        Else
            cbbTabelaPreco.ListIndex = -1
            'cbbTabelaPreco.Visible = False
            'lblTabela.Visible = False
        End If
        
        lblCabFuncionario.Caption = oFuncionario.Funcionario
        
        ' Preenche campos
        txbData.Text = oLancamento.Data
        cbbFuncionario.Text = oFuncionario.Funcionario
        cbbEmpresa.Text = RTrim(oEmpresa.Empresa)
        cbbFilial.Text = RTrim(oEmpresa.Filial)
        cbbTipoExame.Text = oLancamento.TipoExame
        txbPendencia.Text = oLancamento.Pendencia
        txbLiberacao.Text = IIf(oLancamento.Liberacao = 0, "", oLancamento.Liberacao)
        txbComentario.Text = oLancamento.Comentario
        cbbStatus.Text = oLancamento.Status
        
        If oLancamento.MedicoId = 0 Then
            cbbMedico.ListIndex = -1
        Else
            oMedico.Carrega oLancamento.MedicoId
            cbbMedico.Text = oMedico.Medico
        End If
        
        Call lstExamesPopular
    End If

End Sub
Private Sub lstExamesPopular()

    Dim cPreco As Currency

    sSQL = "SELECT * "
    sSQL = sSQL & "FROM tbl_lancamentos "
    sSQL = sSQL & "WHERE id = " & oLancamento.Id
    
    Set rst = New ADODB.RecordSet
    
    rst.Open sSQL, cnn, adOpenStatic
    
    With lstExames
        .Clear
        .ColumnCount = 5
        .ColumnWidths = "200pt; 0pt; 60pt; 60pt; 40pt;"
        .Font = "Consolas"
        
        Do Until rst.EOF
            .AddItem
            
            oExame.Carrega rst.Fields("exame_id").Value
            
            .List(.ListCount - 1, 0) = oExame.Exame
            .List(.ListCount - 1, 1) = rst.Fields("exame_id").Value
            
            If rst.Fields("tabela_id").Value > 0 Then
                cPreco = oTabelaPreco.CarregaPreco(rst.Fields("tabela_id").Value, rst.Fields("exame_id").Value)
            Else
                cPreco = 0#
            End If
            
            .List(.ListCount - 1, 2) = Space(9 - Len(Format(cPreco, "#,##0.00"))) & Format(cPreco, "#,##0.00")
            .List(.ListCount - 1, 3) = Space(9 - Len(Format(rst.Fields("preco").Value, "#,##0.00"))) & Format(rst.Fields("preco").Value, "#,##0.00")
            .List(.ListCount - 1, 4) = IIf(rst.Fields("trava_preco").Value = True, "Sim", "Não")
            
            rst.MoveNext
        Loop
    End With
    
    Set rst = Nothing
End Sub
Private Sub btnLiberacao_Click()
    dtDatabase = IIf(txbLiberacao.Text = Empty, Date, txbLiberacao.Text)
    txbLiberacao.Text = GetCalendario
End Sub
Private Sub cbbFilial_AfterUpdate()
    If cbbEmpresa.ListIndex > -1 And cbbEmpresa.Text <> "" Then
        If cbbFilial.ListIndex = -1 And cbbFilial.Text <> "" Then
            If oEmpresa.Existe(cbbEmpresa.Text, cbbFilial.Text) = False Then
                vbResposta = MsgBox("Esta Filial não existe. Deseja cadastrá-la?", vbQuestion + vbYesNo)
                
                If vbResposta = vbYes Then
                    With oEmpresa
                        .Empresa = RTrim(cbbEmpresa.Text)
                        .Filial = RTrim(cbbFilial.Text)
                        .TabelaPrecoId = 0
                        .Funcionarios = 0
                        .PrecoPercapto = 0
                        .Pagamento = ""
                        .RazaoSocial = ""
                        .Cnpj = ""
    
                        .Inclui
                    End With
                    Call ComboBoxPopularFiliais
                End If
            End If
        End If
    End If
End Sub
Private Sub cbbEmpresa_AfterUpdate()
    
    If cbbEmpresa.ListIndex > -1 Then
        Call ComboBoxPopularFiliais
        cbbFilial.ListIndex = 0
        oEmpresa.Carrega2 cbbEmpresa.Text, cbbFilial.Text
        If oEmpresa.Pagamento <> "P" Then
            cbbTabelaPreco.Visible = True: lblTabela.Visible = True
            
            If oLancamento.TabelaId = 0 Then
                cbbTabelaPreco.ListIndex = -1
            Else
                oTabelaPreco.Carrega oLancamento.TabelaId
                cbbTabelaPreco.Text = oTabelaPreco.Tabela
            End If
            
            cbbTipoExame.SetFocus
        Else
            cbbTabelaPreco.ListIndex = -1
            cbbTabelaPreco.Visible = False: lblTabela.Visible = False
        End If
    Else
        If cbbEmpresa.Text <> "" Then
            vbResposta = MsgBox("Esta Empresa não existe. Deseja cadastrá-la?", vbQuestion + vbYesNo)
            If vbResposta = vbYes Then
                With oEmpresa
                    .Empresa = RTrim(cbbEmpresa.Text)
                    .Filial = ""
                    .TabelaPrecoId = 0
                    .Funcionarios = 0
                    .PrecoPercapto = 0
                    .Pagamento = ""
                    .RazaoSocial = ""
                    .Cnpj = ""
                    
                    .Inclui
                End With
                Call ComboBoxPopularEmpresas
                Call ComboBoxPopularFiliais
            Else
                cbbEmpresa.ListIndex = -1
            End If
        End If
    End If

End Sub
Private Sub cbbFuncionario_AfterUpdate()
    
    Dim l As Long
    
    If cbbFuncionario.ListIndex > -1 Then
        l = BuscaUltimoLancamento("Funcionário", CLng(cbbFuncionario.List(cbbFuncionario.ListIndex, 1)))
        If l > 0 Then
            oLancamento.Carrega l
            oEmpresa.Carrega oLancamento.EmpresaID
            cbbEmpresa.Text = oEmpresa.Empresa
            Call ComboBoxPopularFiliais
            cbbFilial.Text = oEmpresa.Filial
            If oEmpresa.Pagamento = "T" Then
                cbbTabelaPreco.Visible = True: lblTabela.Visible = True
                oTabelaPreco.Carrega oEmpresa.TabelaPrecoId
                cbbTabelaPreco.Text = oTabelaPreco.Tabela
                'TODO Carregar preços
                'cbbTipoExame.SetFocus
            Else
                cbbTabelaPreco.ListIndex = -1
                'cbbTabelaPreco.Visible = False: lblTabela.Visible = False
            End If
        End If
    Else
        If cbbFuncionario.Text <> "" Then
            vbResposta = MsgBox("Este Funcionário não existe. Deseja cadastrá-lo?", vbQuestion + vbYesNo)
            If vbResposta = vbYes Then
                With oFuncionario
                    .Funcionario = cbbFuncionario.Text
                    .EmpresaID = 0
                    .Ativo = True
                    .Inclui
                End With
                Call ComboBoxPopularFuncionarios
            Else
                cbbFuncionario.ListIndex = -1
            End If
        End If
    End If
End Sub
Private Sub btnIncluir_Click()
    Call PosDecisaoTomada("Inclusão")
    lstPrincipal.ListIndex = -1
    cbbStatus.Text = "LIBERAR"
End Sub
Private Sub btnAlterar_Click()
    Call PosDecisaoTomada("Alteração")
End Sub
Private Sub btnExcluir_Click()
    Call PosDecisaoTomada("Exclusão")
End Sub
Private Sub PosDecisaoTomada(Decisao As String)

    btnCancelar.Visible = True: btnConfirmar.Visible = True
    btnConfirmar.Caption = "Confirmar " & Decisao
    btnCancelar.Caption = "Cancelar " & Decisao
    
    btnIncluir.Visible = False: btnAlterar.Visible = False: btnExcluir.Visible = False
    
    If Decisao = "Inclusão" Then
        Call Campos("Limpar")
    End If
    
    If Decisao <> "Exclusão" Then
        Call Campos("Habilitar")
        
        If MultiPage1.Value = 0 Then
            MultiPage1.Value = 1
        End If
        
        If Decisao = "Inclusão" Then
            txbData.Text = Date
            If MultiPage1.Value = 1 Then
                cbbFuncionario.SetFocus
            End If
        Else
            If MultiPage1.Value = 1 Then
                cbbFuncionario.SetFocus
            End If
        End If
            
    End If
    
    lstPrincipal.Enabled = False
    lstPrincipal.ForeColor = &H80000010
    
    cbbFltEmpresa.Enabled = False: lblFltEmpresa.Enabled = False
    cbbFltFuncionario.Enabled = False: lblFltFuncionario.Enabled = False
    cbbFltStatus.Enabled = False: lblFltStatus.Enabled = False
    btnFiltrar.Enabled = False
    btnPaginaInicial.Enabled = False
    btnPaginaAnterior.Enabled = False
    btnPaginaSeguinte.Enabled = False
    btnPaginaFinal.Enabled = False
    
    btnExInclui.Visible = True
    btnExAltera.Visible = True
    btnExExclui.Visible = True
End Sub
Private Sub btnCancelar_Click()
    
    btnIncluir.Visible = True: btnAlterar.Visible = True: btnExcluir.Visible = True
    btnConfirmar.Visible = False: btnCancelar.Visible = False
    
    Call Campos("Limpar")
    Call Campos("Desabilitar")
    
    btnAlterar.Enabled = False
    btnExcluir.Enabled = False
    btnIncluir.SetFocus
    
    lstPrincipal.Enabled = True
    lstPrincipal.ForeColor = &H80000008
    
    ' Tira a seleção
    If lstPrincipal.ListIndex >= 0 Then lstPrincipal.Selected(lstPrincipal.ListIndex) = False
    
    MultiPage1.Value = 0
    
    Set oLancamento = Nothing
    Set oLancamento = New cLancamento
    
    cbbFltEmpresa.Enabled = True: lblFltEmpresa.Enabled = True
    cbbFltFuncionario.Enabled = True: lblFltFuncionario.Enabled = True
    cbbFltStatus.Enabled = True: lblFltStatus.Enabled = True
    btnFiltrar.Enabled = True
    btnPaginaInicial.Enabled = True
    btnPaginaAnterior.Enabled = True
    btnPaginaSeguinte.Enabled = True
    btnPaginaFinal.Enabled = True
    cbbExame.ListIndex = -1
    txbPrecoTabela.Text = Empty
    txbPrecoCobrado.Text = Empty
    btnExInclui.Visible = False
    btnExAltera.Visible = False
    btnExExclui.Visible = False
    btnExCancelar.Visible = False
    btnExConfirmar.Visible = False
End Sub
Private Sub lstPrincipal_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    MultiPage1.Value = 1
End Sub
Private Sub btnData_Click()
    dtDatabase = IIf(txbData.Text = Empty, Date, txbData.Text)
    txbData.Text = GetCalendario
End Sub
Private Sub btnPaginaFinal_Click()
    
    If scrPagina.Value < myRst.PageCount Then
        myRst.AbsolutePage = myRst.PageCount
        scrPagina.Value = myRst.PageCount
        Call ListBoxPrincipalPopular(myRst.PageCount)
        Call Campos("Limpar")
    End If
    
End Sub
Private Sub btnPaginaInicial_Click()
  
    If scrPagina.Value <> 1 Then
        myRst.AbsolutePage = 1
        scrPagina.Value = 1
        Call ListBoxPrincipalPopular(lPagina)
        Call Campos("Limpar")
    End If

End Sub
Private Sub btnPaginaAnterior_Click()

    If scrPagina.Value > 2 Then
        myRst.AbsolutePage = scrPagina.Value - 1
        scrPagina.Value = scrPagina.Value - 1
        Call ListBoxPrincipalPopular(scrPagina.Value)
        Call Campos("Limpar")
    Else
        If myRst.AbsolutePage = 1 Then
            Exit Sub
        Else
            myRst.AbsolutePage = 1
            scrPagina.Value = 1
            Call Campos("Limpar")
        End If
    End If
    
End Sub
Private Sub btnPaginaSeguinte_Click()
    
    If scrPagina.Value < myRst.PageCount Then
        myRst.AbsolutePage = scrPagina.Value + 1
        scrPagina.Value = scrPagina.Value + 1
        Call ListBoxPrincipalPopular(scrPagina.Value)
        Call Campos("Limpar")
    End If

End Sub
Private Sub scrbar_Change()
    lstExames.TopIndex = scrbar.Value
    Call Campos("Limpar")
End Sub
Private Sub btnFiltrar_Click()
    
    Set myRst = New ADODB.RecordSet
    Set myRst = oLancamento.RetornaLancamentos(cbbFltEmpresa.List(cbbFltEmpresa.ListIndex, 1), cbbFltFuncionario.List(cbbFltFuncionario.ListIndex, 1), cbbFltStatus.List(cbbFltStatus.ListIndex, 1))
    
    If myRst.PageCount > 0 Then
    
        With scrPagina
            .Min = 1
            .Max = myRst.PageCount
        End With
        
        lPagina = myRst.PageCount
        
        myRst.AbsolutePage = myRst.PageCount
        
        scrPagina.Value = lPagina
        
        Call ListBoxPrincipalPopular(scrPagina.Value)
        
        If cbbFltStatus = "LIBERAR" Then
            chbVarios.Visible = True
            btnPromoveLiberar.Visible = True: lblPromoveLiberar.Visible = True
            btnPromoveLiberado.Visible = False: lblPromoveLiberado.Visible = False
            btnPromoveExtrato.Visible = False: lblPromoveExtrato.Visible = False
        ElseIf cbbFltStatus = "LIBERADO" Then
            chbVarios.Visible = True
            btnPromoveLiberar.Visible = False: lblPromoveLiberar.Visible = False
            btnPromoveLiberado.Visible = True: lblPromoveLiberado.Visible = True
            btnPromoveExtrato.Visible = False: lblPromoveExtrato.Visible = False
        ElseIf cbbFltStatus = "EXTRATO" Then
            chbVarios.Visible = True
            btnPromoveLiberar.Visible = False: lblPromoveLiberar.Visible = False
            btnPromoveLiberado.Visible = False: lblPromoveLiberado.Visible = False
            btnPromoveExtrato.Visible = True: lblPromoveExtrato.Visible = True
        Else
            chbVarios.Visible = False
            btnPromoveLiberar.Visible = False: lblPromoveLiberar.Visible = False
            btnPromoveLiberado.Visible = False: lblPromoveLiberado.Visible = False
            btnPromoveExtrato.Visible = False: lblPromoveExtrato.Visible = False
        End If
    Else
        lstPrincipal.Clear
        lblPaginaAtual.Caption = "Não há dados"
    End If
        
    btnCancelar.Visible = False: btnConfirmar.Visible = False
    btnAlterar.Enabled = False: btnExcluir.Enabled = False
    MultiPage1.Value = 0
    Call Campos("Limpar")
    
End Sub
Private Sub ListBoxPrincipalPopular(Pagina As Long)

    Dim lPosicao    As Long
    Dim lCount      As Long
    
    With lstPrincipal
        .Clear
        .ColumnCount = 8 ' Funcionário, ID, Empresa, Filial
        .ColumnWidths = "55pt; 55pt; 160pt; 97pt; 120pt; 50pt; 30pt; 30pt;"
        .Enabled = True
        .Font = "Consolas"
        
        lCount = 1
        
        While Not myRst.EOF = True And lCount <= myRst.PageSize

            .AddItem

            oFuncionario.Carrega myRst.Fields("funcionario_id").Value
            oEmpresa.Carrega myRst.Fields("empresa_id").Value
            
            oLancamento.TipoExame = myRst.Fields("tipo_exame").Value

            .List(.ListCount - 1, 0) = Format(myRst.Fields("id").Value, "0000000000")
            .List(.ListCount - 1, 1) = myRst.Fields("data").Value
            .List(.ListCount - 1, 2) = oFuncionario.Funcionario
            .List(.ListCount - 1, 3) = oLancamento.TipoExame
            .List(.ListCount - 1, 4) = oEmpresa.Empresa & IIf(oEmpresa.Filial = "", "", " : " & oEmpresa.Filial)
            .List(.ListCount - 1, 5) = myRst.Fields("status").Value
            .List(.ListCount - 1, 6) = Space(2 - Len(Format(myRst.Fields("count_exames").Value, "00"))) & Format(myRst.Fields("count_exames").Value, "00")
            .List(.ListCount - 1, 7) = Space(6 - Len(Format(myRst.Fields("sum_preco").Value, "#,##0.00"))) & Format(myRst.Fields("sum_preco").Value, "#,##0.00")

            lCount = lCount + 1
            myRst.MoveNext
            
        Wend

    End With
   
    lblPaginaAtual.Caption = "Página " & Format(scrPagina.Value, "#,##0") & " de " & Format(myRst.PageCount, "#,##0")

End Sub
Private Sub EventosCampos()

    ' Declara variáveis
    Dim oControle As MSForms.control
    Dim oEvento As c_EventosCampo
    Dim sTag As String
    
    ' Laço para percorrer todos os TextBox e atribuir eventos
    ' de acordo com o tipo de cada campo
    For Each oControle In Me.Controls
        
        If TypeName(oControle) = "TextBox" Then
            If Len(oControle.Tag) > 3 Then
                    
                sTag = IIf(oControle.Tag = "", "", Mid(oControle.Tag, 4, Len(oControle.Tag) - 3))
                
                If sTag = "DATE" Then
                    Set oEvento = New c_EventosCampo
                    Set oEvento.cData = oControle
                    oControles.Add oEvento
                ElseIf sTag = "MOEDA" Then
                    Set oEvento = New c_EventosCampo
                    Set oEvento.cMoeda = oControle
                    oControles.Add oEvento
                End If
            End If
            
        End If
    Next

End Sub
Private Sub Campos(Acao As String)
    
    ' Critério de validação
    ' Exemplo: Tag = 'ssMOEDA'
    ' 1º letra -> Se valida o campo (s = sim ou n = não)
    ' 2º letra -> Se limpa (s = sim ou n = não)
    ' 3º letra -> Se bloqueia (s = sim ou n = não)
    ' 4º letra em diante -> Formatação (MOEDA, DATA, etc)

    ' Declara variáveis
    Dim oControle As MSForms.control
    
    ' Laço para percorrer todos os TextBox e atribuir eventos
    ' de acordo com o tipo de cada campo
    If Acao = "Limpar" Then
        For Each oControle In Me.Controls
        
            If Mid(oControle.Tag, 2, 1) = "s" Then
                If TypeName(oControle) = "TextBox" Then
                    oControle.Text = ""
                ElseIf TypeName(oControle) = "ListBox" Then
                    oControle.Clear
                ElseIf TypeName(oControle) = "Label" Then
                    oControle.Caption = ""
                ElseIf TypeName(oControle) = "ComboBox" Then
                    oControle.ListIndex = -1
                ElseIf TypeName(oControle) = "CheckBox" Then
                    oControle.Value = False
                End If
            End If
        Next oControle
    Else
        For Each oControle In Me.Controls
            ' Bloqueia ou desbloqueia os campos necessários
            If Mid(oControle.Tag, 3, 1) = "s" Then
                If Acao = "Habilitar" Then
                    oControle.Enabled = True
                    txbPrecoTabela.Enabled = False
                    lstExames.ForeColor = &H80000008
                ElseIf Acao = "Desabilitar" Then
                    oControle.Enabled = False
                    lstExames.ForeColor = &H80000010
                End If
            End If
        Next oControle
    End If
    
End Sub
Private Sub ComboBoxPopularTabelas()
    
    Dim sTabelaPreco As String
    
    ' Carrega combo Fornecedores
    sSQL = "SELECT id, tabela "
    sSQL = sSQL & "FROM tbl_tabelas_preco "
    sSQL = sSQL & "WHERE deletado = False "
    sSQL = sSQL & "GROUP BY id, tabela ORDER BY tabela"
    
    ' Cria novo objeto recordset
    Set rst = cnn.Execute(sSQL, adLockReadOnly)
    
    With cbbTabelaPreco
        sTabelaPreco = cbbTabelaPreco.Text
        .Clear
        Do Until rst.EOF
            .AddItem
            .List(.ListCount - 1, 0) = rst.Fields("tabela").Value
            .List(.ListCount - 1, 1) = rst.Fields("id").Value
            rst.MoveNext
        Loop
    End With
    
    Set rst = Nothing
    
    If sTabelaPreco = "" Then cbbTabelaPreco.ListIndex = -1 Else cbbTabelaPreco.Text = sTabelaPreco
    
    
End Sub
Private Sub ComboBoxPopularMedicos()
    
    Dim sMedico As String
    
    ' Carrega combo Fornecedores
    sSQL = "SELECT id, medico FROM tbl_medicos ORDER BY medico"
    
    ' Cria novo objeto recordset
    Set rst = New ADODB.RecordSet
    
    ' Atribui resultado da consulta SQL ao recordset
    With rst
        .CursorLocation = adUseServer
        .Open Source:=sSQL, _
              ActiveConnection:=cnn, _
              CursorType:=adOpenStatic, _
              LockType:=adLockOptimistic, _
              Options:=adCmdText
    End With
    
    With cbbMedico
        sMedico = cbbMedico.Text
        .Clear
        Do Until rst.EOF
            .AddItem
            .List(.ListCount - 1, 0) = rst.Fields("medico").Value
            .List(.ListCount - 1, 1) = rst.Fields("id").Value
            rst.MoveNext
        Loop
    End With
    
    Set rst = Nothing
    
    If sMedico = "" Then cbbMedico.ListIndex = -1 Else cbbMedico.Text = sMedico
    
End Sub
Private Sub ComboBoxPopularStatus()
    With cbbStatus
        .AddItem "LIBERAR"
        .AddItem "LIBERADO"
        .AddItem "EXTRATO"
        .AddItem "FATURADO"
    End With
End Sub
Private Sub ComboBoxPopularTiposExame()
    With cbbTipoExame
        .AddItem "ADMISSIONAL"
        .AddItem "PERIÓDICO"
        .AddItem "DEMISSIONAL"
        .AddItem "MUDANÇA DE FUNÇÃO"
        .AddItem "RETORNO AO TRABALHO"
        .AddItem "COMPLEMENTAR"
        .AddItem "AVALIAÇÃO CLÍNICA"
        .AddItem "RETORNO COM MÉDICO"
    End With
End Sub
Private Sub ComboBoxPopularEmpresas()

    sSQL = "SELECT empresa "
    sSQL = sSQL & "FROM tbl_empresas "
    sSQL = sSQL & "WHERE deletado = False "
    sSQL = sSQL & "GROUP BY empresa "
    sSQL = sSQL & "ORDER BY empresa "
    
    Set rst = New ADODB.RecordSet
    
    With rst
        .CursorLocation = adUseServer
        .Open Source:=sSQL, _
              ActiveConnection:=cnn, _
              CursorType:=adOpenStatic, _
              LockType:=adLockOptimistic, _
              Options:=adCmdText
    End With
    
    With cbbEmpresa
        .Clear
        .ColumnCount = 1
        .ColumnWidths = "170pt;"
        Do Until rst.EOF
            .AddItem
            .List(.ListCount - 1, 0) = RTrim(rst.Fields("empresa").Value)
            
            rst.MoveNext
        Loop
    End With
    
    Set rst = Nothing
End Sub
Private Sub ComboBoxPopularFiliais()

    sSQL = "SELECT filial "
    sSQL = sSQL & "FROM tbl_empresas "
    sSQL = sSQL & "WHERE deletado = False and "
    sSQL = sSQL & "empresa = '" & Replace(cbbEmpresa.Text, "'", "`") & "' "
    sSQL = sSQL & "ORDER BY filial "

    
    Set rst = New ADODB.RecordSet
    
    With rst
        .CursorLocation = adUseServer
        .Open Source:=sSQL, _
              ActiveConnection:=cnn, _
              CursorType:=adOpenStatic, _
              LockType:=adLockReadOnly, _
              Options:=adCmdText
    End With
    
    With cbbFilial
        .Clear
        .ColumnCount = 1
        .ColumnWidths = "100pt"
        Do Until rst.EOF
            .AddItem
            .List(.ListCount - 1, 0) = RTrim(rst.Fields("filial").Value)
            
            rst.MoveNext
        Loop
    End With
    
    Set rst = Nothing
End Sub
Private Sub ComboBoxPopularFuncionarios()

    sSQL = "SELECT funcionario, id "
    sSQL = sSQL & "FROM tbl_funcionarios "
    sSQL = sSQL & "WHERE ativo = True AND "
    sSQL = sSQL & "deletado = False "
    sSQL = sSQL & "ORDER BY funcionario "
    
    Set rst = New ADODB.RecordSet
    
    With rst
        .CursorLocation = adUseServer
        .Open Source:=sSQL, _
              ActiveConnection:=cnn, _
              CursorType:=adOpenStatic, _
              LockType:=adLockOptimistic, _
              Options:=adCmdText
    End With
    
    With cbbFuncionario
        .Clear
        .ColumnCount = 2
        .ColumnWidths = "170pt; 0pt;"
        Do Until rst.EOF
            .AddItem
            .List(.ListCount - 1, 0) = rst.Fields("funcionario").Value
            .List(.ListCount - 1, 1) = rst.Fields("id").Value
            
            rst.MoveNext
        Loop
    End With
    
    Set rst = Nothing
End Sub
Private Sub ComboBoxPopularExames()
    
    ' Carrega combo Fornecedores
    sSQL = "SELECT exame, id FROM tbl_exames WHERE deletado = False ORDER BY exame"
    
    ' Cria novo objeto recordset
    Set rst = New ADODB.RecordSet
    
    rst.Open sSQL, cnn, adOpenStatic
    
    With cbbExame
        .Clear
        Do Until rst.EOF
            .AddItem
            .List(.ListCount - 1, 0) = rst.Fields("exame").Value
            .List(.ListCount - 1, 1) = rst.Fields("id").Value
            rst.MoveNext
        Loop
    End With
    
    Set rst = Nothing
    
End Sub
Private Sub cbbFltStatusPopular()

    With cbbFltStatus
        .AddItem
        .List(.ListCount - 1, 0) = "***TODOS***"
        .List(.ListCount - 1, 1) = "%%"
        .AddItem
        .List(.ListCount - 1, 0) = "LIBERAR"
        .List(.ListCount - 1, 1) = "LIBERAR"
        .AddItem
        .List(.ListCount - 1, 0) = "LIBERADO"
        .List(.ListCount - 1, 1) = "LIBERADO"
        .AddItem
        .List(.ListCount - 1, 0) = "EXTRATO"
        .List(.ListCount - 1, 1) = "EXTRATO"
        .AddItem
        .List(.ListCount - 1, 0) = "FATURADO"
        .List(.ListCount - 1, 1) = "FATURADO"
    End With
    
    cbbFltStatus.ListIndex = 0
End Sub
Private Sub cbbFltFuncionarioPopular()

    sSQL = "SELECT funcionario, tbl_funcionarios.id as funcionario_id "
    sSQL = sSQL & "FROM tbl_funcionarios RIGHT JOIN (tbl_lancamentos "
    sSQL = sSQL & "LEFT JOIN tbl_empresas ON tbl_lancamentos.empresa_id = tbl_empresas.id) "
    sSQL = sSQL & "ON tbl_funcionarios.id = tbl_lancamentos.funcionario_id "
    
    If cbbFltEmpresa.ListIndex > 0 Then
        sSQL = sSQL & "WHERE tbl_lancamentos.empresa_id = " & CLng(cbbFltEmpresa.List(cbbFltEmpresa.ListIndex, 1)) & " "
    End If
    
    sSQL = sSQL & "GROUP BY "
    sSQL = sSQL & "tbl_lancamentos.empresa_id, tbl_funcionarios.funcionario, tbl_funcionarios.id "
    
    sSQL = sSQL & "ORDER BY funcionario "
    
    Dim rstFuncionario As New ADODB.RecordSet
    
    rstFuncionario.Open sSQL, cnn
    
    With cbbFltFuncionario
        .Clear
        .ColumnCount = 2
        .ColumnWidths = "170pt; 0pt;"
        .AddItem
        .List(.ListCount - 1, 0) = "***TODOS***"
        .List(.ListCount - 1, 1) = "%%"
        Do Until rstFuncionario.EOF
            .AddItem
            .List(.ListCount - 1, 0) = rstFuncionario.Fields("funcionario").Value
            .List(.ListCount - 1, 1) = rstFuncionario.Fields("funcionario_id").Value
            
            rstFuncionario.MoveNext
        Loop
    End With
    
    Set rstFuncionario = Nothing
    
    On Error Resume Next
    cbbFltFuncionario.ListIndex = 0

End Sub
Private Sub cbbFltEmpresasPopular()

    sSQL = "SELECT tbl_lancamentos.empresa_id as empresa_id, tbl_empresas.empresa, tbl_empresas.filial "
    sSQL = sSQL & "FROM tbl_lancamentos "
    sSQL = sSQL & "LEFT JOIN tbl_empresas ON tbl_lancamentos.empresa_id = tbl_empresas.id "
    sSQL = sSQL & "GROUP BY tbl_lancamentos.empresa_id, tbl_empresas.empresa, tbl_empresas.filial "
    sSQL = sSQL & "ORDER BY tbl_empresas.empresa, tbl_empresas.filial"
    
    Dim rstEmpresas As New ADODB.RecordSet
    
    rstEmpresas.Open sSQL, cnn, adOpenStatic
    
    With cbbFltEmpresa
        .Clear
        .ColumnCount = 4
        .ColumnWidths = "170pt; 0pt; 100pt; 80pt;"
        .AddItem
        .List(.ListCount - 1, 0) = "***TODOS***"
        .List(.ListCount - 1, 1) = "%%"
        Do Until rstEmpresas.EOF
            .AddItem
            .List(.ListCount - 1, 0) = rstEmpresas.Fields("empresa").Value
            .List(.ListCount - 1, 1) = rstEmpresas.Fields("empresa_id").Value
            .List(.ListCount - 1, 2) = rstEmpresas.Fields("filial")
            
            rstEmpresas.MoveNext
        Loop
    End With
    
    Set rstEmpresas = Nothing
    
    On Error Resume Next
    cbbFltEmpresa.ListIndex = 0

End Sub
Private Function BuscaUltimoLancamento(BuscarPor As String, Id As Long) As Long
    
    Dim r As New ADODB.RecordSet
    
    If BuscarPor = "Funcionário" Then
        sSQL = "SELECT id, data, funcionario_id, empresa_id, tabela_id, medico_id "
        sSQL = sSQL & "FROM tbl_lancamentos "
        sSQL = sSQL & "GROUP BY id, data, funcionario_id, empresa_id, tabela_id, medico_id "
        sSQL = sSQL & "HAVING (((funcionario_id) = " & Id & "))"
        sSQL = sSQL & "ORDER BY id DESC "
    End If

    r.Open sSQL, cnn, adOpenStatic
    
    If r.EOF = False Then
        BuscaUltimoLancamento = r.Fields("id").Value
    Else
        BuscaUltimoLancamento = 0
    End If
    
    Set rst = Nothing
End Function
Private Sub chbVarios_Click()
    If chbVarios.Value = True Then
        lstPrincipal.MultiSelect = fmMultiSelectMulti
        Call Campos("Limpar")
        btnIncluir.Visible = False
        btnAlterar.Visible = False
        btnExcluir.Visible = False
    Else
        lstPrincipal.MultiSelect = fmMultiSelectSingle
        btnIncluir.Visible = True
        btnAlterar.Visible = True
        btnExcluir.Visible = True
    End If
End Sub
Private Sub UserForm_Terminate()
    
    ' Destrói objeto da classe cProduto
    'Set oFuncionario = Nothing
    'Set oEmpresa = Nothing
    Call Desconecta
End Sub
