VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fRecebimentos 
   Caption         =   ":: Cadastro de Recebimentos ::"
   ClientHeight    =   9705
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13320
   OleObjectBlob   =   "fRecebimentos.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fRecebimentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private oRecebimento        As New cRecebimento
Private oObra               As New cObra
Private oCliente            As New cCliente
Private oConta              As New cConta
Private oRecebimentoItem    As New cRecebimentoItem
Private oTituloReceber      As New cTituloReceber
Private oContaMovimento     As New cContaMovimento

Private colControles        As New Collection
Private myRst               As New ADODB.Recordset

Private Const sTable As String = "tbl_recebimentos"
Private Const sCampoOrderBy As String = "data"

Private Sub UserForm_Initialize()

    Call cbbObraPopular
    Call cbbContaPopular
    
    Call cbbFltObraPopular
    
    Call EventosCampos
    
    Call btnFiltrar_Click
    
    Call btnCancelar_Click

End Sub
Private Sub UserForm_Terminate()
    
    ' Destrói objeto da classe cProduto
    Set oRecebimento = Nothing
    Set oObra = Nothing
    Set oCliente = Nothing
    Set oConta = Nothing
    Set oRecebimentoItem = Nothing
    Set oTituloReceber = Nothing
    Set oContaMovimento = Nothing
    Set myRst = Nothing
    
    Call Desconecta
    
End Sub

Private Sub lstPrincipalPopular()

    Dim lPosicao    As Long
    Dim lCount      As Long
    
    Set myRst = oRecebimento.Recordset
    
    With scrPagina
        .Min = IIf(myRst.PageCount = 0, 1, myRst.PageCount)
        .Max = myRst.PageCount
    End With
    
    If myRst.PageCount > 0 Then
        myRst.AbsolutePage = myRst.PageCount
    End If
    
    scrPagina.Value = myRst.PageCount
    
    With lstPrincipal
        .Clear
        .ColumnCount = 3 ' Funcionário, ID, Empresa, Filial
        .ColumnWidths = "55pt; 55pt;"
        .Enabled = True
        .Font = "Consolas"
        
        lCount = 1
        
        While Not myRst.EOF = True And lCount <= myRst.PageSize

            .AddItem

            oObra.Carrega myRst.Fields("obra_id").Value

            .List(.ListCount - 1, 0) = Format(myRst.Fields("id").Value, "0000000000")
            .List(.ListCount - 1, 1) = myRst.Fields("data").Value
            .List(.ListCount - 1, 2) = oObra.Bairro & ": " & oCliente.Nome & ": " & oObra.Endereco

            lCount = lCount + 1
            myRst.MoveNext
            
        Wend

    End With
   
    lblPaginaAtual.Caption = "Página " & Format(scrPagina.Value, "#,##0") & " de " & Format(myRst.PageCount, "#,##0")

End Sub
Private Sub EventosCampos()

    ' Declara variáveis
    Dim oControle   As MSForms.control
    Dim oEvento     As c_EventoCampo
    Dim sTag        As String
    Dim iType       As Integer
    Dim bNullable   As Boolean
    Dim sField()    As String

    ' Laço para percorrer todos os TextBox e atribuir eventos
    ' de acordo com o tipo de cada campo
    For Each oControle In Me.Controls

        If Len(oControle.Tag) > 0 Then

            If TypeName(oControle) = "TextBox" Then

                Set oEvento = New c_EventoCampo

                With oEvento
                
                    sField() = Split(oControle.Tag, ".")

                    oControle.ControlTipText = cat.Tables(sField(0)).Columns(sField(1)).Properties("Description").Value

                    .FieldType = cat.Tables(sField(0)).Columns(sField(1)).Type
                    .MaxLength = cat.Tables(sField(0)).Columns(sField(1)).DefinedSize
                    .Nullable = cat.Tables(sField(0)).Columns(sField(1)).Properties("Nullable")

                    Set .cGeneric = oControle

                End With

                colControles.Add oEvento

            End If
        End If
    Next

End Sub
Private Sub btnCancelar_Click()

    btnIncluir.Visible = True: btnExcluir.Visible = True
    btnConfirmar.Visible = False: btnCancelar.Visible = False

    Call Campos("Limpar")
    Call Campos("Desabilitar")

    btnExcluir.Enabled = False
    btnIncluir.SetFocus
    
    MultiPage1.Value = 0

    ' Tira a seleção
    lstPrincipal.ListIndex = -1: lstPrincipal.ForeColor = &H80000008: lstPrincipal.Enabled = True:

End Sub
Private Sub cbbObraPopular()
    
    Dim idx         As Integer
    Dim col         As New Collection
    Dim n           As Variant

    Set col = oObra.Listar("bairro")
    
    idx = cbbObra.ListIndex
    
    With cbbObra
        .Clear
        .ColumnCount = 2
        .ColumnWidths = "180pt; 0pt;"
    End With
    
    
    For Each n In col
        
        oObra.Carrega CLng(n)
        
        oCliente.Carrega oObra.ClienteID
    
        With cbbObra
            .AddItem
            .List(.ListCount - 1, 0) = oObra.Bairro & ": " & oCliente.Nome & ": " & oObra.Endereco
            .List(.ListCount - 1, 1) = oObra.ID
        End With
        
    Next n
    
    cbbObra.ListIndex = idx

End Sub
Private Sub Campos(Acao As String)

    If Acao = "Desabilitar" Then
        txbData.Enabled = False: lblData.Enabled = False: btnData.Enabled = False
        cbbObra.Enabled = False: lblObra.Enabled = False
        
        frmTipo.Enabled = False
        optManual.Enabled = False
        optAutomatico.Enabled = False
        
        frmFormaPagamento.Visible = False
        
        lblHdValorPgto.Enabled = False
        lblHdConta.Enabled = False
        
        frmTitulo.Enabled = False
        txbValorBaixar.Enabled = False: lblValorBaixar.Enabled = False
        
        frmTitulosAbertos.Enabled = False
        lblHdVencimento.Enabled = False
        lblHdValorTitulo.Enabled = False
        lblHdValorBaixado.Enabled = False
        lblHdValorBaixar.Enabled = False
        lblHdDesconto.Enabled = False
        lblHdAcrescimo.Enabled = False
        lblHdObservacao.Enabled = False

        Call btnPgtoCancelar_Click
        
        btnPgtoInclui.Visible = False
        btnPgtoAltera.Visible = False
        btnPgtoExclui.Visible = False
        lstRecebimentos.Enabled = False: lstRecebimentos.ForeColor = &H80000010
        lstTitulos.Enabled = False: lstTitulos.ForeColor = &H80000010
        
    ElseIf Acao = "Habilitar" Then
        txbData.Enabled = True: lblData.Enabled = True: btnData.Enabled = True
        cbbObra.Enabled = True: lblObra.Enabled = True
        
        frmTipo.Enabled = True
        optManual.Enabled = True
        'optAutomatico.Enabled = True

        
    ElseIf Acao = "Limpar" Then
        lblCabID.Caption = ""
        lblCabData.Caption = ""
        txbData.Text = ""
        cbbObra.ListIndex = -1
        optManual.Value = False
        optAutomatico.Value = False
        txbValorPgto.Text = Format(0, "#,##0.00")
        cbbConta.ListIndex = -1
        txbValorBaixar.Text = Empty
        
        lstTitulos.Clear
        lstRecebimentos.Clear
        lstPrincipal.ListIndex = -1
        
        lblTotalBaixar.Caption = Format(0, "#,##0.00")
        lblTotalPagamentos.Caption = Format(0, "#,##0.00")
        lblTotalTitulos.Caption = Format(0, "#,##0.00")
        
    End If
    
    Call Filtros("Habilitar")

End Sub

Private Sub btnIncluir_Click()
    Call PosDecisaoTomada("Inclusão")
    lstPrincipal.ListIndex = -1
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
    
    btnIncluir.Visible = False: btnExcluir.Visible = False
    
    If Decisao = "Inclusão" Then
        Call Campos("Limpar")
    End If
    
    If Decisao <> "Exclusão" Then
    
        Call Campos("Habilitar")
        
        MultiPage1.Value = 1
        
        If Decisao = "Inclusão" Then
            txbData.Text = Date
            If MultiPage1.Value = 1 Then
                cbbObra.SetFocus
            End If
        Else
            If MultiPage1.Value = 1 Then
                cbbObra.SetFocus
            End If
        End If
        
    Else
        MultiPage1.Value = 0
    End If
    
    lstPrincipal.Enabled = False
    lstPrincipal.ForeColor = &H80000010
    frmFormaPagamento.Visible = False
    
    Call Filtros("Desabilitar")
    
    btnPaginaInicial.Enabled = False
    btnPaginaAnterior.Enabled = False
    btnPaginaSeguinte.Enabled = False
    btnPaginaFinal.Enabled = False
    
End Sub
Private Sub cbbObra_AfterUpdate()
    
    If cbbObra.ListIndex > -1 And cbbObra.Text <> "" Then
        
        Call lstTitulosPopular(CLng(cbbObra.List(cbbObra.ListIndex, 1)))
        
        MultiPage1.Value = 2
        
    End If
End Sub
Private Sub lstTitulosPopular(ObraID As Long)

    Dim r       As New ADODB.Recordset
    Dim cVlrPg  As Currency
    Dim cSaldo  As Currency
    Dim cVlrBx  As Currency

    If lstPrincipal.ListIndex = -1 Then
    
        sSQL = "SELECT * "
        sSQL = sSQL & "FROM tbl_titulos_receber "
        sSQL = sSQL & "WHERE "
        sSQL = sSQL & "obra_id = " & ObraID
        
        r.Open sSQL, cnn, adOpenStatic
    
        With lstTitulos
            .Clear
            .ColumnCount = 8
            .ColumnWidths = "60pt; 65pt; 65pt; 65pt; 62pt; 62pt; 60pt; 0pt;"
            .Font = "Consolas"
            
            Do Until r.EOF
                
                cVlrBx = oTituloReceber.GetValorBaixado(r.Fields("r_e_c_n_o_").Value)
                cSaldo = r.Fields("valor").Value - cVlrBx
                                
                If cSaldo > 0 Then
                                
                    .AddItem
                    
                    .List(.ListCount - 1, 0) = r.Fields("vencimento").Value
                    .List(.ListCount - 1, 1) = Space(12 - Len(Format(r.Fields("valor").Value, "#,##0.00"))) & Format(r.Fields("valor").Value, "#,##0.00")
                    .List(.ListCount - 1, 2) = Space(12 - Len(Format(cVlrBx, "#,##0.00"))) & Format(cVlrBx, "#,##0.00")
                    .List(.ListCount - 1, 3) = Space(12 - Len(Format(0, "#,##0.00"))) & Format(0, "#,##0.00")
                    .List(.ListCount - 1, 4) = Space(12 - Len(Format(0, "#,##0.00"))) & Format(0, "#,##0.00")
                    .List(.ListCount - 1, 5) = Space(12 - Len(Format(0, "#,##0.00"))) & Format(0, "#,##0.00")
                    .List(.ListCount - 1, 6) = r.Fields("observacao").Value
                    .List(.ListCount - 1, 7) = r.Fields("r_e_c_n_o_").Value
                    
                End If
                
                r.MoveNext
            Loop
            
        End With
        
    Else
        
        sSQL = "SELECT * "
        sSQL = sSQL & "FROM tbl_recebimentos_itens "
        sSQL = sSQL & "WHERE "
        sSQL = sSQL & "recebimento_id = " & oRecebimento.ID
        
        r.Open sSQL, cnn, adOpenStatic
        
        With lstTitulos
            .Clear
            .ColumnCount = 8
            .ColumnWidths = "60pt; 65pt; 65pt; 65pt; 62pt; 62pt; 60pt; 0pt;"
            .Font = "Consolas"
            
            Do Until r.EOF
            
                oTituloReceber.Carrega r.Fields("titulo_id").Value
                
                .AddItem
                
                .List(.ListCount - 1, 0) = oTituloReceber.Vencimento
                .List(.ListCount - 1, 1) = Space(12 - Len(Format(oTituloReceber.Valor, "#,##0.00"))) & Format(oTituloReceber.Valor, "#,##0.00")
                .List(.ListCount - 1, 2) = Space(12 - Len(Format(r.Fields("valor_baixado").Value, "#,##0.00"))) & Format(r.Fields("valor_baixado").Value, "#,##0.00")
                .List(.ListCount - 1, 3) = Space(12 - Len(Format(0, "#,##0.00"))) & Format(0, "#,##0.00")
                .List(.ListCount - 1, 4) = Space(12 - Len(Format(r.Fields("valor_desconto").Value, "#,##0.00"))) & Format(r.Fields("valor_desconto").Value, "#,##0.00")
                .List(.ListCount - 1, 5) = Space(12 - Len(Format(r.Fields("valor_acrescimo").Value, "#,##0.00"))) & Format(r.Fields("valor_acrescimo").Value, "#,##0.00")
                .List(.ListCount - 1, 6) = oTituloReceber.Observacao
                .List(.ListCount - 1, 7) = r.Fields("r_e_c_n_o_").Value
                
                r.MoveNext
            Loop
            
        End With
        
    End If
    
    Set r = Nothing
    
    Call TotalizaTitulos
    
End Sub
Private Sub TotalizaTitulos()

    Dim cTotal As Currency
    Dim i As Integer
    
    For i = 0 To lstTitulos.ListCount - 1
        cTotal = cTotal + CCur(lstTitulos.List(i, 1))
    Next i
    
    lblTotalTitulos.Caption = Format(cTotal, "#,##0.00")

End Sub
Private Sub optManual_Click()
    Call PosDecisaoTipo
End Sub
Private Sub optAutomatico_Click()
    Call PosDecisaoTipo
End Sub
Private Sub PosDecisaoTipo()

    Dim i As Integer

    If optManual.Value = True And lstPrincipal.ListIndex = -1 Then
        frmTitulosAbertos.Enabled = True
        lblHdVencimento.Enabled = True
        lblHdValorTitulo.Enabled = True
        lblHdValorBaixado.Enabled = True
        lblHdValorBaixar.Enabled = True
        lblHdDesconto.Enabled = True
        lblHdAcrescimo.Enabled = True
        lblHdObservacao.Enabled = True
        lstTitulos.Enabled = True: lstTitulos.ForeColor = &H80000008
        frmTotalBaixar.Visible = True
    Else
        frmTitulo.Enabled = False
        lblHdVencimento.Enabled = False
        lblHdValorTitulo.Enabled = False
        lblHdValorBaixado.Enabled = False
        lblHdValorBaixar.Enabled = False
        lblHdDesconto.Enabled = False
        lblHdAcrescimo.Enabled = False
        lblHdObservacao.Enabled = False
        
        For i = 0 To lstTitulos.ListCount - 1
            lstTitulos.List(i, 3) = Space(9 - Len(Format(0, "#,##0.00"))) & Format(0, "#,##0.00")
        Next i
        
        lstTitulos.Enabled = False: lstTitulos.ForeColor = &H80000010
        frmTotalBaixar.Visible = False
    End If

End Sub
Private Sub btnPgtoInclui_Click()

    Call AcaoPgto("Incluir", True)

End Sub
Private Sub btnPgtoAltera_Click()

    Call AcaoPgto("Alterar", True)

End Sub
Private Sub btnPgtoExclui_Click()

    Call AcaoPgto("Excluir", True)

End Sub
Private Sub btnPgtoCancelar_Click()

    Call AcaoPgto("Cancelar", False)
    
End Sub
Private Sub btnPgtoConfirmar_Click()

    Dim sDecisaoLancamento  As String
    Dim sDecisaoItem        As String
    
    sDecisaoLancamento = Replace(btnConfirmar.Caption, "Confirmar ", "")
    sDecisaoItem = btnPgtoConfirmar.Caption
    
    If sDecisaoItem = "Incluir" Then
    
        If ValidaPgto = True Then
            
            With lstRecebimentos
                .ColumnCount = 3
                .ColumnWidths = "65pt; 60pt; 0pt;"
                .Font = "Consolas"
                .AddItem
                
                .List(.ListCount - 1, 0) = Space(12 - Len(Format(CDbl(txbValorPgto.Text), "#,##0.00"))) & Format(CDbl(txbValorPgto.Text), "#,##0.00")
                .List(.ListCount - 1, 1) = cbbConta.List(cbbConta.ListIndex, 0)
                .List(.ListCount - 1, 2) = cbbConta.List(cbbConta.ListIndex, 1)
                
            End With
            
            Call btnPgtoCancelar_Click

        End If
    ElseIf sDecisaoItem = "Alterar" Then
        If ValidaPgto = True Then
            With lstRecebimentos
                .List(.ListIndex, 0) = Space(12 - Len(Format(CDbl(txbValorPgto.Text), "#,##0.00"))) & Format(CDbl(txbValorPgto.Text), "#,##0.00")
                .List(.ListIndex, 1) = cbbConta.List(cbbConta.ListIndex, 0)
                .List(.ListIndex, 2) = cbbConta.List(cbbConta.ListIndex, 1)
            End With
            
            Call btnPgtoCancelar_Click
        End If
    ElseIf sDecisaoItem = "Excluir" Then
        lstRecebimentos.RemoveItem (lstRecebimentos.ListIndex)
        Call btnPgtoCancelar_Click
    End If
    
    Call TotalizaPagamentos
    
End Sub
Private Function ValidaPgto() As Boolean
    ValidaPgto = False
    
    If cbbConta.ListIndex = -1 Then
        MsgBox "Campo 'Conta' é obrigatório", vbCritical
        MultiPage1.Value = 2: cbbConta.SetFocus
    ElseIf txbValorPgto.Text = Empty Then
        MsgBox "Campo 'Valor pgto.' é obrigatório", vbCritical
        MultiPage1.Value = 2: txbValorPgto.SetFocus
    Else
        ValidaPgto = True
    End If
    
End Function
Private Sub AcaoPgto(Acao As String, Habilitar As Boolean)
    
    btnPgtoConfirmar.Caption = Acao
    
    If Acao = "Incluir" Then
        
        lstRecebimentos.ListIndex = -1
        cbbConta.ListIndex = -1
        txbValorPgto.Text = Format(CCur(lblTotalBaixar.Caption) - CCur(lblTotalPagamentos.Caption), "#,##0.00")
        
    End If
    
    If Habilitar = True Then
        txbValorPgto.Enabled = Habilitar: lblValorPgto.Enabled = Habilitar
        cbbConta.Enabled = Habilitar: lblConta.Enabled = Habilitar
        
        btnPgtoInclui.Visible = Not Habilitar
        btnPgtoAltera.Visible = Not Habilitar
        btnPgtoExclui.Visible = Not Habilitar
        btnPgtoCancelar.Visible = Habilitar
        btnPgtoConfirmar.Visible = Habilitar
        lstRecebimentos.Enabled = Not Habilitar: lstRecebimentos.ForeColor = &H80000010
        btnConfirmar.Enabled = Not Habilitar
        btnCancelar.Enabled = Not Habilitar
    Else
        lstRecebimentos.ListIndex = -1
        txbValorPgto.Enabled = Habilitar: lblValorPgto.Enabled = Habilitar: txbValorPgto.Text = Empty
        cbbConta.Enabled = Habilitar: lblConta.Enabled = Habilitar: cbbConta.ListIndex = -1
        
        btnPgtoInclui.Visible = Not Habilitar
        btnPgtoAltera.Visible = Not Habilitar
        btnPgtoExclui.Visible = Not Habilitar
        btnPgtoCancelar.Visible = Habilitar
        btnPgtoConfirmar.Visible = Habilitar
        lstRecebimentos.Enabled = Not Habilitar: lstRecebimentos.ForeColor = &H80000008
        btnConfirmar.Enabled = Not Habilitar
        btnCancelar.Enabled = Not Habilitar
    End If
    
End Sub
Private Sub cbbContaPopular()
    
    Dim idx         As Integer
    Dim col         As New Collection
    Dim n           As Variant

    Set col = oConta.Listar("nome")
    
    idx = cbbConta.ListIndex
    
    cbbConta.Clear
    
    For Each n In col
        
        oConta.Carrega CLng(n)
    
        With cbbConta
            .AddItem
            .List(.ListCount - 1, 0) = oConta.Nome
            .List(.ListCount - 1, 1) = oConta.ID
        End With
        
    Next n
    
    cbbConta.ListIndex = idx

End Sub

Private Sub TotalizaPagamentos()

    Dim cTotal As Currency
    Dim i As Integer
    
    For i = 0 To lstRecebimentos.ListCount - 1
        cTotal = cTotal + CCur(lstRecebimentos.List(i, 0))
    Next i
    
    lblTotalPagamentos.Caption = Format(cTotal, "#,##0.00")

End Sub
Private Sub btnConfirmar_Click()
    
    Dim vbResposta As VBA.VbMsgBoxResult
    Dim sDecisao As String
    Dim i As Integer
    
    sDecisao = Replace(btnConfirmar.Caption, "Confirmar ", "")
    
    If Valida(sDecisao) = True Then
    
        vbResposta = MsgBox("Deseja realmente fazer a " & sDecisao & "?", vbYesNo + vbQuestion, "Pergunta")
        
        If vbResposta = vbYes Then
        
            ' Cabeçalho do recebimento
            If sDecisao = "Inclusão" Then
                
                oRecebimento.Inclui
                
                ' Itens do pagamento
                For i = 0 To lstTitulos.ListCount - 1
                        
                    If CCur(lstTitulos.List(i, 3)) > 0 Then
                        With oRecebimentoItem
                        
                            oTituloReceber.Carrega CLng(lstTitulos.List(i, 7))
                            
                            .RecebimentoID = oRecebimento.ID
                            .TituloID = oTituloReceber.Recno
                            .ValorBaixado = CCur(lstTitulos.List(i, 3))
                            .DataBaixa = oRecebimento.Data
                            .ObraID = oTituloReceber.ObraID
                            .ValorDesconto = CCur(lstTitulos.List(i, 4))
                            .ValorAcrescimo = CCur(lstTitulos.List(i, 5))
                                                        
                            .Inclui
                        End With
                    End If
                Next i
                
                ' Itens da forma de recebimento
                For i = 0 To lstRecebimentos.ListCount - 1
                    
                    With oContaMovimento
                        .ContaID = CLng(lstRecebimentos.List(i, 2))
                        .CliForID = oRecebimento.ObraID
                        .Data = oRecebimento.Data
                        .PagRec = "R"
                        .Valor = CCur(lstRecebimentos.List(i, 0))
                        .TabelaOrigem = "tbl_recebimentos"
                        .RecnoOrigem = oRecebimento.ID
                        
                        oObra.Carrega oRecebimento.ObraID
                        
                        .CategoriaID = oObra.CategoriaID
                    
                        .Inclui
                    End With
        
                Next i
            
            ElseIf sDecisao = "Exclusão" Then
            
                For i = 0 To lstTitulos.ListCount - 1
                    
                    oRecebimentoItem.Exclui CLng(lstTitulos.List(i, 7))
                    
                Next i
            
                oRecebimentoItem.ExcluiMovimentacaoEmContas oRecebimento.ID
                oRecebimento.Exclui oRecebimento.ID
            End If
            
            Call btnFiltrar_Click
            
            ' Exibe mensagem de sucesso na decisão tomada (inclusão, alteração ou exclusão do registro).
            MsgBox sDecisao & " realizada com sucesso.", vbInformation, sDecisao & " de registro"
            
            Call btnCancelar_Click
            
        ElseIf vbResposta = vbNo Then
        
            If sDecisao = "Exclusão" Then
                Call btnCancelar_Click
            End If
            
        End If
    
    End If
    
End Sub
Private Function Valida(Decisao As String) As Boolean
    
    Valida = False
    
    If Decisao = "Inclusão" Then
        If txbData.Text = Empty Then
            MsgBox "Campo 'Data' é obrigatório", vbCritical
            MultiPage1.Value = 1: txbData.SetFocus
        ElseIf cbbObra.ListIndex = -1 Then
            MsgBox "Campo 'Obra' é obrigatório", vbCritical
            MultiPage1.Value = 1: cbbObra.SetFocus
        ElseIf optManual.Value = False And optAutomatico.Value = False Then
            MsgBox "Escolha o tipo de recebimento", vbCritical
            MultiPage1.Value = 2: optManual.SetFocus
        ElseIf lblTotalPagamentos.Caption = Format(0, "#,##0.00") Then
            MsgBox "Não há recebimentos apontados", vbCritical
            MultiPage1.Value = 2
        ElseIf lblTotalBaixar.Caption = Format(0, "#,##0.00") Then
            MsgBox "Não há baixas apontadas", vbCritical
            MultiPage1.Value = 2
        ElseIf CCur(lblTotalBaixar.Caption) <> CCur(lblTotalPagamentos.Caption) Then
            MsgBox "'Total à baixar' e 'Total de recebimentos' está divergente!", vbCritical
            MultiPage1.Value = 2
        Else
            
            If optManual.Value = True And lblTotalBaixar.Caption = Format(0, "#,##0.00") Then
                MsgBox "Você precisa informar o valor que será baixado de cada título.", vbCritical
                MultiPage1.Value = 2: frmTitulosAbertos.SetFocus: Exit Function
            Else
                With oRecebimento
                    .Data = CDate(txbData.Text)
                    .ObraID = CLng(cbbObra.List(cbbObra.ListIndex, 1))
                    
                    If optManual.Value = True Then
                        .TipoBaixa = "M"
                    Else
                        .TipoBaixa = "A"
                    End If
                    
                    .ValorRecebido = CCur(lblTotalPagamentos.Caption)
                    
                End With
                
                Valida = True
            End If
        End If
    Else
        Valida = True
    End If

End Function
Private Sub lstRecebimentos_Change()

    Dim n As Integer

    If lstRecebimentos.ListIndex > -1 And btnPgtoConfirmar.Caption <> "Alterar" Then
        
        txbValorPgto.Text = lstRecebimentos.List(lstRecebimentos.ListIndex, 0)
        
        For n = 0 To cbbConta.ListCount - 1
            If CInt(cbbConta.List(n, 1)) = CInt(lstRecebimentos.List(lstRecebimentos.ListIndex, 2)) Then
                cbbConta.ListIndex = n
                Exit For
            End If
        Next n
        
        btnPgtoAltera.Enabled = True
        btnPgtoExclui.Enabled = True
    End If
End Sub
Private Sub lstTitulos_Change()

    Dim cSaldo      As Currency
    Dim cValorExtra As Currency
    
    If lstTitulos.ListIndex > -1 Then
        
        frmTitulo.Enabled = True
        txbValorBaixar.Enabled = True: lblValorBaixar.Enabled = True
        optDesconto.Visible = True
        optAcrescimo.Visible = True
        
        If CCur(lstTitulos.List(lstTitulos.ListIndex, 3)) = 0 Then
            cSaldo = CCur(lstTitulos.List(lstTitulos.ListIndex, 1)) - CCur(lstTitulos.List(lstTitulos.ListIndex, 2))
            txbValorBaixar.Text = Format(cSaldo, "#,##0.00")
        Else
            txbValorBaixar.Text = lstTitulos.List(lstTitulos.ListIndex, 3)
        End If
        
        If CCur(lstTitulos.List(lstTitulos.ListIndex, 4)) > 0 Or CCur(lstTitulos.List(lstTitulos.ListIndex, 5)) > 0 Then
                                
            If CCur(lstTitulos.List(lstTitulos.ListIndex, 4)) > 0 Then
                optDesconto.Value = True
                cValorExtra = CCur(lstTitulos.List(lstTitulos.ListIndex, 4))
            Else
                optAcrescimo.Value = True
                cValorExtra = CCur(lstTitulos.List(lstTitulos.ListIndex, 5))
            End If
            
            txbDescontoAcrescimo.Text = Format(cValorExtra, "#,##0.00")
        Else
            optDesconto.Value = False
            optAcrescimo.Value = False
            lblDescontoAcrescimo.Visible = False
            txbDescontoAcrescimo.Text = Format(0, "#,##0.00"): txbDescontoAcrescimo.Visible = False
        End If
        
        btnTituloCancelar.Visible = True
        btnTituloConfirmar.Visible = True
        
    End If

End Sub

Private Sub TotalizaBaixar()

    Dim cTotal As Currency
    Dim i As Integer
    
    For i = 0 To lstTitulos.ListCount - 1
        cTotal = cTotal + CCur(lstTitulos.List(i, 3)) - CCur(lstTitulos.List(i, 4)) + CCur(lstTitulos.List(i, 5))
    Next i
    
    If cTotal > 0 Then
        frmFormaPagamento.Visible = True
        btnPgtoInclui.Visible = True
        btnPgtoAltera.Visible = True
        btnPgtoExclui.Visible = True
    Else
        frmFormaPagamento.Visible = False
    End If
    
    lblTotalBaixar.Caption = Format(cTotal, "#,##0.00")

End Sub
Private Sub lstPrincipal_Change()

    Dim n As Long
    
    If lstPrincipal.ListIndex > -1 Then
    
        btnExcluir.Enabled = True
        
        oRecebimento.Carrega CLng(lstPrincipal.List(lstPrincipal.ListIndex, 0))
        
        lblCabID.Caption = Format(oRecebimento.ID, "0000000000")
        lblCabData.Caption = oRecebimento.Data
        
        txbData.Text = oRecebimento.Data
        
        For n = 0 To cbbObra.ListCount - 1
            If CLng(cbbObra.List(n, 1)) = oRecebimento.ObraID Then
                cbbObra.ListIndex = n
                Exit For
            End If
        Next n
        
        If oRecebimento.TipoBaixa = "M" Then
            optManual.Value = True
        Else
            optAutomatico.Value = True
        End If
        
        Call lstTitulosPopular(oRecebimento.ObraID)
        
        frmFormaPagamento.Visible = True
        Call lstRecebimentosPopular(oRecebimento.ID)
    
    End If

End Sub
Private Sub lstRecebimentosPopular(RecebimentoID As Long)

    Dim r       As New ADODB.Recordset
    Dim cVlrPg  As Currency
    Dim cSaldo  As Currency

    If lstPrincipal.ListIndex > -1 Then
    
        sSQL = "SELECT * "
        sSQL = sSQL & "FROM tbl_contas_movimentos "
        sSQL = sSQL & "WHERE "
        sSQL = sSQL & "tabela_origem = 'tbl_recebimentos' "
        sSQL = sSQL & "and recno_origem = " & RecebimentoID & " "
        sSQL = sSQL & "ORDER BY r_e_c_n_o_"
        
        r.Open sSQL, cnn, adOpenStatic
    
        With lstRecebimentos
                .ColumnCount = 4
                .ColumnWidths = "65pt; 60pt; 0pt; 0pt;"
                .Font = "Consolas"
                .Clear
                
                Do Until r.EOF
                
                    .AddItem
                
                    .List(.ListCount - 1, 0) = Space(12 - Len(Format(r.Fields("valor").Value, "#,##0.00"))) & Format(r.Fields("valor").Value, "#,##0.00")
                    
                    oConta.Carrega r.Fields("conta_id").Value
                    
                    .List(.ListCount - 1, 1) = oConta.Nome
                    .List(.ListCount - 1, 2) = r.Fields("conta_id").Value
                    .List(.ListCount - 1, 3) = r.Fields("r_e_c_n_o_").Value
                
                    r.MoveNext
                Loop
            
        End With
    
        Set r = Nothing
    
        Call TotalizaPagamentos
    
    End If

End Sub
Private Sub btnData_Click()
    dtDate = IIf(txbData.Text = Empty, Date, txbData.Text)
    txbData.Text = GetCalendario
End Sub
Private Sub optDesconto_Click()
    
    lblDescontoAcrescimo.Visible = True
    txbDescontoAcrescimo.Visible = True
    txbDescontoAcrescimo.Text = Format(0, "#,##0.00")

End Sub
Private Sub optAcrescimo_Click()

    Call optDesconto_Click

End Sub
Private Sub lstTitulos_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    Call btnTituloConfirmar_Click

End Sub
Private Sub btnTituloConfirmar_Click()

    Dim cSaldo          As Currency
    Dim cBaixar         As Currency
    Dim cValorExtra     As Currency
    Dim iColunaValor    As Integer
    Dim iColunaZero     As Integer

    If lstTitulos.ListIndex > -1 Then
        
        cSaldo = CCur(lstTitulos.List(lstTitulos.ListIndex, 1)) - CCur(lstTitulos.List(lstTitulos.ListIndex, 2))
        cBaixar = CCur(txbValorBaixar.Text)
        
        If cBaixar > cSaldo Then
            lstTitulos.List(lstTitulos.ListIndex, 3) = Space(12 - Len(Format(cSaldo, "#,##0.00"))) & Format(cSaldo, "#,##0.00")
        Else
            lstTitulos.List(lstTitulos.ListIndex, 3) = Space(12 - Len(Format(cBaixar, "#,##0.00"))) & Format(cBaixar, "#,##0.00")
        End If
        
        If optDesconto.Value = True Or optAcrescimo.Value = True Then
        
            cValorExtra = CCur(txbDescontoAcrescimo.Text)
            
            If optDesconto.Value = True Then
                iColunaValor = 4
                iColunaZero = 5
            Else
                iColunaValor = 5
                iColunaZero = 4
            End If
            
            lstTitulos.List(lstTitulos.ListIndex, iColunaValor) = Space(12 - Len(Format(cValorExtra, "#,##0.00"))) & Format(cValorExtra, "#,##0.00")
            lstTitulos.List(lstTitulos.ListIndex, iColunaZero) = Space(12 - Len(Format(0, "#,##0.00"))) & Format(0, "#,##0.00")
            
        End If
        
        Call TotalizaBaixar
        
        Call btnTituloCancelar_Click
        
    End If

End Sub
Private Sub btnTituloCancelar_Click()

    txbValorBaixar.Text = Format(0, "#,##0.00")
    txbValorBaixar.Enabled = False: lblValorBaixar.Enabled = False
    
    optDesconto.Visible = False: optDesconto.Value = False
    optAcrescimo.Visible = False: optAcrescimo.Value = False
    
    txbDescontoAcrescimo.Visible = False: lblDescontoAcrescimo.Visible = False
    
    frmTitulo.Enabled = False
    
    lstTitulos.ListIndex = -1
    
    btnTituloCancelar.Visible = False
    btnTituloConfirmar.Visible = False

End Sub
Private Sub cbbFltObraPopular()
    
    Dim col         As New Collection
    Dim n           As Variant

    Set col = oObra.Listar("bairro")
    
    With cbbFltObra
        .Clear
        .ColumnCount = 2
        .ColumnWidths = "180pt; 0pt;"
        .AddItem
        .List(.ListCount - 1, 0) = "***TODOS***"
        .List(.ListCount - 1, 1) = 0
    End With
    
    For Each n In col
        
        oObra.Carrega CLng(n)
        
        oCliente.Carrega oObra.ClienteID
    
        With cbbFltObra
            .AddItem
            .List(.ListCount - 1, 0) = oObra.Bairro & ": " & oCliente.Nome & ": " & oObra.Endereco
            .List(.ListCount - 1, 1) = oObra.ID
        End With
        
    Next n
    
    cbbFltObra.ListIndex = 0

End Sub
Private Sub Filtros(Acao As String)

    Dim b As Boolean
    
    b = IIf(Acao = "Habilitar", True, False)

    cbbFltObra.Enabled = b: lblFltObra.Enabled = b
    btnFiltrar.Enabled = b
    frmFiltro.Enabled = b

End Sub
Private Sub btnFiltrar_Click()

    Dim lObraID As Long
    
    If cbbFltObra.ListIndex = -1 Then
        lObraID = 0
    Else
        lObraID = CLng(cbbFltObra.List(cbbFltObra.ListIndex, 1))
    End If

    Set myRst = oRecebimento.Recordset(lObraID)
    
    If myRst.PageCount > 0 Then
    
        myRst.AbsolutePage = myRst.PageCount
    
        With scrPagina
            .Max = myRst.PageCount
            .Value = myRst.PageCount
        End With
        
        Call scrPagina_Change
        
    End If

End Sub
Private Sub btnPaginaSeguinte_Click()
    scrPagina.Value = scrPagina.Value + 1
End Sub
Private Sub btnPaginaAnterior_Click()
    scrPagina.Value = scrPagina.Value - 1
End Sub
Private Sub btnPaginaInicial_Click()
    scrPagina.Value = 1
End Sub
Private Sub btnPaginaFinal_Click()
    scrPagina.Value = myRst.PageCount
End Sub
Private Sub scrPagina_Change()

    ' Trata botões de navegação
    If scrPagina.Value = myRst.PageCount And scrPagina.Value > 1 Then
        btnPaginaInicial.Enabled = True
        btnPaginaAnterior.Enabled = True
        btnPaginaSeguinte.Enabled = False
        btnPaginaFinal.Enabled = False
        scrPagina.Enabled = True
    ElseIf scrPagina.Value = 1 And myRst.PageCount = 1 Then
        btnPaginaInicial.Enabled = False
        btnPaginaAnterior.Enabled = False
        btnPaginaSeguinte.Enabled = False
        btnPaginaFinal.Enabled = False
        scrPagina.Enabled = False
    ElseIf scrPagina.Value > 1 And scrPagina.Value < myRst.PageCount Then
        btnPaginaInicial.Enabled = True
        btnPaginaAnterior.Enabled = True
        btnPaginaSeguinte.Enabled = True
        btnPaginaFinal.Enabled = True
        scrPagina.Enabled = True
    ElseIf scrPagina.Value = 1 And myRst.PageCount > 1 Then
        btnPaginaInicial.Enabled = False
        btnPaginaAnterior.Enabled = False
        btnPaginaSeguinte.Enabled = True
        btnPaginaFinal.Enabled = True
        scrPagina.Enabled = True
    End If

    Call Campos("Limpar")
    
    On Error Resume Next
    myRst.AbsolutePage = scrPagina.Value
    
    Call lstPrincipalPopular

End Sub
Private Sub cbbConta_AfterUpdate()

    Dim vbResposta As VbMsgBoxResult
    Dim idx As Integer
    Dim n As Integer
    
    If cbbConta.ListIndex = -1 And cbbConta.Text <> "" Then
        
        vbResposta = MsgBox("Esta conta não existe, deseja cadastrá-lo?", vbQuestion + vbYesNo)
        
        If vbResposta = vbYes Then
            
            oConta.Nome = RTrim(cbbConta.Text)
            oConta.SaldoInicial = 0
            oConta.Inclui
            idx = oConta.ID
            Call cbbContaPopular
            
            For n = 0 To cbbConta.ListCount - 1
                If CInt(cbbConta.List(n, 1)) = idx Then
                    cbbConta.ListIndex = n
                    Exit For
                End If
            Next n
        Else
            cbbConta.ListIndex = -1
        End If

    End If
End Sub
