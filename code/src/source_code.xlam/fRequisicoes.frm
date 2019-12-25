VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fRequisicoes 
   Caption         =   ":: Cadastro de Requisições ::"
   ClientHeight    =   9480
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13215
   OleObjectBlob   =   "fRequisicoes.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fRequisicoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private oRequisicao         As New cRequisicao
Private oObra               As New cObra
Private oCliente            As New cCliente
Private oEtapa              As New cEtapa
Private oFornecedor         As New cFornecedor
Private oProduto            As New cProduto
Private oCompraItem         As New cCompraItem
Private oRequisicaoItem     As New cRequisicaoItem
Private oUM                 As New cUnidadeMedida

Private colControles        As New Collection
Private myRst               As ADODB.RecordSet
Private lPagina             As Long

Private Const sTable As String = "tbl_requisicoes"
Private Const sCampoOrderBy As String = "data"

Private Sub UserForm_Initialize()

    Call cbbObraPopular
    Call cbbEtapaPopular
    Call EventosCampos
    
    Set myRst = New ADODB.RecordSet
    Set myRst = oRequisicao.RecordSet
    
    With scrPagina
        .Min = IIf(myRst.PageCount = 0, 1, myRst.PageCount)
        .Max = myRst.PageCount
    End With
    
    lPagina = myRst.PageCount
    
    If myRst.PageCount > 0 Then
        myRst.AbsolutePage = myRst.PageCount
    End If
    
    scrPagina.Value = lPagina
    
    Call lstPrincipalPopular(lPagina)
    
    Call btnCancelar_Click

End Sub
Private Sub UserForm_Terminate()
    
    ' Destrói objeto da classe cProduto
    Set oRequisicao = Nothing
    Call Desconecta
    
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
            .List(.ListCount - 1, 1) = oObra.Id
        End With
        
    Next n
    
    cbbObra.ListIndex = idx

End Sub
Private Sub cbbEtapaPopular()
    
    Dim idx         As Integer
    Dim col         As New Collection
    Dim n           As Variant

    Set col = oEtapa.Listar("nome")
    
    idx = cbbEtapa.ListIndex
    
    With cbbEtapa
        .Clear
        .ColumnCount = 2
        .ColumnWidths = "60pt; 0pt;"
    End With
    
    For Each n In col
        
        oEtapa.Carrega CLng(n)
    
        With cbbEtapa
            .AddItem
            .List(.ListCount - 1, 0) = oEtapa.Nome
            .List(.ListCount - 1, 1) = oEtapa.Id
        End With
        
    Next n
    
    cbbEtapa.ListIndex = idx

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
Private Sub lstPrincipalPopular(Pagina As Long)

    Dim lPosicao    As Long
    Dim lCount      As Long
    
    With lstPrincipal
        .Clear
        .ColumnCount = 2 ' Funcionário, ID, Empresa, Filial
        .ColumnWidths = "55pt; 55pt;"
        .Enabled = True
        .Font = "Consolas"
        
        lCount = 1
        
        While Not myRst.EOF = True And lCount <= myRst.PageSize

            .AddItem
            .List(.ListCount - 1, 0) = Format(myRst.Fields("id").Value, "0000000000")
            .List(.ListCount - 1, 1) = myRst.Fields("data").Value

            lCount = lCount + 1
            myRst.MoveNext
            
        Wend

    End With
   
    lblPaginaAtual.Caption = "Página " & Format(scrPagina.Value, "#,##0") & " de " & Format(myRst.PageCount, "#,##0")

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

Private Sub Campos(Acao As String)

    If Acao = "Desabilitar" Then
        txbData.Enabled = False: lblData.Enabled = False: btnData.Enabled = False
 '       cbbFornecedor.Enabled = False: lblFornecedor.Enabled = False
        
'        frmTipo.Enabled = False
'        optManual.Enabled = False
'        optAutomatico.Enabled = False
'
'        frmFormaPagamento.Enabled = False
'        txbValorPgto.Enabled = False: lblValorPgto.Enabled = False
'        cbbConta.Enabled = False: lblConta.Enabled = False
'
'        lblHdValorPgto.Enabled = False
'        lblHdConta.Enabled = False
'
'        frmTitulo.Enabled = False
'        txbValorBaixar.Enabled = False: lblValorBaixar.Enabled = False
'
'        lblExtrato.Enabled = False
'        lblHdVencimento.Enabled = False
'        lblHdValorTitulo.Enabled = False
'        lblHdValorBaixado.Enabled = False
'        lblHdValorBaixar.Enabled = False
'        lblHdObservacao.Enabled = False
'
'        Call btnPgtoCancelar_Click
'        btnPgtoInclui.Visible = False
'        btnPgtoAltera.Visible = False
'        btnPgtoExclui.Visible = False
'        lstPgtos.Enabled = False: lstPgtos.ForeColor = &H80000010
'        lstTitulos.Enabled = False: lstTitulos.ForeColor = &H80000010
        
    ElseIf Acao = "Habilitar" Then
        txbData.Enabled = True: lblData.Enabled = True: btnData.Enabled = True
'        cbbFornecedor.Enabled = True: lblFornecedor.Enabled = True
        
'        frmTipo.Enabled = True
'        optManual.Enabled = True
'        'optAutomatico.Enabled = True
'
'        frmFormaPagamento.Enabled = True
'        lblHdValorPgto.Enabled = True
'        lblHdConta.Enabled = True
'        lstPgtos.Enabled = True: lstPgtos.ForeColor = &H80000008
'        btnPgtoInclui.Visible = True
'        btnPgtoAltera.Visible = True
'        btnPgtoExclui.Visible = True
'
'        lblExtrato.Enabled = True
        
    ElseIf Acao = "Limpar" Then
        lblCabID.Caption = ""
        lblCabData.Caption = ""
        txbData.Text = ""
'        cbbFornecedor.ListIndex = -1
'        optManual.Value = False
'        optAutomatico.Value = False
'        txbValorPgto.Text = Format(0, "#,##0.00")
'        cbbConta.ListIndex = -1
'        txbValorBaixar.Text = Empty
'
        lstCompraItens.Clear
        lstCompraItens.ListIndex = -1
        lstRequisicoes.Clear
        lstRequisicoes.ListIndex = -1
        
        frmItemSelecionado.Visible = True
        frmRequisitar.Visible = True
        
    End If

End Sub
Private Sub btnIncluir_Click()
    Call PosDecisaoTomada("Inclusão")
    lstPrincipal.ListIndex = -1
    Call lstCompraItensPopular
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
                txbData.SetFocus
            End If
        Else
            If MultiPage1.Value = 1 Then
                txbData.SetFocus
            End If
        End If
            
    End If
    
    lstPrincipal.Enabled = False
    lstPrincipal.ForeColor = &H80000010
    
    btnPaginaInicial.Enabled = False
    btnPaginaAnterior.Enabled = False
    btnPaginaSeguinte.Enabled = False
    btnPaginaFinal.Enabled = False
    
End Sub
Private Sub lstCompraItensPopular()

    Dim r       As New ADODB.RecordSet
    Dim dQtdBx  As Currency
    Dim dSaldo  As Currency
    Dim dRequisitado As Currency

    If lstPrincipal.ListIndex = -1 Then
    
        sSQL = "SELECT * "
        sSQL = sSQL & "FROM tbl_compras_itens "
        sSQL = sSQL & "WHERE "
        sSQL = sSQL & "requisitado = False"
        
        r.Open sSQL, cnn, adOpenStatic
    
        With lstCompraItens
            .Clear
            .ColumnCount = 10
            .ColumnWidths = "60pt; 60pt; 80pt; 80pt; 60pt; 60pt; 60pt; 0pt; 60pt; 40pt;"
            ' 0 - Data da compra do item
            ' 1 - Número da compra
            ' 2 - Nome do fornecedor
            ' 3 - Nome do produto
            ' 4 - Quantidade do item
            ' 5 - Valor unitário do item
            ' 6 - Valor total do item
            ' 7 - Recno do item comprado
            ' 8 - Quantidade requisitada do item
            .Font = "Consolas"
            
            Do Until r.EOF
                
                dQtdBx = oCompraItem.GetQtdeBaixada(r.Fields("r_e_c_n_o_").Value)
                dSaldo = r.Fields("quantidade").Value - dQtdBx
                dRequisitado = oRequisicaoItem.GetQtdeRequisitada(r.Fields("r_e_c_n_o_").Value)
                                
                If dSaldo > 0 Then
                    
                    oFornecedor.Carrega r.Fields("fornecedor_id").Value
                    oProduto.Carrega r.Fields("produto_id").Value
                    oUM.Carrega r.Fields("um_id").Value
                    
                    .AddItem
                    
                    .List(.ListCount - 1, 0) = r.Fields("data").Value
                    .List(.ListCount - 1, 1) = Format(r.Fields("compra_id").Value, "0000000000")
                    .List(.ListCount - 1, 2) = oFornecedor.Nome
                    .List(.ListCount - 1, 3) = oProduto.Nome
                    .List(.ListCount - 1, 4) = Space(9 - Len(Format(r.Fields("quantidade").Value, "#,##0.00"))) & Format(r.Fields("quantidade").Value, "#,##0.00")
                    .List(.ListCount - 1, 5) = Space(9 - Len(Format(r.Fields("unitario").Value, "#,##0.00"))) & Format(r.Fields("unitario").Value, "#,##0.00")
                    .List(.ListCount - 1, 6) = Space(9 - Len(Format(r.Fields("total").Value, "#,##0.00"))) & Format(r.Fields("total").Value, "#,##0.00")
                    .List(.ListCount - 1, 7) = r.Fields("r_e_c_n_o_").Value
                    .List(.ListCount - 1, 8) = Space(9 - Len(Format(dRequisitado, "#,##0.00"))) & Format(dRequisitado, "#,##0.00")
                    .List(.ListCount - 1, 9) = oUM.Abreviacao
                    
                End If
                
                r.MoveNext
            Loop
            
        End With
    Else
    End If
    
    Set r = Nothing
    
End Sub
Private Sub lstCompraItens_Change()

    Dim i As Integer
        
    If lstCompraItens.ListIndex > -1 Then
    
        i = lstCompraItens.ListIndex
        
        lblItemProduto.Caption = lstCompraItens.List(i, 3)
        txbQtde.Text = Format(lstCompraItens.List(i, 4) - lstCompraItens.List(i, 8), "#,##0.00")
        lblItemUnitario.Caption = lstCompraItens.List(i, 5)
        lblItemTotal.Caption = Format(CDbl(lstCompraItens.List(i, 5)) * (CDbl(lstCompraItens.List(i, 4)) - CDbl(lstCompraItens.List(i, 8))), "#,##0.00")
        
        txbQtde.Enabled = True: lblQtde.Enabled = True
        cbbObra.Enabled = True: lblObra.Enabled = True
        cbbEtapa.Enabled = True: lblEtapa.Enabled = True
        btnRequisitar.Enabled = True
        
        lstRequisicoes.ListIndex = -1
        btnRequisicaoExclui.Enabled = False
        
        
    End If
        
End Sub
Private Sub btnRequisitar_Click()

    If ValidaItem = True Then
    
        Dim cVlrTotal As Currency
    
        With lstRequisicoes
            .ColumnCount = 9
            .ColumnWidths = "0pt; 85pt; 55pt; 55pt; 55pt; 240pt; 0pt; 60pt; 0pt;"
            ' Colunas
            ' 0 - Recno do item da compra
            ' 1 - Descrição do item
            ' 2 - Quantidade do item
            ' 3 - Preço unitário do item
            ' 4 - Preço total do item
            ' 5 - Descrição da obra
            ' 6 - Código da obra
            ' 7 - Descrição da etapa da obra
            ' 8 - Código da etapa da obra
            
            .Font = "Consolas"
            
            
            oCompraItem.Carrega CLng(lstCompraItens.List(lstCompraItens.ListIndex, 7))
        
            .AddItem
            .List(.ListCount - 1, 0) = oCompraItem.Recno
            .List(.ListCount - 1, 1) = lblItemProduto.Caption
            .List(.ListCount - 1, 2) = Space(9 - Len(Format(CDbl(txbQtde.Text), "#,##0.00"))) & Format(CDbl(txbQtde.Text), "#,##0.00")
            .List(.ListCount - 1, 3) = Space(9 - Len(Format(CCur(lblItemUnitario.Caption), "#,##0.00"))) & Format(CCur(lblItemUnitario.Caption), "#,##0.00")
            
            cVlrTotal = CCur(txbQtde.Text) * CCur(lblItemUnitario.Caption)
            
            .List(.ListCount - 1, 4) = Space(9 - Len(Format(cVlrTotal, "#,##0.00"))) & Format(cVlrTotal, "#,##0.00")
            .List(.ListCount - 1, 5) = cbbObra.List(cbbObra.ListIndex, 0)
            .List(.ListCount - 1, 6) = cbbObra.List(cbbObra.ListIndex, 1)
            .List(.ListCount - 1, 7) = cbbEtapa.List(cbbEtapa.ListIndex, 0)
            .List(.ListCount - 1, 8) = cbbEtapa.List(cbbEtapa.ListIndex, 1)
            
        End With
        
        Call AtualizaColunaRequisitado(CDbl(txbQtde.Text), lstCompraItens.ListIndex)
   
    End If

End Sub
Private Function ValidaItem() As Boolean

    ValidaItem = False
    
    If cbbObra.ListIndex = -1 Then
        MsgBox "Campo 'Obra' é obrigatório", vbCritical
        MultiPage1.Value = 2: cbbObra.SetFocus
    ElseIf cbbEtapa.ListIndex = -1 Then
        MsgBox "Campo 'Etapa' é obrigatório", vbCritical
        MultiPage1.Value = 2: cbbEtapa.SetFocus
    ElseIf CDbl(txbQtde.Text) > (CDbl(lstCompraItens.List(lstCompraItens.ListIndex, 4)) - CDbl(lstCompraItens.List(lstCompraItens.ListIndex, 8))) Then
        MsgBox "Item sem saldo para requisitar", vbCritical
        MultiPage1.Value = 2
    Else
        ValidaItem = True
    End If

End Function
Private Sub txbQtde_AfterUpdate()
    lblItemTotal.Caption = Format(CCur(txbQtde.Text) * CCur(lblItemUnitario.Caption), "#,##0.00")
    txbQtde.Text = Format(txbQtde.Text, "#,##0.00")
End Sub
Private Sub AtualizaColunaRequisitado(Quantidade As Double, Indice As Integer)

    'Dim i As Integer
    Dim dRequisitado As Double
    
    'i = lstCompraItens.ListIndex
    
    dRequisitado = CDbl(lstCompraItens.List(Indice, 8)) + Quantidade
    
    lstCompraItens.List(Indice, 8) = Space(9 - Len(Format(dRequisitado, "#,##0.00"))) & Format(dRequisitado, "#,##0.00")
    
    lblItemProduto.Caption = ""
    txbQtde.Text = Format(0, "#,##0.00")
    lblItemUnitario.Caption = Format(0, "#,##0.00")
    lblItemTotal.Caption = Format(0, "#,##0.00")
    cbbObra.ListIndex = -1
    cbbEtapa.ListIndex = -1
    lstCompraItens.ListIndex = -1
    
    txbQtde.Enabled = False: lblQtde.Enabled = False
    cbbObra.Enabled = False: lblObra.Enabled = False
    cbbEtapa.Enabled = False: lblEtapa.Enabled = False
    btnRequisitar.Enabled = False

End Sub
Private Sub lstRequisicoes_Click()

    Dim i As Integer
        
    If lstRequisicoes.ListIndex > -1 Then
    
        i = lstRequisicoes.ListIndex
        
        btnRequisicaoExclui.Enabled = True
        lstCompraItens.ListIndex = -1
        
        txbQtde.Enabled = False: lblQtde.Enabled = False
        cbbObra.Enabled = False: lblObra.Enabled = False
        cbbEtapa.Enabled = False: lblEtapa.Enabled = False
        btnRequisitar.Enabled = False
    End If

End Sub
Private Sub btnRequisicaoExclui_Click()
    
    If lstRequisicoes.ListIndex > -1 Then
        
        Dim d As Double
        Dim i As Integer
        
        d = CDbl(lstRequisicoes.List(lstRequisicoes.ListIndex, 2))
        
        For i = 0 To lstCompraItens.ListCount - 1
            If lstRequisicoes.List(lstRequisicoes.ListIndex, 0) = lstCompraItens.List(i, 7) Then
                Call AtualizaColunaRequisitado(d * -1, i)
            End If
        Next i
        
        lstRequisicoes.RemoveItem (lstRequisicoes.ListIndex)
        
        MsgBox "Item excluído com sucesso!", vbInformation
        
        btnRequisicaoExclui.Enabled = False
        
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
        
            ' Cabeçalho da requisição
            If sDecisao = "Inclusão" Then
                oRequisicao.Inclui
            End If
            
            ' Itens requisitados
            For i = 0 To lstRequisicoes.ListCount - 1
            
                If sDecisao = "Inclusão" Then
                    
                    With oRequisicaoItem
                    
                        oCompraItem.Carrega CLng(lstRequisicoes.List(i, 0))
                        
                        .RequisicaoID = oRequisicao.Id
                        .ProdutoID = oCompraItem.ProdutoID
                        .ObraID = CLng(lstRequisicoes.List(i, 6))
                        .EtapaID = CLng(lstRequisicoes.List(i, 8))
                        .Qtde = CDbl(lstRequisicoes.List(i, 2))
                        .UmID = oCompraItem.UmID
                        .Unitario = CCur(lstRequisicoes.List(i, 3))
                        .Total = CCur(lstRequisicoes.List(i, 4))
                        .Data = oRequisicao.Data
                        .TabelaOrigem = "tbl_compras_itens"
                        .RecnoOrigem = oCompraItem.Recno
                        
                        .Inclui
                    End With
                    
                    If oCompraItem.Quantidade = oCompraItem.GetQtdeBaixada(oRequisicaoItem.RecnoOrigem) Then
                        oCompraItem.ItemTotalmenteRequisitado oCompraItem.Recno
                    End If
                    
                ElseIf sDecisao = "Exclusão" Then
                
                    oCompraItem.Carrega CLng(lstRequisicoes.List(i, 0))
                
                    With oRequisicaoItem
                        .Recno = CLng(lstRequisicoes.List(i, 9))
                        .Exclui .Recno
                    End With
                    
                    If oCompraItem.Quantidade > oCompraItem.GetQtdeBaixada(oCompraItem.Recno) Then
                        oCompraItem.CancelaRequisicaoTotalItem oCompraItem.Recno
                    End If
                    
                End If
            Next i
            
            If sDecisao = "Exclusão" Then
                oRequisicao.Exclui oRequisicao.Id
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
            Set myRst = oRequisicao.RecordSet
        
            With scrPagina
                .Min = 1
                .Max = myRst.PageCount
            End With
            
            If myRst.PageCount > 0 Then
                lPagina = myRst.PageCount
                myRst.AbsolutePage = myRst.PageCount
                scrPagina.Value = lPagina
            End If
            
            Call lstPrincipalPopular(lPagina)
            
            ' Exibe mensagem de sucesso na decisão tomada (inclusão, alteração ou exclusão do registro).
            MsgBox sDecisao & " realizada com sucesso.", vbInformation, sDecisao & " de registro"
            
            MultiPage1.Value = 0
            
            Call btnCancelar_Click
            
        ElseIf vbResposta = vbNo Then
            Call btnCancelar_Click
        End If
    
    End If
    
End Sub
Private Function Valida() As Boolean
    
    Valida = False
    
    If txbData.Text = Empty Then
        MsgBox "Campo 'Data' é obrigatório", vbCritical
        MultiPage1.Value = 1: txbData.SetFocus
    Else
        If lstRequisicoes.ListCount = 0 Then
            MsgBox "Não há itens requisitados.", vbCritical
            MultiPage1.Value = 2
        Else
            With oRequisicao
                .Data = CDate(txbData.Text)
            End With
            
            Valida = True
        End If
    
    End If
    

End Function
Private Sub lstPrincipal_Change()

    Dim n As Long
    
    If lstPrincipal.ListIndex > -1 Then
    
        btnExcluir.Enabled = True
        
        oRequisicao.Carrega CLng(lstPrincipal.List(lstPrincipal.ListIndex, 0))
        
        lblCabID.Caption = Format(oRequisicao.Id, "0000000000")
        lblCabData.Caption = oRequisicao.Data
        
        txbData.Text = oRequisicao.Data
        
        Call lstRequisicoesPopular(oRequisicao.Id)
        
        lstRequisicoes.Enabled = False
        btnRequisicaoExclui.Enabled = False
        frmRequisitar.Visible = False
        frmItemSelecionado.Visible = False
        
    
    End If
    
End Sub
Private Sub lstRequisicoesPopular(RequisicaoID As Long)

    Dim r       As New ADODB.RecordSet

    If lstPrincipal.ListIndex > -1 Then
    
        sSQL = "SELECT * "
        sSQL = sSQL & "FROM tbl_requisicoes_itens "
        sSQL = sSQL & "WHERE "
        sSQL = sSQL & "requisicao_id = " & RequisicaoID & " "
        sSQL = sSQL & "ORDER BY r_e_c_n_o_"
        
        r.Open sSQL, cnn, adOpenStatic
    
        With lstRequisicoes
                .Clear
                .ColumnCount = 10
                .ColumnWidths = "0pt; 85pt; 55pt; 55pt; 55pt; 240pt; 0pt; 60pt; 0pt; 0pt;"
                ' Colunas
                ' 0 - Recno do item da compra
                ' 1 - Descrição do item
                ' 2 - Quantidade do item
                ' 3 - Preço unitário do item
                ' 4 - Preço total do item
                ' 5 - Descrição da obra
                ' 6 - Código da obra
                ' 7 - Descrição da etapa da obra
                ' 8 - Código da etapa da obra
                ' 9 - Recno do item requisitado
                
                .Font = "Consolas"
                
                Do Until r.EOF

                    oProduto.Carrega r.Fields("produto_id").Value
                    oObra.Carrega r.Fields("obra_id").Value
                    oCliente.Carrega oObra.ClienteID
                    oEtapa.Carrega r.Fields("etapa_id").Value
                
                    .AddItem
                    .List(.ListCount - 1, 0) = r.Fields("recno_origem").Value
                    .List(.ListCount - 1, 1) = oProduto.Nome
                    .List(.ListCount - 1, 2) = Space(9 - Len(Format(r.Fields("quantidade").Value, "#,##0.00"))) & Format(r.Fields("quantidade").Value, "#,##0.00")
                    .List(.ListCount - 1, 3) = Space(9 - Len(Format(r.Fields("unitario").Value, "#,##0.00"))) & Format(r.Fields("unitario").Value, "#,##0.00")
                    .List(.ListCount - 1, 4) = Space(9 - Len(Format(r.Fields("total").Value, "#,##0.00"))) & Format(r.Fields("total").Value, "#,##0.00")
                    .List(.ListCount - 1, 5) = oObra.Bairro & Space(30 - Len(oObra.Bairro)) & " | " & oCliente.Nome
                    .List(.ListCount - 1, 6) = r.Fields("obra_id").Value
                    .List(.ListCount - 1, 7) = oEtapa.Nome
                    .List(.ListCount - 1, 8) = r.Fields("etapa_id").Value
                    .List(.ListCount - 1, 9) = r.Fields("r_e_c_n_o_").Value
                
                    r.MoveNext
                Loop
            
        End With
    
        Set r = Nothing
    
    End If

End Sub
Private Sub cbbEtapa_AfterUpdate()

    Dim vbResposta As VbMsgBoxResult
    Dim idx As Integer
    Dim n As Integer
    
    If cbbEtapa.ListIndex = -1 And cbbEtapa.Text <> "" Then
        
        vbResposta = MsgBox("Esta etapa não existe, deseja cadastrá-la?", vbQuestion + vbYesNo)
        
        If vbResposta = vbYes Then
            
            oEtapa.Nome = RTrim(cbbEtapa.Text)
            oEtapa.Inclui
            
            idx = oEtapa.Id
            
            Call cbbEtapaPopular
            
            For n = 0 To cbbEtapa.ListCount - 1
                If CInt(cbbEtapa.List(n, 1)) = idx Then
                    cbbEtapa.ListIndex = n
                    Exit For
                End If
            Next n
        Else
            cbbEtapa.ListIndex = -1
        End If

    End If

End Sub
