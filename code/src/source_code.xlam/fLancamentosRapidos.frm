VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fLancamentosRapidos 
   Caption         =   ":: Lancamentos à vista ::"
   ClientHeight    =   9075
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12885
   OleObjectBlob   =   "fLancamentosRapidos.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fLancamentosRapidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private oLancamentoRapido   As New cLancamentoRapido
Private oConta              As New cConta
Private oFornecedor         As New cFornecedor
Private oObra               As New cObra
Private oEtapa              As New cEtapa
Private oCliente            As New cCliente
Private oProduto            As New cProduto
Private oUM                 As New cUnidadeMedida
Private oContaMovimento     As New cContaMovimento
Private oRequisicao         As New cRequisicao
Private oRequisicaoItem     As New cRequisicaoItem
Private oCategoria          As New cCategoria

Private colControles        As New Collection
Private myRst               As ADODB.Recordset

Private Const sTable As String = "tbl_compras"
Private Const sCampoOrderBy As String = "data"

Private Sub UserForm_Initialize()
    
    Call cbbContaPopular
    Call cbbPagRecPopular
    Call cbbFornecedorPopular
    Call cbbObra2Popular
    Call cbbObraPopular
    Call cbbEtapaPopular
    Call cbbProdutoPopular
    Call cbbUMPopular
    
    Call EventosCampos
    
    Set myRst = oLancamentoRapido.Recordset
    
    Call scrPaginaAtualiza(False, Null)
    
    Call btnCancelar_Click

End Sub
Private Sub UserForm_Terminate()
    
    Set oLancamentoRapido = Nothing
    Set oConta = Nothing
    Set oFornecedor = Nothing
    Set oObra = Nothing
    Set oEtapa = Nothing
    Set oCliente = Nothing
    Set oProduto = Nothing
    Set oUM = Nothing
    
    Call Desconecta
    
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
Private Sub cbbPagRecPopular()

    With cbbPagRec
    
        .Clear
        
        .AddItem
        .List(.ListCount - 1, 0) = "Pagamento"
        .List(.ListCount - 1, 1) = "P"
        
        .AddItem
        .List(.ListCount - 1, 0) = "Recebimento"
        .List(.ListCount - 1, 1) = "R"
    End With

End Sub
Private Sub cbbFornecedorPopular()
    
    Dim idx         As Integer
    Dim col         As New Collection
    Dim n           As Variant

    Set col = oFornecedor.Listar("nome")
    
    idx = cbbFornecedor.ListIndex
    
    cbbFornecedor.Clear
    
    For Each n In col
        
        oFornecedor.Carrega CLng(n)
    
        With cbbFornecedor
            .AddItem
            .List(.ListCount - 1, 0) = oFornecedor.Nome
            .List(.ListCount - 1, 1) = oFornecedor.ID
        End With
        
    Next n
    
    cbbFornecedor.ListIndex = idx

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
Private Sub chbRequisita_Click()
    If chbRequisita.Value = False Then
        MultiPage1.Pages(2).Visible = False
        frmTotalRequisicao.Visible = False
        txbValor.Enabled = True: lblValor.Enabled = True
    Else
        MultiPage1.Pages(2).Visible = True
        'MultiPage1.Value = 2
        frmTotalRequisicao.Visible = True
        txbValor.Enabled = False: lblValor.Enabled = False
    End If
End Sub
Private Sub btnData_Click()
    dtDate = IIf(txbData.Text = Empty, Date, txbData.Text)
    txbData.Text = GetCalendario
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
            .List(.ListCount - 1, 1) = oEtapa.ID
        End With
        
    Next n
    
    cbbEtapa.ListIndex = idx

End Sub
Private Sub cbbProdutoPopular()
    
    Dim idx         As Integer
    Dim col         As New Collection
    Dim n           As Variant

    Set col = oProduto.Listar("nome")
    
    idx = cbbProduto.ListIndex
    
    cbbProduto.Clear
    
    For Each n In col
        
        oProduto.Carrega CLng(n)
    
        With cbbProduto
            .AddItem
            .List(.ListCount - 1, 0) = oProduto.Nome
            .List(.ListCount - 1, 1) = oProduto.ID
        End With
        
    Next n
    
    cbbProduto.ListIndex = idx

End Sub
Private Sub cbbUMPopular()
    
    Dim idx         As Integer
    Dim col         As New Collection
    Dim n           As Variant

    Set col = oUM.Listar("abreviacao")
    
    idx = cbbUM.ListIndex
    
    cbbUM.Clear
    
    For Each n In col
        
        oUM.Carrega CLng(n)
    
        With cbbUM
            .AddItem
            .List(.ListCount - 1, 0) = oUM.Abreviacao
            .List(.ListCount - 1, 1) = oUM.ID
        End With
        
    Next n
    
    cbbUM.ListIndex = idx

End Sub
Private Sub lstPrincipalPopular(Pagina As Variant)

    Dim lCount      As Long
    
    If myRst.PageCount > 0 Then
    
        'If IsNull(Pagina) Then
            'myRst.AbsolutePage = myRst.PageCount
        'Else
            myRst.AbsolutePage = Pagina
        'End If
        
        If myRst.AbsolutePage = adPosEOF Then
            lblPaginaAtual.Caption = "Página " & Format(myRst.PageCount, "#,##0") & " de " & Format(myRst.PageCount, "#,##0")
        Else
            lblPaginaAtual.Caption = "Página " & Format(myRst.AbsolutePage, "#,##0") & " de " & Format(myRst.PageCount, "#,##0")
        End If
    
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
        
'        With scrPagina
'            .Max = myRst.PageCount:
'            .Value = myRst.PageCount
'        End With
    Else
        lstPrincipal.Clear
    End If
        

End Sub
Private Sub btnCancelar_Click()

    btnIncluir.Visible = True: btnExcluir.Visible = True
    btnConfirmar.Visible = False: btnCancelar.Visible = False

    Call Campos("Limpar")
    Call Campos("Desabilitar")

    btnExcluir.Enabled = False
    btnIncluir.SetFocus
    
    MultiPage1.Value = 0

    lstPrincipal.ListIndex = -1: lstPrincipal.ForeColor = &H80000008: lstPrincipal.Enabled = True:

End Sub
Private Sub Campos(Acao As String)
    
    Dim sDecisao As String
    
    sDecisao = Replace(btnConfirmar.Caption, "Confirmar ", "")

    If Acao = "Desabilitar" Then
        txbData.Enabled = False: lblData.Enabled = False: btnData.Enabled = False
        cbbConta.Enabled = False: lblConta.Enabled = False
        cbbPagRec.Enabled = False: lblPagRec.Enabled = False
        txbValor.Enabled = False: lblValor.Enabled = False
        
        cbbFornecedor.Enabled = False: lblFornecedor.Enabled = False
        cbbObra2.Enabled = False: lblObra2.Enabled = False
        cbbCategoria.Enabled = False: lblCategoria.Enabled = False
        chbRequisita.Enabled = False
        
        frmRequisicao.Enabled = False
        lblHdProduto.Enabled = False
        lblHdQtde.Enabled = False
        lblHdUM.Enabled = False
        lblHdUnitario.Enabled = False
        lblHdTotal.Enabled = False
        lblHdObra.Enabled = False
        lblHdEtapa.Enabled = False
        
        Call btnItemCancelar_Click
        btnItemInclui.Visible = False
        btnItemAltera.Visible = False
        btnItemExclui.Visible = False
        lstRequisicoes.Enabled = False: lstRequisicoes.ForeColor = &H80000010
        
    ElseIf Acao = "Habilitar" Then
        txbData.Enabled = True: lblData.Enabled = True: btnData.Enabled = True
        cbbConta.Enabled = True: lblConta.Enabled = True
        cbbPagRec.Enabled = True: lblPagRec.Enabled = True
        txbValor.Enabled = True: lblValor.Enabled = True
        cbbFornecedor.Enabled = True: lblFornecedor.Enabled = True
        cbbCategoria.Enabled = True: lblCategoria.Enabled = True
        cbbObra2.Enabled = True: lblObra2.Enabled = True
        chbRequisita.Enabled = True
        
        frmRequisicao.Enabled = True
        lblHdProduto.Enabled = True
        lblHdQtde.Enabled = True
        lblHdUM.Enabled = True
        lblHdUnitario.Enabled = True
        lblHdTotal.Enabled = True
        lblHdObra.Enabled = True
        lblHdEtapa.Enabled = True
        lstRequisicoes.Enabled = True: lstRequisicoes.ForeColor = &H80000008
        
        btnItemInclui.Visible = True
        btnItemAltera.Visible = True
        btnItemExclui.Visible = True
        
    ElseIf Acao = "Limpar" Then
        lblCabID.Caption = ""
        'lblCabData.Caption = ""
        txbData.Text = Empty
        cbbConta.ListIndex = -1
        cbbFornecedor.ListIndex = -1
        cbbCategoria.Clear: cbbCategoria.ListIndex = -1
        cbbPagRec.ListIndex = -1
        txbValor.Text = Format(0, "#,##0.00")
        chbRequisita.Value = False
        
        lstRequisicoes.Clear
        lstPrincipal.ListIndex = -1
    End If

End Sub
Private Sub cbbFornecedor_AfterUpdate()

    Dim vbResposta As VbMsgBoxResult
    Dim idx As Integer
    Dim n As Integer
    
    If cbbFornecedor.ListIndex = -1 And cbbFornecedor.Text <> "" Then
        
        vbResposta = MsgBox("Este fornecedor não existe, deseja cadastrá-lo?", vbQuestion + vbYesNo)
        
        If vbResposta = vbYes Then
            
            oFornecedor.Nome = RTrim(cbbFornecedor.Text)
            oFornecedor.Inclui
            idx = oFornecedor.ID
            Call cbbFornecedorPopular
            
            For n = 0 To cbbFornecedor.ListCount - 1
                If CInt(cbbFornecedor.List(n, 1)) = idx Then
                    cbbFornecedor.ListIndex = n
                    Exit For
                End If
            Next n
        Else
            cbbFornecedor.ListIndex = -1
        End If

    End If
End Sub
Private Sub btnIncluir_Click()
    Call PosDecisaoTomada("Inclusão")
    lstPrincipal.ListIndex = -1
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
        
        If MultiPage1.Value = 0 Then
            MultiPage1.Value = 1
        End If
        
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
    
'    cbbFltEmpresa.Enabled = False: lblFltEmpresa.Enabled = False
'    cbbFltFuncionario.Enabled = False: lblFltFuncionario.Enabled = False
'    cbbFltStatus.Enabled = False: lblFltStatus.Enabled = False
'    btnFiltrar.Enabled = False
    btnPaginaInicial.Enabled = False
    btnPaginaAnterior.Enabled = False
    btnPaginaSeguinte.Enabled = False
    btnPaginaFinal.Enabled = False
    
End Sub
Private Sub cbbObra2Popular()
    
    Dim idx         As Integer
    Dim col         As New Collection
    Dim n           As Variant

    Set col = oObra.Listar("bairro")
    
    idx = cbbObra2.ListIndex
    
    With cbbObra2
        .Clear
        .ColumnCount = 4
        .ColumnWidths = "180pt; 0pt; 0pt; 0pt;"
    End With
    
    
    For Each n In col
        
        oObra.Carrega CLng(n)
        
        oCliente.Carrega oObra.ClienteID
    
        With cbbObra2
            .AddItem
            .List(.ListCount - 1, 0) = oObra.Bairro & ": " & oCliente.Nome & ": " & oObra.Endereco
            .List(.ListCount - 1, 1) = oObra.ID
            .List(.ListCount - 1, 1) = oCliente.Nome
            .List(.ListCount - 1, 1) = oObra.Endereco
        End With
        
    Next n
    
    cbbObra2.ListIndex = idx

End Sub
Private Sub cbbPagRec_Change()

    If cbbPagRec.ListIndex > -1 Then
        If cbbPagRec.List(cbbPagRec.ListIndex, 1) = "P" Then
            lblFornecedor.Visible = True: cbbFornecedor.Visible = True
            chbRequisita.Visible = True
            lblObra2.Visible = False: cbbObra2.Visible = False
        Else
            lblFornecedor.Visible = False: cbbFornecedor.Visible = False
            chbRequisita.Visible = False
            lblObra2.Visible = True: cbbObra2.Visible = True
        End If
        
        Call cbbCategoriaPopular(cbbPagRec.List(cbbPagRec.ListIndex, 1))
        
    End If

End Sub
Private Sub btnConfirmar_Click()
    
    Dim vbResposta  As VBA.VbMsgBoxResult
    Dim sDecisao    As String
    Dim i           As Integer
    Dim arr()       As String
    Dim idxProduto  As Long
    Dim idxUM       As Long
    Dim idxObra     As Long
    Dim idxEtapa    As Long
    
    sDecisao = Replace(btnConfirmar.Caption, "Confirmar ", "")
    
    If Valida = True Then
    
        vbResposta = MsgBox("Deseja realmente fazer a " & sDecisao & "?", vbYesNo + vbQuestion, "Pergunta")
        
        If vbResposta = vbYes Then
        
            ' Cabeçalho da compra
            If sDecisao = "Inclusão" Then
            
                oLancamentoRapido.Inclui
                
                With oContaMovimento
                    .ContaID = CLng(cbbConta.List(cbbConta.ListIndex, 1))
                    .CliForID = oLancamentoRapido.CliForID
                    .CategoriaID = oLancamentoRapido.CategoriaID
                    .Data = CDate(txbData.Text)
                    .PagRec = cbbPagRec.List(cbbPagRec.ListIndex, 1)
                    .Valor = CCur(txbValor.Text)
                    .TabelaOrigem = "tbl_lancamentos_rapidos"
                    .RecnoOrigem = oLancamentoRapido.ID
                    .Inclui
                End With
                
                If oLancamentoRapido.Requisitado = True Then
                
                    oRequisicao.Data = CDate(txbData.Text)
                    oRequisicao.Inclui
                    
                    oLancamentoRapido.AtualizaCampoRequisicaoID oLancamentoRapido.ID, oRequisicao.ID
                    
                    For i = 0 To lstRequisicoes.ListCount - 1
                    
                        arr() = Split(lstRequisicoes.List(i, 1), ";")
                        
                        idxProduto = CLng(arr(0))
                        idxUM = CLng(arr(1))
                        idxObra = CLng(arr(2))
                        idxEtapa = CLng(arr(3))
                        
                        With oRequisicaoItem
                            .RequisicaoID = oRequisicao.ID
                            .ProdutoID = idxProduto
                            .ObraID = idxObra
                            .EtapaID = idxEtapa
                            .Qtde = CDbl(lstRequisicoes.List(i, 3))
                            .UmID = idxUM
                            .Unitario = CCur(lstRequisicoes.List(i, 5))
                            .Total = CCur(lstRequisicoes.List(i, 6))
                            .Data = CDate(txbData.Text)
                            .TabelaOrigem = "tbl_lancamentos_rapidos"
                            .RecnoOrigem = oLancamentoRapido.ID
                            .Inclui
                        End With
    
                    Next i
                    
                End If
            
            ElseIf sDecisao = "Exclusão" Then
            
                oLancamentoRapido.ExcluiMovimentacaoContaVinculada oLancamentoRapido.ID
                
                If chbRequisita.Value = True Then
                
                    oLancamentoRapido.ExcluiRequisicaoVinculada oLancamentoRapido.RequisicaoID
                    
                End If
                
                oLancamentoRapido.Exclui oLancamentoRapido.ID
                
            End If
            
            Set myRst = oLancamentoRapido.Recordset
            
            Call scrPaginaAtualiza(True, sDecisao)
            
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
Private Function Valida() As Boolean
    
    Valida = False
    
    If txbData.Text = Empty Then
        MsgBox "Campo 'Data' é obrigatório", vbCritical
        MultiPage1.Value = 1: txbData.SetFocus
    ElseIf cbbPagRec.ListIndex = -1 Then
        MsgBox "Campo 'Pgto/Recebto' é obrigatório", vbCritical
        MultiPage1.Value = 1: cbbPagRec.SetFocus
    ElseIf cbbConta.ListIndex = -1 Then
        MsgBox "Campo 'Conta' é obrigatório", vbCritical
        MultiPage1.Value = 1: cbbConta.SetFocus
    ElseIf cbbCategoria.ListIndex = -1 Then
        MsgBox "Campo 'Categoria' é obrigatório", vbCritical
        MultiPage1.Value = 1: cbbCategoria.SetFocus
    Else
        If cbbPagRec.List(cbbPagRec.ListIndex, 1) = "P" Then
            If cbbFornecedor.ListIndex = -1 Then
                MsgBox "Campo 'Fornecedor' é obrigatório", vbCritical
                MultiPage1.Value = 1: cbbFornecedor.SetFocus
            Else
                
                If chbRequisita.Value = False Then
                
                    If txbValor.Text = Empty Or CCur(txbValor.Text = 0) Then
                        MsgBox "Campo 'Valor' é inválido", vbCritical
                        MultiPage1.Value = 1: txbValor.SetFocus
                    Else
                        With oLancamentoRapido
                            .Data = CDate(txbData.Text)
                            .ContaID = CLng(cbbConta.List(cbbConta.ListIndex, 1))
                            .CliForID = CLng(cbbFornecedor.List(cbbFornecedor.ListIndex, 1))
                            .CategoriaID = CLng(cbbCategoria.List(cbbCategoria.ListIndex, 1))
                            .PagRec = cbbPagRec.List(cbbPagRec.ListIndex, 1)
                            .Valor = CCur(txbValor.Text)
                            .Requisitado = chbRequisita.Value
                        End With
                    
                        Valida = True
                        
                    End If
                    
                Else
                    
                    If lstRequisicoes.ListCount = 0 Then
                        MsgBox "Não há itens requisitados", vbCritical
                        MultiPage1.Value = 2
                    Else
                        With oLancamentoRapido
                            .Data = CDate(txbData.Text)
                            .ContaID = CLng(cbbConta.List(cbbConta.ListIndex, 1))
                            .CliForID = CLng(cbbFornecedor.List(cbbFornecedor.ListIndex, 1))
                            .CategoriaID = CLng(cbbCategoria.List(cbbCategoria.ListIndex, 1))
                            .PagRec = cbbPagRec.List(cbbPagRec.ListIndex, 1)
                            .Valor = CCur(txbValor.Text)
                            .Requisitado = chbRequisita.Value
                        End With
                    
                        Valida = True
                        
                    End If
                    
                End If
            End If
        ElseIf cbbPagRec.List(cbbPagRec.ListIndex, 1) = "R" Then
            If cbbObra2.ListIndex = -1 Then
                MsgBox "Campo 'Obra' é obrigatório", vbCritical
                MultiPage1.Value = 1: cbbObra2.SetFocus
            Else
                With oLancamentoRapido
                    .Data = CDate(txbData.Text)
                    .ContaID = CLng(cbbConta.List(cbbConta.ListIndex, 1))
                    .CliForID = CLng(cbbObra2.List(cbbObra2.ListIndex, 1))
                    .CategoriaID = CLng(cbbCategoria.List(cbbCategoria.ListIndex, 1))
                    .PagRec = cbbPagRec.List(cbbPagRec.ListIndex, 1)
                    .Valor = CCur(txbValor.Text)
                    .Requisitado = chbRequisita.Value
                End With
            
                Valida = True
                
            End If
        End If
        
    End If

End Function
Private Sub lstPrincipal_Change()

    Dim n As Long

    If lstPrincipal.ListIndex > -1 Then
    
        btnExcluir.Enabled = True
        
        ' Carrega informações do lançamento
        oLancamentoRapido.Carrega (CLng(lstPrincipal.List(lstPrincipal.ListIndex, 0)))
        
        ' Preenche cabeçalho
        lblCabID.Caption = IIf(oLancamentoRapido.ID = 0, "", Format(oLancamentoRapido.ID, "0000000000"))
        lblCabData.Caption = oLancamentoRapido.Data
        
        'oFornecedor.Carrega oCompra.FornecedorID
        
        'lblCabFuncionario.Caption = oFuncionario.Funcionario
        
        ' Preenche campos
        txbData.Text = oLancamentoRapido.Data
        
        For n = 0 To cbbPagRec.ListCount - 1
            If cbbPagRec.List(n, 1) = oLancamentoRapido.PagRec Then
                cbbPagRec.ListIndex = n
                Exit For
            End If
        Next n
        
        txbValor.Text = Format(oLancamentoRapido.Valor, "#,##0.00")
        
        For n = 0 To cbbConta.ListCount - 1
            If CLng(cbbConta.List(n, 1)) = oLancamentoRapido.ContaID Then
                cbbConta.ListIndex = n
                Exit For
            End If
        Next n
        
        If oLancamentoRapido.PagRec = "P" Then
            For n = 0 To cbbFornecedor.ListCount - 1
                If CLng(cbbFornecedor.List(n, 1)) = oLancamentoRapido.CliForID Then
                    cbbFornecedor.ListIndex = n
                    Exit For
                End If
            Next n
        Else
            For n = 0 To cbbObra2.ListCount - 1
                If CLng(cbbObra2.List(n, 1)) = oLancamentoRapido.CliForID Then
                    cbbObra2.ListIndex = n
                    Exit For
                End If
            Next n
        End If
        
        Call cbbCategoriaPopular(cbbPagRec.List(cbbPagRec.ListIndex, 1))
        
        For n = 0 To cbbCategoria.ListCount - 1
            If CLng(cbbCategoria.List(n, 1)) = oLancamentoRapido.CategoriaID Then
                cbbCategoria.ListIndex = n
                Exit For
            End If
        Next n
        
        
        chbRequisita.Value = oLancamentoRapido.Requisitado
        
        If chbRequisita.Value = True Then
            Call lstRequisicoesPopular(oLancamentoRapido.RequisicaoID)
        End If
        
    End If

End Sub
Private Sub btnItemInclui_Click()

    Call AcaoItem("Incluir", True)

End Sub
Private Sub btnItemAltera_Click()

    Call AcaoItem("Alterar", True)

End Sub
Private Sub btnItemExclui_Click()

    Call AcaoItem("Excluir", True)

End Sub
Private Sub btnItemCancelar_Click()

    Call AcaoItem("Cancelar", False)
    
End Sub
Private Sub btnItemConfirmar_Click()

    Dim sDecisaoLancamento  As String
    Dim sDecisaoItem        As String
    Dim dVlrUnitario        As Double
    
    sDecisaoLancamento = Replace(btnConfirmar.Caption, "Confirmar ", "")
    sDecisaoItem = btnItemConfirmar.Caption
    
    ' Configura ListBox
    With lstRequisicoes
        .ColumnCount = 9
        .ColumnWidths = "0pt; 0pt; 85pt; 55pt; 18pt; 55pt; 55pt; 240pt; 60pt"
        ' Colunas
            ' 0 - Recno do produto na tbl_requisicoes_itens
            ' 1 - Códigos: Produto;Unidade de medida;Obra;Etapa
            ' 2 - Produto
            ' 3 - Quantidade do item
            ' 4 - Unidade de medida
            ' 5 - Preço unitário
            ' 6 - Preço total
            ' 7 - Obra
            ' 8 - Etapa
        .Font = "Consolas"
    End With
    
    If sDecisaoItem = "Incluir" Then

        If ValidaItem = True Then
            
            With lstRequisicoes
            
                .AddItem
                .List(.ListCount - 1, 0) = ""
                .List(.ListCount - 1, 1) = cbbProduto.List(cbbProduto.ListIndex, 1) & ";" & cbbUM.List(cbbUM.ListIndex, 1) & ";" & cbbObra.List(cbbObra.ListIndex, 1) & ";" & cbbEtapa.List(cbbEtapa.ListIndex, 1)
                .List(.ListCount - 1, 2) = cbbProduto.List(cbbProduto.ListIndex, 0)
                .List(.ListCount - 1, 3) = Space(9 - Len(Format(CDbl(txbQtde.Text), "#,##0.00"))) & Format(CDbl(txbQtde.Text), "#,##0.00")
                .List(.ListCount - 1, 4) = cbbUM.List(cbbUM.ListIndex, 0)
                
                dVlrUnitario = Round(CCur(txbTotal.Text) / CDbl(txbQtde.Text), 2)
                
                .List(.ListCount - 1, 5) = Space(9 - Len(Format(dVlrUnitario, "#,##0.00"))) & Format(dVlrUnitario, "#,##0.00")
                .List(.ListCount - 1, 6) = Space(9 - Len(Format(txbTotal.Text, "#,##0.00"))) & Format(txbTotal.Text, "#,##0.00")
                .List(.ListCount - 1, 7) = cbbObra.List(cbbObra.ListIndex, 0) ' & Space(30 - Len(cbbObra.List(cbbObra.ListIndex, 2))) & " | " & cbbObra.List(cbbObra.ListIndex, 3)
                .List(.ListCount - 1, 8) = cbbEtapa.List(cbbEtapa.ListIndex, 0)
                
            End With
            
        End If
    
    ElseIf sDecisaoItem = "Alterar" Then
        
        If ValidaItem = True Then
        
            With lstRequisicoes
            
                .List(.ListIndex, 0) = ""
                .List(.ListIndex, 1) = cbbProduto.List(cbbProduto.ListIndex, 1) & ";" & cbbUM.List(cbbUM.ListIndex, 1) & ";" & cbbObra.List(cbbObra.ListIndex, 1) & ";" & cbbEtapa.List(cbbEtapa.ListIndex, 1)
                .List(.ListIndex, 2) = cbbProduto.List(cbbProduto.ListIndex, 0)
                .List(.ListIndex, 3) = Space(9 - Len(Format(CDbl(txbQtde.Text), "#,##0.00"))) & Format(CDbl(txbQtde.Text), "#,##0.00")
                .List(.ListIndex, 4) = cbbUM.List(cbbUM.ListIndex, 0)
                
                dVlrUnitario = Round(CCur(txbTotal.Text) / CDbl(txbQtde.Text), 2)
                
                .List(.ListIndex, 5) = Space(9 - Len(Format(dVlrUnitario, "#,##0.00"))) & Format(dVlrUnitario, "#,##0.00")
                .List(.ListIndex, 6) = Space(9 - Len(Format(txbTotal.Text, "#,##0.00"))) & Format(txbTotal.Text, "#,##0.00")
                .List(.ListIndex, 7) = cbbObra.List(cbbObra.ListIndex, 0) '& Space(30 - Len(cbbObra.List(cbbObra.ListIndex, 0))) & " | " & cbbObra.List(cbbObra.ListIndex, 2)
                .List(.ListIndex, 8) = cbbEtapa.List(cbbEtapa.ListIndex, 0)
                
            End With
        
        End If
    ElseIf sDecisaoItem = "Excluir" Then
    
        lstRequisicoes.RemoveItem (lstRequisicoes.ListIndex)
         
    End If
    
    Call btnItemCancelar_Click
    
    Call TotalizaRequisicoes

End Sub
Private Sub AcaoItem(Acao As String, Habilitar As Boolean)
    
    btnItemConfirmar.Caption = Acao
    
    If Acao = "Incluir" Then
        lstRequisicoes.ListIndex = -1
        cbbObra.ListIndex = -1
        cbbEtapa.ListIndex = -1
        cbbProduto.ListIndex = -1
        cbbUM.ListIndex = -1
        txbQtde.Text = Format(0, "#,##0.00")
        txbUnitario.Text = Format(0, "#,##0.00")
        txbTotal.Text = Format(0, "#,##0.00")
    End If
    
    If Habilitar = True Then
        
        cbbObra.Enabled = Habilitar: lblObra.Enabled = Habilitar
        cbbEtapa.Enabled = Habilitar: lblEtapa.Enabled = Habilitar
        cbbProduto.Enabled = Habilitar: lblProduto.Enabled = Habilitar
        cbbUM.Enabled = Habilitar: lblUM.Enabled = Habilitar
        txbQtde.Enabled = Habilitar: lblQtde.Enabled = Habilitar
        cbbUM.Enabled = Habilitar: lblUM.Enabled = Habilitar
        txbUnitario.Enabled = Habilitar: lblUnitario.Enabled = Habilitar
        txbTotal.Enabled = Habilitar: lblTotal.Enabled = Habilitar
        
        btnItemInclui.Visible = Not Habilitar
        btnItemAltera.Visible = Not Habilitar
        btnItemExclui.Visible = Not Habilitar
        btnItemCancelar.Visible = Habilitar
        btnItemConfirmar.Visible = Habilitar
        lstRequisicoes.Enabled = Not Habilitar: lstRequisicoes.ForeColor = &H80000010
        btnConfirmar.Enabled = Not Habilitar
        btnCancelar.Enabled = Not Habilitar
    Else
        lstRequisicoes.ListIndex = -1
        cbbObra.Enabled = Habilitar: lblObra.Enabled = Habilitar: cbbObra.ListIndex = -1
        cbbEtapa.Enabled = Habilitar: lblEtapa.Enabled = Habilitar: cbbEtapa.ListIndex = -1
        cbbProduto.Enabled = Habilitar: lblProduto.Enabled = Habilitar: cbbProduto.ListIndex = -1
        cbbUM.Enabled = Habilitar: lblUM.Enabled = Habilitar: cbbUM.ListIndex = -1
        txbQtde.Enabled = Habilitar: lblQtde.Enabled = Habilitar: txbQtde.Text = Empty
        cbbUM.Enabled = Habilitar: lblUM.Enabled = Habilitar: cbbUM.ListIndex = -1
        txbUnitario.Enabled = Habilitar: lblUnitario.Enabled = Habilitar: txbUnitario.Text = Empty
        txbTotal.Enabled = Habilitar: lblTotal.Enabled = Habilitar: txbTotal.Text = Empty
        
        btnItemInclui.Visible = Not Habilitar
        btnItemAltera.Visible = Not Habilitar
        btnItemExclui.Visible = Not Habilitar
        btnItemCancelar.Visible = Habilitar
        btnItemConfirmar.Visible = Habilitar
        lstRequisicoes.Enabled = Not Habilitar: lstRequisicoes.ForeColor = &H80000008
        btnConfirmar.Enabled = Not Habilitar
        btnCancelar.Enabled = Not Habilitar
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
    Else
        ValidaItem = True
    End If

End Function
Private Sub lstRequisicoes_Change()

    Dim i As Integer
    
    Dim idxProduto  As Long
    Dim idxUM       As Long
    Dim idxObra     As Long
    Dim idxEtapa    As Long
    Dim arr()       As String
    
    
    If lstRequisicoes.ListIndex > -1 And btnItemConfirmar.Caption <> "Alterar" Then
    
        arr() = Split(lstRequisicoes.List(lstRequisicoes.ListIndex, 1), ";")
        
        idxProduto = CLng(arr(0))
        idxUM = CLng(arr(1))
        idxObra = CLng(arr(2))
        idxEtapa = CLng(arr(3))
    
        For i = 0 To cbbProduto.ListCount - 1
            If CLng(cbbProduto.List(i, 1)) = idxProduto Then
                cbbProduto.ListIndex = i: Exit For
            End If
        Next i
        
        For i = 0 To cbbUM.ListCount - 1
            If CLng(cbbUM.List(i, 1)) = idxUM Then
                cbbUM.ListIndex = i: Exit For
            End If
        Next i

        For i = 0 To cbbObra.ListCount - 1
            If CLng(cbbObra.List(i, 1)) = idxObra Then
                cbbObra.ListIndex = i: Exit For
            End If
        Next i
        
        For i = 0 To cbbEtapa.ListCount - 1
            If CLng(cbbEtapa.List(i, 1)) = idxEtapa Then
                cbbEtapa.ListIndex = i: Exit For
            End If
        Next i
        
        txbQtde.Text = lstRequisicoes.List(lstRequisicoes.ListIndex, 3)
        txbUnitario.Text = lstRequisicoes.List(lstRequisicoes.ListIndex, 5)
        txbTotal.Text = lstRequisicoes.List(lstRequisicoes.ListIndex, 6)
        
        btnItemAltera.Enabled = True
        btnItemExclui.Enabled = True
    End If
End Sub
Private Sub TotalizaRequisicoes()

    Dim cTotal As Currency
    Dim i As Integer
    
    For i = 0 To lstRequisicoes.ListCount - 1
        cTotal = cTotal + CCur(lstRequisicoes.List(i, 6))
    Next i
    
    lblTotalRequisicoes.Caption = Format(cTotal, "#,##0.00")
    txbValor.Text = Format(cTotal, "#,##0.00")

End Sub
Private Sub lstRequisicoesPopular(RequisicaoID As Long)

    Dim r               As New ADODB.Recordset
    Dim dVlrUnitario    As Currency

    sSQL = "SELECT * "
    sSQL = sSQL & "FROM tbl_requisicoes_itens "
    sSQL = sSQL & "WHERE requisicao_id = " & RequisicaoID & " "
    sSQL = sSQL & "ORDER BY r_e_c_n_o_"
    
    r.Open sSQL, cnn, adOpenStatic
    
    With lstRequisicoes
        .Clear
        .ColumnCount = 9
        .ColumnWidths = "0pt; 0pt; 85pt; 55pt; 18pt; 55pt; 55pt; 240pt; 60pt"
            ' Colunas
            ' 0 - Recno do produto na tbl_requisicoes_itens
            ' 1 - Códigos: Produto;Unidade de medida;Obra;Etapa
            ' 2 - Produto
            ' 3 - Quantidade do item
            ' 4 - Unidade de medida
            ' 5 - Preço unitário
            ' 6 - Preço total
            ' 7 - Obra
            ' 8 - Etapa
        .Font = "Consolas"
        
        Do Until r.EOF
            .AddItem
            
            oProduto.Carrega r.Fields("produto_id").Value
            oUM.Carrega r.Fields("um_id").Value
            oObra.Carrega r.Fields("obra_id").Value
            oEtapa.Carrega r.Fields("etapa_id").Value
            oCliente.Carrega oObra.ClienteID
            
            .List(.ListCount - 1, 0) = r.Fields("r_e_c_n_o_").Value
            .List(.ListCount - 1, 1) = r.Fields("produto_id").Value & ";" & r.Fields("um_id").Value & ";" & r.Fields("obra_id").Value & ";" & r.Fields("etapa_id").Value
            .List(.ListCount - 1, 2) = oProduto.Nome
            .List(.ListCount - 1, 3) = Space(9 - Len(Format(r.Fields("quantidade").Value, "#,##0.00"))) & Format(r.Fields("quantidade").Value, "#,##0.00")
            .List(.ListCount - 1, 4) = oUM.Abreviacao
            
            dVlrUnitario = Round(r.Fields("total").Value / r.Fields("quantidade").Value, 2)
            
            .List(.ListCount - 1, 5) = Space(9 - Len(Format(dVlrUnitario, "#,##0.00"))) & Format(dVlrUnitario, "#,##0.00")
            .List(.ListCount - 1, 6) = Space(9 - Len(Format(r.Fields("total").Value, "#,##0.00"))) & Format(r.Fields("total").Value, "#,##0.00")
            .List(.ListCount - 1, 7) = oObra.Bairro & ": " & oCliente.Nome
            .List(.ListCount - 1, 8) = oEtapa.Nome
            
            r.MoveNext
        Loop
    End With
    
    Set r = Nothing
    
    Call TotalizaRequisicoes
    
End Sub
Private Sub cbbCategoriaPopular(PagRec As String)
    
    Dim idx         As Integer
    Dim col         As New Collection
    Dim n           As Variant

    Set col = oCategoria.Listar("categoria, subcategoria, item_subcategoria", PagRec)
    
    'idx = cbbCategoria.ListIndex
    
    With cbbCategoria
        .Clear
        .ColumnCount = 4
        .ColumnWidths = "180pt; 0pt; 180pt; 100pt;"
    End With
    
    
    For Each n In col
        
        oCategoria.Carrega CLng(n)
    
        With cbbCategoria
            .AddItem
            .List(.ListCount - 1, 0) = oCategoria.Categoria & ": " & oCategoria.Subcategoria & IIf(oCategoria.ItemSubcategoria = "", "", ": " & oCategoria.ItemSubcategoria)
            .List(.ListCount - 1, 1) = oCategoria.ID
            .List(.ListCount - 1, 2) = oCategoria.Subcategoria
            .List(.ListCount - 1, 3) = oCategoria.ItemSubcategoria
        End With
        
    Next n
    
    cbbCategoria.ListIndex = -1

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
Private Sub cbbProduto_AfterUpdate()

    Dim vbResposta As VbMsgBoxResult
    Dim idx As Integer
    Dim n As Integer
    
    If cbbProduto.ListIndex = -1 And cbbProduto.Text <> "" Then
        
        vbResposta = MsgBox("Este produto não existe, deseja cadastrá-lo?", vbQuestion + vbYesNo)
        
        If vbResposta = vbYes Then
            
            oProduto.Nome = RTrim(cbbProduto.Text)
            oProduto.Inclui
            idx = oProduto.ID
            Call cbbProdutoPopular
            
            For n = 0 To cbbProduto.ListCount - 1
                If CInt(cbbProduto.List(n, 1)) = idx Then
                    cbbProduto.ListIndex = n
                    Exit For
                End If
            Next n
        Else
            cbbProduto.ListIndex = -1
        End If

    End If

End Sub
Private Sub cbbUM_AfterUpdate()

    Dim vbResposta As VbMsgBoxResult
    Dim idx As Integer
    Dim n As Integer
    
    If cbbUM.ListIndex = -1 And cbbUM.Text <> "" Then
        
        vbResposta = MsgBox("Esta unidade de medida não existe. Deseja cadastrá-la?", vbQuestion + vbYesNo)
        
        If vbResposta = vbYes Then
        
            oUM.Abreviacao = RTrim(cbbUM.Text)
            oUM.Nome = ""
            oUM.Inclui
            
            idx = oUM.ID
            
            Call cbbUMPopular
            
            For n = 0 To cbbUM.ListCount - 1
                If CInt(cbbUM.List(n, 1)) = idx Then
                    cbbUM.ListIndex = n
                    Exit For
                End If
            Next n
        Else
            cbbUM.ListIndex = -1
        End If
        
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
    
    Call lstPrincipalPopular(scrPagina.Value)

End Sub
Private Sub scrPaginaAtualiza(AfetouBanco As Boolean, Optional Decisao As Variant)

    If Not IsNull(Decisao) Then
        If Decisao = "Inclusão" Then
            With scrPagina
                .Max = myRst.PageCount
                .Value = myRst.PageCount
            End With
            Call scrPagina_Change
        ElseIf Decisao = "Exclusão" Then
            If scrPagina.Max = myRst.PageCount Then
                Call scrPagina_Change
            ElseIf myRst.PageCount = 0 Then
                Call scrPagina_Change
            Else
                With scrPagina
                    .Value = myRst.PageCount
                    .Max = myRst.PageCount
                End With
            End If
            
        End If
    Else
        With scrPagina
            .Max = myRst.PageCount
            .Value = myRst.PageCount
        End With
        Call scrPagina_Change
    End If
    
End Sub
Private Sub txbQtde_AfterUpdate()

    If txbTotal.Text = Empty Then
        txbQtde.Text = Format(0, "#,##0.00")
    Else
        txbTotal.Text = Format(CDbl(txbQtde.Text) * CCur(txbUnitario.Text), "#,##0.00")
    End If

End Sub
Private Sub txbUnitario_AfterUpdate()

    If txbUnitario.Text = Empty Then
        txbUnitario.Text = Format(0, "#,##0.00")
    Else
        txbTotal.Text = Format(CDbl(txbQtde.Text) * CCur(txbUnitario.Text), "#,##0.00")
    End If

End Sub
Private Sub txbTotal_AfterUpdate()

    If txbTotal.Text = Empty Then
        txbTotal.Text = Format(0, "#,##0.00")
    Else
        txbUnitario.Text = Format(CDbl(txbTotal.Text) / CCur(txbQtde.Text), "#,##0.00")
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
            
            idx = oEtapa.ID
            
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
