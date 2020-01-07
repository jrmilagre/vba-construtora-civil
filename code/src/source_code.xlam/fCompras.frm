VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fCompras 
   Caption         =   ":: Cadastro de Compras ::"
   ClientHeight    =   10560
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13320
   OleObjectBlob   =   "fCompras.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private oCompra             As New cCompra
Private oFornecedor         As New cFornecedor
Private oProduto            As New cProduto
Private oCompraItem         As New cCompraItem
Private oTituloPagar        As New cTituloPagar
Private oUM                 As New cUnidadeMedida
Private oCategoria          As New cCategoria

Private colControles        As New Collection
Private bListBoxOrdenando   As Boolean
Private myRst               As ADODB.Recordset
Private bChangeScrPag       As Boolean

Private Const sTable As String = "tbl_compras"
Private Const sCampoOrderBy As String = "data"

Private Sub UserForm_Initialize()

    Call cbbFornecedorPopular
    Call cbbProdutoPopular
    Call cbbUMPopular
    Call cbbCategoriaPopular
    
    Call cbbFltFornecedorPopular
    
    Call EventosCampos
    
    Call btnFiltrar_Click
    
    Call btnCancelar_Click

End Sub
Private Sub UserForm_Terminate()
    
    ' Destrói objeto da classe cProduto
    Set oCompra = Nothing
    Set oFornecedor = Nothing
    Set oProduto = Nothing
    Set oCompraItem = Nothing
    Set oTituloPagar = Nothing
    Set oUM = Nothing
    Set oCategoria = Nothing
    
    Set myRst = Nothing
    
    Call Desconecta
    
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
Private Sub lstItens_Change()

    Dim i As Integer
    
    If lstItens.ListIndex > -1 And btnItemConfirmar.Caption <> "Alterar" Then
    
        For i = 0 To cbbProduto.ListCount - 1
            If cbbProduto.List(i, 1) = lstItens.List(lstItens.ListIndex, 1) Then
                cbbProduto.ListIndex = i: Exit For
            End If
        Next i
        
        For i = 0 To cbbUM.ListCount - 1
            If cbbUM.List(i, 1) = lstItens.List(lstItens.ListIndex, 7) Then
                cbbUM.ListIndex = i: Exit For
            End If
        Next i
        
        txbQtde.Text = lstItens.List(lstItens.ListIndex, 2)
        txbUnitario.Text = lstItens.List(lstItens.ListIndex, 3)
        txbTotal.Text = lstItens.List(lstItens.ListIndex, 4)
        
        btnItemAltera.Enabled = True
        btnItemExclui.Enabled = True
    End If
End Sub
Private Sub btnConfirmar_Click()
    
    Dim vbResposta As VBA.VbMsgBoxResult
    Dim sDecisao As String
    Dim i As Integer
    
    sDecisao = Replace(btnConfirmar.Caption, "Confirmar ", "")
    
    If Valida(sDecisao) = True Then
    
        vbResposta = MsgBox("Deseja realmente fazer a " & sDecisao & "?", vbYesNo + vbQuestion, "Pergunta")
        
        If vbResposta = vbYes Then
        
            ' Cabeçalho da compra
            If sDecisao = "Inclusão" Then
                oCompra.Inclui
            
                ' Itens das compras
                For i = 0 To lstItens.ListCount - 1
                
                    With oCompraItem
                        .ProdutoID = CLng(lstItens.List(i, 1))
                        .Quantidade = CDbl(lstItens.List(i, 2))
                        .UmID = CLng(lstItens.List(i, 7))
                        .Unitario = CDbl(lstItens.List(i, 3))
                        .Total = CCur(lstItens.List(i, 4))
                        .Data = oCompra.Data
                        .FornecedorID = oCompra.FornecedorID
                        .CompraID = oCompra.ID
                        
                        If sDecisao = "Inclusão" Then
                            .Inclui
                        ElseIf sDecisao = "Exclusão" Then
                            .Recno = CLng(lstItens.List(i, 5))
                            .Exclui .Recno
                        End If
                        
                    End With
                    
                Next i
                
                ' Títulos das compras
                For i = 0 To lstTitulos.ListCount - 1
                
                    With oTituloPagar
                        .CompraID = oCompra.ID
                        .FornecedorID = oCompra.FornecedorID
                        .Observacao = lstTitulos.List(i, 2)
                        .Vencimento = CDate(lstTitulos.List(i, 0))
                        .Valor = CCur(lstTitulos.List(i, 1))
                        .Data = oCompra.Data

                        If sDecisao = "Inclusão" Then
                            .Inclui

                        ElseIf sDecisao = "Exclusão" Then
                            .Recno = CLng(lstItens.List(i, 5))
                            .Exclui .Recno
                        End If
                        
                    End With
                    
                Next i
            
            ElseIf sDecisao = "Exclusão" Then
            
                oCompraItem.Exclui oCompra.ID
                oTituloPagar.Exclui oCompra.ID
                oCompra.Exclui oCompra.ID
                
            End If
            
            Call btnFiltrar_Click
            
            ' Exibe mensagem de sucesso na decisão tomada (inclusão, alteração ou exclusão do registro).
            MsgBox sDecisao & " realizada com sucesso.", vbInformation, sDecisao & " de registro"
            
            Call btnCancelar_Click
            
        ElseIf vbResposta = vbNo Then
        
            Call btnCancelar_Click
            
        End If
        
    Else
    
        If sDecisao = "Exclusão" Then
        
            Call btnCancelar_Click
            
        End If
    End If
    
End Sub
Private Sub btnIncluir_Click()
    Call PosDecisaoTomada("Inclusão")
End Sub
Private Sub btnExcluir_Click()
    Call PosDecisaoTomada("Exclusão")
End Sub
Private Sub PosDecisaoTomada(Decisao As String)

    btnCancelar.Visible = True: btnConfirmar.Visible = True
    btnConfirmar.Caption = "Confirmar " & Decisao
    btnCancelar.Caption = "Cancelar " & Decisao
    
    btnIncluir.Visible = False: btnExcluir.Visible = False
    
    MultiPage1.Value = 1
    
    If Decisao = "Inclusão" Then
        MultiPage1.Value = 1
        MultiPage1.Pages(0).Enabled = False
        Call Campos("Limpar")
        Call Campos("Habilitar")
        txbData.Text = Date
        cbbFornecedor.SetFocus
    End If
    
End Sub
Private Sub lstPrincipalPopular(Pagina As Long)

    Dim lPosicao    As Long
    Dim lCount      As Long
    
    myRst.AbsolutePage = Pagina
    
    With lstPrincipal
        .Clear
        .ColumnCount = 3 ' Funcionário, ID, Empresa, Filial
        .ColumnWidths = "55pt; 55pt; 180pt;"
        .Enabled = True
        .Font = "Consolas"
        
        lCount = 1
        
        While Not myRst.EOF = True And lCount <= myRst.PageSize

            .AddItem

            oFornecedor.Carrega myRst.Fields("fornecedor_id").Value

            .List(.ListCount - 1, 0) = Format(myRst.Fields("id").Value, "0000000000")
            .List(.ListCount - 1, 1) = myRst.Fields("data").Value
            .List(.ListCount - 1, 2) = oFornecedor.Nome
            
'            .List(.ListCount - 1, 4) = oEmpresa.Empresa & IIf(oEmpresa.Filial = "", "", " : " & oEmpresa.Filial)
'            .List(.ListCount - 1, 5) = myRst.Fields("status").Value
'            .List(.ListCount - 1, 6) = Space(2 - Len(Format(myRst.Fields("count_exames").Value, "00"))) & Format(myRst.Fields("count_exames").Value, "00")
'            .List(.ListCount - 1, 7) = Space(6 - Len(Format(myRst.Fields("sum_preco").Value, "#,##0.00"))) & Format(myRst.Fields("sum_preco").Value, "#,##0.00")

            lCount = lCount + 1
            myRst.MoveNext
            
        Wend

    End With
   
    ' Posiciona scroll de navegação em páginas
    lblPaginaAtual.Caption = Pagina
    lblNumeroPaginas.Caption = myRst.PageCount
    bChangeScrPag = False: scrPagina.Value = CLng(lblPaginaAtual.Caption): bChangeScrPag = True
    
    ' Trata os botões de navegação
    Call TrataBotoesNavegacao

End Sub
Private Sub btnData_Click()
    dtDate = IIf(txbData.Text = Empty, Date, txbData.Text)
    txbData.Text = GetCalendario
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
Private Sub lstPrincipal_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    MultiPage1.Value = 1
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

Private Sub lstPrincipal_Change()

    Dim n As Long

    If lstPrincipal.ListIndex > -1 Then
    
        btnExcluir.Enabled = True
        
        ' Carrega informações do lançamento
        oCompra.Carrega (CLng(lstPrincipal.List(lstPrincipal.ListIndex, 0)))
        
        ' Preenche cabeçalho
        lblCabID.Caption = IIf(oCompra.ID = 0, "", Format(oCompra.ID, "0000000000"))
        lblCabData.Caption = oCompra.Data
        
        oFornecedor.Carrega oCompra.FornecedorID
        
        'lblCabFuncionario.Caption = oFuncionario.Funcionario
        
        ' Preenche campos
        txbData.Text = oCompra.Data
                
        For n = 0 To cbbFornecedor.ListCount - 1
            If CLng(cbbFornecedor.List(n, 1)) = oCompra.FornecedorID Then
                cbbFornecedor.ListIndex = n
                Exit For
            End If
        Next n
        
        For n = 0 To cbbCategoria.ListCount - 1
            If CLng(cbbCategoria.List(n, 1)) = oCompra.CategoriaID Then
                cbbCategoria.ListIndex = n
                Exit For
            End If
        Next n
        
        Call lstItensPopular(CLng(lblCabID.Caption))
        Call lstTitulosPopular(CLng(lblCabID.Caption))
    End If

End Sub
Private Sub lstItensPopular(CompraID As Long)

    Dim r           As New ADODB.Recordset
    Dim cUnitario   As Currency

    sSQL = "SELECT * "
    sSQL = sSQL & "FROM tbl_compras_itens "
    sSQL = sSQL & "WHERE compra_id = " & CompraID
    
    r.Open sSQL, cnn, adOpenStatic
    
    With lstItens
        .Clear
        .ColumnCount = 8
        .ColumnWidths = "200pt; 0pt; 60pt; 60pt; 60pt; 0pt; 40pt; 0pt"
        .Font = "Consolas"
        
        Do Until r.EOF
            .AddItem
            
            oProduto.Carrega r.Fields("produto_id").Value
            oUM.Carrega r.Fields("um_id").Value
            
            .List(.ListCount - 1, 0) = oProduto.Nome
            .List(.ListCount - 1, 1) = r.Fields("produto_id").Value
            .List(.ListCount - 1, 2) = Space(9 - Len(Format(r.Fields("quantidade").Value, "#,##0.00"))) & Format(r.Fields("quantidade").Value, "#,##0.00")
            
            cUnitario = r.Fields("total").Value / r.Fields("quantidade").Value
            
            .List(.ListCount - 1, 3) = Space(9 - Len(Format(cUnitario, "#,##0.00"))) & Format(cUnitario, "#,##0.00")
            .List(.ListCount - 1, 4) = Space(9 - Len(Format(r.Fields("total").Value, "#,##0.00"))) & Format(r.Fields("total").Value, "#,##0.00")
            .List(.ListCount - 1, 5) = r.Fields("r_e_c_n_o_").Value
            .List(.ListCount - 1, 6) = oUM.Abreviacao
            .List(.ListCount - 1, 7) = r.Fields("um_id").Value
            
            r.MoveNext
        Loop
    End With
    
    Set r = Nothing
    
    Call TotalizaItens
    
End Sub
Private Sub TotalizaItens()

    Dim cTotal As Currency
    Dim i As Integer
    
    For i = 0 To lstItens.ListCount - 1
        cTotal = cTotal + CCur(lstItens.List(i, 4))
    Next i
    
    lblTotalItens.Caption = Format(cTotal, "#,##0.00")

End Sub
Private Sub btnCancelar_Click()

    btnIncluir.Visible = True: btnExcluir.Visible = True
    btnConfirmar.Visible = False: btnCancelar.Visible = False
    
    lstPrincipal.ListIndex = -1

    Call Campos("Limpar")
    Call Campos("Desabilitar")

    btnExcluir.Enabled = False
    btnIncluir.SetFocus
    
    MultiPage1.Value = 0

End Sub
Private Sub Campos(Acao As String)

    If Acao = "Desabilitar" Then
        txbData.Enabled = False: lblData.Enabled = False: btnData.Enabled = False
        cbbFornecedor.Enabled = False: lblFornecedor.Enabled = False
        cbbCategoria.Enabled = False: lblCategoria.Enabled = False
        
        frmItem.Enabled = False
        lblHdProduto.Enabled = False
        lblHdQuant.Enabled = False
        lblHdUnitario.Enabled = False
        lblHdTotal.Enabled = False
        lblHdUM.Enabled = False
        Call btnItemCancelar_Click
        btnItemInclui.Visible = False
        btnItemAltera.Visible = False
        btnItemExclui.Visible = False
        lstItens.Enabled = False: lstItens.ForeColor = &H80000010
        
        frmTitulo.Enabled = False
        lblHdVencimento.Enabled = False
        lblHdValor.Enabled = False
        lblHdObservacao.Enabled = False

        Call btnTituloCancelar_Click
        
        btnTituloInclui.Visible = False
        btnTituloAltera.Visible = False
        btnTituloExclui.Visible = False
        lstTitulos.Enabled = False: lstTitulos.ForeColor = &H80000010
        
        MultiPage1.Pages(0).Enabled = True
        
    ElseIf Acao = "Habilitar" Then
        txbData.Enabled = True: lblData.Enabled = True: btnData.Enabled = True
        cbbFornecedor.Enabled = True: lblFornecedor.Enabled = True
        cbbCategoria.Enabled = True: lblCategoria.Enabled = True
        frmItem.Enabled = True
        
        lstItens.Enabled = True: lstItens.ForeColor = &H80000008
        lblHdProduto.Enabled = True
        lblHdProduto.Enabled = True
        lblHdQuant.Enabled = True
        lblHdUnitario.Enabled = True
        lblHdTotal.Enabled = True
        lblHdUM.Enabled = True
        btnItemInclui.Visible = True
        btnItemAltera.Visible = True
        btnItemExclui.Visible = True
        
        frmTitulo.Enabled = True
        lstTitulos.Enabled = True: lstTitulos.ForeColor = &H80000008
        lblHdVencimento.Enabled = True
        lblHdValor.Enabled = True
        lblHdObservacao.Enabled = True

        btnTituloInclui.Visible = True
        btnTituloAltera.Visible = True
        btnTituloExclui.Visible = True
        
        MultiPage1.Pages(0).Enabled = False
        
    ElseIf Acao = "Limpar" Then
        lblCabID.Caption = ""
        lblCabData.Caption = ""
        txbData.Text = ""
        cbbFornecedor.ListIndex = -1
        cbbCategoria.ListIndex = -1
        
        lblTotalItens.Caption = ""
        lblTotalTitulos.Caption = ""
        
        lstItens.Clear
        lstTitulos.Clear
        
        lstPrincipal.ListIndex = -1
    End If

End Sub

Private Function Valida(Decisao As String) As Boolean
    
    Valida = False
    
    If Decisao = "Inclusão" Then
    
        If txbData.Text = Empty Then
            MsgBox "Campo 'Data' é obrigatório", vbCritical
            MultiPage1.Value = 1: txbData.SetFocus
        ElseIf cbbFornecedor.ListIndex = -1 Then
            MsgBox "Campo 'Fornecedor' é obrigatório", vbCritical
            MultiPage1.Value = 1: cbbFornecedor.SetFocus
        ElseIf cbbCategoria.ListIndex = -1 Then
            MsgBox "Campo 'Categoria' é obrigatório", vbCritical
            MultiPage1.Value = 1: cbbCategoria.SetFocus
        Else
            If lstItens.ListCount = 0 Then
                MsgBox "Não há itens apontados na compra", vbCritical
                MultiPage1.Value = 2: btnItemInclui.SetFocus
            ElseIf lstTitulos.ListCount = 0 Then
                MsgBox "Não há títulos apontados na compra", vbCritical
                MultiPage1.Value = 3: btnTituloInclui.SetFocus
            Else
                With oCompra
                    .Data = CDate(txbData.Text)
                    .FornecedorID = CLng(cbbFornecedor.List(cbbFornecedor.ListIndex, 1))
                    .CategoriaID = CLng(cbbCategoria.List(cbbCategoria.ListIndex, 1))
                End With
                
                Valida = True
            End If
        End If
    ElseIf Decisao = "Exclusão" Then
    
        If oCompra.ExisteRequisicao(oCompra.ID) = True Then
            Exit Function
        ElseIf oCompra.ExistePagamento(oCompra.ID) = True Then
            Exit Function
        Else
            Valida = True
        End If
    
    End If

End Function
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
Private Sub AcaoItem(Acao As String, Habilitar As Boolean)
    
    btnItemConfirmar.Caption = Acao
    
    If Acao = "Incluir" Then
        lstItens.ListIndex = -1
        cbbProduto.ListIndex = -1
        txbQtde.Text = Format(0, "#,##0.00")
        cbbUM.ListIndex = -1
        txbUnitario.Text = Format(0, "#,##0.00")
        txbTotal.Text = Format(0, "#,##0.00")
    End If
    
    If Habilitar = True Then
        
        cbbProduto.Enabled = Habilitar: lblProduto.Enabled = Habilitar
        txbQtde.Enabled = Habilitar: lblQtde.Enabled = Habilitar
        cbbUM.Enabled = Habilitar: lblUM.Enabled = Habilitar
        txbUnitario.Enabled = Habilitar: lblUnitario.Enabled = Habilitar
        txbTotal.Enabled = Habilitar: lblTotal.Enabled = Habilitar
        
        btnItemInclui.Visible = Not Habilitar
        btnItemAltera.Visible = Not Habilitar
        btnItemExclui.Visible = Not Habilitar
        btnItemCancelar.Visible = Habilitar
        btnItemConfirmar.Visible = Habilitar
        lstItens.Enabled = Not Habilitar: lstItens.ForeColor = &H80000010
        btnConfirmar.Enabled = Not Habilitar
        btnCancelar.Enabled = Not Habilitar
    Else
        lstItens.ListIndex = -1
        cbbProduto.Enabled = Habilitar: lblProduto.Enabled = Habilitar: cbbProduto.ListIndex = -1
        txbQtde.Enabled = Habilitar: lblQtde.Enabled = Habilitar: txbQtde.Text = Empty
        cbbUM.Enabled = Habilitar: lblUM.Enabled = Habilitar: cbbUM.ListIndex = -1
        txbUnitario.Enabled = Habilitar: lblUnitario.Enabled = Habilitar: txbUnitario.Text = Empty
        txbTotal.Enabled = Habilitar: lblTotal.Enabled = Habilitar: txbTotal.Text = Empty
        
        btnItemInclui.Visible = Not Habilitar
        btnItemAltera.Visible = Not Habilitar
        btnItemExclui.Visible = Not Habilitar
        btnItemCancelar.Visible = Habilitar
        btnItemConfirmar.Visible = Habilitar
        lstItens.Enabled = Not Habilitar: lstItens.ForeColor = &H80000008
        btnConfirmar.Enabled = Not Habilitar
        btnCancelar.Enabled = Not Habilitar
    End If
    
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
Private Sub btnItemConfirmar_Click()

    Dim sDecisaoLancamento  As String
    Dim sDecisaoItem        As String
    Dim cUnitario           As Currency
    
    sDecisaoLancamento = Replace(btnConfirmar.Caption, "Confirmar ", "")
    sDecisaoItem = btnItemConfirmar.Caption
    
    If sDecisaoItem = "Incluir" Then
    
        If ValidaItem = True Then
            
            With lstItens
                .ColumnCount = 8
                .ColumnWidths = "200pt; 0pt; 60pt; 60pt; 60pt; 0pt; 40pt; 0pt;"
                    ' COLUNAS:
                    '   0 - Código do produto
                    '   1 - Descrição do produto
                    '   2 - Quantidade
                    '   3 - Preço unitário
                    '   4 - Total
                    '   5 -
                    '   6 - Código da unidade de medida
                    '   7 - Descrição da unidade de medida
                .Font = "Consolas"
                .AddItem
                
                .List(.ListCount - 1, 0) = cbbProduto.List(cbbProduto.ListIndex, 0)
                .List(.ListCount - 1, 1) = cbbProduto.List(cbbProduto.ListIndex, 1)
                .List(.ListCount - 1, 2) = Space(9 - Len(Format(CDbl(txbQtde.Text), "#,##0.00"))) & Format(CDbl(txbQtde.Text), "#,##0.00")
                
                cUnitario = CCur(txbTotal.Text) / CDbl(txbQtde.Text)
                
                .List(.ListCount - 1, 3) = Space(9 - Len(Format(cUnitario, "#,##0.00"))) & Format(cUnitario, "#,##0.00")
                .List(.ListCount - 1, 4) = Space(9 - Len(Format(CCur(txbTotal.Text), "#,##0.00"))) & Format(CCur(txbTotal.Text), "#,##0.00")
                .List(.ListCount - 1, 6) = cbbUM.List(cbbUM.ListIndex, 0)
                .List(.ListCount - 1, 7) = cbbUM.List(cbbUM.ListIndex, 1)
                
            End With

        End If
        
    ElseIf sDecisaoItem = "Alterar" Then
    
        If ValidaItem = True Then
        
            With lstItens
                .List(.ListIndex, 0) = cbbProduto.List(cbbProduto.ListIndex, 0)
                .List(.ListIndex, 1) = cbbProduto.List(cbbProduto.ListIndex, 1)
                .List(.ListIndex, 2) = Space(9 - Len(Format(CDbl(txbQtde.Text), "#,##0.00"))) & Format(CDbl(txbQtde.Text), "#,##0.00")
                
                cUnitario = CCur(txbTotal.Text) / CDbl(txbQtde.Text)
                
                .List(.ListIndex, 3) = Space(9 - Len(Format(cUnitario, "#,##0.00"))) & Format(cUnitario, "#,##0.00")
                .List(.ListIndex, 4) = Space(9 - Len(Format(CCur(txbTotal.Text), "#,##0.00"))) & Format(CCur(txbTotal.Text), "#,##0.00")
                .List(.ListIndex, 6) = cbbUM.List(cbbUM.ListIndex, 0)
                .List(.ListIndex, 7) = cbbUM.List(cbbUM.ListIndex, 1)
            End With
        
        End If
        
    ElseIf sDecisaoItem = "Excluir" Then
    
        lstItens.RemoveItem (lstItens.ListIndex)
        
    End If
    
    Call btnItemCancelar_Click
    
    Call TotalizaItens
    
End Sub
Private Function ValidaItem() As Boolean
    ValidaItem = False
    If cbbProduto.ListIndex = -1 Then
        MsgBox "Campo 'Produto' é obrigatório", vbCritical
        MultiPage1.Value = 2: cbbProduto.SetFocus: Exit Function
    ElseIf txbQtde.Text = Empty Then
        MsgBox "Campo 'Quantidade' é obrigatório", vbCritical
        MultiPage1.Value = 2: txbQtde.SetFocus: Exit Function
    ElseIf txbUnitario.Text = Empty Then
        MsgBox "Campo 'Unitário' é obrigatório", vbCritical
        MultiPage1.Value = 2: txbUnitario.SetFocus: Exit Function
    Else
        ValidaItem = True
    End If
End Function
Private Sub btnTituloInclui_Click()

    Call AcaoTitulo("Incluir", True)

End Sub
Private Sub btnTituloAltera_Click()

    Call AcaoTitulo("Alterar", True)

End Sub
Private Sub btnTituloExclui_Click()

    Call AcaoTitulo("Excluir", True)

End Sub
Private Sub btnTituloCancelar_Click()

    Call AcaoTitulo("Cancelar", False)
    
End Sub
Private Sub AcaoTitulo(Acao As String, Habilitar As Boolean)
    
    btnTituloConfirmar.Caption = Acao
    
    If Acao = "Incluir" Then
        lstTitulos.ListIndex = -1
        txbVencimento.Text = Date
        txbValor.Text = Format(0, "#,##0.00")
        txbObservacao.Text = Empty
    End If
    
    If Habilitar = True Then
        
        txbVencimento.Enabled = Habilitar: lblVencimento.Enabled = Habilitar: btnVencimento.Enabled = Habilitar
        txbValor.Enabled = Habilitar: lblValor.Enabled = Habilitar
        txbObservacao.Enabled = Habilitar: lblObservacao.Enabled = Habilitar
        
        btnTituloInclui.Visible = Not Habilitar
        btnTituloAltera.Visible = Not Habilitar
        btnTituloExclui.Visible = Not Habilitar
        btnTituloCancelar.Visible = Habilitar
        btnTituloConfirmar.Visible = Habilitar
        
        lstTitulos.Enabled = Not Habilitar: lstTitulos.ForeColor = &H80000010
        btnConfirmar.Enabled = Not Habilitar
        btnCancelar.Enabled = Not Habilitar
    Else
        lstTitulos.ListIndex = -1

        txbVencimento.Enabled = Habilitar: lblVencimento.Enabled = Habilitar: txbVencimento.Text = Empty: btnVencimento.Enabled = Habilitar
        txbValor.Enabled = Habilitar: lblValor.Enabled = Habilitar: txbValor.Text = Empty
        txbObservacao.Enabled = Habilitar: lblObservacao.Enabled = Habilitar: txbObservacao.Text = Empty
        
        btnTituloInclui.Visible = Not Habilitar
        btnTituloAltera.Visible = Not Habilitar
        btnTituloExclui.Visible = Not Habilitar
        btnTituloCancelar.Visible = Habilitar
        btnTituloConfirmar.Visible = Habilitar
        lstTitulos.Enabled = Not Habilitar: lstTitulos.ForeColor = &H80000008
        btnConfirmar.Enabled = Not Habilitar
        btnCancelar.Enabled = Not Habilitar
    End If
    
End Sub
Private Sub btnVencimento_Click()
    dtDate = IIf(txbVencimento.Text = Empty, Date, txbVencimento.Text)
    txbVencimento.Text = GetCalendario
End Sub
Private Sub btnTituloConfirmar_Click()

    Dim sDecisaoLancamento  As String
    Dim sDecisaoTitulo      As String
    
    sDecisaoLancamento = Replace(btnConfirmar.Caption, "Confirmar ", "")
    sDecisaoTitulo = btnTituloConfirmar.Caption
    
    If sDecisaoTitulo = "Incluir" Then
    
        If ValidaTitulo = True Then
            
            With lstTitulos
                .ColumnCount = 4
                .ColumnWidths = "60pt; 60pt; 135pt; 0pt;"
                .Font = "Consolas"
                .AddItem
                
                .List(.ListCount - 1, 0) = txbVencimento.Text
                .List(.ListCount - 1, 1) = Space(9 - Len(Format(CDbl(txbValor.Text), "#,##0.00"))) & Format(CDbl(txbValor.Text), "#,##0.00")
                .List(.ListCount - 1, 2) = txbObservacao.Text
                
            End With
            
            Call btnTituloCancelar_Click

        End If
    ElseIf sDecisaoTitulo = "Alterar" Then
        If ValidaTitulo = True Then
            With lstTitulos
                .List(.ListIndex, 0) = txbVencimento.Text
                .List(.ListIndex, 1) = Space(9 - Len(Format(CDbl(txbValor.Text), "#,##0.00"))) & Format(CDbl(txbValor.Text), "#,##0.00")
                .List(.ListIndex, 2) = txbObservacao.Text
                .List(.ListIndex, 3) = .List(.ListIndex, 3)
            End With
            
            Call btnTituloCancelar_Click
            
        End If
    ElseIf sDecisaoTitulo = "Excluir" Then
        lstTitulos.RemoveItem (lstTitulos.ListIndex)
        Call btnTituloCancelar_Click
    End If
    
    Call TotalizaTitulos
End Sub
Private Function ValidaTitulo() As Boolean

    ValidaTitulo = False
    
    If txbVencimento.Text = Empty Then
        MsgBox "Campo 'Vencimento' é obrigatório", vbCritical
        MultiPage1.Value = 3: txbVencimento.SetFocus: Exit Function
    ElseIf txbValor.Text = Empty Then
        MsgBox "Campo 'Valor' é obrigatório", vbCritical
        MultiPage1.Value = 3: txbValor.SetFocus: Exit Function
    ElseIf txbObservacao.Text = Empty Then
        MsgBox "Campo 'Observação' é obrigatório", vbCritical
        MultiPage1.Value = 3: txbObservacao.SetFocus: Exit Function
    Else
        ValidaTitulo = True
    End If
    
End Function
Private Sub lstTitulosPopular(CompraID As Long)

    Dim r       As New ADODB.Recordset
    Dim cTotal As Currency

    sSQL = "SELECT * "
    sSQL = sSQL & "FROM tbl_titulos_pagar "
    sSQL = sSQL & "WHERE compra_id = " & CompraID
    
    r.Open sSQL, cnn, adOpenStatic
    
    With lstTitulos
        .Clear
        .ColumnCount = 4
        .ColumnWidths = "60pt; 60pt; 135pt; 0pt;"
        .Font = "Consolas"
        
        Do Until r.EOF
            .AddItem
            
            .List(.ListCount - 1, 0) = r.Fields("vencimento").Value
            .List(.ListCount - 1, 1) = Space(9 - Len(Format(r.Fields("valor").Value, "#,##0.00"))) & Format(r.Fields("valor").Value, "#,##0.00")
            .List(.ListCount - 1, 2) = r.Fields("observacao").Value
            .List(.ListCount - 1, 3) = r.Fields("r_e_c_n_o_").Value
            
            r.MoveNext
        Loop
    End With
    
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
Private Sub lstTitulos_Change()

    Dim n As Integer

    If lstTitulos.ListIndex > -1 And btnTituloConfirmar.Caption <> "Alterar" Then
        txbVencimento.Text = lstTitulos.List(lstTitulos.ListIndex, 0)
        txbValor.Text = lstTitulos.List(lstTitulos.ListIndex, 1)
        txbObservacao.Text = lstTitulos.List(lstTitulos.ListIndex, 2)
        
        oTituloPagar.Carrega CLng(lstTitulos.List(lstTitulos.ListIndex, 3))
        
        For n = 0 To cbbCategoria.ListCount - 1
            If CLng(cbbCategoria.List(n, 1)) = oTituloPagar.CategoriaID Then
                cbbCategoria.ListIndex = n
                Exit For
            End If
        Next n
        
        btnTituloAltera.Enabled = True
        btnTituloExclui.Enabled = True
    End If
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
Private Sub cbbCategoriaPopular()
    
    Dim idx         As Integer
    Dim col         As New Collection
    Dim n           As Variant

    Set col = oCategoria.Listar("categoria, subcategoria, item_subcategoria", "P")
    
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
Private Sub cbbFltFornecedorPopular()
    
    Dim idx         As Integer
    Dim col         As New Collection
    Dim n           As Variant

    Set col = oFornecedor.Listar("nome")
    
    idx = cbbFltFornecedor.ListIndex
    
    With cbbFltFornecedor
        .Clear
        .AddItem
        .List(.ListCount - 1, 0) = "***TODOS***"
        .List(.ListCount - 1, 1) = 0
    
    
        For Each n In col
            
            oFornecedor.Carrega CLng(n)
        
            .AddItem
            .List(.ListCount - 1, 0) = oFornecedor.Nome
            .List(.ListCount - 1, 1) = oFornecedor.ID
            
        Next n
        
        .ListIndex = 0
    
    End With

End Sub
Private Sub btnFiltrar_Click()

    Dim lFornecedorID As Long
    
    If cbbFltFornecedor.ListIndex = -1 Then
        lFornecedorID = 0
    Else
        lFornecedorID = CLng(cbbFltFornecedor.List(cbbFltFornecedor.ListIndex, 1))
    End If

    Set myRst = oCompra.Recordset(lFornecedorID)
    
    If myRst.PageCount > 0 Then
    
        myRst.AbsolutePage = myRst.PageCount
        
        bChangeScrPag = False
        
        With scrPagina
            .Max = myRst.PageCount
            .Value = myRst.PageCount
        End With
        
        Call lstPrincipalPopular(myRst.PageCount)
    Else
    
        lstPrincipal.Clear
        
    End If

End Sub
Private Sub TrataBotoesNavegacao()

    If CLng(lblPaginaAtual.Caption) = myRst.PageCount And CLng(lblPaginaAtual.Caption) > 1 Then
    
        btnPaginaInicial.Enabled = True
        btnPaginaAnterior.Enabled = True
        btnPaginaFinal.Enabled = False
        btnPaginaSeguinte.Enabled = False
        
    ElseIf CLng(lblPaginaAtual.Caption) < myRst.PageCount And CLng(lblPaginaAtual.Caption) = 1 Then
    
        btnPaginaInicial.Enabled = False
        btnPaginaAnterior.Enabled = False
        btnPaginaFinal.Enabled = True
        btnPaginaSeguinte.Enabled = True
        
    ElseIf CLng(lblPaginaAtual.Caption) = myRst.PageCount And CLng(lblPaginaAtual.Caption) = 1 Then
    
        btnPaginaInicial.Enabled = False
        btnPaginaAnterior.Enabled = False
        btnPaginaFinal.Enabled = False
        btnPaginaSeguinte.Enabled = False
    
    Else
    
        btnPaginaInicial.Enabled = True
        btnPaginaAnterior.Enabled = True
        btnPaginaFinal.Enabled = True
        btnPaginaSeguinte.Enabled = True
        
    End If

End Sub
Private Sub btnPaginaInicial_Click()
    
    Call lstPrincipalPopular(1)
    
End Sub
Private Sub btnPaginaAnterior_Click()

    Call lstPrincipalPopular(CLng(lblPaginaAtual.Caption) - 1)
    
End Sub
Private Sub btnPaginaSeguinte_Click()

    Call lstPrincipalPopular(CLng(lblPaginaAtual.Caption) + 1)

End Sub
Private Sub btnPaginaFinal_Click()

    Call lstPrincipalPopular(myRst.PageCount)
    
End Sub
Private Sub btnRegistroAnterior_Click()

        If lstPrincipal.ListIndex > 0 Then
        
            lstPrincipal.ListIndex = lstPrincipal.ListIndex - 1
            
        ElseIf lstPrincipal.ListIndex = 0 And CLng(lblPaginaAtual.Caption) > 1 Then
            
            Call lstPrincipalPopular(CLng(lblPaginaAtual.Caption) - 1)
            
            lstPrincipal.ListIndex = myRst.PageSize - 1
            
        ElseIf CLng(lblPaginaAtual.Caption) = 1 And lstPrincipal.ListIndex = 0 Then
        
            MsgBox "Primeiro registro"
            Exit Sub
            
        Else
        
            lstPrincipal.ListIndex = -1
            
        End If
        
End Sub
Private Sub btnRegistroSeguinte_Click()

    If lstPrincipal.ListIndex = -1 Then
        
        lstPrincipal.ListIndex = 0
    
    ElseIf lstPrincipal.ListIndex = myRst.PageSize - 1 And CLng(lblPaginaAtual.Caption) < myRst.PageCount Then
        
        Call lstPrincipalPopular(CLng(lblPaginaAtual.Caption) + 1)
        
        lstPrincipal.ListIndex = 0
        
    ElseIf CLng(lblPaginaAtual.Caption) = myRst.PageCount And (lstPrincipal.ListIndex + 1) = lstPrincipal.ListCount Then
    
        MsgBox "Último registro"
        Exit Sub
        
    Else
    
        lstPrincipal.ListIndex = lstPrincipal.ListIndex + 1
    
    End If
    
End Sub
Private Sub scrPagina_Change()

    If bChangeScrPag = True Then
        
        Call lstPrincipalPopular(scrPagina.Value)
        
    End If

End Sub
