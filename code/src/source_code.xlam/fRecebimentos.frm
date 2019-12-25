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
Private myRst               As ADODB.RecordSet
Private lPagina             As Long

Private Const sTable As String = "tbl_recebimentos"
Private Const sCampoOrderBy As String = "data"

Private Sub UserForm_Initialize()

    Call cbbObraPopular
    Call cbbContaPopular
    Call EventosCampos
    
    Set myRst = New ADODB.RecordSet
    Set myRst = oRecebimento.RecordSet
    
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
    Set oRecebimento = Nothing
    Call Desconecta
    
End Sub

Private Sub lstPrincipalPopular(Pagina As Long)

    Dim lPosicao    As Long
    Dim lCount      As Long
    
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
        
        frmFormaPagamento.Enabled = False
        txbValorPgto.Enabled = False: lblValorPgto.Enabled = False
        cbbConta.Enabled = False: lblConta.Enabled = False
        
        lblHdValorPgto.Enabled = False
        lblHdConta.Enabled = False
        
        frmTitulo.Enabled = False
        txbValorBaixar.Enabled = False: lblValorBaixar.Enabled = False
        
        lblExtrato.Enabled = False
        lblHdVencimento.Enabled = False
        lblHdValorTitulo.Enabled = False
        lblHdValorBaixado.Enabled = False
        lblHdValorBaixar.Enabled = False
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

        frmFormaPagamento.Enabled = True
        lblHdValorPgto.Enabled = True
        lblHdConta.Enabled = True
        lstRecebimentos.Enabled = True: lstRecebimentos.ForeColor = &H80000008
        btnPgtoInclui.Visible = True
        btnPgtoAltera.Visible = True
        btnPgtoExclui.Visible = True

        lblExtrato.Enabled = True
        
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
    End If

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
            
    End If
    
    lstPrincipal.Enabled = False
    lstPrincipal.ForeColor = &H80000010
    
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

    Dim r       As New ADODB.RecordSet
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
            .ColumnCount = 7
            .ColumnWidths = "60pt; 65pt; 65pt; 65pt; 60pt; 0pt; 0pt;"
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
                    .List(.ListCount - 1, 4) = r.Fields("observacao").Value
                    .List(.ListCount - 1, 5) = r.Fields("r_e_c_n_o_").Value
                    
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
            .ColumnCount = 7
            .ColumnWidths = "60pt; 65pt; 65pt; 65pt; 60pt; 0pt; 0pt;"
            .Font = "Consolas"
            
            Do Until r.EOF
            
                oTituloReceber.Carrega r.Fields("titulo_id").Value
                
                .AddItem
                
                .List(.ListCount - 1, 0) = r.Fields("titulo_vencimento").Value
                .List(.ListCount - 1, 1) = Space(12 - Len(Format(r.Fields("titulo_valor").Value, "#,##0.00"))) & Format(r.Fields("titulo_valor").Value, "#,##0.00")
                .List(.ListCount - 1, 2) = Space(12 - Len(Format(r.Fields("valor_baixado").Value, "#,##0.00"))) & Format(r.Fields("valor_baixado").Value, "#,##0.00")
                .List(.ListCount - 1, 3) = Space(12 - Len(Format(0, "#,##0.00"))) & Format(0, "#,##0.00")
                .List(.ListCount - 1, 4) = oTituloReceber.Observacao
                .List(.ListCount - 1, 5) = r.Fields("r_e_c_n_o_").Value
                
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
        frmTitulo.Enabled = True
        lblHdVencimento.Enabled = True
        lblHdValorTitulo.Enabled = True
        lblHdValorBaixado.Enabled = True
        lblHdValorBaixar.Enabled = True
        lblHdObservacao.Enabled = True
        lstTitulos.Enabled = True: lstTitulos.ForeColor = &H80000008
        frmTotalBaixar.Visible = True
    Else
        frmTitulo.Enabled = False
        lblHdVencimento.Enabled = False
        lblHdValorTitulo.Enabled = False
        lblHdValorBaixado.Enabled = False
        lblHdValorBaixar.Enabled = False
        lblHdObservacao.Enabled = False
        
        For i = 0 To lstTitulos.ListCount - 1
            lstTitulos.List(i, 3) = Space(9 - Len(Format(0, "#,##0.00"))) & Format(0, "#,##0.00")
        Next i
        
        lstTitulos.Enabled = False: lstTitulos.ForeColor = &H80000010
        frmTotalBaixar.Visible = False
    End If

    frmFormaPagamento.Visible = True
    lblHdValorPgto.Visible = True
    lblHdConta.Visible = True
    lstRecebimentos.Visible = True

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
        MultiPage1.Value = 2: cbbConta.SetFocus: Exit Function
    ElseIf txbValorPgto.Text = Empty Then
        MsgBox "Campo 'Valor pgto.' é obrigatório", vbCritical
        MultiPage1.Value = 2: txbValorPgto.SetFocus: Exit Function
    Else
        ValidaPgto = True
    End If
End Function
Private Sub AcaoPgto(Acao As String, Habilitar As Boolean)
    
    btnPgtoConfirmar.Caption = Acao
    
    If Acao = "Incluir" Then
        
        lstRecebimentos.ListIndex = -1
        cbbConta.ListIndex = -1
        
        If lblTotalBaixar.Caption <> "" Then
            txbValorPgto.Text = Format(lblTotalBaixar.Caption, "#,##0.00")
        Else
            txbValorPgto.Text = Format(0, "#,##0.00")
        End If
        
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
        
            ' Cabeçalho da compra
            If sDecisao = "Inclusão" Then
                oRecebimento.Inclui
            End If
            
            ' Itens do pagamento
            For i = 0 To lstTitulos.ListCount - 1
            
                If sDecisao = "Inclusão" Then
                    
                    If CCur(lstTitulos.List(i, 3)) > 0 Then
                        With oRecebimentoItem
                        
                            oTituloReceber.Carrega CLng(lstTitulos.List(i, 5))
                            
                            .RecebimentoID = oRecebimento.ID
                            .TituloID = oTituloReceber.Recno
                            .ValorBaixado = CCur(lstTitulos.List(i, 3))
                            .DataBaixa = oRecebimento.Data
                            .ObraID = oTituloReceber.ObraID
                            .TituloValor = oTituloReceber.Valor
                            .TituloData = oTituloReceber.Data
                            .TituloVencimento = oTituloReceber.Vencimento
                                                        
                            .Inclui
                        End With
                    End If
                ElseIf sDecisao = "Exclusão" Then
                    With oRecebimentoItem
                        .Recno = CLng(lstTitulos.List(i, 5))
                        .Exclui .Recno
                    End With
                End If
            Next i
            
            ' Itens da forma de recebimento
            For i = 0 To lstRecebimentos.ListCount - 1

                If sDecisao = "Inclusão" Then
                
                    With oContaMovimento
                        .ContaID = CLng(lstRecebimentos.List(i, 2))
                        .CliForID = oRecebimento.ObraID
                        
                        .Data = oRecebimento.Data
                        .PagRec = "R"
                        .Valor = CCur(lstRecebimentos.List(i, 0))
                        .TabelaOrigem = "tbl_recebimentos"
                        .RecnoOrigem = oRecebimento.ID
                        
                        oObra.Carrega oTituloReceber.ObraID
                        
                        .CategoriaID = oObra.CategoriaID
                    
                        .Inclui
                    End With
                    
                ElseIf sDecisao = "Exclusão" Then
                
                    With oContaMovimento
                        .Recno = CLng(lstRecebimentos.List(i, 3))
                        .Exclui .Recno
                    End With
                    
                End If

            Next i
            
            If sDecisao = "Exclusão" Then
                oRecebimento.Exclui oRecebimento.ID
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
            Set myRst = oRecebimento.RecordSet
        
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
Private Function Valida(Decisao As String) As Boolean
    
    Valida = False
    
    If Decisao = "Inclusão" Then
        If txbData.Text = Empty Then
            MsgBox "Campo 'Data' é obrigatório", vbCritical
            MultiPage1.Value = 1: txbData.SetFocus: Exit Function
        ElseIf cbbObra.ListIndex = -1 Then
            MsgBox "Campo 'Fornecedor' é obrigatório", vbCritical
            MultiPage1.Value = 1: cbbObra.SetFocus: Exit Function
        ElseIf optManual.Value = False And optAutomatico.Value = False Then
            MsgBox "Escolha o tipo de pagamento", vbCritical
            MultiPage1.Value = 2: optManual.SetFocus: Exit Function
        ElseIf lblTotalPagamentos.Caption = "" Then
            MsgBox "Não há pagamentos apontados", vbCritical
            MultiPage1.Value = 2: btnPgtoInclui.SetFocus: Exit Function
        Else
            
            If optManual.Value = True And lblTotalBaixar.Caption = "" Then
                MsgBox "Você precisa informar o valor que será baixado de cada título.", vbCritical
                MultiPage1.Value = 2: lblExtrato.SetFocus: Exit Function
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

    If lstTitulos.ListIndex > -1 Then
        txbValorBaixar.Enabled = True: lblValorBaixar.Enabled = True
        txbValorBaixar.Text = lstTitulos.List(lstTitulos.ListIndex, 3)
    End If

End Sub
Private Sub txbValorBaixar_AfterUpdate()

    Dim cSaldo As Currency
    Dim cBaixar As Currency
    
    
    
    If lstTitulos.ListIndex > -1 Then
        
        cSaldo = CCur(lstTitulos.List(lstTitulos.ListIndex, 1)) - CCur(lstTitulos.List(lstTitulos.ListIndex, 2))
        cBaixar = CCur(txbValorBaixar.Text)
        
        If cBaixar > cSaldo Then
            lstTitulos.List(lstTitulos.ListIndex, 3) = Space(12 - Len(Format(cSaldo, "#,##0.00"))) & Format(cSaldo, "#,##0.00")
        Else
            lstTitulos.List(lstTitulos.ListIndex, 3) = Space(12 - Len(Format(cBaixar, "#,##0.00"))) & Format(cBaixar, "#,##0.00")
        End If

        txbValorBaixar.Enabled = False: lblValorBaixar.Enabled = False: txbValorBaixar.Text = Format(0, "#,##0.00")
        lstTitulos.ListIndex = -1
        
        Call TotalizaBaixar
        
    End If

End Sub
Private Sub TotalizaBaixar()

    Dim cTotal As Currency
    Dim i As Integer
    
    For i = 0 To lstTitulos.ListCount - 1
        cTotal = cTotal + CCur(lstTitulos.List(i, 3))
    Next i
    
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
        Call lstRecebimentosPopular(oRecebimento.ID)
    
    End If

End Sub
Private Sub lstRecebimentosPopular(RecebimentoID As Long)

    Dim r       As New ADODB.RecordSet
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

