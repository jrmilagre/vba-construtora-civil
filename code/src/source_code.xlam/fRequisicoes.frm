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
        .ColumnCount = 4
        .ColumnWidths = "100pt; 0pt; 100pt; 200pt;"
    End With
    
    
    For Each n In col
        
        oObra.Carrega CLng(n)
        
        oCliente.Carrega oObra.ClienteID
    
        With cbbObra
            .AddItem
            .List(.ListCount - 1, 0) = oObra.Bairro
            .List(.ListCount - 1, 1) = oObra.ID
            .List(.ListCount - 1, 2) = oCliente.Nome
            .List(.ListCount - 1, 3) = oObra.Endereco
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
'        lstPgtos.Clear
'        lstPrincipal.ListIndex = -1
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
Private Sub lstTitulosPopular()

    Dim r       As New ADODB.RecordSet
    Dim dQtdBx  As Currency
    Dim dSaldo  As Currency

    If lstPrincipal.ListIndex = -1 Then
    
        sSQL = "SELECT * "
        sSQL = sSQL & "FROM tbl_compras_itens "
        sSQL = sSQL & "WHERE "
        sSQL = sSQL & "requisitado = False"
        
        r.Open sSQL, cnn, adOpenStatic
    
        With lstTitulos
            .Clear
            .ColumnCount = 7
            .ColumnWidths = "60pt; 60pt; 60pt; 60pt; 60pt; 0pt; 0pt;"
            .Font = "Consolas"
            
            Do Until r.EOF
                
                dQtdBx = oCompraItem.GetQtdeBaixada(r.Fields("r_e_c_n_o_").Value)
                dSaldo = r.Fields("quantidade").Value - dQtdBx
                                
                If dSaldo > 0 Then
                                
                    .AddItem
                    
                    oFornecedor.Carrega r.Fields("fornecedor_id").Value
                    oProduto.Carrega r.Fields("produto_id").Value
                    
                    .AddItem
                    
                    .List(.ListCount - 1, 0) = r.Fields("data").Value
                    .List(.ListCount - 1, 1) = Format(r.Fields("compra_id").Value, "0000000000")
                    .List(.ListCount - 1, 2) = oFornecedor.Nome
                    .List(.ListCount - 1, 3) = oProduto.Nome
                    
                    .List(.ListCount - 1, 4) = Space(9 - Len(Format(r.Fields("quantidade").Value, "#,##0.00"))) & Format(r.Fields("quantidade").Value, "#,##0.00")
                    .List(.ListCount - 1, 5) = Space(9 - Len(Format(r.Fields("unitario").Value, "#,##0.00"))) & Format(r.Fields("unitario").Value, "#,##0.00")
                    .List(.ListCount - 1, 6) = Space(9 - Len(Format(r.Fields("total").Value, "#,##0.00"))) & Format(r.Fields("total").Value, "#,##0.00")
                    .List(.ListCount - 1, 7) = r.Fields("r_e_c_n_o_").Value
                    
                End If
                
                r.MoveNext
            Loop
            
        End With
        
    Else
        
        sSQL = "SELECT * "
        sSQL = sSQL & "FROM tbl_pagamentos_itens "
        sSQL = sSQL & "WHERE "
        sSQL = sSQL & "pagamento_id = " & oPagamento.ID
        
        r.Open sSQL, cnn, adOpenStatic
        
        With lstTitulos
            .Clear
            .ColumnCount = 7
            .ColumnWidths = "60pt; 60pt; 60pt; 60pt; 60pt; 0pt; 0pt;"
            .Font = "Consolas"
            
            Do Until r.EOF
            
                oFornecedor.Carrega r.Fields("fornecedor_id").Value
                oProduto.Carrega r.Fields("produto_id").Value
                
                .AddItem
                
                .List(.ListCount - 1, 0) = r.Fields("data").Value
                .List(.ListCount - 1, 1) = Format(r.Fields("compra_id").Value, "0000000000")
                .List(.ListCount - 1, 2) = oFornecedor.Nome
                .List(.ListCount - 1, 3) = oProduto.Nome
                
                .List(.ListCount - 1, 4) = Space(9 - Len(Format(r.Fields("quantidade").Value, "#,##0.00"))) & Format(r.Fields("quantidade").Value, "#,##0.00")
                .List(.ListCount - 1, 5) = Space(9 - Len(Format(r.Fields("unitario").Value, "#,##0.00"))) & Format(r.Fields("unitario").Value, "#,##0.00")
                .List(.ListCount - 1, 6) = Space(9 - Len(Format(r.Fields("total").Value, "#,##0.00"))) & Format(r.Fields("total").Value, "#,##0.00")
                .List(.ListCount - 1, 7) = r.Fields("r_e_c_n_o_").Value
                
                r.MoveNext
            Loop
            
        End With
        
    End If
    
    Set r = Nothing
    
    Call TotalizaTitulos
    
End Sub
