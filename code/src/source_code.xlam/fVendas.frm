VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fVendas 
   Caption         =   ":: Cadastro de Vendas ::"
   ClientHeight    =   10560
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13200
   OleObjectBlob   =   "fVendas.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fVendas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private oObra               As New cObra
Private oTipoObra           As New cTipoObra
Private oCliente            As New cCliente
Private oCategoria          As New cCategoria
Private oTituloReceber      As New cTituloReceber
Private oUF                 As New cUF

Private colControles        As New Collection
Private myRst               As ADODB.Recordset
Private bChangeScrPag       As Boolean

Private lPagina             As Long

Private Const sTable As String = "tbl_obras"
Private Const sCampoOrderBy As String = "endereco"

Private Sub UserForm_Initialize()
     
    Call cbbTipoPopular
    Call cbbClientePopular
    Call cbbCategoriaPopular
    Call cbbUFPopular
    
    Call cbbFltStatusPopular
    
    Call EventosCampos
    
    Call btnFiltrar_Click
    
    Call btnCancelar_Click

End Sub
Private Sub UserForm_Terminate()
    
    ' Destrói objeto da classe cProduto
    Set oObra = Nothing
    Set oTipoObra = Nothing
    Set oCliente = Nothing
    Set oCategoria = Nothing
    Set oTituloReceber = Nothing
    Set oUF = Nothing
    
    Set myRst = Nothing
    
    Call Desconecta
End Sub
Private Sub cbbClientePopular()
    
    Dim idx         As Integer
    Dim col         As New Collection
    Dim n           As Variant

    Set col = oCliente.Listar("nome")
    
    idx = cbbCliente.ListIndex
    
    cbbCliente.Clear
    
    For Each n In col
        
        oCliente.Carrega CLng(n)
    
        With cbbCliente
            .AddItem
            .List(.ListCount - 1, 0) = oCliente.Nome
            .List(.ListCount - 1, 1) = oCliente.ID
        End With
        
    Next n
    
    cbbCliente.ListIndex = idx

End Sub
Private Sub lblHdNome_Click():
    Call lstPrincipalPopular(sCampoOrderBy)
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

Private Sub btnIncluir_Click()
    Call PosDecisaoTomada("Inclusão")
End Sub
Private Sub btnAlterar_Click()
    Call PosDecisaoTomada("Alteração")
End Sub
Private Sub btnExcluir_Click()
    Call PosDecisaoTomada("Exclusão")
End Sub
Private Sub PosDecisaoTomada(Decisao As String)

    btnConfirmar.Visible = True: btnCancelar.Visible = True
    btnConfirmar.Caption = "Confirmar " & VBA.vbNewLine & Decisao
    btnCancelar.Caption = "Cancelar " & VBA.vbNewLine & Decisao
    
    btnIncluir.Visible = False: btnAlterar.Visible = False: btnExcluir.Visible = False
    
    MultiPage1.Value = 1
    
    If Decisao = "Inclusão" Then
        MultiPage1.Value = 1
        MultiPage1.Pages(0).Enabled = False
        Call Campos("Limpar")
        txbData.Text = Date
    End If
    
    If Decisao = "Alteração" Then
        MultiPage1.Pages(2).Enabled = False
    End If
    
    If Decisao <> "Exclusão" Then
        Call Campos("Habilitar")
        cbbCliente.SetFocus
    End If
    
End Sub

Private Sub lstPrincipal_Change()

    Dim n As Long
    Dim iTipoID As Integer

    btnAlterar.Enabled = True
    btnExcluir.Enabled = True
    
    If lstPrincipal.ListIndex >= 0 Then
        oObra.Carrega (CLng(lstPrincipal.List(lstPrincipal.ListIndex, 0)))
        oCliente.Carrega oObra.ClienteID
    End If
    
    ' Preenche informações do cabeçalho
    lblCabID.Caption = Format(IIf(oObra.ID = 0, "", oObra.ID), "00000")
    lblCabCliente.Caption = oCliente.Nome
    lblCabBairro.Caption = oObra.Bairro
    lblCabEndereco.Caption = oObra.Endereco
    
    
    txbBairro.Text = oObra.Bairro
    txbCidade.Text = oObra.Cidade
    cbbUF.Text = oObra.UF
    txbEndereco.Text = oObra.Endereco
    txbData.Text = oObra.Data
    
    For n = 0 To cbbTipo.ListCount - 1
        If CInt(cbbTipo.List(n, 1)) = oObra.TipoID Then
            cbbTipo.ListIndex = n
            Exit For
        End If
    Next n
    
    For n = 0 To cbbCategoria.ListCount - 1
        If CLng(cbbCategoria.List(n, 1)) = oObra.CategoriaID Then
            cbbCategoria.ListIndex = n
            Exit For
        End If
    Next n
    
    If oObra.ClienteID = Null Then
        cbbCliente.ListIndex = -1
    Else
        For n = 0 To cbbCliente.ListCount - 1
            If CInt(cbbCliente.List(n, 1)) = oObra.ClienteID Then
                cbbCliente.ListIndex = n
                Exit For
            End If
        Next n
    End If
    
    chbEncerrada.Value = oObra.Encerrada
    
    txbQtde.Text = IIf(IsNull(oObra.QtdeCasas), Format(0, "#,##0"), oObra.QtdeCasas)
    txbObraM2.Text = IIf(IsNull(oObra.ObraM2), Format(0, "#,##0.00"), oObra.ObraM2)
    txbTerrenoM2.Text = IIf(IsNull(oObra.TerrenoM2), Format(0, "#,##0.00"), oObra.TerrenoM2)
    txbPrecoM2.Text = IIf(IsNull(oObra.PrecoM2), Format(0, "#,##0.00"), oObra.PrecoM2)
    txbMuroPortao.Text = IIf(IsNull(oObra.PrecoMuroPortao), Format(0, "#,##0.00"), oObra.PrecoMuroPortao)
    txbPrazoDias.Text = IIf(IsNull(oObra.PrazoDias), Format(0, "#,##0"), oObra.PrazoDias)
    txbMetrosFrente.Text = IIf(IsNull(oObra.MetrosFrente), Format(0, "#,##0.00"), oObra.MetrosFrente)
    
    Call lstTitulosPopular(CLng(lblCabID.Caption))

End Sub
Private Sub btnCancelar_Click()
    
    btnIncluir.Visible = True: btnAlterar.Visible = True: btnExcluir.Visible = True
    btnConfirmar.Visible = False: btnCancelar.Visible = False
    
    ' Tira a seleção
    lstPrincipal.ListIndex = -1
    
    Call Campos("Limpar")
    Call Campos("Desabilitar")
    
    btnAlterar.Enabled = False
    btnExcluir.Enabled = False
    btnIncluir.SetFocus
   
    MultiPage1.Value = 0
    
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
                .ColumnWidths = "60pt; 65pt; 135pt; 0pt;"
                .Font = "Consolas"
                .AddItem
                
                .List(.ListCount - 1, 0) = txbVencimento.Text
                .List(.ListCount - 1, 1) = Space(12 - Len(Format(CDbl(txbValor.Text), "#,##0.00"))) & Format(CDbl(txbValor.Text), "#,##0.00")
                .List(.ListCount - 1, 2) = txbObservacao.Text
                
            End With
            
            Call btnTituloCancelar_Click

        End If
    ElseIf sDecisaoTitulo = "Alterar" Then
        If ValidaTitulo = True Then
            With lstTitulos
                .List(.ListIndex, 0) = txbVencimento.Text
                .List(.ListIndex, 1) = Space(12 - Len(Format(CDbl(txbValor.Text), "#,##0.00"))) & Format(CDbl(txbValor.Text), "#,##0.00")
                .List(.ListIndex, 2) = txbObservacao.Text
                If Not IsNull(.List(.ListIndex, 3)) Then
                    .List(.ListIndex, 3) = .List(.ListIndex, 3)
                End If
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
        MultiPage1.Value = 2: txbVencimento.SetFocus
    ElseIf txbValor.Text = Empty Then
        MsgBox "Campo 'Valor' é obrigatório", vbCritical
        MultiPage1.Value = 2: txbValor.SetFocus
    ElseIf txbObservacao.Text = Empty Then
        MsgBox "Campo 'Observação' é obrigatório", vbCritical
        MultiPage1.Value = 2: txbObservacao.SetFocus
    Else
        ValidaTitulo = True
    End If
    
End Function
Private Sub Campos(Acao As String)

    If Acao = "Desabilitar" Then
        txbEndereco.Enabled = False: lblEndereco.Enabled = False
        txbBairro.Enabled = False: lblBairro.Enabled = False
        txbCidade.Enabled = False: lblCidade.Enabled = False
        cbbUF.Enabled = False: lblUF.Enabled = False
        cbbTipo.Enabled = False: lblTipo.Enabled = False
        cbbCliente.Enabled = False: lblCliente.Enabled = False
        txbData.Enabled = False: lblData.Enabled = False: btnData.Enabled = False
        cbbCategoria.Enabled = False: lblCategoria.Enabled = False
        chbEncerrada.Enabled = False
        txbQtde.Enabled = False: lblQtde.Enabled = False
        txbObraM2.Enabled = False: lblObraM2.Enabled = False
        txbTerrenoM2.Enabled = False: lblTerrenoM2.Enabled = False
        txbPrecoM2.Enabled = False: lblPrecoM2.Enabled = False
        txbMuroPortao.Enabled = False: lblMuroPortao.Enabled = False
        txbPrazoDias.Enabled = False: lblPrazoDias.Enabled = False
        txbMetrosFrente.Enabled = False: lblMetrosFrente.Enabled = False
        
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
        MultiPage1.Pages(2).Enabled = True
        
    ElseIf Acao = "Habilitar" Then
        txbEndereco.Enabled = True: lblEndereco.Enabled = True
        txbBairro.Enabled = True: lblBairro.Enabled = True
        txbCidade.Enabled = True: lblCidade.Enabled = True
        cbbUF.Enabled = True: lblUF.Enabled = True
        cbbTipo.Enabled = True: lblTipo.Enabled = True
        cbbCliente.Enabled = True: lblCliente.Enabled = True
        txbData.Enabled = True: lblData.Enabled = True: btnData.Enabled = True
        cbbCategoria.Enabled = True: lblCategoria.Enabled = True
        chbEncerrada.Enabled = True
        txbQtde.Enabled = True: lblQtde.Enabled = True
        txbObraM2.Enabled = True: lblObraM2.Enabled = True
        txbTerrenoM2.Enabled = True: lblTerrenoM2.Enabled = True
        txbPrecoM2.Enabled = True: lblPrecoM2.Enabled = True
        txbMuroPortao.Enabled = True: lblMuroPortao.Enabled = True
        txbPrazoDias.Enabled = True: lblPrazoDias.Enabled = True
        txbMetrosFrente.Enabled = True: lblMetrosFrente.Enabled = True
        
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
        lblCabEndereco.Caption = ""
        txbEndereco.Text = ""
        txbBairro.Text = ""
        txbCidade.Text = ""
        cbbUF.ListIndex = -1
        cbbTipo.ListIndex = -1
        cbbCliente.ListIndex = -1
        cbbCategoria.ListIndex = -1
        txbData.Text = Empty
        chbEncerrada.Value = False
        txbQtde.Text = Format(0, "#,##0")
        txbObraM2.Text = Format(0, "#,##0.00")
        txbTerrenoM2.Text = Format(0, "#,##0.00")
        txbPrecoM2.Text = Format(0, "#,##0.00")
        txbMuroPortao.Text = Format(0, "#,##0.00")
        txbPrazoDias.Text = Format(0, "#,##0")
        txbMetrosFrente.Text = Format(0, "#,##0.00")
        
        lblTotalTitulos.Caption = Format(0, "#,##0.00")
        
        lstTitulos.Clear
        
    End If

End Sub
Private Sub lstPrincipalPopular(Pagina As Long)

    Dim lPosicao    As Long
    Dim lCount      As Long
    
    With lstPrincipal
        .Clear
        .ColumnCount = 6
        .ColumnWidths = "55pt; 55pt; 180pt; 0pt; 120pt; 180pt;"
            ' COLUNAS:
            '   - Código (ID)
            '   - Data
            '   - Cliente
            '   - Bairro
            '   - Endereço
            '   - Encerrada?
        .Enabled = True
        .Font = "Consolas"
        
        lCount = 1
        
        While Not myRst.EOF = True And lCount <= myRst.PageSize

            .AddItem

            oCliente.Carrega myRst.Fields("cliente_id").Value

            .List(.ListCount - 1, 0) = Format(myRst.Fields("id").Value, "0000000000")
            .List(.ListCount - 1, 1) = myRst.Fields("data").Value
            .List(.ListCount - 1, 2) = oCliente.Nome
            .List(.ListCount - 1, 3) = myRst.Fields("encerrada").Value
            .List(.ListCount - 1, 4) = myRst.Fields("bairro").Value
            .List(.ListCount - 1, 5) = myRst.Fields("endereco").Value
            
'            .List(.ListCount - 1, 4) = oEmpresa.Empresa & IIf(oEmpresa.Filial = "", "", " : " & oEmpresa.Filial)
'            .List(.ListCount - 1, 5) = myRst.Fields("status").Value
'            .List(.ListCount - 1, 6) = Space(2 - Len(Format(myRst.Fields("count_exames").Value, "00"))) & Format(myRst.Fields("count_exames").Value, "00")
'            .List(.ListCount - 1, 7) = Space(6 - Len(Format(myRst.Fields("sum_preco").Value, "#,##0.00"))) & Format(myRst.Fields("sum_preco").Value, "#,##0.00")

            lCount = lCount + 1
            
            myRst.MoveNext
            
        Wend

    End With
    
    Call ColoreLegenda
    
    ' Posiciona scroll de navegação em páginas
    lblPaginaAtual.Caption = Pagina
    lblNumeroPaginas.Caption = myRst.PageCount
    bChangeScrPag = False: scrPagina.Value = CLng(lblPaginaAtual.Caption): bChangeScrPag = True
    
    ' Trata os botões de navegação
    Call TrataBotoesNavegacao

End Sub

Private Function Valida(Decisao As String) As Boolean
    
    Valida = False
    
    If Decisao = "Inclusão" Or Decisao = "Alteração" Then
    
        If txbEndereco.Text = Empty Then
            MsgBox "'Endereço' é um campo obrigatório", vbInformation: MultiPage1.Value = 1: txbEndereco.SetFocus
        ElseIf cbbTipo.ListIndex = -1 Then
            MsgBox "'Tipo' é um campo obrigatório", vbInformation: cbbTipo.SetFocus
        ElseIf txbData.Text = Empty Then
            MsgBox "'Data' é um campo obrigatório", vbInformation: MultiPage1.Value = 1: txbData.SetFocus
        ElseIf cbbCategoria.ListIndex = -1 Then
            MsgBox "'Categoria' é um campo obrigatório", vbInformation: MultiPage1.Value = 1: cbbCategoria.SetFocus
        ElseIf lstTitulos.ListCount = 0 Then
            MsgBox "Não há títulos à receber apontados na obra", vbCritical
            MultiPage1.Value = 2: btnTituloInclui.SetFocus
        Else
                    
            With oObra
                .Endereco = txbEndereco.Text
                .Bairro = txbBairro.Text
                .Cidade = txbCidade.Text
                .UF = cbbUF.Text
                .TipoID = IIf(cbbTipo.ListIndex = -1, 0, CInt(cbbTipo.List(cbbTipo.ListIndex, 1)))
                .ClienteID = CLng(cbbCliente.List(cbbCliente.ListIndex, 1))
                .Data = CDate(txbData.Text)
                .CategoriaID = CLng(cbbCategoria.List(cbbCategoria.ListIndex, 1))
                .Encerrada = chbEncerrada.Value
                .QtdeCasas = CInt(txbQtde.Text)
                .ObraM2 = CDbl(txbObraM2.Text)
                .TerrenoM2 = CDbl(txbTerrenoM2.Text)
                .PrecoM2 = CCur(txbPrecoM2.Text)
                .PrecoMuroPortao = CCur(txbMuroPortao.Text)
                .PrazoDias = CInt(txbPrazoDias.Text)
                .MetrosFrente = CDbl(txbMetrosFrente.Text)
            End With
            
            Valida = True
                
        End If
        
    ElseIf Decisao = "Exclusão" Then
    
        If oObra.ExisteRecebimento(oObra.ID) = True Then
            Exit Function
        Else
            Valida = True
        End If
        
    End If
    
End Function
Private Sub cbbTipoPopular()
    
    Dim idx         As Integer
    Dim col         As New Collection
    Dim n           As Variant

    Set col = oTipoObra.Listar("nome")
    
    idx = cbbTipo.ListIndex
    
    cbbTipo.Clear
    
    For Each n In col
        
        oTipoObra.Carrega CLng(n)
    
        With cbbTipo
            .AddItem
            .List(.ListCount - 1, 0) = oTipoObra.Nome
            .List(.ListCount - 1, 1) = oTipoObra.ID
        End With
        
    Next n
    
    cbbTipo.ListIndex = idx

End Sub
Private Sub cbbTipo_AfterUpdate()
    
    Dim vbResposta As VbMsgBoxResult
    Dim idx As Integer
    Dim n As Integer
    
    If cbbTipo.ListIndex = -1 And cbbTipo.Text <> "" Then
        
        vbResposta = MsgBox("Este Tipo de obra não existe, deseja cadastrá-lo?", vbQuestion + vbYesNo)
        
        If vbResposta = vbYes Then
            oTipoObra.Nome = RTrim(cbbTipo.Text)
            oTipoObra.Inclui
            idx = oTipoObra.ID
            Call cbbTipoPopular
            For n = 0 To cbbTipo.ListCount - 1
                If CInt(cbbTipo.List(n, 1)) = idx Then
                    cbbTipo.ListIndex = n
                    Exit For
                End If
            Next n
        Else
            cbbTipo.ListIndex = -1
        End If
        
    End If

End Sub
Private Sub cbbCategoriaPopular()
    
    Dim idx         As Integer
    Dim col         As New Collection
    Dim n           As Variant

    Set col = oCategoria.Listar("categoria, subcategoria, item_subcategoria", "R")
    
    With cbbCategoria
        .Clear
        .ColumnCount = 2
        .ColumnWidths = "180pt; 0pt;"
    End With
    
    
    For Each n In col
        
        oCategoria.Carrega CLng(n)
    
        With cbbCategoria
            .AddItem
            .List(.ListCount - 1, 0) = oCategoria.Categoria & ": " & oCategoria.Subcategoria & IIf(oCategoria.ItemSubcategoria = "", "", ": " & oCategoria.ItemSubcategoria)
            .List(.ListCount - 1, 1) = oCategoria.ID
        End With
        
    Next n
    
    cbbCategoria.ListIndex = -1

End Sub
Private Sub lstTitulosPopular(ObraID As Long)

    Dim r       As New ADODB.Recordset
    Dim cTotal As Currency

    sSQL = "SELECT * "
    sSQL = sSQL & "FROM tbl_titulos_receber "
    sSQL = sSQL & "WHERE obra_id = " & ObraID
    
    r.Open sSQL, cnn, adOpenStatic
    
    With lstTitulos
        .Clear
        .ColumnCount = 4
        .ColumnWidths = "60pt; 65pt; 135pt; 0pt;"
        .Font = "Consolas"
        
        Do Until r.EOF
            .AddItem
            
            .List(.ListCount - 1, 0) = r.Fields("vencimento").Value
            .List(.ListCount - 1, 1) = Space(12 - Len(Format(r.Fields("valor").Value, "#,##0.00"))) & Format(r.Fields("valor").Value, "#,##0.00")
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
Private Sub lstTitulos_Change()

    Dim n As Integer

    If lstTitulos.ListIndex > -1 And btnTituloConfirmar.Caption <> "Alterar" Then
        txbVencimento.Text = lstTitulos.List(lstTitulos.ListIndex, 0)
        txbValor.Text = lstTitulos.List(lstTitulos.ListIndex, 1)
        txbObservacao.Text = lstTitulos.List(lstTitulos.ListIndex, 2)
        
        btnTituloAltera.Enabled = True
        btnTituloExclui.Enabled = True
    End If
End Sub
Private Sub btnConfirmar_Click()
    
    Dim vbResposta As VBA.VbMsgBoxResult
    Dim sDecisao As String
    Dim i As Integer
    
    sDecisao = Replace(btnConfirmar.Caption, "Confirmar " & vbNewLine, "")
    
    If Valida(sDecisao) = True Then
        vbResposta = MsgBox("Deseja realmente fazer a " & sDecisao & "?", vbYesNo + vbQuestion, "Pergunta")
        
        If vbResposta = vbYes Then
        
            ' Cabeçalho da compra
            If sDecisao = "Inclusão" Then
                oObra.Inclui
                
                ' Títulos das compras (DOING)
                For i = 0 To lstTitulos.ListCount - 1
                
                    With oTituloReceber
                        .ObraID = oObra.ID
                        .ClienteID = oObra.ClienteID
                        .Observacao = lstTitulos.List(i, 2)
                        .Vencimento = CDate(lstTitulos.List(i, 0))
                        .Valor = CCur(lstTitulos.List(i, 1))
                        .Data = CDate(txbData.Text)
                        .Inclui
                    End With
                    
                Next i
                
            ElseIf sDecisao = "Alteração" Then
                oObra.Altera
            ElseIf sDecisao = "Exclusão" Then
                oTituloReceber.Exclui oObra.ID
                oObra.Exclui oObra.ID
            End If
            
            ' Clica no botão filtrar para chamar a rotina de popular lstPrincipal
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
Private Sub btnData_Click()
    dtDate = IIf(txbData.Text = Empty, Date, txbData.Text)
    txbData.Text = GetCalendario
End Sub
Private Sub cbbUFPopular()

    Dim idx         As Integer
    Dim col         As New Collection
    Dim n           As Variant
    Dim a()         As String

    Set col = oUF.Listar
    
    idx = cbbUF.ListIndex
    
    cbbUF.Clear
    cbbUF.ColumnCount = 2
    cbbUF.ColumnWidths = "20pt; 120pt;"
    
    For Each n In col
    
        a = Split(n, ";")
    
        With cbbUF
            .AddItem
            .List(.ListCount - 1, 0) = a(0)
            .List(.ListCount - 1, 1) = a(1)
        End With
        
    Next n
    
    cbbUF.ListIndex = idx
    

End Sub
Private Sub cbbCliente_AfterUpdate()

    Dim vbResposta As VbMsgBoxResult
    Dim idx As Integer
    Dim n As Integer
    
    If cbbCliente.ListIndex = -1 And cbbCliente.Text <> "" Then
        
        vbResposta = MsgBox("Este Cliente não existe, deseja cadastrá-lo?", vbQuestion + vbYesNo)
        
        If vbResposta = vbYes Then
            
            oCliente.Nome = RTrim(cbbCliente.Text)
            oCliente.Inclui
            
            idx = oCliente.ID
            
            Call cbbClientePopular
            
            For n = 0 To cbbCliente.ListCount - 1
                If CInt(cbbCliente.List(n, 1)) = idx Then
                    cbbCliente.ListIndex = n
                    Exit For
                End If
            Next n
        Else
            cbbCliente.ListIndex = -1
        End If
        
    End If

End Sub
Private Sub cbbFltStatusPopular()
    
    With cbbFltStatus
        .Clear
        .ColumnCount = 2
        .ColumnWidths = "65pt; 0pt;"
        .AddItem
        .List(.ListCount - 1, 0) = "***TODOS***"
        .List(.ListCount - 1, 1) = "All"
        .AddItem
        .List(.ListCount - 1, 0) = "Em andamento"
        .List(.ListCount - 1, 1) = False
        .AddItem
        .List(.ListCount - 1, 0) = "Encerrada"
        .List(.ListCount - 1, 1) = True
    End With
    
    cbbFltStatus.ListIndex = 0

End Sub
Private Sub btnFiltrar_Click()

    Dim vStatus As Variant
    
    If cbbFltStatus.ListIndex <= 0 Then
        vStatus = Null
    Else
        vStatus = CInt(cbbFltStatus.List(cbbFltStatus.ListIndex, 1))
    End If

    Set myRst = oObra.Recordset(vStatus)
    
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
Private Sub ColoreLegenda()

    Dim idx         As Integer
    Dim c           As control
    
    For Each c In fVendas.Controls
        
        If TypeName(c) = "Label" And c.Tag = "status" Then
            
            idx = CInt(Mid(c.name, 2, 2))
            
            If idx <= (lstPrincipal.ListCount - 1) Then
                If CBool(lstPrincipal.List(idx, 3)) = False Then
                    c.BackColor = &HC000& ' Verde
                ElseIf CBool(lstPrincipal.List(idx, 3)) = True Then
                    c.BackColor = &HC0& ' Vermelho
                End If
            Else
                c.BackColor = &H8000000F
            End If
        
        End If
        
    Next c
    
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
