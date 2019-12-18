VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fCompras 
   Caption         =   ":: Cadastro de Compras ::"
   ClientHeight    =   9705
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

Private colControles        As New Collection
Private bListBoxOrdenando   As Boolean
Private myRst               As ADODB.RecordSet
Private lPagina             As Long

Private Const sTable As String = "tbl_compras"
Private Const sCampoOrderBy As String = "data"

Private Sub UserForm_Initialize()

    Call cbbFornecedorPopular
    Call cbbProdutoPopular
    Call EventosCampos
    
    Set myRst = New ADODB.RecordSet
    Set myRst = oCompra.RecordSet
    
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
    
    ' Destr�i objeto da classe cProduto
    Set oCompra = Nothing
    Call Desconecta
    
End Sub
Private Sub cbbFornecedor_AfterUpdate()

    Dim vbResposta As VbMsgBoxResult
    Dim idx As Integer
    Dim n As Integer
    
    If cbbFornecedor.ListIndex = -1 And cbbFornecedor.Text <> "" Then
        
        vbResposta = MsgBox("Este fornecedor n�o existe, deseja cadastr�-lo?", vbQuestion + vbYesNo)
        
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
    If lstItens.ListIndex > -1 And btnItemConfirmar.Caption <> "Alterar" Then
        cbbProduto.Text = lstItens.List(lstItens.ListIndex, 0)
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
    
    If Valida = True Then
        vbResposta = MsgBox("Deseja realmente fazer a " & sDecisao & "?", vbYesNo + vbQuestion, "Pergunta")
        
        If vbResposta = vbYes Then
        
            ' Cabe�alho da compra
            If sDecisao = "Inclus�o" Then
                oCompra.Inclui
            ElseIf sDecisao = "Altera��o" Then
                oCompra.Altera CLng(lblCabID.Caption)
            End If
            
            ' Itens das compras
            For i = 0 To lstItens.ListCount - 1
            
                With oCompraItem
                    .ProdutoID = CLng(lstItens.List(i, 1))
                    .Quantidade = CDbl(lstItens.List(i, 2))
                    .Unitario = CDbl(lstItens.List(i, 3))
                    .Data = oCompra.Data
                    .FornecedorID = oCompra.FornecedorID
                    .CompraID = oCompra.ID
                    
                    If Not IsNull(lstItens.List(i, 5)) Then
                        .Recno = CLng(lstItens.List(i, 5))
                    Else
                        If sDecisao = "Inclus�o" Then
                            .Inclui
                        ElseIf sDecisao = "Altera��o" Then
                            .Altera .Recno
                        ElseIf sDecisao = "Exclus�o" Then
                            .Exclui .Recno
                        End If
                        
                    End If
                    
                End With
                
            Next i
            
            ' T�tulos das compras (DOING)
            For i = 0 To lstTitulos.ListCount - 1
            
                With oTituloPagar
                    .CompraID = oCompra.ID
                    .FornecedorID = oCompra.FornecedorID
                    .Observacao = lstTitulos.List(i, 2)
                    .Vencimento = CDate(lstTitulos.List(i, 0))
                    .Valor = CCur(lstTitulos.List(i, 1))
                    .Data = oCompra.Data
                    
                    If Not IsNull(lstItens.List(i, 5)) Then
                        .Recno = CLng(lstItens.List(i, 5))
                    Else
                        If sDecisao = "Inclus�o" Then
                            .Inclui
                        ElseIf sDecisao = "Altera��o" Then
                            '.Altera .Recno
                        ElseIf sDecisao = "Exclus�o" Then
                            .Exclui .Recno
                        End If
                        
                    End If
                    
                End With
                
            Next i
            
            If sDecisao = "Exclus�o" Then
                oCompra.Exclui oCompra.ID
            End If
            
            
            'TODO (Inclui titulos)
            'For i = 0 To lstTitulos.ListCount - 1
            'Next i
                
            
            If sDecisao = "Inclus�o" Then
                If lstPrincipal.ListCount < myRst.PageSize Then
                    lPagina = Trim(Mid(lblPaginaAtual.Caption, InStr(1, lblPaginaAtual.Caption, "de") + 3, Len(lblPaginaAtual.Caption)))
                Else
                    lPagina = Trim(Mid(lblPaginaAtual.Caption, InStr(1, lblPaginaAtual.Caption, "de") + 3, Len(lblPaginaAtual.Caption))) + 1
                End If
            Else
                lPagina = Trim(Mid(lblPaginaAtual.Caption, InStr(1, lblPaginaAtual.Caption, "de") + 3, Len(lblPaginaAtual.Caption)))
            End If
            
            Set myRst = New ADODB.RecordSet
            Set myRst = oCompra.RecordSet
        
            With scrPagina
                .Min = 1
                .Max = myRst.PageCount
            End With
            
            myRst.AbsolutePage = myRst.PageCount
            scrPagina.Value = lPagina
            
            Call lstPrincipalPopular(lPagina)
            
            ' Exibe mensagem de sucesso na decis�o tomada (inclus�o, altera��o ou exclus�o do registro).
            MsgBox sDecisao & " realizada com sucesso.", vbInformation, sDecisao & " de registro"
            
            MultiPage1.Value = 0
            
            Call btnCancelar_Click
            
        ElseIf vbResposta = vbNo Then
            Call btnCancelar_Click
        End If
    
    End If
    
End Sub
Private Sub btnIncluir_Click()
    Call PosDecisaoTomada("Inclus�o")
    lstPrincipal.ListIndex = -1
End Sub
Private Sub btnAlterar_Click()
    Call PosDecisaoTomada("Altera��o")
End Sub
Private Sub btnExcluir_Click()
    Call PosDecisaoTomada("Exclus�o")
End Sub
Private Sub PosDecisaoTomada(Decisao As String)

    btnCancelar.Visible = True: btnConfirmar.Visible = True
    btnConfirmar.Caption = "Confirmar " & Decisao
    btnCancelar.Caption = "Cancelar " & Decisao
    
    btnIncluir.Visible = False: btnAlterar.Visible = False: btnExcluir.Visible = False
    
    If Decisao = "Inclus�o" Then
        Call Campos("Limpar")
    End If
    
    If Decisao <> "Exclus�o" Then
        Call Campos("Habilitar")
        
        If MultiPage1.Value = 0 Then
            MultiPage1.Value = 1
        End If
        
        If Decisao = "Inclus�o" Then
            txbData.Text = Date
            If MultiPage1.Value = 1 Then
                cbbFornecedor.SetFocus
            End If
        Else
            If MultiPage1.Value = 1 Then
                cbbFornecedor.SetFocus
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
Private Sub lstPrincipalPopular(Pagina As Long)

    Dim lPosicao    As Long
    Dim lCount      As Long
    
    With lstPrincipal
        .Clear
        .ColumnCount = 8 ' Funcion�rio, ID, Empresa, Filial
        .ColumnWidths = "55pt; 55pt; 160pt;"
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
   
    lblPaginaAtual.Caption = "P�gina " & Format(scrPagina.Value, "#,##0") & " de " & Format(myRst.PageCount, "#,##0")

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
'Private Sub lblHdNome_Click():
'    Call lstPrincipalPopular(sCampoOrderBy)
'End Sub
'
Private Sub lstPrincipal_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    MultiPage1.Value = 1
End Sub
Private Sub EventosCampos()

    ' Declara vari�veis
    Dim oControle   As MSForms.control
    Dim oEvento     As c_EventoCampo
    Dim sTag        As String
    Dim iType       As Integer
    Dim bNullable   As Boolean
    Dim sField()    As String

    ' La�o para percorrer todos os TextBox e atribuir eventos
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
    
        btnAlterar.Enabled = True: btnExcluir.Enabled = True
        
        ' Carrega informa��es do lan�amento
        oCompra.Carrega (CLng(lstPrincipal.List(lstPrincipal.ListIndex, 0)))
        
        ' Preenche cabe�alho
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
        
        Call lstItensPopular(CLng(lblCabID.Caption))
        Call lstTitulosPopular(CLng(lblCabID.Caption))
    End If

End Sub
Private Sub lstItensPopular(CompraID As Long)

    Dim r       As New ADODB.RecordSet
    Dim cTotal As Currency

    sSQL = "SELECT * "
    sSQL = sSQL & "FROM tbl_compras_itens "
    sSQL = sSQL & "WHERE compra_id = " & CompraID
    
    r.Open sSQL, cnn, adOpenStatic
    
    With lstItens
        .Clear
        .ColumnCount = 6
        .ColumnWidths = "200pt; 0pt; 60pt; 60pt; 60pt; 0pt;"
        .Font = "Consolas"
        
        Do Until r.EOF
            .AddItem
            
            oProduto.Carrega r.Fields("produto_id").Value
            
            .List(.ListCount - 1, 0) = oProduto.Nome
            .List(.ListCount - 1, 1) = r.Fields("produto_id").Value
            .List(.ListCount - 1, 2) = Space(9 - Len(Format(r.Fields("quantidade").Value, "#,##0.00"))) & Format(r.Fields("quantidade").Value, "#,##0.00")
            .List(.ListCount - 1, 3) = Space(9 - Len(Format(r.Fields("unitario").Value, "#,##0.00"))) & Format(r.Fields("unitario").Value, "#,##0.00")
            
            cTotal = r.Fields("quantidade").Value * r.Fields("unitario").Value
            
            .List(.ListCount - 1, 4) = Space(9 - Len(Format(cTotal, "#,##0.00"))) & Format(cTotal, "#,##0.00")
            .List(.ListCount - 1, 5) = r.Fields("r_e_c_n_o_").Value
            
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

    btnIncluir.Visible = True: btnAlterar.Visible = True: btnExcluir.Visible = True
    btnConfirmar.Visible = False: btnCancelar.Visible = False

    Call Campos("Limpar")
    Call Campos("Desabilitar")

    btnAlterar.Enabled = False
    btnExcluir.Enabled = False
    btnIncluir.SetFocus
    
    MultiPage1.Value = 0

    ' Tira a sele��o
    lstPrincipal.ListIndex = -1: lstPrincipal.ForeColor = &H80000008: lstPrincipal.Enabled = True:

End Sub
Private Sub Campos(Acao As String)

    If Acao = "Desabilitar" Then
        txbData.Enabled = False: lblData.Enabled = False: btnData.Enabled = False
        cbbFornecedor.Enabled = False: lblFornecedor.Enabled = False
        
        frmItem.Enabled = False
        lblHdProduto.Enabled = False
        lblHdQuant.Enabled = False
        lblHdUnitario.Enabled = False
        lblHdTotal.Enabled = False
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
        
    ElseIf Acao = "Habilitar" Then
        txbData.Enabled = True: lblData.Enabled = True: btnData.Enabled = True
        cbbFornecedor.Enabled = True: lblFornecedor.Enabled = True
        frmItem.Enabled = True
        
        lstItens.Enabled = True: lstItens.ForeColor = &H80000008
        lblHdProduto.Enabled = True
        lblHdProduto.Enabled = True
        lblHdQuant.Enabled = True
        lblHdUnitario.Enabled = True
        lblHdTotal.Enabled = True
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
        
    ElseIf Acao = "Limpar" Then
        lblCabID.Caption = ""
        lblCabData.Caption = ""
        txbData.Text = ""
        cbbFornecedor.ListIndex = -1
        lstItens.Clear
        lstTitulos.Clear
        
        lstPrincipal.ListIndex = -1
    End If

End Sub
'Private Sub ListBoxOrdenar()
'
'    Dim ini, fim, i, j  As Long
'    Dim sCol01          As String
'    Dim sCol02          As String
'
'    bListBoxOrdenando = True
'
'    With lstPrincipal
'
'        ini = 0
'        fim = .ListCount - 1 '4 itens(0 - 3)
'
'        For i = ini To fim - 1  ' La�o para comparar cada item com todos os outros itens
'            For j = i + 1 To fim    ' La�o para comparar item com o pr�ximo item
'                If .List(i) > .List(j) Then
'                    sCol01 = .List(j, 0)
'                    sCol02 = .List(j, 1)
'                    .List(j, 0) = .List(i, 0)
'                    .List(j, 1) = .List(i, 1)
'                    .List(i, 0) = sCol01
'                    .List(i, 1) = sCol02
'                End If
'            Next j
'        Next i
'    End With
'
'    bListBoxOrdenando = False
'
'End Sub

Private Function Valida() As Boolean
    
    Valida = False
    
    If txbData.Text = Empty Then
        MsgBox "Campo 'Data' � obrigat�rio", vbCritical
        MultiPage1.Value = 1: txbData.SetFocus: Exit Function
    ElseIf cbbFornecedor.ListIndex = -1 Then
        MsgBox "Campo 'Fornecedor' � obrigat�rio", vbCritical
        MultiPage1.Value = 1: cbbFornecedor.SetFocus: Exit Function
    Else
        If lstItens.ListCount = 0 Then
            MsgBox "N�o h� itens apontados na compra", vbCritical
            MultiPage1.Value = 2: btnItemInclui.SetFocus: Exit Function
        ElseIf lstTitulos.ListCount = 0 Then
            MsgBox "N�o h� t�tulos apontados na compra", vbCritical
            MultiPage1.Value = 3: btnTituloInclui.SetFocus: Exit Function
        Else
            With oCompra
                .Data = CDate(txbData.Text)
                .FornecedorID = CLng(cbbFornecedor.List(cbbFornecedor.ListIndex, 1))
            End With
            
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
        txbUnitario.Text = Format(0, "#,##0.00")
        txbTotal.Text = Format(0, "#,##0.00")
    End If
    
    If Habilitar = True Then
        
        cbbProduto.Enabled = Habilitar: lblProduto.Enabled = Habilitar
        txbQtde.Enabled = Habilitar: lblQtde.Enabled = Habilitar
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
    
    sDecisaoLancamento = Replace(btnConfirmar.Caption, "Confirmar ", "")
    sDecisaoItem = btnItemConfirmar.Caption
    
    If sDecisaoItem = "Incluir" Then
    
        If ValidaItem = True Then
            
            With lstItens
                .ColumnCount = 6
                .ColumnWidths = "200pt; 0pt; 60pt; 60pt; 60pt;"
                .Font = "Consolas"
                .AddItem
                
                .List(.ListCount - 1, 0) = cbbProduto.List(cbbProduto.ListIndex, 0)
                .List(.ListCount - 1, 1) = cbbProduto.List(cbbProduto.ListIndex, 1)
                .List(.ListCount - 1, 2) = Space(9 - Len(Format(CDbl(txbQtde.Text), "#,##0.00"))) & Format(CDbl(txbQtde.Text), "#,##0.00")
                .List(.ListCount - 1, 3) = Space(9 - Len(Format(CDbl(txbUnitario.Text), "#,##0.00"))) & Format(CDbl(txbUnitario.Text), "#,##0.00")
                .List(.ListCount - 1, 4) = Space(9 - Len(Format(CDbl(txbTotal.Text), "#,##0.00"))) & Format(CDbl(txbTotal.Text), "#,##0.00")
                
            End With
            
            Call btnItemCancelar_Click

        End If
    ElseIf sDecisaoItem = "Alterar" Then
        If ValidaItem = True Then
            With lstItens
                .List(.ListIndex, 0) = cbbProduto.List(cbbProduto.ListIndex, 0)
                .List(.ListIndex, 1) = cbbProduto.List(cbbProduto.ListIndex, 1)
                .List(.ListIndex, 2) = Space(9 - Len(Format(CDbl(txbQtde.Text), "#,##0.00"))) & Format(CDbl(txbQtde.Text), "#,##0.00")
                .List(.ListIndex, 3) = Space(9 - Len(Format(CDbl(txbUnitario.Text), "#,##0.00"))) & Format(CDbl(txbUnitario.Text), "#,##0.00")
                .List(.ListIndex, 4) = Space(9 - Len(Format(CDbl(txbTotal.Text), "#,##0.00"))) & Format(CDbl(txbTotal.Text), "#,##0.00")
            End With
            
            Call btnItemCancelar_Click
        End If
    ElseIf sDecisaoItem = "Excluir" Then
        lstItens.RemoveItem (lstItens.ListIndex)
        Call btnItemCancelar_Click
    End If
    
    Call TotalizaItens
    
End Sub
Private Function ValidaItem() As Boolean
    ValidaItem = False
    If cbbProduto.ListIndex = -1 Then
        MsgBox "Campo 'Produto' � obrigat�rio", vbCritical
        MultiPage1.Value = 2: cbbProduto.SetFocus: Exit Function
    ElseIf txbQtde.Text = Empty Then
        MsgBox "Campo 'Quantidade' � obrigat�rio", vbCritical
        MultiPage1.Value = 2: txbQtde.SetFocus: Exit Function
    ElseIf txbUnitario.Text = Empty Then
        MsgBox "Campo 'Unit�rio' � obrigat�rio", vbCritical
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
Private Sub txbQtde_AfterUpdate()
    txbQtde.Text = Format(txbQtde.Text, "#,##0.00")
    txbTotal.Text = Format(CDbl(txbQtde.Text) * CDbl(txbUnitario.Text), "#,##0.00")
End Sub
Private Sub txbUnitario_AfterUpdate()
    txbQtde.Text = Format(txbQtde.Text, "#,##0.00")
    txbTotal.Text = Format(CDbl(txbQtde.Text) * CDbl(txbUnitario.Text), "#,##0.00")
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
                .ColumnCount = 3
                .ColumnWidths = "60pt; 60pt; 180pt;"
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
        MsgBox "Campo 'Vencimento' � obrigat�rio", vbCritical
        MultiPage1.Value = 3: txbVencimento.SetFocus: Exit Function
    ElseIf txbValor.Text = Empty Then
        MsgBox "Campo 'Valor' � obrigat�rio", vbCritical
        MultiPage1.Value = 3: txbValor.SetFocus: Exit Function
    ElseIf txbObservacao.Text = Empty Then
        MsgBox "Campo 'Observa��o' � obrigat�rio", vbCritical
        MultiPage1.Value = 3: txbObservacao.SetFocus: Exit Function
    Else
        ValidaTitulo = True
    End If
    
End Function
Private Sub lstTitulosPopular(CompraID As Long)

    Dim r       As New ADODB.RecordSet
    Dim cTotal As Currency

    sSQL = "SELECT * "
    sSQL = sSQL & "FROM tbl_titulos_pagar "
    sSQL = sSQL & "WHERE compra_id = " & CompraID
    
    r.Open sSQL, cnn, adOpenStatic
    
    With lstTitulos
        .Clear
        .ColumnCount = 4
        .ColumnWidths = "60pt; 60pt; 60pt; 0pt;"
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
    If lstTitulos.ListIndex > -1 And btnTituloConfirmar.Caption <> "Alterar" Then
        txbVencimento.Text = lstTitulos.List(lstTitulos.ListIndex, 0)
        txbValor.Text = lstTitulos.List(lstTitulos.ListIndex, 1)
        txbObservacao.Text = lstTitulos.List(lstTitulos.ListIndex, 2)
        
        btnTituloAltera.Enabled = True
        btnTituloExclui.Enabled = True
    End If
End Sub
Private Sub cbbFornecedor_Change()
    If cbbFornecedor.ListIndex > -1 And cbbFornecedor.Text <> "" Then
        MultiPage1.Value = 2
    End If
End Sub
