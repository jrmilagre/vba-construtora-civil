VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fTitulosPagar 
   Caption         =   ":: Cadastro de Títulos à Pagar ::"
   ClientHeight    =   10560
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13320
   OleObjectBlob   =   "fTitulosPagar.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fTitulosPagar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private oTituloPagar        As New cTituloPagar
Private oFornecedor         As New cFornecedor

Private colControles        As New Collection
Private bListBoxOrdenando   As Boolean
Private myRst               As New ADODB.Recordset

Private Const sTable As String = "tbl_titulos_pagar"
Private Const sCampoOrderBy As String = "vencimento"

Private Sub UserForm_Initialize()

    Call cbbFornecedorPopular
    
    Call cbbFltFornecedorPopular
    
    Call EventosCampos
        
    Call btnFiltrar_Click
    
    Call btnCancelar_Click

End Sub
Private Sub UserForm_Terminate()
    
    ' Destrói objeto da classe cProduto
    Set oTituloPagar = Nothing
    Set oFornecedor = Nothing
    
    Call Desconecta
    
End Sub

Private Sub btnConfirmar_Click()
    
    Dim vbResposta As VBA.VbMsgBoxResult
    Dim sDecisao As String
    Dim i As Integer
    
    sDecisao = Replace(btnConfirmar.Caption, "Confirmar ", "")
    
    If Valida(sDecisao) = True Then
    
        vbResposta = MsgBox("Deseja realmente fazer a " & sDecisao & "?", vbYesNo + vbQuestion, "Pergunta")
        
        If vbResposta = vbYes Then
        
            If sDecisao = "Inclusão" Then
                oTituloPagar.Inclui
            ElseIf sDecisao = "Alteração" Then
                oTituloPagar.AlteraTitulo oTituloPagar.Recno
            ElseIf sDecisao = "Exclusão" Then
                oTituloPagar.ExcluiTitulo oTituloPagar.Recno
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
            txbVencimento.Text = Date + 30
            If MultiPage1.Value = 1 Then
                txbValor.Text = Format(0, "#,##0.00")
                txbValor.SelStart = 0
                txbValor.SelLength = Len(txbValor.Text)
                txbValor.SetFocus
                txbBaixado.Text = Format(0, "#,##0.00")
                txbSaldo.Text = Format(0, "#,##0.00")
            End If
        Else
            If MultiPage1.Value = 1 Then
                txbValor.SetFocus
            End If
        End If
            
    End If
    
    lstPrincipal.Enabled = False
    lstPrincipal.ForeColor = &H80000010
    
    Call Filtros("Desabilitar")

    btnPaginaInicial.Enabled = False
    btnPaginaAnterior.Enabled = False
    btnPaginaSeguinte.Enabled = False
    btnPaginaFinal.Enabled = False
    
End Sub
Private Sub lstPrincipalPopular()

    Dim lCount      As Long
    Dim cVlrBxd     As Currency
    Dim cVlrSld     As Currency
    Dim c           As control
    
    ' Numera a página posicionada
    If myRst.AbsolutePage = adPosEOF Then
        lblPaginaAtual.Caption = "Página " & Format(myRst.PageCount, "#,##0") & " de " & Format(myRst.PageCount, "#,##0")
    Else
        lblPaginaAtual.Caption = "Página " & Format(myRst.AbsolutePage, "#,##0") & " de " & Format(myRst.PageCount, "#,##0")
    End If
    
    With lstPrincipal
        .Clear
        .ColumnCount = 7
        .ColumnWidths = "55pt; 55pt; 55pt; 65pt; 65pt; 65pt; 100pt;"
        .Enabled = True
        .Font = "Consolas"
        
        lCount = 1
        
        While Not myRst.EOF = True And lCount <= myRst.PageSize

            .AddItem

            .List(.ListCount - 1, 0) = Format(myRst.Fields("r_e_c_n_o_").Value, "0000000000")
            .List(.ListCount - 1, 1) = myRst.Fields("data").Value
            .List(.ListCount - 1, 2) = myRst.Fields("vencimento").Value
            .List(.ListCount - 1, 3) = Space(12 - Len(Format(myRst.Fields("valor").Value, "#,##0.00"))) & Format(myRst.Fields("valor").Value, "#,##0.00")
            
            cVlrBxd = oTituloPagar.GetValorBaixado(myRst.Fields("r_e_c_n_o_").Value)
            cVlrSld = myRst.Fields("valor").Value - cVlrBxd
            
            .List(.ListCount - 1, 4) = Space(12 - Len(Format(cVlrBxd, "#,##0.00"))) & Format(cVlrBxd, "#,##0.00")
            .List(.ListCount - 1, 5) = Space(12 - Len(Format(cVlrSld, "#,##0.00"))) & Format(cVlrSld, "#,##0.00")
            .List(.ListCount - 1, 6) = myRst.Fields("observacao").Value

            lCount = lCount + 1
            
            myRst.MoveNext
            
        Wend

    End With
    
    ' Colore status
    Dim idx As Integer
    
    For Each c In fTitulosReceber.Controls
        
        If TypeName(c) = "Label" And c.Tag = "status" Then
            'Stop
            
            idx = CInt(Mid(c.name, 2, 2))
            
            If idx <= (lstPrincipal.ListCount - 1) Then
                If CDate(lstPrincipal.List(idx, 2)) > (Date + 3) Then
                    c.BackColor = &HC000& ' Verde
                ElseIf CDate(lstPrincipal.List(idx, 2)) < Date Then
                    c.BackColor = &HC0& ' Vermelho
                Else
                    c.BackColor = &HFFFF&         ' Amarelo
                End If
            Else
                c.BackColor = &H8000000F
            End If
        
        End If
        
    Next c
   
End Sub
Private Sub btnData_Click()
    dtDate = IIf(txbData.Text = Empty, Date, txbData.Text)
    txbData.Text = GetCalendario
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

        btnAlterar.Enabled = True
        btnExcluir.Enabled = True

        ' Carrega informações do lançamento
        oTituloPagar.Carrega CLng(lstPrincipal.List(lstPrincipal.ListIndex, 0))

        ' Preenche cabeçalho
        lblCabID.Caption = IIf(oTituloPagar.Recno = 0, "", Format(oTituloPagar.Recno, "0000000000"))
        lblCabData.Caption = oTituloPagar.Data
'
'        oFornecedor.Carrega oCompra.FornecedorID
'
'        'lblCabFuncionario.Caption = oFuncionario.Funcionario
'
        ' Preenche campos
        txbData.Text = oTituloPagar.Data
        txbVencimento.Text = oTituloPagar.Vencimento
        txbValor.Text = Format(oTituloPagar.Valor, "#,##0.00")
        txbBaixado.Text = lstPrincipal.List(lstPrincipal.ListIndex, 4)
        txbSaldo.Text = lstPrincipal.List(lstPrincipal.ListIndex, 5)
        txbObservacao.Text = oTituloPagar.Observacao
        
        For n = 0 To cbbFornecedor.ListCount - 1
            If CLng(cbbFornecedor.List(n, 1)) = oTituloPagar.FornecedorID Then
                cbbFornecedor.ListIndex = n
                Exit For
            End If
        Next n

    End If

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

    ' Tira a seleção
    lstPrincipal.ListIndex = -1: lstPrincipal.ForeColor = &H80000008: lstPrincipal.Enabled = True:

End Sub
Private Sub Campos(Acao As String)

    Dim sDecisao As String
    
    sDecisao = Replace(btnConfirmar.Caption, "Confirmar ", "")

    If Acao = "Desabilitar" Then
        
        txbData.Enabled = False: lblData.Enabled = False: btnData.Enabled = False
        txbVencimento.Enabled = False: lblVencimento.Enabled = False: btnVencimento.Enabled = False
        txbValor.Enabled = False: lblValor.Enabled = False
        txbObservacao.Enabled = False: lblObservacao.Enabled = False
        cbbFornecedor.Enabled = False: lblFornecedor.Enabled = False

        
    ElseIf Acao = "Habilitar" Then
        
        If sDecisao = "Inclusão" Then
            txbData.Enabled = True: lblData.Enabled = True: btnData.Enabled = True
            cbbFornecedor.Enabled = True: lblFornecedor.Enabled = True
        End If
        
        txbVencimento.Enabled = True: lblVencimento.Enabled = True: btnVencimento.Enabled = True
        txbValor.Enabled = True: lblValor.Enabled = True
        txbObservacao.Enabled = True: lblObservacao.Enabled = True
        
        
    ElseIf Acao = "Limpar" Then
        lblCabID.Caption = ""
        lblCabData.Caption = ""
        txbData.Text = Empty
        txbVencimento.Text = Empty
        txbValor.Text = Empty
        txbBaixado.Text = Empty
        txbSaldo.Text = Empty
        txbObservacao.Text = Empty
        cbbFornecedor.ListIndex = -1
    End If
    
    Call Filtros("Habilitar")

End Sub
Private Function Valida(Decisao As String) As Boolean
    
    Valida = False
    
    If Decisao = "Inclusão" Or Decisao = "Alteração" Then
    
        
        If txbData.Text = Empty Then
            MsgBox "Campo 'Data' é obrigatório", vbCritical
            MultiPage1.Value = 1: txbData.SetFocus
        ElseIf txbVencimento.Text = Empty Then
            MsgBox "Campo 'Vencimento' é obrigatório", vbCritical
            MultiPage1.Value = 1: txbVencimento.SetFocus
        ElseIf txbValor.Text = Empty Or CCur(txbValor.Text) = 0 Then
            MsgBox "Campo 'Valor' é obrigatório", vbCritical
            MultiPage1.Value = 1: txbValor.SetFocus
        ElseIf txbObservacao.Text = Empty Then
            MsgBox "Preencha o campo 'Observação', pode ser importante no futuro!", vbCritical
            MultiPage1.Value = 1: txbObservacao.SetFocus
        ElseIf cbbFornecedor.ListIndex = -1 Then
            MsgBox "Campo 'Fornecedor' é obrigatório", vbCritical
            MultiPage1.Value = 1: cbbFornecedor.SetFocus
        ElseIf CCur(txbValor.Text) < CCur(txbBaixado.Text) Then
            MsgBox "O campo 'Valor' não pode ser menor que o campo 'Baixado'", vbCritical
            MultiPage1.Value = 1: txbValor.SetFocus
        Else

            With oTituloPagar
                .Data = CDate(txbData.Text)
                .Vencimento = CDate(txbVencimento.Text)
                .Valor = CCur(txbValor.Text)
                .FornecedorID = CLng(cbbFornecedor.List(cbbFornecedor.ListIndex, 1))
                
                oFornecedor.Carrega CLng(cbbFornecedor.List(cbbFornecedor.ListIndex, 1))
                
                .CompraID = Null
                .Observacao = RTrim(txbObservacao.Text)
            End With
            
            Valida = True

        End If
    ElseIf Decisao = "Exclusão" Then
    
        If oTituloPagar.ExistePagamento(oTituloPagar.FornecedorID, oTituloPagar.Recno) = True Then
            Exit Function
        Else
            Valida = True
        End If
    
    End If

End Function

Private Sub btnVencimento_Click()
    dtDate = IIf(txbVencimento.Text = Empty, Date, txbVencimento.Text)
    txbVencimento.Text = GetCalendario
End Sub

Private Sub cbbFornecedorPopular()
    
    Dim col         As New Collection
    Dim n           As Variant

    Set col = oFornecedor.Listar("nome")
    
    With cbbFltFornecedor
        .Clear
        .ColumnCount = 2
        .ColumnWidths = "180pt; 0pt;"
    End With
    
    For Each n In col
        
        oFornecedor.Carrega CLng(n)
    
        With cbbFltFornecedor
            .AddItem
            .List(.ListCount - 1, 0) = oFornecedor.Nome
            .List(.ListCount - 1, 1) = oFornecedor.ID
        End With
        
    Next n
    
    cbbFltFornecedor.ListIndex = -1

End Sub
Private Sub cbbFltFornecedorPopular()
    
    Dim col         As New Collection
    Dim n           As Variant

    Set col = oFornecedor.Listar("nome")
    
    With cbbFltFornecedor
        .Clear
        .ColumnCount = 2
        .ColumnWidths = "180pt; 0pt;"
        .AddItem
        .List(.ListCount - 1, 0) = "***TODOS***"
        .List(.ListCount - 1, 1) = 0
    End With
    
    For Each n In col
        
        oFornecedor.Carrega CLng(n)
    
        With cbbFltFornecedor
            .AddItem
            .List(.ListCount - 1, 0) = oFornecedor.Nome
            .List(.ListCount - 1, 1) = oFornecedor.ID
        End With
        
    Next n
    
    cbbFltFornecedor.ListIndex = 0

End Sub
Private Sub Filtros(Acao As String)

    Dim b As Boolean
    
    b = IIf(Acao = "Habilitar", True, False)

    cbbFltFornecedor.Enabled = b: lblFltFornecedor.Enabled = b
    btnFiltrar.Enabled = b
    frmFiltro.Enabled = b

End Sub
Private Sub btnFiltrar_Click()

    Dim lFornecedorID As Long
    
    If cbbFltFornecedor.ListIndex = -1 Then
        lFornecedorID = 0
    Else
        lFornecedorID = CLng(cbbFltFornecedor.List(cbbFltFornecedor.ListIndex, 1))
    End If

    Set myRst = oTituloPagar.Recordset(lFornecedorID)
    
    If myRst.PageCount > 0 Then
    
        myRst.AbsolutePage = myRst.PageCount
        
        With scrPagina
            .Max = myRst.PageCount
            .Value = myRst.PageCount
        End With
        
        Call scrPagina_Change
        
    End If

End Sub
Private Sub txbValor_AfterUpdate()

    Dim sDecisao As String
    
    sDecisao = Replace(btnConfirmar.Caption, "Confirmar ", "")
    
    If sDecisao = "Inclusão" Then
        txbSaldo.Text = Format(txbValor.Text, "#,##0.00")
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
