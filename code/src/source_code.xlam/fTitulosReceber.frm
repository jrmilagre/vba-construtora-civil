VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fTitulosReceber 
   Caption         =   ":: Cadastro de T�tulos � Receber ::"
   ClientHeight    =   10560
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13320
   OleObjectBlob   =   "fTitulosReceber.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fTitulosReceber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private oTituloReceber      As New cTituloReceber
Private oObra               As New cObra
Private oCliente            As New cCliente

Private colControles        As New Collection
Private bListBoxOrdenando   As Boolean
Private myRst               As New ADODB.RecordSet

Private Const sTable As String = "tbl_titulos_receber"
Private Const sCampoOrderBy As String = "vencimento"

Private Sub UserForm_Initialize()

    Call cbbObraPopular
    
    Call cbbFltObraPopular
    
    Call EventosCampos
        
    Call btnFiltrar_Click
    
    Call btnCancelar_Click

End Sub
Private Sub UserForm_Terminate()
    
    ' Destr�i objeto da classe cProduto
    Set oTituloReceber = Nothing
    Set oObra = Nothing
    Set oCliente = Nothing
    
    Call Desconecta
    
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
        
            If sDecisao = "Inclus�o" Then
                oTituloReceber.Inclui
            ElseIf sDecisao = "Altera��o" Then
                oTituloReceber.AlteraTitulo oTituloReceber.Recno
            ElseIf sDecisao = "Exclus�o" Then
                oTituloReceber.ExcluiTitulo oTituloReceber.Recno
            End If
            
            ' Clica no bot�o filtrar para chamar a rotina de popular lstPrincipal
            Call btnFiltrar_Click
            
            ' Exibe mensagem de sucesso na decis�o tomada (inclus�o, altera��o ou exclus�o do registro).
            MsgBox sDecisao & " realizada com sucesso.", vbInformation, sDecisao & " de registro"
            
            Call btnCancelar_Click
            
        ElseIf vbResposta = vbNo Then
        
            Call btnCancelar_Click
            
        End If
        
    Else
    
        If sDecisao = "Exclus�o" Then
            
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
    
    ' Numera a p�gina posicionada
    If myRst.AbsolutePage = adPosEOF Then
        lblPaginaAtual.Caption = "P�gina " & Format(myRst.PageCount, "#,##0") & " de " & Format(myRst.PageCount, "#,##0")
    Else
        lblPaginaAtual.Caption = "P�gina " & Format(myRst.AbsolutePage, "#,##0") & " de " & Format(myRst.PageCount, "#,##0")
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
            
            cVlrBxd = oTituloReceber.GetValorBaixado(myRst.Fields("r_e_c_n_o_").Value)
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

        btnAlterar.Enabled = True
        btnExcluir.Enabled = True

        ' Carrega informa��es do lan�amento
        oTituloReceber.Carrega CLng(lstPrincipal.List(lstPrincipal.ListIndex, 0))

        ' Preenche cabe�alho
        lblCabID.Caption = IIf(oTituloReceber.Recno = 0, "", Format(oTituloReceber.Recno, "0000000000"))
        lblCabData.Caption = oTituloReceber.Data
'
'        oFornecedor.Carrega oCompra.FornecedorID
'
'        'lblCabFuncionario.Caption = oFuncionario.Funcionario
'
        ' Preenche campos
        txbData.Text = oTituloReceber.Data
        txbVencimento.Text = oTituloReceber.Vencimento
        txbValor.Text = Format(oTituloReceber.Valor, "#,##0.00")
        txbBaixado.Text = lstPrincipal.List(lstPrincipal.ListIndex, 4)
        txbSaldo.Text = lstPrincipal.List(lstPrincipal.ListIndex, 5)
        txbObservacao.Text = oTituloReceber.Observacao
        
        For n = 0 To cbbObra.ListCount - 1
            If CLng(cbbObra.List(n, 1)) = oTituloReceber.ObraID Then
                cbbObra.ListIndex = n
                Exit For
            End If
        Next n

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
            .List(.ListCount - 1, 3) = Space(9 - Len(Format(r.Fields("unitario").Value, "#,##0.00"))) & Format(r.Fields("unitario").Value, "#,##0.00")
            
            cTotal = r.Fields("quantidade").Value * r.Fields("unitario").Value
            
            .List(.ListCount - 1, 4) = Space(9 - Len(Format(cTotal, "#,##0.00"))) & Format(cTotal, "#,##0.00")
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

    Dim sDecisao As String
    
    sDecisao = Replace(btnConfirmar.Caption, "Confirmar ", "")

    If Acao = "Desabilitar" Then
        
        txbData.Enabled = False: lblData.Enabled = False: btnData.Enabled = False
        txbVencimento.Enabled = False: lblVencimento.Enabled = False: btnVencimento.Enabled = False
        txbValor.Enabled = False: lblValor.Enabled = False
        txbObservacao.Enabled = False: lblObservacao.Enabled = False
        cbbObra.Enabled = False: lblObra.Enabled = False

        
    ElseIf Acao = "Habilitar" Then
        
        If sDecisao = "Inclus�o" Then
            txbData.Enabled = True: lblData.Enabled = True: btnData.Enabled = True
            cbbObra.Enabled = True: lblObra.Enabled = True
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
        cbbObra.ListIndex = -1
    End If
    
    Call Filtros("Habilitar")

End Sub
Private Function Valida(Decisao As String) As Boolean
    
    Valida = False
    
    If Decisao = "Inclus�o" Or Decisao = "Altera��o" Then
    
        
        If txbData.Text = Empty Then
            MsgBox "Campo 'Data' � obrigat�rio", vbCritical
            MultiPage1.Value = 1: txbData.SetFocus
        ElseIf txbVencimento.Text = Empty Then
            MsgBox "Campo 'Vencimento' � obrigat�rio", vbCritical
            MultiPage1.Value = 1: txbVencimento.SetFocus
        ElseIf txbValor.Text = Empty Or CCur(txbValor.Text) = 0 Then
            MsgBox "Campo 'Valor' � obrigat�rio", vbCritical
            MultiPage1.Value = 1: txbValor.SetFocus
        ElseIf txbObservacao.Text = Empty Then
            MsgBox "Preencha o campo 'Observa��o', pode ser importante no futuro!", vbCritical
            MultiPage1.Value = 1: txbObservacao.SetFocus
        ElseIf cbbObra.ListIndex = -1 Then
            MsgBox "Campo 'Obra' � obrigat�rio", vbCritical
            MultiPage1.Value = 1: cbbObra.SetFocus
        ElseIf CCur(txbValor.Text) < CCur(txbBaixado.Text) Then
            MsgBox "O campo 'Valor' n�o pode ser menor que o campo 'Baixado'", vbCritical
            MultiPage1.Value = 1: txbValor.SetFocus
        Else

            With oTituloReceber
                .Data = CDate(txbData.Text)
                .Vencimento = CDate(txbVencimento.Text)
                .Valor = CCur(txbValor.Text)
                .ObraID = CLng(cbbObra.List(cbbObra.ListIndex, 1))
                
                oObra.Carrega CLng(cbbObra.List(cbbObra.ListIndex, 1))
                
                .ClienteID = oObra.ClienteID
                .Observacao = RTrim(txbObservacao.Text)
            End With
            
            Valida = True

        End If
    ElseIf Decisao = "Exclus�o" Then
    
        If oTituloReceber.ExisteRecebimento(oTituloReceber.ObraID, oTituloReceber.Recno) = True Then
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

Private Sub btnItemConfirmar_Click()

    Dim sDecisaoLancamento  As String
    Dim sDecisaoItem        As String
    
    sDecisaoLancamento = Replace(btnConfirmar.Caption, "Confirmar ", "")
    sDecisaoItem = btnItemConfirmar.Caption
    
    If sDecisaoItem = "Incluir" Then
    
        If ValidaItem = True Then
            
            With lstItens
                .ColumnCount = 8
                .ColumnWidths = "200pt; 0pt; 60pt; 60pt; 60pt; 0pt; 40pt; 0pt;"
                .Font = "Consolas"
                .AddItem
                
                .List(.ListCount - 1, 0) = cbbProduto.List(cbbProduto.ListIndex, 0)
                .List(.ListCount - 1, 1) = cbbProduto.List(cbbProduto.ListIndex, 1)
                .List(.ListCount - 1, 2) = Space(9 - Len(Format(CDbl(txbQtde.Text), "#,##0.00"))) & Format(CDbl(txbQtde.Text), "#,##0.00")
                .List(.ListCount - 1, 3) = Space(9 - Len(Format(CDbl(txbUnitario.Text), "#,##0.00"))) & Format(CDbl(txbUnitario.Text), "#,##0.00")
                .List(.ListCount - 1, 4) = Space(9 - Len(Format(CDbl(txbTotal.Text), "#,##0.00"))) & Format(CDbl(txbTotal.Text), "#,##0.00")
                .List(.ListCount - 1, 6) = cbbUM.List(cbbUM.ListIndex, 0)
                .List(.ListCount - 1, 7) = cbbUM.List(cbbUM.ListIndex, 1)
                
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
                .List(.ListIndex, 6) = cbbUM.List(cbbUM.ListIndex, 0)
                .List(.ListIndex, 7) = cbbUM.List(cbbUM.ListIndex, 1)
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

Private Sub cbbProduto_AfterUpdate()

    Dim vbResposta As VbMsgBoxResult
    Dim idx As Integer
    Dim n As Integer
    
    If cbbProduto.ListIndex = -1 And cbbProduto.Text <> "" Then
        
        vbResposta = MsgBox("Este produto n�o existe, deseja cadastr�-lo?", vbQuestion + vbYesNo)
        
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
        
        vbResposta = MsgBox("Esta unidade de medida n�o existe. Deseja cadastr�-la?", vbQuestion + vbYesNo)
        
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
Private Sub cbbObraPopular()
    
    Dim idx         As Long
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

    Set myRst = oTituloReceber.RecordSet(lObraID)
    
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
    
    If sDecisao = "Inclus�o" Then
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

    ' Trata bot�es de navega��o
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
