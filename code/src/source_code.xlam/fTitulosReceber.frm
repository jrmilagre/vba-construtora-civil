VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fTitulosReceber 
   Caption         =   ":: Cadastro de Títulos à Receber ::"
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
Private myRst               As New ADODB.Recordset
Private bChangeScrPag       As Boolean

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
    
    ' Destrói objeto da classe cProduto
    Set oTituloReceber = Nothing
    Set oObra = Nothing
    Set oCliente = Nothing
    
    Set myRst = Nothing
    
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
                oTituloReceber.Inclui
            ElseIf sDecisao = "Alteração" Then
                oTituloReceber.AlteraTitulo oTituloReceber.Recno
            ElseIf sDecisao = "Exclusão" Then
                oTituloReceber.ExcluiTitulo oTituloReceber.Recno
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
    Else
        MultiPage1.Value = 1
        MultiPage1.Pages(0).Enabled = False
    End If
    
End Sub
Private Sub lstPrincipalPopular(Pagina As Long)

    Dim lCount      As Long
    Dim cVlrBxd     As Currency
    Dim cVlrSld     As Currency
    
        
    
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
    
    Call ColoreLegenda
    
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
        oTituloReceber.Carrega CLng(lstPrincipal.List(lstPrincipal.ListIndex, 0))

        ' Preenche cabeçalho
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
        cbbObra.Enabled = False: lblObra.Enabled = False
        
        MultiPage1.Pages(0).Enabled = True
        
    ElseIf Acao = "Habilitar" Then
        
        If sDecisao = "Inclusão" Then
            txbData.Enabled = True: lblData.Enabled = True: btnData.Enabled = True
            cbbObra.Enabled = True: lblObra.Enabled = True
        End If
        
        txbVencimento.Enabled = True: lblVencimento.Enabled = True: btnVencimento.Enabled = True
        txbValor.Enabled = True: lblValor.Enabled = True
        txbObservacao.Enabled = True: lblObservacao.Enabled = True
        
        MultiPage1.Pages(0).Enabled = False
        
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
        ElseIf cbbObra.ListIndex = -1 Then
            MsgBox "Campo 'Obra' é obrigatório", vbCritical
            MultiPage1.Value = 1: cbbObra.SetFocus
        ElseIf CCur(txbValor.Text) < CCur(txbBaixado.Text) Then
            MsgBox "O campo 'Valor' não pode ser menor que o campo 'Baixado'", vbCritical
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
    ElseIf Decisao = "Exclusão" Then
    
        If oTituloReceber.ExisteRecebimento(oTituloReceber.ObraID, oTituloReceber.Recno) = True Then
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
Private Sub btnFiltrar_Click()

    Dim lObraID As Long
    
    If cbbFltObra.ListIndex = -1 Then
        lObraID = 0
    Else
        lObraID = CLng(cbbFltObra.List(cbbFltObra.ListIndex, 1))
    End If

    Set myRst = oTituloReceber.Recordset(lObraID)
    
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
Private Sub txbValor_AfterUpdate()

    Dim sDecisao As String
    
    sDecisao = Replace(btnConfirmar.Caption, "Confirmar ", "")
    
    If sDecisao = "Inclusão" Then
        txbSaldo.Text = Format(txbValor.Text, "#,##0.00")
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

Private Sub ColoreLegenda()

    Dim idx         As Integer
    Dim c           As control
    
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
