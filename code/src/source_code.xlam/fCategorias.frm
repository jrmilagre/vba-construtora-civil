VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fCategorias 
   Caption         =   ":: Cadastro de Categorias ::"
   ClientHeight    =   7800
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9120
   OleObjectBlob   =   "fCategorias.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fCategorias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private oCategoria          As New cCategoria
Private colControles        As New Collection
Private bListBoxOrdenando   As Boolean
Private Const sTable As String = "tbl_categorias"
Private Const sCampoOrderBy As String = "pag_rec DESC, categoria, subcategoria, item_subcategoria"

Private Sub UserForm_Initialize()
    
    Call cbbPagRecPopular
    Call lstPrincipalPopular(sCampoOrderBy)
    Call EventosCampos
    Call Campos("Desabilitar")
    
    btnCancelar.Visible = False: btnConfirmar.Visible = False
    btnAlterar.Enabled = False
    btnExcluir.Enabled = False
    
    MultiPage1.Value = 0

End Sub
Private Sub lblHdNome_Click():
    Call lstPrincipalPopular(sCampoOrderBy)
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
    
    ' La�o para percorrer todos os TextBox e atribuir eventos
    ' de acordo com o tipo de cada campo
    For Each oControle In Me.Controls
    
        If Len(oControle.Tag) > 0 Then
        
            If TypeName(oControle) = "TextBox" Then
                
                Set oEvento = New c_EventoCampo
                
                With oEvento
                    
                    oControle.ControlTipText = cat.Tables(sTable).Columns(oControle.Tag).Properties("Description").Value
                    
                    .FieldType = cat.Tables(sTable).Columns(oControle.Tag).Type
                    .MaxLength = cat.Tables(sTable).Columns(oControle.Tag).DefinedSize
                    .Nullable = cat.Tables(sTable).Columns(oControle.Tag).Properties("Nullable")
                    
                    Set .cGeneric = oControle
                    
                End With
                    
                colControles.Add oEvento
                
            End If
        End If
    Next

End Sub


' Bot�o confirmar
Private Sub btnConfirmar_Click()
    
    Dim vbResposta  As VbMsgBoxResult
    Dim sDecisao    As String
    
    sDecisao = Replace(btnConfirmar.Caption, "Confirmar ", "")
    
    If Valida = True Then
        
        vbResposta = MsgBox("Deseja realmente fazer a " & sDecisao & "?", vbYesNo + vbQuestion, "Pergunta")
        
        If vbResposta = vbYes Then
        
            If sDecisao = vbNewLine & "Inclus�o" Then
            
                oCategoria.Inclui
                
                Call lstPrincipalPopular(sCampoOrderBy)
                
            ElseIf sDecisao = vbNewLine & "Altera��o" Then
                
                oCategoria.Altera
                
                Call lstPrincipalPopular(sCampoOrderBy)
                    
            ElseIf sDecisao = vbNewLine & "Exclus�o" Then
                        
                oCategoria.Exclui
                
                Call lstPrincipalPopular(sCampoOrderBy)
                
            End If
            
            ' Exibe mensagem de sucesso na decis�o tomada (inclus�o, altera��o ou exclus�o do registro).
            MsgBox sDecisao & " realizada com sucesso.", vbInformation, sDecisao & " de registro"
                
        ElseIf vbResposta = vbNo Then
        
            ' Se a resposta for n�o, executa a rotina atribu�da ao clique do bot�o cancelar
            Call btnCancelar_Click
            
        End If
        
        Call Campos("Limpar")                   ' Chama sub-rotina para limpar campos e objeto
        lstPrincipal.Enabled = True      ' Habilita ListBox
        Call Campos("Desabilitar")     ' Chama sub-rotina para desabilitar campos
        
        btnConfirmar.Visible = False: btnCancelar.Visible = False
        
        btnIncluir.Visible = True: btnAlterar.Visible = True: btnExcluir.Visible = True
        
        
        btnAlterar.Enabled = False          ' Desabilita bot�o alterar
        btnExcluir.Enabled = False          ' Desabilita bot�o excluir
        btnIncluir.SetFocus                 ' Coloca o foco no bot�o incluir
        
        ' Tira a sele��o
        If lstPrincipal.ListIndex >= 0 Then lstPrincipal.Selected(lstPrincipal.ListIndex) = False
        
        MultiPage1.Value = 0
        
    End If
End Sub

'Private Sub txtPesquisa_Change()
'
'    bPesquisando = True
'    Call PopulaListBox
'    bPesquisando = False
'End Sub
Private Sub btnIncluir_Click()
    Call PosDecisaoTomada("Inclus�o")
End Sub
Private Sub btnAlterar_Click()
    Call PosDecisaoTomada("Altera��o")
End Sub
Private Sub btnExcluir_Click()
    Call PosDecisaoTomada("Exclus�o")
End Sub
Private Sub PosDecisaoTomada(Decisao As String)

    btnConfirmar.Visible = True: btnCancelar.Visible = True
    btnConfirmar.Caption = "Confirmar " & VBA.vbNewLine & Decisao
    btnCancelar.Caption = "Cancelar " & VBA.vbNewLine & Decisao
    
    btnIncluir.Visible = False: btnAlterar.Visible = False: btnExcluir.Visible = False
    
    MultiPage1.Value = 1
    
    If Decisao = "Inclus�o" Then
        lstPrincipal.ListIndex = -1
        Call Campos("Limpar")
    End If
    
    If Decisao <> "Exclus�o" Then
        Call Campos("Habilitar")
        cbbPagRec.SetFocus
    End If
    
    lstPrincipal.Enabled = False
    
    'txtPesquisa.Enabled = False
    
    
End Sub

Private Sub lstPrincipal_Change()

    Dim n As Integer

    If bListBoxOrdenando = False Then
    
        If btnAlterar.Enabled = False Then btnAlterar.Enabled = True
        If btnExcluir.Enabled = False Then btnExcluir.Enabled = True
        
        If lstPrincipal.ListIndex >= 0 Then
            oCategoria.Carrega (CLng(lstPrincipal.List(lstPrincipal.ListIndex, 1)))
        End If
        
        lblID.Caption = Format(IIf(oCategoria.ID = 0, "", oCategoria.ID), "00000")
        lblCabNome.Caption = oCategoria.Categoria
        txbCategoria.Text = oCategoria.Categoria
        txbSubcategoria.Text = oCategoria.Subcategoria
        txbItemSubcategoria.Text = oCategoria.ItemSubcategoria
        
        For n = 0 To cbbPagRec.ListCount - 1
            If cbbPagRec.List(n, 1) = oCategoria.PagRec Then
                cbbPagRec.ListIndex = n
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
    'txtPesquisa.Enabled = True
    btnIncluir.SetFocus
    
    lstPrincipal.Enabled = True
   
    MultiPage1.Value = 0
    
    ' Tira a sele��o
    lstPrincipal.ListIndex = -1

End Sub
Private Sub Campos(Acao As String)

    If Acao = "Desabilitar" Then
        cbbPagRec.Enabled = False: lblPagRec.Enabled = False
        txbCategoria.Enabled = False: lblCategoria.Enabled = False
        txbSubcategoria.Enabled = False: lblSubcategoria.Enabled = False
        txbItemSubcategoria.Enabled = False: lblItemSubcategoria.Enabled = False
        
    ElseIf Acao = "Habilitar" Then
        cbbPagRec.Enabled = True: lblPagRec.Enabled = True
        txbCategoria.Enabled = True: lblCategoria.Enabled = True
        txbSubcategoria.Enabled = True: lblSubcategoria.Enabled = True
        txbItemSubcategoria.Enabled = True: lblItemSubcategoria.Enabled = True
    ElseIf Acao = "Limpar" Then
        lblID.Caption = ""
        lblCabNome.Caption = ""
        cbbPagRec.ListIndex = -1
        txbCategoria.Text = ""
        txbSubcategoria.Text = ""
        txbItemSubcategoria.Text = ""
        
        
        lstPrincipal.ListIndex = -1
    End If

End Sub
Private Sub ListBoxOrdenar()
    
    Dim ini, fim, i, j  As Long
    Dim sCol01          As String
    Dim sCol02          As String
    
    bListBoxOrdenando = True
    
    With lstPrincipal
        
        ini = 0
        fim = .ListCount - 1 '4 itens(0 - 3)
        
        For i = ini To fim - 1  ' La�o para comparar cada item com todos os outros itens
            For j = i + 1 To fim    ' La�o para comparar item com o pr�ximo item
                If .List(i) > .List(j) Then
                    sCol01 = .List(j, 0)
                    sCol02 = .List(j, 1)
                    .List(j, 0) = .List(i, 0)
                    .List(j, 1) = .List(i, 1)
                    .List(i, 0) = sCol01
                    .List(i, 1) = sCol02
                End If
            Next j
        Next i
    End With
    
    bListBoxOrdenando = False
    
End Sub
Private Sub lstPrincipalPopular(OrderBy As String)

    Dim col As New Collection
    
    Set col = oCategoria.Listar(OrderBy, "T")
    
    With lstPrincipal
        .Clear                              ' Limpa ListBox
        .Enabled = True                     ' Habilita ListBox
        .ColumnCount = 4                    ' Determina n�mero de colunas
        .ColumnWidths = "170 pt; 0pt; 150pt; 150pt;"      ' Configura largura das colunas
        .Font = "Consolas"
        
        Dim n As Variant
        
        For Each n In col
            .AddItem
            oCategoria.Carrega CLng(n)
            .List(.ListCount - 1, 0) = oCategoria.Categoria
            .List(.ListCount - 1, 1) = oCategoria.ID
            .List(.ListCount - 1, 2) = oCategoria.Subcategoria
            .List(.ListCount - 1, 3) = oCategoria.ItemSubcategoria
        Next n
        
    End With
    
    Call Campos("Limpar")
    
End Sub

Private Function Valida() As Boolean
    
    Valida = False
    
    If txbCategoria.Text = Empty Then
        MsgBox "'Categoria' � um campo obrigat�rio", vbInformation: txbCategoria.SetFocus
    ElseIf txbSubcategoria.Text = Empty Then
        MsgBox "'Subcategoria' � um campo obrigat�rio", vbInformation: txbSubcategoria.SetFocus
    Else
        ' Envia valores preenchidos no formul�rio para o objeto
        With oCategoria
            .PagRec = cbbPagRec.List(cbbPagRec.ListIndex, 1)
            .Categoria = txbCategoria.Text
            .Subcategoria = txbSubcategoria.Text
            .ItemSubcategoria = txbItemSubcategoria.Text
        End With
        
        Valida = True
    End If
    
End Function
Private Sub UserForm_Terminate()
    
    ' Destr�i objeto da classe cProduto
    Set oCategoria = Nothing
    Call Desconecta
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
