VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fObras 
   Caption         =   ":: Cadastro de Obras ::"
   ClientHeight    =   7800
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9120
   OleObjectBlob   =   "fObras.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fObras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private oObra               As New cObra
Private oTipoObra           As New cTipoObra
Private oCliente            As New cCliente
Private colControles        As New Collection
Private bListBoxOrdenando   As Boolean
Private Const sTable As String = "tbl_obras"
Private Const sCampoOrderBy As String = "endereco"
Private Sub UserForm_Initialize()
     
    Call cbbTipoPopular
    Call cbbClientePopular
    Call lstPrincipalPopular(sCampoOrderBy)
    Call EventosCampos
    Call Campos("Desabilitar")
    
    btnCancelar.Visible = False: btnConfirmar.Visible = False
    btnAlterar.Enabled = False
    btnExcluir.Enabled = False
    
    MultiPage1.Value = 0

End Sub
Private Sub UserForm_Terminate()
    
    ' Destr�i objeto da classe cProduto
    Set oObra = Nothing
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
            
                ' Chama m�todo para incluir registro no banco de dados
                oObra.Inclui
                Call lstPrincipalPopular(sCampoOrderBy)
                
            ElseIf sDecisao = vbNewLine & "Altera��o" Then
                
                ' Chama m�todo para alterar dados no banco de dados
                oObra.Altera
                Call lstPrincipalPopular(sCampoOrderBy)
                    
            ElseIf sDecisao = vbNewLine & "Exclus�o" Then
                        
                ' Chama m�todo para deletar registro do banco de dados
                oObra.Exclui
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
        txbEndereco.SetFocus
    End If
    
    lstPrincipal.Enabled = False
    
    'txtPesquisa.Enabled = False
    
    
End Sub

Private Sub lstPrincipal_Change()

    Dim n As Long
    Dim iTipoID As Integer

    If bListBoxOrdenando = False Then
    
        If btnAlterar.Enabled = False Then btnAlterar.Enabled = True
        If btnExcluir.Enabled = False Then btnExcluir.Enabled = True
        
        If lstPrincipal.ListIndex >= 0 Then
            oObra.Carrega (CLng(lstPrincipal.List(lstPrincipal.ListIndex, 1)))
        End If
        
        lblID.Caption = Format(IIf(oObra.ID = 0, "", oObra.ID), "00000")
        lblCabEndereco.Caption = oObra.Endereco
        
        txbBairro.Text = oObra.Bairro
        txbCidade.Text = oObra.Cidade
        cbbUF.Text = oObra.UF
        txbEndereco.Text = oObra.Endereco
        
        For n = 0 To cbbTipo.ListCount - 1
            If CInt(cbbTipo.List(n, 1)) = oObra.TipoID Then
                cbbTipo.ListIndex = n
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
                
    End If

End Sub
Private Sub btnCancelar_Click()
    
    btnIncluir.Visible = True: btnAlterar.Visible = True: btnExcluir.Visible = True
    btnConfirmar.Visible = False: btnCancelar.Visible = False
    
    ' Tira a sele��o
    lstPrincipal.ListIndex = -1
    
    Call Campos("Limpar")
    Call Campos("Desabilitar")
    
    btnAlterar.Enabled = False
    btnExcluir.Enabled = False
    'txtPesquisa.Enabled = True
    btnIncluir.SetFocus
    
    lstPrincipal.Enabled = True
   
    MultiPage1.Value = 0
    
    


End Sub
Private Sub Campos(Acao As String)

    If Acao = "Desabilitar" Then
        txbEndereco.Enabled = False: lblEndereco.Enabled = False
        txbBairro.Enabled = False: lblBairro.Enabled = False
        txbCidade.Enabled = False: lblCidade.Enabled = False
        cbbUF.Enabled = False: lblUF.Enabled = False
        cbbTipo.Enabled = False: lblTipo.Enabled = False
        cbbCliente.Enabled = False: lblCliente.Enabled = False
    ElseIf Acao = "Habilitar" Then
        txbEndereco.Enabled = True: lblEndereco.Enabled = True
        txbBairro.Enabled = True: lblBairro.Enabled = True
        txbCidade.Enabled = True: lblCidade.Enabled = True
        cbbUF.Enabled = True: lblUF.Enabled = True
        cbbTipo.Enabled = True: lblTipo.Enabled = True
        cbbCliente.Enabled = True: lblCliente.Enabled = True
    ElseIf Acao = "Limpar" Then
        lblID.Caption = ""
        lblCabEndereco.Caption = ""
        txbEndereco.Text = ""
        txbBairro.Text = ""
        txbCidade.Text = ""
        cbbUF.ListIndex = -1
        cbbTipo.ListIndex = -1
        cbbCliente.ListIndex = -1
        
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
    
    Set col = oObra.Listar(OrderBy)
    
    With lstPrincipal
        .Clear                              ' Limpa ListBox
        .Enabled = True                     ' Habilita ListBox
        .ColumnCount = 3                    ' Determina n�mero de colunas
        .ColumnWidths = "170 pt; 0pt; 180pt;"      ' Configura largura das colunas
        .Font = "Consolas"
        
        Dim n As Variant
        
        For Each n In col
            .AddItem
            oObra.Carrega CLng(n)
            .List(.ListCount - 1, 0) = oObra.Endereco
            .List(.ListCount - 1, 1) = oObra.ID
        Next n
        
    End With
    
    Call Campos("Limpar")
    
End Sub

Private Function Valida() As Boolean
    
    Valida = False
    
    If txbEndereco.Text = Empty Then
        MsgBox "'Endere�o' � um campo obrigat�rio", vbInformation: txbEndereco.SetFocus
    ElseIf cbbTipo.ListIndex = -1 Then
        MsgBox "'Tipo' � um campo obrigat�rio", vbInformation: cbbTipo.SetFocus
    Else
        ' Envia valores preenchidos no formul�rio para o objeto
        With oObra
            .Endereco = txbEndereco.Text
            .Bairro = txbBairro.Text
            .Cidade = txbCidade.Text
            .UF = cbbUF.Text
            .TipoID = IIf(cbbTipo.ListIndex = -1, 0, CInt(cbbTipo.List(cbbTipo.ListIndex, 1)))
            
            If cbbCliente.ListIndex = -1 Then .ClienteID = Null Else .ClienteID = CLng(cbbCliente.List(cbbCliente.ListIndex, 1))
            
        End With
        
        Valida = True
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
        
        vbResposta = MsgBox("Este Tipo de obra n�o existe, deseja cadastr�-lo?", vbQuestion + vbYesNo)
        
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
