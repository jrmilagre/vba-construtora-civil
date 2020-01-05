VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fPedreiros 
   Caption         =   ":: Cadastro de Pedreiros ::"
   ClientHeight    =   7800
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9090
   OleObjectBlob   =   "fPedreiros.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fPedreiros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private oPedreiro         As New cPedreiro
Private colControles        As New Collection
Private bListBoxOrdenando   As Boolean
Private Const sTable As String = "tbl_pedreiros"
Private Const sCampoOrderBy As String = "nome"

Private Sub UserForm_Initialize()

    If Conecta = True Then
     
        Call lstPrincipalPopular(sCampoOrderBy)
        Call EventosCampos
        Call Campos("Desabilitar")
        
        btnCancelar.Visible = False: btnConfirmar.Visible = False
        btnAlterar.Enabled = False
        btnExcluir.Enabled = False
        
        MultiPage1.Value = 0
        
    End If

End Sub
Private Sub lblHdFornecedor_Click():
    Call lstPrincipalPopular(sCampoOrderBy)
End Sub

Private Sub lblHdEndereco_Click()
    Call lstPrincipalPopular("apelido")
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
    
    ' Laço para percorrer todos os TextBox e atribuir eventos
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


' Botão confirmar
Private Sub btnConfirmar_Click()
    
    Dim vbResposta  As VbMsgBoxResult
    Dim sDecisao    As String
    
    sDecisao = Replace(btnConfirmar.Caption, "Confirmar ", "")
    
    If Valida = True Then
        
        vbResposta = MsgBox("Deseja realmente fazer a " & sDecisao & "?", vbYesNo + vbQuestion, "Pergunta")
        
        If vbResposta = vbYes Then
        
            If sDecisao = vbNewLine & "Inclusão" Then
            
                oPedreiro.Crud Create
                
                Call lstPrincipalPopular(sCampoOrderBy)

                
            ElseIf sDecisao = vbNewLine & "Alteração" Then
                
                oPedreiro.Crud Update, CLng(lstPrincipal.List(lstPrincipal.ListIndex, 1))
                
                Call lstPrincipalPopular(sCampoOrderBy)

                    
            ElseIf sDecisao = vbNewLine & "Exclusão" Then
                        
                oPedreiro.Crud Delete, CLng(lstPrincipal.List(lstPrincipal.ListIndex, 1))
                
                Call lstPrincipalPopular(sCampoOrderBy)
                
            End If
            
            ' Exibe mensagem de sucesso na decisão tomada (inclusão, alteração ou exclusão do registro).
            MsgBox sDecisao & " realizada com sucesso.", vbInformation, sDecisao & " de registro"
                
        ElseIf vbResposta = vbNo Then
        
            ' Se a resposta for não, executa a rotina atribuída ao clique do botão cancelar
            Call btnCancelar_Click
            
        End If
        
        Call Campos("Limpar")                   ' Chama sub-rotina para limpar campos e objeto
        lstPrincipal.Enabled = True      ' Habilita ListBox
        Call Campos("Desabilitar")     ' Chama sub-rotina para desabilitar campos
        
        btnConfirmar.Visible = False: btnCancelar.Visible = False
        
        btnIncluir.Visible = True: btnAlterar.Visible = True: btnExcluir.Visible = True
        
        
        btnAlterar.Enabled = False          ' Desabilita botão alterar
        btnExcluir.Enabled = False          ' Desabilita botão excluir
        btnIncluir.SetFocus                 ' Coloca o foco no botão incluir
        
        ' Tira a seleção
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
        lstPrincipal.ListIndex = -1
        Call Campos("Limpar")
    End If
    
    If Decisao <> "Exclusão" Then
        Call Campos("Habilitar")
        txbNome.SetFocus
    End If
    
    lstPrincipal.Enabled = False
    
    'txtPesquisa.Enabled = False
    
    
End Sub

Private Sub lstPrincipal_Change()

    If bListBoxOrdenando = False Then
    
        If btnAlterar.Enabled = False Then btnAlterar.Enabled = True
        If btnExcluir.Enabled = False Then btnExcluir.Enabled = True
        
        If lstPrincipal.ListIndex >= 0 Then
            oPedreiro.Crud Read, CLng(lstPrincipal.List(lstPrincipal.ListIndex, 1))
        End If
        
        lblID.Caption = Format(IIf(oPedreiro.ID = 0, "", oPedreiro.ID), "00000")
        lblCabNome.Caption = oPedreiro.Nome
        txbNome.Text = oPedreiro.Nome
        txbApelido.Text = oPedreiro.Apelido
        txbPrecoM2.Text = Format(oPedreiro.PrecoM2, "#,##0.00")
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
    
    ' Tira a seleção
    lstPrincipal.ListIndex = -1

End Sub
Private Sub Campos(Acao As String)

    If Acao = "Desabilitar" Then
        txbNome.Enabled = False: lblNome.Enabled = False
        txbApelido.Enabled = False: lblApelido.Enabled = False
        txbPrecoM2.Enabled = False: lblPrecoM2.Enabled = False
    ElseIf Acao = "Habilitar" Then
        txbNome.Enabled = True: lblNome.Enabled = True
        txbApelido.Enabled = True: lblApelido.Enabled = True
        txbPrecoM2.Enabled = True: lblPrecoM2.Enabled = True
    ElseIf Acao = "Limpar" Then
        lblID.Caption = ""
        lblCabNome.Caption = ""
        txbNome.Text = ""
        txbApelido.Text = ""
        txbPrecoM2.Text = ""
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
        
        For i = ini To fim - 1  ' Laço para comparar cada item com todos os outros itens
            For j = i + 1 To fim    ' Laço para comparar item com o próximo item
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
    
    Set col = oPedreiro.Listar(OrderBy)
    
    With lstPrincipal
        .Clear                              ' Limpa ListBox
        .Enabled = True                     ' Habilita ListBox
        .ColumnCount = 3                    ' Determina número de colunas
        .ColumnWidths = "170 pt; 0pt; 180pt;"      ' Configura largura das colunas
        .Font = "Consolas"
        
        Dim n As Variant
        
        For Each n In col
            .AddItem
            oPedreiro.Carrega CLng(n)
            .List(.ListCount - 1, 0) = oPedreiro.Nome
            .List(.ListCount - 1, 1) = oPedreiro.ID
            .List(.ListCount - 1, 2) = oPedreiro.Apelido
        Next n
        
    End With
    
    Call Campos("Limpar")
    
End Sub

Private Function Valida() As Boolean
        
    Valida = False
    
    If txbNome.Text = Empty Then
        MsgBox "'Nome' é um campo obrigatório", vbInformation: txbNome.SetFocus
    Else
        ' Envia valores preenchidos no formulário para o objeto
        With oPedreiro
            .Nome = txbNome.Text
            .Apelido = txbApelido.Text
            .PrecoM2 = txbPrecoM2.Text
        End With
        
        Valida = True
    End If
    
End Function
Private Sub UserForm_Terminate()
    
    ' Destrói objeto da classe cProduto
    Set oPedreiro = Nothing
    Call Desconecta
End Sub
