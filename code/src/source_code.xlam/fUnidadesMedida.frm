VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fUnidadesMedida 
   Caption         =   ":: Cadastro de Unidades de Medida ::"
   ClientHeight    =   7800
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9120
   OleObjectBlob   =   "fUnidadesMedida.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fUnidadesMedida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private oUnidadeMedida      As New cUnidadeMedida
Private colControles        As New Collection
Private bListBoxOrdenando   As Boolean

Private Const sTable        As String = "tbl_unidades_medida"
Private Const sCampoOrderBy As String = "abreviacao"

Private Sub UserForm_Initialize()
     
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
                oUnidadeMedida.Inclui
                Call lstPrincipalPopular(sCampoOrderBy)
                ' Inclui registro na ListBox
'                With lstPrincipal
'                    .AddItem
'                    .List(.ListCount - 1, 0) = oUnidadeMedida.NomeFantasia
'                    .List(.ListCount - 1, 1) = oUnidadeMedida.ID
'                    .List(.ListCount - 1, 2) = oUnidadeMedida.Endereco
'                End With
                    
                
                'Call ListBoxOrdenar
                
            ElseIf sDecisao = vbNewLine & "Altera��o" Then
                
                ' Chama m�todo para alterar dados no banco de dados
                oUnidadeMedida.Altera
                Call lstPrincipalPopular(sCampoOrderBy)
                ' Replica as altera��es na ListBox
'                With lstPrincipal
'                    .List(.ListIndex, 0) = oUnidadeMedida.NomeFantasia
'                    .List(.ListIndex, 2) = oUnidadeMedida.Endereco
'                End With
                
                
                'Call ListBoxOrdenar
                    
            ElseIf sDecisao = vbNewLine & "Exclus�o" Then
                        
                ' Chama m�todo para deletar registro do banco de dados
                oUnidadeMedida.Exclui
                Call lstPrincipalPopular(sCampoOrderBy)
                ' Remove item da ListBox
                'lstPrincipal.RemoveItem (lstPrincipal.ListIndex)
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
        txbAbreviacao.SetFocus
    End If
    
    lstPrincipal.Enabled = False
    
    'txtPesquisa.Enabled = False
    
    
End Sub

Private Sub lstPrincipal_Change()

    If bListBoxOrdenando = False Then
    
        If btnAlterar.Enabled = False Then btnAlterar.Enabled = True
        If btnExcluir.Enabled = False Then btnExcluir.Enabled = True
        
        If lstPrincipal.ListIndex >= 0 Then
            oUnidadeMedida.Carrega (CLng(lstPrincipal.List(lstPrincipal.ListIndex, 1)))
        End If
        
        lblID.Caption = Format(IIf(oUnidadeMedida.Id = 0, "", oUnidadeMedida.Id), "00000")
        lblCabNome.Caption = oUnidadeMedida.Abreviacao
        txbNome.Text = oUnidadeMedida.Nome
        txbAbreviacao.Text = oUnidadeMedida.Abreviacao
                
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
        txbNome.Enabled = False: lblNome.Enabled = False
        txbAbreviacao.Enabled = False: lblAbreviacao.Enabled = False
    ElseIf Acao = "Habilitar" Then
        txbNome.Enabled = True: lblNome.Enabled = True
        txbAbreviacao.Enabled = True: lblAbreviacao.Enabled = True
    ElseIf Acao = "Limpar" Then
        lblID.Caption = ""
        lblCabNome.Caption = ""
        txbNome.Text = Empty
        txbAbreviacao.Text = Empty
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
    
    Set col = oUnidadeMedida.Listar(OrderBy)
    
    With lstPrincipal
        .Clear                              ' Limpa ListBox
        .Enabled = True                     ' Habilita ListBox
        .ColumnCount = 3                    ' Determina n�mero de colunas
        .ColumnWidths = "40 pt; 0pt; 170pt;"      ' Configura largura das colunas
        .Font = "Consolas"
        
        Dim n As Variant
        
        For Each n In col
            .AddItem
            oUnidadeMedida.Carrega CLng(n)
            .List(.ListCount - 1, 0) = oUnidadeMedida.Abreviacao
            .List(.ListCount - 1, 1) = oUnidadeMedida.Id
            .List(.ListCount - 1, 2) = oUnidadeMedida.Nome
        Next n
        
    End With
    
    Call Campos("Limpar")
    
End Sub

Private Function Valida() As Boolean
    
    Valida = False
    
    If txbNome.Text = Empty Then
        MsgBox "'Nome' � um campo obrigat�rio", vbInformation: txbNome.SetFocus
    ElseIf txbAbreviacao.Text = Empty Then
        MsgBox "'Abrevia��o' � um campo obrigat�rio", vbInformation: txbAbreviacao.SetFocus
    Else
        ' Envia valores preenchidos no formul�rio para o objeto
        With oUnidadeMedida
            .Nome = txbNome.Text
            .Abreviacao = txbAbreviacao.Text
        End With
        
        Valida = True
    End If
    
End Function
Private Sub UserForm_Terminate()
    
    ' Destr�i objeto da classe cProduto
    Set oUnidadeMedida = Nothing
    Call Desconecta
End Sub
