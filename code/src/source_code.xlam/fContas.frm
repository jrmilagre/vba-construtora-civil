VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fContas 
   Caption         =   ":: Cadastro de Contas ::"
   ClientHeight    =   9045
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11325
   OleObjectBlob   =   "fContas.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fContas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private oConta              As New cConta
Private oContaMovimento     As New cContaMovimento

Private colControles        As New Collection
Private myRst               As ADODB.Recordset

'Private bListBoxOrdenando   As Boolean

Private Const sTable As String = "tbl_contas"
Private Const sCampoOrderBy As String = "nome"

Private Sub UserForm_Initialize()
     
    Call lstPrincipalPopular(sCampoOrderBy)
    Call EventosCampos
    Call Campos("Desabilitar")
    
    btnCancelar.Visible = False: btnConfirmar.Visible = False
    btnAlterar.Enabled = False
    btnExcluir.Enabled = False
    
    MultiPage1.Value = 0
    MultiPage1.Pages(2).Visible = False

End Sub
Private Sub UserForm_Terminate()
    
    Set oConta = Nothing
    Set myRst = Nothing
    
    Call Desconecta
    
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
                oConta.Inclui
                Call lstPrincipalPopular(sCampoOrderBy)
                
            ElseIf sDecisao = vbNewLine & "Altera��o" Then
                
                ' Chama m�todo para alterar dados no banco de dados
                oConta.Altera
                Call lstPrincipalPopular(sCampoOrderBy)
                    
            ElseIf sDecisao = vbNewLine & "Exclus�o" Then
                        
                ' Chama m�todo para deletar registro do banco de dados
                oConta.Exclui
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
        txbSaldoInicial.Text = Format(0, "#,##0.00")
        lstPrincipal.ListIndex = -1
        Call Campos("Limpar")
    End If
    
    If Decisao <> "Exclus�o" Then
        Call Campos("Habilitar")
        txbNome.SetFocus
    End If
    
    lstPrincipal.Enabled = False
    
    'txtPesquisa.Enabled = False
    
    
End Sub

Private Sub lstPrincipal_Change()

    If lstPrincipal.ListIndex > -1 Then
    
        btnAlterar.Enabled = True
        btnExcluir.Enabled = True
        
        oConta.Carrega (CLng(lstPrincipal.List(lstPrincipal.ListIndex, 1)))
        
        lblID.Caption = Format(IIf(oConta.ID = 0, "", oConta.ID), "00000")
        lblCabNome.Caption = oConta.Nome
        txbNome.Text = oConta.Nome
        txbSaldoInicial.Text = Format(oConta.SaldoInicial, "#,##0.00")
        
        MultiPage1.Pages(2).Visible = True
        
        Call lstExtratoPopular
                
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
    MultiPage1.Pages(2).Visible = False

End Sub
Private Sub Campos(Acao As String)

    If Acao = "Desabilitar" Then
        txbNome.Enabled = False: lblNome.Enabled = False
        txbSaldoInicial.Enabled = False: lblSaldoInicial.Enabled = False
    ElseIf Acao = "Habilitar" Then
        txbNome.Enabled = True: lblNome.Enabled = True
        txbSaldoInicial.Enabled = True: lblSaldoInicial.Enabled = True
    ElseIf Acao = "Limpar" Then
        lblID.Caption = ""
        lblCabNome.Caption = ""
        txbNome.Text = Empty
        txbSaldoInicial.Text = Format(0, "#,##0.00")
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
Private Sub lstPrincipalPopular(OrderBy As String)

    Dim col As New Collection
    
    Set col = oConta.Listar(OrderBy)
    
    With lstPrincipal
        .Clear                              ' Limpa ListBox
        .Enabled = True                     ' Habilita ListBox
        .ColumnCount = 3                    ' Determina n�mero de colunas
        .ColumnWidths = "170 pt; 0pt; 180pt;"      ' Configura largura das colunas
        .Font = "Consolas"
        
        Dim n As Variant
        
        For Each n In col
            .AddItem
            oConta.Carrega CLng(n)
            .List(.ListCount - 1, 0) = oConta.Nome
            .List(.ListCount - 1, 1) = oConta.ID
            .List(.ListCount - 1, 2) = Space(15 - Len(Format(oConta.SaldoInicial, "#,##0.00"))) & Format(oConta.SaldoInicial, "#,##0.00")
        Next n
        
    End With
    
    Call Campos("Limpar")
    
End Sub

Private Function Valida() As Boolean
    
    Valida = False
    
    If txbNome.Text = Empty Then
        MsgBox "'Nome' � um campo obrigat�rio", vbInformation: MultiPage1.Page = 1: txbNome.SetFocus
    ElseIf txbSaldoInicial.Text = Empty Then
        MsgBox "'Saldo inicial' � um campo obrigat�rio", vbInformation: MultiPage1.Page = 1: txbSaldoInicial.SetFocus
    Else
        ' Envia valores preenchidos no formul�rio para o objeto
        With oConta
            .Nome = txbNome.Text
            .SaldoInicial = CCur(txbSaldoInicial.Text)
        End With
        
        Valida = True
    End If
    
End Function
Private Sub lstExtratoPopular()

    Dim lCount      As Integer
    Dim cValor      As Currency
    Dim sHistorico  As String

    Set myRst = oContaMovimento.Recordset(CLng(lstPrincipal.List(lstPrincipal.ListIndex, 1)))

    If myRst.PageCount > 0 Then
    
        myRst.AbsolutePage = myRst.PageCount
        
        If myRst.AbsolutePage = adPosEOF Then
            lblPaginaAtual.Caption = "P�gina " & Format(myRst.PageCount, "#,##0") & " de " & Format(myRst.PageCount, "#,##0")
        Else
            lblPaginaAtual.Caption = "P�gina " & Format(myRst.AbsolutePage, "#,##0") & " de " & Format(myRst.PageCount, "#,##0")
        End If
    
        With lstExtrato
            .Clear
            .ColumnCount = 6
            .ColumnWidths = "0pt; 55pt; 65pt; 65pt; 65pt; 180pt;"
                ' Colunas:
                '   - ID (oculta)
                '   - Data
                '   - D�bito
                '   - Cr�dito
                '   - Hist�rico
            
            .Enabled = True
            .Font = "Consolas"
            
            lCount = 1
            
            While Not myRst.EOF = True And lCount <= myRst.PageSize
    
                .AddItem
    
                .List(.ListCount - 1, 0) = Format(myRst.Fields("r_e_c_n_o_").Value, "0000000000")
                .List(.ListCount - 1, 1) = myRst.Fields("data").Value
                
                cValor = myRst.Fields("valor").Value
                
                If myRst.Fields("pag_rec").Value = "P" Then
                    .List(.ListCount - 1, 2) = Space(12 - Len(Format(cValor, "#,##0.00"))) & Format(cValor, "#,##0.00")
                    .List(.ListCount - 1, 3) = Space(12 - Len("-")) & "-"
                Else
                    .List(.ListCount - 1, 2) = Space(12 - Len("-")) & "-"
                    .List(.ListCount - 1, 3) = Space(12 - Len(Format(cValor, "#,##0.00"))) & Format(cValor, "#,##0.00")
                End If
                
                .List(.ListCount - 1, 4) = Space(12 - Len("-")) & "-"
                
                If myRst.Fields("tabela_origem").Value = "tbl_lancamentos_rapidos" Then
                    sHistorico = "Lan�amento � vista"
                ElseIf myRst.Fields("tabela_origem").Value = "tbl_recebimentos" Then
                    sHistorico = "Recebimento"
                ElseIf myRst.Fields("tabela_origem").Value = "tbl_pagamentos" Then
                    sHistorico = "Pagamento"
                End If
                
                .List(.ListCount - 1, 5) = sHistorico & ": " & Format(myRst.Fields("recno_origem").Value, "0000000000")
    
                lCount = lCount + 1
                
                myRst.MoveNext
                
            Wend
    
        End With
        
    Else
        lstExtrato.Clear
    End If
    

End Sub
