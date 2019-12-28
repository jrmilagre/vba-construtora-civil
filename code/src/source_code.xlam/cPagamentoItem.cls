VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cPagamentoItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type TPagamentoItem
    PagamentoID         As Long
    TituloID            As Date
    ValorBaixado        As Currency
    DataBaixa           As Date
    FornecedorID        As Long
    ValorDesconto       As Currency
    ValorAcrescimo      As Currency
    
    Recno               As Long
End Type

Private This As TPagamentoItem

Private Const sTable As String = "tbl_pagamentos_itens"

Public Property Get PagamentoID() As Long: PagamentoID = This.PagamentoID: End Property
Public Property Let PagamentoID(ByVal Value As Long): This.PagamentoID = Value: End Property

Public Property Get TituloID() As Long: TituloID = This.TituloID: End Property
Public Property Let TituloID(ByVal Value As Long): This.TituloID = Value: End Property

Public Property Get ValorBaixado() As Currency: ValorBaixado = This.ValorBaixado: End Property
Public Property Let ValorBaixado(ByVal Value As Currency): This.ValorBaixado = Value: End Property

Public Property Get DataBaixa() As Date: DataBaixa = This.DataBaixa: End Property
Public Property Let DataBaixa(ByVal Value As Date): This.DataBaixa = Value: End Property

Public Property Get FornecedorID() As Long: FornecedorID = This.FornecedorID: End Property
Public Property Let FornecedorID(ByVal Value As Long): This.FornecedorID = Value: End Property

Public Property Get ValorDesconto() As Currency: ValorDesconto = This.ValorDesconto: End Property
Public Property Let ValorDesconto(ByVal Value As Currency): This.ValorDesconto = Value: End Property

Public Property Get ValorAcrescimo() As Currency: ValorAcrescimo = This.ValorAcrescimo: End Property
Public Property Let ValorAcrescimo(ByVal Value As Currency): This.ValorAcrescimo = Value: End Property

Public Property Get Recno() As Long: Recno = This.Recno: End Property
Public Property Let Recno(ByVal Value As Long): This.Recno = Value: End Property

Public Sub Carrega(ID As Long)

    sSQL = "SELECT * FROM " & sTable & " WHERE r_e_c_n_o_ = " & ID
    
    Set rst = New ADODB.Recordset
    
    rst.Open sSQL, cnn, adOpenStatic
    
    With rst
        Me.PagamentoID = .Fields("pagamento_id").Value
        Me.TituloID = .Fields("titulo_id").Value
        Me.ValorBaixado = .Fields("valor_baixado").Value
        Me.DataBaixa = .Fields("data_baixa").Value
        Me.ValorDesconto = .Fields("valor_desconto").Value
        Me.ValorAcrescimo = .Fields("valor_acrescimo").Value
        Me.Recno = .Fields("r_e_c_n_o_").Value
    End With
    
End Sub
Public Sub Inclui()

    'Set oProduto = New cProduto

    Set rst = New ADODB.Recordset

    With rst
        .Open Source:=sTable, _
              ActiveConnection:=cnn, _
              CursorType:=adOpenDynamic, _
              LockType:=adLockOptimistic, _
              Options:=adCmdTable
        .AddNew
        
        Me.Recno = .Fields("r_e_c_n_o_").Value
        
        .Fields("pagamento_id").Value = Me.PagamentoID
        .Fields("titulo_id").Value = Me.TituloID
        .Fields("valor_baixado").Value = Me.ValorBaixado
        .Fields("data_baixa").Value = Me.DataBaixa
        .Fields("fornecedor_id").Value = Me.FornecedorID
        .Fields("valor_desconto").Value = Me.ValorDesconto
        .Fields("valor_acrescimo").Value = Me.ValorAcrescimo
        
        .Update
    End With
    
    Set rst = Nothing
       
End Sub
Public Sub Exclui(PagamentoID As Long)
    
    sSQL = "DELETE FROM " & sTable & " WHERE pagamento_id = " & PagamentoID
    
    cnn.Execute sSQL
    
    sSQL = "DELETE FROM tbl_contas_movimentos "
    sSQL = sSQL & "WHERE tabela_origem = 'tbl_pagamentos' "
    sSQL = sSQL & "and recno_origem = " & PagamentoID
    
    cnn.Execute sSQL

End Sub

