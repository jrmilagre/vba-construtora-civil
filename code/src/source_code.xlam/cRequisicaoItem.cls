VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cRequisicaoItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type TRequisicaoItem
    RequisicaoID        As Long
    ItemCompraID        As Long
    ProdutoID           As Long
    ObraID              As Currency
    EtapaID             As Long
    Qtde                As Double
    Unitario            As Currency
    Total               As Currency
    Data                As Date
    Recno               As Long
End Type

Private This As TRequisicaoItem

Private Const sTable As String = "tbl_requisicoes_itens"

Public Property Get RequisicaoID() As Long: RequisicaoID = This.RequisicaoID: End Property
Public Property Let RequisicaoID(ByVal Value As Long): This.RequisicaoID = Value: End Property

Public Property Get ItemCompraID() As Long: ItemCompraID = This.ItemCompraID: End Property
Public Property Let ItemCompraID(ByVal Value As Long): This.ItemCompraID = Value: End Property

Public Property Get ProdutoID() As Long: ProdutoID = This.ProdutoID: End Property
Public Property Let ProdutoID(ByVal Value As Long): This.ProdutoID = Value: End Property

Public Property Get ObraID() As Long: ObraID = This.ObraID: End Property
Public Property Let ObraID(ByVal Value As Long): This.ObraID = Value: End Property

Public Property Get EtapaID() As Long: EtapaID = This.EtapaID: End Property
Public Property Let EtapaID(ByVal Value As Long): This.EtapaID = Value: End Property

Public Property Get Qtde() As Double: Qtde = This.Qtde: End Property
Public Property Let Qtde(ByVal Value As Double): This.Qtde = Value: End Property

Public Property Get Unitario() As Currency: Unitario = This.Unitario: End Property
Public Property Let Unitario(ByVal Value As Currency): This.Unitario = Value: End Property

Public Property Get Total() As Currency: Total = This.Total: End Property
Public Property Let Total(ByVal Value As Currency): This.Total = Value: End Property

Public Property Get Data() As Date: Data = This.Data: End Property
Public Property Let Data(ByVal Value As Date): This.Data = Value: End Property

Public Property Get Recno() As Long: Recno = This.Recno: End Property
Public Property Let Recno(ByVal Value As Long): This.Recno = Value: End Property

Public Sub Carrega(ID As Long)

    sSQL = "SELECT * FROM " & sTable & " WHERE r_e_c_n_o_ = " & ID
    
    Set rst = New ADODB.RecordSet
    
    rst.Open sSQL, cnn, adOpenStatic
    
    With rst
        Me.PagamentoID = .Fields("pagamento_id").Value
        Me.TituloID = .Fields("titulo_id").Value
        Me.ValorBaixado = .Fields("valor_baixado").Value
        Me.DataBaixa = .Fields("data_baixa").Value
        Me.Recno = .Fields("r_e_c_n_o_").Value
    End With
    
End Sub
Public Sub Inclui()

    'Set oProduto = New cProduto

    Set rst = New ADODB.RecordSet

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
        .Fields("titulo_valor").Value = Me.TituloValor
        .Fields("titulo_data").Value = Me.TituloData
        .Fields("titulo_vencimento").Value = Me.TituloVencimento
        
        .Update
    End With
    
    Set rst = Nothing
       
End Sub
Public Sub Altera(Recno As Long)
    
    Set rst = New ADODB.RecordSet
    
    ' Comando SQL
    sSQL = "SELECT * FROM " & sTable & " WHERE r_e_c_n_o_ = " & Recno

    'Armazena na variável o comando que fará a consulta SQL no BD Access
    With rst
        .Open Source:=sSQL, _
              ActiveConnection:=cnn, _
              CursorType:=adOpenDynamic, _
              LockType:=adLockOptimistic
        
        .Fields("pagamento_id").Value = Me.PagamentoID
        .Fields("titulo_id").Value = Me.TituloID
        .Fields("valor_baixado").Value = Me.ValorBaixado
        .Fields("data_baixa").Value = Me.DataBaixa
        .Fields("fornecedor_id").Value = Me.FornecedorID
        .Fields("titulo_valor").Value = Me.TituloValor
        .Fields("titulo_data").Value = Me.TituloData
        .Fields("titulo_vencimento").Value = Me.TituloVencimento
        
        ' Grava alterações
        .Update
    End With
    
    Set rst = Nothing
End Sub
Public Sub Exclui(Recno As Long)
    
    sSQL = "DELETE FROM " & sTable & " WHERE r_e_c_n_o_ = " & Recno
    
    cnn.Execute sSQL

End Sub

