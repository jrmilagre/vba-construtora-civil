VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cContaMovimento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type TContaMovimento
    ContaID         As Long
    Data            As Date
    EntradaSaida    As String
    Valor           As Currency
    TabelaOrigem    As String
    RecnoOrigem     As Long
    Recno           As Long
End Type

Private This As TContaMovimento

Private Const sTable As String = "tbl_contas_movimentos"

Public Property Get ContaID() As Long: ContaID = This.ContaID: End Property
Public Property Let ContaID(ByVal Value As Long): This.ContaID = Value: End Property

Public Property Get Data() As Date: Data = This.Data: End Property
Public Property Let Data(ByVal Value As Date): This.Data = Value: End Property

Public Property Get EntradaSaida() As String: EntradaSaida = This.EntradaSaida: End Property
Public Property Let EntradaSaida(ByVal Value As String): This.EntradaSaida = Value: End Property

Public Property Get Valor() As Currency: Valor = This.Valor: End Property
Public Property Let Valor(ByVal Value As Currency): This.Valor = Value: End Property

Public Property Get TabelaOrigem() As String: TabelaOrigem = This.TabelaOrigem: End Property
Public Property Let TabelaOrigem(ByVal Value As String): This.TabelaOrigem = Value: End Property

Public Property Get RecnoOrigem() As Long: RecnoOrigem = This.RecnoOrigem: End Property
Public Property Let RecnoOrigem(ByVal Value As Long): This.RecnoOrigem = Value: End Property

Public Property Get Recno() As Long: Recno = This.Recno: End Property
Public Property Let Recno(ByVal Value As Long): This.Recno = Value: End Property

Public Sub Carrega(ID As Long)

    sSQL = "SELECT * FROM " & sTable & " WHERE r_e_c_n_o_ = " & ID
    
    Set rst = New ADODB.RecordSet
    
    rst.Open sSQL, cnn, adOpenStatic
    
    With rst
'        Me.PagamentoID = .Fields("pagamento_id").Value
'        Me.TituloID = .Fields("titulo_id").Value
'        Me.ValorBaixado = .Fields("valor_baixado").Value
'        Me.DataBaixa = .Fields("data_baixa").Value
'        Me.Recno = .Fields("r_e_c_n_o_").Value
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
        
        .Fields("conta_id").Value = Me.ContaID
        .Fields("data").Value = Me.Data
        .Fields("e_s").Value = Me.EntradaSaida
        .Fields("valor").Value = Me.Valor
        .Fields("tabela_origem").Value = Me.TabelaOrigem
        .Fields("recno_origem").Value = Me.RecnoOrigem
        
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
        
'        .Fields("pagamento_id").Value = Me.PagamentoID
'        .Fields("titulo_id").Value = Me.TituloID
'        .Fields("valor_baixado").Value = Me.ValorBaixado
'        .Fields("data_baixa").Value = Me.DataBaixa
'        .Fields("fornecedor_id").Value = Me.FornecedorID
'        .Fields("titulo_valor").Value = Me.TituloValor
'        .Fields("titulo_data").Value = Me.TituloData
'        .Fields("titulo_vencimento").Value = Me.TituloVencimento
        
        ' Grava alterações
        .Update
    End With
    
    Set rst = Nothing
End Sub
Public Sub Exclui(Recno As Long)
    
    sSQL = "DELETE FROM " & sTable & " WHERE r_e_c_n_o_ = " & Recno
    
    cnn.Execute sSQL

End Sub

