VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cCompraItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type TCompraItem
    CompraID        As Long
    ProdutoID       As Long
    Quantidade      As Double
    Unitario        As Currency
    Data            As Date
    FornecedorID    As Long
    QuantReq        As Double
    Recno           As Long
End Type

Private This As TCompraItem

Private Const sTable As String = "tbl_compras_itens"

Public Property Get CompraID() As Long: CompraID = This.CompraID: End Property
Public Property Let CompraID(ByVal Value As Long): This.CompraID = Value: End Property

Public Property Get ProdutoID() As Long: ProdutoID = This.ProdutoID: End Property
Public Property Let ProdutoID(ByVal Value As Long): This.ProdutoID = Value: End Property

Public Property Get Quantidade() As Double: Quantidade = This.Quantidade: End Property
Public Property Let Quantidade(ByVal Value As Double): This.Quantidade = Value: End Property

Public Property Get Unitario() As Currency: Unitario = This.Unitario: End Property
Public Property Let Unitario(ByVal Value As Currency): This.Unitario = Value: End Property

Public Property Get Data() As Date: Data = This.Data: End Property
Public Property Let Data(ByVal Value As Date): This.Data = Value: End Property

Public Property Get FornecedorID() As Variant: FornecedorID = This.FornecedorID: End Property
Public Property Let FornecedorID(ByVal Value As Variant): This.FornecedorID = Value: End Property

Public Property Get QuantReq() As Double: QuantReq = This.QuantReq: End Property
Public Property Let QuantReq(ByVal Value As Double): This.QuantReq = Value: End Property

Public Property Get Recno() As Long: Recno = This.Recno: End Property
Public Property Let Recno(ByVal Value As Long): This.Recno = Value: End Property

Public Sub Carrega(ID As Long)

    sSQL = "SELECT * FROM " & sTable & " WHERE r_e_c_n_o_ = " & ID
    
    Set rst = New ADODB.RecordSet
    
    rst.Open sSQL, cnn, adOpenStatic
    
    With rst
        Me.CompraID = .Fields("compra_id").Value
        Me.Data = .Fields("data").Value
        Me.FornecedorID = IIf(IsNull(.Fields("fornecedor_id").Value), Null, .Fields("fornecedor_id").Value)
        Me.Quantidade = .Fields("quantidade").Value
        Me.Unitario = .Fields("unitario").Value
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
        
        .Fields("compra_id").Value = Me.CompraID
        .Fields("produto_id").Value = Me.ProdutoID
        .Fields("quantidade").Value = Me.Quantidade
        .Fields("unitario").Value = Me.Unitario
        .Fields("total").Value = Me.Quantidade * Me.Unitario
        .Fields("data").Value = CLng(Data)
        .Fields("fornecedor_id").Value = FornecedorID
        
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
        
        .Fields("compra_id").Value = Me.CompraID
        .Fields("produto_id").Value = Me.ProdutoID
        .Fields("quantidade").Value = Me.Quantidade
        .Fields("unitario").Value = Me.Unitario
        .Fields("total").Value = Me.Quantidade * Me.Unitario
        .Fields("data").Value = CLng(Data)
        .Fields("fornecedor_id").Value = FornecedorID
        
        ' Grava alterações
        .Update
    End With
    
    Set rst = Nothing
End Sub
Public Sub Exclui(Recno As Long)

    Dim rTmp As New ADODB.RecordSet
    
    sSQL = "SELECT * FROM tbl_compras_itens WHERE r_e_c_n_o_ = " & Recno
    
    rTmp.Open sSQL, cnn
    
    Set rTmp = Nothing
    
    sSQL = "DELETE FROM tbl_compras_itens WHERE r_e_c_n_o_ = " & Recno
    
    cnn.Execute sSQL

End Sub
Public Function GetQtdeBaixada(ItemCompraID As Long) As Currency

    Dim rTmp As New ADODB.RecordSet

    sSQL = "SELECT sum(quantidade) as quantidade FROM tbl_requisicoes_itens WHERE itemcompra_id = " & ItemCompraID
    
    rTmp.Open sSQL, cnn
    
    If Not IsNull(rTmp.Fields("quantidade").Value) Then
        GetQtdeBaixada = rTmp.Fields("quantidade").Value
    Else
        GetQtdeBaixada = 0
    End If
    
    Set rTmp = Nothing
    
End Function

