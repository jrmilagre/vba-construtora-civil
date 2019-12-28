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
    ProdutoID           As Long
    ObraID              As Currency
    EtapaID             As Long
    Qtde                As Double
    UmID                As Long
    Unitario            As Currency
    Total               As Currency
    Data                As Date
    TabelaOrigem        As String
    RecnoOrigem         As Long
    Recno               As Long
End Type

Private This As TRequisicaoItem

Private Const sTable As String = "tbl_requisicoes_itens"

Public Property Get RequisicaoID() As Long: RequisicaoID = This.RequisicaoID: End Property
Public Property Let RequisicaoID(ByVal Value As Long): This.RequisicaoID = Value: End Property

Public Property Get ProdutoID() As Long: ProdutoID = This.ProdutoID: End Property
Public Property Let ProdutoID(ByVal Value As Long): This.ProdutoID = Value: End Property

Public Property Get ObraID() As Long: ObraID = This.ObraID: End Property
Public Property Let ObraID(ByVal Value As Long): This.ObraID = Value: End Property

Public Property Get EtapaID() As Long: EtapaID = This.EtapaID: End Property
Public Property Let EtapaID(ByVal Value As Long): This.EtapaID = Value: End Property

Public Property Get Qtde() As Double: Qtde = This.Qtde: End Property
Public Property Let Qtde(ByVal Value As Double): This.Qtde = Value: End Property

Public Property Get UmID() As Long: UmID = This.UmID: End Property
Public Property Let UmID(ByVal Value As Long): This.UmID = Value: End Property

Public Property Get Unitario() As Currency: Unitario = This.Unitario: End Property
Public Property Let Unitario(ByVal Value As Currency): This.Unitario = Value: End Property

Public Property Get Total() As Currency: Total = This.Total: End Property
Public Property Let Total(ByVal Value As Currency): This.Total = Value: End Property

Public Property Get Data() As Date: Data = This.Data: End Property
Public Property Let Data(ByVal Value As Date): This.Data = Value: End Property

Public Property Get TabelaOrigem() As String: TabelaOrigem = This.TabelaOrigem: End Property
Public Property Let TabelaOrigem(ByVal Value As String): This.TabelaOrigem = Value: End Property

Public Property Get RecnoOrigem() As Long: RecnoOrigem = This.RecnoOrigem: End Property
Public Property Let RecnoOrigem(ByVal Value As Long): This.RecnoOrigem = Value: End Property

Public Property Get Recno() As Long: Recno = This.Recno: End Property
Public Property Let Recno(ByVal Value As Long): This.Recno = Value: End Property

Public Sub Carrega(ID As Long)

    sSQL = "SELECT * FROM " & sTable & " WHERE r_e_c_n_o_ = " & ID
    
    Set rst = New ADODB.Recordset
    
    rst.Open sSQL, cnn, adOpenStatic
    
    With rst
        Me.RequisicaoID = .Fields("requisicao_id").Value
        Me.ProdutoID = .Fields("produto_id").Value
        Me.ObraID = .Fields("obra_id").Value
        Me.EtapaID = .Fields("etapa_id").Value
        Me.Qtde = .Fields("quantidade").Value
        Me.UmID = .Fields("um_id").Value
        Me.Unitario = .Fields("unitario").Value
        Me.Total = .Fields("total").Value
        Me.Data = .Fields("data").Value
        Me.TabelaOrigem = .Fields("tabela_origem").Value
        Me.RecnoOrigem = .Fields("recno_origem").Value
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
        
        .Fields("requisicao_id").Value = Me.RequisicaoID
        .Fields("produto_id").Value = Me.ProdutoID
        .Fields("obra_id").Value = Me.ObraID
        .Fields("etapa_id").Value = Me.EtapaID
        .Fields("quantidade").Value = Me.Qtde
        .Fields("um_id").Value = Me.UmID
        .Fields("unitario").Value = Me.Unitario
        .Fields("total").Value = Me.Total
        .Fields("data").Value = Me.Data
        .Fields("tabela_origem").Value = Me.TabelaOrigem
        .Fields("recno_origem").Value = Me.RecnoOrigem
        
        .Update
    End With
    
    Set rst = Nothing
       
End Sub
Public Sub Exclui(Recno As Long)
    
    sSQL = "DELETE FROM " & sTable & " WHERE r_e_c_n_o_ = " & Recno
    
    cnn.Execute sSQL

End Sub
Public Function GetQtdeRequisitada(ItemCompraID As Long) As Double

    Dim rTmp As New ADODB.Recordset

    sSQL = "SELECT sum(quantidade) as quantidade "
    sSQL = sSQL & "FROM tbl_requisicoes_itens "
    sSQL = sSQL & "WHERE "
    sSQL = sSQL & "tabela_origem = 'tbl_compras_itens'  "
    sSQL = sSQL & "and recno_origem = " & ItemCompraID
    
    rTmp.Open sSQL, cnn
    
    If Not IsNull(rTmp.Fields("quantidade").Value) Then
        GetQtdeRequisitada = rTmp.Fields("quantidade").Value
    Else
        GetQtdeRequisitada = 0
    End If
    
    Set rTmp = Nothing
    
End Function

