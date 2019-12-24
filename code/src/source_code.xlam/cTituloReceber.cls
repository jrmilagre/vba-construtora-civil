VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTituloReceber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type TTituloPagar
    ObraID        As Long
    ClienteID    As Long
    CategoriaID     As Long
    Observacao      As String
    Vencimento      As Date
    Valor           As Currency
    Data            As Date
    Recno           As Long
End Type

Private This As TTituloPagar

Private Const sTable As String = "tbl_titulos_receber"

Public Property Get ObraID() As Long: ObraID = This.ObraID: End Property
Public Property Let ObraID(ByVal Value As Long): This.ObraID = Value: End Property

Public Property Get ClienteID() As Long: ClienteID = This.ClienteID: End Property
Public Property Let ClienteID(ByVal Value As Long): This.ClienteID = Value: End Property

Public Property Get CategoriaID() As Long: CategoriaID = This.CategoriaID: End Property
Public Property Let CategoriaID(ByVal Value As Long): This.CategoriaID = Value: End Property

Public Property Get Observacao() As String: Observacao = This.Observacao: End Property
Public Property Let Observacao(ByVal Value As String): This.Observacao = Value: End Property

Public Property Get Vencimento() As Date: Vencimento = This.Vencimento: End Property
Public Property Let Vencimento(ByVal Value As Date): This.Vencimento = Value: End Property

Public Property Get Valor() As Currency: Valor = This.Valor: End Property
Public Property Let Valor(ByVal Value As Currency): This.Valor = Value: End Property

Public Property Get Data() As Date: Data = This.Data: End Property
Public Property Let Data(ByVal Value As Date): This.Data = Value: End Property

Public Property Get Recno() As Long: Recno = This.Recno: End Property
Public Property Let Recno(ByVal Value As Long): This.Recno = Value: End Property

Public Sub Carrega(ID As Long)

    sSQL = "SELECT * FROM " & sTable & " WHERE r_e_c_n_o_ = " & ID
    
    Set rst = New ADODB.RecordSet
    
    rst.Open sSQL, cnn, adOpenStatic
    
    With rst
        Me.ObraID = .Fields("obra_id").Value
        Me.ClienteID = .Fields("cliente_id").Value
        Me.Observacao = .Fields("observacao").Value
        Me.Vencimento = .Fields("vencimento").Value
        Me.Valor = .Fields("valor").Value
        Me.Data = .Fields("data").Value
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
        
        .Fields("obra_id").Value = Me.ObraID
        .Fields("cliente_id").Value = Me.ClienteID
        .Fields("observacao").Value = Me.Observacao
        .Fields("vencimento").Value = CLng(Me.Vencimento)
        .Fields("valor").Value = Me.Valor
        .Fields("data").Value = CLng(Me.Data)
        
        .Update
    End With
    
    Set rst = Nothing
       
End Sub
Public Sub Exclui(ObraID As Long)
    
    sSQL = "DELETE FROM tbl_titulos_receber WHERE obra_id = " & ObraID
    
    cnn.Execute sSQL

End Sub
Public Function GetValorBaixado(TituloID As Long) As Currency

    Dim rTmp As New ADODB.RecordSet

    sSQL = "SELECT sum(valor_baixado) as valor_baixado FROM tbl_recebimentos_itens WHERE titulo_id = " & TituloID
    
    rTmp.Open sSQL, cnn
    
    If Not IsNull(rTmp.Fields("valor_baixado").Value) Then
        GetValorBaixado = rTmp.Fields("valor_baixado").Value
    Else
        GetValorBaixado = 0
    End If
    
    Set rTmp = Nothing
    
End Function

