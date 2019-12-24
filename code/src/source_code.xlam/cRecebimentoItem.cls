VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cRecebimentoItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type TRecebimentoItem
    RecebimentoID       As Long
    TituloID            As Date
    ValorBaixado        As Currency
    DataBaixa           As Date
    ObraID              As Long
    TituloValor         As Currency
    TituloData          As Date
    TituloVencimento    As Date
    Recno               As Long
End Type

Private This As TRecebimentoItem

Private Const sTable As String = "tbl_recebimentos_itens"

Public Property Get RecebimentoID() As Long: RecebimentoID = This.RecebimentoID: End Property
Public Property Let RecebimentoID(ByVal Value As Long): This.RecebimentoID = Value: End Property

Public Property Get TituloID() As Long: TituloID = This.TituloID: End Property
Public Property Let TituloID(ByVal Value As Long): This.TituloID = Value: End Property

Public Property Get ValorBaixado() As Currency: ValorBaixado = This.ValorBaixado: End Property
Public Property Let ValorBaixado(ByVal Value As Currency): This.ValorBaixado = Value: End Property

Public Property Get DataBaixa() As Date: DataBaixa = This.DataBaixa: End Property
Public Property Let DataBaixa(ByVal Value As Date): This.DataBaixa = Value: End Property

Public Property Get ObraID() As Long: ObraID = This.ObraID: End Property
Public Property Let ObraID(ByVal Value As Long): This.ObraID = Value: End Property

Public Property Get TituloValor() As Currency: TituloValor = This.TituloValor: End Property
Public Property Let TituloValor(ByVal Value As Currency): This.TituloValor = Value: End Property

Public Property Get TituloData() As Date: TituloData = This.TituloData: End Property
Public Property Let TituloData(ByVal Value As Date): This.TituloData = Value: End Property

Public Property Get TituloVencimento() As Date: TituloVencimento = This.TituloVencimento: End Property
Public Property Let TituloVencimento(ByVal Value As Date): This.TituloVencimento = Value: End Property

Public Property Get Recno() As Long: Recno = This.Recno: End Property
Public Property Let Recno(ByVal Value As Long): This.Recno = Value: End Property

Public Sub Carrega(ID As Long)

    sSQL = "SELECT * FROM " & sTable & " WHERE r_e_c_n_o_ = " & ID
    
    Set rst = New ADODB.RecordSet
    
    rst.Open sSQL, cnn, adOpenStatic
    
    With rst
        Me.RecebimentoID = .Fields("recebimento_id").Value
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
        
        .Fields("recebimento_id").Value = Me.RecebimentoID
        .Fields("titulo_id").Value = Me.TituloID
        .Fields("valor_baixado").Value = Me.ValorBaixado
        .Fields("data_baixa").Value = Me.DataBaixa
        .Fields("obra_id").Value = Me.ObraID
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
        
        .Fields("recebimento_id").Value = Me.RecebimentoID
        .Fields("titulo_id").Value = Me.TituloID
        .Fields("valor_baixado").Value = Me.ValorBaixado
        .Fields("data_baixa").Value = Me.DataBaixa
        .Fields("obra_id").Value = Me.ObraID
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

