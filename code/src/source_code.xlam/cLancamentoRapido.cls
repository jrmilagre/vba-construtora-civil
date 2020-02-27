VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cLancamentoRapido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type TLancamentoRapido
    ID              As Long
    Data            As Date
    ContaID         As Long
    CliForID        As Long
    CategoriaID     As Long
    PagRec          As String
    Valor           As Currency
    Requisitado     As Boolean
    RequisicaoID    As Variant
End Type

Private This As TLancamentoRapido

Private Const sTable As String = "tbl_lancamentos_rapidos"

Public Property Get ID() As Long: ID = This.ID: End Property
Public Property Let ID(ByVal Value As Long): This.ID = Value: End Property

Public Property Get Data() As Date: Data = This.Data: End Property
Public Property Let Data(ByVal Value As Date): This.Data = Value: End Property

Public Property Get ContaID() As Long: ContaID = This.ContaID: End Property
Public Property Let ContaID(ByVal Value As Long): This.ContaID = Value: End Property

Public Property Get CliForID() As Long: CliForID = This.CliForID: End Property
Public Property Let CliForID(ByVal Value As Long): This.CliForID = Value: End Property

Public Property Get CategoriaID() As Long: CategoriaID = This.CategoriaID: End Property
Public Property Let CategoriaID(ByVal Value As Long): This.CategoriaID = Value: End Property

Public Property Get Valor() As Currency: Valor = This.Valor: End Property
Public Property Let Valor(ByVal Value As Currency): This.Valor = Value: End Property

Public Property Get PagRec() As String: PagRec = This.PagRec: End Property
Public Property Let PagRec(ByVal Value As String): This.PagRec = Value: End Property

Public Property Get Requisitado() As Boolean: Requisitado = This.Requisitado: End Property
Public Property Let Requisitado(ByVal Value As Boolean): This.Requisitado = Value: End Property

Public Property Get RequisicaoID() As Variant: RequisicaoID = This.RequisicaoID: End Property
Public Property Let RequisicaoID(ByVal Value As Variant): This.RequisicaoID = Value: End Property


' Rotina para carregar dados
Public Sub Carrega(ID As Long)

    ' Consulta SQL
    sSQL = "SELECT * FROM " & sTable & " WHERE id = " & ID
    
    ' Cria objeto Recordset
    Set rst = New ADODB.Recordset
    
    ' Atribui resultado da consulta SQL ao recordset
    rst.Open sSQL, cnn, adOpenStatic
    
    '---Propriedades
    With rst
        Me.ID = .Fields("id").Value
        Me.Data = .Fields("data").Value
        Me.ContaID = .Fields("conta_id").Value
        Me.CliForID = .Fields("clifor_id").Value
        Me.CategoriaID = .Fields("categoria_id").Value
        Me.PagRec = .Fields("pag_rec").Value
        Me.Valor = .Fields("valor").Value
        Me.Requisitado = .Fields("requisitado").Value
        
        If .Fields("requisitado").Value = True Then
            Me.RequisicaoID = .Fields("requisicao_id").Value
        Else
            Me.RequisicaoID = Null
        End If
    End With
    
End Sub
Public Sub Inclui()

    Set rst = New ADODB.Recordset

    'Armazena na variável o comando que fará a consulta SQL no BD Access
    With rst
        .Open Source:=sTable, _
              ActiveConnection:=cnn, _
              CursorType:=adOpenDynamic, _
              LockType:=adLockOptimistic, _
              Options:=adCmdTable
        .AddNew
        
        Me.ID = .Fields("id").Value
        .Fields("data").Value = Me.Data
        .Fields("conta_id").Value = Me.ContaID
        .Fields("clifor_id").Value = Me.CliForID
        .Fields("categoria_id").Value = Me.CategoriaID
        .Fields("pag_rec").Value = Me.PagRec
        .Fields("valor").Value = Me.Valor
        .Fields("requisitado").Value = Me.Requisitado
        
        .Update
    End With
    
    Set rst = Nothing
       
End Sub

Public Sub Exclui(ID As Long)
        
    sSQL = "DELETE FROM " & sTable & " WHERE id = " & ID
    cnn.Execute sSQL
    
End Sub
Public Function Recordset() As ADODB.Recordset

    sSQL = "SELECT * "
    sSQL = sSQL & "FROM " & sTable & " "
    sSQL = sSQL & "ORDER BY id"
    
    Set Recordset = New ADODB.Recordset
    
    With Recordset
        .CursorLocation = adUseServer
        .CursorType = adOpenStatic
        .LockType = adLockReadOnly
        .PageSize = 24
        .Open sSQL, cnn
    End With
    
End Function
Public Sub AtualizaCampoRequisicaoID(LancamentoID As Long, RequisicaoID As Long)

    sSQL = "UPDATE tbl_lancamentos_rapidos "
    sSQL = sSQL & "SET requisicao_id = " & RequisicaoID & " "
    sSQL = sSQL & "WHERE id = " & LancamentoID
    
    cnn.Execute sSQL

End Sub
Public Sub ExcluiMovimentacaoContaVinculada(RequisicaoID As Long)

    sSQL = "DELETE FROM tbl_contas_movimentos "
    sSQL = sSQL & "WHERE "
    sSQL = sSQL & "tabela_origem = 'tbl_lancamentos_rapidos' "
    sSQL = sSQL & "and recno_origem = " & RequisicaoID
    
    cnn.Execute sSQL

End Sub
Public Sub ExcluiRequisicaoVinculada(RequisicaoID As Long)

    sSQL = "DELETE FROM tbl_requisicoes_itens "
    sSQL = sSQL & "WHERE requisicao_id = " & RequisicaoID
    
    cnn.Execute sSQL
    
    sSQL = "DELETE FROM tbl_requisicoes "
    sSQL = sSQL & "WHERE id = " & RequisicaoID
    
    cnn.Execute sSQL

End Sub

