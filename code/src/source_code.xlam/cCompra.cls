VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type TCompra
    ID              As Long
    Data            As Date
    FornecedorID    As Long
    CategoriaID     As Long
End Type

Private This As TCompra

Private Const sTable As String = "tbl_compras"

Public Property Get ID() As Long: ID = This.ID: End Property
Public Property Let ID(ByVal Value As Long): This.ID = Value: End Property

Public Property Get Data() As Date: Data = This.Data: End Property
Public Property Let Data(ByVal Value As Date): This.Data = Value: End Property

Public Property Get FornecedorID() As Variant: FornecedorID = This.FornecedorID: End Property
Public Property Let FornecedorID(ByVal Value As Variant): This.FornecedorID = Value: End Property

Public Property Get CategoriaID() As Variant: CategoriaID = This.CategoriaID: End Property
Public Property Let CategoriaID(ByVal Value As Variant): This.CategoriaID = Value: End Property

' Rotina para carregar dados
Public Sub Carrega(ID As Long)

    ' Consulta SQL
    sSQL = "SELECT * FROM " & sTable & " WHERE id = " & ID
    
    ' Cria objeto Recordset
    Set rst = New ADODB.RecordSet
    
    ' Atribui resultado da consulta SQL ao recordset
    rst.Open sSQL, cnn, adOpenStatic
    
    '---Propriedades
    With rst
        Me.ID = .Fields("id").Value
        Me.Data = .Fields("data").Value
        Me.FornecedorID = IIf(IsNull(.Fields("fornecedor_id").Value), Null, .Fields("fornecedor_id").Value)
        Me.CategoriaID = .Fields("categoria_id").Value
    End With
    
End Sub
Public Sub Inclui()

    Set rst = New ADODB.RecordSet

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
        .Fields("fornecedor_id").Value = Me.FornecedorID
        .Fields("categoria_id").Value = Me.CategoriaID
        
        .Update
    End With
    
    Set rst = Nothing
       
End Sub
' Rotina para alterar registro no banco de dados
Public Sub Altera(ID As Long)
    
    Set rst = New ADODB.RecordSet
    
    ' Comando SQL
    sSQL = "SELECT * FROM " & sTable & " WHERE id = " & ID

    'Armazena na variável o comando que fará a consulta SQL no BD Access
    With rst
        .Open Source:=sSQL, _
              ActiveConnection:=cnn, _
              CursorType:=adOpenDynamic, _
              LockType:=adLockOptimistic
        
        ' Atribui novos valores aos campos
        .Fields("data").Value = Me.Data
        .Fields("fornecedor_id").Value = Me.FornecedorID
        .Fields("categoria_id").Value = Me.CategoriaID
        
        ' Grava alterações
        .Update
    End With
    
    Set rst = Nothing
End Sub
Public Sub Exclui(ID As Long)
        
    sSQL = "DELETE FROM " & sTable & " WHERE id = " & ID
    cnn.Execute sSQL
    
End Sub
Public Function RecordSet() As ADODB.RecordSet

    sSQL = "SELECT * "
    sSQL = sSQL & "FROM " & sTable
    
    Set RecordSet = New ADODB.RecordSet
    
    With RecordSet
        .CursorLocation = adUseServer
        .CursorType = adOpenStatic
        .LockType = adLockReadOnly
        .PageSize = 24
        .Open sSQL, cnn
    End With
    
End Function
Public Function ExisteRequisicao(CompraID As Long) As Boolean

    Dim rTmp    As New ADODB.RecordSet
    Dim sMsg    As String
    
    ExisteRequisicao = False
    
    sSQL = "SELECT tbl_requisicoes_itens.requisicao_id "
    sSQL = sSQL & "FROM tbl_requisicoes_itens "
    sSQL = sSQL & "INNER JOIN tbl_compras_itens ON tbl_requisicoes_itens.recno_origem = tbl_compras_itens.r_e_c_n_o_ "
    sSQL = sSQL & "GROUP BY tbl_requisicoes_itens.tabela_origem, tbl_compras_itens.compra_id, tbl_requisicoes_itens.requisicao_id "
    sSQL = sSQL & "HAVING (((tbl_requisicoes_itens.tabela_origem)='tbl_compras_itens') AND ((tbl_compras_itens.compra_id)=" & CompraID & "));"
    
    rTmp.Open sSQL, cnn, adOpenStatic
    
    If rTmp.EOF = False Then
    
        sMsg = "Não posso excluir pois existem requisições vinculadas a essa compra!"
        sMsg = sMsg & vbNewLine
        sMsg = sMsg & vbNewLine
        sMsg = sMsg & "Veja o número das requisições:"
        sMsg = sMsg & vbNewLine
        sMsg = sMsg & vbNewLine
        
        Do Until rTmp.EOF
        
            sMsg = sMsg & "Nº:" & Format(rTmp.Fields("requisicao_id").Value, "0000000000") & vbNewLine
        
            rTmp.MoveNext
        Loop
        
        sMsg = sMsg & vbNewLine
        sMsg = sMsg & "Fim!"
        
        MsgBox sMsg, vbCritical, "EXISTEM REQUISIÇÕES"
            
        ExisteRequisicao = True
    Else
        ExisteRequisicao = False
    End If
        
    Set rTmp = Nothing

End Function
Public Function ExistePagamento(CompraID As Long) As Boolean

    Dim rTmp    As New ADODB.RecordSet
    Dim sMsg    As String
    
    ExisteBaixaDeTitulos = False
    
    sSQL = "SELECT tbl_pagamentos_itens.pagamento_id "
    sSQL = sSQL & "FROM tbl_titulos_pagar "
    sSQL = sSQL & "INNER JOIN tbl_pagamentos_itens ON tbl_titulos_pagar.r_e_c_n_o_ = tbl_pagamentos_itens.titulo_id "
    sSQL = sSQL & "GROUP BY tbl_titulos_pagar.compra_id, tbl_pagamentos_itens.pagamento_id "
    sSQL = sSQL & "HAVING (((tbl_titulos_pagar.compra_id)=" & CompraID & "))"
    
    rTmp.Open sSQL, cnn, adOpenStatic
    
    If rTmp.EOF = False Then
    
        sMsg = "Não posso excluir pois existem pagamentos vinculados a essa compra!"
        sMsg = sMsg & vbNewLine
        sMsg = sMsg & vbNewLine
        sMsg = sMsg & "Veja o número dos pagamentos:"
        sMsg = sMsg & vbNewLine
        sMsg = sMsg & vbNewLine
        
        Do Until rTmp.EOF
        
            sMsg = sMsg & "Nº:" & Format(rTmp.Fields("pagamento_id").Value, "0000000000") & vbNewLine
        
            rTmp.MoveNext
        Loop
        
        sMsg = sMsg & vbNewLine
        sMsg = sMsg & "Fim!"
        
        MsgBox sMsg, vbCritical, "EXISTEM PAGAMENTOS"
            
        ExistePagamento = True
    Else
        ExistePagamento = False
    End If
        
    Set rTmp = Nothing

End Function
