VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cPagamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type TPagamento
    Id              As Long
    Data            As Date
    FornecedorID    As Long
    TipoBaixa       As String
    ValorPago       As Currency
End Type

Private This As TPagamento

Private Const sTable As String = "tbl_pagamentos"

Public Property Get Id() As Long: Id = This.Id: End Property
Public Property Let Id(ByVal Value As Long): This.Id = Value: End Property

Public Property Get Data() As Date: Data = This.Data: End Property
Public Property Let Data(ByVal Value As Date): This.Data = Value: End Property

Public Property Get FornecedorID() As Long: FornecedorID = This.FornecedorID: End Property
Public Property Let FornecedorID(ByVal Value As Long): This.FornecedorID = Value: End Property

Public Property Get TipoBaixa() As String: TipoBaixa = This.TipoBaixa: End Property
Public Property Let TipoBaixa(ByVal Value As String): This.TipoBaixa = Value: End Property

Public Property Get ValorPago() As Currency: ValorPago = This.ValorPago: End Property
Public Property Let ValorPago(ByVal Value As Currency): This.ValorPago = Value: End Property

' Rotina para carregar dados
Public Sub Carrega(Id As Long)

    ' Consulta SQL
    sSQL = "SELECT * FROM " & sTable & " WHERE id = " & Id
    
    ' Cria objeto Recordset
    Set rst = New ADODB.RecordSet
    
    ' Atribui resultado da consulta SQL ao recordset
    rst.Open sSQL, cnn, adOpenStatic
    
    '---Propriedades
    With rst
        Me.Id = .Fields("id").Value
        Me.Data = .Fields("data").Value
        Me.FornecedorID = .Fields("fornecedor_id").Value
        Me.TipoBaixa = .Fields("tipo_baixa").Value
        Me.ValorPago = .Fields("valor_pago").Value
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
        
        Me.Id = .Fields("id").Value
        .Fields("data").Value = Me.Data
        .Fields("fornecedor_id").Value = Me.FornecedorID
        .Fields("tipo_baixa").Value = Me.TipoBaixa
        .Fields("valor_pago").Value = Me.ValorPago
        
        .Update
    End With
    
    Set rst = Nothing
       
End Sub

Public Sub Exclui(Id As Long)
        
    sSQL = "DELETE FROM " & sTable & " WHERE id = " & Id
    cnn.Execute sSQL
    
End Sub
Public Function RecordSet() As ADODB.RecordSet

    sSQL = "SELECT id, data, fornecedor_id, tipo_baixa, valor_pago "
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
