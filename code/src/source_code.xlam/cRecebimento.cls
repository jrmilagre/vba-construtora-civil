VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cRecebimento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type TRecebimento
    ID              As Long
    Data            As Date
    ObraID          As Long
    TipoBaixa       As String
    ValorRecebido   As Currency
End Type

Private This As TRecebimento

Private Const sTable As String = "tbl_recebimentos"

Public Property Get ID() As Long: ID = This.ID: End Property
Public Property Let ID(ByVal Value As Long): This.ID = Value: End Property

Public Property Get Data() As Date: Data = This.Data: End Property
Public Property Let Data(ByVal Value As Date): This.Data = Value: End Property

Public Property Get ObraID() As Long: ObraID = This.ObraID: End Property
Public Property Let ObraID(ByVal Value As Long): This.ObraID = Value: End Property

Public Property Get TipoBaixa() As String: TipoBaixa = This.TipoBaixa: End Property
Public Property Let TipoBaixa(ByVal Value As String): This.TipoBaixa = Value: End Property

Public Property Get ValorRecebido() As Currency: ValorRecebido = This.ValorRecebido: End Property
Public Property Let ValorRecebido(ByVal Value As Currency): This.ValorRecebido = Value: End Property

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
        Me.ObraID = .Fields("obra_id").Value
        Me.TipoBaixa = .Fields("tipo_baixa").Value
        Me.ValorRecebido = .Fields("valor_recebido").Value
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
        .Fields("obra_id").Value = Me.ObraID
        .Fields("tipo_baixa").Value = Me.TipoBaixa
        .Fields("valor_recebido").Value = Me.ValorRecebido
        
        .Update
    End With
    
    Set rst = Nothing
       
End Sub

Public Sub Exclui(ID As Long)
        
    sSQL = "DELETE FROM " & sTable & " WHERE id = " & ID
    cnn.Execute sSQL
    
End Sub
Public Function RecordSet() As ADODB.RecordSet

    sSQL = "SELECT id, data, obra_id, tipo_baixa, valor_recebido "
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

