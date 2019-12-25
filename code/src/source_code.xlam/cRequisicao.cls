VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cRequisicao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type TRequisicao
    Id              As Long
    Data            As Date
End Type

Private This As TRequisicao

Private Const sTable As String = "tbl_requisicoes"

Public Property Get Id() As Long: Id = This.Id: End Property
Public Property Let Id(ByVal Value As Long): This.Id = Value: End Property

Public Property Get Data() As Date: Data = This.Data: End Property
Public Property Let Data(ByVal Value As Date): This.Data = Value: End Property

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
        
        .Update
    End With
    
    Set rst = Nothing
       
End Sub

Public Sub Exclui(Id As Long)
        
    sSQL = "DELETE FROM " & sTable & " WHERE id = " & Id
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

