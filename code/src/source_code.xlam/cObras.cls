VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cObras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type TProduto
    ID              As Long
    Endereco        As String
End Type

Private This As TProduto

Private Const sTable As String = "tbl_obras"

Public Property Get ID() As Long: ID = This.ID: End Property
Public Property Let ID(ByVal Value As Long): This.ID = Value: End Property

Public Property Get Endereco() As String: Endereco = This.Endereco: End Property
Public Property Let Endereco(ByVal Value As String): This.Endereco = Value: End Property

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
        Me.Endereco = .Fields("nome").Value
    End With
    
End Sub

' Rotina para incluir registro no banco de dados
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
        .Fields("endereco").Value = Me.Endereco
        
        .Update
    End With
    
    Set rst = Nothing
       

End Sub

' Rotina para alterar registro no banco de dados
Public Sub Altera()
    
    Set rst = New ADODB.Recordset
    
    ' Comando SQL
    sSQL = "SELECT * FROM " & sTable & " WHERE id = " & Me.ID

    'Armazena na variável o comando que fará a consulta SQL no BD Access
    With rst
        .Open Source:=sSQL, _
              ActiveConnection:=cnn, _
              CursorType:=adOpenDynamic, _
              LockType:=adLockOptimistic
        
        ' Atribui novos valores aos campos
        .Fields("endereco").Value = Me.Endereco
        
        ' Grava alterações
        .Update
    End With
    
    Set rst = Nothing
End Sub

' Rotina para Excluir registro do banco de dados
Public Sub Exclui()
    
    sSQL = "UPDATE " & sTable & " SET deletado = True WHERE id = " & Me.ID
    
    cnn.Execute sSQL
    
End Sub
Public Function Listar(OrderBy As String) As Collection

    Dim col As New Collection

    sSQL = "SELECT id "
    sSQL = sSQL & "FROM " & sTable & " "
    sSQL = sSQL & "WHERE deletado = False "
    sSQL = sSQL & "ORDER BY " & OrderBy

    Set rst = New ADODB.Recordset
    
    rst.Open sSQL, cnn, adOpenStatic

    Do Until rst.EOF = True
        col.Add rst.Fields("id").Value
        rst.MoveNext
    Loop

    Set rst = Nothing
    
    Set Listar = col

End Function
