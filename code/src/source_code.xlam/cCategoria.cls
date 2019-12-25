VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cCategoria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type TCategoria
    Id                  As Long
    PagRec              As String
    Categoria           As String
    Subcategoria        As String
    ItemSubcategoria    As String
End Type

Private This As TCategoria

Private Const sTable As String = "tbl_categorias"

Public Property Get Id() As Long: Id = This.Id: End Property
Public Property Let Id(ByVal Value As Long): This.Id = Value: End Property

Public Property Get PagRec() As String: PagRec = This.PagRec: End Property
Public Property Let PagRec(ByVal Value As String): This.PagRec = Value: End Property

Public Property Get Categoria() As String: Categoria = This.Categoria: End Property
Public Property Let Categoria(ByVal Value As String): This.Categoria = Value: End Property

Public Property Get Subcategoria() As String: Subcategoria = This.Subcategoria: End Property
Public Property Let Subcategoria(ByVal Value As String): This.Subcategoria = Value: End Property

Public Property Get ItemSubcategoria() As String: ItemSubcategoria = This.ItemSubcategoria: End Property
Public Property Let ItemSubcategoria(ByVal Value As String): This.ItemSubcategoria = Value: End Property

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
        Me.PagRec = .Fields("pag_rec").Value
        Me.Categoria = .Fields("categoria").Value
        Me.Subcategoria = .Fields("subcategoria").Value
        Me.ItemSubcategoria = IIf(IsNull(.Fields("item_subcategoria").Value), "", .Fields("item_subcategoria").Value)
    End With
    
End Sub

' Rotina para incluir registro no banco de dados
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
        .Fields("pag_rec").Value = Me.PagRec
        .Fields("categoria").Value = Me.Categoria
        .Fields("subcategoria").Value = Me.Subcategoria
        .Fields("item_subcategoria").Value = Me.ItemSubcategoria
        
        .Update
    End With
    
    Set rst = Nothing
       

End Sub

' Rotina para alterar registro no banco de dados
Public Sub Altera()
    
    Set rst = New ADODB.RecordSet
    
    ' Comando SQL
    sSQL = "SELECT * FROM " & sTable & " WHERE id = " & Me.Id

    'Armazena na variável o comando que fará a consulta SQL no BD Access
    With rst
        .Open Source:=sSQL, _
              ActiveConnection:=cnn, _
              CursorType:=adOpenDynamic, _
              LockType:=adLockOptimistic
        
        ' Atribui novos valores aos campos
        .Fields("pag_rec").Value = Me.PagRec
        .Fields("categoria").Value = Me.Categoria
        .Fields("subcategoria").Value = Me.Subcategoria
        .Fields("item_subcategoria").Value = Me.ItemSubcategoria
        
        ' Grava alterações
        .Update
    End With
    
    Set rst = Nothing
End Sub

' Rotina para Excluir registro do banco de dados
Public Sub Exclui()
    
    sSQL = "UPDATE " & sTable & " SET deletado = True WHERE id = " & Me.Id
    
    cnn.Execute sSQL
    
End Sub
Public Function Listar(OrderBy As String, PagRec As String) As Collection

    Dim col As New Collection

    sSQL = "SELECT id "
    sSQL = sSQL & "FROM " & sTable & " "
    sSQL = sSQL & "WHERE deletado = False "
    
    If PagRec <> "T" Then
        sSQL = sSQL & "and pag_rec = '" & PagRec & "' "
    End If
    
    sSQL = sSQL & "ORDER BY " & OrderBy

    Set rst = New ADODB.RecordSet
    
    rst.Open sSQL, cnn, adOpenStatic

    Do Until rst.EOF = True
        col.Add rst.Fields("id").Value
        rst.MoveNext
    Loop

    Set rst = Nothing
    
    Set Listar = col

End Function
Public Function Existe(Categoria As String) As Boolean

    Existe = False
    
    ' Comando SQL
    sSQL = "SELECT * FROM " & sTable & " WHERE categoria = '" & Replace(Categoria, "'", "`") & "'"

    ' Atribui resultado da consulta SQL ao recordset
    Set rst = cnn.Execute(sSQL)
    
    If rst.EOF = False Then
        Me.Id = rst.Fields("id").Value
        Existe = True
    End If
    
    ' Destrói Recordset
    Set rst = Nothing
        
End Function

