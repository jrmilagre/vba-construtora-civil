VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cPedreiro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type TPedreiro
    ID              As Long
    Nome            As String
    Apelido         As String
    PrecoM2         As Currency
End Type

Private This As TPedreiro
Private Const sTable As String = "tbl_pedreiros"

Public Property Get ID() As Long: ID = This.ID: End Property
Public Property Let ID(ByVal Value As Long): This.ID = Value: End Property

Public Property Get Nome() As String: Nome = This.Nome: End Property
Public Property Let Nome(ByVal Value As String): This.Nome = Value: End Property

Public Property Get Apelido() As String: Apelido = This.Apelido: End Property
Public Property Let Apelido(ByVal Value As String): This.Apelido = Value: End Property

Public Property Get PrecoM2() As Currency: PrecoM2 = This.PrecoM2: End Property
Public Property Let PrecoM2(ByVal Value As Currency): This.PrecoM2 = Value: End Property

Public Sub Crud(Acao As Crud, Optional ID As Long)
    
    If ID > 0 Then
        sSQL = "SELECT * FROM " & sTable & " WHERE id = " & ID
    End If

    ' Cria objeto Recordset
    Set rst = New ADODB.Recordset
    
    If Acao = Read Then
        
        With rst
            .Open sSQL, cnn, adOpenStatic, adLockReadOnly
        
            Me.ID = .Fields("id").Value
            Me.Nome = .Fields("nome").Value
            Me.Apelido = .Fields("apelido").Value
            Me.PrecoM2 = .Fields("preco_m2").Value
            
        End With
        
    ElseIf Acao = Create Or Acao = Update Then
        
        With rst
        
            If Acao = Create Then
                .Open sTable, cnn, adOpenDynamic, adLockOptimistic, adCmdTable
                .AddNew
                
                Me.ID = .Fields("id").Value
                
            ElseIf Acao = Update Then
            
                .Open sSQL, cnn, adOpenDynamic, adLockOptimistic
                
            End If
            
            .Fields("nome").Value = Me.Nome
            .Fields("apelido").Value = Me.Apelido
            .Fields("preco_m2").Value = Me.PrecoM2
            
            .Update
            
        End With
        
    ElseIf Acao = Delete Then
    
        With rst
        
            .Open sSQL, cnn, adOpenDynamic, adLockOptimistic
            
            .Delete
            
        End With
        
    End If
    
    Set rst = Nothing

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
Public Function Existe(Nome As String) As Boolean

    Existe = False
    
    ' Comando SQL
    sSQL = "SELECT * FROM " & sTable & " WHERE nome = '" & Replace(Nome, "'", "`") & "'"

    ' Atribui resultado da consulta SQL ao recordset
    Set rst = cnn.Execute(sSQL)
    
    ' (EOF = True) = Fornecedor n�o existe
    ' (EOF = False) = Fornecedor existe
    If rst.EOF = False Then
        Me.ID = rst.Fields("id").Value
        Existe = True
    End If
    
    ' Destr�i Recordset
    Set rst = Nothing
        
End Function
