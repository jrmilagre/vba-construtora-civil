VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cObra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type TTipoObra
    ID              As Long
    Endereco        As String
    Bairro          As String
    Cidade          As String
    UF              As String
    TipoID          As Integer
    ClienteID       As Variant
    Data            As Date
    CategoriaID     As Long
    Encerrada       As Boolean
End Type

Private This As TTipoObra

Private Const sTable As String = "tbl_obras"

Public Property Get ID() As Long: ID = This.ID: End Property
Public Property Let ID(ByVal Value As Long): This.ID = Value: End Property

Public Property Get Endereco() As String: Endereco = This.Endereco: End Property
Public Property Let Endereco(ByVal Value As String): This.Endereco = Value: End Property

Public Property Get TipoID() As Integer: TipoID = This.TipoID: End Property
Public Property Let TipoID(ByVal Value As Integer): This.TipoID = Value: End Property

Public Property Get Bairro() As String: Bairro = This.Bairro: End Property
Public Property Let Bairro(ByVal Value As String): This.Bairro = Value: End Property

Public Property Get Cidade() As String: Cidade = This.Cidade: End Property
Public Property Let Cidade(ByVal Value As String): This.Cidade = Value: End Property

Public Property Get UF() As String: UF = This.UF: End Property
Public Property Let UF(ByVal Value As String): This.UF = Value: End Property

Public Property Get ClienteID() As Variant: ClienteID = This.ClienteID: End Property
Public Property Let ClienteID(ByVal Value As Variant): This.ClienteID = Value: End Property

Public Property Get Data() As Date: Data = This.Data: End Property
Public Property Let Data(ByVal Value As Date): This.Data = Value: End Property

Public Property Get CategoriaID() As Variant: CategoriaID = This.CategoriaID: End Property
Public Property Let CategoriaID(ByVal Value As Variant): This.CategoriaID = Value: End Property

Public Property Get Encerrada() As Boolean: Encerrada = This.Encerrada: End Property
Public Property Let Encerrada(ByVal Value As Boolean): This.Encerrada = Value: End Property

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
        Me.Endereco = .Fields("endereco").Value
        Me.Bairro = .Fields("bairro").Value
        Me.Cidade = .Fields("cidade").Value
        Me.UF = .Fields("uf").Value
        Me.TipoID = .Fields("tipo_id").Value
        Me.ClienteID = IIf(IsNull(.Fields("cliente_id").Value), Null, .Fields("cliente_id").Value)
        Me.Data = .Fields("data").Value
        Me.CategoriaID = .Fields("categoria_id").Value
        Me.Encerrada = .Fields("encerrada").Value
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
        .Fields("bairro").Value = Me.Bairro
        .Fields("cidade").Value = Me.Cidade
        .Fields("uf").Value = Me.UF
        .Fields("tipo_id").Value = Me.TipoID
        .Fields("cliente_id").Value = Me.ClienteID
        .Fields("data").Value = Me.Data
        .Fields("categoria_id").Value = Me.CategoriaID
        .Fields("encerrada").Value = Me.Encerrada
        
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
        .Fields("bairro").Value = Me.Bairro
        .Fields("cidade").Value = Me.Cidade
        .Fields("uf").Value = Me.UF
        .Fields("tipo_id").Value = Me.TipoID
        .Fields("cliente_id").Value = Me.ClienteID
        .Fields("data").Value = Me.Data
        .Fields("categoria_id").Value = Me.CategoriaID
        .Fields("encerrada").Value = Me.Encerrada
        
        ' Grava alterações
        .Update
    End With
    
    Set rst = Nothing
End Sub

' Rotina para Excluir registro do banco de dados
Public Sub Exclui(ID As Long)
    
    sSQL = "DELETE FROM " & sTable & " WHERE id = " & ID
    
    cnn.Execute sSQL
    
End Sub
Public Function Listar(OrderBy As String) As Collection

    Dim col As New Collection

    sSQL = "SELECT id "
    sSQL = sSQL & "FROM " & sTable & " "
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
Public Function Recordset(Status As Variant) As ADODB.Recordset

    sSQL = "SELECT * "
    sSQL = sSQL & "FROM " & sTable & " "
    
    If Not IsNull(Status) Then
        sSQL = sSQL & "WHERE encerrada = " & Status
    End If
    
    Set Recordset = New ADODB.Recordset
    
    With Recordset
        .CursorLocation = adUseServer
        .CursorType = adOpenStatic
        .LockType = adLockReadOnly
        .PageSize = 24
        .Open sSQL, cnn
    End With
    
End Function
Public Function ExisteRecebimento(ObraID As Long) As Boolean

    Dim rTmp    As New ADODB.Recordset
    Dim sMsg    As String
    
    ExisteRecebimento = False
    
    sSQL = "SELECT tbl_recebimentos_itens.recebimento_id "
    sSQL = sSQL & "FROM tbl_titulos_receber "
    sSQL = sSQL & "INNER JOIN tbl_recebimentos_itens ON tbl_titulos_receber.r_e_c_n_o_ = tbl_recebimentos_itens.titulo_id "
    sSQL = sSQL & "GROUP BY tbl_recebimentos_itens.recebimento_id, tbl_titulos_receber.obra_id "
    sSQL = sSQL & "HAVING (((tbl_titulos_receber.obra_id)=" & ObraID & "));"

    
    rTmp.Open sSQL, cnn, adOpenStatic
    
    If rTmp.EOF = False Then
    
        sMsg = "Não posso excluir pois existem recebimentos vinculados a essa obra!"
        sMsg = sMsg & vbNewLine
        sMsg = sMsg & vbNewLine
        sMsg = sMsg & "Veja o número dos recebimentos:"
        sMsg = sMsg & vbNewLine
        sMsg = sMsg & vbNewLine
        
        Do Until rTmp.EOF
        
            sMsg = sMsg & "Nº:" & Format(rTmp.Fields("recebimento_id").Value, "0000000000") & vbNewLine
        
            rTmp.MoveNext
        Loop
        
        sMsg = sMsg & vbNewLine
        sMsg = sMsg & "Fim!"
        
        MsgBox sMsg, vbCritical, "EXISTEM RECEBIMENTOS"
            
        ExisteRecebimento = True
    Else
        ExisteRecebimento = False
    End If
        
    Set rTmp = Nothing

End Function
