VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTituloReceber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type TTituloReceber
    ObraID        As Long
    ClienteID    As Long
    CategoriaID     As Long
    Observacao      As String
    Vencimento      As Date
    Valor           As Currency
    Data            As Date
    Recno           As Long
End Type

Private This As TTituloReceber

Private Const sTable As String = "tbl_titulos_receber"

Public Property Get ObraID() As Long: ObraID = This.ObraID: End Property
Public Property Let ObraID(ByVal Value As Long): This.ObraID = Value: End Property

Public Property Get ClienteID() As Long: ClienteID = This.ClienteID: End Property
Public Property Let ClienteID(ByVal Value As Long): This.ClienteID = Value: End Property

Public Property Get CategoriaID() As Long: CategoriaID = This.CategoriaID: End Property
Public Property Let CategoriaID(ByVal Value As Long): This.CategoriaID = Value: End Property

Public Property Get Observacao() As String: Observacao = This.Observacao: End Property
Public Property Let Observacao(ByVal Value As String): This.Observacao = Value: End Property

Public Property Get Vencimento() As Date: Vencimento = This.Vencimento: End Property
Public Property Let Vencimento(ByVal Value As Date): This.Vencimento = Value: End Property

Public Property Get Valor() As Currency: Valor = This.Valor: End Property
Public Property Let Valor(ByVal Value As Currency): This.Valor = Value: End Property

Public Property Get Data() As Date: Data = This.Data: End Property
Public Property Let Data(ByVal Value As Date): This.Data = Value: End Property

Public Property Get Recno() As Long: Recno = This.Recno: End Property
Public Property Let Recno(ByVal Value As Long): This.Recno = Value: End Property

Public Sub Carrega(ID As Long)

    sSQL = "SELECT * FROM " & sTable & " WHERE r_e_c_n_o_ = " & ID
    
    Set rst = New ADODB.Recordset
    
    rst.Open sSQL, cnn, adOpenStatic
    
    With rst
        Me.ObraID = .Fields("obra_id").Value
        Me.ClienteID = .Fields("cliente_id").Value
        Me.Observacao = .Fields("observacao").Value
        Me.Vencimento = .Fields("vencimento").Value
        Me.Valor = .Fields("valor").Value
        Me.Data = .Fields("data").Value
        Me.Recno = .Fields("r_e_c_n_o_").Value
    End With
    
End Sub

Public Sub Inclui()

    Set rst = New ADODB.Recordset

    With rst
        .Open Source:=sTable, _
              ActiveConnection:=cnn, _
              CursorType:=adOpenDynamic, _
              LockType:=adLockOptimistic, _
              Options:=adCmdTable
        .AddNew
        
        Me.Recno = .Fields("r_e_c_n_o_").Value
        
        .Fields("obra_id").Value = Me.ObraID
        .Fields("cliente_id").Value = Me.ClienteID
        .Fields("observacao").Value = Me.Observacao
        .Fields("vencimento").Value = CLng(Me.Vencimento)
        .Fields("valor").Value = Me.Valor
        .Fields("data").Value = CLng(Me.Data)
        
        .Update
    End With
    
    Set rst = Nothing
       
End Sub
Public Sub AlteraTitulo(TituloID As Long)
    
    Set rst = New ADODB.Recordset
    
    ' Comando SQL
    sSQL = "SELECT * FROM " & sTable & " WHERE r_e_c_n_o_ = " & TituloID

    'Armazena na variável o comando que fará a consulta SQL no BD Access
    With rst
        .Open Source:=sSQL, _
              ActiveConnection:=cnn, _
              CursorType:=adOpenDynamic, _
              LockType:=adLockOptimistic
        
        ' Atribui novos valores aos campos
        .Fields("obra_id").Value = Me.ObraID
        .Fields("cliente_id").Value = Me.ClienteID
        .Fields("observacao").Value = Me.Observacao
        .Fields("vencimento").Value = CLng(Me.Vencimento)
        .Fields("valor").Value = Me.Valor
        .Fields("data").Value = CLng(Me.Data)
        
        .Update
        
        ' Grava alterações
        .Update
    End With
    
    Set rst = Nothing
End Sub
Public Sub Exclui(ObraID As Long)
    
    sSQL = "DELETE FROM tbl_titulos_receber WHERE obra_id = " & ObraID
    
    cnn.Execute sSQL

End Sub
Public Sub ExcluiTitulo(TituloID As Long)
    
    sSQL = "DELETE FROM tbl_titulos_receber WHERE r_e_c_n_o_ = " & TituloID

    cnn.Execute sSQL

End Sub
Public Function GetValorBaixado(TituloID As Long) As Currency

    Dim rTmp As New ADODB.Recordset

    sSQL = "SELECT sum(valor_baixado) as valor_baixado FROM tbl_recebimentos_itens WHERE titulo_id = " & TituloID
    
    rTmp.Open sSQL, cnn
    
    If Not IsNull(rTmp.Fields("valor_baixado").Value) Then
        GetValorBaixado = rTmp.Fields("valor_baixado").Value
    Else
        GetValorBaixado = 0
    End If
    
    Set rTmp = Nothing
    
End Function
Public Function Recordset(Optional ObraID As Long) As ADODB.Recordset

    sSQL = "SELECT * "
    sSQL = sSQL & "FROM " & sTable & " "
    sSQL = sSQL & "WHERE (valor - valor_recebido) > 0 "
    
    If ObraID > 0 Then
        sSQL = sSQL & "and obra_id = " & ObraID & " "
    End If
    
    sSQL = sSQL & "ORDER BY vencimento"
    
    Set Recordset = New ADODB.Recordset
    
    With Recordset
        .CursorLocation = adUseServer
        .CursorType = adOpenStatic
        .LockType = adLockReadOnly
        .PageSize = 24
        .Open sSQL, cnn
    End With
    
End Function
Public Function ExisteRecebimento(ObraID As Long, TituloID As Long) As Boolean

    Dim rTmp    As New ADODB.Recordset
    Dim sMsg    As String
    
    ExisteRecebimento = False
    
    sSQL = "SELECT tbl_recebimentos_itens.recebimento_id, tbl_titulos_receber.r_e_c_n_o_ "
    sSQL = sSQL & "FROM tbl_titulos_receber INNER JOIN tbl_recebimentos_itens ON tbl_titulos_receber.r_e_c_n_o_ = tbl_recebimentos_itens.titulo_id "
    sSQL = sSQL & "GROUP BY tbl_recebimentos_itens.recebimento_id, tbl_titulos_receber.obra_id, tbl_titulos_receber.r_e_c_n_o_ "
    sSQL = sSQL & "HAVING (((tbl_titulos_receber.obra_id)=" & ObraID & ") AND ((tbl_titulos_receber.r_e_c_n_o_)=" & TituloID & "));"
     
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
