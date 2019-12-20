Attribute VB_Name = "m_Database"
Option Explicit         ' Obriga a declara��o de vari�veis
Option Private Module   ' Deixa o m�dulo privado (invis�vel)

' BIBLIOTECAS:
' ---> Microsoft ActiveX Data Objects 2.8 Library
' ---> Microsoft ADO Ext. 2.8 for DDL and Security

Public cnn  As ADODB.Connection  ' Objeto de conex�o com o banco de dados
Public rst  As ADODB.RecordSet   ' Objeto de armazenamento de dados
Public cat  As ADOX.Catalog
Public sSQL As String
Private Const sBanco As String = "database.mdb"
Private sCaminho As String

' Fun��o para efetuar conex�o com o banco de dados
Public Function Conecta() As Boolean
    
    ' Declara var�avel
    Dim vbResultado As VBA.VbMsgBoxResult
    Dim sCaminho As String
    
    sCaminho = Mid(wbCode.Path, 1, Len(wbCode.Path) - 5) & _
               Application.PathSeparator & "data" & _
               Application.PathSeparator & sBanco
    
    ' Cria objeto de conex�o com o banco de dados
    Set cnn = New ADODB.Connection
    Set cat = New ADOX.Catalog
    
    ' Inicia status da conex�o como falso (desconectado)
    Conecta = False
    
    ' Se a conex�o der erro, desvia para o r�tulo Sair
    On Error GoTo Sair
    
    ' Com o objeto conex�o, escolhe o provedor e abre o banco de dados
    With cnn
        .Provider = "Microsoft.Jet.OLEDB.4.0"       ' Provedor
        .Open sCaminho
        Set cat.ActiveConnection = cnn
    End With
    
    ' Se a conex�o estiver funcionando, retorna verdadeiro
    Conecta = True
    
    ' Sai da fun��o
    Exit Function

' R�tulo Sair
Sair:
    ' Mensagem caso a conex�o com o banco de dados der problema
    vbResultado = MsgBox("Banco de dados n�o existe ou n�o est� acess�vel:" & vbNewLine & _
           vbNewLine & "Caminho do banco procurado: " & vbNewLine & _
           vbNewLine & sCaminho & vbNewLine & vbNewLine & _
           "Deseja criar o arquivo de banco de dados?", vbInformation + vbYesNo)
    
    If vbResultado = vbYes Then
        Call CriaBancoDeDados(sCaminho)
    Else
        Exit Function
    End If

End Function

' Fun��o para efetuar a desconex�o com o banco de dados
' --- � necess�rio habilitar a biblioteca "Microsoft ActiveX Data Objects 2.8 Library"
' --- para o funcionamento desta fun��o.
Public Sub Desconecta()

    ' Fecha conex�o com o banco de dados
    cnn.Close
    Set cat = Nothing

End Sub
' Procedimento para criar o banco de dados
' --- � necess�rio habilitar a biblioteca "Microsoft ADO Ext. 2.8 for DDL and Security"
' --- para o funcionamento deste procedimento.
Private Sub CriaBancoDeDados(Caminho As String)
     
    ' Declara vari�vel
    Dim oCatalogo As New ADOX.Catalog
     
    ' Cria o banco de dados
    oCatalogo.Create "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho
    
    ' Rotina para criar tabelas
    Call AtualizaBD
    
    ' Mensagem de conclus�o
    MsgBox "Banco de dados criado com sucesso!", vbInformation
    
End Sub

Public Sub AtualizaBD()

    ' Declara vari�veis
    Dim oCatalogo       As New ADOX.Catalog
    Dim sCaminho        As String
    
    sCaminho = Mid(wbCode.Path, 1, Len(wbCode.Path) - 5) & _
               Application.PathSeparator & "data" & _
               Application.PathSeparator & sBanco
    
    ' Cria o banco de dados se n�o existir
    On Error GoTo Conecta
    oCatalogo.Create "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sCaminho

Conecta:
    Set cnn = New ADODB.Connection
    
    ' Abre cat�logo
    With cnn
        .Provider = "Microsoft.Jet.OLEDB.4.0"       ' Provedor
        .Open sCaminho
        Set oCatalogo.ActiveConnection = cnn        ' Instancia o cat�logo
    End With
    
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '

    Dim FilePath As String
    Dim sText As String
    Dim myArray() As String
    Dim sTableName As String
    
    FilePath = Mid(wbCode.Path, 1, Len(wbCode.Path) - 5) & _
               Application.PathSeparator & "data" & _
               Application.PathSeparator & "date_dictionary.csv"
    
    Open FilePath For Input As #1
    
    ' La�o para percorrer o arquivo csv que cont�m o dicion�rio de dados
    Do Until EOF(1)
    
        Line Input #1, sText
        
        ' Ignora o cabe�alho
        If Mid(sText, 1, 5) <> "table" Then
            
            myArray = Split(sText, ";")
                        
            ' VERIFICA SE EXISTE TABELA
            If sTableName <> myArray(0) Then
            
                Dim oTabela         As New ADOX.Table
                Dim bExisteTabela   As Boolean
                
                bExisteTabela = False
                
                For Each oTabela In oCatalogo.Tables
                    If oTabela.Type = "TABLE" Then
                        If oTabela.name = myArray(0) Then
                            bExisteTabela = True
                            Exit For
                        End If
                    End If
                Next oTabela
            Else
                bExisteTabela = True
            End If
            
            sTableName = myArray(0)
            
            ' Se tabela n�o existir, cria tabela no banco de dados
            If bExisteTabela = False Then
        
                With oTabela
                    .name = myArray(0)
                    Set .ParentCatalog = oCatalogo
                End With
            
                oCatalogo.Tables.Append oTabela
            End If
            
            '--- VERIFICA SE EXISTE CAMPO ---------------------------+
            Dim oCampo          As ADOX.Column                      '|
            Dim bExisteCampo    As Boolean                          '|
                                                                    '|
            Set oCampo = New ADOX.Column                            '|
            bExisteCampo = False                                    '|
                                                                    '|
            For Each oCampo In oCatalogo.Tables(myArray(0)).Columns '|
                                                                    '|
                If oCampo.name = myArray(1) Then                    '|
                    bExisteCampo = True                             '|
                    Exit For                                        '|
                End If                                              '|
                                                                    '|
            Next oCampo                                             '|
            '--------------------------------------------------------+
            
            Set oCampo = Nothing
            
            ' Cria o campo na tabela, caso n�o exista
            If bExisteCampo = False Then
            
                Set oCampo = New ADOX.Column
                
                With oCampo
                    Set .ParentCatalog = oCatalogo
                    .name = myArray(1)
                    .Type = CInt(myArray(2))
                    
                    If CInt(myArray(2)) = 202 Then
                        .DefinedSize = CInt(myArray(3))
                    End If
                    
                    If CInt(myArray(3)) <> 13 Then
                        .Properties("Nullable").Value = CBool(myArray(4))
                        .Properties("Autoincrement").Value = CBool(myArray(5))
                        .Properties("Description").Value = CStr(myArray(6))
                    End If
                    
                End With
                
                oCatalogo.Tables(myArray(0)).Columns.Append oCampo
                
                ' Cria chave prim�ria
                If CBool(myArray(7)) = True Then
                        
                    Dim idx As ADOX.Index
                    
                    Set idx = New ADOX.Index

                    With idx
                        .name = "PK_" & myArray(0)
                        .IndexNulls = adIndexNullsAllow
                        .PrimaryKey = True
                        .Unique = True
                        .Columns.Append myArray(1)
                    End With
                    
                    oCatalogo.Tables(myArray(0)).Indexes.Append idx
                    
                    Set idx = Nothing

                End If
                
                ' Cria chave estrangeira
                If myArray(8) <> "False" Then

                    Dim fk As ADOX.Key
                    
                    Set fk = New ADOX.Key
                    
                    Dim fkArr() As String

                    fkArr = Split(myArray(8), ".")

                    With fk
                       .name = "FK_" & fkArr(0) & "->" & fkArr(1) & "=" & myArray(0) & "->" & myArray(1)
                       .Type = adKeyForeign
                       .RelatedTable = fkArr(0)
                       .Columns.Append myArray(1)
                       .Columns(myArray(1)).RelatedColumn = fkArr(1)
                       .UpdateRule = adRICascade
                    End With

                    oCatalogo.Tables(myArray(0)).Keys.Append fk
                    
                    Set fk = Nothing
                    
                End If
                
                Set oCampo = Nothing
                
            End If
        
        End If
    
    Loop
    
    Close #1
    
    Set oCatalogo = Nothing
    
    Call Desconecta
    
    MsgBox "Banco de dados atualizado com sucesso!", vbInformation

End Sub
Public Function Backup(CaminhoBackup As String) As Boolean

    On Error GoTo Sair
    
    Dim FSO As Object
    Dim NewName As String
    Set FSO = CreateObject("scripting.filesystemobject")
    
    sCaminho = Mid(wbCode.Path, 1, Len(wbCode.Path) - 5) & _
               Application.PathSeparator & "data" & _
               Application.PathSeparator & sBanco
    
    FSO.Copyfile sCaminho, CaminhoBackup
    
    MsgBox "Backup realizado com sucesso!", vbInformation
    
    Backup = True
    Exit Function
Sair:
    Backup = False
    MsgBox "Problema no Backup!", vbCritical
    Exit Function
End Function
Public Sub IncluiRegistrosTeste()

    If Conecta = True Then
        sSQL = "INSERT INTO tbl_produtos ([nome]) VALUES ('Cimento') ": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_produtos ([nome]) VALUES ('Cal') ": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_fornecedores ([nome]) VALUES ('Cardoso') ": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_fornecedores ([nome]) VALUES ('Orlando') ": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_tipos_obra ([nome]) VALUES ('Casa') ": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_tipos_obra ([nome]) VALUES ('Sobrado') ": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_pedreiros ([nome], [apelido], [preco_m2]) VALUES ('Aparecido', 'Cidinho', 300.00) ": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_estados ([nome], [uf]) VALUES ('Minas Gerais', 'MG') ": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_estados ([nome], [uf]) VALUES ('S�o Paulo', 'SP') ": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_clientes ([nome]) VALUES ('Jairo Milagre') ": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_clientes ([nome]) VALUES ('Dalila Emanuela') ": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_obras ([endereco], [tipo_id], [bairro], [cidade], [uf], [cliente_id]) VALUES ('Rua dos Gavi�es, 156', 2, 'Parque dos P�ssaros', 'Extrema', 'MG', 1) ": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_obras ([endereco], [tipo_id], [bairro], [cidade], [uf], [cliente_id]) VALUES ('Alameda Joaquim Marcondes da Silveira, 171', 1, 'Campos Olivotti', 'Extrema', 'MG', 2) ": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_compras ([data], [fornecedor_id]) VALUES (" & CLng(CDate("13/12/2019")) & ", 1)": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_compras ([data], [fornecedor_id]) VALUES (" & CLng(CDate("14/12/2019")) & ", 2)": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_compras_itens ([compra_id], [produto_id], [quantidade], [unitario], [total], [data], [fornecedor_id]) VALUES (1, 1, 2, 23.5, 47, " & CLng(CDate("13/12/2019")) & ", 1)": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_compras_itens ([compra_id], [produto_id], [quantidade], [unitario], [total], [data], [fornecedor_id]) VALUES (1, 2, 5, 5, 25, " & CLng(CDate("13/12/2019")) & ", 1)": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_compras_itens ([compra_id], [produto_id], [quantidade], [unitario], [total], [data], [fornecedor_id]) VALUES (2, 1, 5, 24.5, 122.5, " & CLng(CDate("14/12/2019")) & ", 2)": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_compras_itens ([compra_id], [produto_id], [quantidade], [unitario], [total], [data], [fornecedor_id]) VALUES (2, 2, 5, 4.8, 48, " & CLng(CDate("14/12/2019")) & ", 2)": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_titulos_pagar ([compra_id], [fornecedor_id], [observacao], [vencimento], [valor], [data]) VALUES (1, 1, '01/02', " & CLng(CDate("13/01/2019")) & ", 36, " & CLng(CDate("13/12/2019")) & ")": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_titulos_pagar ([compra_id], [fornecedor_id], [observacao], [vencimento], [valor], [data]) VALUES (1, 1, '02/02', " & CLng(CDate("13/02/2020")) & ", 36, " & CLng(CDate("13/12/2019")) & ")": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_titulos_pagar ([compra_id], [fornecedor_id], [observacao], [vencimento], [valor], [data]) VALUES (2, 2, '01/03', " & CLng(CDate("14/01/2020")) & ", 48.83, " & CLng(CDate("14/12/2019")) & ")": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_titulos_pagar ([compra_id], [fornecedor_id], [observacao], [vencimento], [valor], [data]) VALUES (2, 2, '02/03', " & CLng(CDate("14/02/2020")) & ", 48.83, " & CLng(CDate("14/12/2019")) & ")": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_titulos_pagar ([compra_id], [fornecedor_id], [observacao], [vencimento], [valor], [data]) VALUES (2, 2, '03/03', " & CLng(CDate("14/03/2020")) & ", 48.84, " & CLng(CDate("14/12/2019")) & ")": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_contas ([nome], [saldo_inicial]) VALUES ('Dinheiro em caixa', 0) ": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_etapas ([nome]) VALUES ('Alvenaria') ": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_etapas ([nome]) VALUES ('Acabamento') ": cnn.Execute sSQL
        
        Call Desconecta
    End If
    
    

End Sub
