Attribute VB_Name = "m_Database"
Option Explicit         ' Obriga a declaração de variáveis
Option Private Module   ' Deixa o módulo privado (invisível)

' BIBLIOTECAS:
' ---> Microsoft ActiveX Data Objects 2.8 Library
' ---> Microsoft ADO Ext. 2.8 for DDL and Security

Public cnn  As ADODB.Connection  ' Objeto de conexão com o banco de dados
Public rst  As ADODB.Recordset   ' Objeto de armazenamento de dados
Public cat  As ADOX.Catalog
Public sSQL As String
Private Const sBanco As String = "database.mdb"
Private sCaminho As String

' Função para efetuar conexão com o banco de dados
Public Function Conecta() As Boolean
    
    ' Declara varíavel
    Dim vbResultado As VBA.VbMsgBoxResult
    Dim sCaminho As String
    
    sCaminho = Mid(wbCode.Path, 1, Len(wbCode.Path) - 5) & _
               Application.PathSeparator & "data" & _
               Application.PathSeparator & sBanco
    
    ' Cria objeto de conexão com o banco de dados
    Set cnn = New ADODB.Connection
    Set cat = New ADOX.Catalog
    
    ' Inicia status da conexão como falso (desconectado)
    Conecta = False
    
    ' Se a conexão der erro, desvia para o rótulo Sair
    On Error GoTo Sair
    
    ' Com o objeto conexão, escolhe o provedor e abre o banco de dados
    With cnn
        .Provider = "Microsoft.Jet.OLEDB.4.0"       ' Provedor
        .Open sCaminho
        Set cat.ActiveConnection = cnn
    End With
    
    ' Se a conexão estiver funcionando, retorna verdadeiro
    Conecta = True
    
    ' Sai da função
    Exit Function

' Rótulo Sair
Sair:
    ' Mensagem caso a conexão com o banco de dados der problema
    vbResultado = MsgBox("Banco de dados não existe ou não está acessível:" & vbNewLine & _
           vbNewLine & "Caminho do banco procurado: " & vbNewLine & _
           vbNewLine & sCaminho & vbNewLine & vbNewLine & _
           "Deseja criar o arquivo de banco de dados?", vbInformation + vbYesNo)
    
    If vbResultado = vbYes Then
        Call CriaBancoDeDados(sCaminho)
    Else
        Exit Function
    End If

End Function

' Função para efetuar a desconexão com o banco de dados
' --- É necessário habilitar a biblioteca "Microsoft ActiveX Data Objects 2.8 Library"
' --- para o funcionamento desta função.
Public Sub Desconecta()

    ' Fecha conexão com o banco de dados
    cnn.Close
    Set cat = Nothing

End Sub
' Procedimento para criar o banco de dados
' --- É necessário habilitar a biblioteca "Microsoft ADO Ext. 2.8 for DDL and Security"
' --- para o funcionamento deste procedimento.
Private Sub CriaBancoDeDados(Caminho As String)
     
    ' Declara variável
    Dim oCatalogo As New ADOX.Catalog
     
    ' Cria o banco de dados
    oCatalogo.Create "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho
    
    ' Rotina para criar tabelas
    Call AtualizaBD
    
    ' Mensagem de conclusão
    MsgBox "Banco de dados criado com sucesso!", vbInformation
    
End Sub

Public Sub AtualizaBD()

    ' Declara variáveis
    Dim oCatalogo       As New ADOX.Catalog
    Dim sCaminho        As String
    
    sCaminho = Mid(wbCode.Path, 1, Len(wbCode.Path) - 5) & _
               Application.PathSeparator & "data" & _
               Application.PathSeparator & sBanco
    
    ' Cria o banco de dados se não existir
    On Error GoTo Conecta
    oCatalogo.Create "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sCaminho

Conecta:
    Set cnn = New ADODB.Connection
    
    ' Abre catálogo
    With cnn
        .Provider = "Microsoft.Jet.OLEDB.4.0"       ' Provedor
        .Open sCaminho
        Set oCatalogo.ActiveConnection = cnn        ' Instancia o catálogo
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
    
    ' Laço para percorrer o arquivo csv que contém o dicionário de dados
    Do Until EOF(1)
    
        Line Input #1, sText
        
        ' Ignora o cabeçalho
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
            
            ' Se tabela não existir, cria tabela no banco de dados
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
            
            ' Cria o campo na tabela, caso não exista
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
                
                ' Cria chave primária
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
        sSQL = "INSERT INTO tbl_unidades_medida ([nome], [abreviacao]) VALUES ('Saco', 'SC') ": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_unidades_medida ([nome], [abreviacao]) VALUES ('Metro cúbico', 'M3') ": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_unidades_medida ([nome], [abreviacao]) VALUES ('Metro quadrado', 'M2') ": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_unidades_medida ([nome], [abreviacao]) VALUES ('Lata', 'LT') ": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_unidades_medida ([nome], [abreviacao]) VALUES ('Metro linear', 'MT') ": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_unidades_medida ([nome], [abreviacao]) VALUES ('Dia', 'DD') ": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_unidades_medida ([nome], [abreviacao]) VALUES ('Hora', 'HH') ": cnn.Execute sSQL
        
        sSQL = "INSERT INTO tbl_produtos ([nome], [um_id]) VALUES ('Cimento', 1) ": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_produtos ([nome], [um_id]) VALUES ('Cal', 1) ": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_produtos ([nome], [um_id]) VALUES ('Areia média', 2) ": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_produtos ([nome], [um_id]) VALUES ('Pedra', 2) ": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_produtos ([nome], [um_id]) VALUES ('Laje treliça H08', 3) ": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_produtos ([nome], [um_id]) VALUES ('Mão de obra', 3) ": cnn.Execute sSQL
        
        sSQL = "INSERT INTO tbl_categorias ([pag_rec], [categoria], [subcategoria]) VALUES ('R', 'Vendas', 'Obras') ": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_categorias ([pag_rec], [categoria], [subcategoria]) VALUES ('R', 'Vendas', 'Carros') ": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_categorias ([pag_rec], [categoria], [subcategoria]) VALUES ('P', 'Despesas com obras', 'Materiais de construção') ": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_categorias ([pag_rec], [categoria], [subcategoria]) VALUES ('P', 'Despesas com obras', 'Serviços') ": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_categorias ([pag_rec], [categoria], [subcategoria]) VALUES ('P', 'Despesas administrativas', 'Salários') ": cnn.Execute sSQL
        
        sSQL = "INSERT INTO tbl_fornecedores ([nome]) VALUES ('Cardoso Materiais para Construção') ": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_fornecedores ([nome]) VALUES ('Orlando Materiais para Construção') ": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_fornecedores ([nome]) VALUES ('Aparecido (Cidinho)') ": cnn.Execute sSQL
        
        sSQL = "INSERT INTO tbl_tipos_obra ([nome]) VALUES ('Casa') ": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_tipos_obra ([nome]) VALUES ('Sobrado') ": cnn.Execute sSQL
        
        sSQL = "INSERT INTO tbl_pedreiros ([nome], [apelido], [preco_m2]) VALUES ('Aparecido', 'Cidinho', 300.00) ": cnn.Execute sSQL
        
        sSQL = "INSERT INTO tbl_estados ([nome], [uf]) VALUES ('Minas Gerais', 'MG') ": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_estados ([nome], [uf]) VALUES ('São Paulo', 'SP') ": cnn.Execute sSQL
        
        sSQL = "INSERT INTO tbl_clientes ([nome]) VALUES ('Acmo Administração de Bens e Participações Eireli') ": cnn.Execute sSQL
        
        sSQL = "INSERT INTO tbl_obras ([endereco], [tipo_id], [bairro], [cidade], [uf], [cliente_id], [data], [categoria_id]) VALUES ('Alameda Joaquim Marcondes da Silveira, 171', 2, 'Campos Olivotti', 'Extrema', 'MG', 1, " & CLng(CDate("14/12/2019")) & ", 1) ": cnn.Execute sSQL
        
        sSQL = "INSERT INTO tbl_titulos_receber ([obra_id], [cliente_id], [observacao], [vencimento], [valor], [data]) VALUES (1, 1, 'Entrada', " & CLng(CDate("14/12/2019")) & ", 104386.08, " & CLng(CDate("14/12/2019")) & ")": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_titulos_receber ([obra_id], [cliente_id], [observacao], [vencimento], [valor], [data]) VALUES (1, 1, '01/09', " & CLng(CDate("14/01/2020")) & ", 27063.05, " & CLng(CDate("14/12/2019")) & ")": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_titulos_receber ([obra_id], [cliente_id], [observacao], [vencimento], [valor], [data]) VALUES (1, 1, '02/09', " & CLng(CDate("14/02/2020")) & ", 27063.05, " & CLng(CDate("14/12/2019")) & ")": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_titulos_receber ([obra_id], [cliente_id], [observacao], [vencimento], [valor], [data]) VALUES (1, 1, '03/09', " & CLng(CDate("14/03/2020")) & ", 27063.05, " & CLng(CDate("14/12/2019")) & ")": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_titulos_receber ([obra_id], [cliente_id], [observacao], [vencimento], [valor], [data]) VALUES (1, 1, '04/09', " & CLng(CDate("14/04/2020")) & ", 27063.05, " & CLng(CDate("14/12/2019")) & ")": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_titulos_receber ([obra_id], [cliente_id], [observacao], [vencimento], [valor], [data]) VALUES (1, 1, '05/09', " & CLng(CDate("14/05/2020")) & ", 27063.05, " & CLng(CDate("14/12/2019")) & ")": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_titulos_receber ([obra_id], [cliente_id], [observacao], [vencimento], [valor], [data]) VALUES (1, 1, '06/09', " & CLng(CDate("14/06/2020")) & ", 27063.05, " & CLng(CDate("14/12/2019")) & ")": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_titulos_receber ([obra_id], [cliente_id], [observacao], [vencimento], [valor], [data]) VALUES (1, 1, '07/09', " & CLng(CDate("14/07/2020")) & ", 27063.05, " & CLng(CDate("14/12/2019")) & ")": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_titulos_receber ([obra_id], [cliente_id], [observacao], [vencimento], [valor], [data]) VALUES (1, 1, '08/09', " & CLng(CDate("14/08/2020")) & ", 27063.05, " & CLng(CDate("14/12/2019")) & ")": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_titulos_receber ([obra_id], [cliente_id], [observacao], [vencimento], [valor], [data]) VALUES (1, 1, '09/09', " & CLng(CDate("14/09/2020")) & ", 27063.05, " & CLng(CDate("14/12/2019")) & ")": cnn.Execute sSQL
        
        sSQL = "INSERT INTO tbl_compras ([data], [fornecedor_id], [categoria_id]) VALUES (" & CLng(CDate("13/12/2019")) & ", 1, 3)": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_compras ([data], [fornecedor_id], [categoria_id]) VALUES (" & CLng(CDate("14/12/2019")) & ", 2, 3)": cnn.Execute sSQL
        
        sSQL = "INSERT INTO tbl_compras_itens ([compra_id], [produto_id], [quantidade], [um_id], [unitario], [total], [data], [fornecedor_id]) VALUES (1, 1, 2, 1, 23.5, 47, " & CLng(CDate("13/12/2019")) & ", 1)": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_compras_itens ([compra_id], [produto_id], [quantidade], [um_id], [unitario], [total], [data], [fornecedor_id]) VALUES (1, 2, 5, 1, 5, 25, " & CLng(CDate("13/12/2019")) & ", 1)": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_compras_itens ([compra_id], [produto_id], [quantidade], [um_id], [unitario], [total], [data], [fornecedor_id]) VALUES (1, 3, 5, 2, 25, 125, " & CLng(CDate("13/12/2019")) & ", 1)": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_compras_itens ([compra_id], [produto_id], [quantidade], [um_id], [unitario], [total], [data], [fornecedor_id]) VALUES (2, 1, 5, 1, 24.5, 122.5, " & CLng(CDate("14/12/2019")) & ", 2)": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_compras_itens ([compra_id], [produto_id], [quantidade], [um_id], [unitario], [total], [data], [fornecedor_id]) VALUES (2, 2, 5, 1, 4.8, 48, " & CLng(CDate("14/12/2019")) & ", 2)": cnn.Execute sSQL
        
        sSQL = "INSERT INTO tbl_titulos_pagar ([compra_id], [fornecedor_id], [observacao], [vencimento], [valor], [data]) VALUES (1, 1, '01/02', " & CLng(CDate("13/01/2019")) & ", 36, " & CLng(CDate("13/12/2019")) & ")": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_titulos_pagar ([compra_id], [fornecedor_id], [observacao], [vencimento], [valor], [data]) VALUES (1, 1, '02/02', " & CLng(CDate("13/02/2020")) & ", 36, " & CLng(CDate("13/12/2019")) & ")": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_titulos_pagar ([compra_id], [fornecedor_id], [observacao], [vencimento], [valor], [data]) VALUES (2, 2, '01/03', " & CLng(CDate("14/01/2020")) & ", 48.83, " & CLng(CDate("14/12/2019")) & ")": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_titulos_pagar ([compra_id], [fornecedor_id], [observacao], [vencimento], [valor], [data]) VALUES (2, 2, '02/03', " & CLng(CDate("14/02/2020")) & ", 48.83, " & CLng(CDate("14/12/2019")) & ")": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_titulos_pagar ([compra_id], [fornecedor_id], [observacao], [vencimento], [valor], [data]) VALUES (2, 2, '03/03', " & CLng(CDate("14/03/2020")) & ", 48.84, " & CLng(CDate("14/12/2019")) & ")": cnn.Execute sSQL
        
        sSQL = "INSERT INTO tbl_contas ([nome], [saldo_inicial]) VALUES ('Dinheiro em caixa', 0) ": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_contas ([nome], [saldo_inicial]) VALUES ('Santander', 0) ": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_contas ([nome], [saldo_inicial]) VALUES ('Bradesco', 0) ": cnn.Execute sSQL
        
        sSQL = "INSERT INTO tbl_etapas ([nome]) VALUES ('Alvenaria') ": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_etapas ([nome]) VALUES ('Acabamento') ": cnn.Execute sSQL
        
        Call Desconecta
    End If
    
    

End Sub
