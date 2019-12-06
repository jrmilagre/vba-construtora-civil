Attribute VB_Name = "m_Database"
Option Explicit         ' Obriga a declara��o de vari�veis
Option Private Module   ' Deixa o m�dulo privado (invis�vel)

' BIBLIOTECAS:
' ---> Microsoft ActiveX Data Objects 2.8 Library
' ---> Microsoft ADO Ext. 2.8 for DDL and Security

Public cnn  As ADODB.Connection  ' Objeto de conex�o com o banco de dados
Public rst  As ADODB.Recordset   ' Objeto de armazenamento de dados
Public cat  As ADOX.Catalog
Public sSQL As String
Private Const sBanco As String = "database_test.mdb"
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
        If Trim(sText) <> "table;field;type;size;nullable;autoincrement;description" Then
            
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
            
            ' VERIFICA SE EXISTE CAMPO
            Dim oCampo          As ADOX.Column
            Dim bExisteCampo    As Boolean
            
            Set oCampo = New ADOX.Column
            bExisteCampo = False
            
            For Each oCampo In oCatalogo.Tables(myArray(0)).Columns
                
                If oCampo.name = myArray(1) Then
                    bExisteCampo = True
                    Exit For
                End If
                
            Next oCampo
            
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
                
                Set oCampo = Nothing
                
            End If
        
        End If
    
    Loop
    
    Close #1
    
    cnn.Close
    Set oCatalogo = Nothing
    
    MsgBox "Banco de dados atualizado com sucesso!", vbInformation

End Sub
