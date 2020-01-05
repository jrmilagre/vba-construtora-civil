VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fDicionarioDados 
   Caption         =   ":: Dicionário de dados ::"
   ClientHeight    =   5850
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6930
   OleObjectBlob   =   "fDicionarioDados.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fDicionarioDados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()

    If Conecta = True Then
    
        Call cbbTipoCampoPopular
        
        Call lstTabelasPopular
    
    End If

End Sub
Private Sub UserForm_Terminate()

    Call Desconecta

End Sub
Private Sub lstTabelasPopular()
    
    With lstTabelas
    
        .Font = "Consolas"

        For Each tbl In cat.Tables
        
            If tbl.Type = "TABLE" Then
        
                .AddItem tbl.name
            
            End If
            
        Next
        
    End With

End Sub
Private Sub lstTabelas_Change()
    
    Dim sTabela As String
    
    sTabela = lstTabelas.List(lstTabelas.ListIndex, 0)
    
    Call lstCamposPopular(sTabela)
    
End Sub
Private Sub lstCamposPopular(Tabela As String)

    Dim tbl As ADOX.Table
    Dim col As ADOX.Column
    
    Set tbl = cat.Tables(Tabela)
    
    With lstCampos
        .Clear
        .ColumnCount = 2
        .ColumnWidths = "120pt; 55pt;"
        .Font = "Consolas"
    End With
    
    For Each col In tbl.Columns
        
        With lstCampos
        
            .AddItem
            .List(.ListCount - 1, 0) = col.name
        End With
        
    Next

End Sub
Private Sub lstCampos_Change()

    Dim sTabela As String
    Dim sCampo  As String
    Dim n       As Integer
    
    If lstCampos.ListIndex > -1 Then
    
        sTabela = lstTabelas.List(lstTabelas.ListIndex, 0)
        sCampo = lstCampos.List(lstCampos.ListIndex, 0)
        
        Set col = cat.Tables(sTabela).Columns(sCampo)
        
        chbAutoincremento.Value = col.Properties(0).Value
        txbNomeCampo.Text = col.name
        
        For n = 0 To cbbTipoCampo.ListCount - 1
            If CInt(cbbTipoCampo.List(n, 0)) = col.Type Then
                cbbTipoCampo.ListIndex = n
                Exit For
            End If
        Next n
    
    End If

End Sub
Private Sub cbbTipoCampoPopular()

    With cbbTipoCampo
        .Clear
        .ColumnCount = 3
        .ColumnWidths = "0pt; 0pt; 100pt;"
        
        .AddItem
        .List(.ListCount - 1, 0) = "11"
        .List(.ListCount - 1, 1) = "adBoolean"
        .List(.ListCount - 1, 2) = "Sim/Não"
        
        .AddItem
        .List(.ListCount - 1, 0) = "6"
        .List(.ListCount - 1, 1) = "adCurrency"
        .List(.ListCount - 1, 2) = "Moeda"
        
        .AddItem
        .List(.ListCount - 1, 0) = "7"
        .List(.ListCount - 1, 1) = "adDate"
        .List(.ListCount - 1, 2) = "Data/Hora"
        
        .AddItem
        .List(.ListCount - 1, 0) = "202"
        .List(.ListCount - 1, 1) = "adVarWChar"
        .List(.ListCount - 1, 2) = "Texto curto"
        
        .AddItem
        .List(.ListCount - 1, 0) = "3"
        .List(.ListCount - 1, 1) = "adInteger"
        .List(.ListCount - 1, 2) = "Número"
        
    End With

End Sub
