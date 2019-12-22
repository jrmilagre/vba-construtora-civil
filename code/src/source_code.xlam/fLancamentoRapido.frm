VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fLancamentoRapido 
   Caption         =   ":: Lancamentos rápidos ::"
   ClientHeight    =   8430
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12885
   OleObjectBlob   =   "fLancamentoRapido.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fLancamentoRapido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private oLancamentoRapido   As New cLancamentoRapido
Private oConta              As New cConta
Private oFornecedor         As New cFornecedor
Private oObra               As New cObra
Private oEtapa              As New cEtapa
Private oCliente            As New cCliente
Private oProduto            As New cProduto
Private oUM                 As New cUnidadeMedida

Private colControles        As New Collection
Private myRst               As ADODB.RecordSet
Private lPagina             As Long

Private Const sTable As String = "tbl_compras"
Private Const sCampoOrderBy As String = "data"

Private Sub UserForm_Initialize()
    
    Call cbbContaPopular
    Call cbbPagRecPopular
    Call cbbFornecedorPopular
    Call cbbObraPopular
    Call cbbEtapaPopular
    Call cbbProdutoPopular
    Call cbbUMPopular
    
    Call EventosCampos
    
    Set myRst = New ADODB.RecordSet
    Set myRst = oCompra.RecordSet

End Sub
Private Sub UserForm_Terminate()
    
    Set oConta = Nothing
    Call Desconecta
    
End Sub
Private Sub cbbContaPopular()
    
    Dim idx         As Integer
    Dim col         As New Collection
    Dim n           As Variant

    Set col = oConta.Listar("nome")
    
    idx = cbbConta.ListIndex
    
    cbbConta.Clear
    
    For Each n In col
        
        oConta.Carrega CLng(n)
    
        With cbbConta
            .AddItem
            .List(.ListCount - 1, 0) = oConta.Nome
            .List(.ListCount - 1, 1) = oConta.ID
        End With
        
    Next n
    
    cbbConta.ListIndex = idx

End Sub
Private Sub cbbPagRecPopular()

    With cbbPagRec
    
        .Clear
        
        .AddItem
        .List(.ListCount - 1, 0) = "Pagamento"
        .List(.ListCount - 1, 1) = "P"
        
        .AddItem
        .List(.ListCount - 1, 0) = "Recebimento"
        .List(.ListCount - 1, 1) = "R"
    End With

End Sub
Private Sub cbbFornecedorPopular()
    
    Dim idx         As Integer
    Dim col         As New Collection
    Dim n           As Variant

    Set col = oFornecedor.Listar("nome")
    
    idx = cbbFornecedor.ListIndex
    
    cbbFornecedor.Clear
    
    For Each n In col
        
        oFornecedor.Carrega CLng(n)
    
        With cbbFornecedor
            .AddItem
            .List(.ListCount - 1, 0) = oFornecedor.Nome
            .List(.ListCount - 1, 1) = oFornecedor.ID
        End With
        
    Next n
    
    cbbFornecedor.ListIndex = idx

End Sub
Private Sub EventosCampos()

    ' Declara variáveis
    Dim oControle   As MSForms.control
    Dim oEvento     As c_EventoCampo
    Dim sTag        As String
    Dim iType       As Integer
    Dim bNullable   As Boolean
    Dim sField()    As String

    ' Laço para percorrer todos os TextBox e atribuir eventos
    ' de acordo com o tipo de cada campo
    For Each oControle In Me.Controls

        If Len(oControle.Tag) > 0 Then

            If TypeName(oControle) = "TextBox" Then

                Set oEvento = New c_EventoCampo

                With oEvento
                
                    sField() = Split(oControle.Tag, ".")

                    oControle.ControlTipText = cat.Tables(sField(0)).Columns(sField(1)).Properties("Description").Value

                    .FieldType = cat.Tables(sField(0)).Columns(sField(1)).Type
                    .MaxLength = cat.Tables(sField(0)).Columns(sField(1)).DefinedSize
                    .Nullable = cat.Tables(sField(0)).Columns(sField(1)).Properties("Nullable")

                    Set .cGeneric = oControle

                End With

                colControles.Add oEvento

            End If
        End If
    Next

End Sub
Private Sub chbRequisita_Click()
    If chbRequisita.Value = False Then
        MultiPage1.Pages(2).Visible = False
    Else
        MultiPage1.Pages(2).Visible = True
        MultiPage1.Value = 2
    End If
End Sub
Private Sub btnData_Click()
    dtDate = IIf(txbData.Text = Empty, Date, txbData.Text)
    txbData.Text = GetCalendario
End Sub
Private Sub cbbObraPopular()
    
    Dim idx         As Integer
    Dim col         As New Collection
    Dim n           As Variant

    Set col = oObra.Listar("bairro")
    
    idx = cbbObra.ListIndex
    
    With cbbObra
        .Clear
        .ColumnCount = 4
        .ColumnWidths = "100pt; 0pt; 100pt; 200pt;"
    End With
    
    
    For Each n In col
        
        oObra.Carrega CLng(n)
        
        oCliente.Carrega oObra.ClienteID
    
        With cbbObra
            .AddItem
            .List(.ListCount - 1, 0) = oObra.Bairro
            .List(.ListCount - 1, 1) = oObra.ID
            .List(.ListCount - 1, 2) = oCliente.Nome
            .List(.ListCount - 1, 3) = oObra.Endereco
        End With
        
    Next n
    
    cbbObra.ListIndex = idx

End Sub
Private Sub cbbEtapaPopular()
    
    Dim idx         As Integer
    Dim col         As New Collection
    Dim n           As Variant

    Set col = oEtapa.Listar("nome")
    
    idx = cbbEtapa.ListIndex
    
    With cbbEtapa
        .Clear
        .ColumnCount = 2
        .ColumnWidths = "60pt; 0pt;"
    End With
    
    For Each n In col
        
        oEtapa.Carrega CLng(n)
    
        With cbbEtapa
            .AddItem
            .List(.ListCount - 1, 0) = oEtapa.Nome
            .List(.ListCount - 1, 1) = oEtapa.ID
        End With
        
    Next n
    
    cbbEtapa.ListIndex = idx

End Sub
Private Sub cbbProdutoPopular()
    
    Dim idx         As Integer
    Dim col         As New Collection
    Dim n           As Variant

    Set col = oProduto.Listar("nome")
    
    idx = cbbProduto.ListIndex
    
    cbbProduto.Clear
    
    For Each n In col
        
        oProduto.Carrega CLng(n)
    
        With cbbProduto
            .AddItem
            .List(.ListCount - 1, 0) = oProduto.Nome
            .List(.ListCount - 1, 1) = oProduto.ID
        End With
        
    Next n
    
    cbbProduto.ListIndex = idx

End Sub
Private Sub cbbUMPopular()
    
    Dim idx         As Integer
    Dim col         As New Collection
    Dim n           As Variant

    Set col = oUM.Listar("abreviacao")
    
    idx = cbbUM.ListIndex
    
    cbbUM.Clear
    
    For Each n In col
        
        oUM.Carrega CLng(n)
    
        With cbbUM
            .AddItem
            .List(.ListCount - 1, 0) = oUM.Abreviacao
            .List(.ListCount - 1, 1) = oUM.ID
        End With
        
    Next n
    
    cbbUM.ListIndex = idx

End Sub
