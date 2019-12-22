VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fLancamentoDireto 
   Caption         =   ":: Lancamentos diretos ::"
   ClientHeight    =   8430
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12885
   OleObjectBlob   =   "fLancamentoDireto.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fLancamentoDireto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private oConta              As New cConta

Private colControles        As New Collection
Private myRst               As ADODB.RecordSet
Private lPagina             As Long

Private Const sTable As String = "tbl_compras"
Private Const sCampoOrderBy As String = "data"
Private Sub UserForm_Initialize()
    
    Call cbbContaPopular

End Sub
Private Sub UserForm_Terminate()
    
    Set oConta = Nothing
    Call Desconecta
    
End Sub
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



