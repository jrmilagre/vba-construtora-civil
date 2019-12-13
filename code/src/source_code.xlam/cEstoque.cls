VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cEstoque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type TEstoque
    Estoque         As Double
    CustoMedio      As Currency
    CustoTotal      As Currency
End Type

Private This As TEstoque

Public Property Get Quantidade() As Double: Estoque = This.Estoque: End Property
Public Property Let Quantidade(ByVal Value As Double): This.Estoque = Value: End Property

Public Property Get CustoMedio() As Currency: CustoMedio = This.CustoMedio: End Property
Public Property Let CustoMedio(ByVal Value As Currency): This.CustoMedio = Value: End Property

Public Property Get CustoTotal() As Currency: CustoTotal = This.CustoTotal: End Property
Public Property Let CustoTotal(ByVal Value As Currency): This.CustoTotal = Value: End Property

Public Sub DiminuiEstoque(QtdeAtual As Double, _
                           CustoAtual As Currency, _
                           QtdeSaida As Double, _
                           CustoSaida As Currency, _
                           ProdutoID As Long)
                           
    Dim cCTA As Currency ' Custo Total Atual
    Dim cCTM As Currency ' Custo Total Movimento
    Dim cCTF As Currency ' Custo Total Futuro
    Dim dEF As Double    ' Estoque Futuro
    
    cCTA = (QtdeAtual * CustoAtual)
    cCTM = (QtdeSaida * CustoSaida)
    
    dEF = (QtdeAtual - QtdeSaida)
    
    If dEF = 0 Then
        cCTF = 0
    Else
        cCTF = (cCTA - cCTM) / dEF
    End If
    
    sSQL = "UPDATE tbl_produtos "
    sSQL = sSQL & "SET "
    sSQL = sSQL & "estoque = " & dEF & ", "
    sSQL = sSQL & "custo_medio = " & cCTF & " "
    sSQL = sSQL & "WHERE id = " & ProdutoID
        
    cnn.Execute sSQL
    
End Sub
Public Sub AumentaEstoque(QtdeAtual As Double, _
                           CustoAtual As Currency, _
                           QtdeEntrada As Double, _
                           CustoEntrada As Currency, _
                           ProdutoID As Long)
                           
    Dim cCTA As Currency ' Custo Total Atual
    Dim cCTM As Currency ' Custo Total Movimento
    Dim cCTF As Currency ' Custo Total Futuro
    Dim dEF As Double    ' Estoque Futuro
    
    cCTA = (QtdeAtual * CustoAtual)
    cCTM = (QtdeEntrada * CustoEntrada)
    
    dEF = (QtdeAtual + QtdeEntrada)
    
    If dEF = 0 Then
        cCTF = 0
    Else
        cCTF = (cCTA + cCTM) / dEF
    End If
    
    sSQL = "UPDATE tbl_produtos "
    sSQL = sSQL & "SET "
    sSQL = sSQL & "estoque = " & Replace(dEF, ",", ".") & ", "
    sSQL = sSQL & "custo_medio = " & Replace(cCTF, ",", ".") & " "
    sSQL = sSQL & "WHERE id = " & ProdutoID
        
    cnn.Execute sSQL
                               
End Sub
Public Sub AtualizaEstoque(ProdutoID As Long, Qtde As Double, CustoMedio As Currency)



End Sub

