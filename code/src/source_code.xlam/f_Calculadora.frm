VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} f_Calculadora 
   Caption         =   ":: Calculadora"
   ClientHeight    =   2865
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2790
   OleObjectBlob   =   "f_Calculadora.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "f_Calculadora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Dim oCalculadora As New c_Calculadora

Dim dMemoria    As Double
Dim iOperacao   As Integer
Enum Operacao
    Adicao = 1
    Subtracao = 2
    Multiplicacao = 3
    Divisao = 4
End Enum

Private Sub btn0_Click()
    txbVisor.Text = IIf(txbVisor.Text = "0", "0", txbVisor.Text + "0")
End Sub
Private Sub btn1_Click()
    txbVisor.Text = IIf(txbVisor.Text = "0", "1", txbVisor.Text + "1")
End Sub
Private Sub btn2_Click()
    txbVisor.Text = IIf(txbVisor.Text = "0", "2", txbVisor.Text + "2")
End Sub
Private Sub btn3_Click()
    txbVisor.Text = IIf(txbVisor.Text = "0", "3", txbVisor.Text + "3")
End Sub
Private Sub btn4_Click()
    txbVisor.Text = IIf(txbVisor.Text = "0", "4", txbVisor.Text + "4")
End Sub
Private Sub btn5_Click()
    txbVisor.Text = IIf(txbVisor.Text = "0", "5", txbVisor.Text + "5")
End Sub
Private Sub btn6_Click()
    txbVisor.Text = IIf(txbVisor.Text = "0", "6", txbVisor.Text + "6")
End Sub
Private Sub btn7_Click()
    txbVisor.Text = IIf(txbVisor.Text = "0", "7", txbVisor.Text + "7")
End Sub
Private Sub btn8_Click()
    txbVisor.Text = IIf(txbVisor.Text = "0", "8", txbVisor.Text + "8")
End Sub
Private Sub btn9_Click()
    txbVisor.Text = IIf(txbVisor.Text = "0", "9", txbVisor.Text + "9")
End Sub
Private Sub btnSomar_Click()
    dMemoria = CDbl(txbVisor.Text)
    txbVisor.Text = 0
    txbVisor.SelStart = 0
    txbVisor.SelLength = Len(txbVisor.Text)
    iOperacao = Operacao.Adicao
End Sub
Private Sub btnSubtrair_Click()
    dMemoria = CDbl(txbVisor.Text)
    txbVisor.Text = 0
    txbVisor.SelStart = 0
    txbVisor.SelLength = Len(txbVisor.Text)
    iOperacao = Operacao.Subtracao
End Sub
Private Sub btnMultiplicar_Click()
    dMemoria = CDbl(txbVisor.Text)
    txbVisor.Text = 0
    txbVisor.SelStart = 0
    txbVisor.SelLength = Len(txbVisor.Text)
    iOperacao = Operacao.Multiplicacao
End Sub
Private Sub btnDividir_Click()
    dMemoria = CDbl(txbVisor.Text)
    txbVisor.Text = 0
    txbVisor.SelStart = 0
    txbVisor.SelLength = Len(txbVisor.Text)
    iOperacao = Operacao.Divisao
End Sub

Private Sub btnResultado_Click()
    Select Case iOperacao
        Case Operacao.Adicao
            txbVisor.Text = dMemoria + txbVisor.Text
        Case Operacao.Subtracao
            txbVisor.Text = dMemoria - txbVisor.Text
        Case Operacao.Multiplicacao
            txbVisor.Text = dMemoria * txbVisor.Text
        Case Operacao.Divisao
            txbVisor.Text = dMemoria / txbVisor.Text
        Case Else
            MsgBox "Dígito inválido!"
    End Select
    
    iOperacao = 0
End Sub

Private Sub btnUsar_Click()
    oCalculadora.Resultado = CDbl(txbVisor.Text)
End Sub

Private Sub btnVirgula_Click()
    txbVisor.Text = txbVisor.Text + ","
End Sub

Private Sub txbVisor_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode
        Case 13
            If iOperacao > 0 Then
                Call btnResultado_Click: btnUsar.SetFocus
            Else
                Call btnUsar_Click: Unload Me
            End If
        Case 107
            Call btnSomar_Click
    End Select
End Sub

Private Sub txbVisor_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' Permite apenas números
    Select Case KeyAscii
        Case 8                      ' Backspace (seta de apagar)
        Case 48 To 57               ' Números de 0 a 9
        Case 44                     ' Vírgula
        
        If InStr(KeyAscii, ",") Then 'Se o campo já tiver vírgula então ele não adiciona
            KeyAscii = 0 'Não adiciona a vírgula caso ja tenha
        Else
            KeyAscii = 44 'Adiciona uma vírgula
        End If
        
        Case Else
            KeyAscii = 0 'Não deixa nenhuma outra caractere ser escrito
    End Select
End Sub

Private Sub UserForm_Initialize()
    txbVisor.Text = IIf(ccurVisor > 0, Format(ccurVisor, "#,##0.00"), Format(0, "#,##0.00"))
    txbVisor.SelStart = 0
    txbVisor.SelLength = Len(txbVisor.Text)
End Sub
