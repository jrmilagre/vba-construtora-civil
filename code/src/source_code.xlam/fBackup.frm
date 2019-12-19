VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fBackup 
   Caption         =   ":: Backup ::"
   ClientHeight    =   2475
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8340
   OleObjectBlob   =   "fBackup.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()

    Dim rst As New ADODB.RecordSet
    
    sSQL = "SELECT parametro, valor_unico FROM tbl_parametros WHERE parametro = 'backup'"
    
    rst.Open sSQL, cnn
    
    If rst.EOF = True Then
        sSQL = "INSERT INTO tbl_parametros ([parametro], [valor_unico]) VALUES ('backup', 'C:\') ": cnn.Execute sSQL
        txbCaminho.Text = "C:\"
    Else
        txbCaminho.Text = rst.Fields("valor_unico").Value
    End If
    
    Set rst = Nothing

End Sub
Private Sub btnLocalizar_Click()
    Call SelectFolder
End Sub
Private Sub btnBackup_Click()

    Dim b As Boolean

    b = m_Database.Backup(txbCaminho.Text & "\" & Format(Now, "yyyymmdd-hhMMss") & ".mdb")
    
    If b = True Then
        Unload Me
    End If
End Sub

Private Sub SelectFolder()

    Dim sFolder As String

    ' Open the select folder prompt
    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = txbCaminho.Text & "\"
        If .Show = -1 Then ' if OK is pressed
            sFolder = .SelectedItems(1)
        End If
    End With
    
    If sFolder <> "" Then ' if a file was chosen
        txbCaminho.Text = sFolder
    Else
        Exit Sub
        'MsgBox "Caminho inválido"
        'txbCaminho.Text = "C:\"
    End If
    
    sSQL = "UPDATE tbl_parametros SET valor_unico = '" & txbCaminho.Text & "' WHERE parametro = 'backup'"
    
    cnn.Execute sSQL

End Sub
Private Sub UserForm_Terminate()
    
    ' Destrói objeto da classe cProduto
    Call Desconecta
End Sub
