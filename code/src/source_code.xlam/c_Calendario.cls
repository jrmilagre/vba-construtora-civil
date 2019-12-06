VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_Calendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents btnGrupo As MSForms.CommandButton
Attribute btnGrupo.VB_VarHelpID = -1

Private Sub btnGrupo_Click()
    btnGrupo.Parent.Tag = btnGrupo.Tag
    btnGrupo.Parent.Hide
End Sub
