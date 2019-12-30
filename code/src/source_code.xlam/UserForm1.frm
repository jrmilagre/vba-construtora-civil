VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    
    Call MySubroutine1
    
End Sub
Public Sub MySubroutine1()
   Call MySubroutine2(1)
   Call MySubroutine2(2)
   Call MySubroutine2(3)
   Call MySubroutine2(4)
End Sub

Public Sub MySubroutine2(ByVal iValue As Integer)

    'Debug.Print Application.VBE.ActiveCodePane.CodeModule
    'Debug.Print Application.VBE.SelectedVBComponent.name
    'Debug.Print Application.VBE.ActiveCodePane.CodeModule.name
    'Debug.Print Application.VBE.SelectedVBComponent.CodeModule.Parent.name
    Debug.Print Application.VBE.ActiveCodePane.Collection.Count
    

End Sub
