VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RetestUserForm 
   Caption         =   "UserForm1"
   ClientHeight    =   2340
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "RetestUserForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RetestUserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnOK_Click()
    ' Ensure at least one option is selected
    If chkNTC.Value = False And chkFunctional.Value = False Then
        MsgBox "Please select at least one plate to retest.", vbExclamation
    Else
        ' Hide the form if a valid selection is made
        Me.Hide
    End If
End Sub
