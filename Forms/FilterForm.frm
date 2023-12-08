VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FilterForm 
   Caption         =   "Filter "
   ClientHeight    =   3915
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5820
   OleObjectBlob   =   "FilterForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FilterForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()

End Sub
Option Explicit
Private Sub CancelButton_Click()
FilterForm.Hide
PatientIDTextBox.Text = ""

End Sub

Private Sub SubmitButton_Click()

WriteToDocument (PatientIDTextBox.Text)
FilterForm.Hide
PatientIDTextBox.Text = ""

End Sub
