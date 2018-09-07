VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_RandomProducts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Close()

    ' Erase the assigned random numbers.
    RandomRowNumber vbNullString, True
    
End Sub

Private Sub ResetRandomButton_Click()

    ' Erase the assigned random numbers.
    RandomRowNumber vbNullString, True
    ' Rebuild the assigned random numbers by requering the form.
    Me.Requery
    
End Sub
