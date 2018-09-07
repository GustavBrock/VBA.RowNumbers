VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Products"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_AfterDelConfirm(Status As Integer)

    ' Rebuild priority of all records having lower priority (= higher numeric value).
    RowPriority Me.Priority

End Sub

Private Sub Form_AfterInsert()

    ' Rebuild priority of all records having lower priority (= higher numeric value).
    RowPriority Me.Priority
    
End Sub

Private Sub Priority_AfterUpdate()

    ' Rebuild priority of all records having lower priority (= higher numeric value).
    RowPriority Me.Priority
    
End Sub

Private Sub ResetPriorityButton_Click()

    ' Set priority order to match the form's current record order.
    SetRowPriority Me.Priority

End Sub

