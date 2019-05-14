VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_RecordOfRecords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_AfterDelConfirm(Status As Integer)

    Me!RecordCount.Requery

End Sub

Private Sub Form_AfterInsert()

    Me!RecordCount.Requery
    
End Sub

Private Sub Form_Current()

    Static NewRecord    As Boolean
    
    If NewRecord <> Me.NewRecord Then
        Me!RecordCount.Requery
        NewRecord = Me.NewRecord
    End If

End Sub
