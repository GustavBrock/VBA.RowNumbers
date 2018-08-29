Attribute VB_Name = "RowEnumeration"
Option Compare Database
Option Explicit
'
' VBA.RowNumbers V1.1.0
' (c) Gustav Brock, Cactus Data ApS, CPH
' https://github.com/GustavBrock/VBA.RowCount
'
' Functions for enumaration of records in queries and forms,
' either stored or created on the fly.
'

' Creates and returns a sequential record number for records displayed
' in a form, even if no primary or unique key is present.
' For a new record, Null is returned until the record is saved.
'
' Implementation, typical:
'
'   Create a TextBox to display the record number.
'   Set the ControlSource of this to:
'
'       =RecordNumber([Form])
'
'   The returned number will equal the Current Record displayed in the
'   form's record navigator (bottom-left).
'   Optionally, specify another first number than 1, say, 0:
'
'       =RecordNumber([Form],0)
'
'   NB: For localised versions of Access, when entering the expression, type
'
'       =RecordNumber([LocalisedNameOfObjectForm])
'
'   for example:
'
'       =RecordNumber([Formular])
'
'   and press Enter. The expression will update to:
'
'       =RecordNumber([Form])
'
'   If the form can delete records, insert this code line in the
'   AfterDelConfirm event:
'
'       Private Sub Form_AfterDelConfirm(Status As Integer)
'           Me!RecordNumber.Requery
'       End Sub
'
'   If the form can add records, insert this code line in the
'   AfterInsert event:
'
'       Private Sub Form_AfterInsert()
'           Me!RecordNumber.Requery
'       End Sub
'
' Implementation, stand-alone:
'
'   Dim Number As Variant
'
'   Number = RecordNumber(Forms(IndexOfFormInFormsCollection))
'   ' or
'   Number = RecordNumber(Forms("NameOfSomeOpenForm"))
'
'
' 2018-08-23. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function RecordNumber( _
    ByRef Form As Access.Form, _
    Optional ByVal FirstNumber As Long = 1) _
    As Variant

    ' Error code for "There is no current record."
    Const NoCurrentRecord   As Long = 3021
    
    Dim Records             As DAO.Recordset
    
    Dim Number              As Variant
    Dim Prompt              As String
    Dim Buttons             As VbMsgBoxStyle
    Dim Title               As String

    On Error GoTo Err_RecordNumber
    If Form Is Nothing Then
        ' No form object is passed.
        Number = Null
    ElseIf Form.Dirty = True Then
        ' No record number until the record is saved.
        Number = Null
    Else
        Set Records = Form.RecordsetClone
        Records.Bookmark = Form.Bookmark
        Number = FirstNumber + Records.AbsolutePosition
        Set Records = Nothing
    End If
    
Exit_RecordNumber:
    RecordNumber = Number
    Exit Function
    
Err_RecordNumber:
    Select Case Err.Number
        Case NoCurrentRecord
            ' Form is at new record, thus no Bookmark exists.
            ' Ignore and continue.
        Case Else
            ' Unexpected error.
            Prompt = "Error " & Err.Number & ": " & Err.Description
            Buttons = vbCritical + vbOKOnly
            Title = Form.Name
            MsgBox Prompt, Buttons, Title
    End Select
    
    ' Return Null for any error.
    Number = Null
    Resume Exit_RecordNumber

End Function

' Builds consecutive row numbers in a select, append or create query
' with the option of a initial automatic reset.
' Optionally, a grouping key can be passed to reset the row count
' for every group key.
'
' Usage (typical select query):
'   SELECT RowNumber(CStr([ID])) AS RowID, *
'   FROM SomeTable
'   WHERE (RowNumber(CStr([ID])) <> RowNumber("","",True));
'
' Usage (with group key):
'   SELECT RowNumber(CStr([ID]), CStr[GroupID])) AS RowID, *
'   FROM SomeTable
'   WHERE (RowNumber(CStr([ID])) <> RowNumber("","",True));
'
' The Where statement resets the counter when the query is run
' and is needed for browsing a select query.
'
' Usage (typical append query, manual reset):
' 1. Reset counter manually:
'   Call RowNumber(vbNullString)
' 2. Run query:
'   INSERT INTO TempTable ( [RowID] )
'   SELECT RowNumber(CStr([ID])) AS RowID, *
'   FROM SomeTable;
'
' Usage (typical append query, automatic reset):
'   INSERT INTO TempTable ( [RowID] )
'   SELECT RowNumber(CStr([ID])) AS RowID, *
'   FROM SomeTable
'   WHERE (RowNumber("","",True)=0);
'
' 2018-08-23. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function RowNumber( _
    ByVal Key As String, _
    Optional ByVal GroupKey As String, _
    Optional ByVal Reset As Boolean) _
    As Long
    
    ' Uncommon character string to assemble GroupKey and Key as a compound key.
    Const KeySeparator      As String = "¤§¤"
    ' Expected error codes to accept.
    Const CannotAddKey      As Long = 457
    Const CannotRemoveKey   As Long = 5
  
    Static Keys             As New Collection
    Static GroupKeys        As New Collection

    Dim Count               As Long
    Dim CompoundKey         As String
    
    On Error GoTo Err_RowNumber
    
    If Reset = True Then
        ' Erase the collection of keys and group key counts.
        Set Keys = Nothing
        Set GroupKeys = Nothing
    Else
        ' Create a compound key to uniquely identify GroupKey and its Key.
        ' Note: If GroupKey is not used, only one element will be added.
        CompoundKey = GroupKey & KeySeparator & Key
        Count = Keys(CompoundKey)
        
        If Count = 0 Then
            ' This record has not been enumerated.
            '
            ' Will either fail if the group key is new, leaving Count as zero,
            ' or retrieve the count of already enumerated records with this group key.
            Count = GroupKeys(GroupKey) + 1
            If Count > 0 Then
                ' The group key has been recorded.
                ' Remove it to allow it to be recreated holding the new count.
                GroupKeys.Remove (GroupKey)
            Else
                ' This record is the first having this group key.
                ' Thus, the count is 1.
                Count = 1
            End If
            ' (Re)create the group key item with the value of the count of keys.
            GroupKeys.Add Count, GroupKey
        End If

        ' Add the key and its enumeration.
        ' This will be:
        '   Using no group key: Relative to the full recordset.
        '   Using a group key:  Relative to the group key.
        ' Will fail if the key already has been created.
        Keys.Add Count, CompoundKey
    End If
    
    ' Return the key value as this is the row counter.
    RowNumber = Count
  
Exit_RowNumber:
    Exit Function
    
Err_RowNumber:
    Select Case Err
        Case CannotAddKey
            ' Key is present, thus cannot be added again.
            Resume Next
        Case CannotRemoveKey
            ' GroupKey is not present, thus cannot be removed.
            Resume Next
        Case Else
            ' Some other error. Ignore.
            Resume Exit_RowNumber
    End Select

End Function

' Set the priority order of a record relative to the other records of a form.
'
' The table/query bound to the form must have an updatable numeric field for
' storing the priority of the record. Default value of this should be Null.
'
' Requires:
'   A numeric, primary key, typical an AutoNumber field.
'
' Usage:
'   To be called from the AfterUpdate event of the Priority textbox:
'
'       Private Sub Priority_AfterUpdate()
'           RowPriority Me.Priority
'       End Sub
'
'   Optionally, if the control holding the primary key is not named Id:
'
'       Private Sub Priority_AfterUpdate()
'           RowPriority Me.Priority, NameOfPrimaryKeyControl
'       End Sub
'
' 2018-08-27. Gustav Brock, Cactus Data ApS, CPH.
'
Public Sub RowPriority( _
    ByRef TextBox As Access.TextBox, _
    Optional ByVal IdControlName As String = "Id")

    Dim Form            As Access.Form
    Dim Records         As DAO.Recordset
    
    Dim RecordId        As Long
    Dim NewPriority     As Long
    Dim PriorityFix     As Long
    Dim FieldName       As String
    Dim IdFieldName     As String
    
    Set Form = TextBox.Parent
    ' Save record.
    Form.Dirty = False
    
    ' Priority control can have any Name.
    FieldName = TextBox.ControlSource
    ' Id (primary key) control can have any name.
    IdFieldName = Form.Controls(IdControlName).ControlSource
    
    ' Prepare form.
    DoCmd.Hourglass True
    Form.Repaint
    Form.Painting = False
    
    ' Current Id and priority.
    RecordId = Form.Controls(IdControlName).Value
    PriorityFix = Nz(TextBox.Value, 0)
    If PriorityFix <= 0 Then
        PriorityFix = 1
        TextBox.Value = PriorityFix
        Form.Dirty = False
    End If
    
    ' Rebuild priority list.
    Set Records = Form.RecordsetClone
    Records.MoveFirst
    While Not Records.EOF
        If Records.Fields(IdFieldName).Value <> RecordId Then
        NewPriority = NewPriority + 1
            If NewPriority = PriorityFix Then
                ' Move this record to next lower priority.
                NewPriority = NewPriority + 1
            End If
            If Nz(Records.Fields(FieldName).Value, 0) = NewPriority Then
                ' Priority hasn't changed for this record.
            Else
                ' Assign new priority.
                Records.Edit
                    Records.Fields(FieldName).Value = NewPriority
                Records.Update
            End If
        End If
        Records.MoveNext
    Wend
    
    ' Reorder form and relocate record position.
    Form.Requery
    Set Records = Form.RecordsetClone
    Records.FindFirst "[" & IdFieldName & "] = " & RecordId & ""
    Form.Bookmark = Records.Bookmark
   
    ' Present form.
    Form.Painting = True
    DoCmd.Hourglass False
    
    Set Records = Nothing
    Set Form = Nothing
    
End Sub

' Set the priority order of the records to match a form's current record order.
'
' The table/query bound to the form must have an updatable numeric field for
' storing the priority of the records. Default value of this should be Null.
'
' Usage:
'   To be called from, say, a button click on the form.
'   The textbox Me.Priority is bound to the Priority field of the table:
'
'       Private Sub ResetPriorityButton_Click()
'           SetRowPriority Me.Priority
'       End Sub
'
'
' 2018-08-27. Gustav Brock, Cactus Data ApS, CPH.
'
Public Sub SetRowPriority(ByRef TextBox As Access.TextBox)

    Const FirstNumber   As Long = 1
    
    Dim Form            As Access.Form
    Dim Records         As DAO.Recordset
    
    Dim FieldName       As String
        
    Set Form = TextBox.Parent
    Set Records = Form.RecordsetClone
    
    ' TextBox can have any Name.
    FieldName = TextBox.ControlSource
    
    ' Pause form painting to speed up rebuilding of the records' priority.
    Form.Painting = False
    
    ' Set each record's priority to match its current position in the form.
    Records.MoveFirst
    While Not Records.EOF
        If Records.Fields(FieldName).Value = FirstNumber + Records.AbsolutePosition Then
            ' No update needed.
        Else
            ' Assign and save adjusted priority.
            Records.Edit
                Records.Fields(FieldName).Value = FirstNumber + Records.AbsolutePosition
            Records.Update
        End If
        Records.MoveNext
    Wend
    
    ' Repaint form.
    Form.Painting = True

    Set Records = Nothing
    Set Form = Nothing

End Sub

