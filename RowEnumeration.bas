Attribute VB_Name = "RowEnumeration"
Option Compare Database
Option Explicit
'
' VBA.RowNumbers V1.4.3
' (c) Gustav Brock, Cactus Data ApS, CPH
' https://github.com/GustavBrock/VBA.RowCount
'
' Functions for enumeration of records in queries and forms,
' either stored or created on the fly.

' Enumerations.

'   Ranking strategies. Numeric values match those of:
'   https://se.mathworks.com/matlabcentral/fileexchange/70301-ranknum
Public Enum ApRankingStrategy
    apDense = 1
    apOrdinal = 2
    apStandardCompetition = 3
    apModifiedCompetition = 4
    apFractional = 5
End Enum

'   Ranking orders.
Public Enum ApRankingOrder
    apDescending = 0
    apAscending = 1
End Enum

' Builds random row numbers in a select, append, or create query
' with the option of a initial automatic reset.
'
' Usage (typical select query with random ordering):
'   SELECT RandomRowNumber(CStr([ID])) AS RandomRowID, *
'   FROM SomeTable
'   WHERE (RandomRowNumber(CStr([ID])) <> RandomRowNumber("",True))
'   ORDER BY RandomRowNumber(CStr([ID]));
'
' The Where statement shuffles the sequence when the query is run.
'
' Usage (typical select query for a form with random ordering):
'   SELECT RandomRowNumber(CStr([ID])) AS RandomRowID, *
'   FROM SomeTable
'   ORDER BY RandomRowNumber(CStr([ID]));
'
' The RandomRowID values will resist reordering and refiltering of the form.
' The sequence can be shuffled at will from, for example, a button click:
'
'   Private Sub ResetRandomButton_Click()
'       RandomRowNumber vbNullString, True
'       Me.Requery
'   End Sub
'
' and erased each time the form is closed:
'
'   Private Sub Form_Close()
'       RandomRowNumber vbNullString, True
'   End Sub
'
' Usage (typical append query, manual reset):
' 1. Reset random counter manually:
'   Call RandomRowNumber(vbNullString, True)
' 2. Run query:
'   INSERT INTO TempTable ( [RandomRowID] )
'   SELECT RandomRowNumber(CStr([ID])) AS RandomRowID, *
'   FROM SomeTable;
'
' Usage (typical append query, automatic reset):
'   INSERT INTO TempTable ( [RandomRowID] )
'   SELECT RandomRowNumber(CStr([ID])) AS RandomRowID, *
'   FROM SomeTable
'   WHERE (RandomRowNumber("",True)=0);
'
' 2018-09-11. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function RandomRowNumber( _
    ByVal Key As String, _
    Optional Reset As Boolean) _
    As Single

    ' Error codes.
    ' This key is already associated with an element of this collection.
    Const KeyIsInUse        As Long = 457
    
    Static Keys             As New Collection
  
    On Error GoTo Err_RandomRowNumber
    
    If Reset = True Then
        Set Keys = Nothing
    Else
        Keys.Add Rnd(-Timer * Keys.Count), Key
    End If
    
    RandomRowNumber = Keys(Key)
    
Exit_RandomRowNumber:
    Exit Function
    
Err_RandomRowNumber:
    Select Case Err
        Case KeyIsInUse
            ' Key is present.
            Resume Next
        Case Else
            ' Some other error.
            Resume Exit_RandomRowNumber
    End Select

End Function

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
' 2018-09-14. Gustav Brock, Cactus Data ApS, CPH.
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
    ElseIf Form.NewRecord = True Then
        ' No record number on a new record.
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

' Returns the count of records in form Form.
'
' Implementation, typical:
'
'   Create a TextBox to display the record count.
'   Set the ControlSource of this to:
'
'       =RecordCount([Form])
'
'   NB: For localised versions of Access, when entering the expression, type
'
'       =RecordCount([LocalisedNameOfObjectForm])
'
'   for example:
'
'       =RecordCount([Formular])
'
'   and press Enter. The expression will update to:
'
'       =RecordCount([Form])
'
'   If the form can delete records, insert this code line in the
'   AfterDelConfirm event:
'
'       Private Sub Form_AfterDelConfirm(Status As Integer)
'           Me!RecordCount.Requery
'       End Sub
'
'
'   If the form can add records, insert this code line in the
'   AfterInsert and OnCurrent events respectively:
'
'       Private Sub Form_AfterInsert()
'           Me!RecordCount.Requery
'       End Sub
'
'       Private Sub Form_Current()'
'           Static NewRecord    As Boolean
'
'           If NewRecord <> Me.NewRecord Then
'               Me!RecordCount.Requery
'               NewRecord = Me.NewRecord
'           End If
'       End Sub
'
' 2018-09-14. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function RecordCount( _
    ByRef Form As Access.Form) _
    As Long

    Dim Records As DAO.Recordset
    Dim Count   As Long

    Set Records = Form.RecordsetClone
    If Not Records.EOF Then
        Records.MoveLast
    End If
    Count = Records.RecordCount + Abs(Form.NewRecord)
    Records.Close

    RecordCount = Count

End Function

' Builds consecutive row numbers in a select, append, or create query
' with the option of a initial automatic reset.
' Optionally, a grouping key can be passed to reset the row count
' for every group key.
'
' Usage (typical select query having an ID with an index):
'   SELECT RowNumber(CStr([ID])) AS RowID, *
'   FROM SomeTable
'   WHERE (RowNumber(CStr([ID])) <> RowNumber("","",True));
'
' Usage (typical select query having an ID without an index):
'   SELECT RowNumber(CStr([ID])) AS RowID, *
'   FROM SomeTable
'   WHERE (RowNumber("","",True)=0);
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
'   Call RowNumber(vbNullString, True)
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
' 2020-05-29. Gustav Brock, Cactus Data ApS, CPH.
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
'   and after inserting or deleting records:
'
'       Private Sub Form_AfterDelConfirm(Status As Integer)
'           RowPriority Me.Priority
'       End Sub
'
'       Private Sub Form_AfterInsert()
'           RowPriority Me.Priority
'       End Sub
'
'   Optionally, if the control holding the primary key is not named Id:
'
'       Private Sub Priority_AfterUpdate()
'           RowPriority Me.Priority, NameOfPrimaryKeyControl
'       End Sub
'
'       Private Sub Form_AfterDelConfirm(Status As Integer)
'           RowPriority Me.Priority, NameOfPrimaryKeyControl
'       End Sub
'
'       Private Sub Form_AfterInsert()
'           RowPriority Me.Priority, NameOfPrimaryKeyControl
'       End Sub
'
' 2022-03-12. Gustav Brock, Cactus Data ApS, CPH.
'
Public Sub RowPriority( _
    ByRef TextBox As Access.TextBox, _
    Optional ByVal IdControlName As String = "Id")
    
    ' Error codes.
    ' This action is not supported in transactions.
    Const NotSupported      As Long = 3246

    Dim Form                As Access.Form
    Dim Records             As DAO.Recordset
    
    Dim RecordId            As Long
    Dim NewPriority         As Long
    Dim PriorityFix         As Long
    Dim FieldName           As String
    Dim IdFieldName         As String
    
    Dim Prompt              As String
    Dim Buttons             As VbMsgBoxStyle
    Dim Title               As String
    
    On Error GoTo Err_RowPriority
    
    Set Form = TextBox.Parent
    
    If Form.NewRecord Then
        ' Will happen if the last record of the form is deleted.
        Exit Sub
    Else
        ' Save record.
        Form.Dirty = False
    End If
    
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
    
    ' Disable a filter.
    ' If a filter is applied, only the filtered records
    ' will be reordered, and duplicates might be created.
    Form.FilterOn = False
    
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
    
    ' Set default value for a new record.
    TextBox.DefaultValue = NewPriority + 1
    
    ' Reorder form and relocate record position.
    ' Will fail if more than one record is pasted in.
    Form.Requery
    Set Records = Form.RecordsetClone
    Records.FindFirst "[" & IdFieldName & "] = " & RecordId & ""
    Form.Bookmark = Records.Bookmark
   
PreExit_RowPriority:
    ' Enable a filter.
    Form.FilterOn = True
    ' Present form.
    Form.Painting = True
    DoCmd.Hourglass False
    
    Set Records = Nothing
    Set Form = Nothing
    
Exit_RowPriority:
    Exit Sub
    
Err_RowPriority:
    Select Case Err.Number
        Case NotSupported
            ' Will happen if more than one record is pasted in.
            Resume PreExit_RowPriority
        Case Else
            ' Unexpected error.
            Prompt = "Error " & Err.Number & ": " & Err.Description
            Buttons = vbCritical + vbOKOnly
            Title = Form.Name
            MsgBox Prompt, Buttons, Title
            
            ' Restore form.
            Form.Painting = True
            DoCmd.Hourglass False
            Resume Exit_RowPriority
    End Select
    
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

    Const FirstNumber       As Long = 1
    
    Dim Form                As Access.Form
    Dim Records             As DAO.Recordset
    
    Dim FieldName           As String
        
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

' Loop through a recordset and align the values of a priority field
' to be valid and sequential.
'
' Default name for the priority field is Priority.
' Another name can be specified in parameter FieldName.
'
' Typical usage:
'   1.  Run code or query that updates, deletes, or appends records to
'       a table holding a priority field.
'
'   2.  Open an updatable and sorted DAO recordset (Records) with the table:
'
'       Dim Records As DAO.Recordset
'       Set Records = CurrentDb("Select * From Table Order By SomeField")

'   3.  Call this function, passing it the recordset:
'
'       AlignPriority Records
'
' 2018-09-04. Gustav Brock, Cactus Data ApS, CPH.
'
Public Sub AlignPriority( _
    ByRef Records As DAO.Recordset, _
    Optional FieldName As String)

    Const FirstNumber       As Long = 1
    Const PriorityFieldName As String = "Priority"
    
    Dim Field               As DAO.Field
    
    Dim CurrentPriority     As Long
    Dim NextPriority        As Long
    
    If FieldName = "" Then
        FieldName = PriorityFieldName
    End If
    ' Verify that the field exists.
    For Each Field In Records.Fields
        If Field.Name = FieldName Then
            Exit For
        End If
    Next
    ' If FieldName is not present, exit silently.
    If Field Is Nothing Then Exit Sub
    
    NextPriority = FirstNumber
    ' Set each record's priority to match its current position as
    ' defined by the sorting of the recordset.
    Records.MoveFirst
    While Not Records.EOF
        CurrentPriority = Nz(Field.Value, 0)
        If CurrentPriority = NextPriority Then
            ' No update needed.
        Else
            ' Assign and save adjusted priority.
            Records.Edit
                Field.Value = NextPriority
            Records.Update
        End If
        Records.MoveNext
        NextPriority = NextPriority + 1
    Wend
    
End Sub

' Returns, by the value of a field, the rank of one or more records of a table or query.
' Supports all five common ranking strategies (methods).
'
' Source:
'   WikiPedia: https://en.wikipedia.org/wiki/Ranking
'
' Supports ranking of descending as well as ascending values.
' Any ranking will require one table scan only.
' For strategy Ordinal, a a second field with a subvalue must be used.
'
' Typical usage (table Products of Northwind sample database):
'
'   SELECT Products.*, RowRank("[Standard Cost]","[Products]",[Standard Cost]) AS Rank
'   FROM Products
'   ORDER BY Products.[Standard Cost] DESC;
'
' Typical usage for strategy Ordinal with a second field ([Product Code]) holding the subvalues:
'
'   SELECT Products.*, RowRank("[Standard Cost],[Product Code]","[Products]",[Standard Cost],[Product Code],2) AS Ordinal
'   FROM Products
'   ORDER BY Products.[Standard Cost] DESC;
'
' To obtain a rank, the first three parameters must be passed.
' Four parameters is required for strategy Ordinal to be returned properly.
' The remaining parameters are optional.
'
' The ranking will be cached until Order is changed or RowRank is called to clear the cache.
' To clear the cache, call RowRank with no parameters:
'
'   RowRank
'
' Parameters:
'
'   Expression: One field name for other strategies than Ordinal, two field names for this.
'   Domain:     Table or query name.
'   Value:      The values to rank.
'   SubValue:   The subvalues to rank when using strategy Ordinal.
'   Strategy:   Strategy for the ranking.
'   Order:      The order by which to rank the values (and subvalues).
'
' 2019-07-11. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function RowRank( _
    Optional ByVal Expression As String, _
    Optional ByVal Domain As String, _
    Optional ByVal Value As Variant, _
    Optional ByVal SubValue As Variant, _
    Optional ByVal Strategy As ApRankingStrategy = ApRankingStrategy.apStandardCompetition, _
    Optional ByVal Order As ApRankingOrder = ApRankingOrder.apDescending) _
    As Double
    
    Const SqlMask1          As String = "Select Top 1 {0} From {1}"
    Const SqlMask           As String = "Select {0} From {1} Order By 1 {2}"
    Const SqlOrder          As String = ",{0} {1}"
    Const OrderAsc          As String = "Asc"
    Const OrderDesc         As String = "Desc"
    Const FirstStrategy     As Integer = ApRankingStrategy.apDense
    Const LastStrategy      As Integer = ApRankingStrategy.apFractional
    
    ' Expected error codes to accept.
    Const CannotAddKey      As Long = 457
    Const CannotFindKey     As Long = 5
    ' Uncommon character string to assemble Key and SubKey as a compound key.
    Const KeySeparator      As String = "¤§¤"
    
    ' Array of the collections for the five strategies.
    Static Ranks(FirstStrategy To LastStrategy) As Collection
    ' The last sort order used.
    Static LastOrder        As ApRankingOrder

    Dim Records             As DAO.Recordset
    
    ' Array to hold the rank for each strategy.
    Dim Rank(FirstStrategy To LastStrategy)     As Double
    
    Dim Item                As Integer
    Dim Sql                 As String
    Dim SortCount           As Integer
    Dim SortOrder           As String
    Dim LastKey             As String
    Dim Key                 As String
    Dim SubKey              As String
    Dim Dupes               As Integer
    Dim Delta               As Long
    Dim ThisStrategy        As ApRankingStrategy

    On Error GoTo Err_RowRank
    
    If Expression = "" Then
        ' Erase the collections of keys.
        For Item = LBound(Ranks) To UBound(Ranks)
            Set Ranks(Item) = Nothing
        Next
    Else
        If LastOrder <> Order Or Ranks(FirstStrategy) Is Nothing Then
            ' Initialize the collections and reset their ranks.
            For Item = LBound(Ranks) To UBound(Ranks)
                Set Ranks(Item) = New Collection
                Rank(Item) = 0
            Next
            
            ' Build order clause.
            Sql = Replace(Replace(SqlMask1, "{0}", Expression), "{1}", Domain)
            SortCount = CurrentDb.OpenRecordset(Sql, dbReadOnly).Fields.Count
            
            If Order = ApRankingOrder.apDescending Then
                ' Descending sorting (default).
                SortOrder = OrderDesc
            Else
                ' Ascending sorting.
                SortOrder = OrderAsc
            End If
            LastOrder = Order
            
            ' Build SQL.
            Sql = Replace(Replace(Replace(SqlMask, "{0}", Expression), "{1}", Domain), "{2}", SortOrder)
            ' Add a second sort field, if present.
            If SortCount >= 2 Then
                Sql = Sql & Replace(Replace(SqlOrder, "{0}", 2), "{1}", SortOrder)
            End If

            ' Open ordered recordset.
            Set Records = CurrentDb.OpenRecordset(Sql, dbReadOnly)
            ' Loop the recordset once while creating all the collections of ranks.
            While Not Records.EOF
                Key = CStr(Nz(Records.Fields(0).Value))
                SubKey = ""
                ' Create the sub key if a second field is present.
                If SortCount > 1 Then
                    SubKey = CStr(Nz(Records.Fields(1).Value))
                End If
                
                If LastKey <> Key Then
                    ' Add new entries.
                    For ThisStrategy = FirstStrategy To LastStrategy
                        Select Case ThisStrategy
                            Case ApRankingStrategy.apDense
                                Rank(ThisStrategy) = Rank(ThisStrategy) + 1
                            Case ApRankingStrategy.apStandardCompetition
                                Rank(ThisStrategy) = Rank(ThisStrategy) + 1 + Dupes
                                Dupes = 0
                            Case ApRankingStrategy.apModifiedCompetition
                                Rank(ThisStrategy) = Rank(ThisStrategy) + 1
                            Case ApRankingStrategy.apOrdinal
                                Rank(ThisStrategy) = Rank(ThisStrategy) + 1
                                ' Add entry using both Key and SubKey
                                Ranks(ThisStrategy).Add Rank(ThisStrategy), Key & KeySeparator & SubKey
                            Case ApRankingStrategy.apFractional
                                Rank(ThisStrategy) = Rank(ThisStrategy) + 1 + Delta / 2
                                Delta = 0
                        End Select
                        If ThisStrategy = ApRankingStrategy.apOrdinal Then
                            ' Key with SubKey has been added above for this strategy.
                        Else
                            ' Add key for all other strategies.
                            Ranks(ThisStrategy).Add Rank(ThisStrategy), Key
                        End If
                    Next
                    LastKey = Key
                Else
                    ' Modify entries and/or counters for those strategies that require this for a repeated key.
                    For ThisStrategy = FirstStrategy To LastStrategy
                        Select Case ThisStrategy
                            Case ApRankingStrategy.apDense
                            Case ApRankingStrategy.apStandardCompetition
                                Dupes = Dupes + 1
                            Case ApRankingStrategy.apModifiedCompetition
                                Rank(ThisStrategy) = Rank(ThisStrategy) + 1
                                Ranks(ThisStrategy).Remove Key
                                Ranks(ThisStrategy).Add Rank(ThisStrategy), Key
                            Case ApRankingStrategy.apOrdinal
                                Rank(ThisStrategy) = Rank(ThisStrategy) + 1
                                ' Will fail for a repeated value of SubKey.
                                Ranks(ThisStrategy).Add Rank(ThisStrategy), Key & KeySeparator & SubKey
                            Case ApRankingStrategy.apFractional
                                Rank(ThisStrategy) = Rank(ThisStrategy) + 0.5
                                Ranks(ThisStrategy).Remove Key
                                Ranks(ThisStrategy).Add Rank(ThisStrategy), Key
                                Delta = Delta + 1
                        End Select
                    Next
                End If
                Records.MoveNext
            Wend
            Records.Close
        End If
        
        ' Retrieve the rank for the current strategy.
        If Strategy = ApRankingStrategy.apOrdinal Then
            ' Use both Value and SubValue.
            Key = CStr(Nz(Value)) & KeySeparator & CStr(Nz(SubValue))
        Else
            ' Use Value only.
            Key = CStr(Nz(Value))
        End If
        ' Will fail if key isn't present.
        Rank(Strategy) = Ranks(Strategy).Item(Key)
    End If
    
    RowRank = Rank(Strategy)
    
Exit_RowRank:
    Exit Function
    
Err_RowRank:
    Select Case Err
        Case CannotAddKey
            ' Key is present, thus cannot be added again.
            Resume Next
        Case CannotFindKey
            ' Key is not present, thus cannot be removed.
            Resume Next
        Case Else
            ' Some other error. Ignore.
            Resume Exit_RowRank
    End Select
    
End Function
