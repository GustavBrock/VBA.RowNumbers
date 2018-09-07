Attribute VB_Name = "Demo"
Option Compare Database
Option Explicit
'
' VBA.RowNumbers Demo V1.0.0
' (c) Gustav Brock, Cactus Data ApS, CPH
' https://github.com/GustavBrock/VBA.RowCount
'
' Function demonstrate typical usage for some of the
' functions of module RowEnumeration.
'

' Examples for how to use AlignPriority.
'
' Requires:
'   Table Products from example database Northwind 2007.
'   A field added to this, named Priority, data type Long.
'
' 2018-09-04. Gustav Brock, Cactus Data ApS, CPH.
'
Public Sub PriorityCleanTest()

    Dim Records As DAO.Recordset
    
    Dim Sql     As String
    
    ' Maintain Priority field in table Products.
    ' "Comment out" those not to use.
    '
    ' Assign high priority to invalid priority values for these to be listed first.
''    Sql = "Select * From Products Order By Priority"
    '
    ' Assign low priority to invalid priority values for these to be listed last.
    Sql = "Select * From Products Order By Abs(Priority Is Null), Priority"
    '
    ' Align Priority to a search order, here "ID Asc".
''    Sql = "Select * From Products Order By ID"

    Set Records = CurrentDb.OpenRecordset(Sql)
    
    AlignPriority Records
    
    Records.Close
    
End Sub

