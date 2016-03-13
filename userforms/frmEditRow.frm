VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEditRow 
   Caption         =   "Edit Row"
   ClientHeight    =   5166
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   4193
   OleObjectBlob   =   "frmEditRow.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEditRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' FillRow backend
' This module provides the backend functions frmFillRow
' Contents of the textboxes are refered to in comments by their labels
' REQUIRES: ExcellIO must be imported and initalizeRows must be called before this userform is shown
' Rows must be recognized by this userform before it is shown (they will automatically deleted when this userform is closed)

'***Interface***

'Sub: initalizeRows()
'Purpose: initalizes the rows that the userform recognizes, required to be called before the userform is shown
'Effects: rowStartingCells and rowNames have their values reinitalized

'Sub: addRow(newRowName, newRowStartingCell)
'Purpose: causes the userform to recognize a new row with a name of newRowName starting at newRowStartingCell
'Effects: rowStartingCells and rowNames may have their values changed
'Requires: initalizeRows() must be called first
'          newRowStartingCell must be the location of a (only one) vaild cell
'          there are no other rows currently recognized by the userform with the name of newRowName

'Sub: deleteRow(oldRowName)
'Purpose: causes the userform to forget a currently recognized row with a name of oldRowName
'Effects: rowStartingCells and rowNames may have their values changed
'Requires: initalizeRows() must be called first
'          there must be a row currently recognized by the userform with the name of oldRowName

'***Implementation***

'Tracks the rows on the sheet and the cells that they start at
Private rowNames As Collection
Private rowStartingCells As Collection

'Purpose: fills the cell at starting value (Fill Length - 1) cells to its right
'           with a value of Starting Value and increasing by Fill Increment for each cell already filled
'Effects: may change the value of the cell at Starting Cell and (Fill Length - 1) cells to its right
'Requires: Starting Cell must be the location of only one cell
'          Starting Value and Fill Increment must be valid doubles
'          Fill Length must be a valid integer >= 1
Private Sub btnClear_Click()
    On Error Resume Next

    If lsbSelectedRow.ListIndex < 0 Then
        MsgBox ("No row selected, make sure to select a row by clicking on it")
        Exit Sub
    End If
    
    Call ExcelIO_startEditing
    Call ExcelIO_setRowValue(rowStartingCells(lsbSelectedRow.ListIndex + 1), "")
    Call ExcelIO_stopEditing
End Sub

'Purpose: fills the cell at the starting cell of the selected row and (Fill Length - 1) cells to its right
'           with a value of Starting Value and increasing by Fill Increment for each cell already filled
'Effects: may change the value of the cell at the starting cell of the selected row and (Fill Length - 1) cells to its right
'Requires: a row must be selected
'          Starting Value and Fill Increment must be valid doubles
'          Fill Length must be a valid integer >= 1
Private Sub btnFill_Click()
    On Error Resume Next

    If lsbSelectedRow.ListIndex < 0 Then
        MsgBox ("No row selected, make sure to select a row by clicking on it")
        Exit Sub
    End If
    
    If (Not (IsNumeric(txtStartValue.Text))) Then
        MsgBox ("Invaid starting value, ensure that it is compleatly numeric (ie 12.34)")
        Exit Sub
    End If
    If (Not (IsNumeric(txtLength.Text))) Then
        MsgBox ("Invaid fill length, ensure that it is compleatly numeric (ie 12)")
        Exit Sub
    End If
    If (Not (IsNumeric(txtIncrement.Text))) Then
        MsgBox ("Invaid fill increment, ensure that it is compleatly numeric (ie 12.34), put in 0 if no increment")
        Exit Sub
    End If
    
    Call ExcelIO_startEditing
    Dim startCell As Range
    Set startCell = Range(rowStartingCells(lsbSelectedRow.ListIndex + 1))
    
    Call ExcelIO_setRowValue(startCell.Address, "")
    
    Dim startValue As Double: startValue = CDbl(txtStartValue.Text)
    Dim length As Integer: length = CInt(txtLength.Text)
    Dim increment As Double: increment = CDbl(txtIncrement.Text)
    
    Dim i As Integer: i = 1
    For i = 0 To length - 1
        Cells(startCell.Row, startCell.Column + i).Value2 = startValue + increment * i
    Next
    
    Call ExcelIO_stopEditing
End Sub

'Purpose: initalizes the rows that the userform recognizes, required to be called before the userform is shown
'Effects: rowStartingCells and rowNames have their values reinitalized
Public Sub initalizeRows()
    Set rowStartingCells = New Collection
    Set rowNames = New Collection
End Sub

'Purpose: causes the userform to recognize a new row with a name of newRowName starting at newRowStartingCell
'Effects: rowStartingCells and rowNames may have their values changed
'Requires: newRowStartingCell must be the location of a (only one) vaild cell
'          there are no other rows currently recognized by the userform with the name of newRowName
Public Sub addRow(ByVal newRowName As String, ByVal newRowStartingCell As String)
    Dim rowIterator As Variant
    For Each rowIterator In rowNames
        If (StrComp(rowIterator, newRowName) = 0) Then
            MsgBox ("Row already in userform!")
            Exit Sub
        End If
    Next
    rowNames.Add (newRowName)
    rowStartingCells.Add (newRowStartingCell)
    lsbSelectedRow.AddItem (newRowName)
End Sub

'Purpose: causes the userform to forget a currently recognized row with a name of oldRowName
'Effects: rowStartingCells and rowNames may have their values changed
'Requires: there must be a row currently recognized by the userform with the name of oldRowName
Public Sub deleteRow(ByVal oldRowName As String)
    Dim rowIndex As Integer
    For rowIndex = 1 To rowNames.Count
        If (StrComp(rowNames(rowIndex), oldRowName) = 0) Then
            rowNames.Remove (rowIndex)
            rowStartingCells.Remove (rowIndex)
            lsbSelectedRow.RemoveItem (oldRowName)
            Exit Sub
        End If
    Next
    MsgBox ("Row not in userform!")
End Sub
