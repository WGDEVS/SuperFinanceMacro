Attribute VB_Name = "FindMissing"
' FindMissing module
' This module provides the function to find the missing value of a cell

'''***Interface***

'sub: FindMissing_FindMissingInput(inputCells, outputCell, expectedOutput)
'Purpose: changes the cell with the missing value in inputCells so that the cell at outputCell
'           has the value as the cell at expectedOutput
'Effects: may change the value of the cell with the missing value in inputCells
'Requires: inputCells must be the locations of valid cells,
'            exactly one of which must have a missing value (value = ?)
'          outputCell and expectedOutput must be the location of a (only one) vaild cell

'''***Implementation***

'Purpose: changes the cell with the missing value in inputCells so that the cell at outputCell
'           has the same value as the cell at expectedOutput
'Effects: may change the value of the cell with the missing value in inputCells
'Requires: inputCells must be the locations of valid cells,
'            exactly one of which must have a missing value (value = ?)
'          outputCell and expectedOutput must be the location of a (only one) vaild cell
Sub FindMissing_FindMissingInput(ByVal inputCells As String, ByVal outputCell As String, ByVal expectedOutput As String)
    Call ExcelIO_startEditing
    
    Dim inputCell As Range
    Dim inputCellCount As Integer
    
    'Search the inputCells for missing values
    Dim tempCell As Range
    For Each tempCell In Range(inputCells)
        If Not IsNumeric(tempCell.Value2) Then
            Set inputCell = tempCell
            inputCellCount = inputCellCount + 1
        End If
    Next
    
    'Give an error message if there is no missing value or multiple missing values
    If inputCellCount = 0 Then
        MsgBox ("No missing value found, (mark missing values with ?)")
    ElseIf inputCellCount > 1 Then
        MsgBox ("Too many missing values, re-read the question")
    Else
        inputCell.Value2 = 0
        Range(outputCell).GoalSeek Goal:=Range(expectedOutput).Value, ChangingCell:=inputCell
    End If
    
    Call ExcelIO_stopEditing
End Sub
