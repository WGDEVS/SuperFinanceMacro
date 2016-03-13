Attribute VB_Name = "ExcelIO"
' ExcelIO module
' This module provides the functions and subroutines to read and write from excel cells

'''***Interface***

'Sub: ExcelIO_startEditing()
'Purpose: prepares the active excel sheet for editing, make sure that the ExcelIO_stopEditing()
'           subroutine is called after you are done and before this function is called again
'Effects: imporves the speed at which data is written to the excel sheet by disabling the updating of the display
'Requires: if this subroutine was called before, stopEditing() must be called before this subroutne

'Sub: ExcelIO_stopEditing()
'Purpose: prepares the active excel sheet for displaying, make sure that the ExcelIO_startEditing()
'           subroutine is called before this function is called
'Effects: allows the display of the active excel sheet to update again
'Requires: ExcelIO_startEditing() must be called before this subroutine is called

'Function: ExcelIO_getChecked(controlName)
'Purpose: returns a boolean representing if the control on the active excel
'         sheet with the name controlName is checked
'Requires: controlName must be the name of a vaild control with the checked property

'Sub: ExcelIO_setRowValue(startingPoint, newValue)
'Purpose: sets the value of the cell at startingPoint and any cell to
'           its right to newValue
'Effects: changes the value of the cells at and to the right of cellLocation
'Requires: startingPoint must be the location of a (only one) vaild cell

'Function: ExcelIO_getCellBackColor(cellLocation)
'Purpose: returns a number representing the background color of the cell at
'          cellLocation on the active excel sheet
'Requires: cellLocation must be the location of a (only one) vaild cell

'Sub: ExcelIO_setRowBackColor(startingPoint, color)
'Purpose: sets the background color of the cell at startingPoint and any cell to
'           its right to the color represented by color
'Effects: changes the background color of the cells at and to the right of cellLocation
'Requires: startingPoint must be the location of a (only one) vaild cell

'Sub: ExcelIO_setCellBackColor
'Purpose: sets the background color of the cells at cellLocation to the color
'           represented by color
'Effects: changes the background color of the cells at cellLocation
'Requires: cellLocation must be the location of vaild cells

'''***Implementation***

Private Const EXCEL_SHEET_WIDTH = 16372

' State of the display that the user has set before any function in this module was called
Private screenUpdateState As Boolean
Private statusBarState As Boolean
Private eventsState As Boolean
Private displayPageBreakState As Boolean

'Purpose: prepares the active excel sheet for editing, make sure that the ExcelIO_stopEditing()
'           subroutine is called after you are done and before this function is called again
'Effects: imporves the speed at which data is written to the excel sheet by disabling the updating of the display
'Requires: if this subroutine was called before, stopEditing() must be called before this subroutne
'          is called again
Sub ExcelIO_startEditing()
    screenUpdateState = Application.ScreenUpdating
    statusBarState = Application.DisplayStatusBar
    displayPageBreakState = ActiveSheet.DisplayPageBreaks
    eventsState = Application.EnableEvents

    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.EnableEvents = False
    ActiveSheet.DisplayPageBreaks = False
End Sub

'Purpose: prepares the active excel sheet for displaying, make sure that the ExcelIO_startEditing()
'           subroutine is called before this function is called
'Effects: allows the display of the active excel sheet to update again
'Requires: ExcelIO_startEditing() must be called before this subroutine is called
Sub ExcelIO_stopEditing()
    Application.ScreenUpdating = screenUpdateState
    Application.DisplayStatusBar = statusBarState
    Application.EnableEvents = eventsState
    ActiveSheet.DisplayPageBreaks = displayPageBreaksState
End Sub

'Purpose: returns a boolean representing if the control on the active excel
'         sheet with the name controlName is checked
'Requires: controlName must be the name of a vaild control with the checked property
Function ExcelIO_getChecked(ByVal controlName As String) As Boolean
    getChecked = (ActiveSheet.Shapes(controlName).OLEFormat.Object.Value = 1)
End Function

'Purpose: sets the value of the cell at startingPoint and any cell to
'           its right to newValue
'Effects: changes the value of the cells at and to the right of cellLocation
'Requires: startingPoint must be the location of a (only one) vaild cell
Sub ExcelIO_setRowValue(ByVal startingPoint As String, ByVal newValue As String)
    Range(startingPoint, Cells(Range(startingPoint).Row, EXCEL_SHEET_WIDTH)).Value2 = ""
End Sub

'Purpose: returns a number representing the background color of the cell at
'          cellLocation on the active excel sheet
'Requires: cellLocation must be the location of a (only one) vaild cell
Function ExcelIO_getCellBackColor(ByVal cellLocation As String) As Long
    getCellBackColor = Range(cellLocation).Interior.color
End Function

'Purpose: sets the background color of the cell at startingPoint and any cell to
'           its right to the color represented by color
'Effects: changes the background color of the cells at and to the right of startingPoint
'Requires: startingPoint must be the location of a (only one) vaild cell
Sub ExcelIO_setRowBackColor(ByVal startingPoint As String, ByVal color As Long)
   Range(startingPoint, Cells(Range(startingPoint).Row, EXCEL_SHEET_WIDTH)).Interior.color = color
End Sub

'Purpose: sets the background color of the cells at cellLocation to the color
'           represented by color
'Effects: changes the background color of the cells at cellLocation
'Requires: cellLocation must be the location of vaild cells
Sub ExcelIO_setCellBackColor(ByVal cellLocation As String, ByVal color As Long)
    Range(cellLocation).Interior.color = color
End Sub

