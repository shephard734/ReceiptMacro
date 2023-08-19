Attribute VB_Name = "Module1"
Sub importlastweek()
'   @Author Spencer Shephard
'   @Version 1.0
'   This macro imports Confirmation (Conf') and Notes data from a receipts workbook pre-transformed by Kristen Lukasik.
'   Without these transformations, it will produce undefined behavior.
'   This macro will overwrite most things in the D and E columns of any worksheet it is applied to.
'   Loss of data is possible when overwriting. Use at your own risk.

'   In future versions, I hope to include:
'       -The carrying over of formatting, namely color
'       -The performance of the aforementioned transformations currently done manually by Kristen.


'    Select last week's file using a file browse dialog box.
    Dim destwb As Workbook
    Dim sourcewb As Workbook
    Set destwb = ActiveWorkbook
    
    Dim fdialog As Office.FileDialog
    Set fdialog = Application.FileDialog(msoFileDialogFilePicker)
    
    With fdialog
        .AllowMultiSelect = False
        .Title = "Please select the workbook from which to import notes."
        .Filters.Clear
        .Filters.Add "Excel Workbooks", "*.xls; *.xlsx; *.xlsm", 1
    
    If .Show = True Then
        Set sourcewb = Workbooks.Open(fdialog.SelectedItems(1))
    Else
        MsgBox "A workbook must be selected to import data. Program exiting."
        Exit Sub
    End If
    End With
    
'    Copy last week's data to this week as a new sheet.
    Dim sh As Worksheet
    Set sh = sourcewb.Sheets(1)
    sh.Copy After:=destwb.Sheets("Sheet1")
    
'    Define and name a range in the copied sheet, including PO#, Conf', and Notes columns. Also includes Line# column, which is unused.
    destwb.Sheets(2).Activate
    Range("B65536").End(xlUp).Select
    intbottomrow = ActiveCell.Row
    Range("B7:E" & intbottomrow).Select
    Application.Goto Reference:="lastweek"
    
'   Define the name to be used in the upcoming vlookup formula, including reference to its scope (which is its own worksheet).
    Dim shname As String
    shname = "'" & ActiveSheet.Name & "'!lastweek"
    
'   set formulae on this week's sheet to vlookup data from copied sheet.
    destwb.Sheets("Sheet1").Activate
    Range("B65536").End(xlUp).Select
    intbottomrow = ActiveCell.Row
    Range("D7:D" & intbottomrow).FormulaR1C1 = "=IFNA(IF(VLOOKUP(RC[-2]," & shname & ",3,FALSE)=0,"""",VLOOKUP(RC[-2]," & shname & ",3,FALSE)),"""")"
    Range("E7:E" & intbottomrow).FormulaR1C1 = "=IFNA(IF(VLOOKUP(RC[-3]," & shname & ",4,FALSE)=0,"""",VLOOKUP(RC[-3]," & shname & ",4,FALSE)),"""")"

'    Convert new values from formula to value. This is so we can delete the imported sheet afterwards to clean up our mess.
    Dim Cell_Value As Range
    For Each Cell_Value In Range("D7:E" & intbottomrow)
    If Cell_Value.HasFormula Then
    Cell_Value.Formula = Cell_Value.Value
    End If
    Next Cell_Value
    
'   Cleanup copied sheet and opened workbook. Suppress confirmation message, then re-enable.
    Application.DisplayAlerts = False
    destwb.Sheets(2).Delete
    sourcewb.Close
    Application.DisplayAlerts = True
    
    
End Sub
