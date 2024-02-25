' Require all variables to be declared explicitly
    Option Explicit

' Declare some public variables for this file and path
    Public FilePath As String
    Public ThisFileName As String

'Declare some public variables to hold the file names used with the File DilogBox
    Public RawErrors_File As String
    Public RawPSI_File As String
    Public RawTU_File As String
    Public RawMFH_File As String
    Public RawTransaction_File As String
    Public AppendMFH_File As String

	
Sub Refresh_Data()
    
' Get the path and file name for this file
    FilePath = ThisWorkbook.path & "\"
    ThisFileName = ThisWorkbook.Name
    
' Turn off screen updating and alerts to speed up the processing
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
' Unhide the calculated data sheets
    Sheets("C - TU").Visible = True
    Sheets("C - MFH").Visible = True
    Sheets("C - Error").Visible = True
    Sheets("C - Transaction").Visible = True
    Sheets("C - PSI").Visible = True
    
       
' Call the routine to append the Historical tab with the latest summary data
    Update_Historical_Data
    
' Call the routine to open the new raw data files
    Open_Raw_Data_Files

' Call the routine to update the calculated sheets
    Build_TU_Calc_Sheet
    Build_MFH_Calc_Sheet
    Build_Error_Calc_Sheet
    Build_Transaction_Calc_Sheet
    Build_PSI_Calc_Sheet
'    UpdateData_A
    
' Close the raw data workbooks
     Workbooks(RawTU_File).Close
     Workbooks(RawErrors_File).Close
     Workbooks(RawMFH_File).Close
     Workbooks(RawTransaction_File).Close
     Workbooks(RawPSI_File).Close

' Check if there is a second MFH File and append the C -MFH sheet if needed.
    Append_MFH_Calc_Sheet
    
' Hide the calculated data sheets
    Sheets("C - TU").Visible = False
    Sheets("C - MFH").Visible = False
    Sheets("C - Error").Visible = False
    Sheets("C - Transaction").Visible = False
    Sheets("C - PSI").Visible = False
    
    
' Shift focus to the Summary Sheet
    Sheets("Summary").Select
    Range("A1").Select
    
' Turn Screen Updating and alerts back on
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

  
End Sub
Sub Open_Raw_Data_Files()

'
' Open_RawTU_A Macro
        
' Start in the folder this file is saved in
    ChDir FilePath
    
'Promt user for the Raw TU file and then open the file
    RawTU_File = Application.GetOpenFilename _
    (Title:="Please Select the TU Raw File", _
    FileFilter:="SPOC Exported Files *.csv (*.csv),")
    'Here you can note that we have allowed only *.csv excel files to choose
    'We have also customized the file dialog title

    Workbooks.Open FileName:=RawTU_File
    RawTU_File = ActiveWorkbook.Name

'Promt user for the Raw MFH file and then open the file
    RawMFH_File = Application.GetOpenFilename _
    (Title:="Please Select the MFH Raw file", _
    FileFilter:="SPOC Exported Files *.csv (*.csv),")
    'Here you can note that we have allowed only *.csv excel files to choose
    'We have also customized the file dialog title

    Workbooks.Open FileName:=RawMFH_File
    RawMFH_File = ActiveWorkbook.Name
    
'Promt user for the Raw Errors file and then open the file
    RawErrors_File = Application.GetOpenFilename _
    (Title:="Please Select the Errors Raw file", _
    FileFilter:="SPOC Exported Files *.csv (*.csv),")
    'Here you can note that we have allowed only *.csv excel files to choose
    'We have also customized the file dialog title

    Workbooks.Open FileName:=RawErrors_File
    RawErrors_File = ActiveWorkbook.Name
    
'Promt user for the Raw Transaction file and then open the file
    RawTransaction_File = Application.GetOpenFilename _
    (Title:="Please Select the Transaction Raw file", _
    FileFilter:="SPOC Exported Files *.csv (*.csv),")
    'Here you can note that we have allowed only *.csv excel files to choose
    'We have also customized the file dialog title

    Workbooks.Open FileName:=RawTransaction_File
    RawTransaction_File = ActiveWorkbook.Name
    
'Promt user for the Raw PSI file and then open the file
    RawPSI_File = Application.GetOpenFilename _
    (Title:="Please Select the PSI Raw file", _
    FileFilter:="SPOC Exported Files *.csv (*.csv),")
    'Here you can note that we have allowed only *.csv excel files to choose
    'We have also customized the file dialog title

    Workbooks.Open FileName:=RawPSI_File
    RawPSI_File = ActiveWorkbook.Name

End Sub
Sub Update_Historical_Data()

' Create a new column on the Historical sheet and copy the last values into it
    Windows(ThisFileName).Activate
    Sheets("Historical").Select
    Columns("B:B").Select
    Selection.Copy
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight
    Range("C1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B1").Select
    
End Sub


Sub Build_TU_Calc_Sheet()

    Dim LastRow As Long
    
' Find the last row in the C-TU sheet and then clear the data from row 3 on
    Windows(ThisFileName).Activate
    Sheets("C - TU").Select
    LastRow = ActiveSheet.UsedRange.Rows.Count
    Rows("3:" & LastRow).Select
    Selection.Delete Shift:=xlUp
    
 ' Copy the data from the Raw TU sheet into the C-TU sheet
    Windows(RawTU_File).Activate
    Range("A9").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.Copy
    
    Windows(ThisFileName).Activate
    Sheets("C - TU").Select
    Range("A1").Select
    ActiveSheet.Paste
    
' Fill in the calculated columns to the last used row
    LastRow = ActiveSheet.UsedRange.Rows.Count
    Range("U2:AD2").Select
    Selection.AutoFill Destination:=Range("U2:AD" & LastRow)

' Bold the header row and return focus to A1
    Rows("1:1").Select
    Selection.Font.Bold = True
    Range("A1").Select

End Sub


Sub Build_MFH_Calc_Sheet()

    Dim LastRow As Long
    
' Find the last row in the C-TU sheet and then clear the data from row 3 on
    Windows(ThisFileName).Activate
    Sheets("C - MFH").Select
    LastRow = ActiveSheet.UsedRange.Rows.Count
    Rows("3:" & LastRow).Select
    Selection.Delete Shift:=xlUp
    
 ' Copy the data from the Raw TU sheet into the C-TU sheet
    Windows(RawMFH_File).Activate
    Range("A9").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.Copy
    
    Windows(ThisFileName).Activate
    Sheets("C - MFH").Select
    Range("A1").Select
    ActiveSheet.Paste
    
' Add the non-formula value to cell L1
    Range("L1").Value = "SEQ"
    
' Fill in the calculated columns to the last used row
    LastRow = ActiveSheet.UsedRange.Rows.Count
    Range("N2:AC2").Select
    Selection.AutoFill Destination:=Range("N2:AC" & LastRow)

' Bold the header row and return focus to A1
    Rows("1:1").Select
    Selection.Font.Bold = True
    Range("A1").Select

End Sub


Sub Build_Error_Calc_Sheet()

    Dim LastRow As Long
    
' Find the last row in the C-TU sheet and then clear the data from row 3 on
    Windows(ThisFileName).Activate
    Sheets("C - Error").Select
    LastRow = ActiveSheet.UsedRange.Rows.Count
    Rows("3:" & LastRow).Select
    Selection.Delete Shift:=xlUp
    
 ' Copy the data from the Raw TU sheet into the C-TU sheet
    Windows(RawErrors_File).Activate
    Range("A9").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.Copy
    
    Windows(ThisFileName).Activate
    Sheets("C - Error").Select
    Range("A1").Select
    ActiveSheet.Paste
    
' Add the non-formula value to cell L1
    Range("G1").Value = "SEQ"
    
' Fill in the calculated columns to the last used row
    LastRow = ActiveSheet.UsedRange.Rows.Count
    Range("I2:Y2").Select
    Selection.AutoFill Destination:=Range("I2:Y" & LastRow)

' Bold the header row and return focus to A1
    Rows("1:1").Select
    Selection.Font.Bold = True
    Range("A1").Select

End Sub


Sub Build_Transaction_Calc_Sheet()

    Dim LastRow As Long
    
' Find the last row in the C-TU sheet and then clear the data from row 3 on
    Windows(ThisFileName).Activate
    Sheets("C - Transaction").Select
    LastRow = ActiveSheet.UsedRange.Rows.Count
    Rows("3:" & LastRow).Select
    Selection.Delete Shift:=xlUp
    
 ' Copy the data from the Raw TU sheet into the C-TU sheet
    Windows(RawTransaction_File).Activate
    Range("A9").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.Copy
    
    Windows(ThisFileName).Activate
    Sheets("C - Transaction").Select
    Range("A1").Select
    ActiveSheet.Paste
    
' Clear the data from column N, label it, and highlight it yellow
    Columns("N:N").Select
    Selection.ClearContents
    Range("N1").Value = "Blank"
    
    Columns("N:N").Select
    With Selection.Interior
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
' Fill in the calculated columns to the last used row
    LastRow = ActiveSheet.UsedRange.Rows.Count
    Range("O2:AA2").Select
    Selection.AutoFill Destination:=Range("O2:AA" & LastRow)

' Bold the header row and return focus to A1
    Rows("1:1").Select
    Selection.Font.Bold = True
    Range("A1").Select

End Sub

Sub Build_PSI_Calc_Sheet()

    Dim LastRow As Long
    
' Find the last row in the C-TU sheet and then clear the data from row 3 on
    Windows(ThisFileName).Activate
    Sheets("C - PSI").Select
    LastRow = ActiveSheet.UsedRange.Rows.Count
    Rows("3:" & LastRow).Select
    Selection.Delete Shift:=xlUp
    
 ' Copy the data from the Raw PSI sheet into the C-PSI sheet
    Windows(RawPSI_File).Activate
    Range("A9").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.Copy
    
    Windows(ThisFileName).Activate
    Sheets("C - PSI").Select
    Range("A1").Select
    ActiveSheet.Paste
    
' Add the non-formula value to cell L1
    Range("H1").Value = "Blank"
    
' Fill in the calculated columns to the last used row
    LastRow = ActiveSheet.UsedRange.Rows.Count
    Range("I2:W2").Select
    Selection.AutoFill Destination:=Range("I2:W" & LastRow)

' Bold the header row and return focus to A1
    Rows("1:1").Select
    Selection.Font.Bold = True
    Range("A1").Select

End Sub

Sub Append_MFH_Calc_Sheet()

    Dim LastRow As Long
    Dim MSG_Box_Answer As Integer
    
'Ask user if there is a second MFH file.  If ues continue.  If no, exit routine
    MSG_Box_Answer = MsgBox("Is there a second MFH file to append the data with?", vbYesNo + vbQuestion, "Append MFH Data?")
    If MSG_Box_Answer = vbNo Then Exit Sub
 
'Promt user for the Append MFH file and then open the file
    AppendMFH_File = Application.GetOpenFilename _
    (Title:="Please Select the 2nd MFH file", _
    FileFilter:="SPOC Exported Files *.csv (*.csv),")
    'Here you can note that we have allowed only *.csv excel files to choose
    'We have also customized the file dialog title

    Workbooks.Open FileName:=AppendMFH_File
    AppendMFH_File = ActiveWorkbook.Name
    
' Find the last row in the C-TU sheet
    Windows(ThisFileName).Activate
    Sheets("C - MFH").Select
    LastRow = ActiveSheet.UsedRange.Rows.Count
    
 ' Copy the data from the Raw TU sheet into the C-TU sheet
    Windows(AppendMFH_File).Activate
    Range("A9").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.Copy
    
    Windows(ThisFileName).Activate
    Sheets("C - MFH").Select
    Range("A" & LastRow + 1).Select
    ActiveSheet.Paste
    
' Fill in the calculated columns to the last used row
    LastRow = ActiveSheet.UsedRange.Rows.Count
    Range("N2:AC2").Select
    Selection.AutoFill Destination:=Range("N2:AC" & LastRow)

' Close the raw AppendMFH workbook
     Workbooks(AppendMFH_File).Close

End Sub