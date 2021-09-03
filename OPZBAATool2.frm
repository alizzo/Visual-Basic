VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} OPZBAATool 
   Caption         =   "Nuvance Health - OP ZBAA"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   OleObjectBlob   =   "OPZBAATool.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "OPZBAATool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub Cancel_Click()
Unload Me
End Sub

Private Sub ComboBox_Change()

End Sub

Private Sub UserForm_Initialize()

With ComboBox


    .AddItem "NDH"
    .AddItem "PHC"
    .AddItem "SH"
    .AddItem "VBMC"
    
End With
End Sub



Private Sub Continue_Click()


'
' OPZBAA Macro
'

'
  
   

    Dim DropDown As String
    
    DropDown = ComboBox.Value


    

'Crosswalk to the payor mix
    Dim crossWalk As Workbook
    
    ' if the file name of the payor mix cross walk changes please update the following line
    Dim crossWalkFileName As String
    crossWalkFileName = "File.xlsx"

    Dim crossWalkNamePath As String
    crossWalkNamePath = "C:\user\Desktop\" & crossWalkFileName
  

'PATH crosswalk to the ZBAA
    Dim opZBAAcrossWalknamePath As String
    opZBAAcrossWalknamePath = "C:\user\Desktop\"
    
' File names of ZBAA Crosswalks
    Dim NDHopZBAAcrossWalkFileName As String
    NDHopZBAAcrossWalkFileName = "HospitalCrosswalk1.xlsx"
    
    Dim VBMCopZBAAcrossWalkFileName As String
    VBMCopZBAAcrossWalkFileName = "HospitalCrosswalk2.xlsx"
    
    Dim PHCopZBAAcrossWalkFileName As String
    PHCopZBAAcrossWalkFileName = "HospitalCrosswalk3.xlsx"


    Dim SHopZBAAcrossWalkFileName As String
    SHopZBAAcrossWalkFileName = "HospitalCrosswalk4.xlsx"



' Declare a variable to store the crosswalk file to be executed in the vlookup
    Dim ZBAAfileReference As String

    'HorizontalCrossWalkReference is a reference to the column in the crossWalkFileName crosswalk
    Dim HorizontalCrossWalkReference As Variant

    ' Declare the end of the data set & store to variable "LastRow"
    Dim i As Long
    Dim LastRow As Long
    LastRow = Range("A" & Rows.Count).End(xlUp).Row
   
   ' Stores the value of the conditional path baed on users input into "DropDown"
    Dim ZBAAFileAndPath As String

'clears the pop up window after the user has selected a choice from the combo (Could run it blank - error checking might be needed...)
    Unload Me


    







'Formatting & Delete unnecessary columns
    Columns("A:F").Select
    Selection.Delete Shift:=xlToLeft
    Cells.Select
    Cells.EntireColumn.AutoFit
    Columns("F:G").Select
    Selection.Delete Shift:=xlToLeft
    Columns("G:H").Select
    Selection.Delete Shift:=xlToLeft
    Columns("G:G").Select
    Selection.Delete Shift:=xlToLeft
    Columns("H:I").Select
    Selection.Delete Shift:=xlToLeft
    Columns("J:J").Select
    Selection.Delete Shift:=xlToLeft
    Columns("L:M").Select
    Selection.Delete Shift:=xlToLeft
    Columns("M:R").Select
    Selection.Delete Shift:=xlToLeft
    
    
    Columns("M:R").Select
    Selection.Delete Shift:=xlToLeft
    
    

    Columns("P:Q").Select
    Selection.Delete Shift:=xlToLeft
    
    Columns("R:R").Select
    Range("R1").Value = "Payor Mix"

    Range(Selection, Selection.End(xlToLeft)).Select
    Selection.AutoFilter
    
    
    Columns("F:F").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("F1").Value = "Concatenate"
    Range("F2").Select
    Application.CutCopyMode = False

    
    ActiveCell.FormulaR1C1 = "=RC[-1]&RC[9]"
    
    
    
    
    Range("F2").Select
    Selection.AutoFill Destination:=Range("F2:F" & LastRow)
    Range("F2:F" & LastRow).Select
    Columns("F:F").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    
    Columns("G:G").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
    Range("G1").Value = "Revenue Test Category"
  



    ' Name the active sheet "OP Cerner Charges"
    ActiveSheet.Select
    ActiveSheet.Name = "OP Cerner Charges"
    Sheets("OP Cerner Charges").Select



    
    
            If DropDown = "NDH" Then
                HorizontalCrossWalkReference = 10
     
                ZBAAFileAndPath = opZBAAcrossWalknamePath & NDHopZBAAcrossWalkFileName
                Workbooks.Open Filename:=ZBAAFileAndPath
                ActiveWindow.WindowState = xlMinimized
                ActiveWindow.WindowState = xlMaximized
                

                ZBAAfileReference = NDHopZBAAcrossWalkFileName
                
                
            ElseIf DropDown = "VBMC" Then
                HorizontalCrossWalkReference = 11
               
                ZBAAFileAndPath = opZBAAcrossWalknamePath & VBMCopZBAAcrossWalkFileName
                Workbooks.Open Filename:=ZBAAFileAndPath
                ActiveWindow.WindowState = xlMinimized
                ActiveWindow.WindowState = xlMaximized
                   
                ZBAAfileReference = VBMCopZBAAcrossWalkFileName
                    
                    
            ElseIf DropDown = "SH" Then
                HorizontalCrossWalkReference = 13
             
                ZBAAFileAndPath = opZBAAcrossWalknamePath & SHopZBAAcrossWalkFileName
                Workbooks.Open Filename:=ZBAAFileAndPath
                ActiveWindow.WindowState = xlMinimized
                ActiveWindow.WindowState = xlMaximized
                    
                ZBAAfileReference = SHopZBAAcrossWalkFileName

                
            ElseIf DropDown = "PHC" Then
                HorizontalCrossWalkReference = 12
            
                ZBAAFileAndPath = opZBAAcrossWalknamePath & PHCopZBAAcrossWalkFileName
                Workbooks.Open Filename:=ZBAAFileAndPath
                ActiveWindow.WindowState = xlMinimized
                ActiveWindow.WindowState = xlMaximized
                    
                ZBAAfileReference = PHCopZBAAcrossWalkFileName
                    
            
                
            End If
    
    
    
 
    

    
        For i = 2 To LastRow
             On Error Resume Next
             Range("G" & i).Value = Application.WorksheetFunction.VLookup(Worksheets("OP Cerner Charges").Range("F" & i).Value, _
             Workbooks(ZBAAfileReference).Worksheets("OP ZBAA Crosswalk").Range("A:B"), 2, 0)
             On Error GoTo 0
              
            If Range("G" & i).Value = 0 Then
                  Range("G" & i) = "#N/A"
              End If
        Next i
                

   ' Close Service crosswalk
   Workbooks(ZBAAfileReference).Close
   
    
    
   ' Payor Mix Crosswalk
    Workbooks.Open Filename:=crossWalkNamePath
    ActiveWindow.WindowState = xlMinimized
    ActiveWindow.WindowState = xlMaximized
    
    
     Application.ScreenUpdating = False
    
     For i = 2 To LastRow
         On Error Resume Next
         Range("T" & i).Value = Application.WorksheetFunction.VLookup(Worksheets("OP Cerner Charges").Range("N" & i).Value, _
         Workbooks(crossWalkFileName).Worksheets("HCRA crosswalk").Range("D:P"), HorizontalCrossWalkReference, 0)
         On Error GoTo 0
              
         If Range("T" & i).Value = 0 Then
           Range("T" & i) = "#N/A"
        End If
    Next i
                
    Application.ScreenUpdating = True

   Workbooks(crossWalkFileName).Close



' remove duplicates
   ActiveSheet.Range("$A$1:$T" & LastRow).RemoveDuplicates Columns:=2, Header:=xlYes






    

'changes the tab names and creates a new copy with name :" OP ZBAA"
    ActiveSheet.Copy After:=Worksheets(Sheets.Count)
    On Error Resume Next
    ActiveSheet.Name = "OP ZBAA"


'Payments
Dim rng As Range

'header
Range("J1").Value = "Payments"

'invert payment value
Set rng = Range("J2:J" & LastRow)
rng = Evaluate(rng.Address & "*-1")
'rng = rng.NumberFormat = "#,##0"



   Application.ScreenUpdating = False

' deletes entire rows if the payment value is zero (well between .01 and -.01 cent)
For r = LastRow To 2 Step -1
If (Cells(r, 10) > -0.01) And (Cells(r, 10) < 0.01) Then
Rows(r).Delete
End If
Next r



'Encounter Balances between -100 and 100



    Application.Calculation = xlCalculationManual
    Dim y As Long
    For y = Range("M" & Rows.Count).End(xlUp).Row To 2 Step -1
        If Not (Range("M" & y).Value >= -100.001 And Range("M" & y).Value <= 100.001) Then
            Range("M" & y).EntireRow.Delete
        End If
    Next y
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True







'pivot table

'Declare Variables
Dim PSheet As Worksheet
Dim DSheet As Worksheet
Dim PCache As PivotCache
Dim PTable As PivotTable
Dim PRange As Range
Dim LastRow1 As Long
Dim LastCol As Long

'Insert a New Blank Worksheet
On Error Resume Next
Application.DisplayAlerts = False
Worksheets("PivotTable").Delete
Sheets.Add After:=ActiveSheet
ActiveSheet.Name = "PivotTable"
Application.DisplayAlerts = True
Set PSheet = Worksheets("PivotTable")
Set DSheet = Worksheets("OP ZBAA")

'Define Data Range
LastRow1 = DSheet.Cells(Rows.Count, 1).End(xlUp).Row
LastCol = DSheet.Cells(1, Columns.Count).End(xlToLeft).Column
Set PRange = DSheet.Cells(1, 1).Resize(LastRow1, LastCol)

'Define Pivot Cache
Set PCache = ActiveWorkbook.PivotCaches.Create _
(SourceType:=xlDatabase, SourceData:=PRange). _
CreatePivotTable(TableDestination:=PSheet.Cells(2, 2), _
TableName:="OP ZBAA Table")

'Insert Blank Pivot Table
Set PTable = PCache.CreatePivotTable _
(TableDestination:=PSheet.Cells(1, 1), TableName:="OP ZBAA Table")

'Insert Row Fields
With ActiveSheet.PivotTables("OP ZBAA Table").PivotFields("Revenue Test Category")
.Orientation = xlRowField
.Position = 1
End With

' potential second row field
'With ActiveSheet.PivotTables("OP ZBAA").PivotFields("Total Charges")
'.Orientation = xlRowField
'.Position = 2
'End With


    ActiveSheet.PivotTables("OP ZBAA Table").AddDataField ActiveSheet.PivotTables( _
        "OP ZBAA Table").PivotFields("Encounter"), "Count of Encounter", xlCount
        

    ActiveSheet.PivotTables("OP ZBAA Table").AddDataField ActiveSheet.PivotTables( _
        "OP ZBAA Table").PivotFields("Total Charges"), "Sum of Total Charges", xlSum
        
    ActiveSheet.PivotTables("OP ZBAA Table").AddDataField ActiveSheet.PivotTables( _
        "OP ZBAA Table").PivotFields("Total Adjustments"), "Sum of Adjustments", xlSum
        
    ActiveSheet.PivotTables("OP ZBAA Table").AddDataField ActiveSheet.PivotTables( _
        "OP ZBAA Table").PivotFields("Payments"), "Sum of Payments", xlSum
    
    
    ActiveSheet.PivotTables("OP ZBAA Table").CalculatedFields.Add "ZBAA", _
        "=Payments /'Total Charges'", True
    ActiveSheet.PivotTables("OP ZBAA Table").PivotFields("ZBAA").Orientation = _
        xlDataField









'Insert Column Fields
'With ActiveSheet.PivotTables("OP ZBAA").PivotFields("Total Charges")
'.Orientation = xlColumnField
'.Position = 1
'End With

'Insert Data Field
'With ActiveSheet.PivotTables("OP ZBAA")
'.PivotFields ("Total Charges")
'.Orientation = xlDataField
'.Function = xlSum
'.NumberFormat = "#,##0"
'.Name = "Charges"
'End With

'Format Pivot Table
ActiveSheet.PivotTables("OP ZBAA").ShowTableStyleRowStripes = True
ActiveSheet.PivotTables("OP ZBAA").TableStyle2 = "PivotStyleMedium9"










'changes the tab names and creates a new copy with name :" ZBAA Bal(-100<x<100)"
'   Worksheets("OP Cerner Charges").Copy After:=Worksheets(Sheets.Count)
'   On Error Resume Next
'   ActiveSheet.Name = "ZBAA Bal(-2500<x<2500)"


'   Dim NewLastRow As Long
'    NewLastRow = Range("B" & Rows.Count).End(xlUp).Row


' deletes entire rows if the payment value is zero (well between .01 and -.01 cent)
'For r = NewLastRow To 2 Step -1
'If (Cells(r, 10) > -0.01) And (Cells(r, 10) < 0.01) Then
'Rows(r).Delete
'End If
'Next r



'Encounter Balances between -100 and 100

   ' Application.Calculation = xlCalculationManual
   ' Dim A As Long
   ' For A = Range("M" & Rows.Count).End(xlUp).Row To 2 Step -1
   '     If Not (Range("M" & A).Value >= -2500.001 And Range("M" & A).Value <= 2500.001) Then
   '         Range("M" & A).EntireRow.Delete
   '     End If
   ' Next A
   ' Application.Calculation = xlCalculationAutomatic
   










End Sub






