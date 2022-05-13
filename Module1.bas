Attribute VB_Name = "Module1"
' Created by Katie Conlon, Data Analytics Intern. June-August 2019.
' Last updated by Katie Conlon on 7/26/19.

''' Used to aid error debugging.
Option Explicit

''' DECLARE PUBLIC VARIABLES
' Used in subs: getFilePath, openDataFile, shipperUPCLookup, & finalFormatting
Public filepath As String
' Used in subs: collectModes & inconsistentDataEntry
Public sampleNbr() As String
Public modeHeightCase() As Single
Public modeDepthCase() As Single
Public modeWidthCase() As Single
Public modeWeightCase() As Single
Public modeHeightCU() As Single
Public modeDepthCU() As Single
Public modeWidthCU() As Single
Public modeWeightCU() As Single
' Used in sub: inconsistentDataEntry
Public comparisonID As String
' Used in subs: collectUserInput & inconsistentDataEntry
Public devAllowed As Single
' Used in subs: collectUserInput & finalFormatting
Public yearQrt As String
' Used in sub: finalFormatting
Public locationCityState As String
Sub runErrorChecking()
' PURPOSE: Runs all necessary subs.

Call collectUserInput
Call openDataFile
Call deleteRow2
Call deleteRowsWOSampleNbr
Call intialFormatting
Call collectModes
Call decodeabilityShipper
Call decodeabilityCU
Call inconsistentDataEntry
Call missingData
Call shipperUPCLookup
Call gtinCheck
Call wrongProdCode
Call finalFormatting

Range("A1").Select

End Sub
Function getFilePath()
' PURPOSE: Locate the filepath of the current workbook
    
    filepath = ActiveWorkbook.Path
    
End Function
Private Function IsInArray(valToBeFound As Variant, arr As Variant) As Boolean
' PURPOSE: check for SampleNbr position in array.

'DEVELOPER: Ryan Wells (wellsr.com)
'DESCRIPTION: Function to check if a value is in an array of values
'INPUT: Pass the function a value to search for and an array of values of any data type.
'OUTPUT: True if is in array, false otherwise

Dim element As Variant

On Error GoTo IsInArrayError: 'array is empty
    For Each element In arr
        If element = valToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next element
Exit Function

IsInArrayError:
On Error GoTo 0
IsInArray = False

End Function
Private Function WhereInArray(arr1 As Variant, vFind As Variant) As Variant
'PURPOSE: Function to check where a value is in an array
'DEVELOPER: Ryan Wells (wellsr.com)

Dim i As Long

'Loop through all rows in the array until SampleNbr is found.
For i = LBound(arr1) To UBound(arr1)
    
    'If SampleNbr found in array, then collect the row number as variable i
    If arr1(i) = vFind Then
        WhereInArray = i
        Exit Function
    End If
    
Next i

'If SampleNbr isn't found in the array, then set to null.
WhereInArray = Null

End Function
Sub collectUserInput()
' PURPOSE: Collects the amount of deviation allowed and the year/quarter entered by the user.

'' Collect deviation allowed.
' Checks to ensure user completed the user-defined field on the Instructions tab.
If Range("A:A").Find("Deviation Allowed:").Offset(0, 1).Value = "" Then
    
    MsgBox ("You must fill both in the Deviation Allowed and the Year/Quarter.")
    End
    
    Else
    
    devAllowed = Range("A:A").Find("Deviation Allowed:").Offset(0, 1).Value
    
End If

'' Collect Year/Quarter.
' Checks to ensure user completed the user-defined field on the Instructions tab.
If Range("A:A").Find("Year & Quarter:").Offset(0, 1).Value = "" Then
    
    MsgBox ("You must fill both in the Deviation Allowed and the Year/Quarter.")
    End
    
    Else
    
    ' Collect year & quarter from Instructions tab.
    yearQrt = Range("A:A").Find("Year & Quarter:").Offset(0, 1).Value

End If

End Sub
Sub openDataFile()
' PURPOSE: Opens raw data xls file.
    
Call getFilePath
Workbooks.Open (filepath & "\RawFile.xls")

End Sub
Sub deleteRow2()
' PURPOSE: Deletes row 2 (a blank row).

Range("A2").EntireRow.Delete

End Sub
Sub deleteRowsWOSampleNbr()
' PURPOSE: Deletes rows without data based on a blank SampleNbr.

Dim nRows As Integer 'total rows
Dim i As Integer 'row counter

' Calculate number of rows
nRows = Range("H1").CurrentRegion.Rows.Count - 1

' Loops through all rows.
For i = 1 To nRows

    ' Checks if the cell value = "-" (i.e. a blank). If blank, the sub logic applies.
    If Range("E1").Offset(i, 0).Value = "-" Then
    
        ' Delete rows with no SampleNbr (-) indicating we have no data and it should be removed.
        Range("E1").Offset(i, 0).EntireRow.Delete
        
        ' Recalculates number of rows (since some have just been deleted)
        nRows = nRows - 1
        
        ' Resets the row counter back a step (so it doesn't the new row i, since deleting a row advances i+1 to i)
        i = i - 1
         
    End If

    ' Checks if cell value is null (i.e. =""). If so, then stop, because we're at the end of the data.
    If Range("E1").Offset(i, 0).Value = "" Then Exit Sub

Next i
    
End Sub
Sub intialFormatting()
' PURPOSE: Formats the following fields as numbers: Shipper UPC, Consumer Unit UPC, Shipper UPC Decodeability & CU UPC Decodeability. Also formats the very first SampleNbr row of data.
' NOTE: Shipper UPC Decodeability & CU UPC Decodeability formatting is REQUIRED for decodeability sub to work!

''' Sort data by SampleNbr, then Consumer Unit UPC (descending or else it puts blanks to the top for some reason), then Case Counter field. Note: sounds like can't sort more than 3 fields in VBA v2007 or earlier.
Range(Range("A1"), Range("AA1").End(xlDown)).Sort Key1:=Range("E1"), Order1:=xlAscending, Key2:=Range("K1"), Order2:=xlDescending, Key3:=Range("O1"), Order3:=xlAscending, Header:=xlYes

''' Format first SampleNbr row (i.e. row 2), because it gets skipped in formatting logic below. (Note: needs to be earlier in process than the mode deviation check -- otherwise might overwrite the vbRed error callout.)
With Range("A2").EntireRow
    .Interior.Color = vbYellow
    .Font.Bold = True
End With

''' Change to number format.

[G:G].Select
With Selection
    .NumberFormat = "General"
    .Value = .Value
End With

[I:I].Select
With Selection

    ' Fix numbers that have commas or semicolons instead of periods.
    .Replace What:=",", Replacement:=".", _
    lookat:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
    SearchFormat:=False, ReplaceFormat:=False

    .Replace What:=";", Replacement:=".", _
    lookat:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
    SearchFormat:=False, ReplaceFormat:=False

    .NumberFormat = "0.#0"
    .Value = .Value
    
End With

[K:K].Select
With Selection
    .NumberFormat = "General"
    .Value = .Value
End With

[R:R].Select
With Selection

    ' Fix numbers with commas or semicolons instead of periods.
    .Replace What:=",", Replacement:=".", _
    lookat:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
    SearchFormat:=False, ReplaceFormat:=False
    
    .Replace What:=";", Replacement:=".", _
    lookat:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
    SearchFormat:=False, ReplaceFormat:=False
    
    .NumberFormat = "0.#0"
    .Value = .Value
    
End With

[T:AA].Select
With Selection
    
    ' Fix numbers with commas instead of periods.
    .Replace What:=",", Replacement:=".", _
    lookat:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
    SearchFormat:=False, ReplaceFormat:=False
    
    .NumberFormat = "0.#0;###.#0"
    .Value = .Value
    
End With
    
End Sub
Sub decodeabilityShipper()
' PURPOSE: Identifies Shipper UPC decodeability issues

'DECODABILITY GRADE CONVERSION SCALE:
'A >= 62
'B 50 - 61
'C 37 - 49
'D 25 - 36
'F <= 24

Dim nRows As Integer 'total rows
Dim i As Integer 'row counter

' Activate RawFile Workbook, Worksheet Sheet1
Workbooks("RawFile.xls").Activate
Worksheets("Sheet1").Activate

' Calculates number of rows
nRows = Range("H1").CurrentRegion.Rows.Count - 1

' Loops through rows. Highlights & fixes cells with inappropriate Letter Grade.
For i = 1 To nRows
        
    ' GREEN FILL cells that don't have an appropriate Letter Grade or dash (indicating blank).
    If Range("H1").Offset(i, 0) <> "A" And Range("H1").Offset(i, 0) <> "B" And Range("H1").Offset(i, 0) <> "C" And Range("H1").Offset(i, 0) <> "D" And Range("H1").Offset(i, 0) <> "F" And Range("H1").Offset(i, 0) <> "-" Then
        Range("H1").Offset(i, 0).Interior.Color = vbGreen
    End If
        
    ' Check letter grade A against decodeability scale.
    If Range("H1").Offset(i, 0) = "A" And Range("I1").Offset(i, 0) < 0.62 Then
            
            Range("H1").Offset(i, 0).Interior.Color = vbRed
            
            If Range("I1").Offset(i, 0) >= 0.62 Then
                Range("H1").Offset(i, 0).Value = "A"
            ElseIf Range("I1").Offset(i, 0) >= 0.5 And Range("I1").Offset(i, 0) <= 0.61 Then
                Range("H1").Offset(i, 0).Value = "B"
            ElseIf Range("I1").Offset(i, 0) >= 0.37 And Range("I1").Offset(i, 0) <= 0.49 Then
                Range("H1").Offset(i, 0).Value = "C"
            ElseIf Range("I1").Offset(i, 0) >= 0.25 And Range("I1").Offset(i, 0) <= 0.36 Then
                Range("H1").Offset(i, 0).Value = "D"
            ElseIf Range("I1").Offset(i, 0) <= 0.24 Then
                Range("H1").Offset(i, 0).Value = "F"
            Else: Range("H1").Offset(i, 0).Interior.Color = vbGreen
            End If
    
    ' Check letter grade B against decodeability scale.
    ElseIf Range("H1").Offset(i, 0) = "B" And (Range("I1").Offset(i, 0) < 0.5 Or Range("I1").Offset(i, 0) > 0.61) Then
            
            Range("H1").Offset(i, 0).Interior.Color = vbRed
            
            If Range("I1").Offset(i, 0) >= 0.62 Then
                Range("H1").Offset(i, 0).Value = "A"
            ElseIf Range("I1").Offset(i, 0) >= 0.5 And Range("I1").Offset(i, 0) <= 0.61 Then
                Range("H1").Offset(i, 0).Value = "B"
            ElseIf Range("I1").Offset(i, 0) >= 0.37 And Range("I1").Offset(i, 0) <= 0.49 Then
                Range("H1").Offset(i, 0).Value = "C"
            ElseIf Range("I1").Offset(i, 0) >= 0.25 And Range("I1").Offset(i, 0) <= 0.36 Then
                Range("H1").Offset(i, 0).Value = "D"
            ElseIf Range("I1").Offset(i, 0) <= 0.24 Then
                Range("H1").Offset(i, 0).Value = "F"
            Else: Range("H1").Offset(i, 0).Interior.Color = vbGreen
            End If

    ' Check letter grade C against decodeability scale.
    ElseIf Range("H1").Offset(i, 0) = "C" And (Range("I1").Offset(i, 0) < 0.37 Or Range("I1").Offset(i, 0) > 0.49) Then
            
            Range("H1").Offset(i, 0).Interior.Color = vbRed
            
            If Range("I1").Offset(i, 0) >= 0.62 Then
                Range("H1").Offset(i, 0).Value = "A"
            ElseIf Range("I1").Offset(i, 0) >= 0.5 And Range("I1").Offset(i, 0) <= 0.61 Then
                Range("H1").Offset(i, 0).Value = "B"
            ElseIf Range("I1").Offset(i, 0) >= 0.37 And Range("I1").Offset(i, 0) <= 0.49 Then
                Range("H1").Offset(i, 0).Value = "C"
            ElseIf Range("I1").Offset(i, 0) >= 0.25 And Range("I1").Offset(i, 0) <= 0.36 Then
                Range("H1").Offset(i, 0).Value = "D"
            ElseIf Range("I1").Offset(i, 0) <= 0.24 Then
                Range("H1").Offset(i, 0).Value = "F"
            Else: Range("H1").Offset(i, 0).Interior.Color = vbGreen
            End If

    ' Check letter grade D against decodeability scale.
    ElseIf Range("H1").Offset(i, 0) = "D" And (Range("I1").Offset(i, 0) < 0.25 Or Range("I1").Offset(i, 0) > 0.36) Then
            
            Range("H1").Offset(i, 0).Interior.Color = vbRed
            
            If Range("I1").Offset(i, 0) >= 0.62 Then
                Range("H1").Offset(i, 0).Value = "A"
            ElseIf Range("I1").Offset(i, 0) >= 0.5 And Range("I1").Offset(i, 0) <= 0.61 Then
                Range("H1").Offset(i, 0).Value = "B"
            ElseIf Range("I1").Offset(i, 0) >= 0.37 And Range("I1").Offset(i, 0) <= 0.49 Then
                Range("H1").Offset(i, 0).Value = "C"
            ElseIf Range("I1").Offset(i, 0) >= 0.25 And Range("I1").Offset(i, 0) <= 0.36 Then
                Range("H1").Offset(i, 0).Value = "D"
            ElseIf Range("I1").Offset(i, 0) <= 0.24 Then
                Range("H1").Offset(i, 0).Value = "F"
            Else: Range("H1").Offset(i, 0).Interior.Color = vbGreen
            End If

    ' Check letter grade F against decodeability scale.
    ElseIf Range("H1").Offset(i, 0) = "F" And Range("I1").Offset(i, 0) > 0.24 Then
            
            Range("H1").Offset(i, 0).Interior.Color = vbRed
            
            Range("H1").Offset(i, 0).Interior.Color = vbRed
            
            If Range("I1").Offset(i, 0) >= 0.62 Then
                Range("H1").Offset(i, 0).Value = "A"
            ElseIf Range("I1").Offset(i, 0) >= 0.5 And Range("I1").Offset(i, 0) <= 0.61 Then
                Range("H1").Offset(i, 0).Value = "B"
            ElseIf Range("I1").Offset(i, 0) >= 0.37 And Range("I1").Offset(i, 0) <= 0.49 Then
                Range("H1").Offset(i, 0).Value = "C"
            ElseIf Range("I1").Offset(i, 0) >= 0.25 And Range("I1").Offset(i, 0) <= 0.36 Then
                Range("H1").Offset(i, 0).Value = "D"
            ElseIf Range("I1").Offset(i, 0) <= 0.24 Then
                Range("H1").Offset(i, 0).Value = "F"
            Else: Range("H1").Offset(i, 0).Interior.Color = vbGreen
            End If
                
    End If
    
Next i

End Sub
Sub decodeabilityCU()
' PURPOSE: Identifies CU UPC decodeability issues

'DECODABILITY GRADE CONVERSION SCALE:
'A >= 62
'B 50 - 61
'C 37 - 49
'D 25 - 36
'F <= 24

Dim nRows As Integer 'total rows
Dim i As Integer 'row counter

' Activate Worksheet Sheet1
Worksheets("Sheet1").Activate

' Calculates number of rows
nRows = Range("Q1").CurrentRegion.Rows.Count - 1

' Loops through rows. Highlights & fixes cells with inappropriate Letter Grade.
For i = 1 To nRows
        
    ' BLUE FILL cells that don't have an appropriate Letter Grade or dash (indicating blank).
    If Range("Q1").Offset(i, 0) <> "A" And Range("Q1").Offset(i, 0) <> "B" And Range("Q1").Offset(i, 0) <> "C" And Range("Q1").Offset(i, 0) <> "D" And Range("Q1").Offset(i, 0) <> "F" And Range("Q1").Offset(i, 0) <> "-" Then
        Range("Q1").Offset(i, 0).Interior.Color = vbGreen
    End If
        
    ' Check letter grade A against decodeability scale.
    If Range("Q1").Offset(i, 0) = "A" And Range("R1").Offset(i, 0) < 0.62 Then
            
            Range("Q1").Offset(i, 0).Interior.Color = vbRed
            
            If Range("R1").Offset(i, 0) >= 0.62 Then
                Range("Q1").Offset(i, 0).Value = "A"
            ElseIf Range("R1").Offset(i, 0) >= 0.5 And Range("R1").Offset(i, 0) <= 0.61 Then
                Range("Q1").Offset(i, 0).Value = "B"
            ElseIf Range("R1").Offset(i, 0) >= 0.37 And Range("R1").Offset(i, 0) <= 0.49 Then
                Range("Q1").Offset(i, 0).Value = "C"
            ElseIf Range("R1").Offset(i, 0) >= 0.25 And Range("R1").Offset(i, 0) <= 0.36 Then
                Range("Q1").Offset(i, 0).Value = "D"
            ElseIf Range("R1").Offset(i, 0) <= 0.24 Then
                Range("Q1").Offset(i, 0).Value = "F"
            Else: Range("Q1").Offset(i, 0).Interior.Color = vbGreen
            End If
    
    ' Check letter grade B against decodeability scale.
    ElseIf Range("Q1").Offset(i, 0) = "B" And (Range("R1").Offset(i, 0) < 0.5 Or Range("R1").Offset(i, 0) > 0.61) Then
            
            Range("Q1").Offset(i, 0).Interior.Color = vbRed
            
            If Range("R1").Offset(i, 0) >= 0.62 Then
                Range("Q1").Offset(i, 0).Value = "A"
            ElseIf Range("R1").Offset(i, 0) >= 0.5 And Range("R1").Offset(i, 0) <= 0.61 Then
                Range("Q1").Offset(i, 0).Value = "B"
            ElseIf Range("R1").Offset(i, 0) >= 0.37 And Range("R1").Offset(i, 0) <= 0.49 Then
                Range("Q1").Offset(i, 0).Value = "C"
            ElseIf Range("R1").Offset(i, 0) >= 0.25 And Range("R1").Offset(i, 0) <= 0.36 Then
                Range("Q1").Offset(i, 0).Value = "D"
            ElseIf Range("R1").Offset(i, 0) <= 0.24 Then
                Range("Q1").Offset(i, 0).Value = "F"
            Else: Range("Q1").Offset(i, 0).Interior.Color = vbGreen
            End If

    ' Check letter grade C against decodeability scale.
    ElseIf Range("Q1").Offset(i, 0) = "C" And (Range("R1").Offset(i, 0) < 0.37 Or Range("R1").Offset(i, 0) > 0.49) Then
            
            Range("Q1").Offset(i, 0).Interior.Color = vbRed
            
            If Range("R1").Offset(i, 0) >= 0.62 Then
                Range("Q1").Offset(i, 0).Value = "A"
            ElseIf Range("R1").Offset(i, 0) >= 0.5 And Range("R1").Offset(i, 0) <= 0.61 Then
                Range("Q1").Offset(i, 0).Value = "B"
            ElseIf Range("R1").Offset(i, 0) >= 0.37 And Range("R1").Offset(i, 0) <= 0.49 Then
                Range("Q1").Offset(i, 0).Value = "C"
            ElseIf Range("R1").Offset(i, 0) >= 0.25 And Range("R1").Offset(i, 0) <= 0.36 Then
                Range("Q1").Offset(i, 0).Value = "D"
            ElseIf Range("R1").Offset(i, 0) <= 0.24 Then
                Range("Q1").Offset(i, 0).Value = "F"
            Else: Range("Q1").Offset(i, 0).Interior.Color = vbGreen
            End If

    ' Check letter grade D against decodeability scale.
    ElseIf Range("Q1").Offset(i, 0) = "D" And (Range("R1").Offset(i, 0) < 0.25 Or Range("R1").Offset(i, 0) > 0.36) Then
            
            Range("Q1").Offset(i, 0).Interior.Color = vbRed
            
            If Range("R1").Offset(i, 0) >= 0.62 Then
                Range("Q1").Offset(i, 0).Value = "A"
            ElseIf Range("R1").Offset(i, 0) >= 0.5 And Range("R1").Offset(i, 0) <= 0.61 Then
                Range("Q1").Offset(i, 0).Value = "B"
            ElseIf Range("R1").Offset(i, 0) >= 0.37 And Range("R1").Offset(i, 0) <= 0.49 Then
                Range("Q1").Offset(i, 0).Value = "C"
            ElseIf Range("R1").Offset(i, 0) >= 0.25 And Range("R1").Offset(i, 0) <= 0.36 Then
                Range("Q1").Offset(i, 0).Value = "D"
            ElseIf Range("R1").Offset(i, 0) <= 0.24 Then
                Range("Q1").Offset(i, 0).Value = "F"
            Else: Range("Q1").Offset(i, 0).Interior.Color = vbGreen
            End If

    ' Check letter grade F against decodeability scale.
    ElseIf Range("Q1").Offset(i, 0) = "F" And Range("R1").Offset(i, 0) > 0.24 Then
            
            Range("Q1").Offset(i, 0).Interior.Color = vbRed
            
            Range("Q1").Offset(i, 0).Interior.Color = vbRed
            
            If Range("R1").Offset(i, 0) >= 0.62 Then
                Range("Q1").Offset(i, 0).Value = "A"
            ElseIf Range("R1").Offset(i, 0) >= 0.5 And Range("R1").Offset(i, 0) <= 0.61 Then
                Range("Q1").Offset(i, 0).Value = "B"
            ElseIf Range("R1").Offset(i, 0) >= 0.37 And Range("R1").Offset(i, 0) <= 0.49 Then
                Range("Q1").Offset(i, 0).Value = "C"
            ElseIf Range("R1").Offset(i, 0) >= 0.25 And Range("R1").Offset(i, 0) <= 0.36 Then
                Range("Q1").Offset(i, 0).Value = "D"
            ElseIf Range("R1").Offset(i, 0) <= 0.24 Then
                Range("Q1").Offset(i, 0).Value = "F"
            Else: Range("Q1").Offset(i, 0).Interior.Color = vbGreen
            End If
                
    End If
    
Next i

End Sub
Sub inconsistentDataEntry()
' PURPOSE: Identifies inconsistent data entry for Case & Consumer Unit height, depth, width, & weight.

' Declare variables
Dim i As Integer 'row counter
Dim nRows As Integer 'number of rows
Dim pos As Integer 'row index of active SampleNbr in array
Dim absoluteDev As Single 'captures deviation from the mode/median one value at a time

' Activate raw data sheet
Worksheets("Sheet1").Activate

' Count rows of data (using SampleNbr field)
nRows = Range(Range("E1"), Range("E1").End(xlDown)).Rows.Count - 1

'' Loop through raw data.
'   Compare each measurement to the Pivot Table average for that SampleNbr and measurement type.
'   Highlight any values that are higher or lower than the mode/median for that SampleNbr/measurement type combination (based on the "Deviation Allowed # entered on Instructions tab).
For i = 1 To nRows

    ' Collect raw data SampleNbr
    comparisonID = Range("E1").Offset(i, 0).Value
        
    ' Find & collect position for Case Measurements in the array based on SampleNbr
    pos = WhereInArray(sampleNbr, comparisonID)
        
' Checks to ensure there are at least some measurement values in this row (not just blanks). If all blanks, skips to the next row.
If Range("T1").Offset(i, 0).Value <> "-" Or Range("U1").Offset(i, 0).Value <> "-" Or Range("V1").Offset(i, 0).Value <> "-" Or Range("W1").Offset(i, 0).Value <> "-" Or Range("X1").Offset(i, 0).Value <> "-" Or Range("Y1").Offset(i, 0).Value <> "-" Or Range("Z1").Offset(i, 0).Value <> "-" Or Range("AA1").Offset(i, 0).Value <> "-" Then
        
    ' Check to see if there are any values for that measurement type for this row. If there are ANY, then proceed. If all are blank, then skip. (Note: this is necessary to prevent errors.)
    If Range("T1").Offset(i, 0).Value <> "-" Or Range("U1").Offset(i, 0).Value <> "-" Or Range("V1").Offset(i, 0).Value <> "-" Or Range("W1").Offset(i, 0).Value <> "-" Then
        
        ''' Collect values' absolute deviation from the mode/median & highlight red if deviation >= "Deviation Allowed" (entered by the user).
        
        '   Case Height
        If Range("T1").Offset(i, 0).Value <> "-" Then
            absoluteDev = Abs(Range("T1").Offset(i, 0).Value - modeHeightCase(pos))
            If absoluteDev >= devAllowed Then Range("T1").Offset(i, 0).Interior.Color = vbGreen
        End If
        
        '   Case Depth
        If Range("U1").Offset(i, 0).Value <> "-" Then
            absoluteDev = Abs(Range("U1").Offset(i, 0).Value - modeDepthCase(pos))
            If absoluteDev >= devAllowed Then Range("U1").Offset(i, 0).Interior.Color = vbGreen
        End If
        
        '   Case Width
        If Range("V1").Offset(i, 0).Value <> "-" Then
            absoluteDev = Abs(Range("V1").Offset(i, 0).Value - modeWidthCase(pos))
            If absoluteDev >= devAllowed Then Range("V1").Offset(i, 0).Interior.Color = vbGreen
        End If
        
        '   Case Weight
        If Range("W1").Offset(i, 0).Value <> "-" Then
            absoluteDev = Abs(Range("W1").Offset(i, 0).Value - modeWeightCase(pos))
            If absoluteDev >= devAllowed Then Range("W1").Offset(i, 0).Interior.Color = vbGreen
        End If
        
    End If
        
    ' Check to see if there are any values for that measurement type for this row. If there are ANY, then proceed. If all are blank, then skip. (Note: this is necessary to prevent errors.)
    If Range("X1").Offset(i, 0).Value <> "-" Or Range("Y1").Offset(i, 0).Value <> "-" Or Range("Z1").Offset(i, 0).Value <> "-" Or Range("AA1").Offset(i, 0).Value <> "-" Then
          
        ''' Collect values' absolute deviation from the mean & highlight red if deviation >= "Deviation Allowed" (entered by the user).
        
        '   CU Height
        If Range("X1").Offset(i, 0).Value <> "-" Then
            absoluteDev = Abs(Range("X1").Offset(i, 0).Value - modeHeightCU(pos))
            If absoluteDev >= devAllowed Then Range("X1").Offset(i, 0).Interior.Color = vbGreen
        End If
        
        '   CU Depth
        If Range("Y1").Offset(i, 0).Value <> "-" Then
            absoluteDev = Abs(Range("Y1").Offset(i, 0).Value - modeDepthCU(pos))
            If absoluteDev >= devAllowed Then Range("Y1").Offset(i, 0).Interior.Color = vbGreen
        End If
        
        '   CU Width
        If Range("Z1").Offset(i, 0).Value <> "-" Then
            absoluteDev = Abs(Range("Z1").Offset(i, 0).Value - modeWidthCU(pos))
            If absoluteDev >= devAllowed Then Range("Z1").Offset(i, 0).Interior.Color = vbGreen
        End If
        
        '   CU Weight
        If Range("AA1").Offset(i, 0).Value <> "-" Then
            absoluteDev = Abs(Range("AA1").Offset(i, 0).Value - modeWeightCU(pos))
            If absoluteDev >= devAllowed Then Range("AA1").Offset(i, 0).Interior.Color = vbGreen
        End If
        
    End If
    
End If
            
Next i

End Sub
Sub collectModes()
' PURPOSE: Compute the mode for all measurements of each SampleNbr & collect into an array.

' Declare variables
Dim i As Integer 'row counter
Dim j As Integer 'block counter
Dim k As Integer 'col counter
Dim nRows As Integer 'number of rows
Dim nCols As Integer 'number of column blocks
Dim nBlock As Integer 'measurement block counter

' Count rows of data (using SampleNbr field)
nRows = Range(Range("E1"), Range("E1").End(xlDown)).Rows.Count - 1

' Set counter to zero
nBlock = 0

' Add new workbook called "Modes"
Workbooks.Add
ActiveWorkbook.SaveAs Filename:=filepath & "\" & "Modes_DeleteMe.xlsx"

' Activate RawFile workbook
Workbooks("RawFile.xls").Activate

'' Loop through raw data and copy/paste.
'   DETAIL: For each unique SampleNbr, copy/paste each measurement type into its own column (e.g. one column of all CU Height for that one SampleNbr ONLY). When get to new SampleNbr, that's a new column.
For i = 1 To nRows

    ' Activate RawFile workbook, worksheet "Sheet1"
    Workbooks("RawFile.xls").Activate
    Worksheets("Sheet1").Activate

    ' Checks to see if a new SampleNbr has begun.
    If Range("E1").Offset(i, 0) = Range("E2") Or Range("E1").Offset(i, 0).Value = Range("E1").Offset(i - 1, 0).Value Then

        ' Copy/Paste SampleNbr (below previous data for that same SampleNbr).
        Range("E1,T1:AA1").Offset(i, 0).Copy
        Workbooks("Modes_DeleteMe.xlsx").Activate
        Range("A10000").Offset(0, nBlock * 9).End(xlUp).Offset(1, 0).PasteSpecial

    Else

        ' Copy/Paste SampleNbr (in a new column block for that SampleNbr).
        Range("E1,T1:AA1").Offset(i, 0).Copy
        Workbooks("Modes_DeleteMe.xlsx").Activate
        Range("A2").End(xlToRight).Offset(0, 1).PasteSpecial
        
        ' Format the SampleNbr row
        Workbooks("RawFile.xls").Activate
        Worksheets("Sheet1").Activate
        With Range("A1").Offset(i, 0).EntireRow
            .Interior.Color = vbYellow
            .Font.Bold = True
        End With
        
        ' Increment the block counter by 1.
        nBlock = nBlock + 1

    End If
    
Next i

' Activate Modes workbook
Workbooks("Modes_DeleteMe.xlsx").Activate

' Counts the number of column blocks.
nCols = (Range(Range("A2"), Range("A2").End(xlToRight)).EntireColumn.Count / 9)

' Resizes the array to accomodate the data from this sheet.
ReDim sampleNbr(1 To nCols)
ReDim modeHeightCase(1 To nCols)
ReDim modeDepthCase(1 To nCols)
ReDim modeWidthCase(1 To nCols)
ReDim modeWeightCase(1 To nCols)
ReDim modeHeightCU(1 To nCols)
ReDim modeDepthCU(1 To nCols)
ReDim modeWidthCU(1 To nCols)
ReDim modeWeightCU(1 To nCols)

' Loops through each column block.
For j = 1 To nCols

    ' Loops through each column (within each column block).
    For k = 1 To 8
       
        'Calculate the number of rows for that specific column
        nRows = Range(Range("A1").Offset(0, (j - 1) * 9).Offset(0, k), Range("A10000").Offset(0, (j - 1) * 9).End(xlUp).Offset(0, k)).Rows.Count - 1
        
        'Calculate the mode for measurement columns
        Range("A10000").Offset(0, (j - 1) * 9).Offset(0, k).End(xlUp).Offset(1, 0).FormulaR1C1 = "=MODE(R[-" & nRows & "]C[0]:R[-1]C[0])"
        
        'If Mode can't be computed (e.g. every number in the range is unique), then compute the Median instead.
        If IsError(Range("A10000").Offset(0, (j - 1) * 9).Offset(0, k).End(xlUp).Value) Then
            Range("A10000").Offset(0, (j - 1) * 9).Offset(0, k).End(xlUp).FormulaR1C1 = "=MEDIAN(R[-" & nRows & "]C[0]:R[-1]C[0])"
        
            If IsError(Range("A10000").Offset(0, (j - 1) * 9).Offset(0, k).End(xlUp).Value) Then
                Range("A10000").Offset(0, (j - 1) * 9).Offset(0, k).End(xlUp).Value = 0
            End If
        End If
      
    Next k
    
    ' Collect data into an array.
    sampleNbr(j) = Range("A10000").Offset(0, (j - 1) * 9).End(xlUp).Offset(0, 0).Value
    modeHeightCase(j) = Range("A10000").Offset(0, ((j - 1) * 9) + 1).End(xlUp).Value
    modeDepthCase(j) = Range("A10000").Offset(0, ((j - 1) * 9) + 2).End(xlUp).Value
    modeWidthCase(j) = Range("A10000").Offset(0, ((j - 1) * 9) + 3).End(xlUp).Value
    modeWeightCase(j) = Range("A10000").Offset(0, ((j - 1) * 9) + 4).End(xlUp).Value
    modeHeightCU(j) = Range("A10000").Offset(0, ((j - 1) * 9) + 5).End(xlUp).Value
    modeDepthCU(j) = Range("A10000").Offset(0, ((j - 1) * 9) + 6).End(xlUp).Value
    modeWidthCU(j) = Range("A10000").Offset(0, ((j - 1) * 9) + 7).End(xlUp).Value
    modeWeightCU(j) = Range("A10000").Offset(0, ((j - 1) * 9) + 8).End(xlUp).Value
    
Next j

' Close Modes file (without displaying an alert)
Application.DisplayAlerts = False
Workbooks("Modes_DeleteMe.xlsx").Save
Workbooks("Modes_DeleteMe.xlsx").Close
Application.DisplayAlerts = True

End Sub
Sub missingData()
' PURPOSE: fills in missing values for the following fields: GTIN, Consumer Unit UPC, Description, & Category.

' Declare variables
Dim i As Integer 'row counter
Dim nRows As Integer 'number of rows

' Count rows of data (using SampleNbr field)
nRows = Range(Range("E1"), Range("E1").End(xlDown)).Rows.Count - 1

' Loops through all rows for the relevant fields.
For i = 1 To nRows

    '' GTIN Field
    ' Checks to see if the cell is blank. If it's blank and the previous row isn't row 1 & the previous row belongs to the same SampleNbr, then copies down the value from the cell above.
    If (Range("F1").Offset(i, 0).Value = "-") And (Range("F1").Offset(i - 1, 0).Value <> Range("F1")) And (Range("E1").Offset(i, 0).Value = Range("E1").Offset(i - 1, 0).Value) Then
    
        Range("F1").Offset(i, 0).Value = Range("F1").Offset(i - 1, 0).Value
    
    End If
    
    '' Consumer Unit UPC Field
    ' Checks to see if the cell is blank. If it's blank and the previous row isn't row 1 & the previous row belongs to the same SampleNbr, then copies down the value from the cell above.
    If (Range("K1").Offset(i, 0).Value = "-") And (Range("K1").Offset(i - 1, 0).Value <> Range("K1")) And (Range("E1").Offset(i, 0).Value = Range("E1").Offset(i - 1, 0).Value) Then
    
        Range("K1").Offset(i, 0).Value = Range("K1").Offset(i - 1, 0).Value
    
    End If
    
    '' Description Field
    ' Checks to see if the cell is blank. If it's blank and the previous row isn't row 1 & the previous row belongs to the same SampleNbr, then copies down the value from the cell above.
    If (Range("L1").Offset(i, 0).Value = "-") And (Range("L1").Offset(i - 1, 0).Value <> Range("L1")) And (Range("E1").Offset(i, 0).Value = Range("E1").Offset(i - 1, 0).Value) Then
    
        Range("L1").Offset(i, 0).Value = Range("L1").Offset(i - 1, 0).Value
    
    End If
    
    '' Category Field
    ' Checks to see if the cell is blank. If it's blank and the previous row isn't row 1 & the previous row belongs to the same SampleNbr, then copies down the value from the cell above.
    If (Range("M1").Offset(i, 0).Value = "-") And (Range("M1").Offset(i - 1, 0).Value <> Range("M1")) And (Range("E1").Offset(i, 0).Value = Range("E1").Offset(i - 1, 0).Value) Then
    
        Range("M1").Offset(i, 0).Value = Range("M1").Offset(i - 1, 0).Value
    
    End If

Next i

End Sub
Sub wrongProdCode()
' PURPOSE: Identifies Production Codes that may be wrong.

' Declare variables
Dim i As Integer 'row counter
Dim nRows As Integer 'number of rows

' Count rows of data (using SampleNbr field)
nRows = Range(Range("E1"), Range("E1").End(xlDown)).Rows.Count - 1

' Loops through all the Production Codes.
For i = 1 To nRows

    ' Checks to see if a new SampleNbr has begun AND the Production Code for that cell is inconsistent with the previous one (of the same SampleNbr).
    If Range("E1").Offset(i, 0).Value = Range("E1").Offset(i - 1, 0).Value And Range("S1").Offset(i, 0).Value <> Range("S1").Offset(i - 1, 0).Value Then
        
        'If yes, then highlight the cell.
        Range("S1").Offset(i, 0).Interior.Color = vbGreen
        
    End If
    
Next i

End Sub
Sub shipperUPCLookup()
' PURPOSE: Pulls in the following data based on the Shipper UPC: Description, GTIN, UPC Case Code - Chk Digit, UPC Pack Code - Chk Digit, & Base Prod Code.

' Declare variables
Dim i As Integer 'row counter
Dim nRows As Integer 'number of rows
Dim nRowsDetail As String 'number of rows on Detail tab
Dim lookupSource As Range 'named source range for vlookup

' Activate RawFile workbook
Workbooks("RawFile.xls").Activate

    ' Insert 5 rows after Shipper UPC.
    Range(Range("H1"), Range("L1")).EntireColumn.Insert
    
    ' Name headers for new columns.
    Range("H1") = "DESCRIPTION"
    Range("I1") = "GTIN"
    Range("J1") = "UPC Case Code - Chk Digit"
    Range("K1") = "UPC Pack Code - Chk Digit"
    Range("L1") = "Base Prod Code"

' Add new worksheet called "Detail"
Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = "Detail"

    ' Open "GS1 Detail" Excel file.
    Workbooks.Open (filepath & "\GS1 Detail.xlsx")
    
    ' Count number of rows
    nRowsDetail = Range(Range("A1"), Range("A1").End(xlDown)).Count
        
    ' Copy Detail data from GS1 Detail
    Range(Range("A1").End(xlDown), Range("F1")).Copy
    'lookupSource.Copy

' Activate RawFile workbook/Detail sheet
Workbooks("RawFile.xls").Activate
Worksheets("Detail").Activate

    ' Paste Detail data into RawFile
    Range("A1").PasteSpecial
    
    ' Close Detail file (without displaying an alert)
    Application.DisplayAlerts = False
    Workbooks("GS1 Detail.xlsx").Close
    Application.DisplayAlerts = True
    
    ' Activate Sheet1
    Worksheets("Sheet1").Activate

' Count rows of data (using Shipper UPC field)
nRows = Range(Range("G1"), Range("G1").End(xlDown)).Rows.Count - 1

' Loop through all rows for the 5 new columns.
For i = 1 To nRows
    
    ' Use VLookup to pull over fields from "GS1 Detail" file. (If VLOOKUP = 0, then display a blank. Otherwise, use the VLOOKUP. If result of that is an error, then display a blank.)
    Range("H1").Offset(i, 0).Formula = "=IF((VLOOKUP(G" & i + 1 & ", 'Detail'!A1:F" & nRowsDetail & ",2,FALSE))=0,"""",(VLOOKUP(G" & i + 1 & ", 'Detail'!A1:F" & nRowsDetail & ",2,FALSE)))"
    If IsError(Range("H1").Offset(i, 0).Value) = True Then
        Range("H1").Offset(i, 0).Value = ""
    End If
    
    Range("I1").Offset(i, 0).Formula = "=IF((VLOOKUP(G" & i + 1 & ", 'Detail'!A1:F" & nRowsDetail & ",3,FALSE))=0,"""",(VLOOKUP(G" & i + 1 & ", 'Detail'!A1:F" & nRowsDetail & ",3,FALSE)))"
    If IsError(Range("I1").Offset(i, 0).Value) = True Then
        Range("I1").Offset(i, 0).Value = ""
    End If
    
    Range("J1").Offset(i, 0).Formula = "=IF((VLOOKUP(G" & i + 1 & ", 'Detail'!A1:F" & nRowsDetail & ",4,FALSE))=0,"""",(VLOOKUP(G" & i + 1 & ", 'Detail'!A1:F" & nRowsDetail & ",4,FALSE)))"
    If IsError(Range("J1").Offset(i, 0).Value) = True Then
        Range("J1").Offset(i, 0).Value = ""
    End If
    
    Range("K1").Offset(i, 0).Formula = "=IF((VLOOKUP(G" & i + 1 & ", 'Detail'!A1:F" & nRowsDetail & ",5,FALSE))=0,"""",(VLOOKUP(G" & i + 1 & ", 'Detail'!A1:F" & nRowsDetail & ",5,FALSE)))"
    If IsError(Range("K1").Offset(i, 0).Value) = True Then
        Range("K1").Offset(i, 0).Value = ""
    End If
    
    Range("L1").Offset(i, 0).Formula = "=IF((VLOOKUP(G" & i + 1 & ", 'Detail'!A1:F" & nRowsDetail & ",6,FALSE))=0,"""",(VLOOKUP(G" & i + 1 & ", 'Detail'!A1:F" & nRowsDetail & ",6,FALSE)))"
    If IsError(Range("L1").Offset(i, 0).Value) = True Then
        Range("L1").Offset(i, 0).Value = ""
    End If
    
Next i

End Sub
Sub gtinCheck()
' PURPOSE:
''' Check if Shipper UPC is wholly contained within the GTIN. If NOT, then check VLOOKUP GTIN.
''' If it has a number, then pull the number over to the permanent GTIN field.
''' If it's blank, then highlight the permanent GTIN field to flag it for research.

' Declare variables
Dim i As Integer 'row counter
Dim nRows As Integer 'number of rows
Dim isContained As Boolean 'T/F container
Dim shipperUPC As String 'Shipper UPC searching for

' Activate RawFile workbook
Workbooks("RawFile.xls").Activate

' Count rows of data (using SampleNbr field)
nRows = Range(Range("E1"), Range("E1").End(xlDown)).Rows.Count - 1

' Loop through each GTIN.
For i = 1 To nRows

    ' Determine if Shipper UPC is wholly contained within the permanent GTIN field.
    isContained = InStr(1, Range("F1").Offset(i, 0).Value, Range("G1").Offset(i, 0).Value, vbBinaryCompare)
    
    ' Check if Shipper UPC is wholly contained within the permanent GTIN. If NOT, then check the VLOOKUP GTIN.
        '   If VLOOKUP GTIN has a number, then pull the number over to the permanent GTIN field.
        '   If VLOOKUP GTIN is blank, then highlight the permanent GTIN field to flag it for research.
    If isContained = False Then
        
        ' Check if VLOOKUP GTIN has a number.
        If IsNumeric(Range("I1").Offset(i, 0).Value) = True Then
        
            ' Pull the LOOKUP GTIN over to the permanent GTIN field & highlight permanent GTIN field as an FYI.
            Range("I1").Offset(i, 0).Copy
            Range("F1").Offset(i, 0).PasteSpecial xlPasteValues
            Range("F1").Offset(i, 0).Interior.Color = vbRed
        
        Else
        
            ' Check if VLOOKUP GTIN is blank.
            If IsError(Range("I1").Offset(i, 0).Value) = True Then
                
                ' Highlight the permanent GTIN field to flag it for research.
                Range("F1").Offset(i, 0).Interior.Color = vbGreen
            
            ' Check if VLOOKUP GTIN is displaying #N/A. (This indicates that Shipper UPC isn't in the VLOOKUP file.)
            ElseIf Range("I1").Offset(i, 0).Value = "" Then
            
                ' Highlight the permanent GTIN field to flag it for research.
                Range("F1").Offset(i, 0).Interior.Color = vbGreen
                
            End If
        
        End If

    End If

Next i

End Sub
Sub finalFormatting()
' PURPOSE: Concatenate Location, City, & State fields; formats column headers; formats entire sheet (center-alignment & autofits the columns); renames the sheet; deletes unnecessary sheets; and freezes top row. Save as a new file.

' Declare variables
Dim i As Integer 'row counter
Dim nRows As Integer 'number of rows

' Concatenate Location, City, & State fields.

    ' Insert new column after State column.
    Range("E1").EntireColumn.Insert
    
    ' Add heading for new column
    Range("E1").Value = "Location - City, ST"
    
    ' Count rows of data (using SampleNbr field)
    nRows = Range(Range("F1"), Range("F1").End(xlDown)).Rows.Count - 1

    ' Loop through rows.
    For i = 1 To nRows

        ' Concatenate Location, City, & State fields for each row.
        Range("E1").Offset(i, 0).Formula = Range("B1").Offset(i, 0) & " - " & Range("C1").Offset(i, 0) & ", " & Range("D1").Offset(i, 0)

    Next i
    
    ' Delete the new duplicative Location, City, & State columns.
    Range("B1:D1").EntireColumn.Delete
    
    ' Collect Location/City/State (for renaming the file later)
    locationCityState = Range("B2") & " "

' Format column headers
With Range("A1").EntireRow
    .Interior.Color = RGB(146, 208, 80)
    .Font.Bold = True
End With
       
' Resize all columns to fit the size of the data and to be center aligned.
With Range("A1:AD1").EntireColumn
    .AutoFit
    .HorizontalAlignment = xlCenter
End With

' Copy VLOOKUP fields and paste as values.
With Range(Range("F1"), Range("J1").Offset(nRows, 0))
    .Copy
    .PasteSpecial xlPasteValues
End With

' Save As the raw data file.
Workbooks("RawFile.xls").Activate
ActiveWorkbook.SaveAs Filename:=filepath & "\" & locationCityState & yearQrt & " QA", FileFormat:=xlWorkbookDefault

' Delete Modes workbook & Detail worksheet (without displaying an alert).
Application.DisplayAlerts = False
Kill filepath & "\" & "Modes_DeleteMe.xlsx"
Worksheets("Detail").Delete
Application.DisplayAlerts = True

' Delete duplicate Description & GTIN columns (pulled from LOOKUP file).
Range("F1:G1").EntireColumn.Delete

' Truncate locationCityState if greater than 30 characters (30 is the max characters in a Sheet Name).
If Len(locationCityState) > 30 Then
    locationCityState = Left(locationCityState, 30)
End If

' Rename worksheet to Location/City/State.
Worksheets("Sheet1").Name = locationCityState

' Freeze top row
Rows("2:2").Select
ActiveWindow.FreezePanes = True

End Sub
