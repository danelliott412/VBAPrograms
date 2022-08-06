'This program was written by Dan Elliott to create a bill of materials for specific AutoCAD blocks using a prefabbed "BOM.xlsm" file
'Data is imported into the Excel file using an AutoLISP program and then rearranged and pasted into the BOM file 


Attribute VB_Name = "Module1"
Sub Main_acc_bom()

Dim wb1 As Excel.Workbook  'set file location and variable
Set wb1 = Workbooks("Accessory BOM Program.xlsm")

Dim ws1 As Worksheet     'sets worksheet variable for Bom info sheet
Set ws1 = wb1.Worksheets("Acc BOM")

Dim ws2 As Worksheet     'sets worksheet variable for Bom info sheet
Set ws2 = wb1.Worksheets("Acc_Seq")

Dim ws3 As Worksheet     'sets worksheet variable for Bom info sheet
Set ws3 = wb1.Worksheets("Acc_Dim")

Dim jobnum As String
Dim file As String
Dim xfile As String
Dim msgAns As Integer
    
'************************HIGHLIGHT LINE*******************************

jobnum = ws1.Range("D7").Value 'pulls Job number from cell D7

 xfile = "P:\" & jobnum & "\" & jobnum & "_8_BOM\"

file = Dir(xfile, 16)
If file = "" Then
    msgAns = MsgBox("Job folder not found. Please make sure the Project Number box is filled out correctly and try again.", vbOKOnly, "Accessory BOM Program")

    Exit Sub

Else

    Dim bom As Excel.Workbook  'set file location and variable
    Set bom = Workbooks.Open("C:\string\to\BOM.xlsm")

    ws2.Calculate       'Calculate inputed values from autocad to BOM format
    ws3.Calculate

    Call listofseq
    Call acc_seq        'Call sub functions to create bom
    Call acc_dims
    Call project_info

    Application.Run "'C:\string\to\BOM.xlsm'!NumberBOMSet"

    Call save_bom

    With Application

        .Calculation = xlCalculationAutomatic       'set calculations to be performed Automatically before closing

    End With
    ThisWorkbook.Close SaveChanges:=False       'close without saving

End If
End Sub

Sub acc_seq()

Dim Filename As String

Dim wb1 As Excel.Workbook  'set file location and variable
Set wb1 = Workbooks("Accessory BOM Program.xlsm")

Dim bom As Excel.Workbook  'set file location and variable
Set bom = Workbooks.Open("C:\string\to\BOM.xlsm")
  
Dim ws1 As Worksheet     'sets worksheet variable for accessory consolidation sheet
Set ws1 = wb1.Worksheets("Acc_Seq")

Dim S As Worksheet      'sets worksheet variable for sequence accessory sheet
  
Dim accmark_column As Range 'sets range variable rows to count
Set accmark_column = ws1.Range("A1:A1000")

Dim seq_row As Range 'sets range variable for columns to count
Set seq_row = ws1.Range("B1:BBB1")

Dim sheets As Long 'number of sequence sheets to print

Dim n As Long 'printed column counter
n = 0
Dim m As Long 'printed row counter
Dim u As Long 'first row needs offset by 2 variable
Dim t As Long 'stores original dim i variable
Dim z As Long 'main for loop counter
Dim col As Long 'holds value for row count in equation
Dim o As Long 'holds value back from full row
'************************HIGHLIGHT LINE*******************************

Dim i As Double 'counts by number of rows
Dim j As Long
i = WorksheetFunction.CountA(accmark_column)
j = Application.WorksheetFunction.RoundUp(i / 36, 0) 'round sheets needed for rows up
t = i
Debug.Print j & " value j"
Dim x As Double 'counts by number of columns
Dim y As Long
x = WorksheetFunction.CountA(seq_row)
y = Application.WorksheetFunction.RoundUp(x / 14, 0) 'rounds sheets needed for columns up
Debug.Print y & " value y"

sheets = y * j 'rows * columns boundaries = seq sheets needed
Filename = Left(Application.ThisWorkbook.FullName, 8)

'Dim wb2 As Excel.Workbook  'set file location and variable
'Set wb2 = Workbooks.Open(Filename & "\danel\OneDrive\Desktop\BOM.xlsm")

'************************HIGHLIGHT LINE*******************************

Dim a As Long 'sheet counter and set as 1
a = 1
z = 0

For a = 1 To sheets         'while pages are left to print, run loop
    If j > 1 Then           'if more than 1 sheet needs printed for each sequence set
        If m = 0 Then       'u set as 2 instead of 1 on first set of acc for each seq to avoid seq numbers
            u = 2
        Else: u = 1
        End If
    Else: u = 2
    End If
            
            
    If i > 37 Then
        col = 37
    Else
        col = i
    End If
                                 
    If (x - n) < 14 Then
        o = x - 14
    Else
        o = n
    End If

    Debug.Print a & " value set"
    Debug.Print i & " value I"
    Debug.Print n & " value N"
    Debug.Print x & " value X"
    Debug.Print m & " value M"
    Debug.Print o & " value O"

        'S.Copy After:=ws1
        'ActiveSheet.Name = "S (" & a & ")"
    Application.Run "'C:\string\to\BOM.xlsm'!addseqsheet"
    Application.EnableEvents = False
                                                
    ws1.Range(ws1.Cells(u + (37 * m), 2 + n), ws1.Cells(col + (37 * m), 15 + o)).Copy 'copy acc amounts and values shifted down m
    bom.sheets("S (" & a & ")").Range("B14").PasteSpecial xlPasteValues
    ws1.Range(ws1.Cells(1, 2 + n), ws1.Cells(1, 15 + o)).Copy              'copies seq numbers
    bom.sheets("S (" & a & ")").Range("B13").PasteSpecial xlPasteValues
    ws1.Range(ws1.Cells(u + (37 * m), 1), ws1.Cells(col + (37 * m), 1)).Copy       'copies acc marks
    bom.sheets("S (" & a & ")").Range("A14").PasteSpecial xlPasteValues
        
    z = z + 1
    
    If z = j Then
        n = n + 14
        z = 0
        m = 0
    End If
    
    If i > 37 Then
        i = i - 37          'adjust counter for rows
        m = m + 1
    Else
        i = t
    End If
                  
Next a

Debug.Print "********************done**************************"

Application.EnableEvents = True

End Sub

Sub acc_dims()

Dim wb1 As Excel.Workbook  'set file location and variable
Set wb1 = Workbooks("Accessory BOM Program.xlsm")

Dim bom As Excel.Workbook  'set file location and variable
Set bom = Workbooks.Open("C:\string\to\BOM.xlsm")
  
Dim ws1 As Worksheet     'sets worksheet variable for accessory consolidation sheet
Set ws1 = wb1.Worksheets("Acc_Dim")

Dim S As Worksheet      'sets worksheet variable for sequence accessory sheet
  
Dim accmark_column As Range 'sets range variable rows to count
Set accmark_column = ws1.Range("N2:N1000")

Dim sheets As Long 'number of sequence sheets to print
Dim z As Long 'main for loop counter
Dim n As Long 'stores number of columns already pasted
Dim col As Long 'holds value for row count in equation
Dim row As Long 'holds value for row count in equation
Dim u As Long ' used to remove first empty space

u = 0
col = 17
row = 3

'************************HIGHLIGHT LINE*******************************

Dim i As Double 'counts by number of rows
Dim j As Long

i = WorksheetFunction.CountA(accmark_column)
j = Application.WorksheetFunction.RoundUp(i / 24, 0) 'round sheets needed for rows up

sheets = j 'columns needed/ 24 = total Acc pages needed

Dim a As Long 'sheet counter and set as 1
a = 1
z = 0
n = 0

For a = 1 To sheets  'while pages are left to print, run loop
                                 
    Debug.Print a & " value set"
    Debug.Print i & " value I"
    Debug.Print n & " value I"
    Debug.Print col & " value COL"

    Application.Run "'C:\string\to\BOM.xlsm'!addaccsheet"
    Application.EnableEvents = False
    
    ws1.Range(ws1.Cells(1 + (24 * n), 14), ws1.Cells(24 + (24 * n), 14)).Copy 'copy acc amounts and values shifted down m
    bom.sheets("A (" & a & ")").Range("A17").PasteSpecial xlPasteValues

    Application.EnableEvents = True

    bom.sheets("A (" & a & ")").Calculate

    Application.EnableEvents = False

    For col = 17 To 41
        For row = 3 To 7
     
            If bom.sheets("A (" & a & ")").Cells(col, row).Value = "X" Then
                ws1.Cells(col - 16 + ((a - 1) * 24), row + 13).Copy
                bom.sheets("A (" & a & ")").Cells(col, row).PasteSpecial xlPasteValues
                Debug.Print row & " row " & col & " col"
            ElseIf bom.sheets("A (" & a & ")").Cells(col, 1).Value = "CC1" Then
                col = col + 1
                row = 2
            ElseIf bom.sheets("A (" & a & ")").Cells(col, 1).Value = "CC2" Then
                col = col + 1
                row = 2
            ElseIf bom.sheets("A (" & a & ")").Cells(col, 1).Value = "CC3" Then
                col = col + 1
                row = 2
            ElseIf bom.sheets("A (" & a & ")").Cells(col, 1).Value = "ZC1" Then
                col = col + 1
                row = 2
            ElseIf bom.sheets("A (" & a & ")").Cells(col, 1).Value = "ZC2" Then
                col = col + 1
                row = 2
            ElseIf bom.sheets("A (" & a & ")").Cells(col, 1).Value = "ZC3" Then
                col = col + 1
                row = 2
            ElseIf bom.sheets("A (" & a & ")").Cells(col, 1).Value = "FP6" Then
                col = col + 1
                row = 2
            ElseIf bom.sheets("A (" & a & ")").Cells(col, 1).Value = "FP9" Then
                col = col + 1
                row = 2
            ElseIf bom.sheets("A (" & a & ")").Cells(col, 1).Value = "FP12" Then
                col = col + 1
                row = 2
            ElseIf bom.sheets("A (" & a & ")").Cells(col, 1).Value = "FS1" Then
                col = col + 1
                row = 2
            ElseIf bom.sheets("A (" & a & ")").Cells(col, 1).Value = "FS2" Then
                col = col + 1
                row = 2
            ElseIf bom.sheets("A (" & a & ")").Cells(col, row).Value = "1" Then
                ws1.Cells(col - 16 + ((a - 1) * 24), row + 13).Copy
                bom.sheets("A (" & a & ")").Cells(col, row).PasteSpecial xlPasteValues
            ElseIf bom.sheets("A (" & a & ")").Cells(col, row).Value = "0.5" Then
                ws1.Cells(col - 16 + ((a - 1) * 24), row + 13).Copy
                bom.sheets("A (" & a & ")").Cells(col, row).PasteSpecial xlPasteValues
            End If
            Debug.Print bom.sheets("A (" & a & ")").Cells(col, 3).Value
        Next row
    Next col
   
    If i > 24 Then
        i = i - 24  'adjust counter for rows
        n = n + 1 'multiplier for columns used
    End If
Next a
Debug.Print "********************done**************************"

Application.EnableEvents = True
End Sub
Sub project_info()
Dim wb1 As Excel.Workbook  'set file location and variable
Set wb1 = Workbooks("Accessory BOM Program.xlsm")

Dim bom As Excel.Workbook  'set file location and variable
Set bom = Workbooks.Open("C:\string\to\BOM.xlsm.xlsm")
  
Dim ws1 As Worksheet     'sets worksheet variable for accessory consolidation sheet
Set ws1 = wb1.Worksheets("Acc BOM")

'************************HIGHLIGHT LINE*******************************

Application.EnableEvents = False
                               
ws1.Range(ws1.Cells(6, 4), ws1.Cells(10, 4)).Copy 'copy project info
bom.sheets("ProjInfo").Range("D5").PasteSpecial xlPasteValues   'paste to bom
                
Debug.Print "********************done**************************"

Application.EnableEvents = True
End Sub

Sub Reset(RowNum As Integer)

Dim bom As Excel.Workbook  'set file location and variable
Set bom = Workbooks.Open("C:\string\to\BOM.xlsm.xlsm")

Dim sheet As Worksheet
Set sheet = bom.Worksheets("A (1)")

' Accessory Page Reset Buttons
With Application
    .ScreenUpdating = False
    .EnableEvents = False
End With
sheet.Range("C44:L44").Copy
sheet.Cells(RowNum, 3).PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
With Application
    .ScreenUpdating = True
    .EnableEvents = True
End With

End Sub

Sub listofseq()

Dim wb1 As Excel.Workbook  'set file location and variable
Set wb1 = Workbooks("Accessory BOM Program.xlsm")

Dim ws1 As Worksheet     'sets worksheet variable for Bom program user interface
Set ws1 = wb1.Worksheets("Acc BOM")

Dim ws2 As Worksheet     'sets worksheet variable for acc count data import info sheet
Set ws2 = wb1.Worksheets("acc counts")

Dim Sequence_List As String 'stores values of sequnces user wants to BOM
Dim Sequence_array() As String  'converts string to array that can be read by excel
Dim myCell As Range
Dim Rng_Del As Range
Dim rng As Range

'************************HIGHLIGHT LINE*******************************

Sequence_List = ws1.Cells(7, 11).Value
Debug.Print Sequence_List
If Sequence_List = "" Then
Exit Sub
Else

Sequence_array() = Split(Sequence_List, ", ") 'splits comma delimited list into rows
'use autofliter to hide unwanted data
ws2.Range("A:G").AutoFilter Field:=6, Criteria1:=Sequence_array(), Operator:=xlFilterValues, visibledropdown:=False

Set rng = ws2.Range("A1:G9999")

For Each myCell In rng.Columns(1).Cells
    If myCell.EntireRow.Hidden Then
        If Rng_Del Is Nothing Then
            Set Rng_Del = myCell
        Else
            Set Rng_Del = Union(Rng_Del, myCell)
         End If
    End If

Next

If Not Rng_Del Is Nothing Then Rng_Del.EntireRow.Clear      'delete rows with unwanted seq info
End If

End Sub

Sub save_bom()

Dim wb1 As Excel.Workbook   'set file location and variable
Set wb1 = Workbooks("Accessory BOM Program.xlsm")

Dim bom As Excel.Workbook   'set file location and variable
Set bom = Workbooks.Open("C:\string\to\BOM.xlsm")
  
Dim ws1 As Worksheet        'sets worksheet variable for accessory consolidation sheet
Set ws1 = wb1.Worksheets("Acc BOM")

Dim jobnum As String
Dim file As String
Dim xfile As String
    
'************************HIGHLIGHT LINE*******************************
    
jobnum = ws1.Range("D7").Value  'pulls Job number from cell D7
    
xfile = "P:\" & jobnum & "\"
bom.SaveAs Filename:=xfile & jobnum & " Accessory BOM.xlsm"

End Sub
