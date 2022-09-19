Attribute VB_Name = "Module1"
Sub openRfifolder()
Dim myfile As String

JobNum = ExtractJobNum()

Call Shell("explorer.exe" & " " & "P:\" & JobNum & "\" & JobNum & "_1_CORRESPONDENCE\" & JobNum & "_EMAIL", vbNormalFocus)

End Sub

Sub refreshEmailList()

Call turnofffunctionality

JobNum = ExtractJobNum()

Sheet3.Cells(2, 10) = JobNum
ThisWorkbook.Queries("EmailQuery").Refresh
Dim olApp As Object
Dim MSG As Object
Dim thisFile$
Set olApp = CreateObject("Outlook.Application")
Set MSG = olApp.CreateItem(MailItem)
Dim cellcount1 As Integer
Dim cellcount2 As Integer
Dim Email As String
Dim i
Dim inString



cellcount1 = Sheet3.Cells(2, 1).CurrentRegion.Rows.Count - 1

For i = 1 To cellcount1

If Sheet3.Cells(1, 12) = "" Then
cellcount2 = 0
Else
cellcount2 = Sheet3.Cells(1, 12).CurrentRegion.Count
End If

Email = Sheet3.Cells(i + 1, 1)



    Sheet2.Cells(2 + i, 1) = Sheet3.Cells(1 + i, 4) 'sets time of email

    Set MSG = olApp.CreateItemFromTemplate(Sheet3.Cells(1 + i, 6) & Email)
    
    If InStr(MSG.Body, "<") > 0 Then
    inString = InStr(MSG.Body, "<") - 1
    Else
    inString = Len(MSG.Body)
    End If
    Sheet2.Cells(i + 2, 8) = Trim(Replace(Replace(Left(MSG.Body, inString), Chr(10), ""), Chr(13), "")) 'prints out email body

      With Sheet2
        .Hyperlinks.Add Anchor:=.Cells(2 + i, 2), _
         Address:=Sheet3.Cells(1 + i, 6) & Email, _
         TextToDisplay:=Email                           'sets hyperlink
    End With



Next i

Set MSG = Nothing
Set olApp = Nothing

Call turnonfunctionality


End Sub

Function ExtractJobNum()

Dim JobNum As String
Dim File_Path As String
File_Path = ThisWorkbook.Name

JobNum = Left(File_Path, 9)

ExtractJobNum = JobNum
End Function

Public Sub turnofffunctionality()
Application.Calculation = xlCalculationManual
Application.DisplayStatusBar = False
Application.ScreenUpdating = False
Application.DisplayScrollBars = False


End Sub

Public Sub turnonfunctionality()
Application.Calculation = xlCalculationAutomatic
Application.DisplayStatusBar = True
Application.ScreenUpdating = True
Application.DisplayScrollBars = True


End Sub
