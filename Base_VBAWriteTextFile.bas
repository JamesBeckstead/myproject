Attribute VB_Name = "Base_VBAWriteTextFile"
Option Explicit
Sub ShowUserForm()
    UserForm1.Show
End Sub
Sub WriteToTextFile()
'+-------------------------------------------------------------------------------------+
'|This macro is set up to ensure the process is only run once and is run with today's date.  |
'|This solves two concerns. One, evidence the process ran and by whom.                        |
'|Two, the user can't run the process twice and find no results.                                     |
'+-------------------------------------------------------------------------------------+
Dim userName As String
Dim today As Date
Dim macroName As String
Dim filePath As String
Dim compileData As String
Dim lineFromFile As String
Dim lineItems As Variant
Dim line As Variant
Dim lLine As Integer
Dim lineDate As String
Dim processName As String
Dim userPath As String
Dim pctDone As Integer
Dim counter As Integer

'set the path for the saved file
userPath = Application.DefaultFilePath
filePath = userPath & "\TestFolder" & "\test.txt"
' -----------------------------------------------------------
'open the file
Open filePath For Input As #1
' -----------------------------------------------------------
'create an array of each line in the text file, use this later to
' determine if the process has already run
Do Until EOF(1)
    Line Input #1, lineFromFile
    lineItems = Split(lineFromFile, ",")
Loop
' -----------------------------------------------------------
'determine if the array was created, might not happen if the
' process has never run before, loop through the array
' and extract the date, if it has been run exit the sub and
' terminate the process
If IsArray(lineItems) Then
    For Each line In lineItems
        lLine = Len(line)
        lineDate = Mid(line, 2, lLine - 1)
        If lineDate = Date Then
            MsgBox "This has already been run today."
            End
        Else
            Close #1
            Exit For
        End If
        counter = counter + 1
         ' Update the percentage completed.
'        pctDone = counter / (UBound(lineItems) + 1)
'
'        ' Call subroutine that updates the progress bar.
'        UpdateProgressBar2 pctDone
    Next line
End If
Close #1
' ------------------------------------------------------------
'if the process has not been run today open the file and write
' the date, time, process, and username to the file
' save and close the file once write is complete
Open filePath For Append As #2

processName = "testMacro"
today = Date
compileData = today
compileData = compileData & ", " & Right(Now, 11) & ", " & _
    processName & ", " & Environ("userName")
Write #2, compileData
compileData = ""

Close #2
' ---------------------------------------------------------
'inform programer that file has been updated with data listed above
MsgBox "Done, " & processName
'WriteToTextFile = True

End Sub
'Sub UpdateProgressBar2(pctDone As Integer)
'    With UserForm1
'
'        ' Update the Caption property of the Frame control.
'        .FrameProgress.Caption = Format(pctDone, "0%")
'
'        ' Widen the Label control.
'        .LabelProgress.Width = pctDone * _
'            (.FrameProgress.Width - 10)
'    End With
'
'    ' The DoEvents allows the UserForm to update.
'    DoEvents
'End Sub
