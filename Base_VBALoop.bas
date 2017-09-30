Attribute VB_Name = "Base_VBALoop"
Option Explicit

Sub TimesTables()
Dim startNumber As Long
Dim endNumber As Long
Dim feedNumber As Long
Dim counter As Integer
Dim pctDone As Double
TestProgress.Show
endNumber = 12
For startNumber = 1 To endNumber * 10000
    'If startNumber * endNumber <= 240 Then
        Cells(startNumber, 1) = startNumber & " times " & endNumber & _
            " = "
        Cells(startNumber, 2) = startNumber * endNumber
    'Else
        'Cells(startNumber, 1) = "Number too high"
    'End If
     counter = counter + 1
         ' Update the percentage completed.
'        pctDone = counter / (endNumber * 10000)

    ' Call subroutine that updates the progress bar.
    feedNumber = endNumber * 10000
    TestProgress.Show
Next startNumber

End Sub
Sub showTestProgress()
   TestProgress.Show
End Sub
