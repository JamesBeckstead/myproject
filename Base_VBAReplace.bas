Attribute VB_Name = "Base_VBAReplace"
Option Explicit

Sub TestForProper()
Dim prodCode As String
Dim location As Integer

prodCode = "PD-23-23-45"
prodCode = Replace(prodCode, "-", "")
location = InStr(prodCode, "D")
prodCode = Left(prodCode, location) & "C" & Right(prodCode, Len(prodCode) - location)

Debug.Print (prodCode)
'Debug.Print (lastName & ", " & firstName)

'Application.WorksheetFunction.Proper
End Sub
