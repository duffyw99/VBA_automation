# VBA_automation
scripts I use(d) to automate excel based reports at work


Sub RefreshAllPivots()
  'Refresh all pivots within the worksheet (alternative to ActiveWorkbook.RefreshAll )
Dim pt As PivotTable
Dim WS As Worksheet
    For Each WS In ActiveWorkbook.Worksheets
        For Each pt In WS.PivotTables
                pt.RefreshTable
        Next pt
    Next WS
End Sub
Sub RemoveDataConnection()
' Remove connections to embedded SQL
Dim cn As WorkbookConnection
Dim odbcCn As ODBCConnection, oledbCn As OLEDBConnection

    For Each cn In ThisWorkbook.Connections
       If cn.Type = xlConnectionTypeODBC Then
            Set odbcCn = cn.ODBCConnection
            odbcCn.BackgroundQuery = False
            odbcCn.CommandText = ""
            odbcCn.SavePassword = False
        ElseIf cn.Type = xlConnectionTypeOLEDB Then
            Set oledbCn = cn.OLEDBConnection
            oledbCn.BackgroundQuery = False
            oledbCn.CommandText = ""
            oledbCn.SavePassword = False
        End If
Next

End Sub
Sub SaveFile()
' Save file with dynamic date value in filename
Dim strUserName As String
Dim Year As String
Dim QtrYear As String
Dim MonYear As String
Dim WeekYear As String
Dim Today As String
    
'Use the Application Object to get the Username
strUserName = Application.UserName

'''' Timestamp variables for report name. Update save path appropriately
Year = Format(Now, "YYYY")  'Date variable based on system date
QtrYear = Format(DateAdd("M", -3, Now), "YYYY") & " " & "Q" & (Month(DateAdd("M", -3, Now)) + 2) \ 3 'Prior Qtr
MonYear = Format(DateAdd("M", -1, Now), "YYYY_MM") 'Prior month
WeekYear = Format(DateAdd("WW", -1, Now), "YYYY") & " Week " & Format(DateAdd("WW", -1, Now), "WW") 'Prior week
Today = Format(Date, "YYYY_MM_DD") 'Todays date based on system date

'Set path for where to save file
ActiveWorkbook.SaveAs FileName:= _
"\\my-drive\desktop\subfolder\report_folder\" & _ 'Update path to suit
"\Ad Hoc Report Name " & MonYear & ".xls" 'Update report name, timestamp, and filetype as needed

End Sub
Sub SendEmail()
'Originally sourced from Ron deBruin's site on VBA, customized to suit my workflows
Dim OutApp As Object
Dim OutMail As Object
'Dim fso As New Scripting.FileSystemObject
Dim strbody_link As String
Dim strbody_attach As String
Dim SigString As String
Dim Signature As String
Dim rng As Range
'Dim filename as String

If ActiveWorkbook.Path <> "" Then
Set OutApp = CreateObject("Outlook.Application")
Set OutMail = OutApp.CreateItem(0)

'MUST go to VBA Tools -> References -> check the box "Microsoft Scripting Runtime" for this to work
'FileName = fso.GetBaseName(ActiveWorkbook.Name) ' removes file extension from workbook name

'''' Update in order to copy some cells or a pivot into the body of the email
Set rng = Nothing 'Add additional XXrng variables as neeed
    On Error Resume Next
    'Set rng = Selection.SpecialCells(xlCellTypeVisible) 'Only the visible cells in the selection
    'You can also use a fixed range if you want
    Set rng = Sheets("SUMMARY").Range("A1").CurrentRegion 'CurrentRegion will snap to the boundries of all cells with data (like ctrl+A)
    On Error GoTo 0

'strbody is the message that will appear in the body of your email. Uses HTML formatting
'Defaults to include the report name as a hyperlink to the save location
'Ensure your save location is accessable by audience. Avoid attachments as they eat up Inbox space
strbody_link = "<font size=""3"" face=""Calibri"">" & _
"Hello,<br><br>Please click the link to view the report: " & _
"<A HREF=""file://" & ActiveWorkbook.FullName & _
""">" & ActiveWorkbook.Name & "</A>"

'Alternative strbody for attachments:
strbody_attach = "<font size=""3"" face=""Calibri"">" & _
"Hello,<br><br>Please find the latest " & ActiveWorkbook.Name & " file attached."

'Rename or create copy of your Outlook signature as "MySig"
'VBA will find your custom "MySig.htm" file and add it to the end of this email
'If no "MySig.htm" file is found, VBA will leave the signature blank
'Requires GetBoiler function below
    SigString = Environ("appdata") & _
                "\Microsoft\Signatures\MySig.htm"

    If Dir(SigString) <> "" Then
        Signature = GetBoiler(SigString)
    Else
    Signature = ""
    End If
    On Error Resume Next

With OutMail
.to = "" 'enter email addresses seperated by ";"
.CC = ""
.BCC = ""
.Subject = ActiveWorkbook.Name
.HTMLBody = strbody_link & "<br>" & Signature '& RangetoHTML(rng)
'.Attachments.Add ActiveWorkbook.FullName 'Option to attach the report (update strbody to remove "<A HREF </A>" link)
.Display
'Application.Wait (Now + TimeValue("0:00:02"))
'Application.SendKeys "%s" 'Wait and SendKeys will send the report after a 2 second delay
End With
On Error GoTo 0

Set OutMail = Nothing
Set OutApp = Nothing
Else
MsgBox "The ActiveWorkbook does not have a path, Save the file first."
End If
End Sub
  Function GetBoiler(ByVal sFile As String) As String
'This function is required to use personal outlook signature in email
Dim fso As Object
Dim ts As Object
Set fso = CreateObject("Scripting.FileSystemObject")
Set ts = fso.GetFile(sFile).OpenAsTextStream(1, -2)
GetBoiler = ts.readall
ts.Close
End Function
Function RangetoHTML(rng As Range)
'Allows for excel ranges to be pasted into email body with formatting
Dim fso As Object
Dim ts As Object
Dim TempFile As String
Dim TempWB As Workbook

TempFile = Environ$("temp") & "\" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"

rng.Copy
Set TempWB = Workbooks.Add(1)
With TempWB.Sheets(1)
.Cells(1).PasteSpecial Paste:=8
.Cells(1).PasteSpecial xlPasteValues, , False, False
.Cells(1).PasteSpecial xlPasteFormats, , False, False
.Cells(1).Select
Application.CutCopyMode = False
On Error Resume Next
.DrawingObjects.Visible = True
.DrawingObjects.Delete
On Error GoTo 0
End With

With TempWB.PublishObjects.Add( _
SourceType:=xlSourceRange, _
FileName:=TempFile, _
Sheet:=TempWB.Sheets(1).Name, _
Source:=TempWB.Sheets(1).UsedRange.Address, _
HtmlType:=xlHtmlStatic)
.Publish (True)
End With

Set fso = CreateObject("Scripting.FileSystemObject")
Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
RangetoHTML = ts.readall
ts.Close
RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
"align=left x:publishsource=")

TempWB.Close SaveChanges:=False

Kill TempFile

Set ts = Nothing
Set fso = Nothing
Set TempWB = Nothing
End Function
