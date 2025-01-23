Sub RemoveProtection()

Dim dialogBox As FileDialog
Dim sourceFullName As String
Dim sourceFilePath As String
Dim sourceFileName As String
Dim sourceFileType As String
Dim newFileName As Variant
Dim tempFileName As String
Dim zipFilePath As Variant
Dim oApp As Object
Dim FSO As Object
Dim xmlSheetFile As String
Dim xmlFile As Integer
Dim xmlFileContent As String
Dim xmlStartProtectionCode As Double
Dim xmlEndProtectionCode As Double
Dim xmlProtectionString As String

'Open dialog box to select a file
Set dialogBox = Application.FileDialog(msoFileDialogFilePicker)
dialogBox.AllowMultiSelect = False
dialogBox.Title = "Select file to remove protection from"

If dialogBox.Show = -1 Then
    sourceFullName = dialogBox.SelectedItems(1)
Else
    Exit Sub
End If

'Get folder path, file type and file name from the sourceFullName
sourceFilePath = Left(sourceFullName, InStrRev(sourceFullName, "\"))
sourceFileType = Mid(sourceFullName, InStrRev(sourceFullName, ".") + 1)
sourceFileName = Mid(sourceFullName, Len(sourceFilePath) + 1)
sourceFileName = Left(sourceFileName, InStrRev(sourceFileName, ".") - 1)

'Use the date and time to create a unique file name
tempFileName = "Temp" & Format(Now, " dd-mmm-yy h-mm-ss")

'Copy and rename original file to a zip file with a unique name
newFileName = sourceFilePath & tempFileName & ".zip"
On Error Resume Next
FileCopy sourceFullName, newFileName

If Err.Number <> 0 Then
    MsgBox "Unable to copy " & sourceFullName & vbNewLine _
        & "Check the file is closed and try again"
    Exit Sub
End If
On Error GoTo 0

'Create folder to unzip to
zipFilePath = sourceFilePath & tempFileName & "\"
MkDir zipFilePath

'Extract the files into the newly created folder
Set oApp = CreateObject("Shell.Application")
oApp.Namespace(zipFilePath).CopyHere oApp.Namespace(newFileName).items

'loop through each file in the \xl\worksheets folder of the unzipped file
xmlSheetFile = Dir(zipFilePath & "\xl\worksheets\*.xml*")
Do While xmlSheetFile <> ""

    'Read text of the file to a variable
    xmlFile = FreeFile
    Open zipFilePath & "xl\worksheets\" & xmlSheetFile For Input As xmlFile
    xmlFileContent = Input(LOF(xmlFile), xmlFile)
    Close xmlFile

    'Manipulate the text in the file
    xmlStartProtectionCode = 0
    xmlStartProtectionCode = InStr(1, xmlFileContent, "<sheetProtection")

    If xmlStartProtectionCode > 0 Then

        xmlEndProtectionCode = InStr(xmlStartProtectionCode, _
            xmlFileContent, "/>") + 2 '"/>" is 2 characters long
        xmlProtectionString = Mid(xmlFileContent, xmlStartProtectionCode, _
            xmlEndProtectionCode - xmlStartProtectionCode)
        xmlFileContent = Replace(xmlFileContent, xmlProtectionString, "")

    End If

    'Output the text of the variable to the file
    xmlFile = FreeFile
    Open zipFilePath & "xl\worksheets\" & xmlSheetFile For Output As xmlFile
    Print #xmlFile, xmlFileContent
    Close xmlFile

    'Loop to next xmlFile in directory
    xmlSheetFile = Dir

Loop

'Read text of the xl\workbook.xml file to a variable
xmlFile = FreeFile
Open zipFilePath & "xl\workbook.xml" For Input As xmlFile
xmlFileContent = Input(LOF(xmlFile), xmlFile)
Close xmlFile

'Manipulate the text in the file to remove the workbook protection
xmlStartProtectionCode = 0
xmlStartProtectionCode = InStr(1, xmlFileContent, "<workbookProtection")
If xmlStartProtectionCode > 0 Then

    xmlEndProtectionCode = InStr(xmlStartProtectionCode, _
        xmlFileContent, "/>") + 2 ''"/>" is 2 characters long
    xmlProtectionString = Mid(xmlFileContent, xmlStartProtectionCode, _
        xmlEndProtectionCode - xmlStartProtectionCode)
    xmlFileContent = Replace(xmlFileContent, xmlProtectionString, "")

End If

'Manipulate the text in the file to remove the modify password
xmlStartProtectionCode = 0
xmlStartProtectionCode = InStr(1, xmlFileContent, "<fileSharing")
If xmlStartProtectionCode > 0 Then

    xmlEndProtectionCode = InStr(xmlStartProtectionCode, xmlFileContent, _
        "/>") + 2 ''"/>" is 2 characters long
    xmlProtectionString = Mid(xmlFileContent, xmlStartProtectionCode, _
        xmlEndProtectionCode - xmlStartProtectionCode)
    xmlFileContent = Replace(xmlFileContent, xmlProtectionString, "")

End If

'Output the text of the variable to the file
xmlFile = FreeFile
Open zipFilePath & "xl\workbook.xml" & xmlSheetFile For Output As xmlFile
Print #xmlFile, xmlFileContent
Close xmlFile

'Create empty Zip File
Open sourceFilePath & tempFileName & ".zip" For Output As #1
Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
Close #1

'Move files into the zip file
oApp.Namespace(sourceFilePath & tempFileName & ".zip").CopyHere _
oApp.Namespace(zipFilePath).items
'Keep script waiting until Compressing is done
On Error Resume Next
Do Until oApp.Namespace(sourceFilePath & tempFileName & ".zip").items.Count = _
    oApp.Namespace(zipFilePath).items.Count
    Application.Wait (Now + TimeValue("0:00:01"))
Loop
On Error GoTo 0

'Delete the files & folders created during the sub
Set FSO = CreateObject("scripting.filesystemobject")
FSO.deletefolder sourceFilePath & tempFileName

'Rename the final file back to an xlsx file
Name sourceFilePath & tempFileName & ".zip" As sourceFilePath & sourceFileName _
& "_" & Format(Now, "dd-mmm-yy h-mm-ss") & "." & sourceFileType

'Show message box
MsgBox "The workbook and worksheet protection passwords have been removed.", _
vbInformation + vbOKOnly, Title:="Password protection"

End Sub