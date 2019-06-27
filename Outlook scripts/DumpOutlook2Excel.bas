Attribute VB_Name = "Module1"
'Script to extract emails from outlook to Excel for data analysis
'Date: 27/06/2019

'Main script
Sub ExportMain()
ExportToExcel "C:\temp\test.xlsx", "test@3ds.com"
MsgBox "Process complete.", vbInformation + vbOKOnly, MACRO_NAME
End Sub

'Helper function to export to excel
Sub ExportToExcel(strFilename As String, strFolderPath As String)

Dim olkMsg As Object

Dim olkFld As Object

Dim excApp As Object

Dim excWkb As Object

Dim excWks As Object

Dim intRow As Long


If strFilename <> "" Then

If strFolderPath <> "" Then

Set olkFld = OpenOutlookFolder(strFolderPath)

If TypeName(olkFld) <> "Nothing" Then

Set excApp = CreateObject("Excel.Application")

Set excWkb = excApp.Workbooks.Add()

Set excWks = excWkb.ActiveSheet

'Write Excel Column Headers

With excWks

.Cells(1, 1) = "Subject"

.Cells(1, 2) = "Received"

.Cells(1, 3) = "Sender"

.Cells(1, 4) = "Body"

End With



EnumerateFolders olkFld, 2, excWks

Set olkMsg = Nothing

excWkb.SaveAs strFilename

excWkb.Close

Else

MsgBox "The folder '" & strFolderPath & "' does not exist in Outlook.", vbCritical + vbOKOnly, MACRO_NAME

End If

Else

MsgBox "The folder path was empty.", vbCritical + vbOKOnly, MACRO_NAME

End If

Else

MsgBox "The filename was empty.", vbCritical + vbOKOnly, MACRO_NAME

End If

Set olkMsg = Nothing

Set olkFld = Nothing

Set excWks = Nothing

Set excWkb = Nothing

Set excApp = Nothing

End Sub

'Helper function to open outlook folder
Public Function OpenOutlookFolder(strFolderPath As String) As Outlook.MAPIFolder

Dim arrFolders As Variant

Dim varFolder As Variant

Dim bolBeyondRoot As Boolean

On Error Resume Next

If strFolderPath = "" Then

Set OpenOutlookFolder = Nothing

Else

Do While Left(strFolderPath, 1) = ""

strFolderPath = Right(strFolderPath, Len(strFolderPath) - 1)

Loop

arrFolders = Split(strFolderPath, "")

For Each varFolder In arrFolders

Select Case bolBeyondRoot

Case False

Set OpenOutlookFolder = Outlook.Session.folders(varFolder)

bolBeyondRoot = True

Case True

Set OpenOutlookFolder = OpenOutlookFolder.folders(varFolder)

End Select

If Err.Number <> 0 Then

Set OpenOutlookFolder = Nothing

Exit For

End If

Next

End If

On Error GoTo 0

End Function

'Helper function to extract sender/receiver address
Function GetSMTPAddress(Item As Outlook.MailItem, intOutlookVersion As Integer) As String

Dim olkSnd As Outlook.AddressEntry

Dim olkEnt As Object

On Error Resume Next

Select Case intOutlookVersion

Case Is < 14

If Item.SenderEmailType = "EX" Then

GetSMTPAddress = SMTPEX(Item)

Else

GetSMTPAddress = Item.SenderEmailAddress

End If

Case Else

Set olkSnd = Item.Sender

If olkSnd.AddressEntryUserType = olExchangeUserAddressEntry Then

Set olkEnt = olkSnd.GetExchangeUser

GetSMTPAddress = olkEnt.PrimarySmtpAddress

Else

GetSMTPAddress = Item.SenderEmailAddress

End If

End Select

On Error GoTo 0

Set olkPrp = Nothing

Set olkSnd = Nothing

Set olkEnt = Nothing

End Function

Function GetOutlookVersion() As Integer

Dim arrVer As Variant

arrVer = Split(Outlook.Version, ".")

GetOutlookVersion = arrVer(0)

End Function

Function SMTPEX(olkMsg As Outlook.MailItem) As String

Dim olkPA As Outlook.PropertyAccessor

On Error Resume Next

Set olkPA = olkMsg.PropertyAccessor

SMTPEX = olkPA.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x5D01001E")

On Error GoTo 0

Set olkPA = Nothing

End Function

' Helper function to recursively enumerate through the Outlook folders/subfolders and populate the Excel worksheet
Function EnumerateFolders(ByVal oFolder As Outlook.Folder, ByVal row As Integer, ByRef excWks As Object) As Integer


 Dim folders As Outlook.folders
 Dim Folder As Outlook.Folder
 Dim foldercount As Integer
 Dim olkFld As Outlook.Folder
 Dim msg As Outlook.MailItem
 Dim intVersion As Integer
 intVersion = GetOutlookVersion()
   
 On Error Resume Next
 
 Set folders = oFolder.folders
 foldercount = folders.Count
 
 'Check if there are any folders below oFolder
 If foldercount Then
   For Each Folder In folders
     For Each msg In Folder.Items
       'Only export messages, not receipts or appointment requests, etc.
       If msg.Class = olMail Then
         'Add a row for each field in the message you want to export
         excWks.Cells(row, 1) = msg.subject
         excWks.Cells(row, 2) = msg.ReceivedTime
         excWks.Cells(row, 3) = GetSMTPAddress(msg, intVersion)
         excWks.Cells(row, 4) = msg.Body
         row = row + 1
       End If
     Next
    row = row + EnumerateFolders(Folder, row, excWks)
    
  Next
  End If

  result = row
    
End Function





