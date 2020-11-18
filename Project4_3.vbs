'--------------------------------------------------------------------
' mail.vbs
'--------------------------------------------------------------------
Dim strTo, strSubject, strTextBody, strFrom, strCC, strBCC
Dim strTextBodyFile, strAttachment1, strAttachment2

Dim objFSO
Dim objTextFile
Set objFSO = CreateObject("Scripting.FileSystemObject")

strSubject="Subject of the message"
strSMTPServer="mail.westernsydney.edu.au"
strFrom="18963162@student.westernsydney.edu.au"
strCC=""
strBCC=""
strAttachment1 = "C:\Users\Khodar El-Helou\Desktop\prac8\MyZippedFile.zip"

main

function main
	Dim emailList, emailRegex, emailMatch
	
	Set objExcel = CreateObject("Excel.Application")
	Set objWorkbook = objExcel.Workbooks.Open("C:\Users\Khodar El-Helou\Desktop\prac8\Attachments\records_2.xls")
	
	intRow = 2
	ReDim company(10)
	thisFileName = Left(objWorkbook.FullName , InStrRev(objWorkbook.FullName , ".") - 1)
	objWorkbook.Sheets.Select
	objWorkbook.ActiveSheet.ExportAsFixedFormat 0, thisFileName & ".pdf", 0, 1, 0,,,0
	
	''Read an Excel Spreadsheet
	Set oShell = CreateObject("shell.Application")
	Set file = objFSO.CreateTextFile("C:\Users\Khodar El-Helou\Desktop\prac8\MyZippedFile.zip", true)
	file.write("PK" & chr(5) & chr(6) & string(18,chr(0)))
	file.close
	
	Set zip = oShell.NameSpace("C:\Users\Khodar El-Helou\Desktop\prac8\MyZippedFile.zip")
	Set folder = objFSO.getFolder("C:\Users\Khodar El-Helou\Desktop\prac8\Project3DuplicatedFiles")
	
	Dim filecount
	filecount = 0
	
	for each file in folder.files
	if (Lcase(objFSO.getextensionname(file.name)) <> "vbs") AND _
		(Lcase(objFSO.getextensionname(file.name)) <> "hta") AND _
		(Lcase(objFSO.getextensionname(file.name)) <> "zip") then
		zip.CopyHere file.path
		filecount = filecount +1
		Wscript.Sleep 1000
	end if
	next
	
	Do Until zip.items.count = filecount
	Wscript.Sleep 1000
	Loop
	
	Do Until objExcel.Cells(intRow,1).Value = ""
		
		name = objExcel.Cells(intRow, 1).Value
		strTo = objExcel.Cells(intRow, 2).Value
		clicks = objExcel.Cells(intRow, 3).Value
		price = objExcel.Cells(intRow, 4).Value
		
		mail strFrom, price, clicks, strTo, name, strSMTPServer
		
		intRow = intRow + 1
		
	Loop
	
	objWorkbook.close False
	objExcel.Quit
	
	msgbox "Message sent successfully"
end function

sub mail(strFrom, price, clicks, strTo, name, strSMTPServer)

	Dim objMessage, objConfiguration, objConfigFields, totalPrice
	Set objConfiguration = CreateObject("CDO.Configuration")
	Set objConfigFields = objConfiguration.Fields

	objConfigFields.Item("http://schemas.microsoft.com/cdo/" & "configuration/sendusing") = 2
	objConfigFields.Item("http://schemas.microsoft.com/cdo/" & "configuration/smtpserverport") = 25
	objConfigFields.Item("http://schemas.microsoft.com/cdo/" & "configuration/smtpserver") = strSMTPServer
	objConfigFields.Update
	
	Set objMsg = CreateObject("CDO.Message")
	objMsg.Configuration = objConfiguration
	objMsg.From = strFrom
	objMsg.Subject = name
	Wscript.Echo clicks * price
	totalPrice = "BILL. You will need to pay $" & clicks * price & " at the end of the month. For having ads placed on our website."
	objMsg.TextBody = totalPrice
	objMsg.AddAttachment(strAttachment1)
	objMsg.To = strTo

	objMsg.Send
end sub

function ReadTextFile(strFileName)
	Dim objFSO
	Dim objTextFile
	Dim strReadLine
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objTextFile = objFSO.OpenTextFile(strFileName,1)
	do while not objTextFile.AtEndOfStream 
		strReadLine = strReadLine + objTextFile.ReadLine()
	loop
	objTextFile.Close()

	ReadTextFile = strReadLine
end function