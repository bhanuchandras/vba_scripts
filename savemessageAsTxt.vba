Public Sub SavesMsg_click()
	Dim OlApp As Outlook.Application
	Dim Inbox As Outlook.MAPIFolder
	Dim InboxIntems as Outlook.Items
	Dim Mailobject As Object
	Dim index As Integer
	Dim StrText As String
	Dim FNme As String
	Dim fso As Object
	Dim Fileout As Object
	Dim oMail As Outlook.MailItem
	
	index = 0
	Set OlApp = CreateObject ("Outlook.Application")
	Set Inbox = OlApp.GetNamespace("Mapi").GetDefaultFolder
	Set InboxIntems = Inbox.Items
	Set Fldr = Application.ActiveExplorer.CurrentFolder
	DirName = "C:\Temp\Junk"
	
	For Each item in Fldr.Items
		If InStr(itm.Subject, " Regex for looking Subject") Then
			SubTxt = itm.Subject
			sChr = "_"
			SubTxt = Replace(SubTxt,"'",sChr)
			SubTxt = Replace(SubTxt,"*",sChr)
			SubTxt = Replace(SubTxt,"/",sChr)
			SubTxt = Replace(SubTxt,"\",sChr)
			SubTxt = Replace(SubTxt,"?",sChr)
			SubTxt = Replace(SubTxt,">",sChr)
			SubTxt = Replace(SubTxt,"<",sChr)
			SubTxt = Replace(SubTxt,"|",sChr)
			SubTxt = Replace(SubTxt,Chr(34),sChr)
			FNme = DirName & SubTxt & ".txt"
			'MsgBox itm.Subject
			itm.SaveAs FNme, olTXT
		End If
	Next
End Sub
