<html>
<head>
<title>Sending Mail...</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body>
<%
	'Response.Write("dfgdjkfdjg")

	fnm=request("FirstName")
	lnm=request("LastName")
	title=request("Title")
	company=request("Name")
	addr1=request("Street")
	addr2=request("Street2")
	city=request("City")
	zip=request("ZipCode")
	ph=request("PhoneNumber")
	fax=request("FaxNumber")
	email=request("Email")
	cust_type=request("CustType")
	user=request("Users")
	
	appli1=request("ApplicationUsing1")
		if request("ApplicationUsing1")= "" then appli1="No" else appli1="Yes" end if
	appli2=request("ApplicationUsing2")
		if request("ApplicationUsing2")= "" then appli2="No" else appli2="Yes" end if
	appli3=request("ApplicationUsing3")
		if request("ApplicationUsing3")= "" then appli3="No" else appli3="Yes" end if
	appli4=request("ApplicationUsing4")
		if request("ApplicationUsing4")= "" then appli4="No" else appli4="Yes" end if
	appli5=request("ApplicationUsing5")
		if request("ApplicationUsing5")= "" then appli5="No" else appli5="Yes" end if
	appli6=request("ApplicationUsing6")
		if request("ApplicationUsing6")= "" then appli6="No" else appli6="Yes" end if
	
	domino=request("LotusFaxForDomino")
		
	email_program=request("EmailProgram")
		
	interest1=request("InterestedIn1")
		if request("InterestedIn1")= "" then interest1="No" else interest1="Yes" end if
	interest2=request("InterestedIn2")
		if request("InterestedIn2")= "" then interest2="No" else interest2="Yes" end if
	interest3=request("InterestedIn3")
		if request("InterestedIn3")= "" then interest3="No" else interest3="Yes" end if
	interest4=request("InterestedIn4")
		if request("InterestedIn4")= "" then interest4="No" else interest4="Yes" end if
		
	notes=request("Notes")
	
	mail_contents="First Name :" &fnm  &vbCrLf &_
	              "Last Name :" &lnm   &vbCrLf &_
				  "Title :" &title  &vbCrLf &_
				  "Conmpany :" &company   &vbCrLf &_
				  "Street Address Line1 :" &addr1    &vbCrLf &_
				  "Street Address Line2 :" &addr2    &vbCrLf &_
				  "City :" &city  &vbCrLf &_
				  "Zip/Postal Code :" &zip   &vbCrLf &_
				  "Business Phone :" &ph   &vbCrLf &_
				  "Business Fax :" &fax  &vbCrLf &_
				  "Email  :" &email  &vbCrLf &_
				  "Are you a ...? :" &cust_type  &vbCrLf &_
				  "How many users on your network :" &user  &vbCrLf &_
				  "Applications Used by the Company ----- :"   &vbCrLf &_
				  "Baan :" &appli1  &vbCrLf &_
				  "FileNET :" &appli2   &vbCrLf &_
				  "Oracle :" &appli3   &vbCrLf &_
				  "PeopleSoft :" &appli4   &vbCrLf &_
				  "SAP :" &appli5   &vbCrLf &_
				  "Siebel :" &appli6   &vbCrLf &_
				  "Are you a existing user of Lotus' Fax for Domino ? :" &domino  &vbCrLf &_
				  "What e-mail programs does your company currently use? :" &email_program  &vbCrLf &_
				  "Interested in ----- :"   &vbCrLf &_
				  "RightFax :" &interest1  &vbCrLf &_
				  "CallXpress :" &interest2   &vbCrLf &_
				  "Brooktrout :" &interest3   &vbCrLf &_
				  "HP 9100C DS :" &interest4   &vbCrLf &_
				  "Note :" &notes
				  
	'Response.Write(mail_contents)
	'Response.end
	Dim objMessage
	Set objMessage = Server.CreateObject("CDONTS.Newmail")
	objMessage.To       = "test@rincon.co.in"
	objMessage.From     = email
	objMessage.Subject  = "Feedback"
	objMessage.Body = mail_contents
	
	On Error Resume Next 
	objMessage.Send
	If Err <> 0 Then 
	Response.Write "Error encountered: " & Err.Description 
	End If 

	Set objMessage = Nothing
	
	Response.Redirect "index.htm"
	
%>
</body>
</html>
