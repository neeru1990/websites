<html>
<head>
<title>Online Demo</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<SCRIPT>
<!--
	function cancelFunc()
	{
		history.go(-1)
	}
//-->
</SCRIPT>
</head>

<body>
<%
	
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
	exchange=Request("ChkE")
	lotus=Request("ChkL")
	oracle=Request("ChkO")
	
	if exchange = "on"or lotus = "on" or oracle = "on" then

		if exchange = "on" then sub_from = "Microsoft Exchange"
		if lotus = "on" then sub_from = sub_from & " Lotus Notes"
		if oracle = "on" then sub_from = sub_from & " Oracle"
		
		mail_contents="First Name:" &fnm  &vbCrLf &_
		              "Last Name :" &lnm   &vbCrLf &_
									"Title     :" &title  &vbCrLf &_
									"Company   :" &company   &vbCrLf &_
									"Address 1 :" &addr1    &vbCrLf &_
									"Address 2 :" &addr2    &vbCrLf &_
									"City      :" &city  &vbCrLf &_
									"Zip/Postal:" &zip   &vbCrLf &_
									"Busi Phone:" &ph   &vbCrLf &_
									"Fax       :" &fax  &vbCrLf &_
									"Email     :" &email  &vbCrLf &_
									"You are a/an...? :" &cust_type
				  
		Dim objMessage
		Set objMessage = Server.CreateObject("CDONTS.Newmail")
		objMessage.To       = "sales@rincon.co.in"
		objMessage.From     = email
		objMessage.Subject  = sub_from
		objMessage.Body		  = mail_contents
		
		On Error Resume Next 
		objMessage.Send
		If Err <> 0 Then Response.Write "Error encountered: " & Err.Description 
	
		Set objMessage = Nothing %>
		
		<TABLE width="400" border="1" align="center">
			<TR>
				<TD align="left">
					To view our RightFax demo, you will need the Macromedia Flash plug-in to view the demo. If you do not have the Flash plug-in, visit the <A href='http://www.macromedia.com/shockwave/download/frameset.fhtml?P1_Prod_Version=ShockwaveFlash' target='_blank'>	Macromedia Web site</A> to install it and then return to continue.
				</TD>
			</TR>
			<TR>
				<TD align="center">Online Demos</TD>
			</TR>
			<% if exchange = "on" then %>
			<tr>
				<td><a href='ex_demo.htm'>Exchange Demo<a></td>
			</tr>
			<%if lotus = "on" then%>
			<tr>
				<td><a href='lo_demo.htm'>Lotus Notes Demo<a></td>
			</tr>
			<%if oracle = "on" then%>
			<tr>
				<td><a href='or_demo.htm'>Oracle Demo<a></td>
			</tr>
		</TABLE>
		<%
	else
		Response.write "Please select the demo you would like view.<br>Click <A href='#' onclick='cancelFunc()'>here</A> to go back to form page"
	end if
		%>
</body>
</html>
