<html xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xsl:version="1.0">
	<head>
		<title>Claim your free gift</title>
		<!-- #include file="includes/styles.htm" -->
		<SCRIPT>
window.onload=init;
function init()
{
	for(i in document.links) document.links[i].onfocus = document.links[i].blur;
}

		</SCRIPT>
		<script runat=server language="vbscript">
Dim strTitle, strName, strLastName, strPhone, strEmail, strAddress1, strAddress2, strCity, strState, strZip
Dim strCountry, strComments, strGift1, strGift2, strColPurpose, strHearAbout
Dim strColName, strColLastName, strColTitle, strColEmail, strColAddr1, strColCity, strColState, strColCountry, strColHearAbout

function CheckAvailability ()	
	' If the store is closed, exit now
	Set fso = Server.CreateObject ("Scripting.FileSystemObject")
	if (fso.FileExists (Server.MapPath ("res/closed.flg"))) then Response.Redirect ("closed.asp")

	Set dbConn = Server.CreateObject ("ADODB.Connection")
	dbConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath ("res/POGData.mdb")
		
	
	Set objCommand = CreateObject("ADODB.Command") 
	objCommand.CommandText = "GetCatalogId"
	objCommand.CommandType = 4
	objCommand.Name = "GetCatalogId"
	
	Set objParam0 = objCommand.CreateParameter (, 200, 1, 256, "Bible")
	objCommand.Parameters.Append objParam0
		
	Set dbRst = Server.CreateObject ("ADODB.RecordSet")
	dbConn.Open
	
	objCommand.ActiveConnection = dbConn
	Set dbRst = objCommand.Execute
	
	dbRst.MoveFirst
	entryId = dbRst.Fields ("EntryId").Value
		
	dbRst.Close
	
	Set dbRst = Nothing
	Set objCommand = Nothing
		
	Set objCommand = CreateObject("ADODB.Command") 
	objCommand.CommandText = "GetCatalogCount"
	objCommand.CommandType = 4
	objCommand.Name = "GetCatalogCount"
	
	Set objParam1 = objCommand.CreateParameter (, 21, 1, , entryId)
	objCommand.Parameters.Append objParam1
	
	Set dbRst = Server.CreateObject ("ADODB.RecordSet")
	
	objCommand.ActiveConnection = dbConn
	Set dbRst = objCommand.Execute
	
	dbRst.MoveFirst
	count = dbRst.Fields ("Count").Value
	
	dbRst.Close
	Set dbRst = Nothing
	Set objCommand = Nothing
	dbConn.Close 
	Set dbConn = Nothing
	
	if (count <= 0) then Response.Redirect ("outoffreegift.asp")
end function

function InitValues ()
	strTitle = "<option selected>Select One</option><option>Mr.</option><option>Mrs.</option><option>Ms.</option>"
	strName = ""
	strLastName = ""
	strPhone = ""
	strEmail = ""
	strAddress1 = ""
	strAddress2 = ""
	strCity = ""
	strState = ""
	strZip = ""
	strCountry = ""
	strHearAbout = "<option selected>Select One</option><option>Web Surfing</optin><option>Radio (WHRI-Angel2)</option><option>Radio (KWHR-Angel3)</option><option>Radio (WHRA-Angel5)</option><option>Radio Africa</option><option>Referred by a friend</option><option>Linked from whr.org</option>"
	strComments = ""
	strGift1 = "<INPUT id=""Checkbox1"" type=""checkbox"" value=""ON"" name=""Gift1"">"
	<!-- strGift2 = "<INPUT id=""Checkbox2"" type=""checkbox"" value=""ON"" name=""Gift2"">" -->
	
	strColPurpose = "whitesmoke"
	strColName = "whitesmoke"
	strColLastName = "whitesmoke"
	strColTitle = "whitesmoke"
	strColEmail = "whitesmoke"
	strColAddr1 = "beige"
	strColCity = "beige"
	strColState = "beige"
	strColCountry = "beige"
	strColHearAbout = "beige"
end function

function CheckValues ()
	if (Request.QueryString <> "submit") then exit function
	
	strTitle = "<option>Select One</option><option>Mr.</option><option>Mrs.</option><option>Ms.</option>"
	nIndex = InStr (1, strTitle, Request.Form.Item ("Contact_Title"))
	if (nIndex <> 0) then
		strTitle = left (strTitle, nIndex - 2) & " selected" & mid (strTitle, nIndex - 1)
	end if
	
	strName = Request.Form.Item ("Contact_FirstName")
	strLastName = Request.Form.Item ("Contact_LastName")
	strPhone = Request.Form.Item ("Contact_Phone")
	strEmail = Request.Form.Item ("Contact_Email")
	strAddress1 = Request.Form.Item ("Contact_Address1")
	strAddress2 = Request.Form.Item ("Contact_Address2")
	strCity = Request.Form.Item ("Contact_City")
	strState = Request.Form.Item ("Contact_State")
	strZip = Request.Form.Item ("Contact_Zip")
	strCountry = Request.Form.Item ("Contact_Country")
	
	strHearAbout = "<option>Select One</option><option>Web Surfing</optin><option>Radio (WHRI-Angel2)</option><option>Radio (KWHR-Angel3)</option><option>Radio (WHRA-Angel5)</option><option>Radio Africa</option><option>Referred by a friend</option><option>Linked from whr.org</option>"
	nIndex = InStr (1, strHearAbout, Request.Form.Item ("Contact_HearAbout"))
	if (nIndex <> 0) then
		strHearAbout = left (strHearAbout, nIndex - 2) & " selected" & mid (strHearAbout, nIndex - 1)
	end if
	
	strComments = Request.Form.Item ("Comments")
	
	if (Request.Form.Item ("Gift1") = "ON") then 
		strGift1 = "<INPUT id=""Checkbox1"" type=""checkbox"" value=""ON"" name=""Gift1"" CHECKED>" 
	else
		strGift1 = "<INPUT id=""Checkbox1"" type=""checkbox"" value=""ON"" name=""Gift1"">" 
	end if

	
' if (Request.Form.Item ("Gift2") = "ON") then 
	'	strGift2 = "<INPUT id=""Checkbox2"" type=""checkbox"" value=""ON"" name=""Gift2"" CHECKED>"
	' else
	'	strGift2 = "<INPUT id=""Checkbox2"" type=""checkbox"" value=""ON"" name=""Gift2"">"
	' end if
	
	strError = "  "
	if (Request.Form.Item ("Contact_Title") = "Select One") then 
		strError = strError + "Title, "
		strColTitle = "#CC6633"
	end if
	if (strName = "") then 
		strError = strError + "Name, "
		strColName = "#CC6633"
	end if
	if (strLastName = "") then 
		strError = strError + "Name, "
		strColLastName = "#CC6633"
	end if
	if (strAddress1  = "") then 
		strError = strError + "Address, "
		strColAddr1 = "#CC6633"
	end if
	if (strCity  = "") then 
		strError = strError + "City, "
		strColCity = "#CC6633"
	end if
	if (strState  = "") then 
		strError = strError + "State, "
		strColState = "#CC6633"
	end if
	if (strCountry  = "") then 
		strError = strError + "Country, "
		strColCountry = "#CC6633"
	end if
	if (Request.Form.Item ("Contact_HearAbout") = "Select One") then
	    strError = strError + "How found, "
	    strColHearAbout = "#CC6633"
	end if
	strError = left (strError, Len (strError) - 2)
	
	if (strError <> "") then
		strError = "<font color=""red"" size=4><b>Please enter relevant information in the following fields: " + strError
		strError = strError + "</b></font>."
		Response.Write (strError)
	else
		strFrom = "webmaster@powerofgospel.org"
		strTo = "info@powerofgospel.org"
		strSubject = "Free gift claim"
		strBody = "Title: " & Request.Form.Item ("Contact_Title") & chr (13) & chr (10)
		strBody = strBody & "Name: " & Request.Form.Item ("Contact_FirstName") & " " & Request.Form.Item ("Contact_LastName") & chr (13) & chr (10)
		strBody = strBody & "Phone: " & Request.Form.Item ("Contact_Phone") & chr (13) & chr (10)
		strBody = strBody & "EMail: " & Request.Form.Item ("Contact_Email") & chr (13) & chr (10)
		strBody = strBody & "Address1: " & Request.Form.Item ("Contact_Address1") & chr (13) & chr (10)
		strBody = strBody & "Address2: " & Request.Form.Item ("Contact_Address2") & chr (13) & chr (10)
		strBody = strBody & "City and State: " & Request.Form.Item ("Contact_City") & "," & Request.Form.Item ("Contact_State") & chr (13) & chr (10)
		strBody = strBody & "Zip: " & Request.Form.Item ("Contact_Zip") & chr (13) & chr (10)
		strBody = strBody & "Country: " & Request.Form.Item ("Contact_Country") & chr (13) & chr (10)
		strBody = strBody & "How Found: " & Request.Form.Item ("Contact_HearAbout") & chr (13) & chr (10)
		strBody = strBody & "Comments: " & Request.Form.Item ("Comments") & chr (13) & chr (10)
		strBody = strBody & "Gift of Bible: " & Request.Form.Item ("Gift1") & chr (13) & chr (10)
		' strBody = strBody & "Jesus Video Casette: " & Request.Form.Item ("Gift2") & chr (13) & chr (10)
	
		Set objMsg = CreateObject("CDO.Message") 
		objMsg.Subject = strSubject 
		objMsg.Sender = strFrom 
		objMsg.To = strTo 
		objMsg.TextBody = strBody 
		objMsg.Send

		'objMsg.Send strFrom, strTo, strSubject, strBody
		
		'******************************************************************************************************************
		'insert the user record in the Users table
	    Set dbConn = Server.CreateObject ("ADODB.Connection")
		dbConn.Mode = 3
		
	    dbConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath ("res/POGData.mdb")
	    
	    		
		title = Request.Form.Item ("Contact_Title") 
		firstname = Request.Form.Item ("Contact_FirstName")
		lastname = Request.Form.Item ("Contact_LastName")
		addr1 = Request.Form.Item ("Contact_Address1") 
		addr2 = Request.Form.Item ("Contact_Address2") 
		city = Request.Form.Item ("Contact_City")
		state = Request.Form.Item ("Contact_State")
		country = Request.Form.Item ("Contact_Country") 
		zip = Request.Form.Item ("Contact_Zip") 		
		email = Request.Form.Item ("Contact_Email")
		ph = Request.Form.Item ("Contact_Phone")
		
    	
    	Set objCommand = CreateObject("ADODB.Command") 
		objCommand.CommandText = "InsertUser"
		objCommand.CommandType = 4
		objCommand.Name = "InsertUser"
		
		
		Set objParam0 = objCommand.CreateParameter (, 200, 1, 256, title)
		objCommand.Parameters.Append objParam0
		Set objParam1 = objCommand.CreateParameter (, 200, 1, 256, firstname)
		objCommand.Parameters.Append objParam1
		Set objParam2 = objCommand.CreateParameter (, 200, 1, 256, lastname)
		objCommand.Parameters.Append objParam2
		Set objParam3 = objCommand.CreateParameter (, 200, 1, 256, Addr1)
		objCommand.Parameters.Append objParam3		
		Set objParam4 = objCommand.CreateParameter (, 200, 1, 256, Addr2)
		objCommand.Parameters.Append objParam4		
		Set objParam5 = objCommand.CreateParameter (, 200, 1, 256, city)
		objCommand.Parameters.Append objParam5
		Set objParam6 = objCommand.CreateParameter (, 200, 1, 256, state)
		objCommand.Parameters.Append objParam6
		Set objParam7 = objCommand.CreateParameter (, 200, 1, 256, country)
		objCommand.Parameters.Append objParam7
		Set objParam8 = objCommand.CreateParameter (, 200, 1, 256, zip)
		objCommand.Parameters.Append objParam8
		Set objParam9 = objCommand.CreateParameter (, 200, 1, 256, email)
		objCommand.Parameters.Append objParam9
		Set objParam10 = objCommand.CreateParameter (, 200, 1, 256, ph)
		objCommand.Parameters.Append objParam10		
		
	    dbConn.Open
		src = "SELECT SourceId FROM InfoSource WHERE Source = '" & Request.Form.Item ("Contact_HearAbout") & "'"
		role = "SELECT RoleId from Roles WHERE Name = 'User'"		

		Set dbRst = Server.CreateObject ("ADODB.RecordSet")
		
		Set dbRst = dbConn.Execute(src)		
		dbRst.MoveFirst()
		srcId = dbRst.Fields ("SourceId").Value
		dbRst.Close 
		
		Set dbRst = dbConn.Execute(role)
		dbRst.MoveFirst()
		roleId = dbRst.Fields ("RoleId").Value
		dbRst.Close 
		
		Set objParam11 = objCommand.CreateParameter (, 21, 1, , srcId)
		objCommand.Parameters.Append objParam11
		Set objParam12 = objCommand.CreateParameter (, 21, 1, , roleId)
		objCommand.Parameters.Append objParam12				
				
		
		objCommand.ActiveConnection = dbConn
		objCommand.Execute
		
		'*******************************************************************************************************
		'Create order entry
    	Set objCom = CreateObject("ADODB.Command") 
		objCom.CommandText = "GetUserId"
		objCom.CommandType = 4
		
		
		Set objParam0 = objCom.CreateParameter (, 200, 1, 256, title)
		objCom.Parameters.Append objParam0
		Set objParam1 = objCom.CreateParameter (, 200, 1, 256, firstname)
		objCom.Parameters.Append objParam1
		Set objParam2 = objCom.CreateParameter (, 200, 1, 256, lastname)
		objCom.Parameters.Append objParam2
		Set objParam3 = objCom.CreateParameter (, 200, 1, 256, Addr1)
		objCom.Parameters.Append objParam3		
		Set objParam4 = objCom.CreateParameter (, 200, 1, 256, Addr2)
		objCom.Parameters.Append objParam4		
		Set objParam5 = objCom.CreateParameter (, 200, 1, 256, city)
		objCom.Parameters.Append objParam5
		Set objParam6 = objCom.CreateParameter (, 200, 1, 256, state)
		objCom.Parameters.Append objParam6
		Set objParam7 = objCom.CreateParameter (, 200, 1, 256, country)
		objCom.Parameters.Append objParam7
		Set objParam8 = objCom.CreateParameter (, 200, 1, 256, zip)
		objCom.Parameters.Append objParam8
		
		
		Set dbRst = CreateObject ("ADODB.RecordSet")
				
	    objCom.ActiveConnection = dbConn
		Set dbRst = objCom.Execute

		dbRst.MoveFirst
		userId = dbRst.Fields ("UserId").Value
		
		dbRst.Close


		Set objCommand = CreateObject("ADODB.Command") 
		objCommand.CommandText = "GetCatalogId"
		objCommand.CommandType = 4
		objCommand.Name = "GetCatalogId"
		
		Set objParam0 = objCommand.CreateParameter (, 200, 1, 256, "Bible")
		objCommand.Parameters.Append objParam0
			
		Set dbRst = Server.CreateObject ("ADODB.RecordSet")
	
		objCommand.ActiveConnection = dbConn
		Set dbRst = objCommand.Execute
				
		dbRst.MoveFirst
		entryId = dbRst.Fields ("EntryId").Value
		
		dbRst.Close
		Set dbRst = Nothing
		Set objCommand = Nothing
		
		statusQry = "SELECT StatusId from OrderStatus WHERE Value = 'New'"		
	
		Set dbRst = Server.CreateObject ("ADODB.RecordSet")
		
		Set dbRst = dbConn.Execute(statusQry)		
		dbRst.MoveFirst()
		statusId = dbRst.Fields ("StatusId").Value
		dbRst.Close 
		
		count = 1
		
		Set objCommand = CreateObject("ADODB.Command") 
		objCommand.CommandText = "InsertOrder"
		objCommand.CommandType = 4
		
		Set objParam11 = objCommand.CreateParameter (, 21, 1, , userId)
		objCommand.Parameters.Append objParam11	

		Set objParam12 = objCommand.CreateParameter (, 21, 1, , entryId)
		objCommand.Parameters.Append objParam12	
		
		Set objParam13 = objCommand.CreateParameter (, 21, 1, , statusId)
		objCommand.Parameters.Append objParam13	
		
		Set objParam14 = objCommand.CreateParameter (, 21, 1, , count)
		objCommand.Parameters.Append objParam14	

		objCommand.ActiveConnection = dbConn
		objCommand.Execute
		
		Set objCommand = Nothing
		
		'*******************************************************************************************************
	    
		
		' Now, reduce the count for the catalog by one
			
		Set objCommand = CreateObject("ADODB.Command") 
		objCommand.CommandText = "DecrementCount"
		objCommand.CommandType = 4
		objCommand.Name = "DecrementCount"
				
		Set objParam2 = objCommand.CreateParameter (, 21, 1, , count)
		objCommand.Parameters.Append objParam2	
		
		Set objParam2 = objCommand.CreateParameter (, 21, 1, , entryId)
		objCommand.Parameters.Append objParam2
		
		
		objCommand.ActiveConnection = dbConn
		objCommand.Execute
		
		Set objCommand = Nothing
		dbConn.Close 
		Set dbConn = Nothing	
		
	Response.Redirect ("submitfreegift.asp")
	end if
end function
		</script>
	</head>
	<body bgcolor="#ffffff" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
		<%CheckAvailability ()%>
		<%InitValues ()%>
		<font face="Verdana, Arial, Helvetica, sans-serif" size="2"></font>
		<!-- #include file="includes/topborder.htm" -->
		<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
			<tr>
				<td width="180" height="8" valign="top" bgcolor="#eeeeff"><br>
					<!-- #include file="includes/leftmenu.htm" -->
					<p>&nbsp;</p>
				</td>
				<td width="2" height="8" valign="top"><img src="img/1p.gif" width="1" height="100%"></td>
				<td height="8" valign="top">
					<table width="100%" border="0" cellspacing="0" cellpadding="5" height="261">
						<tr>
							<td valign="top" height="479">
								<table width="100%" border="0" cellspacing="0" cellpadding="5" height="442">
						                    <tr background="img/browngradient.gif"> 
						                      <td background="img/browngradient.gif" height="25"> 
						                        <p><font face="Verdana, Arial, Helvetica, sans-serif" size="3"><b><font size="4" 										color="#FFFFFF">Claim your free gift here</font></b></font></p>
						                      </td>
						                    </tr>

									<tr>
										<td valign="center">
											<table width="100%" height="100%" cellspacing="0" cellpadding="5" bordercolor="#339933">
												<tr>
													<td>
														<FORM METHOD="post" action="freegift.asp?submit" name="Contact_Form">
															<TABLE id="Table1" cellSpacing="1" cellPadding="5" width="100%" border="0">
																<tr>
																	<td>
														<font face="Verdana, Arial, Helvetica, sans-serif" size="2">
We would like to give you a free gift of your own copy of God's Holy Word, the Holy Bible in English. Please fill in the information needed below for us to send your free gift and select the gifts you would like to claim and click submit to receive this free gift. This offer is void where legally prohibited.
<p>
More than any free gift that we can give you, there is an even greater wonderful gift God has for you. It is the gift of Eternal life through savlation. The Bible says, <span style='color:red'>"All have sinned and have fallen short of the glory of God" (Romans 3:23)</span>. Our sin has separated us from God's love and His presence. But, God in His mercy has offered us a gift of restoring that relationship to Him from now through eternity. Bible says, <span style='color:red'>"The gift of God is eternal life in Christ Jesus our Lord". (Romans 6:23)</span>. You can receive that gift from Him right now where ever you are. It is very simple. Bible says, <span style='color:red'>"If you <b>confess</b> with your mouth the Lord Jesus and <b>believe</b> in your heart that God raised Him from the dead, you will be saved. For with the heart one believes unto righteousness, and with the mouth confession is made unto salvation". (Romans 10:9-10)</span>
</p>
<blockquote>
<li>Admit you are a sinner and are in need of this free gift.</li>
<li>Be willing to turn from your sins.</li>
<li>Believe in Jesus Christ as your personal savior.</li>
<li>By repeating the prayer below, invite Jesus to be the master of your life.</li>
</blockquote>
<p>
<b><span style='font-size:10.0pt;color:#993300'>Prayer:</span></b> Dear Lord Jesus: I confess that I am  a sinner and need your forgiveness. I believe you died on the cross for my sins. I want to turn away from my sins. I submit myself to you now. Come into my heart and my life. Help me to trust you and follow you every day of my life. Amen
</p>
<p>
If you want more information about salvation and eternal life, please indicate so in the comments below and we will be happy to provide you more information.
</p>
<p>
<b><span style='font-size:12.0pt;color:#993300'>Note:</span></b> Only one Bible per mailing address
</p>
</font>
																	</td>
																</tr>
																<tr>
																	<td class="TextDecoration2">
																		<%CheckValues ()%>
																	</td>
																</tr>
															</TABLE>
															<P>
																<TABLE id="Table2" cellSpacing="1" cellPadding="5" width="100%" border="0" height="141">
																	<TR bgcolor=<%Response.Write ("""" & strColTitle & """")%>>
																		<TD class="TextDecoration2" width="10"><font color="red">*</font></TD>
																		<TD class="TextDecoration2" width="174"><b>Title</b></TD>
																		<TD class="TextDecoration2"><SELECT id="Select2" name="Contact_Title">
																				<%Response.Write ("""" & strTitle & """")%>
																			</SELECT></TD>
																	</TR>
																	<TR bgcolor=<%Response.Write ("""" & strColName & """")%>>
																		<TD class="TextDecoration2" width="10"><font color="red">*</font></TD>
																		<TD class="TextDecoration2" width="174"><b> First Name</b></TD>
																		<TD class="TextDecoration2">
																			<INPUT TYPE="text" NAME="Contact_FirstName" SIZE="25" ID="Text1" value=<%Response.Write ("""" & strName & """")%>></TD>
																	</TR>
																	<TR bgcolor=<%Response.Write ("""" & strColLastName & """")%>>
																		<TD class="TextDecoration2" width="10"><font color="red">*</font></TD>
																		<TD class="TextDecoration2" width="174"><b> Last Name</b></TD>
																		<TD class="TextDecoration2">
																		<INPUT TYPE="text" NAME="Contact_LastName" SIZE="25" ID="Text2" value=<%Response.Write ("""" & strLastName & """")%>></TD>
																	</TR>

																	<TR bgcolor="whitesmoke">
																		<TD class="TextDecoration2" width="10"></TD>
																		<TD class="TextDecoration2" width="174">Phone</TD>
																		<TD class="TextDecoration2">
																			<INPUT TYPE="text" NAME="Contact_Phone" SIZE="25" MAXLENGTH="25" ID="Text3" value=<%Response.Write ("""" & strPhone & """")%>></TD>
																	</TR>
																	<TR bgcolor=<%Response.Write ("""" & strColEmail & """")%>>
																		<TD class="TextDecoration2" width="10"></TD>
																		<TD class="TextDecoration2" width="174">E-mail</TD>
																		<TD class="TextDecoration2">
																			<INPUT TYPE="text" NAME="Contact_Email" SIZE="25" ID="Text5" value=<%Response.Write ("""" & strEmail & """")%>></TD>
																	</TR>
																	
<TR bgcolor=<%Response.Write ("""" & strColAddr1 & """")%>>
																		<TD class="TextDecoration2" width="10"><font color="red">*</font></TD>
																		<TD class="TextDecoration2" width="174"><b>Address&nbsp; 1</b></TD>
																		<TD class="TextDecoration2">
																			<INPUT TYPE="text" NAME="Contact_Address1" SIZE="25" MAXLENGTH="25" ID="Text6" value=<%Response.Write ("""" & strAddress1 & """")%>><b>&nbsp; Bibles cannot be sent to P.O. Box Addresses</b></TD>
																	</TR>
																	
<TR bgcolor="beige">
																		<TD class="TextDecoration2" width="10"></TD>
																		<TD class="TextDecoration2" width="174">Address 2</TD>
																		<TD class="TextDecoration2">
																			<INPUT id="Text7" type="text" maxLength="25" size="25" name="Contact_Address2" value=<%Response.Write ("""" & strAddress2 & """")%>></TD>
																	</TR>
																	
<TR bgcolor=<%Response.Write ("""" & strColCity & """")%>>
																		<TD class="TextDecoration2" width="10"><font color="red">*</font></TD>
																		<TD class="TextDecoration2" width="174"><b>City</b></TD>
																		<TD class="TextDecoration2">
																			<INPUT id="Text8" type="text" maxLength="25" size="25" name="Contact_City" value=<%Response.Write ("""" & strCity & """")%>></TD>
																	</TR>
																	<TR bgcolor=<%Response.Write ("""" & strColState & """")%>>
																		<TD class="TextDecoration2" width="10"><font color="red">*</font></TD>
																		<TD class="TextDecoration2" width="174"><b> State</b></TD>
																		<TD class="TextDecoration2">
																			<INPUT id="Text4" type="text" maxLength="25" size="25" name="Contact_State" value=<%Response.Write ("""" & strState & """")%>></TD>
																	</TR>																	
<TR bgcolor="beige">
																		<TD class="TextDecoration2" width="10"></TD>
																		<TD class="TextDecoration2" width="174">Zip</TD>
																		<TD class="TextDecoration2">
																			<INPUT id="Text9" type="text" maxLength="25" size="25" name="Contact_Zip" value=<%Response.Write ("""" & strZip & """")%>></TD>
																	</TR>
																	
<TR bgcolor=<%Response.Write ("""" & strColCountry & """")%>>
																		<TD class="TextDecoration2" width="10"><font color="red">*</font></TD>
																		<TD class="TextDecoration2" width="174"><b>Country</b></TD>
																		<TD class="TextDecoration2">
																			<INPUT id="Text10" type="text" maxLength="25" size="25" name="Contact_Country" value=<%Response.Write ("""" & strCountry & """")%>></TD>
																	</TR>
																	<TR bgcolor=<%Response.Write ("""" & strColHearAbout & """")%>>
																		<TD class="TextDecoration2" width="10"><font color="red">*</font></TD>
																		<TD class="TextDecoration2" width="174"><b>How did you find this website?</b></TD>
																		<TD class="TextDecoration2"><SELECT id="Select1" name="Contact_HearAbout">
																				<%Response.Write ("""" & strHearAbout & """")%>
																			</SELECT></TD>
																	</TR>
																	<TR bgcolor="honeydew">
																		<TD class="TextDecoration2" width="10"></TD>
																		<TD class="TextDecoration2" width="174">Comments:</TD>
																		<TD class="TextDecoration2">
																			<TEXTAREA NAME="Comments" ROWS="5" COLS="35" ID="Textarea1" value=<%Response.Write ("""" & strComments & """")%>></TEXTAREA></TD>
																	</TR>
																	<TR bgcolor="honeydew">
																		<TD class="TextDecoration2" width="10"></TD>
																		<td></td>
																		<TD class="TextDecoration2">
																			 <%Response.Write (strGift1)%>&nbsp;Please send me my free copy of the Holy Bible</TD>
																	</TR>
<!-- 
																	<TR bgcolor="honeydew">
																		<TD class="TextDecoration2" width="10"></TD>
																		<td></td>
																		<TD class="TextDecoration2">

				In addition, for a limited time, we also offer an additional second free gift of Jesus film recorded on a VHS Video cassette for addresses in <B>United States and Canada only</B>. Claim that gift if you live in the United States of America or Canada.<BR> 																		 				<%Response.Write (strGift2)%>&nbsp;Please send me a copy of the Jesus Film on video</TD>
																	</TR>
-->
																	<TR>
																		<TD class="TextDecoration2" width="10"></TD>
																		<td></td>
																		<TD class="TextDecoration2">
																			<INPUT TYPE="submit" VALUE="Submit" ID="Submit1" NAME="Submit1">&nbsp; <INPUT TYPE="reset" VALUE="Reset" ID="Reset1" NAME="Reset1"></TD>
																	</TR>
																</TABLE>
															</P>
															&nbsp;
														</FORM>
													</td>
												</tr>
											</table>
											<BR>
										</td>
									</tr>
								</table>
							</td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
		<!-- #include file="includes/bottomborder.htm" -->
	</body>
</html>
