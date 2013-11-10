<html xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xsl:version="1.0">
	<head>
		<title>Word for the week</title>
		<script language="vbscript" runat="server">
		strEmail = Request.Form.Item ("email")
		strLocation = Request.Form.Item ("location")
		strSender = Request.Form.Item ("sender_name")
		strFrom = "webmaster@powerofgospel.org"
		strSubject = "Introducing radio program"
		strBody = strSender & " would like to introduce you to a Christian radio broadcast available on the Internet and on shortwave radio. It is a weekly broadcast in English that you can listen to at the following frequencies and times at your location, " & strLocation & ".<br><br>"
		strBody = strBody & "<table width='100%' border='1' cellspacing='0' bordercolor=#ffcc66><tr bgcolor=#ffcc66>"
		strBody = strBody & "<td class='TextDecoration'><b>Station</b></td><td class='TextDecoration'><b>Frequency</b></td>"
		strBody = strBody & "<td class='TextDecoration'><b>Day of the Week</b></td><td class='TextDecoration'><b>Start Time</b></td></tr>"
			' Query for broadcast schedule
			strQuery = "BroadcastsPerLocation" 
			dim strDay (8)
			strDay (1) = "Sunday"
			strDay (2) = "Monday"
			strDay (3) = "Tuesday"
			strDay (4) = "Wednesday"
			strDay (5) = "Thursday"
			strDay (6) = "Friday"
			strDay (7) = "Saturday"
			' Create a database connection and open it
			Set dbConn = Server.CreateObject ("ADODB.Connection")
			dbConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" &  Server.MapPath ("res/POGData.mdb")
			dbConn.Open
			' Create a command object and stuff a parameter into it - Parameter is the location
			Set objCommand = CreateObject("ADODB.Command") 
			objCommand.CommandText = strQuery
			objCommand.CommandType = 4
			' Parameter is var-char type with no name and length is maximum of 256 characters
			Set objParam = objCommand.CreateParameter (, 200, 1, 256, strLocation)
			objCommand.Parameters.Append objParam
			objCommand.Name = "GetBroadcasts"
			objCommand.ActiveConnection = dbConn
			' Execute the stored procedure query with the query parameter	
			Set dbRst = Server.CreateObject ("ADODB.RecordSet")
			Set dbRst = objCommand.Execute
			' Go to the beginning of the record Set
			dbRst.MoveFirst
			' Display the results in a table
			while not dbRst.EOF
				strBody = strBody & "<tr class=TextDecoration>"
				strBody = strBody & "<td>"
				strBody = strBody & (dbRst.Fields ("StationCallLetters").Value)
				strBody = strBody & "</td>"
				strBody = strBody & "<td>"
				strBody = strBody & (dbRst.Fields ("Frequency").Value)
				strBody = strBody & " KHz"
				strBody = strBody & "</td>"
				strBody = strBody & "<td>"
				' *** TO DO: Add logic to calculate correct day of the week here ***
				strBody = strBody & (strDay (dbRst.Fields ("dayoftheweek").Value))
				strBody = strBody & "</td>"
				strBody = strBody & "<td>"
				' If the timezone adjusts forward, display the forward time otherwise the backward time
			if (dbRst.Fields ("GMTAdjustmentForward").Value = true) then
				strBody = strBody & (dbRst.Fields ("StandardLocalForwardTime").Value)
			else
				strBody = strBody & (dbRst.Fields ("StandardLocalBackwardTime").Value)
			end if
			' If the timezone uses daylight savings, repeat the same for daylight
			if (dbRst.Fields ("FollowsDST").Value = true) then
				strBody = strBody & " (Standard)<br>"
				if (dbRst.Fields ("GMTAdjustmentForward").Value = true) then
					strBody = strBody & (dbRst.Fields ("DaylightLocalForwardTime").Value)
				else
					strBody = strBody & (dbRst.Fields ("DaylightLocalBackwardTime").Value)
				end if
				strBody = strBody & " (Daylight)</td>"
			end if
			' Next row on the display table
			strBody = strBody & "</tr>"
			' Get the next record
			dbRst.MoveNext
		wend 
	' Close the record set and the database connection
	dbRst.Close
	dbConn.Close
		strBody = strBody & "</table><br>In addition, you may also listen to the broadcast over the Internet by visiting <a href='http://www.powerofgospel.org'>http://www.powerofgospel.org</a>. You can also listen to archived radio programs from this website. We pray and believe that the Word of God brings joy and comfort into your life."
		Set objMsg = CreateObject("CDO.Message") 
		objMsg.Subject = strSubject 
		objMsg.Sender = strFrom 
		objMsg.To = strEmail 
		objMsg.HTMLBody = strBody 
		objMsg.Send
		
		function email()
			Response.Write ("Thank you for sending the broadcast schedule to your friend.<br>")
			Response.Write ("<a href=" & Session("strHome") & ">Click here to continue to browse the website.</a>")
		end function
		</script>
	</head>
<body bgcolor="#ffffff" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
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
									<tr>
										<td valign="center">
											<table width="100%" height="100%" cellspacing="0" cellpadding="5" bordercolor="#339933">
												<tr>
													<td align="center">
														<font face="Verdana, Arial, Helvetica" size=4><b><%Email()%></b></font>
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