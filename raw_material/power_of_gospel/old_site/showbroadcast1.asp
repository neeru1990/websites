<html xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xsl:version="1.0">
<head>
<title>Radio Broadcast</title> 
<!-- #include file="includes/styles.htm" -->
<script language="JavaScript" type="text/javascript">
	function submit_form() {
		if(document.radio_email.sender_name.value == '')
			alert("Please give us your name so your friend can know who you are.");
		else if(document.radio_email.email.value == '')
			alert("Please make sure to provide your friend's email to send the broadcast schedule.")
		else
			document.radio_email.submit();
	}
</script>
<SCRIPT>
window.onload=init;
function init()
{
	for(i in document.links) document.links[i].onfocus = document.links[i].blur;
}
</SCRIPT>

<script language="vbscript" runat="server">
Dim strLocation

function InitStrings ()
	strLocation = Request.Form.Item ("location")
	if (strLocation = "") then strLocation = "San Francisco, USA"
end function

function FillLocations ()
	' strQuery = "select * from Locations order by Location"
	strQuery = "GetLocationNames"
	Set dbConn = Server.CreateObject ("ADODB.Connection")
	dbConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath ("res/POGData.mdb")
	dbConn.Open

	Set dbRst = Server.CreateObject ("ADODB.RecordSet")
	dbRst.Open strQuery, dbConn
	dbRst.MoveFirst

	while not dbRst.EOF
		if (dbRst.Fields ("Location").Value = strLocation) then Response.Write ("<option selected>") else Response.Write ("<option>")
		Response.Write (dbRst.Fields ("location").Value)
		Response.Write ("</option>")
		dbRst.MoveNext
	wend 
	dbRst.Close
	dbConn.Close 
end function

function FillBroadcasts ()
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
		Response.Write ("<tr class=TextDecoration>")
		Response.Write ("<td>")
		Response.Write (dbRst.Fields ("StationCallLetters").Value)
		Response.Write ("</td>")
		Response.Write ("<td>")
		Response.Write (dbRst.Fields ("Frequency").Value)
		Response.Write (" KHz")
		Response.Write ("</td>")
		Response.Write ("<td>")
		' *** TO DO: Add logic to calculate correct day of the week here ***
		Response.Write (strDay (dbRst.Fields ("dayoftheweek").Value))
		Response.Write ("</td>")
		Response.Write ("<td>")
		' If the timezone adjusts forward, display the forward time otherwise the backward time
		if (dbRst.Fields ("GMTAdjustmentForward").Value = true) then
			Response.Write (dbRst.Fields ("StandardLocalForwardTime").Value)
		else
			Response.Write (dbRst.Fields ("StandardLocalBackwardTime").Value)
		end if
		' If the timezone uses daylight savings, repeat the same for daylight
		if (dbRst.Fields ("FollowsDST").Value = true) then
			Response.Write (" (Standard)<br>")
			if (dbRst.Fields ("GMTAdjustmentForward").Value = true) then
				Response.Write (dbRst.Fields ("DaylightLocalForwardTime").Value)
			else
				Response.Write (dbRst.Fields ("DaylightLocalBackwardTime").Value)
			end if
			Response.Write (" (Daylight)</td>")
		end if
		' Next row on the display table
		Response.Write ("</tr>")
		' Get the next record
		dbRst.MoveNext
	wend 
	' Close the record set and the database connection
	dbRst.Close
	dbConn.Close 
end function
</script>
	</head>
	<body bgcolor="#ffffff" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
		<% InitStrings () %>
		<font face="Verdana, Arial, Helvetica, sans-serif" size="2"></font>
		<!-- #include file="includes/topborder.htm" -->
		<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0" ID="Table1">
			<tr>
				<td width="180" height="8" valign="top" bgcolor="#eeeeff"><br>
					<!-- #include file="includes/leftmenu.htm" -->
					<p>&nbsp;</p>
				</td>
				<td width="2" height="8" valign="top"><img src="img/1p.gif" width="1" height="100%"></td>
				<td height="8" valign="top">
					<table width="100%" border="0" cellspacing="0" cellpadding="5" height="261" ID="Table2">
						<tr>
							<td valign="top" height="479">
								<table width="100%" border="0" cellspacing="0" cellpadding="5" height="442" ID="Table3">
									<tr>
										<td valign="center">
											<table width="100%" height="100%" cellspacing="0" cellpadding="5" bordercolor="#339933" ID="Table4">
												<tr background="img/browngradient.gif">
													<td background="img/browngradient.gif" height="25">
														<p><font face="Verdana, Arial, Helvetica, sans-serif" size="3"><b><font size="4" color="#ffffff">Radio 
																		Broadcast</font></b></font></p>
													</td>
												</tr>
												<tr>
													<td class="TextDecoration">
														<font color="red">
															<br>
															<b>And Jesus said to them, "Go into all the world and preach the gospel to every 
																creature" - Mark 16:15</b></font>
													</td>
												</tr>
												<tr>
													<td class="textdecoration">
														<div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="4"><b><IMG src="img/world.gif"></b></font></div>
														<P>
															<br>
															<br>
															Our radio program titled, "Power of Gospel time" airs weekly on shortwave radio stations and can be heard in several countries around the world. Following are the call letters of the various stations, frequencies and timings of broadcast. Choose any location in the list and select "Go" to get a list of broadcasts in that area. Introduce these broadcasts to your friends.<br><br>
You can also listen to the broadcasts on demand here on the website. <a href="radio.asp?start=">Click here to hear previous broadcasts</a>. 
														</P>
														<P>
															<form METHOD="post" action="showbroadcast.asp" name="Radio_form" ID="Form1">
																Power of Gospel broadcasts in 
																<select class="TextDecoration4" name="location">
																	<%FillLocations ()%>
																</select>
																<input class="TextDecoration4" name="submit" type="submit" value=" Go "/>
															</form>
														</P>
														<P>
															<table width="100%" border="1" cellspacing="0" bordercolor=#ffcc66>
																<tr bgcolor=#ffcc66>
																	<td class="TextDecoration"><b>Station</b></td>
																	<td class="TextDecoration"><b>Frequency</b></td>
																	<td class="TextDecoration"><b>Day of the Week</b></td>
																	<td class="TextDecoration"><b>Start Time</b></td>
																</tr>
																<%FillBroadcasts ()%>
															</table>
														</P>
														<p>
														<br>
														<br>
															<form METHOD="post" action="email.asp" name="radio_email" ID="Form2">
																<b>Send broadcast schedule to a friend</b><br>
																You can introduce our radio program to your friends. Enter your name, the email address of your friend and choose the city closest to your friend’s home town below. We will send an email on your behalf to your friend.<br><br>
																<table class="TextDecoration4"><tr><td>Your name: &nbsp;&nbsp;</td>
																	<td><input class="TextDecoration4" name="sender_name" type="text" size="25"></td></tr>
																<tr><td>Friend's email id: &nbsp;&nbsp;</td>
																	<td><input class="TextDecoration4" name="email" type="text" size="25"></td></tr>
																<tr><td>Friend's location: &nbsp;&nbsp;</td>
																<td><select class="TextDecoration4" name="location">
																	<%FillLocations ()%>
																</select></td></tr>
																<tr><td><INPUT TYPE="button" class="TextDecoration4" VALUE="Submit" ID="Submit1" NAME="Submit1" onclick="javascript:submit_form();" >&nbsp;
																<INPUT TYPE="reset" class="TextDecoration4" VALUE="Reset" ID="Reset1" NAME="Reset1"></td></tr></table>
															</form>
														</p>
													</td>
												</tr>
											</table>
											<br>
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