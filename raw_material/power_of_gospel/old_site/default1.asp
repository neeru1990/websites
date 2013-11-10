<html xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xsl:version="1.0">
	<head>
		<title>Power of Gospel, Ministries</title>
		<!-- #include file = "includes/styles.htm" -->
		<SCRIPT>
window.onload=init;
function init()
{
	for(i in document.links) document.links[i].onfocus = document.links[i].blur;
}

		</SCRIPT>
		<script language="vbscript" runat="server">

function FillTodaysPromise ()
	Set fso = Server.CreateObject ("Scripting.FileSystemObject")
	Set f = fso.OpenTextFile (Server.MapPath ("res/promise.dat"))
	Dim strPromise, strReference
	
	strNum = f.ReadLine
	
	for i = 0 to (Day (Date) mod CInt (strNum))
		strPromise = f.ReadLine
		strReference = f.ReadLine
	next
	Response.Write ("<tr><td class=""TextDecoration"" height=""1"" bgcolor=""#EEEEFF""><p align=""center"">")
	Response.Write (strPromise)
	Response.Write ("</p></td></tr><tr><td height=""2"" bgcolor=""#EEEEFF""><div align=""center""><font face=""Verdana, Arial, Helvetica, sans-serif"" size=""2""><b><font size=""1"" color=""#FF0000"">")
	Response.Write (strReference)
	Response.Write ("</font></b></font></div></td></tr>")
	f.close
end function

function FillAudioSermon ()
	Set fso = Server.CreateObject ("Scripting.FileSystemObject")
	Set f = fso.OpenTextFile (Server.MapPath ("audiosermons/audioindex.dat"))
	
	strTitle = f.ReadLine
	strDescription = f.ReadLine
	strLink = f.ReadLine
	
	f.close
		
	Response.Write ("<font color=#993300 size=3><b>")
	Response.Write (strTitle)
	Response.Write ("</b></font><br>")
	Response.Write (strDescription)
	Response.Write ("<div align=""right""><a href=audiosermons/" & strLink & "><font face=""Verdana, Arial, Helvetica, sans-serif"" size=""1""><b><font color=""#FF0000"">Listen</font></b></font></a> <a href=audiosermons/" & strLink & "><img src=""img/cdrom.gif"" border=none width=""16"" height=""16"" align=""top""></a><br>")
	Response.Write ("<a href=""audiosermon.asp""><font face=""Verdana, Arial, Helvetica, sans-serif"" size=""1""><b><font color=""#FF0000"">Other Sermons</font></b></font></a> <a href=""audiosermon.asp""><img src=""img/more.gif"" border=none width=""12"" height=""12"" align=""center""></a></div>")
end function

function FillWordForWeek ()
	dateToCheck = Date
	fileName = FormatDateSpecial (dateToCheck)
	fileName = fileName & ".txt"
	Set fso = Server.CreateObject ("Scripting.FileSystemObject")
	while (fso.FileExists (Server.MapPath ("word/" & fileName)) = false)
		dateToCheck = DateAdd ("d", -1, dateToCheck)
		fileName = FormatDateSpecial (dateToCheck)
		fileName = fileName & ".txt"
	Wend
	
	Set f = fso.OpenTextFile (Server.MapPath ("word/" & fileName))
	strTitle = f.ReadLine
	strSynopsis = f.ReadLine
	f.close
	Response.Write ("<font color=#993300 size=3><b>")
	Response.Write (strTitle)
	Response.Write ("</b></font><br>")
	Response.Write (strSynopsis)
	Response.Write ("<div align=""right""><b><a href=""word.asp?" & fileName & """><font face=""Verdana, Arial, Helvetica, sans-serif"" size=""1""><b><font color=""#FF0000"">Read</font></b></font></a> <a href=""word.asp?" & fileName & """><img src=""img/more.gif"" border=none width=""12"" height=""12"" align=""center""></a><br>")
	Response.Write ("<a href=""wordtoday.asp""><font face=""Verdana, Arial, Helvetica, sans-serif"" size=""1"" color=""#FF0000"">Previous weeks' words</font></a></b> <a href=""wordtoday.asp""><img src=""img/more.gif"" border=none width=""12"" height=""12"" align=""center""></a></b></div>")
end function

function FillRadio ()
	dateToCheck = Date
	fileName = FormatDateSpecial (dateToCheck)
	fileName = fileName & ".dat"
	Set fso = Server.CreateObject ("Scripting.FileSystemObject")
	while (fso.FileExists (Server.MapPath ("radio/" & fileName)) = false)
		dateToCheck = DateAdd ("d", -1, dateToCheck)
		fileName = FormatDateSpecial (dateToCheck)
		fileName = fileName & ".dat"
	Wend
	
	Set f = fso.OpenTextFile (Server.MapPath ("radio/" & fileName))
	strTitle = f.ReadLine
	strSynopsis = f.ReadLine
	strLink = f.ReadLine
	f.close
	Response.Write ("<font color=#993300 size=3><b>")
	Response.Write (strTitle)
	Response.Write ("</b></font><br>")
	Response.Write (strSynopsis)
	Response.Write ("<div align=""right""><b>")
	if (Len (strLink) > 2) then
		Response.Write ("<a href=""radio/broadcasts/" & strLink & """><font face=""Verdana, Arial, Helvetica, sans-serif"" size=""1""><b><font color=""#FF0000"">Listen</font></b></font></a> <a href=""radio/broadcasts/" & strLink & """><img src=""img/more.gif"" border=none width=""12"" height=""12"" align=""center""></a><br>")
		Response.Write ("<a href=""radio.asp""><font face=""Verdana, Arial, Helvetica, sans-serif"" size=""1"" color=""#FF0000"">Previous broadcasts</font></a></b> <a href=""radio.asp""><img src=""img/more.gif"" border=none width=""12"" height=""12"" align=""center""></a>")
	end if
	Response.Write ("<br><a href=""showbroadcast.asp""><b><font face=""Verdana, Arial, Helvetica, sans-serif"" size=""1"" color=""#FF0000"">Radio broadcast coverage</font></a></b> <a href=""showbroadcast.asp""><img src=""img/more.gif"" border=none width=""12"" height=""12"" align=""center""></a>")
	Response.Write ("</div>")
end function

function FormatDateSpecial (formatDate)
	strFormat = Year (formatDate)
	
	if (Month (formatDate) < 10) then strFormat = strFormat & "0"
	strFormat = strFormat & Month (formatDate)
	
	if (Day (formatDate) < 10) then strFormat = strFormat & "0"
	strFormat = strFormat & Day (formatDate)
	
	FormatDateSpecial = strFormat
end function

function FillEvents ()
	Set fso = Server.CreateObject ("Scripting.FileSystemObject")
	if (not fso.FileExists (Server.MapPath ("res/events.txt"))) then exit function
	
	Set f = fso.OpenTextFile (Server.MapPath ("res/events.txt"))
	Dim strFirst, strSecond

	strTitle = f.ReadLine
	
	Response.Write ("<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"" height=""144""><tr><td valign=""center"" width=""16""><img src=""img/left.gif"" width=""16"" height=""100%"">")
	Response.Write ("</td><td valign=""center"" width=""90%"" bgcolor=""#3366CC""><div align=""center""><font face=""Verdana, Arial, Helvetica, sans-serif"" size=""2""><b><font color=""#FFFFFF"">" & strTitle & "</font></b></font>")
	Response.Write ("</div></td><td valign=""center"" width=""16""><img src=""img/right.gif"" width=""16"" height=""100%""></td></tr><tr><td valign=""top"" colspan=""3"" height=""126"">")
	Response.Write ("<table width=""100%"" border=""1"" cellspacing=""0"" cellpadding=""0"" bordercolor=""#3366CC""><tr><td valign=""top"" bgcolor=""#fff9ea"">")
	Response.Write ("<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""5"" height=""105""><tr><td class=""TextDecoration"" bgcolor=""#EEEEFF"">")
	
	while (not f.AtEndOfStream)
		strFirst = f.ReadLine
		strSecond = f.ReadLine
		Response.Write ("<div align=""center""><span style='color:red'><b>")
		Response.Write (strFirst)
		Response.Write ("</b></span><br>")
		Response.Write (strSecond)
		Response.Write ("</div><br>")
	wend
	Response.Write ("</td></tr></table></td></tr></table>")
	Response.Write ("</td></tr><tr><td valign=""top"" height=""25"">&nbsp;</td>")
	Response.Write ("<td valign=""top"" height=""25"">&nbsp;</td><td valign=""top"" height=""25"">&nbsp;</td></tr></table>")

end function

function FillCommandment ()
	Set fso = Server.CreateObject ("Scripting.FileSystemObject")
	Set f = fso.OpenTextFile (Server.MapPath ("res/commandment.dat"))
	Dim strPromise, strReference
	
	strNum = f.ReadLine
	
	for i = 0 to (Day (Date) mod CInt (strNum))
		strPromise = f.ReadLine
		strReference = f.ReadLine
	next
	Response.Write ("<tr><td class=""TextDecoration"" height=""1"" bgcolor=""#EEEEFF""><p align=""center"">")
	Response.Write (strPromise)
	Response.Write ("</p></td></tr><tr><td height=""2"" bgcolor=""#EEEEFF""><div align=""center""><font face=""Verdana, Arial, Helvetica, sans-serif"" size=""2""><b><font size=""1"" color=""#FF0000"">")
	Response.Write (strReference)
	Response.Write ("</font></b></font></div></td></tr>")
	f.Close
end function
		</script>
	</head>
	<body bgcolor="#ffffff" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
		<font face="Verdana, Arial, Helvetica, sans-serif" size="2"></font>
		<!-- #include file = "includes/topborder.htm" -->
		<table height="90%" width="100%" border="0" cellspacing="0" cellpadding="0">
			<tr>
				<td width="180" valign="top" bgcolor="#EEEEFF"><br>
					<!-- #include file = "includes/leftmenu.htm" -->
					<p>&nbsp;</p>
				</td>
				<td width="2" valign="top"><img src="img/1p.gif" width="1" height="100%"></td>
				<td valign="top">
					<table width="100%" border="0" cellspacing="0" cellpadding="5">
						<tr>
							<td valign="top" height="479">
								<table width="100%" border="0" cellspacing="0" cellpadding="5" height="442">
									<tr>
										<td valign="top">
											<table width="100%" border="0" cellspacing="0" cellpadding="5" height="162">
												<tr background="img/greengradient.gif">
													<td colspan="2" background="img/greengradient.gif">
														<p><font face="Verdana, Arial, Helvetica, sans-serif" size="3"><b><font size="4" color="#FFFFFF">Word 
																		for the Week</font></b></font></p>
													</td>
												</tr>
												<tr>
													
                      <td width="180"><img src="img/bible.gif" width="158"></td>
													<td class="TextDecoration" width="83%" height="136">
													<%FillWordForWeek ()%>
													</td>
												</tr>
											</table>
											<br>
											<table width="100%" border="0" cellspacing="0" cellpadding="5" height="139">
												<tr>
													<td colspan="2" background="img/browngradient.gif">
														<p><font face="Verdana, Arial, Helvetica, sans-serif" size="3"><b><font size="4" color="#FFFFFF">  
																		This Week on Radio</font></b></font></p>
													</td>
												</tr>
												<tr>
													
                      <td rowspan="2" valign="top" height="38"><img src="img/radio.gif" width="158"></td>
													<td class="TextDecoration" width="83%" height="128">
														<% FillRadio () %>
													</td>
												</tr>
											</table>
											<br>
											<table width="100%" border="0" cellspacing="0" cellpadding="5" height="162" ID="Table4">
												<tr background="img/orangegradient.gif">
													<td colspan="2" background="img/orangegradient.gif">
														<p><font face="Verdana, Arial, Helvetica, sans-serif" size="3"><b><font size="4" color="#FFFFFF">Audio Sermon Highlight</font></b></font></p>
													</td>
												</tr>
												<tr>
													
													<td><img src="img/sermon.gif" width="158"></td>
													<td class="TextDecoration" width="83%" height="136">
													<%FillAudioSermon ()%>
													</td>
												</tr>
											</table>
										</td>
									</tr>
								</table>
							</td>
						</tr>
					</table>
				</td>
				<td width="200" valign="top" align="right">
					<table width="100%" border="0" cellspacing="0" cellpadding="5">
						<tr>
							<td class="TextDecoration"><br>
								<table width="100%" border="0" cellspacing="0" cellpadding="0" height="131">
									<tr>
										<td valign="center" width="16"><img src="img/left.gif" width="16" height="100%">
										</td>
										<td valign="center" width="90%" bgcolor="#3366CC">
											<div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font color="#FFFFFF">Today's 
															Promise </font></b></font>
											</div>
										</td>
										<td valign="center" width="16"><img src="img/right.gif" width="16" height="100%"></td>
									</tr>
									<tr>
										<td valign="top" colspan="3" height="92">
											<table width="100%" border="1" cellspacing="0" cellpadding="0" height="97" bordercolor="#3366CC">
												<tr>
													<td valign="top">
														<table width="100%" border="0" cellspacing="0" cellpadding="5" height="105">
															<%FillTodaysPromise ()%>
														</table>
													</td>
												</tr>
											</table>
										</td>
									</tr>
									<tr>
										<td valign="top" height="3">&nbsp;</td>
										<td valign="top" height="3">&nbsp;</td>
										<td valign="top" height="3">&nbsp;</td>
									</tr>
								</table>
								<br>
								<table width="100%" border="0" cellspacing="0" cellpadding="0" height="131" ID="Table1">
									<tr>
										<td width="16"><img src="img/left.gif" width="16" height="100%">
										</td>
										<td valign="center" width="90%" bgcolor="#3366CC">
											<div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font color="#FFFFFF">Today's 
															Commandment </font></b></font>
											</div>
										</td>
										<td valign="center" width="16"><img src="img/right.gif" width="16" height="100%"></td>
									</tr>
									<tr>
										<td valign="top" colspan="3">
											<table width="100%" border="1" cellspacing="0" cellpadding="0" height="97" bordercolor="#3366CC" ID="Table2">
												<tr>
													<td valign="top">
														<table width="100%" border="0" cellspacing="0" cellpadding="5" height="105" ID="Table3">
															<%FillCommandment ()%>
														</table>
													</td>
												</tr>
											</table>
										</td>
									</tr>
									<tr>
										<td valign="top" height="3">&nbsp;</td>
										<td valign="top" height="3">&nbsp;</td>
										<td valign="top" height="3">&nbsp;</td>
									</tr>
								</table>
								<br>
								<%FillEvents ()%>
								
							</td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
		<!-- #include file="includes/bottomborder.htm"-->
	</body>
</html>
