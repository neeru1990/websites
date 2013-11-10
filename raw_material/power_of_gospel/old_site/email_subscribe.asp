<html xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xsl:version="1.0">
	<head>
		<title>Word for the week</title>
		<script language="vbscript" runat="server">
		strEmail = Request.Form.Item ("email")
		strFrom = "webmaster@powerofgospel.org"
		strTo = "info@powerofgospel.org"
		strSubject = "New Email subscription to Word for the Week from Website"
		Set objMsg = CreateObject("CDO.Message") 
		objMsg.Subject = strSubject 
		objMsg.Sender = strFrom 
		objMsg.To = strTo 
		objMsg.TextBody = strEmail 
		objMsg.Send
		
		function email()
			Response.Write ("Thank you for subscribing to the Word for the Week emails.<br>")
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