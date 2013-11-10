<html xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xsl:version="1.0">
	<head>
		<title>Store Closed</title>
		<!-- #include file="includes/styles.htm" -->
		<SCRIPT>
window.onload=init;
function init()
{
	for(i in document.links) document.links[i].onfocus = document.links[i].blur;
}

		</SCRIPT>
		<script language="vbscript" runat="server">
function SubmitForm
	Response.Write ("We are closed briefly for maintenance.")
	Response.Write ("<br>Please visit this page again after some time.")
	Response.Write ("<br>We apologize the inconvenience.")
end function
		</script>
		<script language="vbscript">
function CheckValues (form)
	MsgBox "I'm checking values here"
	CheckValues = true
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
														<font face="Verdana, Arial, Helvetica" size=4><b><%SubmitForm ()%></b></font>
<font face="Verdana, Arial, Helvetica" size=2><a href="default.asp"><br>Click here to continue to browse the website.</a></font>
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
