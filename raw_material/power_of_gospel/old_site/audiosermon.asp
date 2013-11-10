<html xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xsl:version="1.0">
	<head>
		<title>Audio Sermons</title> 
		<!-- #include file="includes/styles.htm" -->
		<SCRIPT>
window.onload=init;
function init()
{
	for(i in document.links) document.links[i].onfocus = document.links[i].blur;
}

		</SCRIPT>
		<script language="vbscript" runat="server">
function FillAudio ()
	Set fso = Server.CreateObject ("Scripting.FileSystemObject")
	Set f = fso.OpenTextFile (Server.MapPath ("audiosermons/audioindex.dat"))
	
	while (not f.AtEndOfStream)
		strTitle = f.ReadLine
		strDescription = f.ReadLine
		strLink = f.ReadLine
	
		Response.Write ("<tr><td align=""center"" class=""TextDecoration"">")
		if (strLink = "On Request") then 
			Response.Write ("<a href=""guestbook.asp?audio"">Request<br>Audio</a>")
		else
			Response.Write ("<a href=""/audiosermons/" & strLink & """><img border=0 src=""img/cdrom.gif""/><br>Listen</a>")
		end if
		Response.Write ("</td><td class=""TextDecoration"" bgcolor=#FFEECC><font color=#993300><b>")
		Response.Write (strTitle)
		Response.Write ("</b></font></td><td class=""TextDecoration"">")
		Response.Write (strDescription)
		Response.Write ("</td></tr>")
	wend
	f.close
end function
		</script>
	</head>
	<body bgcolor="#ffffff" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
		<font face="Verdana, Arial, Helvetica, sans-serif" size="2"></font>
		<!-- #include file="includes/topborder.htm" -->
		<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
			<tr>
				<td width="180" height="8" valign="top" bgcolor="#EEEEFF"><br>
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
										<td valign="top">
											<table width="100%" border="1" cellspacing="0" cellpadding="5" height="47" bordercolor="#996633">
												<tr background="img/browngradient.gif">
													<td height="12" background="img/browngradient.gif" colspan="3">
														<p><font face="Verdana, Arial, Helvetica, sans-serif" size="3"><b><font size="4" color="#FFFFFF">Audio 
																		Sermons</font></b></font></p>
													</td>
												</tr>
												<%FillAudio ()%>
											</table>
											<br>
											<div class="TextDecoration">You would need Real player to listen to the audio files.  Real player is available for free at <a href="http://www.real.com">http://www.real.com</a>.</div>
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
