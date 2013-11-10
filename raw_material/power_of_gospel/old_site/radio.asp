<html xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xsl:version="1.0">
	<head>
		<title>Radio Broadcasts</title> 
		<!-- #include file="includes/styles.htm" -->
		<SCRIPT>
window.onload=init;
function init()
{
	for(i in document.links) document.links[i].onfocus = document.links[i].blur;
}

		</SCRIPT>
		<script language="vbscript" runat="server">

function CheckDate(start)
	if (start="") then
	 Response.Redirect("radio.asp?start=" & Date)
  end if
end function
		
function PrintFillRadio(start)
 strFillRadio = FillRadio(start)
 Response.Write (strFillRadio)
end function

function FillRadio (start)
	Set fso = Server.CreateObject ("Scripting.FileSystemObject")
	Set f = fso.GetFolder(Server.MapPath ("radio/"))
	Set fc = f.Files
	if (request("start")="") then
	 start = Date
	end if
	EndDate = Dateserial (Year(start), Month(start)-2, Day(start))
	strResponse = ""
	strStart = FormatDateSpecial (start)
	strEnd = FormatDateSpecial (EndDate)
	if (fc.count <> 0) then
		for each f1 in fc
			if (Left (f1.name, 2) = "20") and (((f1.name < strStart) and (f1.name > strEnd)) or (f1.name = strStart)) then
				Set fo = fso.OpenTextFile (f1.Path)
				strWord = fo.ReadLine
				strDescription = fo.ReadLine
				strLink = fo.ReadLine
				fo.close
				strTemp = "<tr><td class=""TextDecoration"">"
				strFile = f1.name
				strNewFile = Mid (strFile, 5, 2)
				strNewFile = strNewFile & "/" & Mid (strFile, 7, 2)
				strNewFile = strNewFile & "/" & Left (strFile, 4)
				strTemp = strTemp + strNewFile
				strTemp = strTemp & "</td><td class=""TextDecoration"" bgcolor=#FFEECC><font color=#993300><b>"
				if (Len (strLink) > 2) then strTemp = strTemp & "<a href=""radio/broadcasts/" & strLink & """>"
				strTemp = strTemp + strWord
				if (Len (strLink) > 2) then strTemp = strTemp & "</a>"
				strTemp = strTemp & "</b></font></td><td class=""TextDecoration"">"
				strTemp = strTemp + strDescription
				strTemp = strTemp & "</td></tr>"
				
				strResponse = strTemp + strResponse
			end if
		next
	end if

  FillRadio = strResponse
end function

function Links(start)
 if (request("start")="") then
  start = Date
 end if
 PrevDate = Dateserial (Year(start), Month(start)+2, Day(start))
 NextDate = Dateserial (Year(start), Month(start)-2, Day(start))
 strStart = FormatDateSpecial(start)
 strDate = FormatDateSpecial(Date)
 
 if(strstart=strDate) then
  strPrev = "<font color=""#996633"">&nbsp;</font>"
 else
  strPrev = "<a href=""radio.asp?start=" & PrevDate & """><b>< Newer</b></a>"
 end if
 
 strNewLine = "<tr bgcolor=""#996633""><td height.""12"" colspan=""2""><p><font face=""Verdana, Arial, Helvetica, sans-serif"" size=""3""><b><font size=""4"" color=""#FFFFFF"">"
 strLink = strNewLine + strPrev & "</font></b></font></p></td><td height=""12"" colspan=""1""><p align=""right"">"
 strLink = strLink & "<font face=""Verdana, Arial, Helvetica, sans-serif"" size=""3""><b><font size=""4"" color=""#FFFFFF"">"
 
 strCheck = FillRadio(NextDate)
 
 if (strCheck="") then
	strNext = "<font color=""#996633"">&nbsp;</font>"
 else
  strNext = "<a href=""radio.asp?start=" & NextDate & """><b>Older ></b></a>"
 end if
 
 strLink = strLink + strNext
 strLink = strLink & "</font></b></font></p></td></tr>"
 
 Response.Write (strLink)
end function

function FormatDateSpecial (formatDate)
	strFormat = Year (formatDate)
	
	if (Month (formatDate) < 10) then strFormat = strFormat & "0"
	strFormat = strFormat & Month (formatDate)
	
	if (Day (formatDate) < 10) then strFormat = strFormat & "0"
	strFormat = strFormat & Day (formatDate)
	
	FormatDateSpecial = strFormat & ".dat"
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
														<p><font face="Verdana, Arial, Helvetica, sans-serif" size="3"><b><font size="4" color="#FFFFFF">Radio 
																		Broadcasts</font></b></font></p>
													</td>
												</tr>
												<%CheckDate (request("start"))%>
												<%PrintFillRadio (request("start"))%>
												</tr>
												<%Links (request("start"))%>
											</table>
											<br>
											<div class="TextDecoration">You would need Real player to listen to the audio files.  Real player is available for free at <a href="http://www.real.com">http://www.real.com</a>.</div>
											<div class="TextDecoration">For broadcast coverage reference go to: <a href="showbroadcast.asp">Broadcast Coverage</a>.</div>
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
