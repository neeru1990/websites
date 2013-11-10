<html xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xsl:version="1.0">
	<head>
		<title>Prayer Request</title>
		<!-- #include file="includes/styles.htm" -->
		<SCRIPT>
window.onload=init;
function init()
{
	for(i in document.links) document.links[i].onfocus = document.links[i].blur;
}

		</SCRIPT>
		<script runat=server language="vbscript">
Dim strType, strTitle, strName, strPhone, strEmail, strAddress1, strAddress2, strCity, strZip
Dim strCountry, strComments, strWordForTheWeek, strMinistryUpdates, strColPurpose
Dim strColName, strColTitle, strColEmail

function InitValues ()
	strTitle = "<option selected>Select One</option><option>Mr.</option><option>Mrs.</option><option>Ms.</option>"
	strName = ""
	strEmail = ""
	strComments = ""
	
	strColName = "whitesmoke"
	strColTitle = "whitesmoke"
	strColEmail = "whitesmoke"
end function

function CheckValues ()
	if (Request.QueryString <> "submit") then exit function
	
	strTitle = "<option>Select One</option><option>Mr.</option><option>Mrs.</option><option>Ms.</option>"
	nIndex = InStr (1, strTitle, Request.Form.Item ("Contact_Title"))
	if (nIndex <> 0) then
		strTitle = left (strTitle, nIndex - 2) & " selected" & mid (strTitle, nIndex - 1)
	end if
	
	strName = Request.Form.Item ("Contact_FullName")
	strEmail = Request.Form.Item ("Contact_Email")
	strComments = Request.Form.Item ("Comments")
	
	strError = "  "
	if (Request.Form.Item ("Contact_Title") = "Select One") then 
		strError = strError + "Title, "
		strColTitle = "#CC6633"
	end if
	if (strName = "") then 
		strError = strError + "Name, "
		strColName = "#CC6633"
	end if
	strError = left (strError, Len (strError) - 2)
	
	if (strError <> "") then
		strError = "<font color=""red""><b>Please enter relevant information in the following fields: " + strError
		strError = strError + "</b></font>."
		Response.Write (strError)
	else
		strFrom = "webmaster@powerofgospel.org"
		strTo = "prayer@powerofgospel.org"
		strSubject = "Prayer Request"
		strBody = "Title: " & Request.Form.Item ("Contact_Title") & chr (13) & chr (10)
		strBody = strBody & "Name: " & Request.Form.Item ("Contact_FullName") & chr (13) & chr (10)
		strBody = strBody & "EMail: " & Request.Form.Item ("Contact_Email") & chr (13) & chr (10)
		strBody = strBody & "Comments: " & Request.Form.Item ("Comments") & chr (13) & chr (10)
		
		Set objMsg = CreateObject("CDO.Message") 
		objMsg.Subject = strSubject 
		objMsg.Sender = strFrom 
		objMsg.To = strTo 
		objMsg.TextBody = strBody 
		objMsg.Send

		Response.Redirect ("submitguestbook.asp")
	end if
end function
		</script>
	</head>
	<body bgcolor="#ffffff" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
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
									<tr>
										<td valign="center">
											<table width="100%" height="100%" cellspacing="0" cellpadding="5" bordercolor="#339933">
												<tr background="img/browngradient.gif"> 
										                      <td background="img/browngradient.gif" height="25"> 
										                        <p><font face="Verdana, Arial, Helvetica, sans-serif" size="3"><b><font size="4" 										color="#FFFFFF">Send us your Prayer Request</font></b></font></p>
						                				      </td>
				    					                        </tr>
												<tr>
													<td>
														<FORM METHOD="post" action="prayerrequest.asp?submit" name="Contact_Form">
															<TABLE id="Table1" cellSpacing="1" cellPadding="5" width="100%" border="0">
																<tr>
																	<td class="TextDecoration">
																		<font color="red"><b>Again I say to you that if two of you agree on earth concerning any thing that they ask, it will be done for them by My Father in heaven. - Mathew 18:19</b></font>
																		<br><br>There is power in prayer of togetherness. 
																		We can join with you and uplift you on your prayer need. Submit your prayer request to us from this page. If you would like to hear back from us, please provide us a phone number or an email address that we can contact you with. The information you are providing to us on this 
																		page is not shared with anyone outside of Power of Gospel ministries.</div>
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
																		<TD class="TextDecoration2" width="174"><b> Name</b></TD>
																		<TD class="TextDecoration2">
																			<INPUT TYPE="text" NAME="Contact_FullName" SIZE="35" ID="Text1" value=<%Response.Write ("""" & strName & """")%>></TD>
																	</TR>
																	<TR bgcolor=<%Response.Write ("""" & strColEmail & """")%>>
																		<TD class="TextDecoration2" width="10"><font color="red">*</font></TD>
																		<TD class="TextDecoration2" width="174"><b>E-mail</b></TD>
																		<TD class="TextDecoration2">
																			<INPUT TYPE="text" NAME="Contact_Email" SIZE="25" ID="Text5" value=<%Response.Write ("""" & strEmail & """")%>></TD>
																	</TR>
																	<TR bgcolor="honeydew">
																		<TD class="TextDecoration2" width="10"></TD>
																		<TD class="TextDecoration2" width="174">Comments:</TD>
																		<TD class="TextDecoration2">
																			<TEXTAREA NAME="Comments" ROWS="5" COLS="35" ID="Textarea1" value=<%Response.Write ("""" & strComments & """")%>></TEXTAREA></TD>
																	</TR>
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
