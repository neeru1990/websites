<html xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xsl:version="1.0">
	<head>
		<title>Send us a note</title>
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
' Dim strFreeGift

function InitValues ()
	strType = "<option selected>Select One</option><option>Comments</option><option>Feedback</option><option>Question</option><option>Praise Report</option><option>Other</option>"

<!-- **** REMOVED AUDIO SERMONS
	if (Request.QueryString = "audio") then
		strType = "<option>Select One</option><option>Comments</option><option>Feedback</option><option>Question</option><option>Praise Report</option><option selected>Request Audio Sermon</option><option>Other</option>"
	else
		strType = "<option selected>Select One</option><option>Comments</option><option>Feedback</option><option>Question</option><option>Praise Report</option><option>Request Audio Sermon</option><option>Other</option>"
	end if
-->
	strTitle = "<option selected>Select One</option><option>Mr.</option><option>Mrs.</option><option>Ms.</option>"
	strName = ""
	strPhone = ""
	strEmail = ""
	strAddress1 = ""
	strAddress2 = ""
	strCity = ""
	strZip = ""
	strCountry = ""
	strComments = ""
	strWordForTheWeek = "<INPUT id=""Checkbox1"" type=""checkbox"" value=""ON"" name=""WordForTheWeek"" CHECKED>"
	strMinistryUpdates = "<INPUT id=""Checkbox2"" type=""checkbox"" value=""ON"" name=""MinistryUpdates"">"
	' strFreeGift = "<INPUT id=""Checkbox3"" type=""checkbox"" value=""OFF"" name=""FreeGift"">"
	
	strColPurpose = "whitesmoke"
	strColName = "whitesmoke"
	strColTitle = "whitesmoke"
	strColEmail = "whitesmoke"
end function

function CheckValues ()
	if (Request.QueryString <> "submit") then exit function
	
	strType = "<option>Select One</option><option>Comments</option><option>Feedback</option><option>Question</option><option>Praise Report</option><option>Other</option>"
	nIndex = InStr (1, strType, Request.Form.Item ("Type"))
	if (nIndex <> 0) then
		strType = left (strType, nIndex - 2) & " selected" & mid (strType, nIndex - 1)
	end if
	
	strTitle = "<option>Select One</option><option>Mr.</option><option>Mrs.</option><option>Ms.</option>"
	nIndex = InStr (1, strTitle, Request.Form.Item ("Contact_Title"))
	if (nIndex <> 0) then
		strTitle = left (strTitle, nIndex - 2) & " selected" & mid (strTitle, nIndex - 1)
	end if
	
	strName = Request.Form.Item ("Contact_FullName")
	strPhone = Request.Form.Item ("Contact_Phone")
	strEmail = Request.Form.Item ("Contact_Email")
	strAddress1 = Request.Form.Item ("Contact_Address1")
	strAddress2 = Request.Form.Item ("Contact_Address2")
	strCity = Request.Form.Item ("Contact_City")
	strZip = Request.Form.Item ("Contact_Zip")
	strCountry = Request.Form.Item ("Contact_Country")
	strComments = Request.Form.Item ("Comments")
	
	if (Request.Form.Item ("WordForTheWeek") = "ON") then 
		strWordForTheWeek = "<INPUT id=""Checkbox1"" type=""checkbox"" value=""ON"" name=""WordForTheWeek"" CHECKED>" 
	else
		strWordForTheWeek = "<INPUT id=""Checkbox1"" type=""checkbox"" value=""ON"" name=""WordForTheWeek"">" 
	end if
	
	if (Request.Form.Item ("MinistryUpdates") = "ON") then 
		strMinistryUpdates = "<INPUT id=""Checkbox2"" type=""checkbox"" value=""ON"" name=""MinistryUpdates"" CHECKED>"
	else
		strMinistryUpdates = "<INPUT id=""Checkbox2"" type=""checkbox"" value=""ON"" name=""MinistryUpdates"">"
	end if

<!--	
	if (Request.Form.Item ("FreeGift") = "ON") then 
		strMinistryUpdates = "<INPUT id=""Checkbox3"" type=""checkbox"" value=""ON"" name=""FreeGift"" CHECKED>"
	else
		strMinistryUpdates = "<INPUT id=""Checkbox3"" type=""checkbox"" value=""ON"" name=""FreeGift"">"
	end if
-->
	
	strError = "  "
	if (Request.Form.Item ("Type") = "Select One") then 
		strError = strError + "Purpose of this note, "
		strColPurpose = "#CC6633"
	end if
	if (Request.Form.Item ("Contact_Title") = "Select One") then 
		strError = strError + "Title, "
		strColTitle = "#CC6633"
	end if
	if (strName = "") then 
		strError = strError + "Name, "
		strColName = "#CC6633"
	end if
	if ((Request.Form.Item ("WordForTheWeek") = "ON" or Request.Form.Item ("MinistryUpdates") = "ON") and strEmail = "") then 
		strError = strError + "E-Mail, "
		strColEmail = "#CC6633"
	end if
	strError = left (strError, Len (strError) - 2)
	
	if (strError <> "") then
		strError = "<font color=""red""><b>Please enter relevant information in the following fields: " + strError
		strError = strError + "</b></font>."
		Response.Write (strError)
	else
		strFrom = "webmaster@powerofgospel.org"
		strTo = "info@powerofgospel.org"
		strSubject = Request.Form.Item ("Type")
		strBody = "Title: " & Request.Form.Item ("Contact_Title") & chr (13) & chr (10)
		strBody = strBody & "Name: " & Request.Form.Item ("Contact_FullName") & chr (13) & chr (10)
		strBody = strBody & "Phone: " & Request.Form.Item ("Contact_Phone") & chr (13) & chr (10)
		strBody = strBody & "EMail: " & Request.Form.Item ("Contact_Email") & chr (13) & chr (10)
		strBody = strBody & "Address1: " & Request.Form.Item ("Contact_Address1") & chr (13) & chr (10)
		strBody = strBody & "Address2: " & Request.Form.Item ("Contact_Address2") & chr (13) & chr (10)
		strBody = strBody & "City and State: " & Request.Form.Item ("Contact_City") & chr (13) & chr (10)
		strBody = strBody & "Zip: " & Request.Form.Item ("Contact_Zip") & chr (13) & chr (10)
		strBody = strBody & "Country: " & Request.Form.Item ("Contact_Country") & chr (13) & chr (10)
		strBody = strBody & "Comments: " & Request.Form.Item ("Comments") & chr (13) & chr (10)
		strBody = strBody & "WordForTheWeek: " & Request.Form.Item ("WordForTheWeek") & chr (13) & chr (10)
		' strBody = strBody & "MinistryUpdates: " & Request.Form.Item ("MinistryUpdates") & chr (13) & chr (10)
		
<!--
	   strBody = strBody & "FreeGift: " & Request.Form.Item ("FreeGift") & chr (13) & chr (10)
	
		
	    if (Request.Form.Item ("FreeGift") = "ON") then
	        strQuery = "select freegiftcount from freegift"
	
	        Set dbConn = Server.CreateObject ("ADODB.Connection")
	        dbConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath ("res/broadcast.mdb")
	        dbConn.Open
        	
	        Set dbRst = Server.CreateObject ("ADODB.RecordSet")
	        dbRst.Open strQuery, dbConn
        	
	        dbRst.MoveFirst
	        nCount = dbRst.Fields ("freegiftcount").Value
	        nCount = nCount - 1
    	    
	        strQuery = "update freegift set freegiftcount=" & nCount
	        dbConn.Execute strQuery
        	
	        dbRst.Close
	        dbConn.Close 
	    end if
-->
	    
		Set objMsg = CreateObject("CDO.Message") 
		objMsg.Subject = strSubject 
		objMsg.Sender = strFrom 
		objMsg.To = strTo 
		objMsg.TextBody = strBody 
		objMsg.Send

		Response.Redirect ("submitguestbook.asp")
	end if
end function

<!--
function FillFreeGift ()
    strQuery = "select freegiftcount from freegift"
	
	Set dbConn = Server.CreateObject ("ADODB.Connection")
	dbConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath ("res/broadcast.mdb")
	dbConn.Open
	
	Set dbRst = Server.CreateObject ("ADODB.RecordSet")
	dbRst.Open strQuery, dbConn
	
	dbRst.MoveFirst
	nCount = dbRst.Fields ("freegiftcount").Value
	
	if (nCount <= 0) then
	    Response.Write ("")
	else
	    Response.Write (strFreeGift)
	    Response.Write (" Send me a free gift")
	end if
	dbRst.Close
	dbConn.Close 
end function
-->
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
						                				        <p><font face="Verdana, Arial, Helvetica, sans-serif" size="3"><b><font size="4" 										color="#FFFFFF">Send us a note</font></b></font></p>
										                      </td>
										                    </tr>
												<tr>
													<td>
														<FORM METHOD="post" action="guestbook.asp?submit" name="Contact_Form">
															<!--webbot bot="SaveResults" U-File="_private/formrslt.htm" S-Format="HTML/DL" S-Label-Fields="TRUE" B-Reverse-Chronology="FALSE" S-Email-Format="TEXT/PRE" S-Email-Address="info@powerofgospel.org" B-Email-Label-Fields="TRUE" S-Builtin-Fields startspan --><input TYPE="hidden" NAME="VTI-GROUP" VALUE="0"><!--webbot bot="SaveResults" i-checksum="43374" endspan -->
															<TABLE id="Table1" cellSpacing="1" cellPadding="5" width="100%" border="0">
																<tr>
																	<td>
																		<font face="Verdana, Arial, Helvetica, sans-serif" size="2">We invite your comments 
																			and feedback. If you have a comment or a question, please send that over to us 
																			from this page. You can also request the Word for the Week to be sent weekly to 
																			your email address. If the Lord has put a burden in your heart to pray for this 
																			ministry, you can request ministry updates as well. The information you are 
																			providing to us on this page is not shared with anyone outside of Power of 
																			Gospel ministries. We look forward to hearing from you. </font>
																	</td>
																</tr>
																<tr>
																	<td>
																		<font face="Verdana, Arial, Helvetica, sans-serif" size="4" color="red"><b>Please do not send Bible requests here.  
																		  Instead, request Bibles <a href="freegift.asp">here</a></b></font>
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
																	<TR bgcolor=<%Response.Write ("""" & strColPurpose & """")%>>
																		<TD class="TextDecoration2" width="10"><font color="red">*</font></TD>
																		<TD class="TextDecoration2" width="174"><b>Purpose of this note:</b></TD>
																		<TD class="TextDecoration2"><SELECT id="Select3" name="Type">
																				<%Response.Write ("""" & strType & """")%>
																			</SELECT>
																		</TD>
																	</TR>
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
																	<TR bgcolor="whitesmoke">
																		<TD class="TextDecoration2" width="10"></TD>
																		<TD class="TextDecoration2" width="174">Phone</TD>
																		<TD class="TextDecoration2">
																			<INPUT TYPE="text" NAME="Contact_Phone" SIZE="25" MAXLENGTH="25" ID="Text3" value=<%Response.Write ("""" & strPhone & """")%>></TD>
																	</TR>
																	<TR bgcolor=<%Response.Write ("""" & strColEmail & """")%>>
																		<TD class="TextDecoration2" width="10"><font color="red">*</font></TD>
																		<TD class="TextDecoration2" width="174"><b>E-mail</b></TD>
																		<TD class="TextDecoration2">
																			<INPUT TYPE="text" NAME="Contact_Email" SIZE="25" ID="Text5" value=<%Response.Write ("""" & strEmail & """")%>></TD>
																	</TR>
																	<TR bgcolor="beige">
																		<TD class="TextDecoration2" width="10"></TD>
																		<TD class="TextDecoration2" width="174">Address&nbsp; 1</TD>
																		<TD class="TextDecoration2">
																			<INPUT TYPE="text" NAME="Contact_Address1" SIZE="25" MAXLENGTH="25" ID="Text6" value=<%Response.Write ("""" & strAddress1 & """")%>></TD>
																	</TR>
																	<TR bgcolor="beige">
																		<TD class="TextDecoration2" width="10"></TD>
																		<TD class="TextDecoration2" width="174">Address 2</TD>
																		<TD class="TextDecoration2">
																			<INPUT id="Text7" type="text" maxLength="25" size="25" name="Contact_Address2" value=<%Response.Write ("""" & strAddress2 & """")%>></TD>
																	</TR>
																	<TR bgcolor="beige">
																		<TD class="TextDecoration2" width="10"></TD>
																		<TD class="TextDecoration2" width="174">City, State</TD>
																		<TD class="TextDecoration2">
																			<INPUT id="Text8" type="text" maxLength="25" size="25" name="Contact_City" value=<%Response.Write ("""" & strCity & """")%>></TD>
																	</TR>
																	<TR bgcolor="beige">
																		<TD class="TextDecoration2" width="10"></TD>
																		<TD class="TextDecoration2" width="174">Zip</TD>
																		<TD class="TextDecoration2">
																			<INPUT id="Text9" type="text" maxLength="25" size="25" name="Contact_Zip" value=<%Response.Write ("""" & strZip & """")%>></TD>
																	</TR>
																	<TR bgcolor="beige">
																		<TD class="TextDecoration2" width="10"></TD>
																		<TD class="TextDecoration2" width="174">Country</TD>
																		<TD class="TextDecoration2">
																			<INPUT id="Text10" type="text" maxLength="25" size="25" name="Contact_Country" value=<%Response.Write ("""" & strCountry & """")%>></TD>
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
																			 <%Response.Write (strWordForTheWeek)%>&nbsp;Send 
																			me Word For The Week emails</TD>
																	</TR>
																	<TR bgcolor="honeydew">
																		<TD class="TextDecoration2" width="10"></TD>
																		<td></td>
																		<TD class="TextDecoration2">
																			 <% ' FillFreeGift ()%> </TD>
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
