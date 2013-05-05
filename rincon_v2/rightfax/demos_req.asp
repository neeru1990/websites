<HTML>
<HEAD>
<TITLE>Rincon > Products > RightFax</TITLE>
<LINK href="../includes/rincon.css" type=text/css rel=stylesheet>
</HEAD>
<BODY>
<SCRIPT language=JavaScript src="../includes/menu_array.js" type=text/javascript></SCRIPT>
<SCRIPT language=JavaScript src="../includes/mmenu.js" type=text/javascript></SCRIPT>
<%	
	nm = request("Name")
	title = request("Title")
	company = request("CompanyName")
	addr1 = request("Street")
	addr2 = request("Street2")
	city = request("City")
	zip = request("ZipCode")
	ph = request("PhoneNumber")
	fax = request("FaxNumber")
	email = request("Email")
	cust_type = request("CustType")
	exchange = request("ChkE")
	lotus = request("ChkL")
	oracle = request("ChkO")

		sub_from = "Demo Request"
		
		mail_contents="Name:" &nm  &vbCrLf &_
			 	"Title     :" &title  &vbCrLf &_
				"Company   :" &company   &vbCrLf &_
				"Address 1 :" &addr1    &vbCrLf &_
				"Address 2 :" &addr2    &vbCrLf &_
				"City      :" &city  &vbCrLf &_
				"Zip/Postal:" &zip   &vbCrLf &_
				"Busi Phone:" &ph   &vbCrLf &_
				"Fax       :" &fax  &vbCrLf &_
				"Email     :" &email  &vbCrLf &_
				"You are a/an...? :" &cust_type
				  
		Dim objMessage
		Set objMessage = Server.CreateObject("CDONTS.Newmail")
		objMessage.To       = "sales@rincon.co.in"
		objMessage.From     = email
		objMessage.Subject  = sub_from
		objMessage.Body	    = mail_contents
		On Error Resume Next 
		objMessage.Send
		If Err <> 0 Then Response.Write "Error encountered: " & Err.Description 
		Set objMessage = Nothing
%>
<TABLE width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <TR> 
    <TD class="pad15"><A href="../index.htm"><IMG src="../images/rincon_logo.gif" alt="Rincon India Solutions Pvt. Ltd." width="192" height="49" border="0"></A></TD>
    <TD colspan="4" class="links"><IMG src="../images/bullet_low.gif" width="25" height="9"><A class="menu" href="../index.htm">Home</A><IMG src="../images/bullet_low.gif" width="25" height="9"><A href="../about.htm" class="menu">About 
      Us</A><IMG src="../images/bullet_low.gif" width="25" height="9"><A class="menu" href="../products.htm">Products</A><IMG src="../images/bullet_low.gif" width="25" height="9"><A class="menu" href="../contact.htm">Contact 
      Us</A></TD>
  </TR>
  <TR> 
    <TD width="222" height="20" bgcolor="#2F5B79">&nbsp; </TD>
    <TD width="116" bgcolor="#5FA7C9">&nbsp;</TD>
    <TD width="239" bgcolor="#A9CFE2">&nbsp;</TD>
    <TD width="31" bgcolor="#013554">&nbsp;</TD>
    <TD width="162" bgcolor="#0270B0">&nbsp;</TD>
  </TR>
</TABLE>
<TABLE width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <TR> 
    <TD class="bg">&nbsp;</TD>
    <TD class="borders"> <H1>RightFax</H1>
      <H2>Online Demo's</H2>
      <% if exchange = "ChkE" then %>
      <P><SCRIPT language="JavaScript"><!-- 
		function exchange(){open('rf_demos/exchange.htm','demo','width=670,height=500,menubar=no,location=no,toolbar=no,scrollbars=no')}
		//--></SCRIPT>
        <STRONG>Gateway for Microsoft Exchange - <A href="javascript:exchange()">View 
        Demo</A></STRONG><BR>
        The RightFax Gateway for Microsoft Exchange lets users send, receive and 
        manage both faxes and email using Microsoft Outlook. The gateway enables 
        Exchange administrators to manage fax functionality in Exchange with complete 
        integration and automatic database synchronization with RightFax.</P>
      <%	 
		end if
		if lotus = "ChkL" then %>
      <P><SCRIPT language="JavaScript"><!--
		function lotus(){open('rf_demos/lotus.htm','demo','width=670,height=500,menubar=no,location=no,toolbar=no,scrollbars=no')}
		//--></SCRIPT>
        <STRONG>Gateway for Lotus Notes - <A href="javascript:lotus()">View Demo</A></STRONG><BR>
        The RightFax Gateway for Lotus Notes lets users send, receive and manage 
        both faxes and emails through their Lotus Notes email databases. The gateway 
        enables Notes administrators to manage fax functionality from their Notes 
        domain with complete integration and automatic database synchronization 
        with RightFax.</P>
      <%
		end if
      if oracle = "ChkO" then %>
      <P><SCRIPT language="JavaScript"><!--
		function oracle(){open('rf_demos/oracle.htm','demo','width=670,height=500,menubar=no,location=no,toolbar=no,scrollbars=no')}
		//--></SCRIPT>
        <STRONG>Integration Options for Oracle - <A href="javascript:oracle()">View 
        Demo</A></STRONG><BR>
        The RightFax electronic document delivery server empowers Oracle users 
        to send mission-critical documents such as invoices, purchase orders, 
        confirmations and releases, faster, more reliably, more securely and less 
        expensively than sending them by traditional fax or mail.</P>
      <%end if%>
       <P>You will need the Macromedia Flash plug-in to view the RightFax demo's. 
        If you do not have the Flash plug-in, visit the <A href="http://www.macromedia.com/shockwave/download/frameset.fhtml?P1_Prod_Version=ShockwaveFlash"
    target="_blank">Macromedia Web site</A> to install it and then return to continue. 
      </P>

      </TD>
  </TR>
</TABLE>
<BR>
</BODY>
</HTML>
