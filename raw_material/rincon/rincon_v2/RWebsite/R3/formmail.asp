<PRE><%@ LANGUAGE="VBScript" %>

<% '***************************************************************************
   '* ASP FormMail                                                            *
   '*                                                                         *
   '* Do not remove this notice.                                              *
   '*                                                                         *
   '* Copyright 1999, 2000 by Mike Hall.                                      *
   '* Please see http://www.brainjar.com for documentation and terms of use.  *
   '***************************************************************************

   '- Customization of these values is required, see documentation. -----------

   referers   = Array("www.brainjar.com", "brainjar.com")
   mailComp   = "ASPMail"
   smtpServer = "rincon.co.in"
   fromAddr   = "web@rincon.co.in"

   '- End required customization section. -------------------------------------

   Response.Buffer = true
   errorMsgs = Array()

   'Check for form data.

   if Request.ServerVariables("Content_Length") = 0 then
     call AddErrorMsg("No form data submitted.")
   end if

   'Check if referer is allowed.

   validReferer = false
   referer = GetHost(Request.ServerVariables("HTTP_REFERER"))
   for each host in referers
     if host = referer then
       validReferer = true
     end if
   next
   if not validReferer then
     if referer = "" then
       call AddErrorMsg("No referer.")
     else
       call AddErrorMsg("Invalid referer: '" & referer & "'.")
     end if
   end if

   'Check for the recipients field.

   if Request.Form("_recipients") = "" then
     call AddErrorMsg("Missing email recipient.")
   end if

   'Check all recipient email addresses.

   recipients = Split(Request.Form("_recipients"), ",")
   for each name in recipients
     name = Trim(name)
     if not IsValidEmail(name) then
       call AddErrorMsg("Invalid email address in recipient list: " & name & ".")
     end if
   next
   recipients = Join(recipients, ",")

   'Get replyTo email address from specified field, if given, and check it.

   name = Trim(Request.Form("_replyToField"))
   if name <> "" then
     replyTo = Request.Form(name)
   else
     replyTo = Request.Form("_replyTo")
   end if
   if replyTo <> "" then
     if not IsValidEmail(replyTo) then
       call AddErrorMsg("Invalid email address in reply-to field: " & replyTo & ".")
     end if
   end if

   'Get subject text.

   subject = Request.Form("_subject")

   'If required fields are specified, check for them.

   if Request.Form("_requiredFields") <> "" then
     required = Split(Request.Form("_requiredFields"), ",")
     for each name in required
       name = Trim(name)
       if Left(name, 1) <> "_" and Request.Form(name) = "" then
         call AddErrorMsg("Missing value for " & name)
       end if
     next
   end if

   'If a field order was given, use it. Otherwise use the order the fields were
   'received in.

   str = ""
   if Request.Form("_fieldOrder") <> "" then
     fieldOrder = Split(Request.Form("_fieldOrder"), ",")
     for each name in fieldOrder
       if str <> "" then
         str = str & ","
       end if
       str = str & Trim(name)
     next
     fieldOrder = Split(str, ",")
   else
     fieldOrder = FormFieldList()
   end if

   'If there were no errors, build the email note and send it.

   if UBound(errorMsgs) < 0 then

     'Build table of form fields and values.

     body = "<table border=0 cellpadding=2 cellspacing=0>" & vbCrLf
     for each name in fieldOrder
       body = body _
            & "<tr valign=top>" _
            & "<td><b>" _
            & name _
            & ":&nbsp;</b></font></td>" _
            & "<td>" _
            & Request.Form(name) _
            & "</td>" _
            & "</tr>" & vbCrLf
     next
     body = body & "</table>" & vbCrLf

     'Add a table with any environmental variables.

     if Request.Form("_envars") <> "" then
       body = body _
            & "<p>" _
            & "<table border=0 cellpadding=2 cellspacing=0>" & vbCrLf
       envars = Split(Request.Form("_envars"), ",")
       for each name in envars
         name = Trim(name)
         body = body _
              & "<tr valign=top>" _
              & "<td><b>" _
              & name _
              & ":&nbsp;</b></td>" _
              & "<td>" _
              & Request.ServerVariables(name) _
              & "</td>" _
              & "</tr>" & vbCrLf
       next
       body = body & "</table>" & vbCrLf
     end if

     'Send it.

     str = SendMail()
     if str <> "" then
       AddErrorMsg(str)
     end if

     'Redirect if a URL was given.

     if Request.Form("_redirect") <> "" then
       Response.Redirect(Request.Form("_redirect"))
     end if

   end if %>

<html>
<head>
<title>Form Mail</title>
<style style="text/css">

body {
  background-color: #ffffff;
  color: #000000;
  font-family: Arial, Helvetica, sans-serif;
  font-size: 10pt;
}

table {
  border: solid 1px #000000;
  border-collapse: collapse;
}

td, th {
  border: solid 1px #000000;
  border-collapse: collapse;
  font-family: Arial, Helvetica, sans-serif;
  font-size: 10pt;
  padding: 2px;
  padding-left: 8px;
  padding-right: 8px;
}

.error {
  color: #c00000;
}

</style>
</head>
<body>

<% if UBound(errorMsgs) >= 0 then %>
<p class="error">Form could not be processed due to the following errors:</p>
<ul>
<%   for each msg in errorMsgs %>
  <li class="error"><% = msg %>
<%   next %>
</ul>
</td></tr></table>
<% else %>
<table cellpadding=0 cellspacing=0 width=450>
<tr style="background-color:#c0c0c0;">
  <th colspan=2 valign=bottom>
  Thank you, the following information has been sent:
  </th>
</tr>
<%   for each name in fieldOrder %>
<tr style="background-color:#ffffff;" valign=top>
  <td><b><% = name %></b>&nbsp;</td>
  <td><% = Request.Form(name) %>&nbsp;</td>
</tr>
<%   next %>
</table>
<% end if %>

</body>
</html>

<% '---------------------------------------------------------------------------
   ' Subroutines and functions.
   '---------------------------------------------------------------------------

   sub AddErrorMsg(msg)

     dim n

    'Add an error message to the list.

     n = UBound(errorMsgs)
     Redim Preserve errorMsgs(n + 1)
     errorMsgs(n + 1) = msg

   end sub

   function GetHost(url)

     dim i, s

     GetHost = ""

     'Strip down to host or IP address and port number, if any.

     if Left(url, 7) = "http://" then
       s = Mid(url, 8)
     elseif Left(url, 8) = "https://" then
       s = Mid(url, 9)
     end if
     i = InStr(s, "/")
     if i > 1 then
       s = Mid(s, 1, i - 1)
     end if

     getHost = s

   end function

   function IsValidEmail(email)

     dim names, name, i, c

     'Check for valid syntax in an email address.

     IsValidEmail = true
     names = Split(email, "@")
     if UBound(names) <> 1 then
       IsValidEmail = false
       exit function
     end if
     for each name in names
       if Len(name) <=  0 then
         IsValidEmail = false
         exit function
       end if
       for i = 1 to Len(name)
         c = Lcase(Mid(name, i, 1))
         if InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 and not IsNumeric(c) then
           IsValidEmail = false
           exit function
         end if
       next
       if Left(name, 1) = "." or Right(name, 1) = "." then
         IsValidEmail = false
         exit function
       end if
     next
     if InStr(names(1), ".") <= 0 then
       IsValidEmail = false
       exit function
     end if
     i = Len(names(1)) - InStrRev(names(1), ".")
     if i <> 2 and i <> 3 then
       IsValidEmail = false
       exit function
     end if
     if InStr(email, "..") > 0 then
       IsValidEmail = false
     end if

   end function

   function FormFieldList()

     dim str, i, name

     'Build an array of form field names ordered as they were received.

     str = ""
     for i = 1 to Request.Form.Count
       for each name in Request.Form
         if Left(name, 1) <> "_" and Request.Form(name) is Request.Form(i) then
           if str <> "" then
             str = str & ","
           end if
           str = str & name
           exit for
         end if
       next
     next
     FormFieldList = Split(str, ",")

   end function

   function SendMail()

     dim mailObj
     dim addrList

     'Send email based on mail component. Uses global variables for parameters
     'because there are so many.

     SendMail = ""

     'Send email (CDONTS version), doesn't support reply-to address and has
     'no error checking.

     if mailComp = "CDONTS" then
       set mailObj = Server.CreateObject("CDONTS.NewMail")
       mailObj.BodyFormat = 0
       mailObj.MailFormat = 0
       mailObj.From = fromAddr
       mailObj.To = recipients
       mailObj.Subject = subject
       mailObj.Body = body
       mailObj.Send
     end if

     'Send email (JMail version).

     if mailComp = "JMail" then
       set mailObj = Server.CreateObject("JMail.SMTPMail")
       mailObj.Silent = true
       mailObj.ServerAddress = smtpServer
       mailObj.Sender = fromAddr
       mailObj.ReplyTo = replyTo
       mailObj.Subject = subject
       addrList = Split(recipients, ",")
       for each addr in addrList
         mailObj.AddRecipient Trim(addr)
       next
       mailObj.ContentType = "text/html"
       mailObj.Body = body
       if not mailObj.Execute then
         SendMail = "Email send failed: " & mailObj.ErrorMessage & "."
       end if
     end if

     'Send email (ASPMail version).

     if mailComp = "ASPMail" then
       set mailObj = Server.CreateObject("SMTPsvg.Mailer")
       mailObj.FromAddress = fromAddr
       mailObj.RemoteHost  = smtpServer
       mailObj.ReplyTo = replyTo
       for each addr in Split(recipients, ",")
         mailObj.AddRecipient "", Trim(addr)
       next
       mailObj.Subject = subject
       mailObj.ContentType = "text/html"
       mailObj.BodyText = body
       if not mailObj.SendMail then
         SendMail = "Email send failed: " & mailObj.Response & "."
       end if
    end if

   end function %>
</PRE>