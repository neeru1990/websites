<%

fname=request("fname")
lname=request("lname")
jt=request("jt")
cname=request("cname")
add1=request("add1")
add2=request("add2")
add=add1+add2
city=request("city")
wph=request("wph")
emailadd=request("emailadd")
url=request("url")
tframe=request("tframe")
nuser=request("nuser")
app=request("app")
app=request("app2")
app=request("app3")
app=request("app4")

acom=request("acom")


stremailto="sales@rincon.net"

strmail="<div align=center>"
strmail=strmail + "<table width=80% border=0 cellspacing=0 cellpadding=0>"
strmail=strmail + "<tr valign=top align=center>"
strmail=strmail + "<td height=499>"
strmail=strmail + "<div align=center>"
strmail=strmail + "<table width=100% border=0 cellspacing=0 cellpadding=0>"
strmail=strmail + "   <tr>"
strmail=strmail + "       <td>&nbsp;</td>"
strmail=strmail + "    </tr>"
strmail=strmail + "</table>"
            
strmail=strmail + "<table width=100% border=0 cellspacing=0 cellpadding=0>"
strmail=strmail + "<tr> "
strmail=strmail +  " <td width=6% >&nbsp;</td>"
strmail=strmail +  " <td width=25% >&nbsp;</td>"
strmail=strmail +  "            <td colspan=3>&nbsp;</td>"
strmail=strmail +   "         </tr>"
strmail=strmail +   "         <tr> "
strmail=strmail +   "           <td width=6% >&nbsp;</td>"
strmail=strmail +   "           <td width=25% class=runb>Title</td>"
strmail=strmail +   "           <td colspan=3>&nbsp; </td>"
strmail=strmail +   "         </tr>"
strmail=strmail +   "         <tr> "
strmail=strmail +   "           <td width=6% >&nbsp;</td>"
strmail=strmail +   "           <td width=25% class=runb>First Name</td>"
strmail=strmail +   "           <td width=24% >&nbsp; </td>"
strmail=strmail +   "           <td width=18% class=runb>Last Name</td>"
strmail=strmail +   "           <td width=27% >&nbsp;</td>"
strmail=strmail +   "         </tr>"
strmail=strmail +   "         <tr> "
strmail=strmail +   "           <td width=6% >&nbsp;</td>"
strmail=strmail +   "           <td width=25% >&nbsp;</td>"
strmail=strmail +   "           <td colspan=3>&nbsp; </td>"
strmail=strmail +   "         </tr>"
strmail=strmail +   "         <tr> "
strmail=strmail +   "           <td width=6% height=18>&nbsp;</td>"
strmail=strmail +   "           <td width=25% height=18 class=runb>Job Title</td>"
strmail=strmail +   "           <td height=18 width=24% >&nbsp; </td>"
strmail=strmail +   "           <td height=18 width=18% class=runb>Company Name</td>"
strmail=strmail +   "           <td height=18 width=27% >&nbsp;</td>"
strmail=strmail +   "         </tr>"
strmail=strmail +   "         <tr>" 
strmail=strmail +   "           <td width=6% >&nbsp;</td>"
strmail=strmail +   "           <td width=25% class=runb>Address</td>"
strmail=strmail +   "           <td colspan=3>&nbsp; </td>"
strmail=strmail +   "         </tr>"
strmail=strmail +   "         <tr> "
strmail=strmail +   "           <td width=6% >&nbsp;</td>"
strmail=strmail +   "           <td width=25% >&nbsp;</td>"
strmail=strmail +   "           <td colspan=3>&nbsp; </td>"
strmail=strmail +   "         </tr>"
strmail=strmail +   "         <tr>" 
strmail=strmail +   "           <td width=6% >&nbsp;</td>"
strmail=strmail +   "           <td width=25% class=runb>City</td>"
strmail=strmail +   "           <td colspan=3>&nbsp; </td>"
strmail=strmail +   "         </tr>"
strmail=strmail +   "         <tr>" 
strmail=strmail +   "           <td width=6% >&nbsp;</td>"
strmail=strmail +   "           <td width=25% class=runb>Work Phone</td>"
strmail=strmail +   "           <td colspan=3>&nbsp; </td>"
strmail=strmail +   "         </tr>
strmail=strmail +   "         <tr>" 
strmail=strmail +   "           <td width=6% >&nbsp;</td>"
strmail=strmail +   "           <td width=25% class=runb>Email</td>"
strmail=strmail +   "<td colspan=3>&nbsp; </td>"
strmail=strmail +   "         </tr>"
strmail=strmail +   "         <tr>" 
strmail=strmail +   "           <td width=6% >&nbsp;</td>"
strmail=strmail +   "           <td width=25% class=runb>Web Site URL</td>"
strmail=strmail +   "           <td colspan=3>&nbsp; </td>"
strmail=strmail +   "         </tr>"
strmail=strmail +   "         <tr>" 
strmail=strmail +   "           <td width=6% height=10>&nbsp;</td>"
strmail=strmail +   "           <td width=25% height=10>&nbsp;</td>"
strmail=strmail +   "           <td colspan=3 height=10>&nbsp;</td>"
strmail=strmail +   "         </tr>"
strmail=strmail +   "         <tr>" 
strmail=strmail +   "           <td width=6% height=5>&nbsp;</td>"
strmail=strmail +   "           <td width=25% height=5 class=runb>User Type</td>"
strmail=strmail +   "           <td colspan=3 height=5>&nbsp; </td>"
strmail=strmail +   "         </tr>"
strmail=strmail +   "         <tr>" 
strmail=strmail +   "          <td width=6% height=7>&nbsp;</td>"
strmail=strmail +   "           <td width=25% height=7>&nbsp;</td>"
strmail=strmail +   "           <td colspan=3 height=7>&nbsp;</td>"
strmail=strmail +   "         </tr>"
strmail=strmail +   "         <tr>" 
strmail=strmail +   "           <td width=6% height=14>&nbsp;</td>"
strmail=strmail +   "           <td width=25% height=14>&nbsp;</td>"
strmail=strmail +   "           <td colspan=3 height=14>&nbsp; </td>"
strmail=strmail +   "         </tr>"
strmail=strmail +   "           <td width=6% height=2>&nbsp;</td>"
strmail=strmail +   "           <td width=25% height=2 class=runb>Purchase timeframe</td>"
strmail=strmail +   "           <td colspan=3 height=2>&nbsp;</td>"
strmail=strmail +   "         </tr>"
strmail=strmail +   "         <tr>" 
strmail=strmail +   "           <td width=6% >&nbsp;</td>"
strmail=strmail +   "           <td colspan=4>&nbsp; </td>"
strmail=strmail +   " </tr>"
strmail=strmail +   "         <tr> "
strmail=strmail +   "           <td width=6% >&nbsp;</td>"
strmail=strmail +   "          <td colspan=4>&nbsp;</td>"
strmail=strmail +   "        </tr>"
strmail=strmail +   "        <tr>" 
strmail=strmail +   "          <td width=6% >&nbsp;</td>"
strmail=strmail +   "          <td width=25% class=runb>Network users </td>"
strmail=strmail +   "          <td colspan=3>&nbsp;</td>"
strmail=strmail +   "        </tr>"
strmail=strmail +   "         <tr>" 
strmail=strmail +   "           <td width=6% >&nbsp;</td>"
strmail=strmail +   "           <td colspan=4>&nbsp; </td>"
strmail=strmail +   "         </tr>"
strmail=strmail +   "         <tr>" 
strmail=strmail +   "           <td width=6% >&nbsp;</td>"
strmail=strmail +   "           <td colspan=4>&nbsp;</td>"
strmail=strmail +   "         </tr>"
strmail=strmail +   "         <tr> "
strmail=strmail +   "           <td width=6% >&nbsp;</td>"
strmail=strmail +   "           <td width=25% class=runb>Applications Used</td>"
strmail=strmail +   "           <td colspan=3>&nbsp;</td>"
strmail=strmail +   "         </tr>"
strmail=strmail +   "         <tr> "
strmail=strmail +   "           <td width=6% >&nbsp;</td>"
strmail=strmail +   "           <td colspan=4>&nbsp; </td>"
strmail=strmail +   "         </tr>"
strmail=strmail +   "         <tr> "
strmail=strmail +   "           <td width=6% >&nbsp;</td>"
strmail=strmail +   "           <td colspan=4>&nbsp;</td>"
strmail=strmail +   "         </tr>"
strmail=strmail +   "         <tr> "
strmail=strmail +   "           <td width=6% >&nbsp;</td>"
strmail=strmail +   "           <td width=25% class=runb>E-mail system used</td>"
strmail=strmail +   "           <td colspan=3>&nbsp;</td>"
strmail=strmail +   "         </tr>"
strmail=strmail +   "         <tr> "
strmail=strmail +   "          <td width=6% height=8>&nbsp;</td>"
strmail=strmail +   "           <td colspan=4 valign=top height=8>&nbsp; </td>"
strmail=strmail +   "         </tr>"
strmail=strmail +   "    <tr>" 
strmail=strmail +   "      <td width=6% >&nbsp;</td>"
strmail=strmail +   "      <td colspan=4 valign=top>&nbsp;</td>"
strmail=strmail +   "    </tr>"
strmail=strmail +   "    <tr>" 
strmail=strmail +   "      <td width=6% >&nbsp;</td>"
strmail=strmail +   "      <td width=25% valign=top class=runb>Additional Comments</td>"
strmail=strmail +   "      <td valign=top colspan=3>&nbsp;</td>"
strmail=strmail +   "    </tr>"
strmail=strmail +   "    <tr>" 
strmail=strmail +   "      <td width=6% height=16>&nbsp;</td>"
strmail=strmail +   "      <td colspan=4 valign=top height=16>&nbsp; </td>"
strmail=strmail +   "    </tr>"
strmail=strmail +   "    <tr>" 
strmail=strmail +   "      <td width=6% >&nbsp;</td>"
strmail=strmail +   "      <td colspan=4 valign=top>&nbsp;</td>"
strmail=strmail +   "    </tr>"
strmail=strmail +   "    <tr>" 
strmail=strmail +   "      <td width=6% height=2>&nbsp;</td>"
strmail=strmail +   "      <td colspan=4 valign=top height=2>&nbsp; </td>"
strmail=strmail +   "    </tr>"
strmail=strmail +   "  </table>"
strmail=strmail +   "  
strmail=strmail +   "  </div>"
strmail=strmail +   " </td>"
strmail=strmail +   "</tr>"
strmail=strmail +  "</table>"
strmail=strmail +"</div>"



set objMail=Server.CreateObject("CDONTS.NewMail")

objMail.Mailformat=0
objMail.BodyFormat=0
objMail.To=stremailto
objMail.Subject="User details for " & 
objMail.From="order@daftaronline.com"
objMail.Body=strmail
objMail.Send

%>



<html>
<head>
<title>Right Fax  - Awards</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

<link rel="stylesheet" href="sty.css"></head>

<body bgcolor="#FFFFFF" background="back.gif" text="#000099" link="#000099" vlink="#000099" alink="#000099" topmargin=0 leftmargin = 0 rightmargin=0>
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center" >
  <tr valign="top" align="center"> 
    <td width="20%" bgcolor="#15399F" valign="top"> 
      <div align="center"> 
        <table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
          <tr valign="top"> 
            <td><img src="rflogo.gif"></td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
          </tr>
          <tr> 
            <td> 
              <div align="center"><a href="rfindex.htm" target="content"><img src="rff.gif" width="110" height="20" border="0"></a></div>
            </td>
          </tr>
          <tr> 
            <td> 
              <div align="center"><a href="rfps.htm" target="content"><img src="rfps.gif" width="110" height="20" border="0"></a></div>
            </td>
          </tr>
          <tr> 
            <td height="2"> 
              <div align="center"><a href="rfws.htm" target="content"><img src="rfwser.gif" width="110" height="20" border="0"></a></div>
            </td>
          </tr>
          <tr> 
            <td height="2"> 
              <div align="center"><a href="rfcomp.htm" target="content"><img src="rfcom.gif" width="110" height="20" border="0"></a></div>
            </td>
          </tr>
          <tr> 
            <td height="2"> 
              <div align="center"><a href="rfsr.htm" target="content"><img src="rfsr.gif" width="110" height="20" border="0"></a></div>
            </td>
          </tr>
          <tr> 
            <td height="2"> 
              <div align="center"><a href="rfaw.htm"><img src="rfaw.gif" width="110" height="20" border="0"></a></div>
            </td>
          </tr>
          <tr> 
            <td height="2"> 
              <div align="center"><img src="rfdown1.gif" width="110" height="20" border="0"></div>
            </td>
          </tr>
          <tr> 
            <td height="2">&nbsp;</td>
          </tr>
          <tr> 
            <td height="2">&nbsp;</td>
          </tr>
          <tr> 
            <td >&nbsp;</td>
          </tr>
          <tr> 
            <td height="2"> 
              <div align="center"></div>
            </td>
          </tr>
          <tr> 
            <td height="2"> 
              <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#CC0000"><b></b></font></div>
            </td>
          </tr>
          <tr> 
            <td height="2">&nbsp;</td>
          </tr>
          <tr> 
            <td height="2">&nbsp;</td>
          </tr>
          <tr> 
            <td height="2">&nbsp;</td>
          </tr>
          <tr> 
            <td height="2">&nbsp;</td>
          </tr>
          <tr> 
            <td height="2">&nbsp;</td>
          </tr>
          <tr> 
            <td height="2">&nbsp;</td>
          </tr>
          <tr> 
            <td height="2">&nbsp;</td>
          </tr>
          <tr> 
            <td height="2">&nbsp;</td>
          </tr>
          <tr> 
            <td height="2">&nbsp;</td>
          </tr>
          <tr> 
            <td height="2">&nbsp;</td>
          </tr>
          <tr> 
            <td height="2">&nbsp;</td>
          </tr>
          <tr> 
            <td height="2">&nbsp;</td>
          </tr>
          <tr> 
            <td height="2">&nbsp;</td>
          </tr>
          <tr> 
            <td height="2">&nbsp;</td>
          </tr>
          <tr> 
            <td height="2">&nbsp;</td>
          </tr>
          <tr> 
            <td height="2">&nbsp;</td>
          </tr>
          <tr> 
            <td height="2">&nbsp;</td>
          </tr>
          <tr> 
            <td height="2">&nbsp;</td>
          </tr>
          <tr> 
            <td height="2"> 
              <div align="center"></div>
            </td>
          </tr>
          <tr> 
            <td height="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#CC0000"><b></b></font></td>
          </tr>
          <tr> 
            <td height="2">&nbsp;</td>
          </tr>
          <tr> 
            <td height="2">&nbsp;</td>
          </tr>
          <tr> 
            <td height="2"> 
              <div align="center"></div>
            </td>
          </tr>
          <tr> 
            <td height="2"> 
              <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#CC0000"><b></b></font></div>
            </td>
          </tr>
          <tr> 
            <td height="2">&nbsp;</td>
          </tr>
          <tr> 
            <td height="2">&nbsp;</td>
          </tr>
          <tr> 
            <td height="2"> 
              <div align="center"></div>
            </td>
          </tr>
          <tr> 
            <td height="2"> 
              <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#CC0000"><b></b></font></div>
            </td>
          </tr>
          <tr> 
            <td height="2">&nbsp;</td>
          </tr>
          <tr> 
            <td height="2">&nbsp;</td>
          </tr>
          <tr> 
            <td height="2"> 
              <div align="center"></div>
            </td>
          </tr>
          <tr> 
            <td height="2"> 
              <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#CC0000"><b></b></font></div>
            </td>
          </tr>
          <tr> 
            <td height="2">&nbsp;</td>
          </tr>
          <tr> 
            <td height="2">&nbsp;</td>
          </tr>
          <tr> 
            <td height="2">&nbsp;</td>
          </tr>
          <tr> 
            <td height="2">&nbsp;</td>
          </tr>
          <tr> 
            <td height="2">
              <div align="center"><img src="avtbot.gif" width="50" ></div>
            </td>
          </tr>
          <tr>
            <td height="2">
              <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#CC0000"><b><font color="#FFFFFF">One 
                Place to Manage Enterprise Communication</font></b></font></div>
            </td>
          </tr>
        </table>
      </div>
    </td>
    <td align="left"> 
      <div align="center"> 
        <table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
          <tr valign="top" align="center"> 
            <td > 
              <div align="center"> </div>
              <div align="center"> 
                <table width="95%" border="0" cellspacing="1" cellpadding="1">
                  <tr valign="top"> 
                    <td class="head1" > 
                      <p align="left"><font class="run"><span class="head1"><span class="head1">RightFAX 
                        </span></span></font><span class="head1">Download </span> 
                      </p>
                      <div align="center">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
                          <tr valign="top" align="center"> 
                            <td> 
                              <div align="center"> 
                                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                  <tr> 
                                    <td class="run"><span class="head3">Brochures</span> 
                                      and <span class="head3">White Papers</span>.</td>
                                  </tr>
                                </table>
                                
                              </div>
                            </td>
                          </tr>
                        </table>
                      </div>
                      <p align="left">&nbsp; </p>
                      </td>
                  </tr>
                </table>
              </div>
            </td>
          </tr>
        </table>
      </div>
    </td>
  </tr>
</table>
</body>
</html>
