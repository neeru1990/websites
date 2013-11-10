<html xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xsl:version="1.0">
	<head>
		<title>Coming Soon</title>
		<!-- #include file="includes/styles.htm" -->
		<SCRIPT>
window.onload=init;
function init()
{
	for(i in document.links) document.links[i].onfocus = document.links[i].blur;
}

		</SCRIPT>
		<script language="vbscript" runat="server">
function FillWord ()
	Set fso = Server.CreateObject ("Scripting.FileSystemObject")
	Set f = fso.GetFolder(Server.MapPath ("word/"))
	Set fc = f.Files
	
	if (fc.count <> 0) then
		For Each f1 in fc
			if (Left (f1.name, 2) = "20") then
				Response.Write ("<tr><td class=""TextDecoration"">")
				strFile = f1.name
				strNewFile = Mid (strFile, 5, 2)
				strNewFile = strNewFile & "/" & Mid (strFile, 7, 2)
				strNewFile = strNewFile & "/" & Left (strFile, 4)
				Response.Write (strNewFile)
				Response.Write ("</td><td class=""TextDecoration"" bgcolor=#F6FFF6><font color=#993300><b>")
				Response.Write ("<a href=""word.asp?" & strFile & """>")
				Set fo = fso.OpenTextFile (f1.Path)
				strWord = fo.ReadLine
				strDescription = fo.ReadLine
				fo.close
				Response.Write (strWord)
				Response.Write ("</a>")
				Response.Write ("</b></font></td><td class=""TextDecoration"">")
				Response.Write (strDescription)
				Response.Write ("</td></tr>")
			end if
		Next
	end if
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
                <td valign="middle"> 
                  <table width="100%" height="100%" cellspacing="0" cellpadding="5" height="47" bordercolor="#339933">
                    <tr> 
                      <td> 
                        <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="4"><b>Coming Soon... Please check back later.</b></font></div>
                      </td>
                    </tr>
                  </table>
                  <br>
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
