<?xml version="1.0" encoding="UTF-8"?>
<rss version="2.0">
<channel>
<title>POWER OF GOSPEL</title>
<link>http://powerofgospel.org</link>
<description>Power of Gospel</description>
<language>en-us</language>
<copyright>YOURCOPYRIGHT</copyright>


<%

function FormatDateSpecial (formatDate)
	strFormat = Year (formatDate)
	
	if (Month (formatDate) < 10) then strFormat = strFormat & "0"
	strFormat = strFormat & Month (formatDate)
	
	if (Day (formatDate) < 10) then strFormat = strFormat & "0"
	strFormat = strFormat & Day (formatDate)
	
	FormatDateSpecial = strFormat
end function


	dateToCheck = Date
	strFormat = FormatDateSpecial(dateToCheck)
	
	fileName = strFormat & ".txt"

	Set fso = Server.CreateObject ("Scripting.FileSystemObject")
	while ((fso.FileExists (Server.MapPath ("word/" & fileName)) = false) and (i < 10))
		dateToCheck = DateAdd ("d", -1, dateToCheck)

		strFormat = FormatDateSpecial(dateToCheck)
		fileName = strFormat & ".txt"
	Wend
	
	Set f = fso.OpenTextFile (Server.MapPath ("word/" & fileName))
	strTitle = f.ReadLine
	strSynopsis = f.ReadLine
	f.close
%>
<lastBuildDate><%=CurrDateT%></lastBuildDate>
<ttl>240</ttl>
<image>
<url>http://powerofgospel.org/img/logo.jpg</url>
<title><u>Word For the Week</u></title>
<link>powerofgospel.org</link>
</image>

<item>

<title>
<%=strTitle%>
</title>

<link>http://powerofgospel.org/word.asp?<%=fileName %></link>
<description>
<%=strSynopsis  %>
</description>

<pubDate><%=Date%></pubDate>

</item>
</channel>
</rss>
