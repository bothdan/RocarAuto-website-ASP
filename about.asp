<% Session.Timeout = 300 %>
<%
Response.expires = 0
Response.expiresabsolute = Now() - 1
Response.addHeader "pragma", "no-cache"
Response.addHeader "cache-control", "private"
Response.CacheControl = "no-cache"
%>
<!--#include file="db.asp"-->
<!--#include file="aspmkrfn.asp"-->
<% 
displayRecs = 20
recRange = 10
%>
<%
dbwhere = ""
masterdetailwhere = ""
searchwhere = ""
a_search = ""
b_search = ""
whereClause = ""
%>
<%
'Get clear search cmd
If Request.QueryString("cmd").Count > 0 Then
	cmd = Request.QueryString("cmd")
	If UCase(cmd) = "RESET" Then
		'reset search criteria
		searchwhere = ""
		Session("news_searchwhere") =searchwhere
    ElseIf UCase(cmd) = "RESETALL" Then
		'reset search criteria
		searchwhere = ""
		Session("news_searchwhere") =searchwhere
	End If
	'reset start record counter (reset command)
	startRec = 1
	Session("news_REC") = startRec
End If
'construct dbwhere
If masterdetailwhere <> "" Then
	dbwhere = dbwhere & "(" & masterdetailwhere & ") AND "
End If
If searchwhere <> "" Then
	dbwhere = dbwhere & "(" & searchwhere & ") AND "
End If
If Len(dbwhere) > 5 Then
	dbwhere = Mid(dbwhere, 1, Len(dbwhere)-5) 'trim right most AND
End If
%>
<%
' Load Default Order
DefaultOrder = ""
DefaultOrderType = ""
'No Default Filter
DefaultFilter = ""
' Check for an Order parameter
OrderBy = ""
If Request.QueryString("order").Count > 0 Then
	OrderBy = Request.QueryString("order")
	' Check If an ASC/DESC toggle is required
	If Session("news_OB") = OrderBy Then
		If Session("news_OT") = "ASC" Then
			Session("news_OT") = "DESC"
		Else
			Session("news_OT") = "ASC"
		End if
	Else
		Session("news_OT") = "ASC"
	End If
	Session("news_OB") = OrderBy
	Session("news_REC") = 1
Else
	OrderBy = Session("news_OB")
	If OrderBy = "" Then
		OrderBy = DefaultOrder
		Session("news_OB") = OrderBy
		Session("news_OT") = DefaultOrderType
	End If
End If
' Open Connection to the database
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str
' Build SQL
strsql = "SELECT * FROM [news]"
If DefaultFilter <> "" Then
	whereClause = whereClause & "(" & DefaultFilter & ") AND "
End If
If dbwhere <> "" Then
	whereClause = whereClause & "(" & dbwhere & ") AND "
End If
If Right(whereClause, 5)=" AND " Then whereClause = Left(whereClause, Len(whereClause)-5)
If whereClause <> "" Then
	strsql = strsql & " WHERE " & whereClause
End If
If OrderBy <> "" Then 
	strsql = strsql & " ORDER BY [" & OrderBy & "] " & Session("news_OT")
End If	
'Response.Write strsql
Set rs = Server.CreateObject("ADODB.Recordset")
rs.cursorlocation = 3
rs.Open strsql, conn, 1, 2
totalRecs = rs.RecordCount
' Check for a START parameter
If Request.QueryString("start").Count > 0 Then
	startRec = Request.QueryString("start")
	Session("news_REC") = startRec
ElseIf Request.QueryString("pageno").Count > 0 Then
	pageno = Request.QueryString("pageno")
	If IsNumeric(pageno) Then
		startRec = (pageno-1)*displayRecs+1
		If startRec <= 0 Then
			startRec = 1
		ElseIf startRec >= ((totalRecs-1)\displayRecs)*displayRecs+1 Then
			startRec = ((totalRecs-1)\displayRecs)*displayRecs+1
		End If
		Session("news_REC") = startRec
	Else
		startRec = Session("news_REC")
		If Not IsNumeric(startRec) Or startRec = "" Then
			'reset start record counter
			startRec = 1
			Session("news_REC") = startRec
		End If
	End If
Else
	startRec = Session("news_REC")
	If Not IsNumeric(startRec) Or startRec = "" Then
		'reSet start record counter
		startRec = 1
		Session("news_REC") = startRec
	End If
End If
%>
<!--#include file="header.asp"-->
<meta name="generator" content="Namo WebEditor v5.0(Trial)">
<table align="center" cellpadding="0" cellspacing="0" width="801">
    <tr>
        <td colspan="3" background="images/contactbg.gif" width="801">
            <p><img src="images/about.gif" width="247" height="32" border="0"></p>
        </td>
    </tr>
    <tr>
        <td width="18" bgcolor="white">
            <p align="center">&nbsp;</p>
        </td>
        <td width="528" bgcolor="white">
            <div align="left">
                <table cellpadding="0" cellspacing="0" width="490">
                    <tr>
                        <td width="490"><font face="Arial"><span style="font-size:12pt;"><br>Welcome to Rocar Auto Sales!</span><span style="font-size:11pt;"><BR><BR>Rocar 
                             Auto Sales is a name you can trust.  We 
at Rocar Auto Sales are commited to give our customers the best prices, quality 
and the profesionalism you deserve. Having happy and satisfied customers is our 
#1 priority. Ask about our extended warranties and protection plans. <BR><BR>We 
have a strong and committed sales staff  satisfying 
our customers' needs. Feel free to browse our inventory online, request more 
information about vehicles, set up a test drive or inquire about our  
financing!<BR><BR>If you don't see what you are looking for, click on CarFinder 
&amp; simply fill out the form &amp; we will let you know when vehicles arrive 
that match your search! Or if you would rather discuss your options with our 
friendly sales staff, click on Directions for  driving directions and 
other contact information. We look forward to serving you! </span></font><br><br>&nbsp;</td>
                    </tr>
                </table>
            </div>
        </td>
        <td width="255" bgcolor="white" valign="top">            <table align="center" cellpadding="0" cellspacing="0" width="214">
                <tr>
                    <td width="12" height="9">
                        <p><font face="Arial"><span style="font-size:5pt;"><img src="images/upleft.gif" width="12" height="10" border="0"></span></font></p>
                    </td>
                    <td width="194" background="images/up.gif" height="9">
                        <p><font face="Arial"><span style="font-size:5pt;">&nbsp;</span></font></p>
                    </td>
                    <td width="8" height="9">
                        <p><font face="Arial"><span style="font-size:5pt;"><img src="images/upright.gif" width="13" height="10" border="0"></span></font></p>
                    </td>
                </tr>
                <tr>
                    <td width="12" rowspan="3" background="images/leftbg.gif" height="211">
                        <p><font face="Arial"><span style="font-size:8pt;">&nbsp;</span></font></p>
                    </td>
                    <td width="194" nowrap bgcolor="white"><SPAN style="font-size:12pt;"><font face="Arial"><b>News 
                        and Events</b></font></SPAN></td>
                    <td width="8" rowspan="3" background="images/rightbg.gif" height="211">
                        <p><font face="Arial"><span style="font-size:8pt;">&nbsp;</span></font></p>
                    </td>
                </tr>
                <tr>
                    <td width="194" height="4" bgcolor="white" align="left" valign="top">

<form method="post">
<table border="0" cellspacing="0" cellpadding="0" width="27">
<tr bgcolor="#708090">
<td width="27" height="7" bgcolor="white">
                <p><font color="white"><span style="font-size:2pt;">&nbsp;</span></font></p>
</td>
</tr>
<%
'Avoid starting record > total records
If CLng(startRec) > CLng(totalRecs) Then
	startRec = totalRecs
End If
'Set the last record to display
stopRec = startRec + displayRecs - 1
'Move to first record directly for performance reason
recCount = startRec - 1
If NOT rs.EOF Then
	rs.MoveFirst
	rs.Move startRec - 1
End If
recActual = 0
Do While (NOT rs.EOF) AND (recCount < stopRec)
	recCount = recCount + 1
	If CLng(recCount) >= CLng(startRec) Then 
		recActual = recActual + 1 %>
<%
	'Set row color
	bgcolor="#FFFFFF"
%>
<%	
	' Display alternate color for rows
	If recCount Mod 2 <> 0 Then
		bgcolor="#F5F5F5"
	End If
%>
<%
	'Load Key for record
	key = rs("ID")
	x_ID = rs("ID")
	x_news = rs("news")
%>
<tr bgcolor="<%= bgcolor %>">
<td width="27">
<p></td>
</tr>
<%
	End If
	rs.MoveNext
Loop
%>
</table>
</form>
                    </td>
                </tr>
                <tr>
                    <td width="194" nowrap bgcolor="white" height="174" align="left" valign="top">
                        <p align="left">&nbsp;<font face="Arial" size="2"><%If isnull(x_news) Then %> 
                        </font><font face="Arial"><span style="font-size:10pt;"><b>No News at this time.</b> 
                        </span><span style="font-size:9pt;"><br> &nbsp;&nbsp;&nbsp;Please, check back later.</span></font> <font face="Arial" size="2"><%else%><% response.write x_news %><%end if%></font>&nbsp;</td>
                </tr>
                <tr>
                    <td width="12">
                        <p><font face="Arial"><span style="font-size:8pt;"><img src="images/dwleft.gif" width="12" height="11" border="0"></span></font></p>
                    </td>
                    <td width="194" background="images/dwbg.gif">
                        <p><font face="Arial"><span style="font-size:5pt;">&nbsp;</span></font></p>
                    </td>
                    <td width="8">
                        <p><font face="Arial"><span style="font-size:5pt;"><img src="images/dwright.gif" width="13" height="11" border="0"></span></font></p>
                    </td>
                </tr>
            </table>
        </td>
    </tr>
</table>
<%
' Close recordSet and connection
rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing %>
<!--#include file="footer.asp"-->
</p>
