<% If Session("rocar_status") <> "login" Then Response.Redirect "login.asp" %>
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
displayRecs = 50
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
		Session("moreinfo_searchwhere") =searchwhere
    ElseIf UCase(cmd) = "RESETALL" Then
		'reset search criteria
		searchwhere = ""
		Session("moreinfo_searchwhere") =searchwhere
	End If
	'reset start record counter (reset command)
	startRec = 1
	Session("moreinfo_REC") = startRec
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
DefaultOrder = "morinfodate"
DefaultOrderType = "DESC"
'No Default Filter
DefaultFilter = ""
' Check for an Order parameter
OrderBy = ""
If Request.QueryString("order").Count > 0 Then
	OrderBy = Request.QueryString("order")
	' Check If an ASC/DESC toggle is required
	If Session("moreinfo_OB") = OrderBy Then
		If Session("moreinfo_OT") = "ASC" Then
			Session("moreinfo_OT") = "DESC"
		Else
			Session("moreinfo_OT") = "ASC"
		End if
	Else
		Session("moreinfo_OT") = "ASC"
	End If
	Session("moreinfo_OB") = OrderBy
	Session("moreinfo_REC") = 1
Else
	OrderBy = Session("moreinfo_OB")
	If OrderBy = "" Then
		OrderBy = DefaultOrder
		Session("moreinfo_OB") = OrderBy
		Session("moreinfo_OT") = DefaultOrderType
	End If
End If
' Open Connection to the database
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str
' Build SQL
strsql = "SELECT * FROM [moreinfo]"
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
	strsql = strsql & " ORDER BY [" & OrderBy & "] " & Session("moreinfo_OT")
End If	
'Response.Write strsql
Set rs = Server.CreateObject("ADODB.Recordset")
rs.cursorlocation = 3
rs.Open strsql, conn, 1, 2
totalRecs = rs.RecordCount
' Check for a START parameter
If Request.QueryString("start").Count > 0 Then
	startRec = Request.QueryString("start")
	Session("moreinfo_REC") = startRec
ElseIf Request.QueryString("pageno").Count > 0 Then
	pageno = Request.QueryString("pageno")
	If IsNumeric(pageno) Then
		startRec = (pageno-1)*displayRecs+1
		If startRec <= 0 Then
			startRec = 1
		ElseIf startRec >= ((totalRecs-1)\displayRecs)*displayRecs+1 Then
			startRec = ((totalRecs-1)\displayRecs)*displayRecs+1
		End If
		Session("moreinfo_REC") = startRec
	Else
		startRec = Session("moreinfo_REC")
		If Not IsNumeric(startRec) Or startRec = "" Then
			'reset start record counter
			startRec = 1
			Session("moreinfo_REC") = startRec
		End If
	End If
Else
	startRec = Session("moreinfo_REC")
	If Not IsNumeric(startRec) Or startRec = "" Then
		'reSet start record counter
		startRec = 1
		Session("moreinfo_REC") = startRec
	End If
End If
%>
<!--#include file="header.asp"-->
<meta name="generator" content="Namo WebEditor v5.0(Trial)">
<body bgcolor="white" text="black" link="green" vlink="black" alink="red">
<table align="center" cellpadding="0" cellspacing="0" width="801" bgcolor="white">
    <tr>
        <td height="1">
            <div align="left">
                <table cellpadding="0" cellspacing="0" bgcolor="white" width="360">
                    <tr>
                        <td width="20" height="14">
                            <p></p>
                        </td>
                        <td width="5" height="14">
                <p align="center"><font face="Arial"><span style="font-size:12pt;">|</span></font></p>
                        </td>
                        <td width="100" bgcolor="white" height="14">
                            <p align="center"><span style="font-size:10pt;"><a href="adminlist.asp"><b><font face="Arial" color="navy">Inventory</font></b></a></span></p>
                        </td>
                        <td width="5" bgcolor="white" height="14">
                <p align="center"><font face="Arial"><span style="font-size:12pt;">|</span></font></p>
                        </td>
                        <td width="100" height="14">
                            <p align="center"><b><span style="font-size:10pt;"><font face="Arial" color="green">Customers</font></span></b></p>
                        </td>
                        <td width="5" height="14">
                <p align="center"><span style="font-size:12pt;"><font face="Arial">|</font></span></p>
                        </td>
                        <td width="120" height="14">
                            <p align="center"><a href="newslist.asp"><b><span style="font-size:10pt;"><font face="Arial" color="navy">News 
                            And Events</font></span></b></a></p>
                        </td>
                        <td width="5" height="14">
                <p align="center"><span style="font-size:12pt;"><font face="Arial">|</font></span></p>
                        </td>
                    </tr>
                </table>
</div>
            
</td>
    </tr>
    <tr>
        <td height="3">

<div align="left">
    <table cellpadding="0" cellspacing="0" width="695" bgcolor="white">
        <tr>
            <td width="20" height="5">
                <p></p>
            </td>
            <td width="5" height="5" bgcolor="white">
                <p align="center"></p>
            </td>
            <td width="100" height="5" bgcolor="#979EE0">
                <p align="center"><a href="contactlist.asp"><span style="font-size:10pt;"><b><font face="Arial" color="white">Contact</font></b></span></a></p>
            </td>
            <td width="5" height="5" bgcolor="white">
                <p align="center"></p>
            </td>
            <td width="100" height="5" bgcolor="#979EE0">
                <p align="center"><a href="makeofferlist.asp"><span style="font-size:10pt;"><font face="Arial" color="white"><b>Offers</b></font></span></a></p>
            </td>
            <td width="5" height="5" bgcolor="white">
                <p align="center"><span style="font-size:10pt;"><font face="Arial" color="white"><b>|</b></font></span></p>
            </td>
            <td width="120" height="5" bgcolor="#6351B7">
                <p align="center"><span style="font-size:10pt;"><font face="Arial" color="white"><b>More 
                Info</b></font></span></p>
            </td>
            <td width="5" height="5" bgcolor="white">
                <p align="center"><span style="font-size:10pt;"><font face="Arial" color="white"><b>|</b></font></span></p>
            </td>
            <td width="100" height="5" bgcolor="#979EE0">
                <p align="center"><a href="findelist.asp"><span style="font-size:10pt;"><font face="Arial" color="white"><b>Find 
                Car</b></font></span></a></p>
            </td>
            <td width="5" height="5" bgcolor="white">
                <p align="center"><span style="font-size:10pt;"><font face="Arial" color="white"><b>|</b></font></span></p>
            </td>
            <td width="120" height="5" bgcolor="#979EE0">
                <p align="center"><a href="creditlist.asp"><span style="font-size:10pt;"><font face="Arial" color="white"><b>Credit 
                Application</b></font></span></a></p>
            </td>
            <td width="5" height="5" bgcolor="white">
                <p align="center"><span style="font-size:10pt;"><font face="Arial" color="white"><b>|</b></font></span></p>
            </td>
            <td width="100" height="5" bgcolor="#979EE0">
                <p align="center"><a href="tradeinlist.asp"><span style="font-size:10pt;"><font face="Arial" color="white"><b>Trade-In</b></font></span></a></p>
            </td>
            <td width="5" height="5">
                <p align="center"></p>
            </td>
        </tr>
    </table>
</div>
</td>
    </tr>
    <tr>
        <td>
            <form method="post">
                <p align="left"><font face="Arial"><span style="font-size:11pt;"><br><b><i>&nbsp;&nbsp;&nbsp;More 
                Info List:</i></b><br></span></font></p>
<table cellspacing="0" cellpadding="0" bgcolor="#CCCCCC" width="750" align="center">
<tr bgcolor="#708090">
<td width="76" bgcolor="white" height="20" background="images/sortbar.gif">
                            <p align="left"><a href='moreinfolist.asp?order=<%= Server.URLEncode("morinfodate") %>'><font color="black" face="Arial"><span style="font-size:8pt;"><b>Date</b></span></font></a><font color="black" face="Webdings"><span style="font-size:8pt;"><b> 
                            <% If OrderBy = "morinfodate" Then %><% If Session("moreinfo_OT") = "ASC" Then %>5<% ElseIf Session("moreinfo_OT") = "DESC" Then %>6<% End If %><% End If %></b></span></font>
</td>
<td width="201" bgcolor="white" height="20" background="images/sortbar.gif">
<a href="moreinfolist.asp?order=<%= Server.URLEncode("first") %>"><font face="Arial" color="black"><span style="font-size:8pt;"><b>First 
                            &amp; Last 
                            Name</b></span></font></a><font color="black" face="Arial"><b><span style="font-size:8pt;">&nbsp;</span></b></font><font color="black"><b><span style="font-size:8pt;"><% If OrderBy = "first" Then %></span></b></font><font color="black" face="Webdings"><b><span style="font-size:8pt;"><% If Session("moreinfo_OT") = "ASC" Then %>5<% ElseIf Session("moreinfo_OT") = "DESC" Then %>6<% End If %></span></b></font><font color="black"><b><span style="font-size:8pt;"><% End If %></span></b></font>
</td>
<td width="119" bgcolor="white" height="20" background="images/sortbar.gif">
                            <p align="center"><a href="moreinfolist.asp?order=<%= Server.URLEncode("phone") %>"><font face="Arial" color="black"><b><span style="font-size:8pt;">Phone</span></b></font></a><font color="black" face="Arial"><b><span style="font-size:8pt;">&nbsp;</span></b></font><font color="black"><b><span style="font-size:8pt;"><% If OrderBy = "phone" Then %></span></b></font><font color="black" face="Webdings"><b><span style="font-size:8pt;"><% If Session("moreinfo_OT") = "ASC" Then %>5<% ElseIf Session("moreinfo_OT") = "DESC" Then %>6<% End If %></span></b></font><font color="black"><b><span style="font-size:8pt;"><% End If %></span></b></font>
</td>
<td width="117" bgcolor="white" height="20" background="images/sortbar.gif">
                            <p align="center"><a href="moreinfolist.asp?order=<%= Server.URLEncode("testdrive") %>"><font color="black" face="Arial"><b><span style="font-size:8pt;">Test 
                            Drive</span></b></font></a><font color="black" face="Arial"><b><span style="font-size:8pt;">&nbsp;</span></b></font><font color="black"><b><span style="font-size:8pt;"><% If OrderBy = "testdrive" Then %></span></b></font><font color="black" face="Webdings"><b><span style="font-size:8pt;"><% If Session("moreinfo_OT") = "ASC" Then %>5<% ElseIf Session("moreinfo_OT") = "DESC" Then %>6<% End If %></span></b></font><font color="black"><b><span style="font-size:8pt;"><% End If %></span></b></font>
</td>
<td width="113" bgcolor="white" height="20" background="images/sortbar.gif">
                            <p align="center"><a href="moreinfolist.asp?order=<%= Server.URLEncode("stock") %>"><font color="black" face="Arial"><b><span style="font-size:8pt;">Stock 
                            #</span></b></font></a><font color="black" face="Arial"><b><span style="font-size:8pt;">&nbsp;</span></b></font><font color="black"><b><span style="font-size:8pt;"><% If OrderBy = "stock" Then %></span></b></font><font color="black" face="Webdings"><b><span style="font-size:8pt;"><% If Session("moreinfo_OT") = "ASC" Then %>5<% ElseIf Session("moreinfo_OT") = "DESC" Then %>6<% End If %></span></b></font><font color="black"><b><span style="font-size:8pt;"><% End If %></span></b></font>
</td>
<td width="124" bgcolor="white" height="20" colspan="3" background="images/sortbar.gif"><font color="black"><span style="font-size:8pt;"><b>&nbsp;</b></span></font></td>
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
	x_first = rs("first")
	x_last = rs("last")
	x_phone = rs("phone")
	x_email = rs("email")
	x_testdrive = rs("testdrive")
	x_comments = rs("comments")
	x_stock = rs("stock")
	x_morinfodate = rs("morinfodate")
%>
<tr bgcolor="<%= bgcolor %>">
<td width="76"><font face="Verdana"><span style="font-size:8pt;">&nbsp;<% response.write x_morinfodate%>&nbsp;</span></font></td>
<td width="201"><font face="Verdana"><span style="font-size:8pt;"><% response.write x_first %>&nbsp;<% response.write x_last %>&nbsp;</span></font></td>
<td width="119">
                            <p align="center"><font face="Verdana"><span style="font-size:8pt;"><% response.write x_phone %></span></font></td>
<td width="117">
                            <p align="center"><font face="Verdana"><span style="font-size:8pt;"><%
Select Case x_testdrive
    Case "Yes" response.write "Yes"
    Case "No" response.write ""
End Select
%></span></font></td>
<td width="113">
                            <p align="center"><font face="Verdana"><span style="font-size:8pt;"><% response.write x_stock %></span></font></td>
<td width="51">
                            <p align="center"><a href="<% key = rs("ID") : If not isnull(key) Then response.write "moreinfoview.asp?key=" & Server.URLEncode(key) Else response.write "javascript:alert('Invalid Record! Key is null');" End If %>"><font face="Arial"><span style="font-size:8pt;"><b>View</b></span></font></a></td>
<td width="36">
                            <p align="center"><a href="<% key = rs("ID") : If not isnull(key) Then response.write "moreinfoedit.asp?key=" & Server.URLEncode(key) Else response.write "javascript:alert('Invalid Record! Key is null');" End If %>"><font face="Arial"><span style="font-size:8pt;"><b>Edit</b></span></font></a></td>
<td width="37">
                            <p align="right"><input type="checkbox" name="key" value="<%= key %>"></td>
</tr>
<%
	End If
	rs.MoveNext
Loop
%>
</table>
<% If recActual > 0 Then %>
<p align="right"><input type="button" name="btndelete" value="DELETE SELECTED" onClick="this.form.action='moreinfodelete.asp';this.form.submit();" style="font-family:Arial; font-size:12;"> 
                &nbsp;&nbsp;&nbsp;&nbsp;</p>
<% End If %>
</form>
<table border="0" cellspacing="0" cellpadding="10" width="390" align="center"><tr><td width="370" height="73" align="center">
                        <font face="Arial"><span style="font-size:8pt;"><%
If totalRecs > 0 Then
	rsEof = (totalRecs < (startRec + displayRecs))
	PrevStart = startRec - displayRecs
	If PrevStart < 1 Then PrevStart = 1
	NextStart = startRec + displayRecs
	If NextStart > totalRecs Then NextStart = startRec
	LastStart = ((totalRecs-1)\displayRecs)*displayRecs+1
	%></span></font><table border="0" cellspacing="0" cellpadding="0" align="center"><tr><td><font face="Arial"><span style="font-size:8pt;">Page&nbsp;</span></font></td>
<!--first page button-->
	<% If CLng(startRec)=1 Then %>
	<td><font face="Arial"><span style="font-size:8pt;"><img src="images/firstdisab.gif" alt="First" width="20" height="15" border="0"></span></font></td>
	<% Else %>
	<td><a href="moreinfolist.asp?start=1"><font face="Arial"><span style="font-size:8pt;"><img src="images/first.gif" alt="First" width="20" height="15" border="0"></span></font></a></td>
	<% End If %>
<!--previous page button-->
	<% If CLng(PrevStart) = CLng(startRec) Then %>
	<td><font face="Arial"><span style="font-size:8pt;"><img src="images/prevdisab.gif" alt="Previous" width="20" height="15" border="0"></span></font></td>
	<% Else %>
	<td><a href="moreinfolist.asp?start=<%=PrevStart%>"><font face="Arial"><span style="font-size:8pt;"><img src="images/prev.gif" alt="Previous" width="20" height="15" border="0"></span></font></a></td>
	<% End If %>
<!--current page number-->
	<td><font face="Arial"><span style="font-size:8pt;"><input type="text" name="pageno" value="<%=(startRec-1)\displayRecs+1%>" size="4" style="font-size: 9pt;"></span></font></td>
<!--next page button-->
	<% If CLng(NextStart) = CLng(startRec) Then %>
	<td><font face="Arial"><span style="font-size:8pt;"><img src="images/nextdisab.gif" alt="Next" width="20" height="15" border="0"></span></font></td>
	<% Else %>
	<td><a href="moreinfolist.asp?start=<%=NextStart%>"><font face="Arial"><span style="font-size:8pt;"><img src="images/next.gif" alt="Next" width="20" height="15" border="0"></span></font></a></td>
	<% End If %>
<!--last page button-->
	<% If CLng(LastStart) = CLng(startRec) Then %>
	<td><font face="Arial"><span style="font-size:8pt;"><img src="images/lastdisab.gif" alt="Last" width="20" height="15" border="0"></span></font></td>
	<% Else %>
	<td><a href="moreinfolist.asp?start=<%=LastStart%>"><font face="Arial"><span style="font-size:8pt;"><img src="images/last.gif" alt="Last" width="20" height="15" border="0"></span></font></a></td>
	<% End If %>
	</tr></table>	
                        <p align="center"><font face="Arial"><span style="font-size:8pt;"><% If CLng(startRec) > CLng(totalRecs) Then startRec = totalRecs
	stopRec = startRec + displayRecs - 1
	recCount = totalRecs - 1
	If rsEOF Then recCount = totalRecs
	If stopRec > recCount Then stopRec = recCount %>
	Records <%= startRec %> to <%= stopRec %> of <%= totalRecs %>
<% Else %>
	No records found</span></font><form>	
                            <font face="Arial"><span style="font-size:8pt;"><% End If %></span></font></form>	

</td></tr></table>
<%
' Close recordSet and connection
rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing %></td>
    </tr>
</table>

<!--#include file="footer.asp"-->
