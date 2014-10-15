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
		Session("contact_searchwhere") =searchwhere
    ElseIf UCase(cmd) = "RESETALL" Then
		'reset search criteria
		searchwhere = ""
		Session("contact_searchwhere") =searchwhere
	End If
	'reset start record counter (reset command)
	startRec = 1
	Session("contact_REC") = startRec
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
DefaultOrder = "contactdate"
DefaultOrderType = "DESC"
'No Default Filter
DefaultFilter = ""
' Check for an Order parameter
OrderBy = ""
If Request.QueryString("order").Count > 0 Then
	OrderBy = Request.QueryString("order")
	' Check If an ASC/DESC toggle is required
	If Session("contact_OB") = OrderBy Then
		If Session("contact_OT") = "ASC" Then
			Session("contact_OT") = "DESC"
		Else
			Session("contact_OT") = "ASC"
		End if
	Else
		Session("contact_OT") = "ASC"
	End If
	Session("contact_OB") = OrderBy
	Session("contact_REC") = 1
Else
	OrderBy = Session("contact_OB")
	If OrderBy = "" Then
		OrderBy = DefaultOrder
		Session("contact_OB") = OrderBy
		Session("contact_OT") = DefaultOrderType
	End If
End If
' Open Connection to the database
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str
' Build SQL
strsql = "SELECT * FROM [contact]"
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
	strsql = strsql & " ORDER BY [" & OrderBy & "] " & Session("contact_OT")
End If	
'Response.Write strsql
Set rs = Server.CreateObject("ADODB.Recordset")
rs.cursorlocation = 3
rs.Open strsql, conn, 1, 2
totalRecs = rs.RecordCount
' Check for a START parameter
If Request.QueryString("start").Count > 0 Then
	startRec = Request.QueryString("start")
	Session("contact_REC") = startRec
ElseIf Request.QueryString("pageno").Count > 0 Then
	pageno = Request.QueryString("pageno")
	If IsNumeric(pageno) Then
		startRec = (pageno-1)*displayRecs+1
		If startRec <= 0 Then
			startRec = 1
		ElseIf startRec >= ((totalRecs-1)\displayRecs)*displayRecs+1 Then
			startRec = ((totalRecs-1)\displayRecs)*displayRecs+1
		End If
		Session("contact_REC") = startRec
	Else
		startRec = Session("contact_REC")
		If Not IsNumeric(startRec) Or startRec = "" Then
			'reset start record counter
			startRec = 1
			Session("contact_REC") = startRec
		End If
	End If
Else
	startRec = Session("contact_REC")
	If Not IsNumeric(startRec) Or startRec = "" Then
		'reSet start record counter
		startRec = 1
		Session("contact_REC") = startRec
	End If
End If
%>
<!--#include file="header.asp"-->
<meta name="generator" content="Namo WebEditor v5.0(Trial)">
<body bgcolor="white" text="black" link="green" vlink="black" alink="red">
<table align="center" width="801" bgcolor="white" cellpadding="0" cellspacing="0">
    <tr>
        <td>
            <div align="left">
                <table cellpadding="0" cellspacing="0" bgcolor="white" width="360">
                    <tr>
                        <td width="20" height="8">
                            <p></p>
                        </td>
                        <td width="5" height="8">
                            <p align="center"><font face="Arial"><span style="font-size:12pt;">|</span></font></p>
                        </td>
                        <td width="100" bgcolor="white" height="8">
                            <p align="center"><span style="font-size:10pt;"><a href="adminlist.asp"><b><font face="Arial" color="navy">Inventory</font></b></a></span></p>
                        </td>
                        <td width="5" bgcolor="white" height="8">
                            <p align="center"><font face="Arial"><span style="font-size:12pt;">|</span></font></p>
                        </td>
                        <td width="100" height="8" bgcolor="white">
                            <p align="center"><b><span style="font-size:10pt;"><font face="Arial" color="green">Customers</font></span></b></p>
                        </td>
                        <td width="5" height="8" bgcolor="white">
                            <p align="center"><span style="font-size:12pt;"><font face="Arial" color="black">|</font></span></p>
                        </td>
                        <td width="120" height="8">
                            <p align="center"><a href="newslist.asp"><b><span style="font-size:10pt;"><font face="Arial" color="navy">News 
                            And Events</font></span></b></a></p>
                        </td>
                        <td width="5" height="8">
                            <p align="center"><span style="font-size:12pt;"><font face="Arial">|</font></span></p>
                        </td>
                    </tr>
                </table>
</div>
            
</td>
    </tr>
    <tr>
        <td height="13">
            <div align="left">
                <table cellpadding="0" cellspacing="0" width="695" bgcolor="white">
                    <tr>
                        <td width="20" height="5">
                            <p><font face="Arial"><span style="font-size:10pt;"><b>&nbsp;</b></span></font></p>
                        </td>
                        <td width="5" height="5" bgcolor="white">
                            <p align="center"><font face="Arial"><span style="font-size:10pt;"><b>&nbsp;</b></span></font></p>
                        </td>
                        <td width="100" height="5" bgcolor="#6351B7">
                            <p align="center"><font face="Arial" color="white"><span style="font-size:10pt;"><b>Contact</b></span></font></p>
                        </td>
                        <td width="5" height="5" bgcolor="white">
                            <p align="center"><span style="font-size:10pt;"><font face="Arial"><b>&nbsp;</b></font></span></p>
                        </td>
                        <td width="100" height="5" bgcolor="#979EE0">
                            <p align="center"><a href="makeofferlist.asp"><span style="font-size:10pt;"><font face="Arial" color="white"><b>Offers</b></font></span></a></p>
                        </td>
                        <td width="5" height="5" bgcolor="white">
                            <p align="center"><span style="font-size:10pt;"><font face="Arial" color="white"><b>|</b></font></span></p>
                        </td>
                        <td width="120" height="5" bgcolor="#979EE0">
                            <p align="center"><a href="moreinfolist.asp"><span style="font-size:10pt;"><font face="Arial" color="white"><b>More 
                            Info</b></font></span></a></p>
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
                            <p align="center"><span style="font-size:10pt;"><font face="Arial"><b>&nbsp;</b></font></span></p>
                        </td>
                    </tr>
                </table>
            </div>
</td>
    </tr>
    <tr>
        <td>
            <p align="left"><font face="Arial"><span style="font-size:11pt;"><br><b><i>&nbsp;&nbsp;&nbsp;Contact 
            List:</i></b><br></span></font></p>
            <table align="center" cellpadding="0" cellspacing="0" width="792" bgcolor="white">
                <tr>
                    <td><table border="0" cellspacing="0" cellpadding="0" width="801"><tr><td width="778" height="7">
<form method="post">
<table border="0" cellspacing="0" cellpadding="0" bgcolor="#CCCCCC" width="750" align="center">
<tr bgcolor="#708090">
<td width="70" background="images/sortbar.gif" height="20" bgcolor="white">
                                                    <p align="center"><a href='contactlist.asp?order=<%= Server.URLEncode("contactdate") %>'><font color="black" face="Arial"><b><span style="font-size:8pt;">Date</span></b></font></a><font color="black" face="Arial"><span style="font-size:8pt;"><b>&nbsp;</b></span></font><font color="black"><span style="font-size:8pt;"><b><% If OrderBy = "contactdate" Then %></b></span></font><font color="black" face="Webdings"><span style="font-size:8pt;"><b><% If Session("contact_OT") = "ASC" Then %>5<% ElseIf Session("contact_OT") = "DESC" Then %>6<% End If %></b></span></font><font color="black"><span style="font-size:8pt;"><b><% End If %></b></span></font>
</td>
<td width="213" background="images/sortbar.gif" height="20" bgcolor="white">
<font color="black" face="Arial"><b><span style="font-size:8pt;">First &amp; Last Name</span></b></font>
</td>
<td width="217" background="images/sortbar.gif" height="20" bgcolor="white">
<font color="black" face="Arial"><b><span style="font-size:8pt;">E-mail</span></b></font>
</td>
<td width="149" background="images/sortbar.gif" height="20" bgcolor="white">
                                                    <p align="center"><font color="black" face="Arial"><b><span style="font-size:8pt;">Phone Number</span></b></font>
</td>
<td width="42" background="images/sortbar.gif" height="20" bgcolor="white"><font color="black" face="Arial"><span style="font-size:8pt;"><b>&nbsp;</b></span></font></td>
<td width="28" background="images/sortbar.gif" height="20" bgcolor="white"><font color="black" face="Arial"><span style="font-size:8pt;"><b>&nbsp;</b></span></font></td>
<td width="31" background="images/sortbar.gif" height="20" bgcolor="white"><font color="black" face="Arial"><span style="font-size:8pt;"><b>&nbsp;</b></span></font></td>
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
	x_first_name = rs("first name")
	x_last_name = rs("last name")
	x_email = rs("email")
	x_phone = rs("phone")
	x_comments = rs("comments")
	x_contactdate = rs("contactdate")
%>
<tr bgcolor="<%= bgcolor %>">
<td width="70">
                                                    <p align="right"><font face="Verdana"><span style="font-size:8pt;"><% response.write x_contactdate %> 
                                                    &nbsp;</span></font></td>
<td width="213"><font face="Verdana"><span style="font-size:8pt;"><% response.write x_first_name %>&nbsp;<% response.write x_last_name %>&nbsp;</span></font></td>
<td width="217"><font face="Verdana"><span style="font-size:8pt;"><% response.write x_email %>&nbsp;</span></font></td>
<td width="149">
                                                    <p align="center"><font face="Verdana"><span style="font-size:8pt;"><% response.write x_phone %></span></font></td>
<td width="42">
                                                    <p align="center"><a href="<% key = rs("ID") : If not isnull(key) Then response.write "contactview.asp?key=" & Server.URLEncode(key) Else response.write "javascript:alert('Invalid Record! Key is null');" End If %>"><font face="Arial"><span style="font-size:8pt;"><b>View</b></span></font></a></td>
<td width="28">
                                                    <p align="center"><a href="<% key = rs("ID") : If not isnull(key) Then response.write "contactedit.asp?key=" & Server.URLEncode(key) Else response.write "javascript:alert('Invalid Record! Key is null');" End If %>"><font face="Arial"><span style="font-size:8pt;"><b>Edit</b></span></font></a></td>
<td width="31">
                                                    <p align="right"><input type="checkbox" name="key" value="<%= key %>"></td>
</tr>
<%
	End If
	rs.MoveNext
Loop
%>
</table>
                                        <font face="Arial"><span style="font-size:9pt;"><% If recActual > 0 Then %></span></font>
<p align="right"><font face="Arial"><span style="font-size:9pt;"><input type="button" name="btndelete" value="DELETE SELECTED" onClick="this.form.action='contactdelete.asp';this.form.submit();" style="font-family:Arial; font-size:12;"></span></font></p>
                                        <font face="Arial"><span style="font-size:9pt;"><% End If %></span></font>
</form>
                                    <font face="Arial"><span style="font-size:9pt;"><%
' Close recordSet and connection
rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing %></span></font>

                    </td>
                </tr>
            </table>
<%
If totalRecs > 0 Then
	rsEof = (totalRecs < (startRec + displayRecs))
	PrevStart = startRec - displayRecs
	If PrevStart < 1 Then PrevStart = 1
	NextStart = startRec + displayRecs
	If NextStart > totalRecs Then NextStart = startRec
	LastStart = ((totalRecs-1)\displayRecs)*displayRecs+1
	%>
                                    <p>
                                    <p>
                                    <p>
                                    <p>
                                    <p>
                                    <p>
                                    <p>
                        <p align="center">
                        <p align="center">
                        <p>
                        <p align="center"><font face="Arial"><span style="font-size:9pt;">	<table border="0" cellspacing="0" cellpadding="0"><tr><td><font face="Arial" size="2">Page</font>&nbsp;</td>
<!--first page button-->
	<% If CLng(startRec)=1 Then %>
	<td><img src="images/firstdisab.gif" alt="First" width="20" height="15" border="0"></td>
	<% Else %>
	<td><a href="contactlist.asp?start=1"><img src="images/first.gif" alt="First" width="20" height="15" border="0"></a></td>
	<% End If %>
<!--previous page button-->
	<% If CLng(PrevStart) = CLng(startRec) Then %>
	<td><img src="images/prevdisab.gif" alt="Previous" width="20" height="15" border="0"></td>
	<% Else %>
	<td><a href="contactlist.asp?start=<%=PrevStart%>"><img src="images/prev.gif" alt="Previous" width="20" height="15" border="0"></a></td>
	<% End If %>
<!--current page number-->
	<td><input type="text" name="pageno" value="<%=(startRec-1)\displayRecs+1%>" size="4" style="font-size: 9pt;"></td>
<!--next page button-->
	<% If CLng(NextStart) = CLng(startRec) Then %>
	<td><img src="images/nextdisab.gif" alt="Next" width="20" height="15" border="0"></td>
	<% Else %>
	<td><a href="contactlist.asp?start=<%=NextStart%>"><img src="images/next.gif" alt="Next" width="20" height="15" border="0"></a></td>
	<% End If %>
<!--last page button-->
	<% If CLng(LastStart) = CLng(startRec) Then %>
	<td><img src="images/lastdisab.gif" alt="Last" width="20" height="15" border="0"></td>
	<% Else %>
	<td><a href="contactlist.asp?start=<%=LastStart%>"><img src="images/last.gif" alt="Last" width="20" height="15" border="0"></a></td>
	<% End If %>
	
<form>	
</form>	

</td></tr></table>
                                    </span></font>
                        <p align="center"><font face="Arial"><span style="font-size:9pt;"><% If CLng(startRec) > CLng(totalRecs) Then startRec = totalRecs
	stopRec = startRec + displayRecs - 1
	recCount = totalRecs - 1
	If rsEOF Then recCount = totalRecs
	If stopRec > recCount Then stopRec = recCount %>
	Records <%= startRec %> to <%= stopRec %> of <%= totalRecs %>
<% Else %>
	No records found</span></font></p>
<form>	
<% End If %></form>	
        </td>
    </tr>
</table>
</td>
    </tr>
</table>
<!--#include file="footer.asp"-->