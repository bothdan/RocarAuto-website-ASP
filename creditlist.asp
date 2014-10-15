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
		Session("credit_searchwhere") =searchwhere
    ElseIf UCase(cmd) = "RESETALL" Then
		'reset search criteria
		searchwhere = ""
		Session("credit_searchwhere") =searchwhere
	End If
	'reset start record counter (reset command)
	startRec = 1
	Session("credit_REC") = startRec
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
DefaultOrder = "creditappdate"
DefaultOrderType = "DESC"
'No Default Filter
DefaultFilter = ""
' Check for an Order parameter
OrderBy = ""
If Request.QueryString("order").Count > 0 Then
	OrderBy = Request.QueryString("order")
	' Check If an ASC/DESC toggle is required
	If Session("credit_OB") = OrderBy Then
		If Session("credit_OT") = "ASC" Then
			Session("credit_OT") = "DESC"
		Else
			Session("credit_OT") = "ASC"
		End if
	Else
		Session("credit_OT") = "ASC"
	End If
	Session("credit_OB") = OrderBy
	Session("credit_REC") = 1
Else
	OrderBy = Session("credit_OB")
	If OrderBy = "" Then
		OrderBy = DefaultOrder
		Session("credit_OB") = OrderBy
		Session("credit_OT") = DefaultOrderType
	End If
End If
' Open Connection to the database
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str
' Build SQL
strsql = "SELECT * FROM [credit]"
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
	strsql = strsql & " ORDER BY [" & OrderBy & "] " & Session("credit_OT")
End If	
'Response.Write strsql
Set rs = Server.CreateObject("ADODB.Recordset")
rs.cursorlocation = 3
rs.Open strsql, conn, 1, 2
totalRecs = rs.RecordCount
' Check for a START parameter
If Request.QueryString("start").Count > 0 Then
	startRec = Request.QueryString("start")
	Session("credit_REC") = startRec
ElseIf Request.QueryString("pageno").Count > 0 Then
	pageno = Request.QueryString("pageno")
	If IsNumeric(pageno) Then
		startRec = (pageno-1)*displayRecs+1
		If startRec <= 0 Then
			startRec = 1
		ElseIf startRec >= ((totalRecs-1)\displayRecs)*displayRecs+1 Then
			startRec = ((totalRecs-1)\displayRecs)*displayRecs+1
		End If
		Session("credit_REC") = startRec
	Else
		startRec = Session("credit_REC")
		If Not IsNumeric(startRec) Or startRec = "" Then
			'reset start record counter
			startRec = 1
			Session("credit_REC") = startRec
		End If
	End If
Else
	startRec = Session("credit_REC")
	If Not IsNumeric(startRec) Or startRec = "" Then
		'reSet start record counter
		startRec = 1
		Session("credit_REC") = startRec
	End If
End If
%>
<!--#include file="header.asp"-->
<meta name="generator" content="Namo WebEditor v5.0(Trial)">
<body bgcolor="white" text="black" link="green" vlink="black" alink="red">
<table align="center" cellpadding="0" cellspacing="0" width="801" bgcolor="white">
    <tr>
        <td height="11">
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
        <td>

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
                <p align="center"><span style="font-size:10pt;"><a href="makeofferlist.asp"><b><font face="Arial" color="white">Offers</font></b></a></span></p>
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
            <td width="120" height="5" bgcolor="#6351B7">
                <p align="center"><span style="font-size:10pt;"><font face="Arial" color="white"><b>Credit 
                Application</b></font></span></p>
            </td>
            <td width="5" height="5" bgcolor="white">
                <p align="center"><span style="font-size:10pt;"><font face="Arial" color="white"><b>|</b></font></span></p>
            </td>
            <td width="100" height="5" bgcolor="#979EE0">
                <p align="center"><span style="font-size:10pt;"><a href="tradeinlist.asp"><b><font face="Arial" color="white">Trade-In</font></b></a></span></p>
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
        <td height="4">
            <p><font face="Arial"><span style="font-size:2pt;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></font></p>
</td>
    </tr>
    <tr>
        <td height="11">
            <div align="left">
                <table cellpadding="0" cellspacing="0" width="632">
                    <tr>
                        <td width="415">
                            <p><font face="Arial"><span style="font-size:10pt;"><b>&nbsp;</b></span></font></p>
                        </td>
                        <td width="5">
                            <p align="center"><font face="Arial"><span style="font-size:10pt;"><b>&nbsp;</b></span></font></p>
                        </td>
                        <td width="100" bgcolor="#FF9900">
                            <p align="center"><font face="Arial" color="white"><span style="font-size:10pt;"><b>Single</b></span></font></p>
                        </td>
                        <td width="5">
                            <p align="center"><span style="font-size:10pt;"><font face="Arial"><b>&nbsp;</b></font></span></p>
                        </td>
                        <td width="102" bgcolor="#FFCC00">
                            <p align="center"><a href="cocreditlist.asp"><span style="font-size:10pt;"><b><font face="Arial" color="white">Co-Applicant</font></b></span></a></p>
                        </td>
                        <td width="5">
                            <p align="center"><span style="font-size:10pt;"><font face="Arial"><b>&nbsp;</b></font></span></p>
                        </td>
                    </tr>
                </table>
            </div>
</td>
    </tr>
    <tr>
        <td>
            <p> <font face="Arial"><span style="font-size:11pt;"><b><i>&nbsp;&nbsp;&nbsp;Applicant 
            List:<br>&nbsp;</i></b></span></font></p>
</td>
    </tr>
    <tr>
        <td><form method="post">
<table border="0" cellspacing="0" cellpadding="0" bgcolor="#CCCCCC" width="750" align="center">
<tr bgcolor="#708090">
<td width="120" height="14" bgcolor="white" background="images/sortbar.gif">
<a href="creditlist.asp?order=<%= Server.URLEncode("creditappdate") %>"><font face="Arial" color="black"><b><span style="font-size:8pt;">Date</span></b></font></a><font face="Arial" color="black"><b><span style="font-size:8pt;">&nbsp;</span></b></font><b><font color="black"><span style="font-size:8pt;"><% If OrderBy = "creditappdate" Then %></span></font></b><font face="Webdings" color="black"><b><span style="font-size:8pt;"><% If Session("credit_OT") = "ASC" Then %>5<% ElseIf Session("credit_OT") = "DESC" Then %>6<% End If %></span></b></font><b><font color="black"><span style="font-size:8pt;"><% End If %></span></font></b>
</td>
<td width="181" height="14" bgcolor="white" background="images/sortbar.gif">
<a href="creditlist.asp?order=<%= Server.URLEncode("last") %>"><font face="Arial" color="black"><b><span style="font-size:8pt;">First 
                            &amp; Last Name</span></b></font></a><font face="Arial" color="black"><b><span style="font-size:8pt;">&nbsp;</span></b></font><b><font color="black"><span style="font-size:8pt;"><% If OrderBy = "last" Then %></span></font></b><font face="Webdings" color="black"><b><span style="font-size:8pt;"><% If Session("credit_OT") = "ASC" Then %>5<% ElseIf Session("credit_OT") = "DESC" Then %>6<% End If %></span></b></font><b><font color="black"><span style="font-size:8pt;"><% End If %></span></font></b>
</td>
<td width="159" height="14" bgcolor="white" background="images/sortbar.gif">
<font face="Arial" color="black"><b><span style="font-size:8pt;">Home Phone</span></b></font>
</td>
<td width="76" height="14" bgcolor="white" background="images/sortbar.gif">
<a href="creditlist.asp?order=<%= Server.URLEncode("stock") %>"><font face="Arial" color="black"><b><span style="font-size:8pt;">Stock</span></b></font></a><font face="Arial" color="black"><b><span style="font-size:8pt;">&nbsp;</span></b></font><b><font color="black"><span style="font-size:8pt;"><% If OrderBy = "stock" Then %></span></font></b><font face="Webdings" color="black"><b><span style="font-size:8pt;"><% If Session("credit_OT") = "ASC" Then %>5<% ElseIf Session("credit_OT") = "DESC" Then %>6<% End If %></span></b></font><b><font color="black"><span style="font-size:8pt;"><% End If %></span></font></b>
</td>
<td width="48" height="14" bgcolor="white" background="images/sortbar.gif"><b><font color="black"><span style="font-size:8pt;">&nbsp;</span></font></b></td>
<td width="44" height="14" bgcolor="white" background="images/sortbar.gif"><b><font color="black"><span style="font-size:8pt;">&nbsp;</span></font></b></td>
<td width="36" height="14" bgcolor="white" background="images/sortbar.gif"><b><font color="black"><span style="font-size:8pt;">&nbsp;</span></font></b></td>
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
	x_email = rs("email")
	x_first = rs("first")
	x_middle = rs("middle")
	x_last = rs("last")
	x_street = rs("street")
	x_apartment = rs("apartment")
	x_city = rs("city")
	x_state = rs("state")
	x_zip = rs("zip")
	x_home_phone = rs("home phone")
	x_ssn = rs("ssn")
	x_dob = rs("dob")
	x_workplace = rs("workplace")
	x_occupation = rs("occupation")
	x_work_street = rs("work street")
	x_work_city = rs("work city")
	x_work_state = rs("work state")
	x_work_zip = rs("work zip")
	x_work_phone = rs("work phone")
	x_worktime = rs("worktime")
	x_net_salary = rs("net salary")
	x_other_income = rs("other income")
	x_initials = rs("initials")
	x_iagree = rs("iagree")
	x_stock = rs("stock")
	x_creditappdate = rs("creditappdate")
%>
<tr bgcolor="<%= bgcolor %>">
<td width="120"><font face="Verdana"><span style="font-size:8pt;"><% response.write x_creditappdate %></span></font></td>
<td width="181"><font face="Verdana"><span style="font-size:8pt;"><% response.write x_first %>&nbsp;<% response.write x_last %></span></font></td>
<td width="159"><font face="Verdana"><span style="font-size:8pt;"><% response.write x_home_phone %>&nbsp;</span></font></td>
<td width="76"><font face="Verdana"><span style="font-size:8pt;"><% response.write x_stock %>&nbsp;</span></font></td>
<td width="48">
                            <p align="center"><a href="<% key = rs("ID") : If not isnull(key) Then response.write "creditview.asp?key=" & Server.URLEncode(key) Else response.write "javascript:alert('Invalid Record! Key is null');" End If %>"><font face="Arial"><span style="font-size:8pt;"><b>View</b></span></font></a></td>
<td width="44">
                            <p align="center"><a href="<% key = rs("ID") : If not isnull(key) Then response.write "creditedit.asp?key=" & Server.URLEncode(key) Else response.write "javascript:alert('Invalid Record! Key is null');" End If %>"><font face="Arial"><span style="font-size:8pt;"><b>Edit</b></span></font></a></td>
<td width="36"><input type="checkbox" name="key" value="<%= key %>"></td>
</tr>
<%
	End If
	rs.MoveNext
Loop
%>
</table>
<% If recActual > 0 Then %>
<p align="right"><input type="button" name="btndelete" value="DELETE SELECTED" onClick="this.form.action='creditdelete.asp';this.form.submit();" style="font-family:Arial; font-size:12;"> 
                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</p>
<% End If %>
</form>
<%
' Close recordSet and connection
rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing %><table border="0" cellspacing="0" cellpadding="10" align="center" width="265"><tr><td width="245">
                        <font face="Arial"><span style="font-size:8pt;"><%
If totalRecs > 0 Then
	rsEof = (totalRecs < (startRec + displayRecs))
	PrevStart = startRec - displayRecs
	If PrevStart < 1 Then PrevStart = 1
	NextStart = startRec + displayRecs
	If NextStart > totalRecs Then NextStart = startRec
	LastStart = ((totalRecs-1)\displayRecs)*displayRecs+1
	%></span></font>
                        <p>
                        <p>
                        <p>
                        <p>
                        <p>
                        <p>
                        <p>
                        <p>
                        <p>
                        <p align="center">
                        <p>
                        <p align="center"><font face="Arial"><span style="font-size:8pt;"><table border="0" cellspacing="0" cellpadding="0"><tr><td><font face="Arial" size="2">Page</font>&nbsp;</td>
<!--first page button-->
	<% If CLng(startRec)=1 Then %>
	<td><img src="images/firstdisab.gif" alt="First" width="20" height="15" border="0"></td>
	<% Else %>
	<td><a href="creditlist.asp?start=1"><img src="images/first.gif" alt="First" width="20" height="15" border="0"></a></td>
	<% End If %>
<!--previous page button-->
	<% If CLng(PrevStart) = CLng(startRec) Then %>
	<td><img src="images/prevdisab.gif" alt="Previous" width="20" height="15" border="0"></td>
	<% Else %>
	<td><a href="creditlist.asp?start=<%=PrevStart%>"><img src="images/prev.gif" alt="Previous" width="20" height="15" border="0"></a></td>
	<% End If %>
<!--current page number-->
	<td><input type="text" name="pageno" value="<%=(startRec-1)\displayRecs+1%>" size="4" style="font-size: 9pt;"></td>
<!--next page button-->
	<% If CLng(NextStart) = CLng(startRec) Then %>
	<td><img src="images/nextdisab.gif" alt="Next" width="20" height="15" border="0"></td>
	<% Else %>
	<td><a href="creditlist.asp?start=<%=NextStart%>"><img src="images/next.gif" alt="Next" width="20" height="15" border="0"></a></td>
	<% End If %>
<!--last page button-->
	<% If CLng(LastStart) = CLng(startRec) Then %>
	<td><img src="images/lastdisab.gif" alt="Last" width="20" height="15" border="0"></td>
	<% Else %>
	<td><a href="creditlist.asp?start=<%=LastStart%>"><img src="images/last.gif" alt="Last" width="20" height="15" border="0"></a></td>
	<% End If %>
	
	<td>&nbsp;<font face="Arial" size="2">of <%=(totalRecs-1)\displayRecs+1%></font></td>
	</td></tr></table>	
                        </span></font>                        <p align="center"><font face="Arial"><span style="font-size:8pt;"><% If CLng(startRec) > CLng(totalRecs) Then startRec = totalRecs
	stopRec = startRec + displayRecs - 1
	recCount = totalRecs - 1
	If rsEOF Then recCount = totalRecs
	If stopRec > recCount Then stopRec = recCount %>
	Records <%= startRec %> to <%= stopRec %> of <%= totalRecs %>
<% Else %>
	No records found</span></font>


<form>	
                            <font face="Arial"><span style="font-size:8pt;"><% End If %></span></font></form>	

</td></tr></table>
</td>
    </tr>
</table>

<!--#include file="footer.asp"-->
