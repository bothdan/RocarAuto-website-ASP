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
		Session("contor_searchwhere") =searchwhere
    ElseIf UCase(cmd) = "RESETALL" Then
		'reset search criteria
		searchwhere = ""
		Session("contor_searchwhere") =searchwhere
	End If
	'reset start record counter (reset command)
	startRec = 1
	Session("contor_REC") = startRec
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
	If Session("contor_OB") = OrderBy Then
		If Session("contor_OT") = "ASC" Then
			Session("contor_OT") = "DESC"
		Else
			Session("contor_OT") = "ASC"
		End if
	Else
		Session("contor_OT") = "ASC"
	End If
	Session("contor_OB") = OrderBy
	Session("contor_REC") = 1
Else
	OrderBy = Session("contor_OB")
	If OrderBy = "" Then
		OrderBy = DefaultOrder
		Session("contor_OB") = OrderBy
		Session("contor_OT") = DefaultOrderType
	End If
End If
' Open Connection to the database
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str
' Build SQL
strsql = "SELECT * FROM [contor]"
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
	strsql = strsql & " ORDER BY [" & OrderBy & "] " & Session("contor_OT")
End If	
'Response.Write strsql
Set rs = Server.CreateObject("ADODB.Recordset")
rs.cursorlocation = 3
rs.Open strsql, conn, 1, 2
totalRecs = rs.RecordCount
' Check for a START parameter
If Request.QueryString("start").Count > 0 Then
	startRec = Request.QueryString("start")
	Session("contor_REC") = startRec
ElseIf Request.QueryString("pageno").Count > 0 Then
	pageno = Request.QueryString("pageno")
	If IsNumeric(pageno) Then
		startRec = (pageno-1)*displayRecs+1
		If startRec <= 0 Then
			startRec = 1
		ElseIf startRec >= ((totalRecs-1)\displayRecs)*displayRecs+1 Then
			startRec = ((totalRecs-1)\displayRecs)*displayRecs+1
		End If
		Session("contor_REC") = startRec
	Else
		startRec = Session("contor_REC")
		If Not IsNumeric(startRec) Or startRec = "" Then
			'reset start record counter
			startRec = 1
			Session("contor_REC") = startRec
		End If
	End If
Else
	startRec = Session("contor_REC")
	If Not IsNumeric(startRec) Or startRec = "" Then
		'reSet start record counter
		startRec = 1
		Session("contor_REC") = startRec
	End If
End If
%>
<!--#include file="header.asp"-->
<p><font face="Arial" size="2">TABLE: contor</font></p>
<table border="0" cellspacing="0" cellpadding="10"><tr><td>
<%
If totalRecs > 0 Then
	rsEof = (totalRecs < (startRec + displayRecs))
	PrevStart = startRec - displayRecs
	If PrevStart < 1 Then PrevStart = 1
	NextStart = startRec + displayRecs
	If NextStart > totalRecs Then NextStart = startRec
	LastStart = ((totalRecs-1)\displayRecs)*displayRecs+1
	%>
<form>	
	<table border="0" cellspacing="0" cellpadding="0"><tr><td><font face="Arial" size="2">Page</font>&nbsp;</td>
<!--first page button-->
	<% If CLng(startRec)=1 Then %>
	<td><img src="images/firstdisab.gif" alt="First" width="20" height="15" border="0"></td>
	<% Else %>
	<td><a href="contorlist.asp?start=1"><img src="images/first.gif" alt="First" width="20" height="15" border="0"></a></td>
	<% End If %>
<!--previous page button-->
	<% If CLng(PrevStart) = CLng(startRec) Then %>
	<td><img src="images/prevdisab.gif" alt="Previous" width="20" height="15" border="0"></td>
	<% Else %>
	<td><a href="contorlist.asp?start=<%=PrevStart%>"><img src="images/prev.gif" alt="Previous" width="20" height="15" border="0"></a></td>
	<% End If %>
<!--current page number-->
	<td><input type="text" name="pageno" value="<%=(startRec-1)\displayRecs+1%>" size="4" style="font-size: 9pt;"></td>
<!--next page button-->
	<% If CLng(NextStart) = CLng(startRec) Then %>
	<td><img src="images/nextdisab.gif" alt="Next" width="20" height="15" border="0"></td>
	<% Else %>
	<td><a href="contorlist.asp?start=<%=NextStart%>"><img src="images/next.gif" alt="Next" width="20" height="15" border="0"></a></td>
	<% End If %>
<!--last page button-->
	<% If CLng(LastStart) = CLng(startRec) Then %>
	<td><img src="images/lastdisab.gif" alt="Last" width="20" height="15" border="0"></td>
	<% Else %>
	<td><a href="contorlist.asp?start=<%=LastStart%>"><img src="images/last.gif" alt="Last" width="20" height="15" border="0"></a></td>
	<% End If %>
	<td><a href="contoradd.asp"><img src="images/addnew.gif" alt="Add new" width="20" height="15" border="0"></a></td>
	<td>&nbsp;<font face="Arial" size="2">of <%=(totalRecs-1)\displayRecs+1%></font></td>
	</td></tr></table>	
</form>	
	<% If CLng(startRec) > CLng(totalRecs) Then startRec = totalRecs
	stopRec = startRec + displayRecs - 1
	recCount = totalRecs - 1
	If rsEOF Then recCount = totalRecs
	If stopRec > recCount Then stopRec = recCount %>
	<font face="Arial" size="2">Records <%= startRec %> to <%= stopRec %> of <%= totalRecs %></font>
<% Else %>
	<font face="Arial" size="2">No records found</font>
<p>
<a href="contoradd.asp"><img src="images/addnew.gif" alt="Add new" width="20" height="15" border="0"></a>
</p>
<% End If %>
</td></tr></table>
<form method="post">
<table border="0" cellspacing="0" cellpadding="0" bgcolor="#CCCCCC">
<tr bgcolor="#708090">
<td>
<a href="contorlist.asp?order=<%= Server.URLEncode("ID") %>"><font color="#FFFFFF"><font face="Arial" size="2">ID&nbsp;</font><% If OrderBy = "ID" Then %><font face="Webdings"><% If Session("contor_OT") = "ASC" Then %>5<% ElseIf Session("contor_OT") = "DESC" Then %>6<% End If %></font><% End If %></font></a>
</td>
<td>
<a href="contorlist.asp?order=<%= Server.URLEncode("contor") %>"><font color="#FFFFFF"><font face="Arial" size="2">contor&nbsp;</font><% If OrderBy = "contor" Then %><font face="Webdings"><% If Session("contor_OT") = "ASC" Then %>5<% ElseIf Session("contor_OT") = "DESC" Then %>6<% End If %></font><% End If %></font></a>
</td>
<td>
<a href="contorlist.asp?order=<%= Server.URLEncode("poston") %>"><font color="#FFFFFF"><font face="Arial" size="2">poston&nbsp;</font><% If OrderBy = "poston" Then %><font face="Webdings"><% If Session("contor_OT") = "ASC" Then %>5<% ElseIf Session("contor_OT") = "DESC" Then %>6<% End If %></font><% End If %></font></a>
</td>
<td>
<a href="contorlist.asp?order=<%= Server.URLEncode("members") %>"><font color="#FFFFFF"><font face="Arial" size="2">members&nbsp;</font><% If OrderBy = "members" Then %><font face="Webdings"><% If Session("contor_OT") = "ASC" Then %>5<% ElseIf Session("contor_OT") = "DESC" Then %>6<% End If %></font><% End If %></font></a>
</td>
<td>
<a href="contorlist.asp?order=<%= Server.URLEncode("onlinem") %>"><font color="#FFFFFF"><font face="Arial" size="2">onlinem&nbsp;</font><% If OrderBy = "onlinem" Then %><font face="Webdings"><% If Session("contor_OT") = "ASC" Then %>5<% ElseIf Session("contor_OT") = "DESC" Then %>6<% End If %></font><% End If %></font></a>
</td>
<td>&nbsp;</td>
<td>&nbsp;</td>
<td>&nbsp;</td>
<td>&nbsp;</td>
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
	x_contor = rs("contor")
	x_poston = rs("poston")
	x_members = rs("members")
	x_onlinem = rs("onlinem")
%>
<tr bgcolor="<%= bgcolor %>">
<td><font face="Arial" size="2"><%= x_ID %>&nbsp;</font></td>
<td><font face="Arial" size="2"><% response.write x_contor %>&nbsp;</font></td>
<td><font face="Arial" size="2"><% response.write x_poston %>&nbsp;</font></td>
<td><font face="Arial" size="2"><% response.write x_members %>&nbsp;</font></td>
<td><font face="Arial" size="2"><% response.write x_onlinem %>&nbsp;</font></td>
<td><a href="<% key = rs("ID") : If not isnull(key) Then response.write "contorview.asp?key=" & Server.URLEncode(key) Else response.write "javascript:alert('Invalid Record! Key is null');" End If %>"><font face="Arial" size="2">View</font></a></td>
<td><a href="<% key = rs("ID") : If not isnull(key) Then response.write "contoredit.asp?key=" & Server.URLEncode(key) Else response.write "javascript:alert('Invalid Record! Key is null');" End If %>"><font face="Arial" size="2">Edit</font></a></td>
<td><a href="<% key = rs("ID") : If not isnull(key) Then response.write "contoradd.asp?key=" & Server.URLEncode(key) Else response.write "javascript:alert('Invalid Record! Key is null');" End If %>"><font face="Arial" size="2">Copy</font></a></td>
<td><input type="checkbox" name="key" value="<%= key %>"><font face="Arial" size="2">Delete</font></td>
</tr>
<%
	End If
	rs.MoveNext
Loop
%>
</table>
<% If recActual > 0 Then %>
<p><input type="button" name="btndelete" value="DELETE SELECTED" onClick="this.form.action='contordelete.asp';this.form.submit();"></p>
<% End If %>
</form>
<%
' Close recordSet and connection
rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing %>
<!--#include file="footer.asp"-->
