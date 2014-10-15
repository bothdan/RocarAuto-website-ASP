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
Response.Buffer = True
'multiple delete records
key = Request.Form("key")
arRecKey = Split(key&"", ",")
If UBound(arRecKey) = -1 Then Response.Redirect "contorlist.asp"
For Each reckey In arRecKey
	'remove spaces
	reckey = trim(reckey)
	' build the SQL
	sqlKey = sqlKey & "("
	sqlKey = sqlKey & "[ID]=" & "" & reckey & "" & " AND "
	If Right(sqlKey, 5)=" AND " Then sqlKey = Left(sqlKey, Len(sqlKey)-5)
	sqlKey = sqlKey & ") OR "
Next
If Right(sqlKey, 4)=" OR " Then sqlKey = Left(sqlKey, Len(sqlKey)-4)
'get action
a=Request.Form("a")
If a="" or IsNull(a) Then
	a="I"	'display with input box
End If
' Open Connection to the database
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str
Select Case a
	Case "I": ' Display
		strsql = "SELECT * FROM [contor] WHERE " & sqlKey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn
		If rs.EOF Then
			Response.Clear
			Response.Redirect "contorlist.asp"
		Else
			rs.MoveFirst
		End If
	Case "D": ' Delete
		strsql = "SELECT * FROM [contor] WHERE " & sqlKey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn, 1, 2
		Do While NOT rs.EOF
			rs.Delete
			rs.MoveNext
		Loop
		rs.Close
		Set rs = Nothing
		conn.Close
		Set conn = Nothing		
		Response.Clear
		Response.Redirect "contorlist.asp"
End Select
%>
<!--#include file="header.asp"-->
<p><font face="Arial" size="2">Delete from TABLE: contor<br><br><a href="contorlist.asp">Back to List</a></font></p>
<form action="contordelete.asp" method="post">
<p>
<input type="hidden" name="a" value="D">
<table border="0" cellspacing="0" cellpadding="0" bgcolor="#CCCCCC">
<tr bgcolor="#708090">
<td><font color="#FFFFFF"><font face="Arial" size="2">ID&nbsp;</font></font></td>
<td><font color="#FFFFFF"><font face="Arial" size="2">contor&nbsp;</font></font></td>
<td><font color="#FFFFFF"><font face="Arial" size="2">poston&nbsp;</font></font></td>
<td><font color="#FFFFFF"><font face="Arial" size="2">members&nbsp;</font></font></td>
<td><font color="#FFFFFF"><font face="Arial" size="2">onlinem&nbsp;</font></font></td>
</tr>
<%
recCount = 0
Do While NOT rs.EOF
	recCount = recCount + 1
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
	x_ID = rs("ID")
	x_contor = rs("contor")
	x_poston = rs("poston")
	x_members = rs("members")
	x_onlinem = rs("onlinem")
%>
<tr bgcolor="<%= bgcolor %>">
<input type="hidden" name="key" value="<%= key %>">
<td><font face="Arial" size="2">
<%= x_ID %>&nbsp;
</font></td>
<td><font face="Arial" size="2">
<% response.write x_contor %>&nbsp;
</font></td>
<td><font face="Arial" size="2">
<% response.write x_poston %>&nbsp;
</font></td>
<td><font face="Arial" size="2">
<% response.write x_members %>&nbsp;
</font></td>
<td><font face="Arial" size="2">
<% response.write x_onlinem %>&nbsp;
</font></td>
</tr>
<%
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing
%>
</table>
<p>
<input type="submit" name="Action" value="CONFIRM DELETE">
</form>
<!--#include file="footer.asp"-->
