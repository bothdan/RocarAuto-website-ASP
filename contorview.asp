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
key = Request.Querystring("key")
If key = "" OR IsNull(key) Then key = Request.Form("key")
If key = "" OR IsNull(key) Then Response.Redirect "contorlist.asp"
'get action
a=Request.Form("a")
If a="" OR IsNull(a) Then
	a="I"	'display with input box
End If
' Open Connection to the database
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str
Select Case a
	Case "I": ' Get a record to display
		tkey = "" & key & ""
		strsql = "SELECT * FROM [contor] WHERE [ID]=" & tkey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn
		If rs.EOF Then
			Response.Clear
			Response.Redirect "contorlist.asp"
		Else
			rs.MoveFirst
		End If
		' Get the field contents
		x_ID = rs("ID")
		x_contor = rs("contor")
		x_poston = rs("poston")
		x_members = rs("members")
		x_onlinem = rs("onlinem")
		rs.Close
		Set rs = Nothing
End Select
%>
<!--#include file="header.asp"-->
<p><font face="Arial" size="2">View TABLE: contor<br><br><a href="contorlist.asp">Back to List</a></font></p>
<p>
<form>
<table border="0" cellspacing="0" cellpadding="0" bgcolor="#CCCCCC">
<tr>
<td bgcolor="#708090"><font color="#FFFFFF"><font face="Arial" size="2">ID</font>&nbsp;</font></td>
<td bgcolor="#F5F5F5"><font face="Arial" size="2"><%= x_ID %></font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#708090"><font color="#FFFFFF"><font face="Arial" size="2">contor</font>&nbsp;</font></td>
<td bgcolor="#F5F5F5"><font face="Arial" size="2"><% response.write x_contor %></font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#708090"><font color="#FFFFFF"><font face="Arial" size="2">poston</font>&nbsp;</font></td>
<td bgcolor="#F5F5F5"><font face="Arial" size="2"><% response.write x_poston %></font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#708090"><font color="#FFFFFF"><font face="Arial" size="2">members</font>&nbsp;</font></td>
<td bgcolor="#F5F5F5"><font face="Arial" size="2"><% response.write x_members %></font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#708090"><font color="#FFFFFF"><font face="Arial" size="2">onlinem</font>&nbsp;</font></td>
<td bgcolor="#F5F5F5"><font face="Arial" size="2"><% response.write x_onlinem %></font>&nbsp;</td>
</tr>
</table>
</form>
<p>
<!--#include file="footer.asp"-->
