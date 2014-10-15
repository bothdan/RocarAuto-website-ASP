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
If UBound(arRecKey) = -1 Then Response.Redirect "moreinfolist.asp"
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
		strsql = "SELECT * FROM [moreinfo] WHERE " & sqlKey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn
		If rs.EOF Then
			Response.Clear
			Response.Redirect "moreinfolist.asp"
		Else
			rs.MoveFirst
		End If
	Case "D": ' Delete
		strsql = "SELECT * FROM [moreinfo] WHERE " & sqlKey
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
		Response.Redirect "moreinfolist.asp"
End Select
%>
<!--#include file="header.asp"-->
<meta name="generator" content="Namo WebEditor v5.0(Trial)">
<table align="center" cellpadding="0" cellspacing="0" width="801" bgcolor="white">
    <tr>
        <td width="819">
<form action="moreinfodelete.asp" method="post">
<p>
<input type="hidden" name="a" value="D"><br><font face="Arial"><span style="font-size:10pt;"><b><i>&nbsp;&nbsp;&nbsp;Delete 
                &nbsp;More Info:<br>&nbsp;</i></b></span></font>
<table cellspacing="0" cellpadding="0" bgcolor="#CCCCCC" align="center" width="750">
<tr bgcolor="#708090">
<td width="41" bgcolor="white"><font face="Verdana" color="black"><b><span style="font-size:9pt;">ID&nbsp;</span></b></font></td>
<td width="210" bgcolor="white"><font face="Verdana" color="black"><b><span style="font-size:9pt;">First 
                            &amp; Last Name</span></b></font></td>
<td width="122" bgcolor="white"><font face="Verdana" color="black"><b><span style="font-size:9pt;">Phone&nbsp;</span></b></font></td>
<td width="222" bgcolor="white"><font face="Verdana" color="black"><b><span style="font-size:9pt;">E-mail&nbsp;</span></b></font></td>
<td width="110" bgcolor="white">
                            <p align="center"><font face="Verdana" color="black"><b><span style="font-size:9pt;">Test 
                            Drive&nbsp;</span></b></font></td>
<td width="65" bgcolor="white">
                            <p align="center"><font face="Verdana" color="black"><b><span style="font-size:9pt;">Stock#</span></b></font></td>
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
	x_first = rs("first")
	x_last = rs("last")
	x_phone = rs("phone")
	x_email = rs("email")
	x_testdrive = rs("testdrive")
	x_comments = rs("comments")
	x_stock = rs("stock")
%>
<tr bgcolor="<%= bgcolor %>">
<input type="hidden" name="key" value="<%= key %>">
<td width="41">
<font face="Arial" size="2"><%= x_ID %>&nbsp;
</font></td>
<td width="210">
<font face="Arial" size="2"><% response.write x_first %>&nbsp;
<% response.write x_last %></font></td>
<td width="122">
<font face="Arial" size="2"><% response.write x_phone %>&nbsp;
</font></td>
<td width="222">
<font face="Arial" size="2"><% response.write x_email %>&nbsp;
</font></td>
<td width="110">
                            <p align="center"><font face="Arial" size="2"><%
Select Case x_testdrive
    Case "Yes" response.write "Yes"
    Case "No" response.write ""
End Select
%></font></td>
<td width="65">
                            <p align="center"><font face="Arial" size="2"><% response.write x_stock %></font></td>
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
<p align="center">
<input type="submit" name="Action" value="CONFIRM DELETE">
</form>
            <p><font face="Arial" size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font><a href="moreinfolist.asp"><font face="Arial" size="2" color="black"><b><img src="images/leftsm.gif" align="texttop" width="16" height="16" border="0"></b></font></a><font face="Arial" size="2" color="black"><b> 
            &nbsp;&nbsp;</b></font><a href="moreinfolist.asp"><font face="Arial" size="2" color="black"><b>Back to List</b></font></a><font face="Arial" size="2" color="black"><b><br>&nbsp;</b></font></p>
        </td>
    </tr>
</table>
<!--#include file="footer.asp"-->
