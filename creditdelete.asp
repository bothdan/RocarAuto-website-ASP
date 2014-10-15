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
If UBound(arRecKey) = -1 Then Response.Redirect "creditlist.asp"
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
		strsql = "SELECT * FROM [credit] WHERE " & sqlKey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn
		If rs.EOF Then
			Response.Clear
			Response.Redirect "creditlist.asp"
		Else
			rs.MoveFirst
		End If
	Case "D": ' Delete
		strsql = "SELECT * FROM [credit] WHERE " & sqlKey
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
		Response.Redirect "creditlist.asp"
End Select
%>
<!--#include file="header.asp"-->
<meta name="generator" content="Namo WebEditor v5.0(Trial)">
<table align="center" cellpadding="0" cellspacing="0" width="801" bgcolor="white">
    <tr>
        <td><form action="creditdelete.asp" method="post">
	<input type="hidden" name="a" value="D">                <br><span style="font-size:10pt;"><b><font face="Arial"><i>&nbsp;&nbsp;&nbsp;Delete 
                Credit:<br>&nbsp;</i></font></b></span>
            

<table border="0" cellspacing="0" cellpadding="0" bgcolor="#CCCCCC" width="750" align="center">
<tr bgcolor="#708090">
<td width="29" bgcolor="white"><font color="black" face="Verdana"><b><span style="font-size:9pt;">ID&nbsp;</span></b></font></td>
<td width="248" bgcolor="white"><font color="black" face="Verdana"><b><span style="font-size:9pt;">First 
                            &amp; Last Name</span></b></font></td>
<td width="168" bgcolor="white"><font color="black" face="Verdana"><b><span style="font-size:9pt;">Home Phone&nbsp;</span></b></font></td>
<td width="59" bgcolor="white"><font color="black" face="Verdana"><b><span style="font-size:9pt;">Stock#</span></b></font></td>
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
%>
<tr bgcolor="<%= bgcolor %>">
<input type="hidden" name="key" value="<%= key %>">
<td width="29">
<font face="Arial" size="2"><%= x_ID %>&nbsp;
</font></td>
<td width="248">
<font face="Arial" size="2"><% response.write x_first %>&nbsp;
<% response.write x_last %></font></td>
<td width="168">
<font face="Arial" size="2"><% response.write x_home_phone %>&nbsp;
</font></td>
<td width="59">
<font face="Arial" size="2"><% response.write x_stock %>&nbsp;
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
<p align="center">
<input type="submit" name="Action" value="CONFIRM DELETE">	
                <p><span style="font-size:14pt;"><b><font face="Arial">&nbsp;&nbsp;&nbsp;<a href="creditlist.asp"><img src="images/back.gif" align="middle" width="16" height="16" border="0"></a> 
</font></b></span><a href="creditlist.asp"><font face="Arial" size="2" color="black"><b>Back to Credit 
List</b></font></a><font face="Arial" size="2" color="black"><b><br></b></font>            </form>
        </td>
    </tr>
</table>
<!--#include file="footer.asp"-->
