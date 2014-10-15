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
If UBound(arRecKey) = -1 Then Response.Redirect "cocreditlist.asp"
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
		strsql = "SELECT * FROM [cocredit] WHERE " & sqlKey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn
		If rs.EOF Then
			Response.Clear
			Response.Redirect "cocreditlist.asp"
		Else
			rs.MoveFirst
		End If
	Case "D": ' Delete
		strsql = "SELECT * FROM [cocredit] WHERE " & sqlKey
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
		Response.Redirect "cocreditlist.asp"
End Select
%>
<!--#include file="header.asp"-->
<meta name="generator" content="Namo WebEditor v5.0(Trial)">
<table align="center" cellpadding="0" cellspacing="0" width="801" bgcolor="white">
    <tr>
        <td>
            <p><br> <font face="Arial"><span style="font-size:10pt;"><b><i>&nbsp;&nbsp;&nbsp;Delete 
            Co-Applicant:</i></b></span></font></p>
<form action="cocreditdelete.asp" method="post">
<p>
<input type="hidden" name="a" value="D">
<table border="0" cellspacing="0" cellpadding="0" bgcolor="#CCCCCC" width="750" align="center">
<tr bgcolor="#708090">
<td bgcolor="white"><font color="black" face="Verdana"><span style="font-size:9pt;"><b>ID&nbsp;</b></span></font></td>
<td bgcolor="white">
<p><font color="black" face="Verdana"><b><span style="font-size:9pt;">First 
                            &amp; Last Name</span></b></font></td>
<td bgcolor="white"><font color="black" face="Verdana"><span style="font-size:9pt;"><b>Home Phone&nbsp;</b></span></font></td>
<td bgcolor="white"> <font color="black" face="Verdana"><span style="font-size:9pt;"><b>Co-</b></span><b><span style="font-size:9pt;">First 
                            &amp; Last Name</span></b><span style="font-size:9pt;"><b>&nbsp;</b></span></font></td>
<td bgcolor="white"><font color="black" face="Verdana"><span style="font-size:9pt;"><b>Co-Home Phone</b></span></font></td>
<td bgcolor="white"><font color="black" face="Verdana"><span style="font-size:9pt;"><b>Stock&nbsp;</b></span></font></td>
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
	x_first_name = rs("first name")
	x_middle = rs("middle")
	x_last_name = rs("last name")
	x_street = rs("street")
	x_aparment = rs("aparment")
	x_city = rs("city")
	x_state = rs("state")
	x_zip = rs("zip")
	x_home_phone = rs("home phone")
	x_work_phone = rs("work phone")
	x_email = rs("email")
	x_ssn = rs("ssn")
	x_dob = rs("dob")
	x_occupation = rs("occupation")
	x_workplace = rs("workplace")
	x_net_salary = rs("net salary")
	x_timework = rs("timework")
	x_first_co = rs("first co")
	x_middle_co = rs("middle co")
	x_last_co = rs("last co")
	x_street_co = rs("street co")
	x_apartment_co = rs("apartment co")
	x_city_co = rs("city co")
	x_state_co = rs("state co")
	x_zip_co = rs("zip co")
	x_home_phone_co = rs("home phone co")
	x_work_phone_co = rs("work phone co")
	x_email_co = rs("email co")
	x_ssn_co = rs("ssn co")
	x_dob_co = rs("dob co")
	x_occupation_co = rs("occupation co")
	x_workplace_co = rs("workplace co")
	x_net_salary_co = rs("net salary co")
	x_timework_co = rs("timework co")
	x_initials = rs("initials")
	x_iagree = rs("iagree")
	x_stock = rs("stock")
%>
<tr bgcolor="<%= bgcolor %>">
<input type="hidden" name="key" value="<%= key %>">
<td>
<font face="Arial" size="2"><%= x_ID %>&nbsp;
</font></td>
<td>
<font face="Arial" size="2"><% response.write x_first_name %>&nbsp;
<% response.write x_last_name %></font></td>
<td>
<font face="Arial" size="2"><% response.write x_home_phone %>&nbsp;
</font></td>
<td>
<font face="Arial" size="2"><% response.write x_first_co %>&nbsp;
<% response.write x_last_co %></font></td>
<td>
<font face="Arial" size="2"><% response.write x_home_phone_co %>&nbsp;
</font></td>
<td>
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
</form>
            <p><font face="Arial" size="2" color="black"><b>&nbsp;&nbsp;&nbsp;</b></font><a href="cocreditlist.asp"><font face="Arial" size="2" color="black"><b><img src="images/back.gif" align="texttop" width="16" height="16" border="0"></b></font></a><font face="Arial" size="2" color="black"><b> 
            &nbsp;</b></font><a href="cocreditlist.asp"><font face="Arial" size="2" color="black"><b>Back to List</b></font></a><font face="Arial" size="2" color="black"><b><br>&nbsp;</b></font></p>
        </td>
    </tr>
</table>
<!--#include file="footer.asp"-->
