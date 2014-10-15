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
If UBound(arRecKey) = -1 Then Response.Redirect "tradeinlist.asp"
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
		strsql = "SELECT * FROM [tradein] WHERE " & sqlKey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn
		If rs.EOF Then
			Response.Clear
			Response.Redirect "tradeinlist.asp"
		Else
			rs.MoveFirst
		End If
	Case "D": ' Delete
		strsql = "SELECT * FROM [tradein] WHERE " & sqlKey
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
		Response.Redirect "tradeinlist.asp"
End Select
%>
<!--#include file="header.asp"-->
<meta name="generator" content="Namo WebEditor v5.0(Trial)">
<table align="center" cellpadding="0" cellspacing="0" width="801" bgcolor="white">
    <tr>
        <td>
            <p><br></p>
<form action="tradeindelete.asp" method="post">
<p>
<input type="hidden" name="a" value="D"><table border="0" cellspacing="0" cellpadding="0" bgcolor="#CCCCCC" width="750" align="center">
<tr bgcolor="#708090">
<td width="45" bgcolor="white"><font color="black" face="Verdana"><b><span style="font-size:9pt;">ID&nbsp;</span></b></font></td>
<td width="73" bgcolor="white">
                                <p align="center"><font color="black" face="Verdana"><b><span style="font-size:9pt;">Year&nbsp;</span></b></font></td>
<td width="190" bgcolor="white"><font color="black" face="Verdana"><b><span style="font-size:9pt;">Make&nbsp;&amp; 
                                Model</span></b></font></td>
<td width="101" bgcolor="white">
                                <p align="center"><font color="black" face="Verdana"><b><span style="font-size:9pt;">Mileage&nbsp;</span></b></font></td>
<td width="64" bgcolor="white">
                                <p align="center"><font color="black" face="Verdana"><b><span style="font-size:9pt;">Stock#</span></b></font></td>
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
	x_home_phone = rs("home phone")
	x_work_phone = rs("work phone")
	x_email = rs("email")
	x_year = rs("year")
	x_make = rs("make")
	x_model = rs("model")
	x_ext_color = rs("ext_color")
	x_vin = rs("vin")
	x_mileage = rs("mileage")
	x_engine = rs("engine")
	x_doors = rs("doors")
	x_transmission = rs("transmission")
	x_drivetrain = rs("drivetrain")
	x_lease_rental = rs("lease_rental")
	x_odometer = rs("odometer")
	x_records = rs("records")
	x_ac = rs("ac")
	x_pw_windows = rs("pw_windows")
	x_pw_locks = rs("pw_locks")
	x_pw_seats = rs("pw_seats")
	x_pw_steering = rs("pw_steering")
	x_cr_ct = rs("cr_ct")
	x_navig = rs("navig")
	x_sunroof = rs("sunroof")
	x_dvd = rs("dvd")
	x_satelit = rs("satelit")
	x_cd_cd_ch = rs("cd_cd_ch")
	x_am_fm = rs("am_fm")
	x_cass = rs("cass")
	x_leather = rs("leather")
	x_alloy = rs("alloy")
	x_spoiler = rs("spoiler")
	x_body = rs("body")
	x_tires = rs("tires")
	x_engine_rate = rs("engine rate")
	x_trans_rate = rs("trans rate")
	x_glass_rate = rs("glass rate")
	x_interior_rate = rs("interior rate")
	x_exhouse_rate = rs("exhouse rate")
	x_lienholders = rs("lienholders")
	x_title = rs("title")
	x_work = rs("work")
	x_new = rs("new")
	x_accidents = rs("accidents")
	x_dameges = rs("dameges")
	x_paint = rs("paint")
	x_salvage = rs("salvage")
	x_comments = rs("comments")
	x_stock = rs("stock")
%>
<tr bgcolor="<%= bgcolor %>">
<input type="hidden" name="key" value="<%= key %>">
<td width="45">
<font face="Arial" size="2"><%= x_ID %>&nbsp;
</font></td>
<td width="73">
                                <p align="center"><font face="Arial" size="2"><%
Select Case x_year
    Case "2010" response.write "2010"
    Case "2009" response.write "2009"
    Case "2008" response.write "2008"
    Case "2007" response.write "2007"
End Select
%>
&nbsp;
</font></td>
<td width="190">
<font face="Arial" size="2"><% response.write x_make %>&nbsp;
<% response.write x_model %>&nbsp;</font></td>
<td width="101">
                                <p align="center"><font face="Arial" size="2"><% if isnumeric(x_mileage) then response.write formatnumber(x_mileage,0,-2,-2,-2) else response.write x_mileage end if %>&nbsp;
</font></td>
<td width="64">
                                <p align="center"><font face="Arial" size="2"><% response.write x_stock %>&nbsp;
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
<p align="left">
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="submit" name="Action" value="CONFIRM DELETE"></form>
            <p><font face="Arial" size="2" color="black"><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</b></font><a href="tradeinlist.asp"><font face="Arial" size="2" color="black"><b><img src="images/leftsm.gif" width="16" height="16" border="0" align="texttop"></b></font></a><font face="Arial" size="2" color="black"><b> 
&nbsp;&nbsp;</b></font><a href="tradeinlist.asp"><font face="Arial" size="2" color="black"><b>Back to Trade-In Appraisal List</b></font></a>&nbsp;<br>&nbsp;</p>
        </td>
    </tr>
</table>
<!--#include file="footer.asp"-->
