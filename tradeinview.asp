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
If key = "" OR IsNull(key) Then Response.Redirect "tradeinlist.asp"
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
		strsql = "SELECT * FROM [tradein] WHERE [ID]=" & tkey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn
		If rs.EOF Then
			Response.Clear
			Response.Redirect "tradeinlist.asp"
		Else
			rs.MoveFirst
		End If
		' Get the field contents
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
		rs.Close
		Set rs = Nothing
End Select
%>
<!--#include file="header.asp"-->
<meta name="generator" content="Namo WebEditor v5.0(Trial)">
<table align="center" cellpadding="0" cellspacing="0" width="802" bgcolor="white">
    <tr>
        <td>
            <p><font face="Arial" size="2"><br></font><font face="Arial" size="2" color="black"><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</b></font><a href="tradeinlist.asp"><font face="Arial" size="2" color="black"><b><img src="images/leftsm.gif" width="16" height="16" border="0" align="texttop"></b></font></a><font face="Arial" size="2" color="black"><b> 
            &nbsp;&nbsp;</b></font><a href="tradeinlist.asp"><font face="Arial" size="2" color="black"><b>Back to Trade-In Appraisal 
            List</b></font></a></p>
<form onSubmit="return EW_checkMyForm(this);"  action="tradeinadd.asp" method="post">

<table border="0" cellspacing="0" cellpadding="0" bgcolor="#CCCCCC" align="center" width="752">
<tr>
<td bgcolor="white" width="283" colspan="2"><font face="Arial"><b><span style="font-size:14pt;">Trade-In Appraisal</span></b></font></td>
<td bgcolor="white" width="312" colspan="2"><font face="Arial" size="2" color="white"><%= x_ID %><input type="hidden" name="x_ID" value="<%= x_ID %>"></font><font color="white">&nbsp;&nbsp;&nbsp;&nbsp;</font><font color="black" face="Arial"><b><span style="font-size:14pt;">For 
                            Stock # </span></b></font><font face="Arial" color="black"><b><span style="font-size:14pt;"><%
response.write x_stock%></span></b></font></td>
<td bgcolor="white" width="157" colspan="2">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td bgcolor="white" width="283" colspan="2"><DIV class="smalltext required">&nbsp;</DIV></td>
<td bgcolor="white" width="131">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="181">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="130">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="27">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td bgcolor="white" width="116"><font face="Arial" size="2" color="black">First 
                            Name:</font></td>
<td bgcolor="white" width="167"><font face="Arial" size="2" color="green"><%
response.write x_first%></font></td>
<td bgcolor="white" width="131"><font face="Arial" size="2">Air Conditioning:</font></td>
<td bgcolor="white" width="181"><font face="Arial" size="2" color="green"><%
response.write x_ac
%></font></td>
<td bgcolor="white" width="130"><font face="Arial" size="2">Rear DVD:</font></td>
<td bgcolor="white" width="27"><font face="Arial" size="2" color="green"><%
response.write x_dvd
%></font></td>
</tr>
<tr>
<td bgcolor="white" width="116"><font face="Arial" size="2" color="black">Last 
                            Name:</font></td>
<td bgcolor="white" width="167"><font face="Arial" size="2" color="green"><%response.write x_last%></font></td>
<td bgcolor="white" width="131"><font face="Arial" size="2">Power Windows:</font></td>
<td bgcolor="white" width="181"><font face="Arial" size="2" color="green"><%
response.write x_pw_windows
%></font></td>
<td bgcolor="white" width="130"><font face="Arial" size="2">Satellite Radio:</font></td>
<td bgcolor="white" width="27"><font face="Arial" size="2" color="green"><%
response.write x_satelit
%></font></td>
</tr>
<tr>
<td bgcolor="white" width="116"><font face="Arial" size="2" color="black">Home Phone:</font></td>
<td bgcolor="white" width="167"><font face="Arial" size="2" color="green"><%response.write x_home_phone%></font></td>
<td bgcolor="white" width="131"><font face="Arial" size="2">Power Locks:</font></td>
<td bgcolor="white" width="181"><font face="Arial" size="2" color="green"><%
response.write x_pw_locks
%></font></td>
<td bgcolor="white" width="130"><font face="Arial" size="2">CD Player / Changer:</font></td>
<td bgcolor="white" width="27"><font face="Arial" size="2" color="green"><%
response.write x_cd_cd_ch
%></font></td>
</tr>
<tr>
<td bgcolor="white" width="116"><font face="Arial" size="2" color="black">Work Phone:</font></td>
<td bgcolor="white" width="167"><font face="Arial" size="2" color="green"><%response.write x_work_phone%></font></td>
<td bgcolor="white" width="131"><font face="Arial" size="2">Power Seats:</font></td>
<td bgcolor="white" width="181"><font face="Arial" size="2" color="green"><%
response.write x_pw_seats
%></font></td>
<td bgcolor="white" width="130"><font face="Arial" size="2">AM/FM Stereo:</font></td>
<td bgcolor="white" width="27"><font face="Arial" size="2" color="green"><%
response.write x_am_fm
%></font></td>
</tr>
<tr>
<td bgcolor="white" width="116"><font face="Arial" size="2" color="black">E-mail 
                            Address:</font></td>
<td bgcolor="white" width="167"><font face="Arial" size="2" color="green"><%response.write x_email%></font></td>
<td bgcolor="white" width="131"><font face="Arial" size="2">Power Steering:</font></td>
<td bgcolor="white" width="181"><font face="Arial" size="2" color="green"><%
response.write x_pw_steering
%></font></td>
<td bgcolor="white" width="130"><font face="Arial" size="2">Cassette:</font></td>
<td bgcolor="white" width="27"><font face="Arial" size="2" color="green"><%
response.write x_cass
%></font></td>
</tr>
<tr>
<td bgcolor="white" width="116" height="23">
                            <p><font face="Arial" size="2" color="black">&nbsp;</font></p>
</td>
<td bgcolor="white" width="167" height="23">
                            <p><font face="Arial" size="2">&nbsp;</font></p>
</td>
<td bgcolor="white" width="131" height="23"><font face="Arial" size="2">Cruise Control:</font></td>
<td bgcolor="white" width="181" height="23"><font face="Arial" size="2" color="green"><%
response.write x_cr_ct
%></font></td>
<td bgcolor="white" width="130" height="23"><font face="Arial" size="2">Leather Interior:</font></td>
<td bgcolor="white" width="27" height="23"><font face="Arial" size="2" color="green"><%
response.write x_leather
%></font></td>
</tr>
<tr>
<td bgcolor="white" width="116" height="23">
                            <p><font face="Arial" size="2" color="black">&nbsp;</font></p>
</td>
<td bgcolor="white" width="167" height="23">
                            <p><font face="Arial" size="2">&nbsp;</font></p>
</td>
<td bgcolor="white" width="131" height="23"><font face="Arial" size="2">Navigation System:</font></td>
<td bgcolor="white" width="181" height="23"><font face="Arial" size="2" color="green"><%
response.write x_navig
%></font></td>
<td bgcolor="white" width="130" height="23"><font face="Arial" size="2">Alloy Wheels:</font></td>
<td bgcolor="white" width="27" height="23"><font face="Arial" size="2" color="green"><%
response.write x_alloy
%></font></td>
</tr>
<tr>
<td bgcolor="white" width="116" height="23">
                            <p><font face="Arial" size="2" color="black">&nbsp;</font></p>
</td>
<td bgcolor="white" width="167" height="23">
                            <p><font face="Arial" size="2">&nbsp;</font></p>
</td>
<td bgcolor="white" width="131" height="23"><font face="Arial" size="2">Sunroof:</font></td>
<td bgcolor="white" width="181" height="23"><font face="Arial" size="2" color="green"><%
response.write x_sunroof
%></font></td>
<td bgcolor="white" width="130" height="23"><font face="Arial" size="2">Spoiler:</font></td>
<td bgcolor="white" width="27" height="23"><font face="Arial" size="2" color="green"><%
response.write x_spoiler
%></font></td>
</tr>
<tr>
<td bgcolor="white" width="752" colspan="6">
                            <p><font face="Arial" size="2">&nbsp;</font></p>
</td>
</tr>
<tr>
<td bgcolor="white" width="116"><font face="Arial" size="2" color="black">Year&nbsp;</font></td>
<td bgcolor="white" width="167">
<p><font face="Arial" size="2" color="green"><%response.write x_year
%>
</font></td>
<td bgcolor="white" width="469" colspan="4">
                             <font face="Arial"><span style="font-size:11pt;">Rate Of 
                             Vehicle On A Scale Of 1 To 10 (10 is Perfect):</span></font></td>
</tr>
<tr>
<td bgcolor="white" width="116"><font face="Arial" size="2" color="black">Make&nbsp;</font></td>
<td bgcolor="white" width="167"><font face="Arial" size="2" color="green"><%response.write x_make%></font></td>
<td bgcolor="white" width="312" colspan="2"><font face="Arial" size="2">&nbsp;</font></td>
<td bgcolor="white" width="157" colspan="2">
                            <p><font face="Arial" size="2">&nbsp;</font></p>
</td>
</tr>
<tr>
<td bgcolor="white" width="116"><font face="Arial" size="2" color="black">Model&nbsp;</font></td>
<td bgcolor="white" width="167"><font face="Arial" size="2" color="green"><%response.write x_model%></font></td>
<td bgcolor="white" width="312" colspan="2"><font face="Arial" size="2">Body (dents, dings, rust, rot, damage):</font></td>
<td bgcolor="white" width="157" colspan="2"><font face="Arial" size="2" color="green"><%
response.write x_body
%>
</font></td>
</tr>
<tr>
<td bgcolor="white" width="116"><font face="Arial" size="2" color="black">Ext color&nbsp;</font></td>
<td bgcolor="white" width="167"><font face="Arial" size="2" color="green"><%response.write x_ext_color%></font></td>
<td bgcolor="white" width="312" colspan="2"><font face="Arial" size="2">Tires (tread wear, mismatched):</font></td>
<td bgcolor="white" width="157" colspan="2"><font face="Arial" size="2" color="green"><%
response.write x_tires
%>
</font></td>
</tr>
<tr>
<td bgcolor="white" width="116"><font face="Arial" size="2" color="black">VIN</font></td>
<td bgcolor="white" width="167"><font face="Arial" size="2" color="green"><%response.write x_vin%></font></td>
<td bgcolor="white" width="312" colspan="2"><font face="Arial" size="2">Engine (running condition, burns oil, knocking):</font></td>
<td bgcolor="white" width="157" colspan="2"><font face="Arial" size="2" color="green"><%
response.write x_engine_rate
%>
</font></td>
</tr>
<tr>
<td bgcolor="white" width="116"><font face="Arial" size="2" color="black">Mileage&nbsp;</font></td>
<td bgcolor="white" width="167"><font face="Arial" size="2" color="green"><%response.write x_mileage%></font></td>
<td bgcolor="white" width="312" colspan="2"><font face="Arial" size="2">Transmission / Clutch (slipping, hard shift, grinds):</font></td>
<td bgcolor="white" width="157" colspan="2"><font face="Arial" size="2" color="green"><%
response.write x_trans_rate
%></font></td>
</tr>
<tr>
<td bgcolor="white" width="116"><font face="Arial" size="2">Engine&nbsp;</font></td>
<td bgcolor="white" width="167"><font face="Arial" size="2" color="green"><%response.write x_engine%></font></td>
<td bgcolor="white" width="312" colspan="2"><font face="Arial" size="2">Glass (chips, scratches, cracks, pitted):</font></td>
<td bgcolor="white" width="157" colspan="2"><font face="Arial" size="2" color="green"><%
response.write x_glass_rate
%>
</font></td>
</tr>
<tr>
<td bgcolor="white" width="116"><font face="Arial" size="2">Doors&nbsp;</font></td>
<td bgcolor="white" width="167"><font face="Arial" size="2" color="green"><%response.write x_doors
%>
</font></td>
<td bgcolor="white" width="312" colspan="2">
<font face="Arial" size="2">Interior (rips, tears, burns, faded/worn, stains):</font></td>
<td bgcolor="white" width="157" colspan="2"><font face="Arial" size="2" color="green"><%
response.write x_interior_rate
%></font></td>
</tr>
<tr>
<td bgcolor="white" width="116"><font face="Arial" size="2">Transmission&nbsp;</font></td>
<td bgcolor="white" width="167"><font face="Arial" size="2" color="green"><%response.write x_transmission
%>
</font></td>
<td bgcolor="white" width="312" colspan="2"><font face="Arial" size="2">Exhaust (rusted, leaking, noisy):</font></td>
<td bgcolor="white" width="157" colspan="2"><font face="Arial" size="2" color="green"><%
response.write x_exhouse_rate
%></font></td>
</tr>
<tr>
<td bgcolor="white" width="116"><font face="Arial" size="2">Drivetrain&nbsp;</font></td>
<td bgcolor="white" width="167"><font face="Arial" size="2" color="green"><%
response.write x_drivetrain
%>
</font></td>
<td bgcolor="white" width="312" colspan="2">
<font face="Arial" size="2">&nbsp;</font></td>
<td bgcolor="white" width="157" colspan="2">
                            <p><font face="Arial" size="2">&nbsp;</font></p>
</td>
</tr>
<tr>
<td bgcolor="white" width="752" colspan="6">
                            <p><font face="Arial" size="2">&nbsp;</font></p>
</td>
</tr>
<tr>
<td bgcolor="white" width="283" colspan="2"><font face="Arial" size="2">Was it ever a lease or rental return?</font></td>
<td bgcolor="white" width="131"><font face="Arial" size="2" color="green"><%
response.write x_lease_rental
%>
</font><font color="green">&nbsp;</font></td>
<td bgcolor="white" width="181"><font face="Arial" size="2">Are there any lienholders and where are they located?</font></td>
<td bgcolor="white" width="157" colspan="2"><font face="Arial" size="2"><input type="text" name="x_lienholders" size="20" maxlength=50 value="<%= Server.HtmlEncode(x_lienholders&"") %>"></font></td>
</tr>
<tr>
<td bgcolor="white" width="283" colspan="2"><font face="Arial" size="2">Is the odometer operational and accurate?</font></td>
<td bgcolor="white" width="131"><font face="Arial" size="2" color="green"><%
response.write x_odometer
%>
</font><font color="green">&nbsp;</font></td>
<td bgcolor="white" width="181"><font face="Arial" size="2">Who holds this title?</font></td>
<td bgcolor="white" width="157" colspan="2"><font face="Arial" size="2"><input type="text" name="x_title" size="20" maxlength=50 value="<%= Server.HtmlEncode(x_title&"") %>"></font></td>
</tr>
<tr>
<td bgcolor="white" width="283" colspan="2"><font face="Arial" size="2">Detailed service records available?</font></td>
<td bgcolor="white" width="131"><font face="Arial" size="2" color="green"><%
response.write x_records
%>
</font><font color="green">&nbsp;</font></td>
<td bgcolor="white" width="181">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="157" colspan="2">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td bgcolor="white" width="752" colspan="6">
&nbsp;</td>
</tr>
<tr>
<td bgcolor="white" width="752" colspan="6">&nbsp;</td>
</tr>
<tr>
<td bgcolor="white" width="752" colspan="6">
&nbsp;
                            <table align="center" border="0" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC">
                                <tr>
                                    <td width="376" bgcolor="white"><font face="Arial" size="2">Do all options and accessories work correctly?</font></td>
                                    <td width="376" bgcolor="white"><font face="Arial" size="2">Did you buy the vehicle new?</font></td>
                                </tr>
                                <tr>
                                    <td width="376" bgcolor="white"><font face="Arial" size="2"><textarea cols=35 rows=4 name="x_work"><%= x_work %></textarea></font></td>
                                    <td width="376" bgcolor="white"><font face="Arial" size="2"><textarea cols=35 rows=4 name="x_new"><%= x_new %></textarea></font></td>
                                </tr>
                                <tr>
                                    <td width="376" bgcolor="white"><font face="Arial" size="2">Has the vehicle ever been in any accidents? <br>Cost of repairs?</font></td>
                                    <td width="376" bgcolor="white"><font face="Arial" size="2">Is there existing damage on the vehicle? Where?</font></td>
                                </tr>
                                <tr>
                                    <td width="376" bgcolor="white"><font face="Arial" size="2"><textarea cols=35 rows=4 name="x_accidents"><%= x_accidents %></textarea></font></td>
                                    <td width="376" bgcolor="white"><font face="Arial" size="2"><textarea cols=35 rows=4 name="x_dameges"><%= x_dameges %></textarea></font></td>
                                </tr>
                                <tr>
                                    <td width="376" bgcolor="white"><font face="Arial" size="2">Has the vehicle ever had paint work performed?</font></td>
                                    <td width="376" bgcolor="white"><font face="Arial" size="2">Is the title designated &quot;Salvage&quot;or &quot;Reconstructed&quot;?<br> Any other title 
declarations?</font></td>
                                </tr>
                                <tr>
                                    <td width="376" bgcolor="white"><font face="Arial" size="2"><textarea cols=35 rows=4 name="x_paint"><%= x_paint %></textarea></font></td>
                                    <td width="376" bgcolor="white"><font face="Arial" size="2"><textarea cols=35 rows=4 name="x_salvage"><%= x_salvage %></textarea></font></td>
                                </tr>
                                <tr>
                                    <td width="752" bgcolor="white" colspan="2"><font face="Arial" size="2">Comments&nbsp;</font></td>
                                </tr>
                                <tr>
                                    <td width="752" bgcolor="white" colspan="2"><font face="Arial" size="2"><textarea cols="65" rows="6" name="x_comments"><%= x_comments %></textarea></font></td>
                                </tr>
                            </table>
</td>
</tr>
<tr>
<td bgcolor="white" width="752" colspan="6">



<font face="Arial" size="2" color="white"><%= x_stock %><input type="hidden" name="x_stock" value="<%= x_stock %>"></font></td>
</tr>
</table>

</form>
            <p><font face="Arial" size="2" color="black"><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</b></font><a href="tradeinlist.asp"><font face="Arial" size="2" color="black"><b><img src="images/leftsm.gif" width="16" height="16" border="0" align="texttop"></b></font></a><font face="Arial" size="2" color="black"><b> 
            &nbsp;&nbsp;</b></font><a href="tradeinlist.asp"><font face="Arial" size="2" color="black"><b>Back to Trade-In Appraisal 
            List</b></font></a><font face="Arial" size="2" color="black"><b><br>&nbsp;</b></font></p>
        </td>
    </tr>
</table>
<!--#include file="footer.asp"-->
