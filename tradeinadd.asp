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
'get action
x_stock = Request.QueryString("st")
a = Request.Form("a")
If (a = "" OR IsNull(a)) Then
	key = Request.Querystring("key")
	If key <> "" Then
		a = "C" 'copy record
	Else
		a = "I" 'display blank record
	End If
End If
' Open Connection to the database
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str
Select Case a
	Case "C": ' Get a record to display
		tkey = "" & key & ""
		strsql = "SELECT * FROM [tradein] WHERE [ID]=" & tkey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn
		If rs.EOF Then
			Response.Clear
			Response.Redirect "tradeinlist.asp"
		Else
			rs.MoveFirst
		' Get the field contents
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
		'x_stock = rs()
		End If
		rs.Close
		Set rs = Nothing
	Case "A": ' Add
		'get fields from form
x_ID = Request.Form("x_ID")
x_first = Request.Form("x_first")
x_last = Request.Form("x_last")
x_home_phone = Request.Form("x_home_phone")
x_work_phone = Request.Form("x_work_phone")
x_email = Request.Form("x_email")
x_year = Request.Form("x_year")
x_make = Request.Form("x_make")
x_model = Request.Form("x_model")
x_ext_color = Request.Form("x_ext_color")
x_vin = Request.Form("x_vin")
x_mileage = Request.Form("x_mileage")
x_engine = Request.Form("x_engine")
x_doors = Request.Form("x_doors")
x_transmission = Request.Form("x_transmission")
x_drivetrain = Request.Form("x_drivetrain")
x_lease_rental = Request.Form("x_lease_rental")
x_odometer = Request.Form("x_odometer")
x_records = Request.Form("x_records")
x_ac = Request.Form("x_ac")
x_pw_windows = Request.Form("x_pw_windows")
x_pw_locks = Request.Form("x_pw_locks")
x_pw_seats = Request.Form("x_pw_seats")
x_pw_steering = Request.Form("x_pw_steering")
x_cr_ct = Request.Form("x_cr_ct")
x_navig = Request.Form("x_navig")
x_sunroof = Request.Form("x_sunroof")
x_dvd = Request.Form("x_dvd")
x_satelit = Request.Form("x_satelit")
x_cd_cd_ch = Request.Form("x_cd_cd_ch")
x_am_fm = Request.Form("x_am_fm")
x_cass = Request.Form("x_cass")
x_leather = Request.Form("x_leather")
x_alloy = Request.Form("x_alloy")
x_spoiler = Request.Form("x_spoiler")
x_body = Request.Form("x_body")
x_tires = Request.Form("x_tires")
x_engine_rate = Request.Form("x_engine_rate")
x_trans_rate = Request.Form("x_trans_rate")
x_glass_rate = Request.Form("x_glass_rate")
x_interior_rate = Request.Form("x_interior_rate")
x_exhouse_rate = Request.Form("x_exhouse_rate")
x_lienholders = Request.Form("x_lienholders")
x_title = Request.Form("x_title")
x_work = Request.Form("x_work")
x_new = Request.Form("x_new")
x_accidents = Request.Form("x_accidents")
x_dameges = Request.Form("x_dameges")
x_paint = Request.Form("x_paint")
x_salvage = Request.Form("x_salvage")
x_comments = Request.Form("x_comments")
'x_stock = st
x_stock = Request.Form("x_stock")
		' Open record
		strsql = "SELECT * FROM [tradein] WHERE 0 = 1"
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn, 1, 2
		rs.AddNew
		tmpFld = Trim(x_first)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("first") = tmpFld
		tmpFld = Trim(x_last)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("last") = tmpFld
		tmpFld = Trim(x_home_phone)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("home phone") = tmpFld
		tmpFld = Trim(x_work_phone)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("work phone") = tmpFld
		tmpFld = Trim(x_email)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("email") = tmpFld
		tmpFld = Trim(x_year)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("year") = tmpFld
		tmpFld = Trim(x_make)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("make") = tmpFld
		tmpFld = Trim(x_model)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("model") = tmpFld
		tmpFld = Trim(x_ext_color)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("ext_color") = tmpFld
		tmpFld = Trim(x_vin)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("vin") = tmpFld
		tmpFld = Trim(x_mileage)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("mileage") = tmpFld
		tmpFld = Trim(x_engine)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("engine") = tmpFld
		tmpFld = Trim(x_doors)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("doors") = tmpFld
		tmpFld = Trim(x_transmission)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("transmission") = tmpFld
		tmpFld = Trim(x_drivetrain)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("drivetrain") = tmpFld
		tmpFld = Trim(x_lease_rental)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("lease_rental") = tmpFld
		tmpFld = Trim(x_odometer)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("odometer") = tmpFld
		tmpFld = Trim(x_records)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("records") = tmpFld
		tmpFld = Trim(x_ac)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("ac") = tmpFld
		tmpFld = Trim(x_pw_windows)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("pw_windows") = tmpFld
		tmpFld = Trim(x_pw_locks)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("pw_locks") = tmpFld
		tmpFld = Trim(x_pw_seats)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("pw_seats") = tmpFld
		tmpFld = Trim(x_pw_steering)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("pw_steering") = tmpFld
		tmpFld = Trim(x_cr_ct)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("cr_ct") = tmpFld
		tmpFld = Trim(x_navig)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("navig") = tmpFld
		tmpFld = Trim(x_sunroof)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("sunroof") = tmpFld
		tmpFld = Trim(x_dvd)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("dvd") = tmpFld
		tmpFld = Trim(x_satelit)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("satelit") = tmpFld
		tmpFld = Trim(x_cd_cd_ch)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("cd_cd_ch") = tmpFld
		tmpFld = Trim(x_am_fm)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("am_fm") = tmpFld
		tmpFld = Trim(x_cass)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("cass") = tmpFld
		tmpFld = Trim(x_leather)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("leather") = tmpFld
		tmpFld = Trim(x_alloy)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("alloy") = tmpFld
		tmpFld = Trim(x_spoiler)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("spoiler") = tmpFld
		tmpFld = Trim(x_body)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("body") = tmpFld
		tmpFld = Trim(x_tires)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("tires") = tmpFld
		tmpFld = Trim(x_engine_rate)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("engine rate") = tmpFld
		tmpFld = Trim(x_trans_rate)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("trans rate") = tmpFld
		tmpFld = Trim(x_glass_rate)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("glass rate") = tmpFld
		tmpFld = Trim(x_interior_rate)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("interior rate") = tmpFld
		tmpFld = Trim(x_exhouse_rate)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("exhouse rate") = tmpFld
		tmpFld = Trim(x_lienholders)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("lienholders") = tmpFld
		tmpFld = Trim(x_title)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("title") = tmpFld
		tmpFld = Trim(x_work)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("work") = tmpFld
		tmpFld = Trim(x_new)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("new") = tmpFld
		tmpFld = Trim(x_accidents)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("accidents") = tmpFld
		tmpFld = Trim(x_dameges)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("dameges") = tmpFld
		tmpFld = Trim(x_paint)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("paint") = tmpFld
		tmpFld = Trim(x_salvage)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("salvage") = tmpFld
		tmpFld = Trim(x_comments)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("comments") = tmpFld
		tmpFld = Trim(x_stock)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("stock") = tmpFld
		rs.Update
		rs.Close
		Set rs = Nothing
		conn.Close
		Set conn = Nothing
		Response.Clear
		Response.Redirect "offerthanks.asp"
End Select
%>
<!--#include file="header.asp"-->
<meta name="generator" content="Namo WebEditor v5.0(Trial)">
<table align="center" cellpadding="0" cellspacing="0" width="802" bgcolor="white">
    <tr>
        <td height="28">
            <p>&nbsp;</p>
        </td>
    </tr>
    <tr>
        <td height="66">
<script language="JavaScript" src="ew.js"></script>
<script language="JavaScript">
<!-- start Javascript
function  EW_checkMyForm(EW_this) {
return true;
}
// end JavaScript -->
</script>
<form onSubmit="return EW_checkMyForm(this);"  action="tradeinadd.asp" method="post">
<p>
<input type="hidden" name="a" value="A">
<table border="0" cellspacing="0" cellpadding="0" bgcolor="#CCCCCC" align="center" width="752">
<tr>
<td bgcolor="white" width="283" colspan="2"><font face="Arial"><b><span style="font-size:14pt;">Trade-In Appraisal</span></b></font></td>
<td bgcolor="white" width="312" colspan="2"><font face="Arial" size="2"><%= x_ID %><input type="hidden" name="x_ID" value="<%= x_ID %>"></font>&nbsp;</td>
<td bgcolor="white" width="157" colspan="2">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td bgcolor="white" width="283" colspan="2"><DIV class="smalltext required"><font face="Arial" color="maroon"><span style="font-size:8pt;">* indicates required fields.</span></font></DIV></td>
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
<td bgcolor="white" width="116"><font face="Arial" size="2" color="maroon">*First 
                            Name:</font></td>
<td bgcolor="white" width="167"><font face="Arial" size="2"><input type="text" name="x_first" size="20" maxlength=50 value="<%= Server.HtmlEncode(x_first&"") %>"></font></td>
<td bgcolor="white" width="131"><font face="Arial" size="2">Air Conditioning:</font></td>
<td bgcolor="white" width="181"><font face="Arial" size="2"><%
x_acList = "<SELECT name='x_ac'><OPTION value=''>- - -</OPTION>"
    x_acList = x_acList & "<OPTION value=""Yes"""
    If x_ac = "Yes" Then
        x_acList = x_acList & " selected"
    End If
    x_acList = x_acList & ">" & "Yes" & "</option>"
    x_acList = x_acList & "<OPTION value=""No"""
    If x_ac = "No" Then
        x_acList = x_acList & " selected"
    End If
    x_acList = x_acList & ">" & "No" & "</option>"
x_acList = x_acList & "</select>"
response.write x_acList
%></font></td>
<td bgcolor="white" width="130"><font face="Arial" size="2">Rear DVD:</font></td>
<td bgcolor="white" width="27"><font face="Arial" size="2"><%
x_dvdList = "<SELECT name='x_dvd'><OPTION value=''>- - -</OPTION>"
    x_dvdList = x_dvdList & "<OPTION value=""Yes"""
    If x_dvd = "Yes" Then
        x_dvdList = x_dvdList & " selected"
    End If
    x_dvdList = x_dvdList & ">" & "Yes" & "</option>"
    x_dvdList = x_dvdList & "<OPTION value=""No"""
    If x_dvd = "No" Then
        x_dvdList = x_dvdList & " selected"
    End If
    x_dvdList = x_dvdList & ">" & "No" & "</option>"
x_dvdList = x_dvdList & "</select>"
response.write x_dvdList
%></font></td>
</tr>
<tr>
<td bgcolor="white" width="116"><font face="Arial" size="2" color="maroon">*Last 
                            Name:</font></td>
<td bgcolor="white" width="167"><font face="Arial" size="2"><input type="text" name="x_last" size="20" maxlength=50 value="<%= Server.HtmlEncode(x_last&"") %>"></font></td>
<td bgcolor="white" width="131"><font face="Arial" size="2">Power Windows:</font></td>
<td bgcolor="white" width="181"><font face="Arial" size="2"><%
x_pw_windowsList = "<SELECT name='x_pw_windows'><OPTION value=''>- - -</OPTION>"
    x_pw_windowsList = x_pw_windowsList & "<OPTION value=""Yes"""
    If x_pw_windows = "Yes" Then
        x_pw_windowsList = x_pw_windowsList & " selected"
    End If
    x_pw_windowsList = x_pw_windowsList & ">" & "Yes" & "</option>"
    x_pw_windowsList = x_pw_windowsList & "<OPTION value=""No"""
    If x_pw_windows = "No" Then
        x_pw_windowsList = x_pw_windowsList & " selected"
    End If
    x_pw_windowsList = x_pw_windowsList & ">" & "No" & "</option>"
x_pw_windowsList = x_pw_windowsList & "</select>"
response.write x_pw_windowsList
%></font></td>
<td bgcolor="white" width="130"><font face="Arial" size="2">Satellite Radio:</font></td>
<td bgcolor="white" width="27"><font face="Arial" size="2"><%
x_satelitList = "<SELECT name='x_satelit'><OPTION value=''>- - -</OPTION>"
    x_satelitList = x_satelitList & "<OPTION value=""Yes"""
    If x_satelit = "Yes" Then
        x_satelitList = x_satelitList & " selected"
    End If
    x_satelitList = x_satelitList & ">" & "Yes" & "</option>"
    x_satelitList = x_satelitList & "<OPTION value=""No"""
    If x_satelit = "No" Then
        x_satelitList = x_satelitList & " selected"
    End If
    x_satelitList = x_satelitList & ">" & "No" & "</option>"
x_satelitList = x_satelitList & "</select>"
response.write x_satelitList
%></font></td>
</tr>
<tr>
<td bgcolor="white" width="116"><font face="Arial" size="2" color="maroon">*Home Phone:</font></td>
<td bgcolor="white" width="167"><font face="Arial" size="2"><input type="text" name="x_home_phone" size="20" maxlength=50 value="<%= Server.HtmlEncode(x_home_phone&"") %>"></font></td>
<td bgcolor="white" width="131"><font face="Arial" size="2">Power Locks:</font></td>
<td bgcolor="white" width="181"><font face="Arial" size="2"><%
x_pw_locksList = "<SELECT name='x_pw_locks'><OPTION value=''>- - -</OPTION>"
    x_pw_locksList = x_pw_locksList & "<OPTION value=""Yes"""
    If x_pw_locks = "Yes" Then
        x_pw_locksList = x_pw_locksList & " selected"
    End If
    x_pw_locksList = x_pw_locksList & ">" & "Yes" & "</option>"
    x_pw_locksList = x_pw_locksList & "<OPTION value=""No"""
    If x_pw_locks = "No" Then
        x_pw_locksList = x_pw_locksList & " selected"
    End If
    x_pw_locksList = x_pw_locksList & ">" & "No" & "</option>"
x_pw_locksList = x_pw_locksList & "</select>"
response.write x_pw_locksList
%></font></td>
<td bgcolor="white" width="130"><font face="Arial" size="2">CD Player / Changer:</font></td>
<td bgcolor="white" width="27"><font face="Arial" size="2"><%
x_cd_cd_chList = "<SELECT name='x_cd_cd_ch'><OPTION value=''>- - -</OPTION>"
    x_cd_cd_chList = x_cd_cd_chList & "<OPTION value=""Yes"""
    If x_cd_cd_ch = "Yes" Then
        x_cd_cd_chList = x_cd_cd_chList & " selected"
    End If
    x_cd_cd_chList = x_cd_cd_chList & ">" & "Yes" & "</option>"
    x_cd_cd_chList = x_cd_cd_chList & "<OPTION value=""No"""
    If x_cd_cd_ch = "No" Then
        x_cd_cd_chList = x_cd_cd_chList & " selected"
    End If
    x_cd_cd_chList = x_cd_cd_chList & ">" & "No" & "</option>"
x_cd_cd_chList = x_cd_cd_chList & "</select>"
response.write x_cd_cd_chList
%></font></td>
</tr>
<tr>
<td bgcolor="white" width="116"><font face="Arial" size="2">Work Phone:</font></td>
<td bgcolor="white" width="167"><font face="Arial" size="2"><input type="text" name="x_work_phone" size="20" maxlength=50 value="<%= Server.HtmlEncode(x_work_phone&"") %>"></font></td>
<td bgcolor="white" width="131"><font face="Arial" size="2">Power Seats:</font></td>
<td bgcolor="white" width="181"><font face="Arial" size="2"><%
x_pw_seatsList = "<SELECT name='x_pw_seats'><OPTION value=''>- - -</OPTION>"
    x_pw_seatsList = x_pw_seatsList & "<OPTION value=""Yes"""
    If x_pw_seats = "Yes" Then
        x_pw_seatsList = x_pw_seatsList & " selected"
    End If
    x_pw_seatsList = x_pw_seatsList & ">" & "Yes" & "</option>"
    x_pw_seatsList = x_pw_seatsList & "<OPTION value=""No"""
    If x_pw_seats = "No" Then
        x_pw_seatsList = x_pw_seatsList & " selected"
    End If
    x_pw_seatsList = x_pw_seatsList & ">" & "No" & "</option>"
x_pw_seatsList = x_pw_seatsList & "</select>"
response.write x_pw_seatsList
%></font></td>
<td bgcolor="white" width="130"><font face="Arial" size="2">AM/FM Stereo:</font></td>
<td bgcolor="white" width="27"><font face="Arial" size="2"><%
x_am_fmList = "<SELECT name='x_am_fm'><OPTION value=''>- - -</OPTION>"
    x_am_fmList = x_am_fmList & "<OPTION value=""Yes"""
    If x_am_fm = "Yes" Then
        x_am_fmList = x_am_fmList & " selected"
    End If
    x_am_fmList = x_am_fmList & ">" & "Yes" & "</option>"
    x_am_fmList = x_am_fmList & "<OPTION value=""No"""
    If x_am_fm = "No" Then
        x_am_fmList = x_am_fmList & " selected"
    End If
    x_am_fmList = x_am_fmList & ">" & "No" & "</option>"
x_am_fmList = x_am_fmList & "</select>"
response.write x_am_fmList
%></font></td>
</tr>
<tr>
<td bgcolor="white" width="116"><font face="Arial" size="2" color="maroon">*E-mail 
                            Address:</font></td>
<td bgcolor="white" width="167"><font face="Arial" size="2"><input type="text" name="x_email" size="20" maxlength=50 value="<%= Server.HtmlEncode(x_email&"") %>"></font></td>
<td bgcolor="white" width="131"><font face="Arial" size="2">Power Steering:</font></td>
<td bgcolor="white" width="181"><font face="Arial" size="2"><%
x_pw_steeringList = "<SELECT name='x_pw_steering'><OPTION value=''>- - -</OPTION>"
    x_pw_steeringList = x_pw_steeringList & "<OPTION value=""Yes"""
    If x_pw_steering = "Yes" Then
        x_pw_steeringList = x_pw_steeringList & " selected"
    End If
    x_pw_steeringList = x_pw_steeringList & ">" & "Yes" & "</option>"
    x_pw_steeringList = x_pw_steeringList & "<OPTION value=""No"""
    If x_pw_steering = "No" Then
        x_pw_steeringList = x_pw_steeringList & " selected"
    End If
    x_pw_steeringList = x_pw_steeringList & ">" & "No" & "</option>"
x_pw_steeringList = x_pw_steeringList & "</select>"
response.write x_pw_steeringList
%></font></td>
<td bgcolor="white" width="130"><font face="Arial" size="2">Cassette:</font></td>
<td bgcolor="white" width="27"><font face="Arial" size="2"><%
x_cassList = "<SELECT name='x_cass'><OPTION value=''>- - -</OPTION>"
    x_cassList = x_cassList & "<OPTION value=""Yes"""
    If x_cass = "Yes" Then
        x_cassList = x_cassList & " selected"
    End If
    x_cassList = x_cassList & ">" & "Yes" & "</option>"
    x_cassList = x_cassList & "<OPTION value=""No"""
    If x_cass = "No" Then
        x_cassList = x_cassList & " selected"
    End If
    x_cassList = x_cassList & ">" & "No" & "</option>"
x_cassList = x_cassList & "</select>"
response.write x_cassList
%></font></td>
</tr>
<tr>
<td bgcolor="white" width="116">
                            <p><font face="Arial" size="2">&nbsp;</font></p>
</td>
<td bgcolor="white" width="167">
                            <p><font face="Arial" size="2">&nbsp;</font></p>
</td>
<td bgcolor="white" width="131"><font face="Arial" size="2">Cruise Control:</font></td>
<td bgcolor="white" width="181"><font face="Arial" size="2"><%
x_cr_ctList = "<SELECT name='x_cr_ct'><OPTION value=''>- - -</OPTION>"
    x_cr_ctList = x_cr_ctList & "<OPTION value=""Yes"""
    If x_cr_ct = "Yes" Then
        x_cr_ctList = x_cr_ctList & " selected"
    End If
    x_cr_ctList = x_cr_ctList & ">" & "Yes" & "</option>"
    x_cr_ctList = x_cr_ctList & "<OPTION value=""No"""
    If x_cr_ct = "No" Then
        x_cr_ctList = x_cr_ctList & " selected"
    End If
    x_cr_ctList = x_cr_ctList & ">" & "No" & "</option>"
x_cr_ctList = x_cr_ctList & "</select>"
response.write x_cr_ctList
%></font></td>
<td bgcolor="white" width="130"><font face="Arial" size="2">Leather Interior:</font></td>
<td bgcolor="white" width="27"><font face="Arial" size="2"><%
x_leatherList = "<SELECT name='x_leather'><OPTION value=''>- - -</OPTION>"
    x_leatherList = x_leatherList & "<OPTION value=""Yes"""
    If x_leather = "Yes" Then
        x_leatherList = x_leatherList & " selected"
    End If
    x_leatherList = x_leatherList & ">" & "Yes" & "</option>"
    x_leatherList = x_leatherList & "<OPTION value=""No"""
    If x_leather = "No" Then
        x_leatherList = x_leatherList & " selected"
    End If
    x_leatherList = x_leatherList & ">" & "No" & "</option>"
x_leatherList = x_leatherList & "</select>"
response.write x_leatherList
%></font></td>
</tr>
<tr>
<td bgcolor="white" width="116">
                            <p><font face="Arial" size="2">&nbsp;</font></p>
</td>
<td bgcolor="white" width="167">
                            <p><font face="Arial" size="2">&nbsp;</font></p>
</td>
<td bgcolor="white" width="131"><font face="Arial" size="2">Navigation System:</font></td>
<td bgcolor="white" width="181"><font face="Arial" size="2"><%
x_navigList = "<SELECT name='x_navig'><OPTION value=''>- - -</OPTION>"
    x_navigList = x_navigList & "<OPTION value=""Yes"""
    If x_navig = "Yes" Then
        x_navigList = x_navigList & " selected"
    End If
    x_navigList = x_navigList & ">" & "Yes" & "</option>"
    x_navigList = x_navigList & "<OPTION value=""No"""
    If x_navig = "No" Then
        x_navigList = x_navigList & " selected"
    End If
    x_navigList = x_navigList & ">" & "No" & "</option>"
x_navigList = x_navigList & "</select>"
response.write x_navigList
%></font></td>
<td bgcolor="white" width="130"><font face="Arial" size="2">Alloy Wheels:</font></td>
<td bgcolor="white" width="27"><font face="Arial" size="2"><%
x_alloyList = "<SELECT name='x_alloy'><OPTION value=''>- - -</OPTION>"
    x_alloyList = x_alloyList & "<OPTION value=""Yes"""
    If x_alloy = "Yes" Then
        x_alloyList = x_alloyList & " selected"
    End If
    x_alloyList = x_alloyList & ">" & "Yes" & "</option>"
    x_alloyList = x_alloyList & "<OPTION value=""No"""
    If x_alloy = "No" Then
        x_alloyList = x_alloyList & " selected"
    End If
    x_alloyList = x_alloyList & ">" & "No" & "</option>"
x_alloyList = x_alloyList & "</select>"
response.write x_alloyList
%></font></td>
</tr>
<tr>
<td bgcolor="white" width="116">
                            <p><font face="Arial" size="2">&nbsp;</font></p>
</td>
<td bgcolor="white" width="167">
                            <p><font face="Arial" size="2">&nbsp;</font></p>
</td>
<td bgcolor="white" width="131"><font face="Arial" size="2">Sunroof:</font></td>
<td bgcolor="white" width="181"><font face="Arial" size="2"><%
x_sunroofList = "<SELECT name='x_sunroof'><OPTION value=''>- - -</OPTION>"
    x_sunroofList = x_sunroofList & "<OPTION value=""Yes"""
    If x_sunroof = "Yes" Then
        x_sunroofList = x_sunroofList & " selected"
    End If
    x_sunroofList = x_sunroofList & ">" & "Yes" & "</option>"
    x_sunroofList = x_sunroofList & "<OPTION value=""No"""
    If x_sunroof = "No" Then
        x_sunroofList = x_sunroofList & " selected"
    End If
    x_sunroofList = x_sunroofList & ">" & "No" & "</option>"
x_sunroofList = x_sunroofList & "</select>"
response.write x_sunroofList
%></font></td>
<td bgcolor="white" width="130"><font face="Arial" size="2">Spoiler:</font></td>
<td bgcolor="white" width="27"><font face="Arial" size="2"><%
x_spoilerList = "<SELECT name='x_spoiler'><OPTION value=''>- - -</OPTION>"
    x_spoilerList = x_spoilerList & "<OPTION value=""Yes"""
    If x_spoiler = "Yes" Then
        x_spoilerList = x_spoilerList & " selected"
    End If
    x_spoilerList = x_spoilerList & ">" & "Yes" & "</option>"
    x_spoilerList = x_spoilerList & "<OPTION value=""No"""
    If x_spoiler = "No" Then
        x_spoilerList = x_spoilerList & " selected"
    End If
    x_spoilerList = x_spoilerList & ">" & "No" & "</option>"
x_spoilerList = x_spoilerList & "</select>"
response.write x_spoilerList
%></font></td>
</tr>
<tr>
<td bgcolor="white" width="752" colspan="6">
                            <p><font face="Arial" size="2">&nbsp;</font></p>
</td>
</tr>
<tr>
<td bgcolor="white" width="116"><font face="Arial" size="2" color="maroon">*Year</font><font face="Arial" size="2">&nbsp;</font></td>
<td bgcolor="white" width="167">
<p><font face="Arial" size="2"><%
x_yearList = "<SELECT name='x_year'><OPTION value=''>Please Select</OPTION>"
    x_yearList = x_yearList & "<OPTION value=""2010"""
    If x_year = "2010" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "2010" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""2009"""
    If x_year = "2009" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "2009" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""2008"""
    If x_year = "2008" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "2008" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""2007"""
    If x_year = "2007" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "2007" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""2006"""
    If x_year = "2006" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "2006" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""2005"""
    If x_year = "2005" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "2005" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""2004"""
    If x_year = "2004" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "2004" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""2003"""
    If x_year = "2003" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "2003" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""2002"""
    If x_year = "2002" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "2002" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""2001"""
    If x_year = "2001" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "2001" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""2000"""
    If x_year = "2000" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "2000" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""1999"""
    If x_year = "1999" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "1999" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""1998"""
    If x_year = "1998" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "1998" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""1997"""
    If x_year = "1997" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "1997" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""1996"""
    If x_year = "1996" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "1996" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""1995"""
    If x_year = "1995" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "1995" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""1994"""
    If x_year = "1994" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "1994" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""1993"""
    If x_year = "1993" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "1993" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""1992"""
    If x_year = "1992" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "1992" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""1991"""
    If x_year = "1991" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "1991" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""1990"""
    If x_year = "1990" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "1990" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""1989"""
    If x_year = "1989" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "1989" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""1988"""
    If x_year = "1988" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "1988" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""1887"""
    If x_year = "1887" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "1987" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""1886"""
    If x_year = "1886" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "1986" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""1885"""
    If x_year = "1885" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "1985" & "</option>"
x_yearList = x_yearList & "</select>"
response.write x_yearList
%>
</font></td>
<td bgcolor="white" width="469" colspan="4">
                            <font face="Arial"><span style="font-size:11pt;">Please Rate Your Vehicle On A Scale Of 1 To 10 (10 is Perfect):</span></font></td>
</tr>
<tr>
<td bgcolor="white" width="116"><font face="Arial" size="2" color="maroon">*Make</font><font face="Arial" size="2">&nbsp;</font></td>
<td bgcolor="white" width="167"><font face="Arial" size="2"><input type="text" name="x_make" size="20" maxlength=50 value="<%= Server.HtmlEncode(x_make&"") %>"></font></td>
<td bgcolor="white" width="312" colspan="2"><font face="Arial" size="2">&nbsp;</font></td>
<td bgcolor="white" width="157" colspan="2">
                            <p><font face="Arial" size="2">&nbsp;</font></p>
</td>
</tr>
<tr>
<td bgcolor="white" width="116"><font face="Arial" size="2" color="maroon">*Model</font><font face="Arial" size="2">&nbsp;</font></td>
<td bgcolor="white" width="167"><font face="Arial" size="2"><input type="text" name="x_model" size="20" maxlength=50 value="<%= Server.HtmlEncode(x_model&"") %>"></font></td>
<td bgcolor="white" width="312" colspan="2"><font face="Arial" size="2">Body (dents, dings, rust, rot, damage):</font></td>
<td bgcolor="white" width="157" colspan="2"><font face="Arial" size="2"><%
x_bodyList = "<SELECT name='x_body'><OPTION value=''>- - -</OPTION>"
    x_bodyList = x_bodyList & "<OPTION value=""1"""
    If x_body = "1" Then
        x_bodyList = x_bodyList & " selected"
    End If
    x_bodyList = x_bodyList & ">" & "1" & "</option>"
    x_bodyList = x_bodyList & "<OPTION value=""2"""
    If x_body = "2" Then
        x_bodyList = x_bodyList & " selected"
    End If
    x_bodyList = x_bodyList & ">" & "2" & "</option>"
    x_bodyList = x_bodyList & "<OPTION value=""3"""
    If x_body = "3" Then
        x_bodyList = x_bodyList & " selected"
    End If
    x_bodyList = x_bodyList & ">" & "3" & "</option>"
    x_bodyList = x_bodyList & "<OPTION value=""4"""
    If x_body = "4" Then
        x_bodyList = x_bodyList & " selected"
    End If
    x_bodyList = x_bodyList & ">" & "4" & "</option>"
    x_bodyList = x_bodyList & "<OPTION value=""5"""
    If x_body = "5" Then
        x_bodyList = x_bodyList & " selected"
    End If
    x_bodyList = x_bodyList & ">" & "5" & "</option>"
    x_bodyList = x_bodyList & "<OPTION value=""6"""
    If x_body = "6" Then
        x_bodyList = x_bodyList & " selected"
    End If
    x_bodyList = x_bodyList & ">" & "6" & "</option>"
    x_bodyList = x_bodyList & "<OPTION value=""7"""
    If x_body = "7" Then
        x_bodyList = x_bodyList & " selected"
    End If
    x_bodyList = x_bodyList & ">" & "7" & "</option>"
    x_bodyList = x_bodyList & "<OPTION value=""8"""
    If x_body = "8" Then
        x_bodyList = x_bodyList & " selected"
    End If
    x_bodyList = x_bodyList & ">" & "8" & "</option>"
    x_bodyList = x_bodyList & "<OPTION value=""9"""
    If x_body = "9" Then
        x_bodyList = x_bodyList & " selected"
    End If
    x_bodyList = x_bodyList & ">" & "9" & "</option>"
    x_bodyList = x_bodyList & "<OPTION value=""10"""
    If x_body = "10" Then
        x_bodyList = x_bodyList & " selected"
    End If
    x_bodyList = x_bodyList & ">" & "10" & "</option>"
x_bodyList = x_bodyList & "</select>"
response.write x_bodyList
%>
</font></td>
</tr>
<tr>
<td bgcolor="white" width="116"><font face="Arial" size="2">Ext color&nbsp;</font></td>
<td bgcolor="white" width="167"><font face="Arial" size="2"><input type="text" name="x_ext_color" size="20" maxlength=50 value="<%= Server.HtmlEncode(x_ext_color&"") %>"></font></td>
<td bgcolor="white" width="312" colspan="2"><font face="Arial" size="2">Tires (tread wear, mismatched):</font></td>
<td bgcolor="white" width="157" colspan="2"><font face="Arial" size="2"><%
x_tiresList = "<SELECT name='x_tires'><OPTION value=''>- - -</OPTION>"
    x_tiresList = x_tiresList & "<OPTION value=""1"""
    If x_tires = "1" Then
        x_tiresList = x_tiresList & " selected"
    End If
    x_tiresList = x_tiresList & ">" & "1" & "</option>"
    x_tiresList = x_tiresList & "<OPTION value=""2"""
    If x_tires = "2" Then
        x_tiresList = x_tiresList & " selected"
    End If
    x_tiresList = x_tiresList & ">" & "2" & "</option>"
    x_tiresList = x_tiresList & "<OPTION value=""3"""
    If x_tires = "3" Then
        x_tiresList = x_tiresList & " selected"
    End If
    x_tiresList = x_tiresList & ">" & "3" & "</option>"
    x_tiresList = x_tiresList & "<OPTION value=""4"""
    If x_tires = "4" Then
        x_tiresList = x_tiresList & " selected"
    End If
    x_tiresList = x_tiresList & ">" & "4" & "</option>"
    x_tiresList = x_tiresList & "<OPTION value=""5"""
    If x_tires = "5" Then
        x_tiresList = x_tiresList & " selected"
    End If
    x_tiresList = x_tiresList & ">" & "5" & "</option>"
    x_tiresList = x_tiresList & "<OPTION value=""6"""
    If x_tires = "6" Then
        x_tiresList = x_tiresList & " selected"
    End If
    x_tiresList = x_tiresList & ">" & "6" & "</option>"
    x_tiresList = x_tiresList & "<OPTION value=""7"""
    If x_tires = "7" Then
        x_tiresList = x_tiresList & " selected"
    End If
    x_tiresList = x_tiresList & ">" & "7" & "</option>"
    x_tiresList = x_tiresList & "<OPTION value=""8"""
    If x_tires = "8" Then
        x_tiresList = x_tiresList & " selected"
    End If
    x_tiresList = x_tiresList & ">" & "8" & "</option>"
    x_tiresList = x_tiresList & "<OPTION value=""9"""
    If x_tires = "9" Then
        x_tiresList = x_tiresList & " selected"
    End If
    x_tiresList = x_tiresList & ">" & "9" & "</option>"
    x_tiresList = x_tiresList & "<OPTION value=""10"""
    If x_tires = "10" Then
        x_tiresList = x_tiresList & " selected"
    End If
    x_tiresList = x_tiresList & ">" & "10" & "</option>"
x_tiresList = x_tiresList & "</select>"
response.write x_tiresList
%>
</font></td>
</tr>
<tr>
<td bgcolor="white" width="116"><font face="Arial" size="2">VIN</font></td>
<td bgcolor="white" width="167"><font face="Arial" size="2"><input type="text" name="x_vin" size="20" maxlength=50 value="<%= Server.HtmlEncode(x_vin&"") %>"></font></td>
<td bgcolor="white" width="312" colspan="2"><font face="Arial" size="2">Engine (running condition, burns oil, knocking):</font></td>
<td bgcolor="white" width="157" colspan="2"><font face="Arial" size="2"><%
x_engine_rateList = "<SELECT name='x_engine_rate'><OPTION value=''>- - -</OPTION>"
    x_engine_rateList = x_engine_rateList & "<OPTION value=""1"""
    If x_engine_rate = "1" Then
        x_engine_rateList = x_engine_rateList & " selected"
    End If
    x_engine_rateList = x_engine_rateList & ">" & "1" & "</option>"
    x_engine_rateList = x_engine_rateList & "<OPTION value=""2"""
    If x_engine_rate = "2" Then
        x_engine_rateList = x_engine_rateList & " selected"
    End If
    x_engine_rateList = x_engine_rateList & ">" & "2" & "</option>"
    x_engine_rateList = x_engine_rateList & "<OPTION value=""3"""
    If x_engine_rate = "3" Then
        x_engine_rateList = x_engine_rateList & " selected"
    End If
    x_engine_rateList = x_engine_rateList & ">" & "3" & "</option>"
    x_engine_rateList = x_engine_rateList & "<OPTION value=""4"""
    If x_engine_rate = "4" Then
        x_engine_rateList = x_engine_rateList & " selected"
    End If
    x_engine_rateList = x_engine_rateList & ">" & "4" & "</option>"
    x_engine_rateList = x_engine_rateList & "<OPTION value=""5"""
    If x_engine_rate = "5" Then
        x_engine_rateList = x_engine_rateList & " selected"
    End If
    x_engine_rateList = x_engine_rateList & ">" & "5" & "</option>"
    x_engine_rateList = x_engine_rateList & "<OPTION value=""6"""
    If x_engine_rate = "6" Then
        x_engine_rateList = x_engine_rateList & " selected"
    End If
    x_engine_rateList = x_engine_rateList & ">" & "6" & "</option>"
    x_engine_rateList = x_engine_rateList & "<OPTION value=""7"""
    If x_engine_rate = "7" Then
        x_engine_rateList = x_engine_rateList & " selected"
    End If
    x_engine_rateList = x_engine_rateList & ">" & "7" & "</option>"
    x_engine_rateList = x_engine_rateList & "<OPTION value=""8"""
    If x_engine_rate = "8" Then
        x_engine_rateList = x_engine_rateList & " selected"
    End If
    x_engine_rateList = x_engine_rateList & ">" & "8" & "</option>"
    x_engine_rateList = x_engine_rateList & "<OPTION value=""9"""
    If x_engine_rate = "9" Then
        x_engine_rateList = x_engine_rateList & " selected"
    End If
    x_engine_rateList = x_engine_rateList & ">" & "9" & "</option>"
    x_engine_rateList = x_engine_rateList & "<OPTION value=""10"""
    If x_engine_rate = "10" Then
        x_engine_rateList = x_engine_rateList & " selected"
    End If
    x_engine_rateList = x_engine_rateList & ">" & "10" & "</option>"
x_engine_rateList = x_engine_rateList & "</select>"
response.write x_engine_rateList
%>
</font></td>
</tr>
<tr>
<td bgcolor="white" width="116"><font face="Arial" size="2" color="maroon">*Mileage</font><font face="Arial" size="2">&nbsp;</font></td>
<td bgcolor="white" width="167"><font face="Arial" size="2"><input type="text" name="x_mileage" size="20" maxlength=50 value="<%= Server.HtmlEncode(x_mileage&"") %>"></font></td>
<td bgcolor="white" width="312" colspan="2"><font face="Arial" size="2">Transmission / Clutch (slipping, hard shift, grinds):</font></td>
<td bgcolor="white" width="157" colspan="2"><font face="Arial" size="2"><%
x_trans_rateList = "<SELECT name='x_trans_rate'><OPTION value=''>- - -</OPTION>"
    x_trans_rateList = x_trans_rateList & "<OPTION value=""1"""
    If x_trans_rate = "1" Then
        x_trans_rateList = x_trans_rateList & " selected"
    End If
    x_trans_rateList = x_trans_rateList & ">" & "1" & "</option>"
    x_trans_rateList = x_trans_rateList & "<OPTION value=""2"""
    If x_trans_rate = "2" Then
        x_trans_rateList = x_trans_rateList & " selected"
    End If
    x_trans_rateList = x_trans_rateList & ">" & "2" & "</option>"
    x_trans_rateList = x_trans_rateList & "<OPTION value=""3"""
    If x_trans_rate = "3" Then
        x_trans_rateList = x_trans_rateList & " selected"
    End If
    x_trans_rateList = x_trans_rateList & ">" & "3" & "</option>"
    x_trans_rateList = x_trans_rateList & "<OPTION value=""4"""
    If x_trans_rate = "4" Then
        x_trans_rateList = x_trans_rateList & " selected"
    End If
    x_trans_rateList = x_trans_rateList & ">" & "4" & "</option>"
    x_trans_rateList = x_trans_rateList & "<OPTION value=""5"""
    If x_trans_rate = "5" Then
        x_trans_rateList = x_trans_rateList & " selected"
    End If
    x_trans_rateList = x_trans_rateList & ">" & "5" & "</option>"
    x_trans_rateList = x_trans_rateList & "<OPTION value=""6"""
    If x_trans_rate = "6" Then
        x_trans_rateList = x_trans_rateList & " selected"
    End If
    x_trans_rateList = x_trans_rateList & ">" & "6" & "</option>"
    x_trans_rateList = x_trans_rateList & "<OPTION value=""7"""
    If x_trans_rate = "7" Then
        x_trans_rateList = x_trans_rateList & " selected"
    End If
    x_trans_rateList = x_trans_rateList & ">" & "7" & "</option>"
    x_trans_rateList = x_trans_rateList & "<OPTION value=""8"""
    If x_trans_rate = "8" Then
        x_trans_rateList = x_trans_rateList & " selected"
    End If
    x_trans_rateList = x_trans_rateList & ">" & "8" & "</option>"
    x_trans_rateList = x_trans_rateList & "<OPTION value=""9"""
    If x_trans_rate = "9" Then
        x_trans_rateList = x_trans_rateList & " selected"
    End If
    x_trans_rateList = x_trans_rateList & ">" & "9" & "</option>"
    x_trans_rateList = x_trans_rateList & "<OPTION value=""10"""
    If x_trans_rate = "10" Then
        x_trans_rateList = x_trans_rateList & " selected"
    End If
    x_trans_rateList = x_trans_rateList & ">" & "10" & "</option>"
x_trans_rateList = x_trans_rateList & "</select>"
response.write x_trans_rateList
%></font></td>
</tr>
<tr>
<td bgcolor="white" width="116"><font face="Arial" size="2">Engine&nbsp;</font></td>
<td bgcolor="white" width="167"><font face="Arial" size="2"><input type="text" name="x_engine" size="20" maxlength=50 value="<%= Server.HtmlEncode(x_engine&"") %>"></font></td>
<td bgcolor="white" width="312" colspan="2"><font face="Arial" size="2">Glass (chips, scratches, cracks, pitted):</font></td>
<td bgcolor="white" width="157" colspan="2"><font face="Arial" size="2"><%
x_glass_rateList = "<SELECT name='x_glass_rate'><OPTION value=''>- - -</OPTION>"
    x_glass_rateList = x_glass_rateList & "<OPTION value=""1"""
    If x_glass_rate = "1" Then
        x_glass_rateList = x_glass_rateList & " selected"
    End If
    x_glass_rateList = x_glass_rateList & ">" & "1" & "</option>"
    x_glass_rateList = x_glass_rateList & "<OPTION value=""2"""
    If x_glass_rate = "2" Then
        x_glass_rateList = x_glass_rateList & " selected"
    End If
    x_glass_rateList = x_glass_rateList & ">" & "2" & "</option>"
    x_glass_rateList = x_glass_rateList & "<OPTION value=""3"""
    If x_glass_rate = "3" Then
        x_glass_rateList = x_glass_rateList & " selected"
    End If
    x_glass_rateList = x_glass_rateList & ">" & "3" & "</option>"
    x_glass_rateList = x_glass_rateList & "<OPTION value=""4"""
    If x_glass_rate = "4" Then
        x_glass_rateList = x_glass_rateList & " selected"
    End If
    x_glass_rateList = x_glass_rateList & ">" & "4" & "</option>"
    x_glass_rateList = x_glass_rateList & "<OPTION value=""5"""
    If x_glass_rate = "5" Then
        x_glass_rateList = x_glass_rateList & " selected"
    End If
    x_glass_rateList = x_glass_rateList & ">" & "5" & "</option>"
    x_glass_rateList = x_glass_rateList & "<OPTION value=""6"""
    If x_glass_rate = "6" Then
        x_glass_rateList = x_glass_rateList & " selected"
    End If
    x_glass_rateList = x_glass_rateList & ">" & "6" & "</option>"
    x_glass_rateList = x_glass_rateList & "<OPTION value=""7"""
    If x_glass_rate = "7" Then
        x_glass_rateList = x_glass_rateList & " selected"
    End If
    x_glass_rateList = x_glass_rateList & ">" & "7" & "</option>"
    x_glass_rateList = x_glass_rateList & "<OPTION value=""8"""
    If x_glass_rate = "8" Then
        x_glass_rateList = x_glass_rateList & " selected"
    End If
    x_glass_rateList = x_glass_rateList & ">" & "8" & "</option>"
    x_glass_rateList = x_glass_rateList & "<OPTION value=""9"""
    If x_glass_rate = "9" Then
        x_glass_rateList = x_glass_rateList & " selected"
    End If
    x_glass_rateList = x_glass_rateList & ">" & "9" & "</option>"
    x_glass_rateList = x_glass_rateList & "<OPTION value=""10"""
    If x_glass_rate = "10" Then
        x_glass_rateList = x_glass_rateList & " selected"
    End If
    x_glass_rateList = x_glass_rateList & ">" & "10" & "</option>"
x_glass_rateList = x_glass_rateList & "</select>"
response.write x_glass_rateList
%>
</font></td>
</tr>
<tr>
<td bgcolor="white" width="116"><font face="Arial" size="2">Doors&nbsp;</font></td>
<td bgcolor="white" width="167"><font face="Arial" size="2"><%
x_doorsList = "<SELECT name='x_doors'><OPTION value=''>- - -</OPTION>"
    x_doorsList = x_doorsList & "<OPTION value=""5"""
    If x_doors = "5" Then
        x_doorsList = x_doorsList & " selected"
    End If
    x_doorsList = x_doorsList & ">" & "5" & "</option>"
    x_doorsList = x_doorsList & "<OPTION value=""4"""
    If x_doors = "4" Then
        x_doorsList = x_doorsList & " selected"
    End If
    x_doorsList = x_doorsList & ">" & "4" & "</option>"
    x_doorsList = x_doorsList & "<OPTION value=""3"""
    If x_doors = "3" Then
        x_doorsList = x_doorsList & " selected"
    End If
    x_doorsList = x_doorsList & ">" & "3" & "</option>"
    x_doorsList = x_doorsList & "<OPTION value=""2"""
    If x_doors = "2" Then
        x_doorsList = x_doorsList & " selected"
    End If
    x_doorsList = x_doorsList & ">" & "2" & "</option>"
    x_doorsList = x_doorsList & "<OPTION value=""1"""
    If x_doors = "1" Then
        x_doorsList = x_doorsList & " selected"
    End If
    x_doorsList = x_doorsList & ">" & "1" & "</option>"
x_doorsList = x_doorsList & "</select>"
response.write x_doorsList
%>
</font></td>
<td bgcolor="white" width="312" colspan="2">
<font face="Arial" size="2">Interior (rips, tears, burns, faded/worn, stains):</font></td>
<td bgcolor="white" width="157" colspan="2"><font face="Arial" size="2"><%
x_interior_rateList = "<SELECT name='x_interior_rate'><OPTION value=''>- - -</OPTION>"
    x_interior_rateList = x_interior_rateList & "<OPTION value=""1"""
    If x_interior_rate = "1" Then
        x_interior_rateList = x_interior_rateList & " selected"
    End If
    x_interior_rateList = x_interior_rateList & ">" & "1" & "</option>"
    x_interior_rateList = x_interior_rateList & "<OPTION value=""2"""
    If x_interior_rate = "2" Then
        x_interior_rateList = x_interior_rateList & " selected"
    End If
    x_interior_rateList = x_interior_rateList & ">" & "2" & "</option>"
    x_interior_rateList = x_interior_rateList & "<OPTION value=""3"""
    If x_interior_rate = "3" Then
        x_interior_rateList = x_interior_rateList & " selected"
    End If
    x_interior_rateList = x_interior_rateList & ">" & "3" & "</option>"
    x_interior_rateList = x_interior_rateList & "<OPTION value=""4"""
    If x_interior_rate = "4" Then
        x_interior_rateList = x_interior_rateList & " selected"
    End If
    x_interior_rateList = x_interior_rateList & ">" & "4" & "</option>"
    x_interior_rateList = x_interior_rateList & "<OPTION value=""5"""
    If x_interior_rate = "5" Then
        x_interior_rateList = x_interior_rateList & " selected"
    End If
    x_interior_rateList = x_interior_rateList & ">" & "5" & "</option>"
    x_interior_rateList = x_interior_rateList & "<OPTION value=""6"""
    If x_interior_rate = "6" Then
        x_interior_rateList = x_interior_rateList & " selected"
    End If
    x_interior_rateList = x_interior_rateList & ">" & "6" & "</option>"
    x_interior_rateList = x_interior_rateList & "<OPTION value=""7"""
    If x_interior_rate = "7" Then
        x_interior_rateList = x_interior_rateList & " selected"
    End If
    x_interior_rateList = x_interior_rateList & ">" & "7" & "</option>"
    x_interior_rateList = x_interior_rateList & "<OPTION value=""8"""
    If x_interior_rate = "8" Then
        x_interior_rateList = x_interior_rateList & " selected"
    End If
    x_interior_rateList = x_interior_rateList & ">" & "8" & "</option>"
    x_interior_rateList = x_interior_rateList & "<OPTION value=""9"""
    If x_interior_rate = "9" Then
        x_interior_rateList = x_interior_rateList & " selected"
    End If
    x_interior_rateList = x_interior_rateList & ">" & "9" & "</option>"
    x_interior_rateList = x_interior_rateList & "<OPTION value=""10"""
    If x_interior_rate = "10" Then
        x_interior_rateList = x_interior_rateList & " selected"
    End If
    x_interior_rateList = x_interior_rateList & ">" & "10" & "</option>"
x_interior_rateList = x_interior_rateList & "</select>"
response.write x_interior_rateList
%></font></td>
</tr>
<tr>
<td bgcolor="white" width="116"><font face="Arial" size="2">Transmission&nbsp;</font></td>
<td bgcolor="white" width="167"><font face="Arial" size="2"><%
x_transmissionList = "<SELECT name='x_transmission'><OPTION value=''>- - - - - - - - -</OPTION>"
    x_transmissionList = x_transmissionList & "<OPTION value=""Automatic"""
    If x_transmission = "Automatic" Then
        x_transmissionList = x_transmissionList & " selected"
    End If
    x_transmissionList = x_transmissionList & ">" & "Automatic" & "</option>"
    x_transmissionList = x_transmissionList & "<OPTION value=""Manual"""
    If x_transmission = "Manual" Then
        x_transmissionList = x_transmissionList & " selected"
    End If
    x_transmissionList = x_transmissionList & ">" & "Manual" & "</option>"
x_transmissionList = x_transmissionList & "</select>"
response.write x_transmissionList
%>
</font></td>
<td bgcolor="white" width="312" colspan="2"><font face="Arial" size="2">Exhaust (rusted, leaking, noisy):</font></td>
<td bgcolor="white" width="157" colspan="2"><font face="Arial" size="2"><%
x_exhouse_rateList = "<SELECT name='x_exhouse_rate'><OPTION value=''>- - -</OPTION>"
    x_exhouse_rateList = x_exhouse_rateList & "<OPTION value=""1"""
    If x_exhouse_rate = "1" Then
        x_exhouse_rateList = x_exhouse_rateList & " selected"
    End If
    x_exhouse_rateList = x_exhouse_rateList & ">" & "1" & "</option>"
    x_exhouse_rateList = x_exhouse_rateList & "<OPTION value=""2"""
    If x_exhouse_rate = "2" Then
        x_exhouse_rateList = x_exhouse_rateList & " selected"
    End If
    x_exhouse_rateList = x_exhouse_rateList & ">" & "2" & "</option>"
    x_exhouse_rateList = x_exhouse_rateList & "<OPTION value=""3"""
    If x_exhouse_rate = "3" Then
        x_exhouse_rateList = x_exhouse_rateList & " selected"
    End If
    x_exhouse_rateList = x_exhouse_rateList & ">" & "3" & "</option>"
    x_exhouse_rateList = x_exhouse_rateList & "<OPTION value=""4"""
    If x_exhouse_rate = "4" Then
        x_exhouse_rateList = x_exhouse_rateList & " selected"
    End If
    x_exhouse_rateList = x_exhouse_rateList & ">" & "4" & "</option>"
    x_exhouse_rateList = x_exhouse_rateList & "<OPTION value=""5"""
    If x_exhouse_rate = "5" Then
        x_exhouse_rateList = x_exhouse_rateList & " selected"
    End If
    x_exhouse_rateList = x_exhouse_rateList & ">" & "5" & "</option>"
    x_exhouse_rateList = x_exhouse_rateList & "<OPTION value=""6"""
    If x_exhouse_rate = "6" Then
        x_exhouse_rateList = x_exhouse_rateList & " selected"
    End If
    x_exhouse_rateList = x_exhouse_rateList & ">" & "6" & "</option>"
    x_exhouse_rateList = x_exhouse_rateList & "<OPTION value=""7"""
    If x_exhouse_rate = "7" Then
        x_exhouse_rateList = x_exhouse_rateList & " selected"
    End If
    x_exhouse_rateList = x_exhouse_rateList & ">" & "7" & "</option>"
    x_exhouse_rateList = x_exhouse_rateList & "<OPTION value=""8"""
    If x_exhouse_rate = "8" Then
        x_exhouse_rateList = x_exhouse_rateList & " selected"
    End If
    x_exhouse_rateList = x_exhouse_rateList & ">" & "8" & "</option>"
    x_exhouse_rateList = x_exhouse_rateList & "<OPTION value=""9"""
    If x_exhouse_rate = "9" Then
        x_exhouse_rateList = x_exhouse_rateList & " selected"
    End If
    x_exhouse_rateList = x_exhouse_rateList & ">" & "9" & "</option>"
    x_exhouse_rateList = x_exhouse_rateList & "<OPTION value=""10"""
    If x_exhouse_rate = "10" Then
        x_exhouse_rateList = x_exhouse_rateList & " selected"
    End If
    x_exhouse_rateList = x_exhouse_rateList & ">" & "10" & "</option>"
x_exhouse_rateList = x_exhouse_rateList & "</select>"
response.write x_exhouse_rateList
%></font></td>
</tr>
<tr>
<td bgcolor="white" width="116"><font face="Arial" size="2">Drivetrain&nbsp;</font></td>
<td bgcolor="white" width="167"><font face="Arial" size="2"><%
x_drivetrainList = "<SELECT name='x_drivetrain'><OPTION value=''>- - -</OPTION>"
    x_drivetrainList = x_drivetrainList & "<OPTION value=""FWD"""
    If x_drivetrain = "FWD" Then
        x_drivetrainList = x_drivetrainList & " selected"
    End If
    x_drivetrainList = x_drivetrainList & ">" & "FWD" & "</option>"
    x_drivetrainList = x_drivetrainList & "<OPTION value=""RWD"""
    If x_drivetrain = "RWD" Then
        x_drivetrainList = x_drivetrainList & " selected"
    End If
    x_drivetrainList = x_drivetrainList & ">" & "RWD" & "</option>"
    x_drivetrainList = x_drivetrainList & "<OPTION value=""AWD"""
    If x_drivetrain = "AWD" Then
        x_drivetrainList = x_drivetrainList & " selected"
    End If
    x_drivetrainList = x_drivetrainList & ">" & "AWD" & "</option>"
x_drivetrainList = x_drivetrainList & "</select>"
response.write x_drivetrainList
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
<td bgcolor="white" width="131"><font face="Arial" size="2"><%
x_lease_rentalList = "<SELECT name='x_lease_rental'><OPTION value=''>- - -</OPTION>"
    x_lease_rentalList = x_lease_rentalList & "<OPTION value=""Yes"""
    If x_lease_rental = "Yes" Then
        x_lease_rentalList = x_lease_rentalList & " selected"
    End If
    x_lease_rentalList = x_lease_rentalList & ">" & "Yes" & "</option>"
    x_lease_rentalList = x_lease_rentalList & "<OPTION value=""No"""
    If x_lease_rental = "No" Then
        x_lease_rentalList = x_lease_rentalList & " selected"
    End If
    x_lease_rentalList = x_lease_rentalList & ">" & "No" & "</option>"
x_lease_rentalList = x_lease_rentalList & "</select>"
response.write x_lease_rentalList
%>
</font>&nbsp;</td>
<td bgcolor="white" width="181"><font face="Arial" size="2">Are there any lienholders and where are they located?</font></td>
<td bgcolor="white" width="157" colspan="2"><font face="Arial" size="2"><input type="text" name="x_lienholders" size="20" maxlength=50 value="<%= Server.HtmlEncode(x_lienholders&"") %>"></font></td>
</tr>
<tr>
<td bgcolor="white" width="283" colspan="2"><font face="Arial" size="2">Is the odometer operational and accurate?</font></td>
<td bgcolor="white" width="131"><font face="Arial" size="2"><%
x_odometerList = "<SELECT name='x_odometer'><OPTION value=''>- - -</OPTION>"
    x_odometerList = x_odometerList & "<OPTION value=""Yes"""
    If x_odometer = "Yes" Then
        x_odometerList = x_odometerList & " selected"
    End If
    x_odometerList = x_odometerList & ">" & "Yes" & "</option>"
    x_odometerList = x_odometerList & "<OPTION value=""No"""
    If x_odometer = "No" Then
        x_odometerList = x_odometerList & " selected"
    End If
    x_odometerList = x_odometerList & ">" & "No" & "</option>"
x_odometerList = x_odometerList & "</select>"
response.write x_odometerList
%>
</font>&nbsp;</td>
<td bgcolor="white" width="181"><font face="Arial" size="2">Who holds this title?</font></td>
<td bgcolor="white" width="157" colspan="2"><font face="Arial" size="2"><input type="text" name="x_title" size="20" maxlength=50 value="<%= Server.HtmlEncode(x_title&"") %>"></font></td>
</tr>
<tr>
<td bgcolor="white" width="283" colspan="2"><font face="Arial" size="2">Detailed service records available?</font></td>
<td bgcolor="white" width="131"><font face="Arial" size="2"><%
x_recordsList = "<SELECT name='x_records'><OPTION value=''>- - -</OPTION>"
    x_recordsList = x_recordsList & "<OPTION value=""Yes"""
    If x_records = "Yes" Then
        x_recordsList = x_recordsList & " selected"
    End If
    x_recordsList = x_recordsList & ">" & "Yes" & "</option>"
    x_recordsList = x_recordsList & "<OPTION value=""No"""
    If x_records = "No" Then
        x_recordsList = x_recordsList & " selected"
    End If
    x_recordsList = x_recordsList & ">" & "No" & "</option>"
x_recordsList = x_recordsList & "</select>"
response.write x_recordsList
%>
</font>&nbsp;</td>
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
<p align="center">
<input type="submit" name="Action" value="Submit">
</form>
            <p>&nbsp;</p>
        </td>
    </tr>
</table>
<!--#include file="footer.asp"-->