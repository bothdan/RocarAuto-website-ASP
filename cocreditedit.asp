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
If key="" OR IsNull(key) Then key = Request.Form("key")
If key="" OR IsNull(key) Then Response.Redirect "cocreditlist.asp"
'get action
a=Request.Form("a")
If a="" OR IsNull(a) Then
	a="I"	'display with input box
End If
'get fields from form
x_ID = Request.Form("x_ID")
x_first_name = Request.Form("x_first_name")
x_middle = Request.Form("x_middle")
x_last_name = Request.Form("x_last_name")
x_street = Request.Form("x_street")
x_aparment = Request.Form("x_aparment")
x_city = Request.Form("x_city")
x_state = Request.Form("x_state")
x_zip = Request.Form("x_zip")
x_home_phone = Request.Form("x_home_phone")
x_work_phone = Request.Form("x_work_phone")
x_email = Request.Form("x_email")
x_ssn = Request.Form("x_ssn")
x_dob = Request.Form("x_dob")
x_occupation = Request.Form("x_occupation")
x_workplace = Request.Form("x_workplace")
x_net_salary = Request.Form("x_net_salary")
x_timework = Request.Form("x_timework")
x_first_co = Request.Form("x_first_co")
x_middle_co = Request.Form("x_middle_co")
x_last_co = Request.Form("x_last_co")
x_street_co = Request.Form("x_street_co")
x_apartment_co = Request.Form("x_apartment_co")
x_city_co = Request.Form("x_city_co")
x_state_co = Request.Form("x_state_co")
x_zip_co = Request.Form("x_zip_co")
x_home_phone_co = Request.Form("x_home_phone_co")
x_work_phone_co = Request.Form("x_work_phone_co")
x_email_co = Request.Form("x_email_co")
x_ssn_co = Request.Form("x_ssn_co")
x_dob_co = Request.Form("x_dob_co")
x_occupation_co = Request.Form("x_occupation_co")
x_workplace_co = Request.Form("x_workplace_co")
x_net_salary_co = Request.Form("x_net_salary_co")
x_timework_co = Request.Form("x_timework_co")
x_initials = Request.Form("x_initials")
x_iagree = Request.Form("x_iagree")
x_stock = Request.Form("x_stock")
' Open Connection to the database
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str
Select Case a
	Case "I": ' Get a record to display
		tkey = "" & key & ""
		strsql = "SELECT * FROM [cocredit] WHERE [ID]=" & tkey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn
		If rs.EOF Then
			Response.Clear
			Response.Redirect "cocreditlist.asp"
		Else
			rs.MoveFirst
		End If
		' Get the field contents
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
		rs.Close
		Set rs = Nothing
	Case "U": ' Update
		' Open record
		tkey = "" & key & ""
		strsql = "SELECT * FROM [cocredit] WHERE [ID]=" & tkey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn, 1, 2
		If rs.EOF Then
			Response.Clear
			Response.Redirect "cocreditlist.asp"
		End If
		tmpFld = Trim(x_first_name)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("first name") = tmpFld
		tmpFld = Trim(x_middle)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("middle") = tmpFld
		tmpFld = Trim(x_last_name)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("last name") = tmpFld
		tmpFld = Trim(x_street)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("street") = tmpFld
		tmpFld = Trim(x_aparment)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("aparment") = tmpFld
		tmpFld = Trim(x_city)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("city") = tmpFld
		tmpFld = Trim(x_state)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("state") = tmpFld
		tmpFld = Trim(x_zip)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("zip") = tmpFld
		tmpFld = Trim(x_home_phone)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("home phone") = tmpFld
		tmpFld = Trim(x_work_phone)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("work phone") = tmpFld
		tmpFld = Trim(x_email)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("email") = tmpFld
		tmpFld = Trim(x_ssn)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("ssn") = tmpFld
		tmpFld = Trim(x_dob)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("dob") = tmpFld
		tmpFld = Trim(x_occupation)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("occupation") = tmpFld
		tmpFld = Trim(x_workplace)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("workplace") = tmpFld
		tmpFld = Trim(x_net_salary)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("net salary") = tmpFld
		tmpFld = Trim(x_timework)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("timework") = tmpFld
		tmpFld = Trim(x_first_co)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("first co") = tmpFld
		tmpFld = Trim(x_middle_co)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("middle co") = tmpFld
		tmpFld = Trim(x_last_co)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("last co") = tmpFld
		tmpFld = Trim(x_street_co)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("street co") = tmpFld
		tmpFld = Trim(x_apartment_co)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("apartment co") = tmpFld
		tmpFld = Trim(x_city_co)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("city co") = tmpFld
		tmpFld = Trim(x_state_co)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("state co") = tmpFld
		tmpFld = Trim(x_zip_co)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("zip co") = tmpFld
		tmpFld = Trim(x_home_phone_co)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("home phone co") = tmpFld
		tmpFld = Trim(x_work_phone_co)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("work phone co") = tmpFld
		tmpFld = Trim(x_email_co)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("email co") = tmpFld
		tmpFld = Trim(x_ssn_co)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("ssn co") = tmpFld
		tmpFld = Trim(x_dob_co)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("dob co") = tmpFld
		tmpFld = Trim(x_occupation_co)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("occupation co") = tmpFld
		tmpFld = Trim(x_workplace_co)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("workplace co") = tmpFld
		tmpFld = Trim(x_net_salary_co)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("net salary co") = tmpFld
		tmpFld = Trim(x_timework_co)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("timework co") = tmpFld
		tmpFld = Trim(x_initials)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("initials") = tmpFld
		tmpFld = Trim(x_iagree)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("iagree") = tmpFld
		tmpFld = x_stock
		If Not IsNumeric(tmpFld) Then tmpFld = 0
		rs("stock") = cLng(tmpFld)
		rs.Update
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


<table border="0" cellspacing="0" cellpadding="0" bgcolor="#CCCCCC" width="797">
<tr>
<td bgcolor="white" width="797" colspan="3">
                            <p><span style="font-size:14pt;"><b><font face="Arial">&nbsp;&nbsp;&nbsp;<a href="cocreditlist.asp"><img src="images/back.gif" align="middle" width="16" height="16" border="0"></a> 
                             
</font></b></span><a href="cocreditlist.asp"><font face="Arial" size="2" color="black"><b>Back to Co-Applicant 
List</b></font></a>
</p>
<script language="JavaScript" src="ew.js"></script>
<script language="JavaScript">
<!-- start Javascript
function  EW_checkMyForm(EW_this) {
if (EW_this.x_stock && !EW_checkinteger(EW_this.x_stock.value)) {
        if (!EW_onError(EW_this, EW_this.x_stock, "TEXT", "Incorrect integer - stock"))
            return false; 
        }
return true;
}
// end JavaScript -->
</script>
<form onSubmit="return EW_checkMyForm(this);"  action="cocreditedit.asp" method="post">
<p>
<input type="hidden" name="a" value="U">
<input type="hidden" name="key" value="<%= key %>">

<table border="0" cellspacing="0" cellpadding="0" bgcolor="#CCCCCC" width="797">
<tr>
<td bgcolor="white" width="23">
<p>
&nbsp;
</td>
<td bgcolor="white" width="390" colspan="2"><span style="font-size:14pt;"><b><font face="Arial">Your Application Information<br>&nbsp;</font></b></span></td>
<td bgcolor="white" width="384" colspan="2"><span style="font-size:14pt;"><b><font face="Arial">Co-Applicant Information<br>&nbsp;</font></b></span></td>
</tr>
<tr>
<td bgcolor="white" width="23">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="163">
                            <p class="tblFinanceForm coapp"><font face="Arial" color="maroon"><span style="font-size:10pt;">*First 
                            Name:</span></font></p>
</td>
<td bgcolor="white" width="227"><font face="Arial" size="2"><input type="text" name="x_first_name" size="25" maxlength=50 value="<%= Server.HtmlEncode(x_first_name&"") %>"></font></td>
<td bgcolor="white" width="160">
                            <p class="tblFinanceForm coapp"><font face="Arial" color="maroon"><span style="font-size:10pt;">*First 
                            Name:</span></font></p>
</td>
<td bgcolor="white" width="224"><font face="Arial" size="2"><input type="text" name="x_first_co" size="25" maxlength=50 value="<%= Server.HtmlEncode(x_first_co&"") %>"></font></td>
</tr>
<tr>
<td bgcolor="white" width="23" height="21">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="163" height="21"><font face="Arial"><span style="font-size:10pt;">Middle Initial:</span></font></td>
<td bgcolor="white" width="227" height="21"><font face="Arial" size="2"><input type="text" name="x_middle" size="25" maxlength=1 value="<%= Server.HtmlEncode(x_middle&"") %>"></font></td>
<td bgcolor="white" width="160" height="21"><font face="Arial" color="maroon"><span style="font-size:10pt;">Middle Initial:</span></font></td>
<td bgcolor="white" width="224" height="21"><font face="Arial" size="2"><input type="text" name="x_middle_co" size="25" maxlength=1 value="<%= Server.HtmlEncode(x_middle_co&"") %>"></font></td>
</tr>
<tr>
<td bgcolor="white" width="23">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="163"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">*Last Name:</font></SPAN></td>
<td bgcolor="white" width="227"><font face="Arial" size="2"><input type="text" name="x_last_name" size="25" maxlength=50 value="<%= Server.HtmlEncode(x_last_name&"") %>"></font></td>
<td bgcolor="white" width="160"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">*Last Name:</font></SPAN></td>
<td bgcolor="white" width="224"><font face="Arial" size="2"><input type="text" name="x_last_co" size="25" maxlength=50 value="<%= Server.HtmlEncode(x_last_co&"") %>"></font></td>
</tr>
<tr>
<td bgcolor="white" width="23">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="163"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">*Street:</font></SPAN></td>
<td bgcolor="white" width="227"><font face="Arial" size="2"><input type="text" name="x_street" size="25" maxlength=50 value="<%= Server.HtmlEncode(x_street&"") %>"></font></td>
<td bgcolor="white" width="160"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">*Street:</font></SPAN></td>
<td bgcolor="white" width="224"><font face="Arial" size="2"><input type="text" name="x_street_co" size="25" maxlength=50 value="<%= Server.HtmlEncode(x_street_co&"") %>"></font></td>
</tr>
<tr>
<td bgcolor="white" width="23">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="163"><font face="Arial"><span style="font-size:10pt;">Apartment:</span></font></td>
<td bgcolor="white" width="227"><font face="Arial" size="2"><input type="text" name="x_aparment" size="25" maxlength=50 value="<%= Server.HtmlEncode(x_aparment&"") %>"></font></td>
<td bgcolor="white" width="160"><font face="Arial"><span style="font-size:10pt;">Apartment:</span></font></td>
<td bgcolor="white" width="224"><font face="Arial" size="2"><input type="text" name="x_apartment_co" size="25" maxlength=50 value="<%= Server.HtmlEncode(x_apartment_co&"") %>"></font></td>
</tr>
<tr>
<td bgcolor="white" width="23">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="163"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">*City:</font></SPAN></td>
<td bgcolor="white" width="227"><font face="Arial" size="2"><input type="text" name="x_city" size="25" maxlength=50 value="<%= Server.HtmlEncode(x_city&"") %>"></font></td>
<td bgcolor="white" width="160"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">*City:</font></SPAN></td>
<td bgcolor="white" width="224"><font face="Arial" size="2"><input type="text" name="x_city_co" size="25" maxlength=50 value="<%= Server.HtmlEncode(x_city_co&"") %>"></font></td>
</tr>
<tr>
<td bgcolor="white" width="23">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="163"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">*State:</font></SPAN></td>
<td bgcolor="white" width="227"><font face="Arial" size="2"><input type="text" name="x_state" size="25" maxlength=50 value="<%= Server.HtmlEncode(x_state&"") %>"></font></td>
<td bgcolor="white" width="160"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">*State:</font></SPAN></td>
<td bgcolor="white" width="224"><font face="Arial" size="2"><input type="text" name="x_state_co" size="25" maxlength=50 value="<%= Server.HtmlEncode(x_state_co&"") %>"></font></td>
</tr>
<tr>
<td bgcolor="white" width="23">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="163"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">*Zip:</font></SPAN></td>
<td bgcolor="white" width="227"><font face="Arial" size="2"><input type="text" name="x_zip" size="25" maxlength=50 value="<%= Server.HtmlEncode(x_zip&"") %>"></font></td>
<td bgcolor="white" width="160"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">*Zip:</font></SPAN></td>
<td bgcolor="white" width="224"><font face="Arial" size="2"><input type="text" name="x_zip_co" size="25" maxlength=50 value="<%= Server.HtmlEncode(x_zip_co&"") %>"></font></td>
</tr>
<tr>
<td bgcolor="white" width="23">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="163"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">*Home</font><font face="Arial"> </font><font face="Arial" color="maroon">Phone:</font></SPAN></td>
<td bgcolor="white" width="227"><font face="Arial" size="2"><input type="text" name="x_home_phone" size="25" maxlength=50 value="<%= Server.HtmlEncode(x_home_phone&"") %>"></font></td>
<td bgcolor="white" width="160"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">*Home</font><font face="Arial"> </font><font face="Arial" color="maroon">Phone:</font></SPAN></td>
<td bgcolor="white" width="224"><font face="Arial" size="2"><input type="text" name="x_home_phone_co" size="25" maxlength=50 value="<%= Server.HtmlEncode(x_home_phone_co&"") %>"></font></td>
</tr>
<tr>
<td bgcolor="white" width="23">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="163"><font face="Arial"><span style="font-size:10pt;">Work Phone:</span></font></td>
<td bgcolor="white" width="227"><font face="Arial" size="2"><input type="text" name="x_work_phone" size="25" maxlength=50 value="<%= Server.HtmlEncode(x_work_phone&"") %>"></font></td>
<td bgcolor="white" width="160"><font face="Arial"><span style="font-size:10pt;">Work Phone:</span></font></td>
<td bgcolor="white" width="224"><font face="Arial" size="2"><input type="text" name="x_work_phone_co" size="25" maxlength=50 value="<%= Server.HtmlEncode(x_work_phone_co&"") %>"></font></td>
</tr>
<tr>
<td bgcolor="white" width="23">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="163"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">*Email</font><font face="Arial"> </font><font face="Arial" color="maroon">Address:</font></SPAN></td>
<td bgcolor="white" width="227"><font face="Arial" size="2"><input type="text" name="x_email" size="25" maxlength=50 value="<%= Server.HtmlEncode(x_email&"") %>"></font></td>
<td bgcolor="white" width="160"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">*Email</font><font face="Arial"> </font><font face="Arial" color="maroon">Address:</font></SPAN></td>
<td bgcolor="white" width="224"><font face="Arial" size="2"><input type="text" name="x_email_co" size="25" maxlength=50 value="<%= Server.HtmlEncode(x_email_co&"") %>"></font></td>
</tr>
<tr>
<td bgcolor="white" width="23">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="163"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">*Social Security Number:</font></SPAN></td>
<td bgcolor="white" width="227"><font face="Arial" size="2"><input type="text" name="x_ssn" size="25" maxlength=50 value="<%= Server.HtmlEncode(x_ssn&"") %>"></font></td>
<td bgcolor="white" width="160"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">*Social</font><font face="Arial"> </font><font face="Arial" color="maroon">Security</font><font face="Arial"> </font><font face="Arial" color="maroon">Number:</font></SPAN></td>
<td bgcolor="white" width="224"><font face="Arial" size="2"><input type="text" name="x_ssn_co" size="25" maxlength=50 value="<%= Server.HtmlEncode(x_ssn_co&"") %>"></font></td>
</tr>
<tr>
<td bgcolor="white" width="23">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="163"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">*Date</font><font face="Arial"> </font><font face="Arial" color="maroon">of</font><font face="Arial"> </font><font face="Arial" color="maroon">Birth:</font></SPAN><font face="Arial"><span style="font-size:10pt;"> </span><span style="font-size:8pt;">(mm/dd/yyyy)</span></font></td>
<td bgcolor="white" width="227"><font face="Arial" size="2"><input type="text" name="x_dob" size="25" maxlength=50 value="<% if isdate(x_dob) then response.write EW_FormatDateTime(x_dob,6) else response.write x_dob end if %>"></font></td>
<td bgcolor="white" width="160"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">*Date</font><font face="Arial"> </font><font face="Arial" color="maroon">of</font><font face="Arial"> </font><font face="Arial" color="maroon">Birth:</font></SPAN><font face="Arial"><span style="font-size:10pt;"> </span><span style="font-size:8pt;">(mm/dd/yyyy)</span></font></td>
<td bgcolor="white" width="224"><font face="Arial" size="2"><input type="text" name="x_dob_co" size="25" maxlength=50 value="<% if isdate(x_dob_co) then response.write EW_FormatDateTime(x_dob_co,6) else response.write x_dob_co end if %>"></font></td>
</tr>
<tr>
<td bgcolor="white" width="23">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="163"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">*Occupation:</font></SPAN></td>
<td bgcolor="white" width="227"><font face="Arial" size="2"><input type="text" name="x_occupation" size="25" maxlength=50 value="<%= Server.HtmlEncode(x_occupation&"") %>"></font></td>
<td bgcolor="white" width="160"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">*Occupation:</font></SPAN></td>
<td bgcolor="white" width="224"><font face="Arial" size="2"><input type="text" name="x_occupation_co" size="25" maxlength=50 value="<%= Server.HtmlEncode(x_occupation_co&"") %>"></font></td>
</tr>
<tr>
<td bgcolor="white" width="23">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="163"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">*Place of Employment:</font></SPAN></td>
<td bgcolor="white" width="227"><font face="Arial" size="2"><input type="text" name="x_workplace" size="25" maxlength=50 value="<%= Server.HtmlEncode(x_workplace&"") %>"></font></td>
<td bgcolor="white" width="160"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">*Place</font><font face="Arial"> </font><font face="Arial" color="maroon">of</font><font face="Arial"> </font><font face="Arial" color="maroon">Employment:</font></SPAN></td>
<td bgcolor="white" width="224"><font face="Arial" size="2"><input type="text" name="x_workplace_co" size="25" maxlength=50 value="<%= Server.HtmlEncode(x_workplace_co&"") %>"></font></td>
</tr>
<tr>
<td bgcolor="white" width="23">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="163"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">*Net Salary:</font></SPAN></td>
<td bgcolor="white" width="227"><font face="Arial" size="2"><input type="text" name="x_net_salary" size="25" maxlength=50 value="<%= Server.HtmlEncode(x_net_salary&"") %>"></font></td>
<td bgcolor="white" width="160"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">*Net</font><font face="Arial"> </font><font face="Arial" color="maroon">Salary:</font></SPAN></td>
<td bgcolor="white" width="224"><font face="Arial" size="2"><input type="text" name="x_net_salary_co" size="25" maxlength=50 value="<%= Server.HtmlEncode(x_net_salary_co&"") %>"></font></td>
</tr>
<tr>
<td bgcolor="white" width="23">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="163"><font face="Arial"><span style="font-size:10pt;">Time At Current Employer:</span></font></td>
<td bgcolor="white" width="227"><font face="Arial" size="2"><input type="text" name="x_timework" size="25" maxlength=50 value="<%= Server.HtmlEncode(x_timework&"") %>"></font></td>
<td bgcolor="white" width="160"><font face="Arial"><span style="font-size:10pt;">Time At Current Employer:</span></font></td>
<td bgcolor="white" width="224"><font face="Arial" size="2"><input type="text" name="x_timework_co" size="25" maxlength=50 value="<%= Server.HtmlEncode(x_timework_co&"") %>"></font></td>
</tr>
<tr>
<td bgcolor="white" width="23">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="774" colspan="4">&nbsp;<font face="Arial" size="2"><%= x_ID %><input type="hidden" name="x_ID" value="<%= x_ID %>"></font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="white" width="23">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="390" colspan="2">                            <p><b><span style="font-size:14pt;"><font face="Arial">Important Privacy Information:<br>&nbsp;</font></span></b></p>
</td>
<td bgcolor="white" width="160">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="224">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td bgcolor="white" width="23">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="163">&nbsp;</td>
<td bgcolor="white" width="227">
<p><SPAN class=required style="font-size:11pt;"><font face="Arial" color="maroon">* I Authorize:</font></SPAN><SPAN class=smalltext style="font-size:8pt;"><font face="Arial">(enter your initials)</font></SPAN></td>
<td bgcolor="white" width="160">
<font face="Arial" size="2"><input type="text" name="x_initials" size="25" maxlength=50 value="<%= Server.HtmlEncode(x_initials&"") %>"></font></td>
<td bgcolor="white" width="224">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td bgcolor="white" width="23">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="163">&nbsp;</td>
<td bgcolor="white" width="227">
<font face="Arial" color="maroon"><span style="font-size:11pt;">*Do 
                            you 
                            agree terms and conditiones?</span></font></td>
<td bgcolor="white" width="160">
<font face="Arial" size="2"><%
x_iagreeList = "<SELECT name='x_iagree'><OPTION value=''>Please Select</OPTION>"
    x_iagreeList = x_iagreeList & "<OPTION value=""Yes"""
    If x_iagree = "Yes" Then
        x_iagreeList = x_iagreeList & " selected"
    End If
    x_iagreeList = x_iagreeList & ">" & "Yes" & "</option>"
    x_iagreeList = x_iagreeList & "<OPTION value=""No"""
    If x_iagree = "No" Then
        x_iagreeList = x_iagreeList & " selected"
    End If
    x_iagreeList = x_iagreeList & ">" & "No" & "</option>"
x_iagreeList = x_iagreeList & "</select>"
response.write x_iagreeList
%></font></td>
<td bgcolor="white" width="224">
                            <p>&nbsp;</p>
</td>
</tr>
</table>


<p align="center">
<input type="submit" name="Action" value="EDIT">
</form>
                        <p>&nbsp;</p>
</td>
</tr>
<tr>
<td bgcolor="white" width="398">
                            <p><span style="font-size:14pt;"><b><font face="Arial">&nbsp;&nbsp;&nbsp;<a href="cocreditlist.asp"><img src="images/back.gif" align="middle" width="16" height="16" border="0"></a> 
                             
</font></b></span><a href="cocreditlist.asp"><font face="Arial" size="2" color="black"><b>Back to Co-Applicant 
List</b></font></a><font face="Arial" size="2" color="black"><b><br>&nbsp;</b></font></p>
</td>
<td bgcolor="white" width="160">
                            <p>&nbsp;<font face="Arial" size="2" color="white"><%= x_stock %><input type="hidden" name="x_stock" value="<%= x_stock %>"></font><font color="white">&nbsp;</font></p>
</td>
<td bgcolor="white" width="224">
                            <p>&nbsp;</p>
</td>
</tr>
</table>

</form>
        </td>
    </tr>
</table>
<!--#include file="footer.asp"-->
