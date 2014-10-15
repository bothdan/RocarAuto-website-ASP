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
If key="" OR IsNull(key) Then Response.Redirect "creditlist.asp"
'get action
a=Request.Form("a")
If a="" OR IsNull(a) Then
	a="I"	'display with input box
End If
'get fields from form
x_ID = Request.Form("x_ID")
x_email = Request.Form("x_email")
x_first = Request.Form("x_first")
x_middle = Request.Form("x_middle")
x_last = Request.Form("x_last")
x_street = Request.Form("x_street")
x_apartment = Request.Form("x_apartment")
x_city = Request.Form("x_city")
x_state = Request.Form("x_state")
x_zip = Request.Form("x_zip")
x_home_phone = Request.Form("x_home_phone")
x_ssn = Request.Form("x_ssn")
x_dob = Request.Form("x_dob")
x_workplace = Request.Form("x_workplace")
x_occupation = Request.Form("x_occupation")
x_work_street = Request.Form("x_work_street")
x_work_city = Request.Form("x_work_city")
x_work_state = Request.Form("x_work_state")
x_work_zip = Request.Form("x_work_zip")
x_work_phone = Request.Form("x_work_phone")
x_worktime = Request.Form("x_worktime")
x_net_salary = Request.Form("x_net_salary")
x_other_income = Request.Form("x_other_income")
x_initials = Request.Form("x_initials")
x_iagree = Request.Form("x_iagree")
x_stock = Request.Form("x_stock")
' Open Connection to the database
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str
Select Case a
	Case "I": ' Get a record to display
		tkey = "" & key & ""
		strsql = "SELECT * FROM [credit] WHERE [ID]=" & tkey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn
		If rs.EOF Then
			Response.Clear
			Response.Redirect "creditlist.asp"
		Else
			rs.MoveFirst
		End If
		' Get the field contents
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
		rs.Close
		Set rs = Nothing
	Case "U": ' Update
		' Open record
		tkey = "" & key & ""
		strsql = "SELECT * FROM [credit] WHERE [ID]=" & tkey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn, 1, 2
		If rs.EOF Then
			Response.Clear
			Response.Redirect "creditlist.asp"
		End If
		tmpFld = Trim(x_email)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("email") = tmpFld
		tmpFld = Trim(x_first)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("first") = tmpFld
		tmpFld = Trim(x_middle)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("middle") = tmpFld
		tmpFld = Trim(x_last)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("last") = tmpFld
		tmpFld = Trim(x_street)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("street") = tmpFld
		tmpFld = Trim(x_apartment)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("apartment") = tmpFld
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
		tmpFld = Trim(x_ssn)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("ssn") = tmpFld
		tmpFld = Trim(x_dob)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("dob") = tmpFld
		tmpFld = Trim(x_workplace)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("workplace") = tmpFld
		tmpFld = Trim(x_occupation)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("occupation") = tmpFld
		tmpFld = Trim(x_work_street)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("work street") = tmpFld
		tmpFld = Trim(x_work_city)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("work city") = tmpFld
		tmpFld = Trim(x_work_state)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("work state") = tmpFld
		tmpFld = Trim(x_work_zip)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("work zip") = tmpFld
		tmpFld = Trim(x_work_phone)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("work phone") = tmpFld
		tmpFld = Trim(x_worktime)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("worktime") = tmpFld
		tmpFld = Trim(x_net_salary)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("net salary") = tmpFld
		tmpFld = Trim(x_other_income)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("other income") = tmpFld
		tmpFld = Trim(x_initials)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("initials") = tmpFld
		tmpFld = Trim(x_iagree)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("iagree") = tmpFld
		tmpFld = Trim(x_stock)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("stock") = tmpFld
		rs.Update
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
        <td background="images/contactbg.gif">
            <p><img src="images/financebg.gif" width="312" height="32" border="0"></p>
        </td>
    </tr>
    <tr>
        <td>
            <p><font face="Arial" size="2"><br></font><span style="font-size:14pt;"><b><font face="Arial">&nbsp;&nbsp;<a href="creditlist.asp"><img src="images/back.gif" align="middle" width="16" height="16" border="0"></a> 
</font></b></span><a href="creditlist.asp"><font face="Arial" size="2" color="black"><b>Back to Credit 
List</b></font></a><span style="font-size:14pt;"><b><font face="Arial">&nbsp;</font></b></span></p>
<script language="JavaScript" src="ew.js"></script>
<script language="JavaScript">
<!-- start Javascript
function  EW_checkMyForm(EW_this) {
return true;
}
// end JavaScript -->
</script>
<form onSubmit="return EW_checkMyForm(this);"  action="creditedit.asp" method="post">
<p>
<input type="hidden" name="a" value="U">
<input type="hidden" name="key" value="<%= key %>">
<table border="0" cellspacing="0" cellpadding="0" bgcolor="#CCCCCC" width="737">
<tr>
<td width="290" bgcolor="white" height="67" align="left" valign="top" colspan="2">
                <p></p>

<p><span style="font-size:14pt;"><b><font face="Arial">&nbsp;&nbsp;&nbsp;Application Information:</font></b></span></td>
<td width="447" bgcolor="white" height="61" colspan="2">
                <p align="left"><font face="Arial" color="white"><span style="font-size:12pt;"><%= x_ID %><input type="hidden" name="x_ID" value="<%= x_ID %>">&nbsp;</span></font><SPAN class=required style="font-size:8pt;"><font color="maroon" face="Arial">.</font></SPAN></td>
</tr>
<tr>
<td width="119" bgcolor="white" height="568" rowspan="22" align="left" valign="top">
                <p>&nbsp;</p>
</td>
<td width="171" bgcolor="white" height="24" valign="bottom">
                            <p><font face="Arial" color="maroon"><span style="font-size:10pt;">*First 
                            Name:</span></font></p>
</td>
<td width="278" bgcolor="white" height="24" valign="bottom"><font face="Arial"><span style="font-size:12pt;"><input type="text" name="x_first" size=30 maxlength=50 value="<%= Server.HtmlEncode(x_first&"") %>"></span></font></td>
<td width="169" bgcolor="white" height="24">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td width="171" bgcolor="white" height="24"><font face="Arial"><span style="font-size:10pt;">Middle Initial:</span></font></td>
<td width="278" bgcolor="white" height="24"><font face="Arial"><span style="font-size:12pt;"><input type="text" name="x_middle" size=30 maxlength=50 value="<%= Server.HtmlEncode(x_middle&"") %>"></span></font></td>
<td width="169" bgcolor="white" height="24">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td width="171" bgcolor="white"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">* Last Name:</font></SPAN></td>
<td width="278" bgcolor="white"><font face="Arial"><span style="font-size:12pt;"><input type="text" name="x_last" size=30 maxlength=50 value="<%= Server.HtmlEncode(x_last&"") %>"></span></font></td>
<td width="169" bgcolor="white">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td width="171" bgcolor="white"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">* Street:</font></SPAN></td>
<td width="278" bgcolor="white"><font face="Arial"><span style="font-size:12pt;"><input type="text" name="x_street" size=30 maxlength=50 value="<%= Server.HtmlEncode(x_street&"") %>"></span></font></td>
<td width="169" bgcolor="white">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td width="171" bgcolor="white"><font face="Arial"><span style="font-size:10pt;">Apartment:</span></font></td>
<td width="278" bgcolor="white"><font face="Arial"><span style="font-size:12pt;"><input type="text" name="x_apartment" size=30 maxlength=50 value="<%= Server.HtmlEncode(x_apartment&"") %>"></span></font></td>
<td width="169" bgcolor="white">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td width="171" bgcolor="white"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">* City:</font></SPAN></td>
<td width="278" bgcolor="white"><font face="Arial"><span style="font-size:12pt;"><input type="text" name="x_city" size=30 maxlength=50 value="<%= Server.HtmlEncode(x_city&"") %>"></span></font></td>
<td width="169" bgcolor="white">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td width="171" bgcolor="white"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">* State:</font></SPAN></td>
<td width="278" bgcolor="white"><font face="Arial"><span style="font-size:12pt;"><input type="text" name="x_state" size=30 maxlength=50 value="<%= Server.HtmlEncode(x_state&"") %>"></span></font></td>
<td width="169" bgcolor="white">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td width="171" bgcolor="white"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">* Zip:</font></SPAN></td>
<td width="278" bgcolor="white"><font face="Arial"><span style="font-size:12pt;"><input type="text" name="x_zip" size=30 maxlength=50 value="<%= Server.HtmlEncode(x_zip&"") %>"></span></font></td>
<td width="169" bgcolor="white">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td width="171" bgcolor="white"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">* Home Phone:</font></SPAN></td>
<td width="278" bgcolor="white"><font face="Arial"><span style="font-size:12pt;"><input type="text" name="x_home_phone" size=30 maxlength=50 value="<%= Server.HtmlEncode(x_home_phone&"") %>"></span></font></td>
<td width="169" bgcolor="white">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td width="171" bgcolor="white"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">* Email Address:</font></SPAN></td>
<td width="278" bgcolor="white"><font face="Arial"><span style="font-size:12pt;"><input type="text" name="x_email" size=30 maxlength=50 value="<%= Server.HtmlEncode(x_email&"") %>"></span></font></td>
<td width="169" bgcolor="white">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td width="171" bgcolor="white" height="24"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">* Social Security Number:</font></SPAN></td>
<td width="278" bgcolor="white" height="24"><font face="Arial"><span style="font-size:12pt;"><input type="text" name="x_ssn" size=30 maxlength=50 value="<%= Server.HtmlEncode(x_ssn&"") %>"></span></font></td>
<td width="169" bgcolor="white" height="24">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td width="171" bgcolor="white"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">* Date of Birth</font></SPAN><SPAN class=required style="font-size:11pt;"><font face="Arial" color="maroon">:</font></SPAN><font face="Arial"><span style="font-size:11pt;"> </span><span style="font-size:8pt;">(mm/dd/yyyy)</span></font></td>
<td width="278" bgcolor="white"><font face="Arial"><span style="font-size:12pt;"><input type="text" name="x_dob" size=30 maxlength=50 value="<% if isdate(x_dob) then response.write EW_FormatDateTime(x_dob,6) else response.write x_dob end if %>">&nbsp;</span></font></td>
<td width="169" bgcolor="white">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td width="171" bgcolor="white"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">* Occupation:</font></SPAN></td>
<td width="278" bgcolor="white"><font face="Arial"><span style="font-size:12pt;"><input type="text" name="x_occupation" size=30 maxlength=50 value="<%= Server.HtmlEncode(x_occupation&"") %>"></span></font></td>
<td width="169" bgcolor="white">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td width="171" bgcolor="white"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">* Place of Employment:</font></SPAN></td>
<td width="278" bgcolor="white"><font face="Arial"><span style="font-size:12pt;"><input type="text" name="x_workplace" size=30 maxlength=50 value="<%= Server.HtmlEncode(x_workplace&"") %>"></span></font></td>
<td width="169" bgcolor="white">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td width="171" bgcolor="white"><span style="font-size:10pt;"><font face="Arial">Work&nbsp;Address:</font></span></td>
<td width="278" bgcolor="white"><font face="Arial"><span style="font-size:12pt;"><input type="text" name="x_work_street" size=30 maxlength=50 value="<%= Server.HtmlEncode(x_work_street&"") %>">&nbsp;</span></font></td>
<td width="169" bgcolor="white">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td width="171" bgcolor="white"><SPAN class=required style="font-size:10pt;"><font face="Arial">City:</font></SPAN></td>
<td width="278" bgcolor="white"><font face="Arial"><span style="font-size:12pt;"><input type="text" name="x_work_city" size=30 maxlength=50 value="<%= Server.HtmlEncode(x_work_city&"") %>">&nbsp;</span></font></td>
<td width="169" bgcolor="white">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td width="171" bgcolor="white"><span style="font-size:10pt;"><font face="Arial">State:</font></span></td>
<td width="278" bgcolor="white"><font face="Arial"><span style="font-size:12pt;"><input type="text" name="x_work_state" size=30 maxlength=50 value="<%= Server.HtmlEncode(x_work_state&"") %>">&nbsp;</span></font></td>
<td width="169" bgcolor="white">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td width="171" bgcolor="white"> <font face="Arial"><span style="font-size:10pt;">Zip:</span></font></td>
<td width="278" bgcolor="white"><font face="Arial"><span style="font-size:12pt;"><input type="text" name="x_work_zip" size=30 maxlength=50 value="<%= Server.HtmlEncode(x_work_zip&"") %>">&nbsp;</span></font></td>
<td width="169" bgcolor="white">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td width="171" bgcolor="white"><font face="Arial"><span style="font-size:10pt;">Work Phone&nbsp;</span></font></td>
<td width="278" bgcolor="white"><font face="Arial"><span style="font-size:12pt;"><input type="text" name="x_work_phone" size=30 maxlength=50 value="<%= Server.HtmlEncode(x_work_phone&"") %>">&nbsp;</span></font></td>
<td width="169" bgcolor="white">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td width="171" bgcolor="white"><font face="Arial"><span style="font-size:10pt;">Time At Current Employer:</span></font></td>
<td width="278" bgcolor="white"><font face="Arial"><span style="font-size:12pt;"><input type="text" name="x_worktime" size=30 maxlength=50 value="<%= Server.HtmlEncode(x_worktime&"") %>">&nbsp;</span></font></td>
<td width="169" bgcolor="white">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td width="171" bgcolor="white"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">* Net Salary:</font></SPAN></td>
<td width="278" bgcolor="white"><font face="Arial"><span style="font-size:12pt;"><input type="text" name="x_net_salary" size=30 maxlength=50 value="<%= Server.HtmlEncode(x_net_salary&"") %>">&nbsp;</span></font></td>
<td width="169" bgcolor="white">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td width="171" bgcolor="white"><font face="Arial"><span style="font-size:10pt;">Other Income:</span></font></td>
<td width="278" bgcolor="white"><font face="Arial"><span style="font-size:12pt;"><input type="text" name="x_other_income" size=30 maxlength=50 value="<%= Server.HtmlEncode(x_other_income&"") %>">&nbsp;</span></font></td>
<td width="169" bgcolor="white">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td width="568" bgcolor="white" colspan="3">
                            <p><br> &nbsp;&nbsp;&nbsp;<b><span style="font-size:14pt;"><font face="Arial">Important Privacy Information:<br>&nbsp;</font></span></b></p>
</td>
<td width="169" bgcolor="white">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td width="119" bgcolor="white">
                            <p>&nbsp;</p>
</td>
<td width="618" bgcolor="white" colspan="3">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td width="119" bgcolor="white">
                            <p><font face="Arial"><span style="font-size:10pt;">&nbsp;</span></font></p>
</td>
<td width="171" bgcolor="white"><SPAN class=required style="font-size:11pt;"><font face="Arial" color="maroon">* I Authorize:</font></SPAN><SPAN class=smalltext style="font-size:8pt;"><font face="Arial">(enter your initials)</font></SPAN></td>
<td width="278" bgcolor="white"><font face="Arial"><span style="font-size:10pt;"><input type="text" name="x_initials" size=30 maxlength=50 value="<%= Server.HtmlEncode(x_initials&"") %>">&nbsp;</span></font></td>
<td width="169" bgcolor="white">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td width="119" bgcolor="white">
                            <p><font face="Arial"><span style="font-size:10pt;">&nbsp;</span></font></p>
</td>
<td width="171" bgcolor="white"><font face="Arial" color="maroon"><span style="font-size:11pt;">*Do 
                            you 
                            agree terms and conditiones?</span></font></td>
<td width="278" bgcolor="white"><font face="Arial"><span style="font-size:10pt;"><%
x_iagreeList = "<SELECT name='x_iagree'><OPTION value=''>Please Select</OPTION>"
    x_iagreeList = x_iagreeList & "<OPTION value=""Yes"""
    If x_iagree = "Yes" Then
        x_iagreeList = x_iagreeList & " selected"
    End If
    x_iagreeList = x_iagreeList & ">" & "I Agree" & "</option>"
    x_iagreeList = x_iagreeList & "<OPTION value=""No"""
    If x_iagree = "No" Then
        x_iagreeList = x_iagreeList & " selected"
    End If
    x_iagreeList = x_iagreeList & ">" & "No" & "</option>"
x_iagreeList = x_iagreeList & "</select>"
response.write x_iagreeList
%>
&nbsp;</span></font></td>
<td width="169" bgcolor="white">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td width="119" bgcolor="white">
                            <p><font face="Arial"><span style="font-size:10pt;">&nbsp;</span></font></p>
</td>
<td width="171" bgcolor="white"><font face="Arial"><span style="font-size:10pt;">&nbsp;</span></font></td>
<td width="278" bgcolor="white"><font face="Arial" color="white"><span style="font-size:10pt;"><%= x_stock%>&nbsp;<input type="hidden" name="x_stock" value="<%= x_stock %>">&nbsp;</span></font></td>
<td width="169" bgcolor="white">
                            <p>&nbsp;</p>
</td>
</tr>
</table>
                <p align="center"><input type="submit" name="Action" value="Update">
</form>

<p><span style="font-size:14pt;"><b><font face="Arial">&nbsp;&nbsp;<a href="creditlist.asp"><img src="images/back.gif" align="middle" width="16" height="16" border="0"></a> 
</font></b></span><a href="creditlist.asp"><font face="Arial" size="2" color="black"><b>Back to Credit 
List</b></font></a><span style="font-size:14pt;"><b><font face="Arial">&nbsp;</font></b></span></p>
        </td>
    </tr>
</table>
<!--#include file="footer.asp"-->
