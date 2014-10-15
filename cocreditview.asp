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
If key = "" OR IsNull(key) Then Response.Redirect "cocreditlist.asp"
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
End Select
%>
<!--#include file="header.asp"-->
<meta name="generator" content="Namo WebEditor v5.0(Trial)">
<table align="center" cellpadding="0" cellspacing="0" width="801" bgcolor="white">
    <tr>
        <td>

<form onSubmit="return EW_checkMyForm(this);"  action="cocreditadd.asp" method="post">

<table border="0" cellspacing="0" cellpadding="0" bgcolor="#CCCCCC" width="797">
<tr>
<td bgcolor="white" width="413" colspan="3">
                            <p><span style="font-size:14pt;"><b><font face="Arial"><br>&nbsp;&nbsp;&nbsp;<a href="cocreditlist.asp"><img src="images/back.gif" align="middle" width="16" height="16" border="0"></a> 
                            </font></b></span><a href="cocreditlist.asp"><font face="Arial" size="2" color="black"><b>Back to Co-Applicant 
                            List</b></font></a><span style="font-size:14pt;"><b><font face="Arial">&nbsp;<br>&nbsp;</font></b></span></p>
</td>
<td bgcolor="white" width="384" colspan="2">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td bgcolor="white" width="23">
                            <p>&nbsp;</p>
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
<td bgcolor="white" width="227"><font face="Verdana" size="2" color="black"><% response.write x_first_name %></font></td>
<td bgcolor="white" width="160">
                            <p class="tblFinanceForm coapp"><font face="Arial" color="maroon"><span style="font-size:10pt;">*First 
                            Name:</span></font></p>
</td>
<td bgcolor="white" width="224"><font face="Verdana" size="2" color="black"><% response.write x_first_co %></font></td>
</tr>
<tr>
<td bgcolor="white" width="23" height="21">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="163" height="21"><font face="Arial"><span style="font-size:10pt;">Middle Initial:</span></font></td>
<td bgcolor="white" width="227" height="21"><font face="Verdana" size="2" color="black"><% response.write x_middle %></font></td>
<td bgcolor="white" width="160" height="21"><font face="Arial" color="maroon"><span style="font-size:10pt;">Middle Initial:</span></font></td>
<td bgcolor="white" width="224" height="21"><font face="Verdana" size="2" color="black"><% response.write x_middle_co %></font></td>
</tr>
<tr>
<td bgcolor="white" width="23">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="163"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">*Last Name:</font></SPAN></td>
<td bgcolor="white" width="227"><font face="Verdana" size="2" color="black"><% response.write x_last_name %></font></td>
<td bgcolor="white" width="160"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">*Last Name:</font></SPAN></td>
<td bgcolor="white" width="224"><font face="Verdana" size="2" color="black"><% response.write x_last_co %></font></td>
</tr>
<tr>
<td bgcolor="white" width="23">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="163"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">*Street:</font></SPAN></td>
<td bgcolor="white" width="227"><font face="Verdana" size="2" color="black"><% response.write x_street %></font></td>
<td bgcolor="white" width="160"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">*Street:</font></SPAN></td>
<td bgcolor="white" width="224"><font face="Verdana" size="2" color="black"><% response.write x_street_co %></font></td>
</tr>
<tr>
<td bgcolor="white" width="23">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="163"><font face="Arial"><span style="font-size:10pt;">Apartment:</span></font></td>
<td bgcolor="white" width="227"><font face="Verdana" size="2" color="black"><% response.write x_aparment %></font></td>
<td bgcolor="white" width="160"><font face="Arial"><span style="font-size:10pt;">Apartment:</span></font></td>
<td bgcolor="white" width="224"><font face="Verdana" size="2" color="black"><% response.write x_apartment_co %></font></td>
</tr>
<tr>
<td bgcolor="white" width="23">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="163"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">*City:</font></SPAN></td>
<td bgcolor="white" width="227"><font face="Verdana" size="2" color="black"><% response.write x_city %></font></td>
<td bgcolor="white" width="160"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">*City:</font></SPAN></td>
<td bgcolor="white" width="224"><font face="Verdana" size="2" color="black"><% response.write x_city_co %></font></td>
</tr>
<tr>
<td bgcolor="white" width="23">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="163"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">*State:</font></SPAN></td>
<td bgcolor="white" width="227"><font face="Verdana" size="2" color="black"><% response.write x_state %></font></td>
<td bgcolor="white" width="160"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">*State:</font></SPAN></td>
<td bgcolor="white" width="224"><font face="Verdana" size="2" color="black"><% response.write x_state_co %></font></td>
</tr>
<tr>
<td bgcolor="white" width="23">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="163"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">*Zip:</font></SPAN></td>
<td bgcolor="white" width="227"><font face="Verdana" size="2" color="black"><% response.write x_zip %></font></td>
<td bgcolor="white" width="160"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">*Zip:</font></SPAN></td>
<td bgcolor="white" width="224"><font face="Verdana" size="2" color="black"><% response.write x_zip_co %></font></td>
</tr>
<tr>
<td bgcolor="white" width="23">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="163"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">*Home</font><font face="Arial"> </font><font face="Arial" color="maroon">Phone:</font></SPAN></td>
<td bgcolor="white" width="227"><font face="Verdana" size="2" color="black"><% response.write x_home_phone %></font></td>
<td bgcolor="white" width="160"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">*Home</font><font face="Arial"> </font><font face="Arial" color="maroon">Phone:</font></SPAN></td>
<td bgcolor="white" width="224"><font face="Verdana" size="2" color="black"><% response.write x_home_phone_co %></font></td>
</tr>
<tr>
<td bgcolor="white" width="23">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="163"><font face="Arial"><span style="font-size:10pt;">Work Phone:</span></font></td>
<td bgcolor="white" width="227"><font face="Verdana" size="2" color="black"><% response.write x_work_phone %></font></td>
<td bgcolor="white" width="160"><font face="Arial"><span style="font-size:10pt;">Work Phone:</span></font></td>
<td bgcolor="white" width="224"><font face="Verdana" size="2" color="black"><% response.write x_work_phone_co %></font></td>
</tr>
<tr>
<td bgcolor="white" width="23">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="163"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">*Email</font><font face="Arial"> </font><font face="Arial" color="maroon">Address:</font></SPAN></td>
<td bgcolor="white" width="227"><font face="Verdana" size="2" color="black"><% response.write x_email %></font></td>
<td bgcolor="white" width="160"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">*Email</font><font face="Arial"> </font><font face="Arial" color="maroon">Address:</font></SPAN></td>
<td bgcolor="white" width="224"><font face="Verdana" size="2" color="black"><% response.write x_email_co %></font></td>
</tr>
<tr>
<td bgcolor="white" width="23">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="163"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">*Social Security Number:</font></SPAN></td>
<td bgcolor="white" width="227"><font face="Verdana" size="2" color="black"><% response.write x_ssn %></font></td>
<td bgcolor="white" width="160"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">*Social</font><font face="Arial"> </font><font face="Arial" color="maroon">Security</font><font face="Arial"> </font><font face="Arial" color="maroon">Number:</font></SPAN></td>
<td bgcolor="white" width="224"><font face="Verdana" size="2" color="black"><% response.write x_ssn_co %></font></td>
</tr>
<tr>
<td bgcolor="white" width="23">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="163"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">*Date</font><font face="Arial"> </font><font face="Arial" color="maroon">of</font><font face="Arial"> </font><font face="Arial" color="maroon">Birth:</font></SPAN><font face="Arial"><span style="font-size:10pt;"> </span><span style="font-size:8pt;">(mm/dd/yyyy)</span></font></td>
<td bgcolor="white" width="227"><font face="Verdana" size="2" color="black"><% if isdate(x_dob) then response.write EW_FormatDateTime(x_dob,6) else response.write x_dob end if %></font></td>
<td bgcolor="white" width="160"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">*Date</font><font face="Arial"> </font><font face="Arial" color="maroon">of</font><font face="Arial"> </font><font face="Arial" color="maroon">Birth:</font></SPAN><font face="Arial"><span style="font-size:10pt;"> </span><span style="font-size:8pt;">(mm/dd/yyyy)</span></font></td>
<td bgcolor="white" width="224"><font face="Verdana" size="2" color="black"><% if isdate(x_dob_co) then response.write EW_FormatDateTime(x_dob_co,6) else response.write x_dob_co end if %></font></td>
</tr>
<tr>
<td bgcolor="white" width="23">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="163"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">*Occupation:</font></SPAN></td>
<td bgcolor="white" width="227"><font face="Verdana" size="2" color="black"><% response.write x_occupation %></font></td>
<td bgcolor="white" width="160"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">*Occupation:</font></SPAN></td>
<td bgcolor="white" width="224"><font face="Verdana" size="2" color="black"><% response.write x_occupation_co %></font></td>
</tr>
<tr>
<td bgcolor="white" width="23">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="163"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">*Place of Employment:</font></SPAN></td>
<td bgcolor="white" width="227"><font face="Verdana" size="2" color="black"><% response.write x_workplace %></font></td>
<td bgcolor="white" width="160"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">*Place</font><font face="Arial"> </font><font face="Arial" color="maroon">of</font><font face="Arial"> </font><font face="Arial" color="maroon">Employment:</font></SPAN></td>
<td bgcolor="white" width="224"><font face="Verdana" size="2" color="black"><% response.write x_workplace_co %></font></td>
</tr>
<tr>
<td bgcolor="white" width="23">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="163"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">*Net Salary:</font></SPAN></td>
<td bgcolor="white" width="227"><font face="Verdana" size="2" color="black"><% if isnumeric(x_net_salary) then response.write formatcurrency(x_net_salary,0,-2,-2,-2) else response.write x_net_salary end if %></font></td>
<td bgcolor="white" width="160"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">*Net</font><font face="Arial"> </font><font face="Arial" color="maroon">Salary:</font></SPAN></td>
<td bgcolor="white" width="224"><font face="Verdana" size="2" color="black"><% if isnumeric(x_net_salary_co) then response.write formatcurrency(x_net_salary_co,0,-2,-2,-2) else response.write x_net_salary_co end if %></font></td>
</tr>
<tr>
<td bgcolor="white" width="23">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="163"><font face="Arial"><span style="font-size:10pt;">Time At Current Employer:</span></font></td>
<td bgcolor="white" width="227"><font face="Verdana" size="2" color="black"><% response.write x_timework %></font></td>
<td bgcolor="white" width="160"><font face="Arial"><span style="font-size:10pt;">Time At Current Employer:</span></font></td>
<td bgcolor="white" width="224"><font color="black" size="2" face="Verdana">&nbsp;</font></td>
</tr>
<tr>
<td bgcolor="white" width="23">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="774" colspan="4"><font color="white">&nbsp;</font><font face="Arial" size="2" color="white"><%= x_ID %><input type="hidden" name="x_ID" value="<%= x_ID %>"></font><font color="white">&nbsp;</font></td>
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
<td bgcolor="white" width="227">&nbsp;<SPAN class=required style="font-size:11pt;"><font face="Arial" color="maroon">* I Authorize:</font></SPAN><SPAN class=smalltext style="font-size:8pt;"><font face="Arial">(enter your initials)</font></SPAN></td>
<td bgcolor="white" width="160">
<font face="Verdana" size="2" color="black"><% response.write x_initials%></font></td>
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
&nbsp;<font face="Arial" color="maroon"><span style="font-size:11pt;">*Do 
                            you 
                            agree terms and conditiones?</span></font></td>
<td bgcolor="white" width="160" valign="top">
<font face="Verdana" size="2"><%response.write x_iagree
%></font></td>
<td bgcolor="white" width="224">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td bgcolor="white" width="413" colspan="3">
                            <p><span style="font-size:14pt;"><b><font face="Arial"><br>&nbsp;&nbsp;&nbsp;<a href="cocreditlist.asp"><img src="images/back.gif" align="middle" width="16" height="16" border="0"></a> 
                            </font></b></span><a href="cocreditlist.asp"><font face="Arial" size="2" color="black"><b>Back to Co-Applicant 
                            List</b></font></a><span style="font-size:14pt;"><b><font face="Arial">&nbsp;</font></b></span><font face="Arial" size="2" color="white"><%= x_stock %><input type="hidden" name="x_stock" value="<%= x_stock %>"><br>&nbsp;</font></p>
</td>
<td bgcolor="white" width="160">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="224">
                            <p>&nbsp;</p>
</td>
</tr>
</table>

</form>
        </td>
    </tr>
    <tr>
        <td>
            
        </td>
    </tr>
</table>
<!--#include file="footer.asp"-->
