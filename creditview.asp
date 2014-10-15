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
If key = "" OR IsNull(key) Then Response.Redirect "creditlist.asp"
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
<form onSubmit="return EW_checkMyForm(this);"  action="creditedit.asp" method="post">

<table border="0" cellspacing="0" cellpadding="0" bgcolor="#CCCCCC" width="737">
<tr>
<td width="290" bgcolor="white" height="67" align="left" valign="top" colspan="2">
                <p></p>

<p><span style="font-size:14pt;"><b><font face="Arial"><br>&nbsp;&nbsp;&nbsp;Application Information:</font></b></span>
                            <p><span style="font-size:14pt;"><b><font face="Arial">&nbsp;&nbsp;<a href="creditlist.asp"><img src="images/back.gif" align="middle" width="16" height="16" border="0"></a> 
                            </font></b></span><a href="creditlist.asp"><font face="Arial" size="2" color="black"><b>Back to Credit 
                            List</b></font></a><span style="font-size:14pt;"><b><font face="Arial">&nbsp;</font></b></span></p>
</td>
<td width="447" bgcolor="white" height="61" colspan="2">
                <p align="left"><font face="Arial" color="white"><span style="font-size:12pt;"><%= x_ID %><input type="hidden" name="x_ID" value="<%= x_ID %>">&nbsp;</span></font><SPAN class=required style="font-size:8pt;"><font color="maroon" face="Arial">.</font></SPAN></td>
</tr>
<tr>
<td width="119" bgcolor="white" height="568" rowspan="22" align="left" valign="top">
                <p></p>
</td>
<td width="171" bgcolor="white" height="24" valign="bottom">
                            <p><font face="Arial" color="maroon"><span style="font-size:10pt;">*First 
                            Name:</span></font></p>
</td>
<td width="278" bgcolor="white" height="24" valign="bottom">
<p><font face="Verdana" size="2" color="black"><% response.write x_first %></font></td>
<td width="169" bgcolor="white" height="24">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td width="171" bgcolor="white" height="24"><font face="Arial"><span style="font-size:10pt;">Middle Initial:</span></font></td>
<td width="278" bgcolor="white" height="24">
<p><font face="Verdana" size="2" color="black"><% response.write x_middle %></font></td>
<td width="169" bgcolor="white" height="24">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td width="171" bgcolor="white"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">* Last Name:</font></SPAN></td>
<td width="278" bgcolor="white">
<p><font face="Verdana" size="2" color="black"><% response.write x_last%></font></td>
<td width="169" bgcolor="white">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td width="171" bgcolor="white"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">* Street:</font></SPAN></td>
<td width="278" bgcolor="white">
<p><font face="Verdana" size="2" color="black"><% response.write x_street%></font></td>
<td width="169" bgcolor="white">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td width="171" bgcolor="white"><font face="Arial"><span style="font-size:10pt;">Apartment:</span></font></td>
<td width="278" bgcolor="white">
<p><font face="Verdana" size="2" color="black"><% response.write x_apartment%></font></td>
<td width="169" bgcolor="white">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td width="171" bgcolor="white"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">* City:</font></SPAN></td>
<td width="278" bgcolor="white">
<p><font face="Verdana" size="2" color="black"><% response.write x_city%></font></td>
<td width="169" bgcolor="white">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td width="171" bgcolor="white"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">* State:</font></SPAN></td>
<td width="278" bgcolor="white">
<p><font face="Verdana" size="2" color="black"><% response.write x_state%></font></td>
<td width="169" bgcolor="white">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td width="171" bgcolor="white"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">* Zip:</font></SPAN></td>
<td width="278" bgcolor="white">
<p><font face="Verdana" size="2" color="black"><% response.write x_zip%></font></td>
<td width="169" bgcolor="white">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td width="171" bgcolor="white"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">* Home Phone:</font></SPAN></td>
<td width="278" bgcolor="white">
<p><font face="Verdana" size="2" color="black"><% response.write x_home_phone%></font></td>
<td width="169" bgcolor="white">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td width="171" bgcolor="white"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">* Email Address:</font></SPAN></td>
<td width="278" bgcolor="white">
<p><font face="Verdana" size="2" color="black"><% response.write x_email %></font></td>
<td width="169" bgcolor="white">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td width="171" bgcolor="white" height="24"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">* Social Security Number:</font></SPAN></td>
<td width="278" bgcolor="white" height="24">
<p><font face="Verdana" size="2" color="black"><% response.write x_ssn%></font></td>
<td width="169" bgcolor="white" height="24">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td width="171" bgcolor="white"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">* Date of Birth</font></SPAN><SPAN class=required style="font-size:11pt;"><font face="Arial" color="maroon">:</font></SPAN><font face="Arial"><span style="font-size:11pt;"> </span><span style="font-size:8pt;">(mm/dd/yyyy)</span></font></td>
<td width="278" bgcolor="white">
<p><font face="Verdana" size="2" color="black"><% if isdate(x_dob) then response.write EW_FormatDateTime(x_dob,6) else response.write x_dob end if %></font></td>
<td width="169" bgcolor="white">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td width="171" bgcolor="white"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">* Occupation:</font></SPAN></td>
<td width="278" bgcolor="white">
<p><font face="Verdana" size="2" color="black"><% response.write x_occupation %></font></td>
<td width="169" bgcolor="white">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td width="171" bgcolor="white"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">* Place of Employment:</font></SPAN></td>
<td width="278" bgcolor="white">
<p><font face="Verdana" size="2" color="black"><% response.write x_workplace%></font></td>
<td width="169" bgcolor="white">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td width="171" bgcolor="white"><span style="font-size:10pt;"><font face="Arial">Work&nbsp;Address:</font></span></td>
<td width="278" bgcolor="white">
<p><font face="Verdana" size="2" color="black"><% response.write x_work_street %></font></td>
<td width="169" bgcolor="white">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td width="171" bgcolor="white"><SPAN class=required style="font-size:10pt;"><font face="Arial">City:</font></SPAN></td>
<td width="278" bgcolor="white">
<p><font face="Verdana" size="2" color="black"><% response.write x_work_city %></font></td>
<td width="169" bgcolor="white">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td width="171" bgcolor="white"><span style="font-size:10pt;"><font face="Arial">State:</font></span></td>
<td width="278" bgcolor="white">
<p><font face="Verdana" size="2" color="black"><% response.write x_work_state %></font></td>
<td width="169" bgcolor="white">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td width="171" bgcolor="white"> <font face="Arial"><span style="font-size:10pt;">Zip:</span></font></td>
<td width="278" bgcolor="white">
<p><font face="Verdana" size="2" color="black"><% response.write x_work_zip %></font></td>
<td width="169" bgcolor="white">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td width="171" bgcolor="white"><font face="Arial"><span style="font-size:10pt;">Work Phone&nbsp;</span></font></td>
<td width="278" bgcolor="white">
<p><font face="Verdana" size="2" color="black"><% response.write x_work_phone %></font></td>
<td width="169" bgcolor="white">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td width="171" bgcolor="white"><font face="Arial"><span style="font-size:10pt;">Time At Current Employer:</span></font></td>
<td width="278" bgcolor="white">
<p><font face="Verdana" size="2" color="black"><% response.write x_worktime %></font></td>
<td width="169" bgcolor="white">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td width="171" bgcolor="white"><SPAN class=required style="font-size:10pt;"><font face="Arial" color="maroon">* Net Salary:</font></SPAN></td>
<td width="278" bgcolor="white">
<p><font face="Verdana" size="2" color="black"><% response.write x_net_salary %></font></td>
<td width="169" bgcolor="white">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td width="171" bgcolor="white"><font face="Arial"><span style="font-size:10pt;">Other Income:</span></font></td>
<td width="278" bgcolor="white">
<p><font face="Verdana" size="2" color="black"><% response.write x_other_income %></font></td>
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
                            <p></p>
</td>
</tr>
<tr>
<td width="119" bgcolor="white">
                            <p><font face="Arial"><span style="font-size:10pt;">&nbsp;</span></font></p>
</td>
<td width="171" bgcolor="white"><SPAN class=required style="font-size:11pt;"><font face="Arial" color="maroon">* I Authorize:</font></SPAN><SPAN class=smalltext style="font-size:8pt;"><font face="Arial">(enter your initials)</font></SPAN></td>
<td width="278" bgcolor="white">
<p><font face="Verdana" color="black"><span style="font-size:10pt;"><% response.write x_initials %></span></font></td>
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
<td width="278" bgcolor="white" valign="top"><font face="Verdana" color="black"><span style="font-size:10pt;"><%response.write x_iagree
%></span></font></td>
<td width="169" bgcolor="white">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td width="290" bgcolor="white" colspan="2">
                            <p><span style="font-size:14pt;"><b><font face="Arial"><br>&nbsp;&nbsp;&nbsp;<a href="creditlist.asp"><img src="images/back.gif" align="middle" width="16" height="16" border="0"></a> 
                            </font></b></span><a href="creditlist.asp"><font face="Arial" size="2" color="black"><b>Back to Credit 
                            List</b></font></a><span style="font-size:14pt;"><b><font face="Arial">&nbsp;</font></b></span></p>
</td>
<td width="278" bgcolor="white"><font face="Arial" color="white"><span style="font-size:10pt;"><%= x_stock%>&nbsp;<input type="hidden" name="x_stock" value="<%= x_stock %>">&nbsp;</span></font></td>
<td width="169" bgcolor="white">
                            <p>&nbsp;</p>
</td>
</tr>
</table>

</form>
        </td>
    </tr>
</table>
<!--#include file="footer.asp"-->
