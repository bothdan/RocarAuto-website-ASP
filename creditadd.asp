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
x_stock = Request.Querystring("st")
k = Request.Querystring("k")

v = Request.Querystring("v")
mi = Request.Querystring("mi")
p = Request.Querystring("p")
mk = Request.Querystring("mk")
m = Request.Querystring("m")
y = Request.Querystring("y")
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
		strsql = "SELECT * FROM [credit] WHERE [ID]=" & tkey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn
		If rs.EOF Then
			Response.Clear
			Response.Redirect "creditlist.asp"
		Else
			rs.MoveFirst
		' Get the field contents
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
		'x_stock = st
		End If
		rs.Close
		Set rs = Nothing
	Case "A": ' Add
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
'x_stock = st
x_stock = Request.Form("x_stock")
		' Open record
		strsql = "SELECT * FROM [credit] WHERE 0 = 1"
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn, 1, 2
		rs.AddNew
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
		Response.Redirect "offerthanks.asp"
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
<script language="JavaScript" src="ew.js"></script>
<script language="JavaScript">
<!-- start Javascript
function  EW_checkMyForm(EW_this) {
return true;
}
// end JavaScript -->
</script>


<form onSubmit="return EW_checkMyForm(this);"  action="creditadd.asp" method="post">
<input type="hidden" name="a" value="A">
<table border="0" cellspacing="0" cellpadding="0" bgcolor="#CCCCCC" width="737">
<tr>
<td width="737" bgcolor="white" height="67" align="left" valign="top" colspan="4">

                            <table border="0" cellspacing="0" bordercolordark="white" bordercolorlight="black" align="center" width="426">
                                <tr>
                                    <td width="151" rowspan="6" align="center" valign="middle">                            <p><font face="Arial" size="2"><img src="car_ph.asp?key=<%= k%>&nr=1" border=0 width="140"></font></p>
                                    </td>
                                    <td width="271" colspan="2"><font face="Arial"><span style="font-size:12pt;"><b><%response.write y%> 
                                        &nbsp;<%response.write mk%> &nbsp;<%response.write m%></b></span></font></td>
                                </tr>
                                <tr>
                                    <td width="271" colspan="2">
                                        <p>&nbsp;</p>
                                    </td>
                                </tr>
                                <tr>
                                    <td width="61">
                                        <p><font face="Arial"><span style="font-size:10pt;">Stock 
                                        #:</span></font></p>
                                    </td>
                                    <td width="208">
                                        <p><font face="Arial"><span style="font-size:10pt;"><%response.write x_stock%></span></font></p>
                                    </td>
                                </tr>
                                <tr>
                                    <td width="61">
                                        <p><font face="Arial"><span style="font-size:10pt;">VIN 
                                        #:</span></font></p>
                                    </td>
                                    <td width="208">
                                        <p><font face="Arial"><span style="font-size:10pt;"><%response.write v%></span></font></p>
                                    </td>
                                </tr>
                                <tr>
                                    <td width="61">
                                        <p><font face="Arial"><span style="font-size:10pt;">Mileage:</span></font></p>
                                    </td>
                                    <td width="208">
                                        <p><span style="font-size:10pt;"><font face="Arial"><%if isnumeric(mi) then response.write formatnumber(mi,0,-2,-2,-2) else response.write mi end if %></font></span></p>
                                    </td>
                                </tr>
                                <tr>
                                    <td width="61">
                                        <p><font face="Arial"><span style="font-size:10pt;">Price:</span></font></p>
                                    </td>
                                    <td width="208">
                                        <p><font face="Arial"><span style="font-size:10pt;"><% if isnumeric(p) then response.write formatcurrency(p,0,-2,-2,-2) else response.write p end if %></span></font></p>
                                    </td>
                                </tr>
                            </table>
</td>
</tr>
<tr>
<td width="290" bgcolor="white" height="67" align="left" valign="top" colspan="2">
                            <p>&nbsp;</p>

<p><span style="font-size:14pt;"><b><font face="Arial">&nbsp;&nbsp;&nbsp;Application Information:</font></b></span></td>
<td width="447" bgcolor="white" height="61" colspan="2">
                            <p align="right"><font face="Arial" color="white"><span style="font-size:12pt;"><%= x_ID %><input type="hidden" name="x_ID" value="<%= x_ID %>">&nbsp;</span></font><a href="cocreditadd.asp?st=<%=x_stock%>&k=<%=k%>&v=<%=v%>&mi=<%=mi%>&p=<%=p%>&mk=<%=mk%>&m=<%=m%>&y=<%=y%>"><span style="font-size:11pt;"><font face="Arial" color="#666666">Click 
                            here to switch to co-applicant form</font></span></a>                            <p align="right">&nbsp;<SPAN class=required style="font-size:8pt;"><font color="maroon" face="Arial">* indicates required fields.</font></SPAN></p>
</td>
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
                            <p><font face="Arial"><textarea name="formtextarea1" rows="15" cols="80" style="font-family:Arial;">
As a valued customer, we want to ensure your private information is kept confidential and only shared with those companies who are authorized either by yourself or as allowed or required by law.  This document explains our privacy policy, gives you reasons why we ask for the type of information we do, and if we do reserve a right to share information with non-affiliated third parties, lets you "opt-out" of our reservation to do so.  Please take a moment to read this entire policy.

Collection of Information

The purchase of a motor vehicle requires considerable accumulation of nonpublic personal information.  For example, if we sell you a vehicle - extending you credit at your request - we will receive information from you in order to determine your creditworthiness.  We may also obtain information from a credit-reporting agency.  We may also obtain information from third parties such as employers, references and insurance companies.

Some of the information we obtain from you may be required by state of federal agencies, such as the Department of Motor Vehicles or the Internal Revenue Service. This information may be required even if you were to pay cash for your vehicle. Examples would be a driver's license or social security number.

Protecting Your Information

We safeguard nonpublic personal information according to established industry standards and procedures.  We maintain physical and electronic safeguards that comply with state and federal law.  We restrict access to nonpublic personal information about you to those employees and outside contractors who need to know the information to provide product or service to you.  We prohibit our employees and agents from giving information about you to anyone in a manner that would violate any applicable law or our privacy policy.

Information Sharing


A) As permitted by federal or state law.

B) For purposes of processing a sale or lease transaction as your request or authorize, such as submitting information to third party financial institutions that may be requested to take an assignment of the contract or verifying insurance coverage information.

C) When using outside service providers to help us provide you with products and services.  Before providing information to our service providers we enter into contractual agreements prohibiting them from disclosing or using the information other than for the purpose it was disclosed.

D) With "Affiliated" companies.  Companies that are affiliated with us include any company that controls us, any company we control, or any company under common control with us.
							</textarea></font></p>
</td>
</tr>
<tr>
<td width="119" bgcolor="white">
                            <p>&nbsp;</p>
</td>
<td width="449" bgcolor="white" colspan="2"><br>&nbsp;&nbsp;&nbsp;<font face="Arial"><span style="font-size:9pt;">&nbsp;I am interested in purchasing  a vehicle and request that my Consumer 
Credit Report be obtained, at no cost to me, in order to help determine the 
types and extent of financing which may be available to me.<br>&nbsp;</span></font></td>
<td width="169" bgcolor="white">
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
<p align="center">
<input type="submit" name="Action" value="Submit">
</form>
        </td>
    </tr>
</table>
<!--#include file="footer.asp"-->
