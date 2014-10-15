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
		strsql = "SELECT * FROM [cocredit] WHERE [ID]=" & tkey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn
		If rs.EOF Then
			Response.Clear
			Response.Redirect "cocreditlist.asp"
		Else
			rs.MoveFirst
		' Get the field contents
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
		'x_stock = rs("stock")
		End If
		rs.Close
		Set rs = Nothing
	Case "A": ' Add
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
		' Open record
		strsql = "SELECT * FROM [cocredit] WHERE 0 = 1"
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn, 1, 2
		rs.AddNew
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
        <td><script language="JavaScript" src="ew.js"></script>
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

<form onSubmit="return EW_checkMyForm(this);"  action="cocreditadd.asp" method="post">
<input type="hidden" name="a" value="A">
<table border="0" cellspacing="0" cellpadding="0" bgcolor="#CCCCCC" width="797">
<tr>
<td bgcolor="white" width="797" colspan="5">
                            <table border="0" cellspacing="0" bordercolordark="white" bordercolorlight="black" align="center" width="426">
                                <tr>
                                    <td width="151" rowspan="6" align="center" valign="middle">                            <p><font face="Arial" size="2"><img src="car_ph.asp?key=<%= k%>&nr=1" border=0 width="140"></font></p>
                                    </td>
                                    <td width="271" colspan="2"><font face="Arial"><span style="font-size:12pt;"><b><%response.write y%> 
            &nbsp;<%response.write mk%> &nbsp;<%response.write m%></b></span></font></td>
                                </tr>
                                <tr>
                                    <td width="271" colspan="2">
            <p></p>
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
<td bgcolor="white" width="23">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="390" colspan="2">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="384" colspan="2">
                            <p align="right"><a href="creditadd.asp?st=<%=x_stock%>&k=<%=k%>&v=<%=v%>&mi=<%=mi%>&p=<%=p%>&mk=<%=mk%>&m=<%=m%>&y=<%=y%>"><font face="Arial" color="#666666"><span style="font-size:11pt;">Click 
                            here to switch to single-applicant form</span></font></a></p>
                            <p align="right">&nbsp;<SPAN class=required style="font-size:8pt;"><font color="maroon" face="Arial">* indicates required fields.</font></SPAN>&nbsp;</p>
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
<td bgcolor="white" width="797" colspan="5">
                            <p align="center"><font face="Arial"><textarea name="formtextarea1" rows="15" cols="80" style="font-family:Arial;">
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
<td bgcolor="white" width="23">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="774" colspan="4">
<p align="center"><br>&nbsp;&nbsp;&nbsp;<font face="Arial"><span style="font-size:9pt;">&nbsp;I am interested in purchasing  a vehicle and request that my Consumer 
Credit <br>Report be obtained, at no cost to me, in order to help determine the 
types and extent of <br>financing which may be available to me. &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br>&nbsp;</span></font></td>
</tr>
<tr>
<td bgcolor="white" width="23">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="163">&nbsp;</td>
<td bgcolor="white" width="227">&nbsp;<SPAN class=required style="font-size:11pt;"><font face="Arial" color="maroon">* I Authorize:</font></SPAN><SPAN class=smalltext style="font-size:8pt;"><font face="Arial">(enter your initials)</font></SPAN></td>
<td bgcolor="white" width="160">
                            <p>&nbsp;<font face="Arial" size="2"><input type="text" name="x_initials" size="25" maxlength=50 value="<%= Server.HtmlEncode(x_initials&"") %>"></font></p>
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
<td bgcolor="white" width="227">&nbsp;<font face="Arial" color="maroon"><span style="font-size:11pt;">*Do 
                            you 
                            agree terms and conditiones?</span></font></td>
<td bgcolor="white" width="160">
                            <p>&nbsp;<font face="Arial" size="2"><%
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
%>
</font>&nbsp;</p>
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
<td bgcolor="white" width="227"><font face="Arial" size="2" color="white"><%= x_stock %><input type="hidden" name="x_stock" value="<%= x_stock %>"></font><font color="white">&nbsp;</font></td>
<td bgcolor="white" width="160">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="224">
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
