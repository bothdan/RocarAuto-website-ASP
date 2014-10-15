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
		strsql = "SELECT * FROM [contact] WHERE [ID]=" & tkey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn
		If rs.EOF Then
			Response.Clear
			Response.Redirect "contactlist.asp"
		Else
			rs.MoveFirst
		' Get the field contents
		x_first_name = rs("first name")
		x_last_name = rs("last name")
		x_email = rs("email")
		x_phone = rs("phone")
		x_comments = rs("comments")
		End If
		rs.Close
		Set rs = Nothing
	Case "A": ' Add
		'get fields from form
x_ID = Request.Form("x_ID")
x_first_name = Request.Form("x_first_name")
x_last_name = Request.Form("x_last_name")
x_email = Request.Form("x_email")
x_phone = Request.Form("x_phone")
x_comments = Request.Form("x_comments")
		' Open record
		strsql = "SELECT * FROM [contact] WHERE 0 = 1"
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn, 1, 2
		rs.AddNew
		tmpFld = Trim(x_first_name)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("first name") = tmpFld
		tmpFld = Trim(x_last_name)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("last name") = tmpFld
		tmpFld = Trim(x_email)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("email") = tmpFld
		tmpFld = Trim(x_phone)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("phone") = tmpFld
		tmpFld = Trim(x_comments)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("comments") = tmpFld
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
<table cellpadding="0" cellspacing="0" width="800" align="center">
    <tr>
        <td width="800" colspan="3" background="images/contactbg.gif">
            <p><img src="images/contactbar.gif" width="269" height="32" border="0"></p>
        </td>
    </tr>
    <tr>
        <td width="253" bgcolor="white" height="64">
            <p align="center"><br> <font face="Arial"><b><span style="font-size:16pt;">Contact Information</span></b></font></td>
        <td width="304" bgcolor="white" rowspan="2" height="370" align="center" valign="top"><script language="JavaScript" src="ew.js"></script>
<script language="JavaScript">
<!-- start Javascript
function  EW_checkMyForm(EW_this) {
return true;
}
// end JavaScript -->
</script>
<form onSubmit="return EW_checkMyForm(this);"  action="contactadd.asp" method="post">
<p>
<input type="hidden" name="a" value="A">
<font face="Arial" size="2"><%= x_ID %><input type="hidden" name="x_ID" value="<%= x_ID %>"></font><table border="0" cellspacing="0" cellpadding="0" bgcolor="#CCCCCC" width="303">
<tr>
<td bgcolor="white" width="125"></td>
<td bgcolor="white" width="178">
                            <p align="right">&nbsp;<SPAN class="required" style="font-size:9pt;"><font face="Arial" color="maroon">* indicates required 
fields.</font></SPAN></td>
</tr>
<tr>
<td bgcolor="white" width="125" height="28">
                            <p align="left"><font color="maroon" face="Arial"><span style="font-size:11pt;">*First Name:</span></font></td>
<td bgcolor="#F5F5F5" width="178" height="28"><font face="Arial"><span style="font-size:11pt;"><input type="text" name="x_first_name" size="26" maxlength=50 value="<%= Server.HtmlEncode(x_first_name&"") %>">&nbsp;</span></font></td>
</tr>
<tr>
<td bgcolor="white" width="125" height="28">
                            <p align="left"><font color="maroon" face="Arial"><span style="font-size:11pt;">*Last Name:</span></font></td>
<td bgcolor="#F5F5F5" width="178" height="28"><font face="Arial"><span style="font-size:11pt;"><input type="text" name="x_last_name" size="26" maxlength=50 value="<%= Server.HtmlEncode(x_last_name&"") %>">&nbsp;</span></font></td>
</tr>
<tr>
<td bgcolor="white" width="125" height="28">
                            <p align="left"><font color="maroon" face="Arial"><span style="font-size:11pt;">*E-mail 
                            Address:</span></font></td>
<td bgcolor="#F5F5F5" width="178" height="28"><font face="Arial"><span style="font-size:11pt;"><input type="text" name="x_email" size="26" maxlength=50 value="<%= Server.HtmlEncode(x_email&"") %>">&nbsp;</span></font></td>
</tr>
<tr>
<td bgcolor="white" width="125" height="31">
                            <p align="left"><font color="black" face="Arial"><span style="font-size:11pt;">Phone 
                            Number:</span></font></td>
<td bgcolor="#F5F5F5" width="178" height="31"><font face="Arial"><span style="font-size:11pt;"><input type="text" name="x_phone" size="26" maxlength=50 value="<%= Server.HtmlEncode(x_phone&"") %>">&nbsp;</span></font></td>
</tr>
<tr>
<td bgcolor="white" width="125" valign="top">
                            <p align="left"><font color="maroon" face="Arial"><span style="font-size:11pt;">*Comments:</span></font></td>
<td bgcolor="#F5F5F5" width="178"><font face="Arial"><span style="font-size:11pt;"><textarea cols="20" rows="4" name="x_comments"><%= x_comments %></textarea>&nbsp;</span></font></td>
</tr>
</table>
<p align="center">
<input type="submit" name="Action" value="Submit">
</form>
            <p>&nbsp;</p>
        </td>
        <td width="243" bgcolor="white" valign="top" rowspan="2" height="370">
            <table align="center" cellpadding="0" cellspacing="0" width="214">
                <tr>
                    <td width="12" height="9">
                        <p><font face="Arial"><span style="font-size:5pt;"><img src="images/upleft.gif" width="12" height="10" border="0"></span></font></p>
                    </td>
                    <td width="194" background="images/up.gif" height="9">
                        <p><font face="Arial"><span style="font-size:5pt;">&nbsp;</span></font></p>
                    </td>
                    <td width="8" height="9">
                        <p><font face="Arial"><span style="font-size:5pt;"><img src="images/upright.gif" width="13" height="10" border="0"></span></font></p>
                    </td>
                </tr>
                <tr>
                    <td width="12" rowspan="17" background="images/leftbg.gif" height="325">
                        <p><font face="Arial"><span style="font-size:8pt;">&nbsp;</span></font></p>
                    </td>
                    <td width="194" nowrap><SPAN style="font-size:10pt;"><font face="Arial"><b>Contact Information</b></font></SPAN></td>
                    <td width="8" rowspan="17" background="images/rightbg.gif" height="325">
                        <p><font face="Arial"><span style="font-size:8pt;">&nbsp;</span></font></p>
                    </td>
                </tr>
                <tr>
                    <td width="194" nowrap>
                        <p><font face="Arial"><span style="font-size:8pt;">&nbsp;</span></font></p>
                    </td>
                </tr>
                <tr>
                    <td width="194" nowrap height="17">
                        <p><font face="Arial"><span style="font-size:9pt;">5136 
                        N. Western Ave.</span></font></p>
                    </td>
                </tr>
                <tr>
                    <td width="194" nowrap height="17">
                        <p><font face="Arial"><span style="font-size:9pt;">Chicago, 
                        IL 60625-2533</span></font></p>
                    </td>
                </tr>
                <tr>
                    <td width="194" nowrap height="17"><span style="font-size:9pt;"><font face="Arial">Phone: (773) 334-0025 </font></span></td>
                </tr>
                <tr>
                    <td width="194" nowrap height="17"><span style="font-size:9pt;"><font face="Arial">Fax: 773-334-5036 </font></span></td>
                </tr>
                <tr>
                    <td width="194" nowrap height="17">
                        <p><font face="Arial"><span style="font-size:9pt;">Email: 
                        </span></font><a href="contactadd.asp"><span style="font-size:9pt;"><b><font face="Arial" color="teal">Email 
                        Us</font></b></span></a></p>
                    </td>
                </tr>
                <tr>
                    <td width="194" nowrap>
                        <p><font face="Arial"><span style="font-size:8pt;">&nbsp;</span></font></p>
                    </td>
                </tr>
                <tr>
                    <td width="194" nowrap><SPAN style="font-size:10pt;"><b><font face="Arial">Dealership Hours</font></b></SPAN></td>
                </tr>
                <tr>
                    <td width="194" nowrap>
                        <p><font face="Arial"><span style="font-size:8pt;">&nbsp;</span></font></p>
                    </td>
                </tr>
                <tr>
                    <td width="194" nowrap height="17"><span style="font-size:9pt;"><font face="Arial">Monday:&nbsp;10:00 AM - 7:00 PM </font></span></td>
                </tr>
                <tr>
                    <td width="194" nowrap height="17"><span style="font-size:9pt;"><font face="Arial">Tuesday:&nbsp;10:00 AM - 7:00 PM </font></span></td>
                </tr>
                <tr>
                    <td width="194" nowrap height="17"><span style="font-size:9pt;"><font face="Arial">Wednesday:&nbsp;10:00 AM - 7:00 PM </font></span></td>
                </tr>
                <tr>
                    <td width="194" nowrap height="17"><span style="font-size:9pt;"><font face="Arial">Thursday:&nbsp;10:00 AM - 7:00 PM </font></span></td>
                </tr>
                <tr>
                    <td width="194" nowrap height="17"><span style="font-size:9pt;"><font face="Arial">Friday:&nbsp;10:00 AM - 7:00 PM </font></span></td>
                </tr>
                <tr>
                    <td width="194" nowrap><span style="font-size:9pt;"><font face="Arial">Saturday:&nbsp;10:00 AM - 6:00 PM</font></span></td>
                </tr>
                <tr>
                    <td width="194" nowrap><span style="font-size:9pt;"><font face="Arial">Sunday:&nbsp;Closed</font></span></td>
                </tr>
                <tr>
                    <td width="12">
                        <p><font face="Arial"><span style="font-size:8pt;"><img src="images/dwleft.gif" width="12" height="11" border="0"></span></font></p>
                    </td>
                    <td width="194" background="images/dwbg.gif">
                        <p><font face="Arial"><span style="font-size:5pt;">&nbsp;</span></font></p>
                    </td>
                    <td width="8">
                        <p><font face="Arial"><span style="font-size:5pt;"><img src="images/dwright.gif" width="13" height="11" border="0"></span></font></p>
                    </td>
                </tr>
            </table>
        </td>
    </tr>
    <tr>
        <td width="253" bgcolor="white" height="306" background="images/silvercar.gif">
            <p>&nbsp;</p>
        </td>
    </tr>
</table>
<!--#include file="footer.asp"-->