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
If key="" OR IsNull(key) Then Response.Redirect "contactlist.asp"
'get action
a=Request.Form("a")
If a="" OR IsNull(a) Then
	a="I"	'display with input box
End If
'get fields from form
x_ID = Request.Form("x_ID")
x_first_name = Request.Form("x_first_name")
x_last_name = Request.Form("x_last_name")
x_email = Request.Form("x_email")
x_phone = Request.Form("x_phone")
x_comments = Request.Form("x_comments")
' Open Connection to the database
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str
Select Case a
	Case "I": ' Get a record to display
		tkey = "" & key & ""
		strsql = "SELECT * FROM [contact] WHERE [ID]=" & tkey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn
		If rs.EOF Then
			Response.Clear
			Response.Redirect "contactlist.asp"
		Else
			rs.MoveFirst
		End If
		' Get the field contents
		x_ID = rs("ID")
		x_first_name = rs("first name")
		x_last_name = rs("last name")
		x_email = rs("email")
		x_phone = rs("phone")
		x_comments = rs("comments")
		rs.Close
		Set rs = Nothing
	Case "U": ' Update
		' Open record
		tkey = "" & key & ""
		strsql = "SELECT * FROM [contact] WHERE [ID]=" & tkey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn, 1, 2
		If rs.EOF Then
			Response.Clear
			Response.Redirect "contactlist.asp"
		End If
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
		Response.Redirect "contactlist.asp"
End Select
%>
<!--#include file="header.asp"-->
<meta name="generator" content="Namo WebEditor v5.0(Trial)">
<table align="center" cellpadding="0" cellspacing="0" width="801" bgcolor="white">
    <tr>
        <td>
            <p><font face="Arial" size="2"><br></font></p>
<script language="JavaScript" src="ew.js"></script>
<script language="JavaScript">
<!-- start Javascript
function  EW_checkMyForm(EW_this) {
return true;
}
// end JavaScript -->
</script>
<form onSubmit="return EW_checkMyForm(this);"  action="contactedit.asp" method="post">
<p>
<input type="hidden" name="a" value="U">
<input type="hidden" name="key" value="<%= key %>">
<table border="0" cellspacing="0" cellpadding="0" bgcolor="#CCCCCC" align="center" width="613">
<tr>
<td bgcolor="white" width="128"><font face="Arial"><b><span style="font-size:14pt;">Edit 
                            Contact<br>&nbsp;</span></b></font></td>
<td bgcolor="white" width="485"><font face="Arial" size="2" color="white"><%= x_ID %><input type="hidden" name="x_ID" value="<%= x_ID %>"></font><font color="white">&nbsp;</font></td>
</tr>
<tr>
<td bgcolor="white" width="128"><font face="Arial"><span style="font-size:10pt;"><b>First Name:</b></span></font></td>
<td bgcolor="white" width="485"><font face="Arial"><span style="font-size:10pt;"><input type="text" name="x_first_name" size=30 maxlength=50 value="<%= Server.HtmlEncode(x_first_name&"") %>">&nbsp;</span></font></td>
</tr>
<tr>
<td bgcolor="white" width="128"><font face="Arial"><span style="font-size:10pt;"><b>Last Name:</b></span></font></td>
<td bgcolor="white" width="485"><font face="Arial"><span style="font-size:10pt;"><input type="text" name="x_last_name" size=30 maxlength=50 value="<%= Server.HtmlEncode(x_last_name&"") %>">&nbsp;</span></font></td>
</tr>
<tr>
<td bgcolor="white" width="128"><font face="Arial"><span style="font-size:10pt;"><b>E-mail:</b></span></font></td>
<td bgcolor="white" width="485"><font face="Arial"><span style="font-size:10pt;"><input type="text" name="x_email" size=30 maxlength=50 value="<%= Server.HtmlEncode(x_email&"") %>">&nbsp;</span></font></td>
</tr>
<tr>
<td bgcolor="white" width="128"><font face="Arial"><span style="font-size:10pt;"><b>Phone:</b></span></font></td>
<td bgcolor="white" width="485"><font face="Arial"><span style="font-size:10pt;"><input type="text" name="x_phone" size=30 maxlength=50 value="<%= Server.HtmlEncode(x_phone&"") %>">&nbsp;</span></font></td>
</tr>
<tr>
<td bgcolor="white" width="128"><font face="Arial"><span style="font-size:10pt;"><b>Comments:</b></span></font></td>
<td bgcolor="white" width="485"><font face="Arial"><span style="font-size:10pt;"><textarea cols=35 rows=4 name="x_comments"><%= x_comments %></textarea>&nbsp;</span></font></td>
</tr>
</table>
<p align="left">
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="submit" name="Action" value="Update">
</form>
            <p><font face="Arial" size="2" color="black"><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</b></font><a href="contactlist.asp"><font face="Arial" size="2" color="black"><b><img src="images/leftsm.gif" align="texttop" width="16" height="16" border="0"></b></font></a><font face="Arial" size="2" color="black"><b> 
            &nbsp;&nbsp;</b></font><a href="contactlist.asp"><font face="Arial" size="2" color="black"><b>Back to Contact List</b></font></a><font face="Arial" size="2" color="black"><b><br>&nbsp;</b></font></p>
        </td>
    </tr>
</table>
<!--#include file="footer.asp"-->
