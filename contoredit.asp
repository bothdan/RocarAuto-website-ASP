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
If key="" OR IsNull(key) Then Response.Redirect "contorlist.asp"
'get action
a=Request.Form("a")
If a="" OR IsNull(a) Then
	a="I"	'display with input box
End If
'get fields from form
x_ID = Request.Form("x_ID")
x_contor = Request.Form("x_contor")
x_poston = Request.Form("x_poston")
x_members = Request.Form("x_members")
x_onlinem = Request.Form("x_onlinem")
' Open Connection to the database
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str
Select Case a
	Case "I": ' Get a record to display
		tkey = "" & key & ""
		strsql = "SELECT * FROM [contor] WHERE [ID]=" & tkey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn
		If rs.EOF Then
			Response.Clear
			Response.Redirect "contorlist.asp"
		Else
			rs.MoveFirst
		End If
		' Get the field contents
		x_ID = rs("ID")
		x_contor = rs("contor")
		x_poston = rs("poston")
		x_members = rs("members")
		x_onlinem = rs("onlinem")
		rs.Close
		Set rs = Nothing
	Case "U": ' Update
		' Open record
		tkey = "" & key & ""
		strsql = "SELECT * FROM [contor] WHERE [ID]=" & tkey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn, 1, 2
		If rs.EOF Then
			Response.Clear
			Response.Redirect "contorlist.asp"
		End If
		tmpFld = x_contor
		If Not IsNumeric(tmpFld) Then tmpFld = 0
		rs("contor") = cLng(tmpFld)
		tmpFld = x_poston
		If Not IsNumeric(tmpFld) Then tmpFld = 0
		rs("poston") = cLng(tmpFld)
		tmpFld = x_members
		If Not IsNumeric(tmpFld) Then tmpFld = 0
		rs("members") = cLng(tmpFld)
		tmpFld = x_onlinem
		If Not IsNumeric(tmpFld) Then tmpFld = 0
		rs("onlinem") = cLng(tmpFld)
		rs.Update
		rs.Close
		Set rs = Nothing
		conn.Close
		Set conn = Nothing
		Response.Clear
		Response.Redirect "contorlist.asp"
End Select
%>
<!--#include file="header.asp"-->
<p><font face="Arial" size="2">Edit TABLE: contor<br><br><a href="contorlist.asp">Back to List</a></font></p>
<script language="JavaScript" src="ew.js"></script>
<script language="JavaScript">
<!-- start Javascript
function  EW_checkMyForm(EW_this) {
if (EW_this.x_contor && !EW_checkinteger(EW_this.x_contor.value)) {
        if (!EW_onError(EW_this, EW_this.x_contor, "TEXT", "Incorrect integer - contor"))
            return false; 
        }
if (EW_this.x_poston && !EW_checkinteger(EW_this.x_poston.value)) {
        if (!EW_onError(EW_this, EW_this.x_poston, "TEXT", "Incorrect integer - poston"))
            return false; 
        }
if (EW_this.x_members && !EW_checkinteger(EW_this.x_members.value)) {
        if (!EW_onError(EW_this, EW_this.x_members, "TEXT", "Incorrect integer - members"))
            return false; 
        }
if (EW_this.x_onlinem && !EW_checkinteger(EW_this.x_onlinem.value)) {
        if (!EW_onError(EW_this, EW_this.x_onlinem, "TEXT", "Incorrect integer - onlinem"))
            return false; 
        }
return true;
}
// end JavaScript -->
</script>
<form onSubmit="return EW_checkMyForm(this);"  action="contoredit.asp" method="post">
<p>
<input type="hidden" name="a" value="U">
<input type="hidden" name="key" value="<%= key %>">
<table border="0" cellspacing="0" cellpadding="0" bgcolor="#CCCCCC">
<tr>
<td bgcolor="#708090"><font color="#FFFFFF"><font face="Arial" size="2">ID</font>&nbsp;</font></td>
<td bgcolor="#F5F5F5"><font face="Arial" size="2"><%= x_ID %><input type="hidden" name="x_ID" value="<%= x_ID %>"></font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#708090"><font color="#FFFFFF"><font face="Arial" size="2">contor</font>&nbsp;</font></td>
<td bgcolor="#F5F5F5"><font face="Arial" size="2"><input type="text" name="x_contor" value="<%= Server.HtmlEncode(x_contor&"") %>"></font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#708090"><font color="#FFFFFF"><font face="Arial" size="2">poston</font>&nbsp;</font></td>
<td bgcolor="#F5F5F5"><font face="Arial" size="2"><input type="text" name="x_poston" value="<%= Server.HtmlEncode(x_poston&"") %>"></font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#708090"><font color="#FFFFFF"><font face="Arial" size="2">members</font>&nbsp;</font></td>
<td bgcolor="#F5F5F5"><font face="Arial" size="2"><input type="text" name="x_members" value="<%= Server.HtmlEncode(x_members&"") %>"></font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#708090"><font color="#FFFFFF"><font face="Arial" size="2">onlinem</font>&nbsp;</font></td>
<td bgcolor="#F5F5F5"><font face="Arial" size="2"><input type="text" name="x_onlinem" value="<%= Server.HtmlEncode(x_onlinem&"") %>"></font>&nbsp;</td>
</tr>
</table>
<p>
<input type="submit" name="Action" value="EDIT">
</form>
<!--#include file="footer.asp"-->
