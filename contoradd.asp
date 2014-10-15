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
		strsql = "SELECT * FROM [contor] WHERE [ID]=" & tkey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn
		If rs.EOF Then
			Response.Clear
			Response.Redirect "contorlist.asp"
		Else
			rs.MoveFirst
		' Get the field contents
		x_contor = rs("contor")
		x_poston = rs("poston")
		x_members = rs("members")
		x_onlinem = rs("onlinem")
		End If
		rs.Close
		Set rs = Nothing
	Case "A": ' Add
		'get fields from form
x_ID = Request.Form("x_ID")
x_contor = Request.Form("x_contor")
x_poston = Request.Form("x_poston")
x_members = Request.Form("x_members")
x_onlinem = Request.Form("x_onlinem")
		' Open record
		strsql = "SELECT * FROM [contor] WHERE 0 = 1"
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn, 1, 2
		rs.AddNew
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
<p><font face="Arial" size="2">Add to TABLE: contor<br><br><a href="contorlist.asp">Back to List</font></a></p>
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
<form onSubmit="return EW_checkMyForm(this);"  action="contoradd.asp" method="post">
<p>
<input type="hidden" name="a" value="A">
<table border="0" cellspacing="0" cellpadding="0" bgcolor="#CCCCCC">
<tr>
<td bgcolor="#708090"><font color="#FFFFFF"><font face="Arial" size="2">ID</font>&nbsp;</font></td>
<td bgcolor="#F5F5F5"><font face="Arial" size="2"><%= x_ID %><input type="hidden" name="x_ID" value="<%= x_ID %>"></font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#708090"><font color="#FFFFFF"><font face="Arial" size="2">contor</font>&nbsp;</font></td>
<td bgcolor="#F5F5F5"><font face="Arial" size="2"><% If isnull(x_contor) or x_contor = "" Then x_contor = 0 'set default value %><input type="text" name="x_contor" value="<%= Server.HtmlEncode(x_contor&"") %>"></font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#708090"><font color="#FFFFFF"><font face="Arial" size="2">poston</font>&nbsp;</font></td>
<td bgcolor="#F5F5F5"><font face="Arial" size="2"><% If isnull(x_poston) or x_poston = "" Then x_poston = 0 'set default value %><input type="text" name="x_poston" value="<%= Server.HtmlEncode(x_poston&"") %>"></font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#708090"><font color="#FFFFFF"><font face="Arial" size="2">members</font>&nbsp;</font></td>
<td bgcolor="#F5F5F5"><font face="Arial" size="2"><% If isnull(x_members) or x_members = "" Then x_members = 0 'set default value %><input type="text" name="x_members" value="<%= Server.HtmlEncode(x_members&"") %>"></font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#708090"><font color="#FFFFFF"><font face="Arial" size="2">onlinem</font>&nbsp;</font></td>
<td bgcolor="#F5F5F5"><font face="Arial" size="2"><% If isnull(x_onlinem) or x_onlinem = "" Then x_onlinem = 0 'set default value %><input type="text" name="x_onlinem" value="<%= Server.HtmlEncode(x_onlinem&"") %>"></font>&nbsp;</td>
</tr>
</table>
<p>
<input type="submit" name="Action" value="ADD">
</form>
<!--#include file="footer.asp"-->
