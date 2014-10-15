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
		strsql = "SELECT * FROM [moreinfo] WHERE [ID]=" & tkey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn
		If rs.EOF Then
			Response.Clear
			Response.Redirect "moreinfolist.asp"
		Else
			rs.MoveFirst
		' Get the field contents
		x_first = rs("first")
		x_last = rs("last")
		x_phone = rs("phone")
		x_email = rs("email")
		x_testdrive = rs("testdrive")
		x_comments = rs("comments")
		x_stock = rs("stock")
		End If
		rs.Close
		Set rs = Nothing
	Case "A": ' Add
		'get fields from form
x_ID = Request.Form("x_ID")
x_first = Request.Form("x_first")
x_last = Request.Form("x_last")
x_phone = Request.Form("x_phone")
x_email = Request.Form("x_email")
x_testdrive = Request.Form("x_testdrive")
x_comments = Request.Form("x_comments")
x_stock = Request.Form("x_stock")
		' Open record
		strsql = "SELECT * FROM [moreinfo] WHERE 0 = 1"
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn, 1, 2
		rs.AddNew
		tmpFld = Trim(x_first)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("first") = tmpFld
		tmpFld = Trim(x_last)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("last") = tmpFld
		tmpFld = Trim(x_phone)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("phone") = tmpFld
		tmpFld = Trim(x_email)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("email") = tmpFld
		tmpFld = Trim(x_testdrive)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("testdrive") = tmpFld
		tmpFld = Trim(x_comments)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("comments") = tmpFld
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
<p><font face="Arial" size="2">Add to TABLE: moreinfo<br><br><a href="moreinfolist.asp">Back to List</a></font></p>
<script language="JavaScript" src="ew.js"></script>
<script language="JavaScript">
<!-- start Javascript
function  EW_checkMyForm(EW_this) {
return true;
}
// end JavaScript -->
</script>
<form onSubmit="return EW_checkMyForm(this);"  action="moreinfoadd.asp" method="post">
<p>
<input type="hidden" name="a" value="A">
<table border="0" cellspacing="0" cellpadding="0" bgcolor="#CCCCCC">
<tr>
<td bgcolor="#708090"><font color="#FFFFFF"><font face="Arial" size="2">ID</font>&nbsp;</font></td>
<td bgcolor="#F5F5F5"><font face="Arial" size="2"><%= x_ID %><input type="hidden" name="x_ID" value="<%= x_ID %>"></font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#708090"><font color="#FFFFFF"><font face="Arial" size="2">first</font>&nbsp;</font></td>
<td bgcolor="#F5F5F5"><font face="Arial" size="2"><input type="text" name="x_first" size=30 maxlength=50 value="<%= Server.HtmlEncode(x_first&"") %>"></font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#708090"><font color="#FFFFFF"><font face="Arial" size="2">last</font>&nbsp;</font></td>
<td bgcolor="#F5F5F5"><font face="Arial" size="2"><input type="text" name="x_last" size=30 maxlength=50 value="<%= Server.HtmlEncode(x_last&"") %>"></font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#708090"><font color="#FFFFFF"><font face="Arial" size="2">phone</font>&nbsp;</font></td>
<td bgcolor="#F5F5F5"><font face="Arial" size="2"><input type="text" name="x_phone" size=30 maxlength=50 value="<%= Server.HtmlEncode(x_phone&"") %>"></font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#708090"><font color="#FFFFFF"><font face="Arial" size="2">email</font>&nbsp;</font></td>
<td bgcolor="#F5F5F5"><font face="Arial" size="2"><input type="text" name="x_email" size=30 maxlength=50 value="<%= Server.HtmlEncode(x_email&"") %>"></font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#708090"><font color="#FFFFFF"><font face="Arial" size="2">testdrive</font>&nbsp;</font></td>
<td bgcolor="#F5F5F5"><font face="Arial" size="2"><%
x_testdriveList = "<SELECT name='x_testdrive'><OPTION value=''>Please Select</OPTION>"
    x_testdriveList = x_testdriveList & "<OPTION value=""Yes"""
    If x_testdrive = "Yes" Then
        x_testdriveList = x_testdriveList & " selected"
    End If
    x_testdriveList = x_testdriveList & ">" & "Yes" & "</option>"
    x_testdriveList = x_testdriveList & "<OPTION value=""No"""
    If x_testdrive = "No" Then
        x_testdriveList = x_testdriveList & " selected"
    End If
    x_testdriveList = x_testdriveList & ">" & "" & "</option>"
x_testdriveList = x_testdriveList & "</select>"
response.write x_testdriveList
%>
</font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#708090"><font color="#FFFFFF"><font face="Arial" size="2">comments</font>&nbsp;</font></td>
<td bgcolor="#F5F5F5"><font face="Arial" size="2"><textarea cols=35 rows=4 name="x_comments"><%= x_comments %></textarea></font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#708090"><font color="#FFFFFF"><font face="Arial" size="2">stock</font>&nbsp;</font></td>
<td bgcolor="#F5F5F5"><font face="Arial" size="2"><input type="text" name="x_stock" size=30 maxlength=50 value="<%= Server.HtmlEncode(x_stock&"") %>"></font>&nbsp;</td>
</tr>
</table>
<p>
<input type="submit" name="Action" value="ADD">
</form>
<!--#include file="footer.asp"-->
