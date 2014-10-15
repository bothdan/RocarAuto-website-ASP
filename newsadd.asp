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
		strsql = "SELECT * FROM [news] WHERE [ID]=" & tkey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn
		If rs.EOF Then
			Response.Clear
			Response.Redirect "newslist.asp"
		Else
			rs.MoveFirst
		' Get the field contents
		x_news = rs("news")
		End If
		rs.Close
		Set rs = Nothing
	Case "A": ' Add
		'get fields from form
x_ID = Request.Form("x_ID")
x_news = Request.Form("x_news")
		' Open record
		strsql = "SELECT * FROM [news] WHERE 0 = 1"
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn, 1, 2
		rs.AddNew
		tmpFld = Trim(x_news)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("news") = tmpFld
		rs.Update
		rs.Close
		Set rs = Nothing
		conn.Close
		Set conn = Nothing
		Response.Clear
		Response.Redirect "newslist.asp"
End Select
%>
<!--#include file="header.asp"-->
<p><font face="Arial" size="2">Add to TABLE: news<br><br><a href="newslist.asp">Back to List</font></a></p>
<script language="JavaScript" src="ew.js"></script>
<script language="JavaScript">
<!-- start Javascript
function  EW_checkMyForm(EW_this) {
return true;
}
// end JavaScript -->
</script>
<form onSubmit="return EW_checkMyForm(this);"  action="newsadd.asp" method="post">
<p>
<input type="hidden" name="a" value="A">
<table border="0" cellspacing="0" cellpadding="0" bgcolor="#CCCCCC">
<tr>
<td bgcolor="#708090"><font color="#FFFFFF"><font face="Arial" size="2">ID</font>&nbsp;</font></td>
<td bgcolor="#F5F5F5"><font face="Arial" size="2"><%= x_ID %><input type="hidden" name="x_ID" value="<%= x_ID %>"></font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#708090"><font color="#FFFFFF"><font face="Arial" size="2">news</font>&nbsp;</font></td>
<td bgcolor="#F5F5F5"><font face="Arial" size="2"><textarea cols=35 rows=4 name="x_news"><%= x_news %></textarea></font>&nbsp;</td>
</tr>
</table>
<p>
<input type="submit" name="Action" value="ADD">
</form>
<!--#include file="footer.asp"-->
