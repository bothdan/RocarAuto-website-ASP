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
x_stock = Request.QueryString("st")
x_make = Request.QueryString("mk")
x_model = Request.QueryString("md")
x_year = Request.QueryString("yr")
x_price = Request.QueryString("pr")
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
		strsql = "SELECT * FROM [makeoffer] WHERE [ID]=" & tkey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn
		If rs.EOF Then
			Response.Clear
			Response.Redirect "makeofferlist.asp"
		Else
			rs.MoveFirst
		' Get the field contents
		x_first = rs("first")
		x_last = rs("last")
		x_phone = rs("phone")
		x_email = rs("email")
		x_offer = rs("offer")
		x_stock = rs("stock")
		x_make = rs("make")
		x_model = rs("model")
		x_year = rs("year")
		x_price = rs("price")
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
x_offer = Request.Form("x_offer")
x_stock = Request.Form("x_stock")
x_make = Request.Form("x_make")
x_model = Request.Form("x_model")
x_year = Request.Form("x_year")
x_price = Request.Form("x_price")
		' Open record
		strsql = "SELECT * FROM [makeoffer] WHERE 0 = 1"
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
		tmpFld = Trim(x_offer)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("offer") = tmpFld
		tmpFld = Trim(x_stock)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("stock") = tmpFld
		tmpFld = Trim(x_make)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("make") = tmpFld
		
		tmpFld = Trim(x_model)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("model") = tmpFld
		
		tmpFld = Trim(x_year)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("year") = tmpFld
		
		tmpFld = Trim(x_price)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("price") = tmpFld
		
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
<p><font face="Arial" size="2">Add to TABLE: makeoffer<br><br><a href="makeofferlist.asp">Back to List</a></font></p>
<script language="JavaScript" src="ew.js"></script>
<script language="JavaScript">
<!-- start Javascript
function  EW_checkMyForm(EW_this) {
return true;
}
// end JavaScript -->
</script>
<form onSubmit="return EW_checkMyForm(this);"  action="makeofferadd.asp" method="post">
<p>
<input type="hidden" name="a" value="A">
<table border="0" cellspacing="0" cellpadding="0" bgcolor="#CCCCCC" width="341">
<tr>
<td bgcolor="#708090" width="41">
                <p>&nbsp;</p>
</td>
<td bgcolor="#F5F5F5" width="300">
                <p>&nbsp;</p>
</td>
</tr>
<tr>
<td bgcolor="#708090" width="41"><font color="#FFFFFF"><font face="Arial" size="2">ID</font>&nbsp;</font></td>
<td bgcolor="#F5F5F5" width="300"><font face="Arial" size="2"><%= x_ID %><input type="hidden" name="x_ID" value="<%= x_ID %>"></font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#708090" width="41"><font color="#FFFFFF"><font face="Arial" size="2">first</font>&nbsp;</font></td>
<td bgcolor="#F5F5F5" width="300"><font face="Arial" size="2"><input type="text" name="x_first" size=30 maxlength=50 value="<%= Server.HtmlEncode(x_first&"") %>"></font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#708090" width="41"><font color="#FFFFFF"><font face="Arial" size="2">last</font>&nbsp;</font></td>
<td bgcolor="#F5F5F5" width="300"><font face="Arial" size="2"><input type="text" name="x_last" size=30 maxlength=50 value="<%= Server.HtmlEncode(x_last&"") %>"></font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#708090" width="41"><font color="#FFFFFF"><font face="Arial" size="2">phone</font>&nbsp;</font></td>
<td bgcolor="#F5F5F5" width="300"><font face="Arial" size="2"><input type="text" name="x_phone" size=30 maxlength=50 value="<%= Server.HtmlEncode(x_phone&"") %>"></font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#708090" width="41"><font color="#FFFFFF"><font face="Arial" size="2">email</font>&nbsp;</font></td>
<td bgcolor="#F5F5F5" width="300"><font face="Arial" size="2"><input type="text" name="x_email" size=30 maxlength=50 value="<%= Server.HtmlEncode(x_email&"") %>"></font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#708090" width="41"><font color="#FFFFFF"><font face="Arial" size="2">offer</font>&nbsp;</font></td>
<td bgcolor="#F5F5F5" width="300"><font face="Arial" size="2"><input type="text" name="x_offer" size=30 maxlength=50 value="<%= Server.HtmlEncode(x_offer&"") %>"></font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#708090" width="41"><font color="#FFFFFF"><font face="Arial" size="2">stock</font>&nbsp;</font></td>
<td bgcolor="#F5F5F5" width="300">
<p><font face="Arial" size="2"><% response.write x_stock %><input type="hidden" name="x_stock" value="<%= x_stock %>"></font>&nbsp;</td>
</tr>
        <tr>
<td bgcolor="#708090" width="41">
                <p>&nbsp;</p>
</td>
<td bgcolor="#F5F5F5" width="300">
<p><font face="Arial" size="2"><% response.write x_make %><input type="hidden" name="x_make" value="<%= x_make %>"></font></td>
        </tr>
        <tr>
<td bgcolor="#708090" width="41">
                <p>&nbsp;</p>
</td>
<td bgcolor="#F5F5F5" width="300">
<p><font face="Arial" size="2"><% response.write x_model %><input type="hidden" name="x_model" value="<%= x_model %>"></font></td>
        </tr>
        <tr>
<td bgcolor="#708090" width="41">
                <p>&nbsp;</p>
</td>
<td bgcolor="#F5F5F5" width="300">
<p><font face="Arial" size="2"><% response.write x_year %><input type="hidden" name="x_year" value="<%= x_year %>"></font></td>
        </tr>
        <tr>
<td bgcolor="#708090" width="41">
                <p>&nbsp;</p>
</td>
<td bgcolor="#F5F5F5" width="300">
<p><font face="Arial" size="2"><% response.write x_price %><input type="hidden" name="x_price" value="<%= x_price %>"></font></td>
        </tr>
</table>
<p>
<input type="submit" name="Action" value="ADD">
</form>
<!--#include file="footer.asp"-->
