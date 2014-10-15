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
If key="" OR IsNull(key) Then Response.Redirect "makeofferlist.asp"
'get action
a=Request.Form("a")
If a="" OR IsNull(a) Then
	a="I"	'display with input box
End If
'get fields from form
x_ID = Request.Form("x_ID")
x_first = Request.Form("x_first")
x_last = Request.Form("x_last")
x_phone = Request.Form("x_phone")
x_email = Request.Form("x_email")
x_offer = Request.Form("x_offer")
x_stock = Request.Form("x_stock")
' Open Connection to the database
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str
Select Case a
	Case "I": ' Get a record to display
		tkey = "" & key & ""
		strsql = "SELECT * FROM [makeoffer] WHERE [ID]=" & tkey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn
		If rs.EOF Then
			Response.Clear
			Response.Redirect "makeofferlist.asp"
		Else
			rs.MoveFirst
		End If
		' Get the field contents
		x_ID = rs("ID")
		x_first = rs("first")
		x_last = rs("last")
		x_phone = rs("phone")
		x_email = rs("email")
		x_offer = rs("offer")
		x_stock = rs("stock")
		rs.Close
		Set rs = Nothing
	Case "U": ' Update
		' Open record
		tkey = "" & key & ""
		strsql = "SELECT * FROM [makeoffer] WHERE [ID]=" & tkey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn, 1, 2
		If rs.EOF Then
			Response.Clear
			Response.Redirect "makeofferlist.asp"
		End If
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
		rs.Update
		rs.Close
		Set rs = Nothing
		conn.Close
		Set conn = Nothing
		Response.Clear
		Response.Redirect "makeofferlist.asp"
End Select
%>
<!--#include file="header.asp"-->
<meta name="generator" content="Namo WebEditor v5.0(Trial)">
<table align="center" cellpadding="0" cellspacing="0" width="801" bgcolor="white">
    <tr>
        <td>
            <p><font face="Arial" size="2"><br> </font></p>
<script language="JavaScript" src="ew.js"></script>
<script language="JavaScript">
<!-- start Javascript
function  EW_checkMyForm(EW_this) {
return true;
}
// end JavaScript -->
</script>
<form onSubmit="return EW_checkMyForm(this);"  action="makeofferedit.asp" method="post">
<p>
<input type="hidden" name="a" value="U">
<input type="hidden" name="key" value="<%= key %>">
<table border="0" cellspacing="0" cellpadding="0" bgcolor="#CCCCCC" width="474">
<tr>
<td bgcolor="white" width="48">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="97"><b><span style="font-size:14pt;"><font face="Arial">Edit 
                            Offer</font></span><span style="font-size:10pt;"><font face="Arial"><br>&nbsp;</font></span></b></td>
<td bgcolor="white" width="329"><font face="Arial" color="white"><span style="font-size:10pt;"><%= x_ID %><input type="hidden" name="x_ID" value="<%= x_ID %>">&nbsp;</span></font></td>
</tr>
<tr>
<td bgcolor="white" width="48">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="97"><font face="Arial"><b><span style="font-size:10pt;">First</span></b></font><b><span style="font-size:10pt;"><font face="Arial">&nbsp;Name:</font></span></b></td>
<td bgcolor="white" width="329"><font face="Arial"><span style="font-size:10pt;"><input type="text" name="x_first" size=30 maxlength=50 value="<%= Server.HtmlEncode(x_first&"") %>">&nbsp;</span></font></td>
</tr>
<tr>
<td bgcolor="white" width="48">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="97"><font face="Arial"><b><span style="font-size:10pt;">Last</span></b></font><b><span style="font-size:10pt;"><font face="Arial">&nbsp;Name:</font></span></b></td>
<td bgcolor="white" width="329"><font face="Arial"><span style="font-size:10pt;"><input type="text" name="x_last" size=30 maxlength=50 value="<%= Server.HtmlEncode(x_last&"") %>">&nbsp;</span></font></td>
</tr>
<tr>
<td bgcolor="white" width="48">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="97"><font face="Arial"><b><span style="font-size:10pt;">Phone</span></b></font><b><span style="font-size:10pt;"><font face="Arial">:</font></span></b></td>
<td bgcolor="white" width="329"><font face="Arial"><span style="font-size:10pt;"><input type="text" name="x_phone" size=30 maxlength=50 value="<%= Server.HtmlEncode(x_phone&"") %>">&nbsp;</span></font></td>
</tr>
<tr>
<td bgcolor="white" width="48">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="97"><font face="Arial"><b><span style="font-size:10pt;">E-mail:</span></b></font></td>
<td bgcolor="white" width="329"><font face="Arial"><span style="font-size:10pt;"><input type="text" name="x_email" size=30 maxlength=50 value="<%= Server.HtmlEncode(x_email&"") %>">&nbsp;</span></font></td>
</tr>
<tr>
<td bgcolor="white" width="48">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="97"><font face="Arial"><b><span style="font-size:10pt;">Offer</span></b></font><b><span style="font-size:10pt;"><font face="Arial">:</font></span></b></td>
<td bgcolor="white" width="329"><font face="Arial"><span style="font-size:10pt;"><input type="text" name="x_offer" size=30 maxlength=50 value="<%= Server.HtmlEncode(x_offer&"") %>">&nbsp;</span></font></td>
</tr>
<tr>
<td bgcolor="white" width="48">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="97"><font face="Arial"><b><span style="font-size:10pt;">Stock</span></b></font><b><span style="font-size:10pt;"><font face="Arial">&nbsp;#:</font></span></b></td>
<td bgcolor="white" width="329"><font face="Arial"><span style="font-size:10pt;"><input type="text" name="x_stock" size=30 maxlength=50 value="<%= Server.HtmlEncode(x_stock&"") %>">&nbsp;</span></font></td>
</tr>
</table>
<p align="center">
<input type="submit" name="Action" value="Update">
</form>
            <p><font face="Arial" size="2" color="black"><b>&nbsp;&nbsp;&nbsp;</b></font><a href="makeofferlist.asp"><font face="Arial" size="2" color="black"><b><img src="images/leftsm.gif" align="texttop" width="16" height="16" border="0"></b></font></a><font face="Arial" size="2" color="black"><b> 
            &nbsp;&nbsp;</b></font><a href="makeofferlist.asp"><font face="Arial" size="2" color="black"><b>Back to Offers List</b></font></a><font face="Arial" size="2" color="black"><b><br>&nbsp;</b></font></p>
        </td>
    </tr>
</table>
<!--#include file="footer.asp"-->
