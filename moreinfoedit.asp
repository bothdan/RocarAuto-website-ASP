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
If key="" OR IsNull(key) Then Response.Redirect "moreinfolist.asp"
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
x_testdrive = Request.Form("x_testdrive")
x_comments = Request.Form("x_comments")
x_stock = Request.Form("x_stock")
' Open Connection to the database
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str
Select Case a
	Case "I": ' Get a record to display
		tkey = "" & key & ""
		strsql = "SELECT * FROM [moreinfo] WHERE [ID]=" & tkey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn
		If rs.EOF Then
			Response.Clear
			Response.Redirect "moreinfolist.asp"
		Else
			rs.MoveFirst
		End If
		' Get the field contents
		x_ID = rs("ID")
		x_first = rs("first")
		x_last = rs("last")
		x_phone = rs("phone")
		x_email = rs("email")
		x_testdrive = rs("testdrive")
		x_comments = rs("comments")
		x_stock = rs("stock")
		rs.Close
		Set rs = Nothing
	Case "U": ' Update
		' Open record
		tkey = "" & key & ""
		strsql = "SELECT * FROM [moreinfo] WHERE [ID]=" & tkey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn, 1, 2
		If rs.EOF Then
			Response.Clear
			Response.Redirect "moreinfolist.asp"
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
		Response.Redirect "moreinfolist.asp"
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
<form onSubmit="return EW_checkMyForm(this);"  action="moreinfoedit.asp" method="post">
<p>
<input type="hidden" name="a" value="U">
<input type="hidden" name="key" value="<%= key %>">
<table border="0" cellspacing="0" cellpadding="0" bgcolor="#CCCCCC" width="490">
<tr>
<td bgcolor="white" width="51">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="144"><font face="Arial"><span style="font-size:14pt;"><b>Edit 
                            More Info:<br>&nbsp;</b></span></font></td>
<td bgcolor="white" width="295"><font face="Arial" color="white"><span style="font-size:10pt;"><%= x_ID %><input type="hidden" name="x_ID" value="<%= x_ID %>">&nbsp;</span></font></td>
</tr>
<tr>
<td bgcolor="white" width="51">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="144"><font face="Arial"><span style="font-size:10pt;"><b>First 
                            Name:</b></span></font></td>
<td bgcolor="white" width="295"><font face="Arial"><span style="font-size:10pt;"><input type="text" name="x_first" size=30 maxlength=50 value="<%= Server.HtmlEncode(x_first&"") %>">&nbsp;</span></font></td>
</tr>
<tr>
<td bgcolor="white" width="51">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="144"><font face="Arial"><span style="font-size:10pt;"><b>Last 
                            Name:</b></span></font></td>
<td bgcolor="white" width="295"><font face="Arial"><span style="font-size:10pt;"><input type="text" name="x_last" size=30 maxlength=50 value="<%= Server.HtmlEncode(x_last&"") %>">&nbsp;</span></font></td>
</tr>
<tr>
<td bgcolor="white" width="51">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="144"><font face="Arial"><span style="font-size:10pt;"><b>Phone:</b></span></font></td>
<td bgcolor="white" width="295"><font face="Arial"><span style="font-size:10pt;"><input type="text" name="x_phone" size=30 maxlength=50 value="<%= Server.HtmlEncode(x_phone&"") %>">&nbsp;</span></font></td>
</tr>
<tr>
<td bgcolor="white" width="51">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="144"><font face="Arial"><span style="font-size:10pt;"><b>E-mail:</b></span></font></td>
<td bgcolor="white" width="295"><font face="Arial"><span style="font-size:10pt;"><input type="text" name="x_email" size=30 maxlength=50 value="<%= Server.HtmlEncode(x_email&"") %>">&nbsp;</span></font></td>
</tr>
<tr>
<td bgcolor="white" width="51">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="144"><font face="Arial"><span style="font-size:10pt;"><b>Test 
                            Drive:</b></span></font></td>
<td bgcolor="white" width="295"><font face="Arial"><span style="font-size:10pt;"><%
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
&nbsp;</span></font></td>
</tr>
<tr>
<td bgcolor="white" width="51">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="144"><font face="Arial"><span style="font-size:10pt;"><b>Comments:</b></span></font></td>
<td bgcolor="white" width="295"><font face="Arial"><span style="font-size:10pt;"><textarea cols=35 rows=4 name="x_comments"><%= x_comments %></textarea>&nbsp;</span></font></td>
</tr>
<tr>
<td bgcolor="white" width="51">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="144"><font face="Arial"><span style="font-size:10pt;"><b>Stock#:</b></span></font></td>
<td bgcolor="white" width="295"><font face="Arial"><span style="font-size:10pt;"><input type="text" name="x_stock" size=30 maxlength=50 value="<%= Server.HtmlEncode(x_stock&"") %>">&nbsp;</span></font></td>
</tr>
</table>
<p align="center">
<input type="submit" name="Action" value="Update">
</form>
            <p><font face="Arial" size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font><a href="moreinfolist.asp"><font face="Arial" size="2" color="black"><b><img src="images/leftsm.gif" align="texttop" width="16" height="16" border="0"></b></font></a><font face="Arial" size="2" color="black"><b> 
            &nbsp;&nbsp;</b></font><a href="moreinfolist.asp"><font face="Arial" size="2" color="black"><b>Back to More Info List</b></font></a><font face="Arial" size="2" color="black"><b><br>&nbsp;</b></font></p>
        </td>
    </tr>
</table>
<!--#include file="footer.asp"-->
