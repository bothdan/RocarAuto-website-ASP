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
If key = "" OR IsNull(key) Then key = Request.Form("key")
If key = "" OR IsNull(key) Then Response.Redirect "contactlist.asp"
'get action
a=Request.Form("a")
If a="" OR IsNull(a) Then
	a="I"	'display with input box
End If
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
End Select
%>
<!--#include file="header.asp"-->
<meta name="generator" content="Namo WebEditor v5.0(Trial)">
<table align="center" cellpadding="0" cellspacing="0" width="801" bgcolor="white">
    <tr>
        <td>
            <p><font face="Arial" size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br> 
            </font></p>
<p>
<form>
<table border="0" cellspacing="0" cellpadding="0" bgcolor="white" align="center" width="669">
<tr>
<td bgcolor="white" width="126"><font face="Arial"><b><span style="font-size:14pt;">View 
                            Contact<br>&nbsp;</span></b></font></td>
<td bgcolor="white" width="543"><font face="Arial" color="white"><span style="font-size:10pt;"><%= x_ID %>&nbsp;</span></font></td>
</tr>
<tr>
<td bgcolor="white" width="126" height="25"><font face="Arial"><span style="font-size:10pt;"><b>First Name&nbsp;</b></span></font></td>
<td bgcolor="white" width="543" height="25"><font face="Arial"><span style="font-size:10pt;"><% response.write x_first_name %>&nbsp;</span></font></td>
</tr>
<tr>
<td bgcolor="white" width="126" height="25"><font face="Arial"><span style="font-size:10pt;"><b>Last Name&nbsp;</b></span></font></td>
<td bgcolor="white" width="543" height="25"><font face="Arial"><span style="font-size:10pt;"><% response.write x_last_name %>&nbsp;</span></font></td>
</tr>
<tr>
<td bgcolor="white" width="126" height="25"><font face="Arial"><span style="font-size:10pt;"><b>E-mail&nbsp;</b></span></font></td>
<td bgcolor="white" width="543" height="25"><font face="Arial"><span style="font-size:10pt;"><% response.write x_email %>&nbsp;</span></font></td>
</tr>
<tr>
<td bgcolor="white" width="126" height="25"><font face="Arial"><span style="font-size:10pt;"><b>Phone&nbsp;</b></span></font></td>
<td bgcolor="white" width="543" height="25"><font face="Arial"><span style="font-size:10pt;"><% response.write x_phone %>&nbsp;</span></font></td>
</tr>
<tr>
<td bgcolor="white" width="126" height="25"><font face="Arial"><span style="font-size:10pt;"><b>Comments&nbsp;</b></span></font></td>
<td bgcolor="white" width="543" height="25"><font face="Arial"><span style="font-size:10pt;"><%= replace(x_comments & "",chr(10),"<br>") %>&nbsp;</span></font></td>
</tr>
</table>
</form>
            <p><font face="Arial" size="2" color="black"><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</b></font><a href="contactlist.asp"><font face="Arial" size="2" color="black"><b><img src="images/leftsm.gif" align="texttop" width="16" height="16" border="0"></b></font></a><font face="Arial" size="2" color="black"><b> 
            &nbsp;&nbsp;</b></font><a href="contactlist.asp"><font face="Arial" size="2" color="black"><b>Back to 
            Contact</b></font></a><font face="Arial" size="2"><br>&nbsp;</font></p>
        </td>
    </tr>
</table>
<!--#include file="footer.asp"-->
