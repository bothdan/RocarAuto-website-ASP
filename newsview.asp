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
If key = "" OR IsNull(key) Then Response.Redirect "newslist.asp"
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
		strsql = "SELECT * FROM [news] WHERE [ID]=" & tkey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn
		If rs.EOF Then
			Response.Clear
			Response.Redirect "newslist.asp"
		Else
			rs.MoveFirst
		End If
		' Get the field contents
		x_ID = rs("ID")
		x_news = rs("news")
		rs.Close
		Set rs = Nothing
End Select
%>
<!--#include file="header.asp"-->
<meta name="generator" content="Namo WebEditor v5.0(Trial)">
<table align="center" cellpadding="0" cellspacing="0" width="801" bgcolor="white">
    <tr>
        <td background="images/contactbg.gif">
            <p><img src="images/newsbg.gif" width="324" height="32" border="0"></p>
        </td>
    </tr>
    <tr>
        <td>
<p>
&nbsp;<form>
                <div align="left">
<table border="0" cellspacing="0" cellpadding="0" bgcolor="#CCCCCC" width="761">
<tr>
<td bgcolor="white" width="280">
                                <p><img src="images/redcar.gif" width="280" height="306" border="0"></p>
</td>
<td bgcolor="white" width="181" valign="top"><font face="Arial"><span style="font-size:14pt;"><b>The 
                                Latest News:</b></span></font></td>
<td bgcolor="white" width="300" valign="top"><font face="Arial"><span style="font-size:14pt;"><%= replace(x_news & "",chr(10),"<br>") %>&nbsp;</span></font></td>
</tr>
</table>
                </div>
</form>
        </td>
    </tr>
</table>
<!--#include file="footer.asp"-->
