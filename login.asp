<!--#include file="db.asp"-->
<!--#include file="header.asp"-->
<%
If Request.Form("submit") <> "" Then
	validpwd = False
	' setup variables
	userid = Request.Form("userid")
	passwd = Request.Form("passwd")
    If not validpwd Then
			Set conn = Server.CreateObject("ADODB.Connection")
			conn.open xDb_Conn_Str
			Set rs = conn.Execute( "Select * from [info] where [user] = '" & userid & "'")
			If Not rs.EOF Then
				If UCase(rs("password")) = UCase(passwd) Then
					 Session("rocar_status_User") = rs("user")
					validpwd = True
				End If
			End If
			rs.Close
			Set rs = Nothing
			conn.Close
			Set conn = Nothing
	End If
	If validpwd Then
		' write cookies
		If Request.Form("rememberme") <> "" Then
			Response.Cookies("rocar")("userid") = userid
			Response.Cookies("rocar").Expires = Date + 365 'change the expiry date of the cookies here
		End If		
		Session("rocar_status") = "login"
		Response.Redirect "adminlist.asp"
	End If
Else
	validpwd = True
End If
%>
<html>
<head>
	<title>Rocar Auto Sales 773-334-0025</title>	
<meta name="generator" content="Namo WebEditor v5.0(Trial)">
<script language="JavaScript" src="ew.js"></script>
<script language="JavaScript">
<!-- start JavaScript
function  EW_checkMyForm(EW_this) {
if  (!EW_hasValue(EW_this.userid, "TEXT" )) {
            if  (!EW_onError(EW_this, EW_this.userid, "TEXT", "Please enter user ID"))
                return false; 
        }
if  (!EW_hasValue(EW_this.passwd, "PASSWORD" )) {
            if  (!EW_onError(EW_this, EW_this.passwd, "PASSWORD", "Please enter password"))
                return false; 
        }
return true;
}
// end JavaScript -->
</script>
</head>
<body leftmargin=0 topmargin=0 marginheight=0 marginwidth=0>
 
<table cellpadding="0" cellspacing="0" width="802" bgcolor="white" align="center">
    <tr>
        <td>
            <p><% If Not validpwd Then %>
            </p>
<p align="left"><font face="Arial" size="2" color="red"><b><i>&nbsp;&nbsp;Incorrect user ID or password</i></b></font></p>
<% End If %>
<form action="login.asp" method="post" onSubmit="return EW_checkMyForm(this);">
                <div align="left">
<table border="0" cellspacing="0" cellpadding="4" width="218">
	<tr>
		<td align="left" width="67"><input type="text" name="userid" size="10" value="<%= request.Cookies("rocar")("userid") %>"></td>
		<td width="80"><input type="password" name="passwd" size="10"></td>
		<td width="47"><input type="submit" name="submit" value="Login"></td>
	</tr>
	<tr>
		<td align="left" width="155" colspan="2"><input type="checkbox" name="rememberme" value="true"><font face="Arial"><span style="font-size:10pt;">Remember 
                                Me</span></font></td>
		<td width="47">
                <p>&nbsp;</p>
</td>
	</tr>	
</table>
                </div>
</form>
        </td>
    </tr>
</table>
</body>
<!--#include file="footer.asp"-->
</html>
