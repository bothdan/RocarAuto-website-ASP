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

<!--#include file="header.asp"-->
<meta name="generator" content="Namo WebEditor v5.0(Trial)">
<body text="black" link="blue" vlink="purple" alink="red">
<table align="center" cellpadding="0" cellspacing="0" width="801" bgcolor="white">
    <tr>
        <td background="images/contactbg.gif">
            <p><img src="images/specialsbg.gif" width="324" height="32" border="0"></p>
        </td>
    </tr>
    <tr>
        <td>
            <p align="center"><font face="Arial"><span style="font-size:14pt;">&nbsp;</span></font></p>
            <p align="center"><br><font face="Arial"><span style="font-size:14pt;">&nbsp;We 
            are sorry, but there are no specials at this moment.<br>Please try 
            again later.<br></span></font></p>
            <p align="center"><font face="Arial"><span style="font-size:14pt;">Thank 
            You!<br>&nbsp;</span></font></p>
            <p align="center"><font face="Arial"><span style="font-size:14pt;">&nbsp;</span></font></p>
        </td>
    </tr>
</table>
<!--#include file="footer.asp"-->