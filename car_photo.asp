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
<%
Response.buffer = True
'get key
key = Request.QueryString("key")
If key = "" or isnull(key) Then
	response.redirect "carlist.asp"
	response.end
End If
Open Connection to the database
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str
tkey = key
strsql = "SELECT * FROM [car] WHERE [ID]=" & tkey

Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open strsql, conn
If Not rs.EOF Then
	'rs.MoveFirst
	 Response.BinaryWrite rs("photo 1")
End If
rs.Close
Set rs = nothing
%>
<!--#include file="car_photo_2_bv.asp"-->
<meta name="generator" content="Namo WebEditor v5.0(Trial)">
<p>
&nbsp;