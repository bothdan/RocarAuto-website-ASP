
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
nr = Request.QueryString("nr")

dim serverVar
serverVar = request(j)
if not isnull(j) then
response.write j

End If 

If key = "" or isnull(key) Then
	response.redirect "carlist.asp"
	response.end
End If
' Open Connection to the database
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str
tkey = key
strsql = "SELECT * FROM [car] WHERE [ID]=" & tkey
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open strsql, conn
If Not rs.EOF Then
	'rs.MoveFirst
	ph = "photo " & nr
	 Response.BinaryWrite rs(ph)
	
End If
rs.Close
Set rs = nothing
%>

<meta name="generator" content="Namo WebEditor v5.0(Trial)">
<p>
&nbsp;