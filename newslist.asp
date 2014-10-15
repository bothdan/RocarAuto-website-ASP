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
displayRecs = 20
recRange = 10
%>
<%
dbwhere = ""
masterdetailwhere = ""
searchwhere = ""
a_search = ""
b_search = ""
whereClause = ""
%>
<%
'Get clear search cmd
If Request.QueryString("cmd").Count > 0 Then
	cmd = Request.QueryString("cmd")
	If UCase(cmd) = "RESET" Then
		'reset search criteria
		searchwhere = ""
		Session("news_searchwhere") =searchwhere
    ElseIf UCase(cmd) = "RESETALL" Then
		'reset search criteria
		searchwhere = ""
		Session("news_searchwhere") =searchwhere
	End If
	'reset start record counter (reset command)
	startRec = 1
	Session("news_REC") = startRec
End If
'construct dbwhere
If masterdetailwhere <> "" Then
	dbwhere = dbwhere & "(" & masterdetailwhere & ") AND "
End If
If searchwhere <> "" Then
	dbwhere = dbwhere & "(" & searchwhere & ") AND "
End If
If Len(dbwhere) > 5 Then
	dbwhere = Mid(dbwhere, 1, Len(dbwhere)-5) 'trim right most AND
End If
%>
<%
' Load Default Order
DefaultOrder = ""
DefaultOrderType = ""
'No Default Filter
DefaultFilter = ""
' Check for an Order parameter
OrderBy = ""
If Request.QueryString("order").Count > 0 Then
	OrderBy = Request.QueryString("order")
	' Check If an ASC/DESC toggle is required
	If Session("news_OB") = OrderBy Then
		If Session("news_OT") = "ASC" Then
			Session("news_OT") = "DESC"
		Else
			Session("news_OT") = "ASC"
		End if
	Else
		Session("news_OT") = "ASC"
	End If
	Session("news_OB") = OrderBy
	Session("news_REC") = 1
Else
	OrderBy = Session("news_OB")
	If OrderBy = "" Then
		OrderBy = DefaultOrder
		Session("news_OB") = OrderBy
		Session("news_OT") = DefaultOrderType
	End If
End If
' Open Connection to the database
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str
' Build SQL
strsql = "SELECT * FROM [news]"
If DefaultFilter <> "" Then
	whereClause = whereClause & "(" & DefaultFilter & ") AND "
End If
If dbwhere <> "" Then
	whereClause = whereClause & "(" & dbwhere & ") AND "
End If
If Right(whereClause, 5)=" AND " Then whereClause = Left(whereClause, Len(whereClause)-5)
If whereClause <> "" Then
	strsql = strsql & " WHERE " & whereClause
End If
If OrderBy <> "" Then 
	strsql = strsql & " ORDER BY [" & OrderBy & "] " & Session("news_OT")
End If	
'Response.Write strsql
Set rs = Server.CreateObject("ADODB.Recordset")
rs.cursorlocation = 3
rs.Open strsql, conn, 1, 2
totalRecs = rs.RecordCount
' Check for a START parameter
If Request.QueryString("start").Count > 0 Then
	startRec = Request.QueryString("start")
	Session("news_REC") = startRec
ElseIf Request.QueryString("pageno").Count > 0 Then
	pageno = Request.QueryString("pageno")
	If IsNumeric(pageno) Then
		startRec = (pageno-1)*displayRecs+1
		If startRec <= 0 Then
			startRec = 1
		ElseIf startRec >= ((totalRecs-1)\displayRecs)*displayRecs+1 Then
			startRec = ((totalRecs-1)\displayRecs)*displayRecs+1
		End If
		Session("news_REC") = startRec
	Else
		startRec = Session("news_REC")
		If Not IsNumeric(startRec) Or startRec = "" Then
			'reset start record counter
			startRec = 1
			Session("news_REC") = startRec
		End If
	End If
Else
	startRec = Session("news_REC")
	If Not IsNumeric(startRec) Or startRec = "" Then
		'reSet start record counter
		startRec = 1
		Session("news_REC") = startRec
	End If
End If
%>
<!--#include file="header.asp"-->
<meta name="generator" content="Namo WebEditor v5.0(Trial)">
<table align="center" cellpadding="0" cellspacing="0" width="801" bgcolor="white">
    <tr>
        <td height="11">

<div align="left">
    <table cellpadding="0" cellspacing="0" bgcolor="white" width="360">
        <tr>
            <td width="20" height="17">
                <p></p>
            </td>
            <td width="5" bgcolor="white" height="17">
                            <p align="center"><span style="font-size:12pt;"><font face="Arial">|</font></span></p>
            </td>
            <td width="100" bgcolor="white" height="17">
                <p align="center"><span style="font-size:10pt;"><a href="adminlist.asp"><b><font face="Arial" color="navy">Inventory</font></b></a></span></p>
            </td>
            <td width="5" height="17" bgcolor="white">
                            <p align="center"><span style="font-size:12pt;"><font face="Arial">|</font></span></p>
            </td>
            <td width="100" height="17">
                <p align="center"><a href="contactlist.asp"><b><span style="font-size:10pt;"><font face="Arial" color="navy">Customers</font></span></b></a></p>
            </td>
            <td width="5" height="17">
                            <p align="center"><span style="font-size:12pt;"><font face="Arial">|</font></span></p>
            </td>
            <td width="120" height="17">
                <p align="center"><b><span style="font-size:10pt;"><font face="Arial" color="green">News 
                            And Events</font></span></b></p>
            </td>
            <td width="5" height="17">
                            <p align="center"><span style="font-size:12pt;"><font face="Arial">|</font></span></p>
            </td>
        </tr>
    </table>
</div>
        </td>
    </tr>
    <tr>
        <td>
            <p><font face="Arial" color="black"><b><span style="font-size:14pt;"><br></span><span style="font-size:11pt;"> 
            &nbsp;&nbsp;&nbsp;<i>News 
            And Events:</i></span></b></font></p>
<form method="post">
                <div align="left">
<table border="0" cellspacing="0" cellpadding="0" bgcolor="#CCCCCC" width="286">
<tr bgcolor="#708090">
<td width="20" bgcolor="white">
                                <p><font face="Arial"><span style="font-size:12pt;">&nbsp;</span></font></p>
</td>
<td width="120" bgcolor="white">
                                <p align="center"><a href="<% key = rs("ID") : If not isnull(key) Then response.write "newsview.asp?key=" & Server.URLEncode(key) Else response.write "javascript:alert('Invalid Record! Key is null');" End If %>"><font face="Arial"><span style="font-size:12pt;">View</span></font></a></td>
<td width="146" bgcolor="white">
                                <p align="center"><a href="<% key = rs("ID") : If not isnull(key) Then response.write "newsedit.asp?key=" & Server.URLEncode(key) Else response.write "javascript:alert('Invalid Record! Key is null');" End If %>"><font face="Arial"><span style="font-size:12pt;">Edit</span></font></a></td>
</tr>
<%
	'End If
	rs.MoveNext
'Loop
%>
</table>
                </div>
</form>
            <p><% If recActual > 0 Then %></p>
<form method="post">

                <p><% End If %>
</form>
<%
' Close recordSet and connection
rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing %>
        </td>
    </tr>
</table>
<!--#include file="footer.asp"-->
