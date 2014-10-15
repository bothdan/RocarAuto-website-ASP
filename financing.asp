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
		Session("credit_searchwhere") =searchwhere
    ElseIf UCase(cmd) = "RESETALL" Then
		'reset search criteria
		searchwhere = ""
		Session("credit_searchwhere") =searchwhere
	End If
	'reset start record counter (reset command)
	startRec = 1
	Session("credit_REC") = startRec
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
	If Session("credit_OB") = OrderBy Then
		If Session("credit_OT") = "ASC" Then
			Session("credit_OT") = "DESC"
		Else
			Session("credit_OT") = "ASC"
		End if
	Else
		Session("credit_OT") = "ASC"
	End If
	Session("credit_OB") = OrderBy
	Session("credit_REC") = 1
Else
	OrderBy = Session("credit_OB")
	If OrderBy = "" Then
		OrderBy = DefaultOrder
		Session("credit_OB") = OrderBy
		Session("credit_OT") = DefaultOrderType
	End If
End If
' Open Connection to the database
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str
' Build SQL
strsql = "SELECT * FROM [credit]"
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
	strsql = strsql & " ORDER BY [" & OrderBy & "] " & Session("credit_OT")
End If	
'Response.Write strsql
Set rs = Server.CreateObject("ADODB.Recordset")
rs.cursorlocation = 3
rs.Open strsql, conn, 1, 2
totalRecs = rs.RecordCount
' Check for a START parameter
If Request.QueryString("start").Count > 0 Then
	startRec = Request.QueryString("start")
	Session("credit_REC") = startRec
ElseIf Request.QueryString("pageno").Count > 0 Then
	pageno = Request.QueryString("pageno")
	If IsNumeric(pageno) Then
		startRec = (pageno-1)*displayRecs+1
		If startRec <= 0 Then
			startRec = 1
		ElseIf startRec >= ((totalRecs-1)\displayRecs)*displayRecs+1 Then
			startRec = ((totalRecs-1)\displayRecs)*displayRecs+1
		End If
		Session("credit_REC") = startRec
	Else
		startRec = Session("credit_REC")
		If Not IsNumeric(startRec) Or startRec = "" Then
			'reset start record counter
			startRec = 1
			Session("credit_REC") = startRec
		End If
	End If
Else
	startRec = Session("credit_REC")
	If Not IsNumeric(startRec) Or startRec = "" Then
		'reSet start record counter
		startRec = 1
		Session("credit_REC") = startRec
	End If
End If
%>
<!--#include file="header.asp"-->
<meta name="generator" content="Namo WebEditor v5.0(Trial)">

<table align="center" cellpadding="0" cellspacing="0" width="801" bgcolor="white">
    <tr>
        <td background="images/contactbg.gif">
            <p><img src="images/financebg.gif" width="312" height="32" border="0"></p>
        </td>
    </tr>
    <tr>
        <td>


<p><font face="Arial"><span style="font-size:10pt;">&nbsp;</span></font></p>

            <p align="center"><font face="Arial"><span style="font-size:10pt;"><img src="images/dollarsign.jpg" align="middle" width="30" height="42" border="0"> 
            To apply for a car loan please browse thru our inventory and follow the link
found inside the vehicle information page .<img src="images/dollarsign.jpg" align="middle" width="30" height="42" border="0"></span></font></p>
            <p><font face="Arial"><span style="font-size:10pt;">&nbsp;</span></font></p>
        </td>
    </tr>
</table>
<!--#include file="footer.asp"-->