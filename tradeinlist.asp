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
displayRecs = 50
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
		Session("tradein_searchwhere") =searchwhere
    ElseIf UCase(cmd) = "RESETALL" Then
		'reset search criteria
		searchwhere = ""
		Session("tradein_searchwhere") =searchwhere
	End If
	'reset start record counter (reset command)
	startRec = 1
	Session("tradein_REC") = startRec
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
DefaultOrder = "tradeindate"
DefaultOrderType = "DESC"
'No Default Filter
DefaultFilter = ""
' Check for an Order parameter
OrderBy = ""

If Request.QueryString("order").Count > 0 Then
	OrderBy = Request.QueryString("order")
	' Check If an ASC/DESC toggle is required
	If Session("tradein_OB") = OrderBy Then
		If Session("tradein_OT") = "ASC" Then
			Session("tradein_OT") = "DESC"
		Else
			Session("tradein_OT") = "ASC"
		End if
	Else
		Session("tradein_OT") = "ASC"
	End If
	Session("tradein_OB") = OrderBy
	Session("tradein_REC") = 1
Else
	OrderBy = Session("tradein_OB")
	If OrderBy = "" Then
		OrderBy = DefaultOrder
		Session("tradein_OB") = OrderBy
		Session("tradein_OT") = DefaultOrderType
	End If
End If
' Open Connection to the database
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str
' Build SQL
strsql = "SELECT * FROM [tradein]"
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
	strsql = strsql & " ORDER BY [" & OrderBy & "] " & Session("tradein_OT")
End If	
'Response.Write strsql
Set rs = Server.CreateObject("ADODB.Recordset")
rs.cursorlocation = 3
rs.Open strsql, conn, 1, 2
totalRecs = rs.RecordCount
' Check for a START parameter
If Request.QueryString("start").Count > 0 Then
	startRec = Request.QueryString("start")
	Session("tradein_REC") = startRec
ElseIf Request.QueryString("pageno").Count > 0 Then
	pageno = Request.QueryString("pageno")
	If IsNumeric(pageno) Then
		startRec = (pageno-1)*displayRecs+1
		If startRec <= 0 Then
			startRec = 1
		ElseIf startRec >= ((totalRecs-1)\displayRecs)*displayRecs+1 Then
			startRec = ((totalRecs-1)\displayRecs)*displayRecs+1
		End If
		Session("tradein_REC") = startRec
	Else
		startRec = Session("tradein_REC")
		If Not IsNumeric(startRec) Or startRec = "" Then
			'reset start record counter
			startRec = 1
			Session("tradein_REC") = startRec
		End If
	End If
Else
	startRec = Session("tradein_REC")
	If Not IsNumeric(startRec) Or startRec = "" Then
		'reSet start record counter
		startRec = 1
		Session("tradein_REC") = startRec
	End If
End If
%>



<!--#include file="header.asp"-->
<meta name="generator" content="Namo WebEditor v5.0(Trial)">
<body bgcolor="white" text="black" link="green" vlink="black" alink="red">
<table align="center" cellpadding="0" cellspacing="0" width="801" bgcolor="white">
    <tr>
        <td height="2">
            <div align="left">
                <table cellpadding="0" cellspacing="0" bgcolor="white" width="360">
                    <tr>
                        <td width="20" height="16">
                            <p></p>
                        </td>
                        <td width="5" height="16">
                <p align="center"><font face="Arial"><span style="font-size:12pt;">|</span></font></p>
                        </td>
                        <td width="100" bgcolor="white" height="16">
                            <p align="center"><span style="font-size:10pt;"><a href="adminlist.asp"><b><font face="Arial" color="navy">Inventory</font></b></a></span></p>
                        </td>
                        <td width="5" bgcolor="white" height="16">
                <p align="center"><font face="Arial"><span style="font-size:12pt;">|</span></font></p>
                        </td>
                        <td width="100" height="16">
                            <p align="center"><b><span style="font-size:10pt;"><font face="Arial" color="green">Customers</font></span></b></p>
                        </td>
                        <td width="5" height="16">
                <p align="center"><span style="font-size:12pt;"><font face="Arial">|</font></span></p>
                        </td>
                        <td width="120" height="16">
                            <p align="center"><a href="newslist.asp"><b><span style="font-size:10pt;"><font face="Arial" color="navy">News 
                            And Events</font></span></b></a></p>
                        </td>
                        <td width="5" height="16">
                <p align="center"><span style="font-size:12pt;"><font face="Arial">|</font></span></p>
                        </td>
                    </tr>
                </table>
</div>
            
</td>
    </tr>
    <tr>
        <td height="8">

<div align="left">
    <table cellpadding="0" cellspacing="0" width="695" bgcolor="white">
        <tr>
            <td width="20" height="5">
                <p></p>
            </td>
            <td width="5" height="5" bgcolor="white">
                <p align="center"></p>
            </td>
            <td width="100" height="5" bgcolor="#979EE0">
                <p align="center"><a href="contactlist.asp"><span style="font-size:10pt;"><b><font face="Arial" color="white">Contact</font></b></span></a></p>
            </td>
            <td width="5" height="5" bgcolor="white">
                <p align="center"></p>
            </td>
            <td width="100" height="5" bgcolor="#979EE0">
                <p align="center"><span style="font-size:10pt;"><a href="makeofferlist.asp"><b><font face="Arial" color="white">Offers</font></b></a></span></p>
            </td>
            <td width="5" height="5" bgcolor="white">
                <p align="center"><span style="font-size:10pt;"><font face="Arial" color="white"><b>|</b></font></span></p>
            </td>
            <td width="120" height="5" bgcolor="#979EE0">
                <p align="center"><a href="moreinfolist.asp"><span style="font-size:10pt;"><font face="Arial" color="white"><b>More 
                Info</b></font></span></a></p>
            </td>
            <td width="5" height="5" bgcolor="white">
                <p align="center"><span style="font-size:10pt;"><font face="Arial" color="white"><b>|</b></font></span></p>
            </td>
            <td width="100" height="5" bgcolor="#979EE0">
                <p align="center"><a href="findelist.asp"><span style="font-size:10pt;"><font face="Arial" color="white"><b>Find 
                Car</b></font></span></a></p>
            </td>
            <td width="5" height="5" bgcolor="white">
                <p align="center"><span style="font-size:10pt;"><font face="Arial" color="white"><b>|</b></font></span></p>
            </td>
            <td width="120" height="5" bgcolor="#979EE0">
                <p align="center"><a href="creditlist.asp"><span style="font-size:10pt;"><font face="Arial" color="white"><b>Credit 
                Application</b></font></span></a></p>
            </td>
            <td width="5" height="5" bgcolor="white">
                <p align="center"><span style="font-size:10pt;"><font face="Arial" color="white"><b>|</b></font></span></p>
            </td>
            <td width="100" height="5" bgcolor="#6351B7">
                <p align="center"><span style="font-size:10pt;"><font face="Arial" color="white"><b>Trade-In</b></font></span></p>
            </td>
            <td width="5" height="5">
                <p align="center"></p>
            </td>
        </tr>
    </table>
</div>
</td>
    </tr>
    <tr>
        <td>
            <form method="post">
                <p align="left"><font face="Arial"><span style="font-size:11pt;"><b><i><br>&nbsp;&nbsp;&nbsp;Trade-In 
                List:<br></i></b></span></font></p>
<table border="0" cellspacing="0" cellpadding="0" bgcolor="#CCCCCC" width="750" align="center">
<tr bgcolor="#708090">
<td width="62" bgcolor="white" background="images/sortbar.gif">

<p><font color="black" face="Verdana"><b><span style="font-size:8pt;">&nbsp;</span></b></font><a href='tradeinlist.asp?order=<%= Server.URLEncode("tradeindate") %>'><font color="black" face="Arial"><b><span style="font-size:8pt;">Date</span></b></font></a><font color="black" face="Arial"><span style="font-size:8pt;"><b>&nbsp;</b></span></font><font color="black"><span style="font-size:8pt;"><b><% If OrderBy = "tradeindate" Then %></b></span></font><font color="black" face="Webdings"><span style="font-size:8pt;"><b><% If Session("tradein_OT") = "ASC" Then %>5<% ElseIf Session("tradein_OT") = "DESC" Then %>6<% End If %></b></span></font><font color="black"><span style="font-size:8pt;"><b><% End If %></b></span></font>
</td>
<td width="57" bgcolor="white" background="images/sortbar.gif">
                            <p align="center"><a href="tradeinlist.asp?order=<%= Server.URLEncode("year") %>"><font face="Arial" color="black"><b><span style="font-size:8pt;">Year</span></b></font></a><font face="Arial" color="black"><b><span style="font-size:8pt;">&nbsp;</span></b></font><font color="black"><b><span style="font-size:8pt;"><% If OrderBy = "year" Then %></span></b></font><font face="Webdings" color="black"><b><span style="font-size:8pt;"><% If Session("tradein_OT") = "ASC" Then %>5<% ElseIf Session("tradein_OT") = "DESC" Then %>6<% End If %></span></b></font><font color="black"><b><span style="font-size:8pt;"><% End If %></span></b></font>
</td>
<td width="221" bgcolor="white" background="images/sortbar.gif">
<a href="tradeinlist.asp?order=<%= Server.URLEncode("make") %>"><font face="Arial" color="black"><b><span style="font-size:8pt;">Make&amp;Model</span></b></font></a><font face="Arial" color="black"><b><span style="font-size:8pt;">&nbsp;</span></b></font><font color="black"><b><span style="font-size:8pt;"><% If OrderBy = "make" Then %></span></b></font><font face="Webdings" color="black"><b><span style="font-size:8pt;"><% If Session("tradein_OT") = "ASC" Then %>5<% ElseIf Session("tradein_OT") = "DESC" Then %>6<% End If %></span></b></font><font color="black"><b><span style="font-size:8pt;"><% End If %></span></b></font>
</td>
<td width="100" bgcolor="white" background="images/sortbar.gif">
                            <p align="center"><a href="tradeinlist.asp?order=<%= Server.URLEncode("mileage") %>"><font face="Arial" color="black"><b><span style="font-size:8pt;">Mileage</span></b></font></a><font face="Arial" color="black"><b><span style="font-size:8pt;">&nbsp;</span></b></font><font color="black"><b><span style="font-size:8pt;"><% If OrderBy = "mileage" Then %></span></b></font><font face="Webdings" color="black"><b><span style="font-size:8pt;"><% If Session("tradein_OT") = "ASC" Then %>5<% ElseIf Session("tradein_OT") = "DESC" Then %>6<% End If %></span></b></font><font color="black"><b><span style="font-size:8pt;"><% End If %></span></b></font>
</td>
<td width="100" bgcolor="white" background="images/sortbar.gif">
                            <p align="center"><a href="tradeinlist.asp?order=<%= Server.URLEncode("transmission") %>"><font face="Arial" color="black"><b><span style="font-size:8pt;">Transmission</span></b></font></a><font face="Arial" color="black"><b><span style="font-size:8pt;">&nbsp;</span></b></font><font color="black"><b><span style="font-size:8pt;"><% If OrderBy = "transmission" Then %></span></b></font><font face="Webdings" color="black"><b><span style="font-size:8pt;"><% If Session("tradein_OT") = "ASC" Then %>5<% ElseIf Session("tradein_OT") = "DESC" Then %>6<% End If %></span></b></font><font color="black"><b><span style="font-size:8pt;"><% End If %></span></b></font>
</td>
<td width="112" bgcolor="white" background="images/sortbar.gif">
                            <p align="center"><a href="tradeinlist.asp?order=<%= Server.URLEncode("stock") %>"><font face="Arial" color="black"><b><span style="font-size:8pt;">Stock</span></b></font></a><font face="Arial" color="black"><b><span style="font-size:8pt;">&nbsp;</span></b></font><font color="black"><b><span style="font-size:8pt;"><% If OrderBy = "stock" Then %></span></b></font><font face="Webdings" color="black"><b><span style="font-size:8pt;"><% If Session("tradein_OT") = "ASC" Then %>5<% ElseIf Session("tradein_OT") = "DESC" Then %>6<% End If %></span></b></font><font color="black"><b><span style="font-size:8pt;"><% End If %></span></b></font>
</td>
<td width="136" colspan="3" bgcolor="white" background="images/sortbar.gif"><font color="black"><span style="font-size:8pt;"><b>&nbsp;</b></span></font></td>
</tr>
<%
'Avoid starting record > total records
If CLng(startRec) > CLng(totalRecs) Then
	startRec = totalRecs
End If
'Set the last record to display
stopRec = startRec + displayRecs - 1
'Move to first record directly for performance reason
recCount = startRec - 1
If NOT rs.EOF Then
	rs.MoveFirst
	rs.Move startRec - 1
End If
recActual = 0
Do While (NOT rs.EOF) AND (recCount < stopRec)
	recCount = recCount + 1
	If CLng(recCount) >= CLng(startRec) Then 
		recActual = recActual + 1 %>
<%
	'Set row color
	bgcolor="#FFFFFF"
%>
<%	
	' Display alternate color for rows
	If recCount Mod 2 <> 0 Then
		bgcolor="#F5F5F5"
	End If
%>
<%
	'Load Key for record
	key = rs("ID")
	x_ID = rs("ID")
	x_first = rs("first")
	x_last = rs("last")
	x_home_phone = rs("home phone")
	x_work_phone = rs("work phone")
	x_email = rs("email")
	x_year = rs("year")
	x_make = rs("make")
	x_model = rs("model")
	x_ext_color = rs("ext_color")
	x_vin = rs("vin")
	x_mileage = rs("mileage")
	x_engine = rs("engine")
	x_doors = rs("doors")
	x_transmission = rs("transmission")
	x_drivetrain = rs("drivetrain")
	x_lease_rental = rs("lease_rental")
	x_odometer = rs("odometer")
	x_records = rs("records")
	x_ac = rs("ac")
	x_pw_windows = rs("pw_windows")
	x_pw_locks = rs("pw_locks")
	x_pw_seats = rs("pw_seats")
	x_pw_steering = rs("pw_steering")
	x_cr_ct = rs("cr_ct")
	x_navig = rs("navig")
	x_sunroof = rs("sunroof")
	x_dvd = rs("dvd")
	x_satelit = rs("satelit")
	x_cd_cd_ch = rs("cd_cd_ch")
	x_am_fm = rs("am_fm")
	x_cass = rs("cass")
	x_leather = rs("leather")
	x_alloy = rs("alloy")
	x_spoiler = rs("spoiler")
	x_body = rs("body")
	x_tires = rs("tires")
	x_engine_rate = rs("engine rate")
	x_trans_rate = rs("trans rate")
	x_glass_rate = rs("glass rate")
	x_interior_rate = rs("interior rate")
	x_exhouse_rate = rs("exhouse rate")
	x_lienholders = rs("lienholders")
	x_title = rs("title")
	x_work = rs("work")
	x_new = rs("new")
	x_accidents = rs("accidents")
	x_dameges = rs("dameges")
	x_paint = rs("paint")
	x_salvage = rs("salvage")
	x_comments = rs("comments")
	x_stock = rs("stock")
	x_tradeindate = rs("tradeindate")
%>
<tr bgcolor="<%= bgcolor %>">
<td width="62"><font face="Verdana"><span style="font-size:8pt;"><% response.write x_tradeindate %></span></font></td>
<td width="57">
                            <p align="center"><font face="Verdana"><span style="font-size:8pt;"><%
Select Case x_year
    Case "2010" response.write "2010"
    Case "2009" response.write "2009"
    Case "2008" response.write "2008"
    Case "2007" response.write "2007"
    Case "2006" response.write "2006"
    Case "2005" response.write "2005"
    Case "2004" response.write "2004"
    Case "2003" response.write "2003"
    Case "2002" response.write "2002"
    Case "2001" response.write "2001"
    Case "2000" response.write "2000"
    Case "1999" response.write "1999"
    Case "1998" response.write "1998"
    Case "1997" response.write "1997"
    Case "1996" response.write "1996"
    Case "1995" response.write "1995"
    Case "1994" response.write "1994"
    Case "1993" response.write "1993"
    Case "1992" response.write "1992"
    Case "1991" response.write "1991"
    Case "1990" response.write "1990"
    Case "1989" response.write "1989"
    Case "1988" response.write "1988"
    Case "1887" response.write "1987"
    Case "1886" response.write "1986"
    Case "1885" response.write "1985"
End Select
%></span></font></td>
<td width="221"><font face="Verdana"><span style="font-size:8pt;">&nbsp;</span></font><a href="tradeinview.asp?key=<%=key%>"><font face="Verdana"><span style="font-size:8pt;"><% response.write x_make %>&nbsp;<% response.write x_model %></span></font></a></td>
<td width="100">
                            <p align="center"><font face="Verdana"><span style="font-size:8pt;"><% if isnumeric(x_mileage) then response.write formatnumber(x_mileage,0,-2,-2,-2) else response.write x_mileage end if %></span></font></td>
<td width="100">
                            <p align="center"><font face="Verdana"><span style="font-size:8pt;"><%
Select Case x_transmission
    Case "Automatic" response.write "Automatic"
    Case "Manual" response.write "Manual"
End Select
%></span></font></td>
<td width="112">
                            <p align="center"><font face="Verdana"><span style="font-size:8pt;"><% response.write x_stock %></span></font></td>
<td width="61">
                            <p align="center"><a href="<% key = rs("ID") : If not isnull(key) Then response.write "tradeinview.asp?key=" & Server.URLEncode(key) Else response.write "javascript:alert('Invalid Record! Key is null');" End If %>"><font face="Arial"><span style="font-size:8pt;"><b>View</b></span></font></a></td>
<td width="43">
                            <p align="center"><a href="<% key = rs("ID") : If not isnull(key) Then response.write "tradeinedit.asp?key=" & Server.URLEncode(key) Else response.write "javascript:alert('Invalid Record! Key is null');" End If %>"><font face="Arial"><span style="font-size:8pt;"><b>Edit</b></span></font></a></td>
<td width="32">
                            <p align="right"><input type="checkbox" name="key" value="<%= key %>"></td>
</tr>
<%
	End If
	rs.MoveNext
Loop
%>
</table>
<% If recActual > 0 Then %>
<p align="right"><input type="button" name="btndelete" value="DELETE SELECTED" onClick="this.form.action='tradeindelete.asp';this.form.submit();" style="font-family:Arial; font-size:12;"> 
                &nbsp;&nbsp;&nbsp;</p>
<% End If %>
</form>
<%
' Close recordSet and connection
rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing %><table border="0" cellspacing="0" cellpadding="10" width="118" align="center"><tr><td width="98" height="40" align="center">
                        <font face="Arial"><span style="font-size:8pt;"><%
If totalRecs > 0 Then
	rsEof = (totalRecs < (startRec + displayRecs))
	PrevStart = startRec - displayRecs
	If PrevStart < 1 Then PrevStart = 1
	NextStart = startRec + displayRecs
	If NextStart > totalRecs Then NextStart = startRec
	LastStart = ((totalRecs-1)\displayRecs)*displayRecs+1
	%>































































































































































































































































































































































































































































































































































































































































































































































































































































































































































































































































<meta name="generator" content="Namo WebEditor v5.0(Trial)"></span></font>

                        <p>
                        <p>
                        <p>
                        <p>
                        <p>
                        <p>
                        <p>
                        <p>
                        <p>
                        <p>
                        <p>
                        <p>
                        <p>
                        <p>
                        <p>
                        <p>	
<p>
<p>
                        <p>
                        <p>
                        <p>
                        <p><table border="0" cellspacing="0" cellpadding="0"><tr><td><font face="Arial"><span style="font-size:8pt;">Page</span></font>&nbsp;</td>
<!--first page button-->
	<% If CLng(startRec)=1 Then %>
	<td><img src="images/firstdisab.gif" alt="First" width="20" height="15" border="0"></td>
	<% Else %>
	<td><a href="tradeinlist.asp?start=1"><img src="images/first.gif" alt="First" width="20" height="15" border="0"></a></td>
	<% End If %>
<!--previous page button-->
	<% If CLng(PrevStart) = CLng(startRec) Then %>
	<td><img src="images/prevdisab.gif" alt="Previous" width="20" height="15" border="0"></td>
	<% Else %>
	<td><a href="tradeinlist.asp?start=<%=PrevStart%>"><img src="images/prev.gif" alt="Previous" width="20" height="15" border="0"></a></td>
	<% End If %>
<!--current page number-->
	<td><input type="text" name="pageno" value="<%=(startRec-1)\displayRecs+1%>" size="4" style="font-size: 9pt;"></td>
<!--next page button-->
	<% If CLng(NextStart) = CLng(startRec) Then %>
	<td><img src="images/nextdisab.gif" alt="Next" width="20" height="15" border="0"></td>
	<% Else %>
	<td><a href="tradeinlist.asp?start=<%=NextStart%>"><img src="images/next.gif" alt="Next" width="20" height="15" border="0"></a></td>
	<% End If %>
<!--last page button-->
	<% If CLng(LastStart) = CLng(startRec) Then %>
	<td><img src="images/lastdisab.gif" alt="Last" width="20" height="15" border="0"></td>
	<% Else %>
	<td><a href="tradeinlist.asp?start=<%=LastStart%>"><img src="images/last.gif" alt="Last" width="20" height="15" border="0"></a></td>
	<% End If %>
	</tr></table>	
                        
                        </span></font></span></font></td></tr></table>
                            <p align="center"><font face="Arial"><span style="font-size:8pt;"><% If CLng(startRec) > CLng(totalRecs) Then startRec = totalRecs
	stopRec = startRec + displayRecs - 1
	recCount = totalRecs - 1
	If rsEOF Then recCount = totalRecs
	If stopRec > recCount Then stopRec = recCount %>
	Records <%= startRec %> to <%= stopRec %> of <%= totalRecs %>
<% Else %>
	No records found<% End If %></span></font></td>
    </tr>
</table>

<!--#include file="footer.asp"-->