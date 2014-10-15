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
		Session("makeoffer_searchwhere") =searchwhere
    ElseIf UCase(cmd) = "RESETALL" Then
		'reset search criteria
		searchwhere = ""
		Session("makeoffer_searchwhere") =searchwhere
	End If
	'reset start record counter (reset command)
	startRec = 1
	Session("makeoffer_REC") = startRec
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
DefaultOrder = "offerdate"
DefaultOrderType = "DESC"
'No Default Filter
DefaultFilter = ""
' Check for an Order parameter
OrderBy = ""
If Request.QueryString("order").Count > 0 Then
	OrderBy = Request.QueryString("order")
	' Check If an ASC/DESC toggle is required
	If Session("makeoffer_OB") = OrderBy Then
		If Session("makeoffer_OT") = "ASC" Then
			Session("makeoffer_OT") = "DESC"
		Else
			Session("makeoffer_OT") = "ASC"
		End if
	Else
		Session("makeoffer_OT") = "ASC"
	End If
	Session("makeoffer_OB") = OrderBy
	Session("makeoffer_REC") = 1
Else
	OrderBy = Session("makeoffer_OB")
	If OrderBy = "" Then
		OrderBy = DefaultOrder
		Session("makeoffer_OB") = OrderBy
		Session("makeoffer_OT") = DefaultOrderType
	End If
End If
' Open Connection to the database
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str
' Build SQL
strsql = "SELECT * FROM [makeoffer]"
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
	strsql = strsql & " ORDER BY [" & OrderBy & "] " & Session("makeoffer_OT")
End If	
'Response.Write strsql
Set rs = Server.CreateObject("ADODB.Recordset")
rs.cursorlocation = 3
rs.Open strsql, conn, 1, 2
totalRecs = rs.RecordCount
' Check for a START parameter
If Request.QueryString("start").Count > 0 Then
	startRec = Request.QueryString("start")
	Session("makeoffer_REC") = startRec
ElseIf Request.QueryString("pageno").Count > 0 Then
	pageno = Request.QueryString("pageno")
	If IsNumeric(pageno) Then
		startRec = (pageno-1)*displayRecs+1
		If startRec <= 0 Then
			startRec = 1
		ElseIf startRec >= ((totalRecs-1)\displayRecs)*displayRecs+1 Then
			startRec = ((totalRecs-1)\displayRecs)*displayRecs+1
		End If
		Session("makeoffer_REC") = startRec
	Else
		startRec = Session("makeoffer_REC")
		If Not IsNumeric(startRec) Or startRec = "" Then
			'reset start record counter
			startRec = 1
			Session("makeoffer_REC") = startRec
		End If
	End If
Else
	startRec = Session("makeoffer_REC")
	If Not IsNumeric(startRec) Or startRec = "" Then
		'reSet start record counter
		startRec = 1
		Session("makeoffer_REC") = startRec
	End If
End If
%>
<!--#include file="header.asp"-->
<meta name="Web Design" content="Dan Both">
<meta name="generator" content="Namo WebEditor v5.0(Trial)">
<body bgcolor="white" text="black" link="green" vlink="black" alink="red">
<table align="center" cellspacing="0" bordercolordark="white" bordercolorlight="black" bgcolor="white" width="801" cellpadding="0">
    <tr>
        <td width="795" height="10">
            <div align="left">
                <table cellpadding="0" cellspacing="0" bgcolor="white" width="360">
                    <tr>
                        <td width="20" height="14">
                            <p></p>
                        </td>
                        <td width="5" height="14">
                <p align="center"><font face="Arial"><span style="font-size:12pt;">|</span></font></p>
                        </td>
                        <td width="100" bgcolor="white" height="14">
                            <p align="center"><span style="font-size:10pt;"><a href="adminlist.asp"><b><font face="Arial" color="navy">Inventory</font></b></a></span></p>
                        </td>
                        <td width="5" bgcolor="white" height="14">
                <p align="center"><font face="Arial"><span style="font-size:12pt;">|</span></font></p>
                        </td>
                        <td width="100" height="14">
                            <p align="center"><b><span style="font-size:10pt;"><font face="Arial" color="green">Customers</font></span></b></p>
                        </td>
                        <td width="5" height="14">
                <p align="center"><span style="font-size:12pt;"><font face="Arial">|</font></span></p>
                        </td>
                        <td width="120" height="14">
                            <p align="center"><a href="newslist.asp"><b><span style="font-size:10pt;"><font face="Arial" color="navy">News 
                            And Events</font></span></b></a></p>
                        </td>
                        <td width="5" height="14">
                <p align="center"><span style="font-size:12pt;"><font face="Arial">|</font></span></p>
                        </td>
                    </tr>
                </table>
</div>
            
</td>
    </tr>
    <tr>
        <td width="795" height="11">

<div align="left">
    <table cellpadding="0" cellspacing="0" width="695" bgcolor="white">
        <tr>
            <td width="20" height="5">
                <p><font face="Arial"><span style="font-size:10pt;"><b>&nbsp;</b></span></font></p>
            </td>
            <td width="5" height="5" bgcolor="white">
                <p align="center"><font face="Arial"><span style="font-size:10pt;"><b>&nbsp;</b></span></font></p>
            </td>
            <td width="100" height="5" bgcolor="#979EE0">
                <p align="center"><a href="contactlist.asp"><span style="font-size:10pt;"><b><font face="Arial" color="white">Contact</font></b></span></a></p>
            </td>
            <td width="5" height="5" bgcolor="white">
                <p align="center"><font face="Arial"><span style="font-size:10pt;"><b>&nbsp;</b></span></font></p>
            </td>
            <td width="100" height="5" bgcolor="#6351B7">
                <p align="center"><span style="font-size:10pt;"><font face="Arial" color="white"><b>Offers</b></font></span></p>
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
            <td width="100" height="5" bgcolor="#979EE0">
                <p align="center"><a href="tradeinlist.asp"><span style="font-size:10pt;"><font face="Arial" color="white"><b>Trade-In</b></font></span></a></p>
            </td>
            <td width="5" height="5">
                <p align="center"><font face="Arial"><span style="font-size:10pt;"><b>&nbsp;</b></span></font></p>
            </td>
        </tr>
    </table>
</div>
</td>
    </tr>
    <tr>
        <td width="795">
                        <p><font face="Arial"><span style="font-size:11pt;"><br>&nbsp;&nbsp;&nbsp;<b><i>Offers 
                        Records:<br></i></b></span></font></p>
            <table align="center" width="801" bgcolor="white" cellspacing="0" bordercolordark="white" bordercolorlight="black" cellpadding="0">
                <tr>
                    <td width="765"><form method="post">
<table border="0" cellspacing="0" cellpadding="0" bgcolor="#CCCCCC" width="750" align="center">
<tr bgcolor="#708090">
<td width="70" bgcolor="white" background="images/sortbar.gif" height="20">
                                        <p align="center"><a href='makeofferlist.asp?order=<%= Server.URLEncode("offerdate") %>'><font face="Arial" color="black"><span style="font-size:8pt;"><b>Date</b></span></font></a><font face="Arial" color="black"><b><span style="font-size:8pt;">&nbsp;</span></b></font><font color="black"><b><span style="font-size:8pt;"><% If OrderBy = "offerdate" Then %></span></b></font><font face="Webdings" color="black"><b><span style="font-size:8pt;"><% If Session("makeoffer_OT") = "ASC" Then %>5<% ElseIf Session("makeoffer_OT") = "DESC" Then %>6<% End If %></span></b></font><font color="black"><b><span style="font-size:8pt;"><% End If %></span></b></font>
</td>
<td width="58" bgcolor="white" background="images/sortbar.gif" height="20">
                <p align="center"><font face="Arial" color="black"><span style="font-size:8pt;"><b>Year</b></span></font></p>

</td>
<td width="145" bgcolor="white" background="images/sortbar.gif" height="20">
                <p><font face="Arial" color="black"><span style="font-size:8pt;"><b>Make 
                                        &amp; 
                Model</b></span></font></p>

</td>
<td width="165" bgcolor="white" background="images/sortbar.gif" height="20">
<font face="Arial" color="black"><span style="font-size:8pt;"><b>First &amp; 
                                        Last Name</b></span></font>
</td>
<td width="95" bgcolor="white" background="images/sortbar.gif" height="20">
                                        <p align="center"><font face="Arial" color="black"><span style="font-size:8pt;"><b>Phone</b></span></font>
</td>
<td width="67" bgcolor="white" background="images/sortbar.gif" height="20">
                                        <p align="right"><font face="Arial" color="black"><span style="font-size:8pt;"><b>Offer 
                                        &nbsp;</b></span></font>
</td>
<td width="60" bgcolor="white" background="images/sortbar.gif" height="20">
                <p align="right"><font face="Arial" color="black"><span style="font-size:8pt;"><b>Price 
                                        &nbsp;</b></span></font></p>

</td>
<td width="41" bgcolor="white" background="images/sortbar.gif" height="20">
                                        <p align="center"><font face="Arial" color="black"><span style="font-size:8pt;"><b>&nbsp;Stock</b></span></font>
</td>
<td width="25" bgcolor="white" background="images/sortbar.gif" height="20"><font face="Arial" color="black"><span style="font-size:8pt;"><b>&nbsp;</b></span></font></td>
<td width="24" bgcolor="white" background="images/sortbar.gif" height="20"><font face="Arial" color="black"><span style="font-size:8pt;"><b>&nbsp;</b></span></font></td>
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
	x_phone = rs("phone")
	x_email = rs("email")
	x_offer = rs("offer")
	x_stock = rs("stock")
	x_year = rs("year")
	x_make = rs("make")
	x_model = rs("model")
	x_price = rs("price")
	x_offerdate= rs("offerdate")
%>
<tr bgcolor="<%= bgcolor %>">
<td width="70" height="19">
                                        <p align="right"><font face="Verdana"><span style="font-size:8pt;"><% response.write x_offerdate %></span></font></td>
<td width="58" height="19">
                                        <p align="center"><font face="Verdana"><span style="font-size:8pt;"><% response.write x_year %></span></font></td>
<td width="145" height="19"><a href="makeofferview.asp?key=<%=key%>"><span style="font-size:8pt;"><font face="Verdana"><% response.write x_make %> &nbsp;<% response.write x_model %></font></span></a></td>
<td width="165" height="19"><font face="Verdana"><span style="font-size:8pt;"><% response.write x_first %>&nbsp;<% response.write x_last %>&nbsp;</span></font></td>
<td width="95" height="19">
                                        <p align="center"><font face="Verdana"><span style="font-size:8pt;"><% response.write x_phone %></span></font></td>
<td width="67" height="19">                        <p align="right"><font face="Verdana" color="#336600"><span style="font-size:8pt;"><% if isnumeric(x_offer) then response.write formatcurrency(x_offer,0,-2,-2,-2) else response.write x_offer end if %>&nbsp;</span></font></p>
</td>
<td width="60" height="19">                        <p align="right"><font face="Verdana" color="red"><span style="font-size:8pt;"><% if isnumeric(x_price) then response.write formatcurrency(x_price,0,-2,-2,-2) else response.write x_price end if %></span></font></p>
</td>
<td width="41" height="19">
                                        <p align="center"><font face="Verdana"><span style="font-size:8pt;">&nbsp;<% response.write x_stock %></span></font></td>
<td width="25" height="19"><a href="<% key = rs("ID") : If not isnull(key) Then response.write "makeofferedit.asp?key=" & Server.URLEncode(key) Else response.write "javascript:alert('Invalid Record! Key is null');" End If %>"><font face="Arial"><span style="font-size:8pt;"><b>Edit</b></span></font></a></td>
<td width="24" height="19"><input type="checkbox" name="key" value="<%= key %>"></td>
</tr>
<%
	End If
	rs.MoveNext
Loop
%>
</table>
<% If recActual > 0 Then %>
<p align="right"><input type="button" name="btndelete" value="DELETE SELECTED" onClick="this.form.action='makeofferdelete.asp';this.form.submit();" style="font-family:Arial; font-size:12;"><% End If %></p>
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
            
                        <p align="center"><font face="Arial"><span style="font-size:12pt;"><b><%
' Display page numbers
If totalRecs > 0 Then
	rsEof = (totalRecs < (startRec + displayRecs))
	' Find out if there should be Backward or Forward Buttons on the table.
	If 	startRec = 1 Then
		isPrev = False
	Else
		isPrev = True
		PrevStart = startRec - displayRecs
		If PrevStart < 1 Then PrevStart = 1 %></b></span></font><a href="makeofferlist.asp?start=<%=x%>"><span style="font-size:12pt;"><font face="Arial" color="black"><b><%
	End If
	If (isPrev OR (NOT rsEof)) Then
		x = 1
		y = 1
		dx1 = ((startRec-1)\(displayRecs*recRange))*displayRecs*recRange+1
		dy1 = ((startRec-1)\(displayRecs*recRange))*recRange+1
		If (dx1+displayRecs*recRange-1) > totalRecs Then
			dx2 = (totalRecs\displayRecs)*displayRecs+1
			dy2 = (totalRecs\displayRecs)+1
		Else
			dx2 = dx1+displayRecs*recRange-1
			dy2 = dy1+recRange-1
		End If
		While x <= totalRecs
			If x >= dx1 AND x <= dx2 Then
				If CLng(startRec) = CLng(x) Then %></b></font></span></a><font face="Arial"><span style="font-size:12pt;"><b>
	</b></span></font><a href="makeofferlist.asp?start=<%=x%>"><span style="font-size:12pt;"><font face="Arial" color="red"><b><%=y%></b></font></span></a><font face="Arial"><span style="font-size:12pt;"><b>
				</b></span></font><a href="makeofferlist.asp?start=<%=x%>"><span style="font-size:12pt;"><font face="Arial" color="black"><b><%	Else %></b></font></span></a><font face="Arial" color="black"><span style="font-size:12pt;"><b>
	</b></span></font><a href="makeofferlist.asp?start=<%=x%>"><font face="Arial" color="black"><span style="font-size:12pt;"><b><%=y%></b></span></font></A><font face="Arial" color="black"><span style="font-size:12pt;"><b>
				</b></span></font><a href="makeofferlist.asp?start=<%=x%>"><span style="font-size:12pt;"><font face="Arial" color="black"><b><%	End If
				x = x + displayRecs
				y = y + 1
			ElseIf x >= (dx1-displayRecs*recRange) AND x <= (dx2+displayRecs*recRange) Then
				If x+recRange*displayRecs < totalRecs Then %></b></font></span></a><font face="Arial" color="black"><span style="font-size:12pt;"><b>
	</b></span></font><a href="makeofferlist.asp?start=<%=x%>"><font face="Arial" color="black"><span style="font-size:12pt;"><b><%=y%></b></span></font></A><font face="Arial" color="black"><span style="font-size:12pt;"><b>-</b></span></font><a href="makeofferlist.asp?start=<%=x%>"><font face="Arial" color="black"><span style="font-size:12pt;"><b><%=y+recRange-1%></b></span></font></A><font face="Arial" color="black"><span style="font-size:12pt;"><b>
				</b></span></font><a href="makeofferlist.asp?start=<%=x%>"><span style="font-size:12pt;"><font face="Arial" color="black"><b><% Else
					ny=(totalRecs-1)\displayRecs+1
						If ny = y Then %></b></font></span></a><font face="Arial" color="black"><span style="font-size:12pt;"><b>
	</b></span></font><a href="makeofferlist.asp?start=<%=x%>"><font face="Arial" color="black"><span style="font-size:12pt;"><b><%=y%></b></span></font></A><font face="Arial" color="black"><span style="font-size:12pt;"><b>
						</b></span></font><a href="makeofferlist.asp?start=<%=x%>"><span style="font-size:12pt;"><font face="Arial" color="black"><b><% Else %></b></font></span></a><font face="Arial" color="black"><span style="font-size:12pt;"><b>
	</b></span></font><a href="makeofferlist.asp?start=<%=x%>"><font face="Arial" color="black"><span style="font-size:12pt;"><b><%=y%></b></span></font></A><font face="Arial" color="black"><span style="font-size:12pt;"><b>-</b></span></font><a href="makeofferlist.asp?start=<%=x%>"><font face="Arial" color="black"><span style="font-size:12pt;"><b><%=ny%></b></span></font></A><font face="Arial" color="black"><span style="font-size:12pt;"><b>
						</b></span></font><a href="makeofferlist.asp?start=<%=x%>"><span style="font-size:12pt;"><font face="Arial" color="black"><b><%	End If
				End If
				x=x+recRange*displayRecs
				y=y+recRange
			Else
				x=x+recRange*displayRecs
				y=y+recRange
			End If
		Wend
	End If
	' Next link
	If NOT rsEof Then
		NextStart = startRec + displayRecs
		isMore = True %></b></font></span></a><font face="Arial"><span style="font-size:12pt;"><b>
	<% Else
		isMore = False
	End If %>
		
	
<% End If %></b></span></font></p>
            <p align="center"><font face="Verdana"><span style="font-size:8pt;"><%
If totalRecs > 0 Then
	rsEof = (totalRecs < (startRec + displayRecs))
	PrevStart = startRec - displayRecs
	If PrevStart < 1 Then PrevStart = 1
	NextStart = startRec + displayRecs
	If NextStart > totalRecs Then NextStart = startRec
	LastStart = ((totalRecs-1)\displayRecs)*displayRecs+1
	%><% If CLng(startRec) > CLng(totalRecs) Then startRec = totalRecs
	stopRec = startRec + displayRecs - 1
	recCount = totalRecs - 1
	If rsEOF Then recCount = totalRecs
	If stopRec > recCount Then stopRec = recCount %>
	Records <%= startRec %> to <%= stopRec %> of <%= totalRecs %>
<% Else %>
	No records found<% End If %></span></font></p>
</td>
    </tr>
</table>

<!--#include file="footer.asp"-->
