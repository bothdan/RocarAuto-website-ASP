<html>
<!--#include file="header.asp"-->
<!--#include file="db.asp"-->
<!--#include file="aspmkrfn.asp"-->
<% If Session("rocar_status") <> "login" Then Response.Redirect "login.asp" %>
<% Session.Timeout = 300 %>
<%
Response.expires = 0
Response.expiresabsolute = Now() - 1
Response.addHeader "pragma", "no-cache"
Response.addHeader "cache-control", "private"
Response.CacheControl = "no-cache"
%>
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
		Session("finde_searchwhere") =searchwhere
    ElseIf UCase(cmd) = "RESETALL" Then
		'reset search criteria
		searchwhere = ""
		Session("finde_searchwhere") =searchwhere
	End If
	'reset start record counter (reset command)
	startRec = 1
	Session("finde_REC") = startRec
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
DefaultOrder = "appdate"
DefaultOrderType = "DESC"
'No Default Filter
DefaultFilter = ""
' Check for an Order parameter
OrderBy = ""
If Request.QueryString("order").Count > 0 Then
	OrderBy = Request.QueryString("order")
	' Check If an ASC/DESC toggle is required
	If Session("finde_OB") = OrderBy Then
		If Session("finde_OT") = "ASC" Then
			Session("finde_OT") = "DESC"
		Else
			Session("finde_OT") = "ASC"
		End if
	Else
		Session("finde_OT") = "ASC"
	End If
	Session("finde_OB") = OrderBy
	Session("finde_REC") = 1
Else
	OrderBy = Session("finde_OB")
	If OrderBy = "" Then
		OrderBy = DefaultOrder
		Session("finde_OB") = OrderBy
		Session("finde_OT") = DefaultOrderType
	End If
End If
' Open Connection to the database
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str
' Build SQL
strsql = "SELECT * FROM [finde]"
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
	strsql = strsql & " ORDER BY [" & OrderBy & "] " & Session("finde_OT")
End If	
'Response.Write strsql
Set rs = Server.CreateObject("ADODB.Recordset")
rs.cursorlocation = 3
rs.Open strsql, conn, 1, 2
totalRecs = rs.RecordCount
' Check for a START parameter
If Request.QueryString("start").Count > 0 Then
	startRec = Request.QueryString("start")
	Session("finde_REC") = startRec
ElseIf Request.QueryString("pageno").Count > 0 Then
	pageno = Request.QueryString("pageno")
	If IsNumeric(pageno) Then
		startRec = (pageno-1)*displayRecs+1
		If startRec <= 0 Then
			startRec = 1
		ElseIf startRec >= ((totalRecs-1)\displayRecs)*displayRecs+1 Then
			startRec = ((totalRecs-1)\displayRecs)*displayRecs+1
		End If
		Session("finde_REC") = startRec
	Else
		startRec = Session("finde_REC")
		If Not IsNumeric(startRec) Or startRec = "" Then
			'reset start record counter
			startRec = 1
			Session("finde_REC") = startRec
		End If
	End If
Else
	startRec = Session("finde_REC")
	If Not IsNumeric(startRec) Or startRec = "" Then
		'reSet start record counter
		startRec = 1
		Session("finde_REC") = startRec
	End If
End If
%>

<meta name="generator" content="Namo WebEditor v5.0(Trial)">
</head>


<body link="green" vlink="black" alink="red" bgcolor="white" text="black" leftmargin="0" marginwidth="0" topmargin="0" marginheight="0">
<table align="center" cellspacing="0" width="801" bordercolordark="white" bordercolorlight="black" bgcolor="white" cellpadding="0">
    <tr>
        <td height="3">
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
        <td height="9">

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
                <p align="center"><a href="makeofferlist.asp"><span style="font-size:10pt;"><font face="Arial" color="white"><b>Offers</b></font></span></a></p>
            </td>
            <td width="5" height="5" bgcolor="white">
                <p align="center"><span style="font-size:10pt;"><font face="Arial" color="white"><b>|</b></font></span></p>
            </td>
            <td width="120" height="5" bgcolor="#979EE0">
                <p align="center"><span style="font-size:10pt;"><a href="moreinfolist.asp"><b><font face="Arial" color="white">More 
                Info</font></b></a></span></p>
            </td>
            <td width="5" height="5" bgcolor="white">
                <p align="center"><span style="font-size:10pt;"><font face="Arial" color="white"><b>|</b></font></span></p>
            </td>
            <td width="100" height="5" bgcolor="#6351B7">
                <p align="center"><span style="font-size:10pt;"><font face="Arial" color="white"><b>Find 
                Car</b></font></span></p>
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
                <p align="center"></p>
            </td>
        </tr>
    </table>
</div>
</td>
    </tr>
    <tr>
        <td>
            <p align="left"><font face="Arial"><span style="font-size:11pt;"><br> 
            &nbsp;&nbsp;&nbsp;<b><i>Find Car Records:<br></i></b></span></font></p>
            <table align="center" cellspacing="0" width="801" bordercolordark="white" bordercolorlight="black" bgcolor="white" cellpadding="0">
                <tr>
                    <td><form method="post">
<table bgcolor="#CCCCCC" width="750" align="center" cellpadding="0" cellspacing="0">
<tr bgcolor="#708090">
<td nowrap width="78" height="20" background="images/sortbar.gif">
                            <p align="center"><a href='findelist.asp?order=<%= Server.URLEncode("appdate") %>'><font color="black" face="Arial"><span style="font-size:8pt;"><b>Date</b></span></font></a><font color="black" face="Webdings"><span style="font-size:8pt;"><b> 
                            <% If OrderBy = "appdate" Then %><% If Session("finde_OT") = "ASC" Then %>5<% ElseIf Session("finde_OT") = "DESC" Then %>6<% End If %><% End If %></b></span></font>
</td>
<td width="212" nowrap height="20" background="images/sortbar.gif">
<font color="black" face="Verdana"><span style="font-size:8pt;"><b>&nbsp;</b></span></font><a href="findelist.asp?order=<%= Server.URLEncode("make") %>"><font color="black" face="Arial"><span style="font-size:8pt;"><b>Make 
                            &amp; Model</b></span></font></a><font color="black" face="Verdana"><span style="font-size:8pt;"><b>&nbsp;<% If OrderBy = "make" Then %><% If Session("finde_OT") = "ASC" Then %></b></span></font><font color="black" face="Webdings"><span style="font-size:8pt;"><b>5<% ElseIf Session("finde_OT") = "DESC" Then %>6</b></span></font><a href="findelist.asp?order=<%= Server.URLEncode("make") %>"><font color="black" face="Verdana"><span style="font-size:8pt;"><b><% End If %><% End If %></b></span></font></a>
</td>
<td width="90" nowrap height="20" background="images/sortbar.gif">
                            <p align="center"><font color="black" face="Arial"><span style="font-size:8pt;"><b>Year</b></span></font>
</td>
<td width="115" nowrap height="20" background="images/sortbar.gif">
                            <p align="center"><font face="Arial" color="black"><span style="font-size:8pt;"><b>Mileage</b></span></font>
</td>
<td width="115" nowrap height="20" background="images/sortbar.gif">
                            <p align="center"><font color="black" face="Arial"><span style="font-size:8pt;"><b>Price</b></span></font>
</td>
<td width="101" nowrap height="20" background="images/sortbar.gif">
                            <p align="center"><a href="findelist.asp?order=<%= Server.URLEncode("time") %>"><font color="black" face="Arial"><span style="font-size:8pt;"><b>Search</b></span></font></a><font color="black" face="Verdana"><span style="font-size:8pt;"><b> 
                            &nbsp;</b></span></font><font color="black" face="Webdings"><span style="font-size:8pt;"><b><% If OrderBy = "time" Then %><% If Session("finde_OT") = "ASC" Then %>5<% ElseIf Session("finde_OT") = "DESC" Then %>6</b></span></font><a href="findelist.asp?order=<%= Server.URLEncode("time") %>"><font color="black" face="Webdings"><span style="font-size:8pt;"><b><% End If %></b></span></font><font color="black" face="Verdana"><span style="font-size:8pt;"><b><% End If %></b></span></font></a>
</td>
<td width="65" nowrap height="20" colspan="2" background="images/sortbar.gif"><font color="black"><span style="font-size:8pt;"><b>&nbsp;</b></span></font></td>
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
	x_first_name = rs("first name")
	x_last_name = rs("last name")
	x_home_phone = rs("home phone")
	x_email = rs("email")
	x_type = rs("type")
	x_yearold = rs("yearold")
	x_yearnew = rs("yearnew")
	x_make = rs("make")
	x_model = rs("model")
	x_bodystyle = rs("bodystyle")
	x_transmission = rs("transmission")
	x_mileagelow = rs("mileagelow")
	x_mileagehi = rs("mileagehi")
	x_pricelow = rs("pricelow")
	x_pricehi = rs("pricehi")
	x_comments = rs("comments")
	x_time = rs("time")
	x_appdate = rs("appdate")
%>
<tr bgcolor="<%= bgcolor %>">
<td width="78" height="13" nowrap>
                            
                                        <p align="right"><font face="Verdana"><span style="font-size:8pt;"><% response.write x_appdate %>&nbsp;</span></font></td>
<td width="212" height="13" nowrap><font face="Verdana"><span style="font-size:8pt;">&nbsp;</span></font><a href="findeview.asp?key=<%= key %>"><font face="Verdana"><span style="font-size:8pt;"><%Select Case x_make
    Case "Acura" response.write "Acura"
    Case "Alfa Romeo" response.write "Alfa Romeo"
    Case "Am General" response.write "Am General"
    Case "Aston Martin" response.write "Aston Martin"
    Case "Audi" response.write "Audi"
    Case "BMW" response.write "BMW"
    Case "Bentley" response.write "Bentley"
    Case "Buick" response.write "Buick"
    Case "Cadillac" response.write "Cadillac"
    Case "Chevrolet" response.write "Chevrolet"
    Case "Chrysler" response.write "Chrysler"
    Case "Dacia" response.write "Dacia"
    Case "Daewoo" response.write "Daewoo"
    Case "Daihatsu" response.write "Daihatsu"
    Case "Dodge" response.write "Dodge"
    Case "Eagle" response.write "Eagle"
    Case "Ferrari" response.write "Ferrari"
    Case "Ford" response.write "Ford"
    Case "GMC" response.write "GMC"
    Case "Geo" response.write "Geo"
    Case "Honda" response.write "Honda"
    Case "Hummer" response.write "Hummer"
    Case "Hyundai" response.write "Hyundai"
    Case "Infiniti" response.write "Infiniti"
    Case "International" response.write "International"
    Case "Isuzu" response.write "Isuzu"
    Case "Jaguar" response.write "Jaguar"
    Case "Jeep" response.write "Jeep"
    Case "Kia" response.write "Kia"
    Case "Lamborghini" response.write "Lamborghini"
    Case "Land Rover" response.write "Land Rover"
    Case "Lexus" response.write "Lexus"
    Case "Lincoln" response.write "Lincoln"
    Case "Lotus" response.write "Lotus"
    Case "Maserati" response.write "Maserati"
    Case "Maybach" response.write "Maybach"
    Case "Mazda" response.write "Mazda"
    Case "Mercedes-Benz" response.write "Mercedes-Benz"
    Case "Mercury" response.write "Mercury"
    Case "Mini" response.write "Mini"
    Case "Mitsubishi" response.write "Mitsubishi"
    Case "Morgan" response.write "Morgan"
    Case "Nissan" response.write "Nissan"
    Case "Oldsmobile" response.write "Oldsmobile"
    Case "Panoz" response.write "Panoz"
    Case "Peugeot" response.write "Peugeot"
    Case "Plymouth" response.write "Plymouth"
    Case "Pontiac" response.write "Pontiac"
    Case "Porsche" response.write "Porche"
    Case "Rolls-Royce" response.write "Rolls-Royce"
    Case "Saab" response.write "Saab"
    Case "Saleen" response.write "Saleen"
    Case "Saturn" response.write "Saturn"
    Case "Scion" response.write "Scion"
    Case "Smart" response.write "Smart"
    Case "Sterling" response.write "Sterling"
    Case "Subaru" response.write "Subaru"
    Case "Suzuki" response.write "Suzuki"
    Case "Tesla" response.write "Tesla"
    Case "Toyota" response.write "Toyota"
    Case "Volkswagen" response.write "Volkswagen"
    Case "Volvo" response.write "Volvo"
    Case "Yugo" response.write "Yugo"
End Select
%></span></font><span style="font-size:8pt;"><font face="Verdana"> </font></span><font face="Verdana"><span style="font-size:8pt;"><% response.write x_model %></span></font></a><font face="Verdana"><span style="font-size:8pt;"> </span></font></td>
<td width="90" height="13" nowrap>
                            <p align="center"><font face="Verdana"><span style="font-size:8pt;"><% response.write x_yearold %>&nbsp;- 
                            <% response.write x_yearnew %></span></font></td>
<td width="115" height="13" nowrap>
                            <p align="center"> 
                            <font face="Verdana"><span style="font-size:8pt;"><% if isnumeric(x_mileagelow) then response.write formatnumber(x_mileagelow,0,-2,-2,-2) else response.write x_mileagelow end if %> 
                            -
                            <% if isnumeric(x_mileagehi) then response.write formatnumber(x_mileagehi,0,-2,-2,-2) else response.write x_mileagehi end if %></span></font></td>
<td width="115" height="13" nowrap>
                            <p align="center"><font face="Verdana"><span style="font-size:8pt;"><% if isnumeric(x_pricelow) then response.write formatcurrency(x_pricelow,0,-2,-2,-2) else response.write x_pricelow end if %> 
                            -
                            <% if isnumeric(x_pricehi) then response.write formatcurrency(x_pricehi,0,-2,-2,-2) else response.write x_pricehi end if %></span></font></td>
<td width="101" height="13" nowrap>
                            <p align="center"><font face="Verdana"><span style="font-size:8pt;"><% response.write x_time %></span></font></td>
<td width="38" height="13" nowrap>
                            <p align="center"><a href="<% key = rs("ID") : If not isnull(key) Then response.write "findeedit.asp?key=" & Server.URLEncode(key) Else response.write "javascript:alert('Invalid Record! Key is null');" End If %>"><font face="Verdana"><span style="font-size:8pt;"><b>Edit</b></span></font></a></td>
<td width="27" height="13" nowrap>
                            <p align="right"><input type="checkbox" name="key" value="<%= key %>"></td>
</tr>
<%
	End If
	rs.MoveNext
Loop
%>
</table>
<% If recActual > 0 Then %>
<p align="right"><input type="button" name="btndelete" value="DELETE SELECTED" onClick="this.form.action='findedelete.asp';this.form.submit();" style="font-family:Arial; font-size:12;"> 
                            </p>
<% End If %>
</form>
<p><%
' Close recordSet and connection
rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing %></p>
                    </td>
                </tr>
            </table>
                        <p align="center"><font face="Arial"><span style="font-size:12pt;"><%
' Display page numbers
If totalRecs > 0 Then
	rsEof = (totalRecs < (startRec + displayRecs))
	' Find out if there should be Backward or Forward Buttons on the table.
	If 	startRec = 1 Then
		isPrev = False
	Else
		isPrev = True
		PrevStart = startRec - displayRecs
		If PrevStart < 1 Then PrevStart = 1 %></span></font><a href="makeofferlist.asp?start=<%=x%>"><span style="font-size:12pt;"><font face="Arial" color="black"><%
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
				If CLng(startRec) = CLng(x) Then %></font></span></a><font face="Arial"><span style="font-size:12pt;">
	</span></font><a href="findelist.asp?start=<%=x%>"><span style="font-size:12pt;"><font face="Arial" color="red"><%=y%></font></span></a><font face="Arial"><span style="font-size:12pt;">
				</span></font><a href="findelist.asp?start=<%=x%>"><span style="font-size:12pt;"><font face="Arial" color="black"><%	Else %></font></span></a><span style="font-size:12pt;"><font face="Arial" color="black">
	</font></span><a href="findelist.asp?start=<%=x%>"><font face="Arial" color="black"><span style="font-size:12pt;"><%=y%></span></font></A><span style="font-size:12pt;"><font face="Arial" color="black">
				</font></span><a href="findelist.asp?start=<%=x%>"><span style="font-size:12pt;"><font face="Arial" color="black"><%	End If
				x = x + displayRecs
				y = y + 1
			ElseIf x >= (dx1-displayRecs*recRange) AND x <= (dx2+displayRecs*recRange) Then
				If x+recRange*displayRecs < totalRecs Then %></font></span></a><span style="font-size:12pt;"><font face="Arial" color="black">
	</font></span><a href="findelist.asp?start=<%=x%>"><font face="Arial" color="black"><span style="font-size:12pt;"><%=y%></span></font></A><span style="font-size:12pt;"><font face="Arial" color="black">-</font></span><a href="findelist.asp?start=<%=x%>"><font face="Arial" color="black"><span style="font-size:12pt;"><%=y+recRange-1%></span></font></A><span style="font-size:12pt;"><font face="Arial" color="black">
				</font></span><a href="findelist.asp?start=<%=x%>"><span style="font-size:12pt;"><font face="Arial" color="black"><% Else
					ny=(totalRecs-1)\displayRecs+1
						If ny = y Then %></font></span></a><span style="font-size:12pt;"><font face="Arial" color="black">
	</font></span><a href="findelist.asp?start=<%=x%>"><font face="Arial" color="black"><span style="font-size:12pt;"><%=y%></span></font></A><span style="font-size:12pt;"><font face="Arial" color="black">
						</font></span><a href="findelist.asp?start=<%=x%>"><span style="font-size:12pt;"><font face="Arial" color="black"><% Else %></font></span></a><span style="font-size:12pt;"><font face="Arial" color="black">
	</font></span><a href="findelist.asp?start=<%=x%>"><font face="Arial" color="black"><span style="font-size:12pt;"><%=y%></span></font></A><span style="font-size:12pt;"><font face="Arial" color="black">-</font></span><a href="findelist.asp?start=<%=x%>"><font face="Arial" color="black"><span style="font-size:12pt;"><%=ny%></span></font></A><span style="font-size:12pt;"><font face="Arial" color="black">
						</font></span><a href="findelist.asp?start=<%=x%>"><span style="font-size:12pt;"><font face="Arial" color="black"><%	End If
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
		isMore = True %></font></span></a><font face="Arial"><span style="font-size:12pt;">
	<% Else
		isMore = False
	End If %>
		
	
<% End If %></span></font></p>
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
<!--#include file="footer.asp"--></p>
