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
Response.Buffer = True
'multiple delete records
key = Request.Form("key")
arRecKey = Split(key&"", ",")
If UBound(arRecKey) = -1 Then Response.Redirect "findelist.asp"
For Each reckey In arRecKey
	'remove spaces
	reckey = trim(reckey)
	' build the SQL
	sqlKey = sqlKey & "("
	sqlKey = sqlKey & "[ID]=" & "" & reckey & "" & " AND "
	If Right(sqlKey, 5)=" AND " Then sqlKey = Left(sqlKey, Len(sqlKey)-5)
	sqlKey = sqlKey & ") OR "
Next
If Right(sqlKey, 4)=" OR " Then sqlKey = Left(sqlKey, Len(sqlKey)-4)
'get action
a=Request.Form("a")
If a="" or IsNull(a) Then
	a="I"	'display with input box
End If
' Open Connection to the database
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str
Select Case a
	Case "I": ' Display
		strsql = "SELECT * FROM [finde] WHERE " & sqlKey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn
		If rs.EOF Then
			Response.Clear
			Response.Redirect "findelist.asp"
		Else
			rs.MoveFirst
		End If
	Case "D": ' Delete
		strsql = "SELECT * FROM [finde] WHERE " & sqlKey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn, 1, 2
		Do While NOT rs.EOF
			rs.Delete
			rs.MoveNext
		Loop
		rs.Close
		Set rs = Nothing
		conn.Close
		Set conn = Nothing		
		Response.Clear
		Response.Redirect "findelist.asp"
End Select
%>
<!--#include file="header.asp"-->
<meta name="generator" content="Namo WebEditor v5.0(Trial)">
<body bgcolor="white" text="black" link="blue" vlink="purple" alink="red">
<table align="center" cellspacing="0" width="803" bgcolor="white" border="1" bordercolordark="white" bordercolorlight="black">
    <tr>
        <td>
<p>
                <font face="Arial"><span style="font-size:14pt;"><b><br> </b></span><span style="font-size:10pt;"><b>&nbsp;&nbsp;&nbsp;<i>Delete 
            Find Car Record</i></b></span></font><input type="hidden" name="a" value="D"><br>

<form action="findedelete.asp" method="post">
<input type="hidden" name="a" value="D">
<table border="0" cellspacing="0" cellpadding="0" bgcolor="#CCCCCC" width="750" align="center">
<tr bgcolor="#708090">
<td width="75" bgcolor="white"><font face="Verdana" color="black"><span style="font-size:9pt;"><b>Date</b></span></font></td>
<td width="90" bgcolor="white">
                            <p align="center"><font face="Verdana" color="black"><span style="font-size:9pt;"><b>Year</b></span></font></td>
<td width="170" bgcolor="white"><font face="Verdana" color="black"><span style="font-size:9pt;"><b>Make&nbsp;</b></span></font></td>
<td width="102" bgcolor="white"><font face="Verdana" color="black"><span style="font-size:9pt;"><b>Model&nbsp;</b></span></font></td>
<td width="176" bgcolor="white">
                            <p align="center"><font face="Verdana" color="black"><span style="font-size:9pt;"><b>Mileage</b></span></font></td>
<td width="160" bgcolor="white">
                            <p align="center"><font face="Verdana" color="black"><span style="font-size:9pt;"><b>Price</b></span></font></td>
</tr>
<%
recCount = 0
Do While NOT rs.EOF
	recCount = recCount + 1
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
<input type="hidden" name="key" value="<%= key %>">
<td width="75"><font face="Verdana"><span style="font-size:10pt;"><% response.write x_appdate %>
                            </span></font></td>
<td width="90">
                            <p align="center"><font face="Verdana"><span style="font-size:10pt;"><% response.write x_yearold %> 
                            -
<% response.write x_yearnew %></span></font></td>
<td width="170">
                            <font face="Verdana"><span style="font-size:10pt;"><%Select Case x_make
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
%></span></font></td>
<td width="102">
                            <font face="Verdana"><span style="font-size:10pt;"><% response.write x_model %></span></font></td>
<td width="176">
                            <p align="center"> 
                            <font face="Arial" size="2"><% if isnumeric(x_mileagelow) then response.write formatnumber(x_mileagelow,0,-2,-2,-2) else response.write x_mileagelow end if %> 
                            </font><font face="Verdana"><span style="font-size:10pt;">-
                            </span></font><font face="Arial" size="2"><% if isnumeric(x_mileagehi) then response.write formatnumber(x_mileagehi,0,-2,-2,-2) else response.write x_mileagehi end if %></font></td>
<td width="160">
                            <p align="center"><font face="Arial" size="2"><% if isnumeric(x_pricelow) then response.write formatcurrency(x_pricelow,0,-2,-2,-2) else response.write x_pricelow end if %></font><font face="Verdana"><span style="font-size:10pt;"> 
                            -
                            </span></font><font face="Arial" size="2"><% if isnumeric(x_pricehi) then response.write formatcurrency(x_pricehi,0,-2,-2,-2) else response.write x_pricehi end if %></font></td>
</tr>
<%
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing
%>
</table>
<p align="center">
<input type="submit" name="Action" value="CONFIRM DELETE">
</form>
<p>
            <font face="Arial" color="black"><span style="font-size:10pt;"><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</b></span></font><a href="findelist.asp"><font face="Arial" color="black"><span style="font-size:10pt;"><b><img src="images/leftsm.gif" align="texttop" width="16" height="16" border="0"></b></span></font></a><font face="Arial" color="black"><span style="font-size:10pt;"><b> 
                &nbsp;&nbsp;</b></span></font><a href="findelist.asp"><font face="Arial" color="black"><span style="font-size:10pt;"><b>Back to Find 
                a Car List</b></span></font></a><font face="Arial" color="black"><span style="font-size:10pt;"><b><br>&nbsp;</b></span></font>        </td>
    </tr>
</table>
<!--#include file="footer.asp"-->