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
If UBound(arRecKey) = -1 Then Response.Redirect "carlist.asp"
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
		strsql = "SELECT * FROM [car] WHERE " & sqlKey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn
		If rs.EOF Then
			Response.Clear
			Response.Redirect "carlist.asp"
		Else
			rs.MoveFirst
		End If
	Case "D": ' Delete
		strsql = "SELECT * FROM [car] WHERE " & sqlKey
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
		Response.Redirect "adminlist.asp"
End Select
%>
<!--#include file="header.asp"-->
<meta name="generator" content="Namo WebEditor v5.0(Trial)">
<table width="802" bgcolor="white" align="center" cellpadding="0" cellspacing="0">
    <tr>
        <td height="158">
<form action="cardelete.asp" method="post">
<p>
<input type="hidden" name="a" value="D"><br> <font face="Arial" color="black"><span style="font-size:14pt;"><b>&nbsp;&nbsp;&nbsp;Delete 
                Inventory Recods:<br>&nbsp;</b></span></font>
<table border="0" cellspacing="0" cellpadding="0" bgcolor="#CCCCCC" width="750" align="center">
<tr bgcolor="#708090">
<td width="30" bgcolor="white">
                            <p align="center"><font color="black" face="Verdana"><b><span style="font-size:9pt;">ID</span></b></font></td>
<td width="62" bgcolor="white">
                            <p align="center"><font color="black" face="Verdana"><b><span style="font-size:9pt;">Stock</span></b></font></td>
<td width="62" bgcolor="white">
                            <p align="center"><font color="black" face="Verdana"><b><span style="font-size:9pt;">Year</span></b></font></td>
<td width="227" bgcolor="white"><font color="black" face="Verdana"><b><span style="font-size:9pt;">Make&nbsp;&amp; 
                            Model</span></b></font></td>
<td width="79" bgcolor="white">
                            <p align="right"><font color="black" face="Verdana"><b><span style="font-size:9pt;">Miles&nbsp;</span></b></font></td>
<td width="79" bgcolor="white">
                            <p align="right"><font color="black" face="Verdana"><b><span style="font-size:9pt;">Price&nbsp;</span></b></font></td>
<td width="82" bgcolor="white">
                            <p align="center"><font color="black" face="Verdana"><b><span style="font-size:9pt;">Status</span></b></font></td>
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
	x_year = rs("year")
	x_make = rs("make")
	x_model = rs("model")
	x_type = rs("type")
	x_miles = rs("miles")
	x_price = rs("price")
	x_doors = rs("doors")
	x_engine = rs("engine")
	x_transmission = rs("transmission")
	x_drivetrain = rs("drivetrain")
	x_ext_color = rs("ext_color")
	x_int_color = rs("int_color")
	x_stock = rs("stock")
	x_vin = rs("vin")
	x_city_mpg = rs("city_mpg")
	x_hwy_mpg = rs("hwy_mpg")
	x_carfax = rs("carfax")
	x_special = rs("special")
	x_status = rs("status")
	x_photo_1 = rs("photo 1")
	x_photo_2 = rs("photo 2")
	x_photo_3 = rs("photo 3")
	x_photo_4 = rs("photo 4")
	x_photo_5 = rs("photo 5")
	x_photo_6 = rs("photo 6")
	x_photo_7 = rs("photo 7")
	x_photo_8 = rs("photo 8")
	x_photo_9 = rs("photo 9")
	x_photo_10 = rs("photo 10")
	x_photo_11 = rs("photo 11")
	x_photo_12 = rs("photo 12")
	x_photo_13 = rs("photo 13")
	x_photo_14 = rs("photo 14")
	x_photo_15 = rs("photo 15")
	x_photo_16 = rs("photo 16")
	x_photo_17 = rs("photo 17")
	x_photo_18 = rs("photo 18")
	x_photo_19 = rs("photo 19")
	x_photo_20 = rs("photo 20")
	x_photo_21 = rs("photo 21")
	x_photo_22 = rs("photo 22")
	x_photo_23 = rs("photo 23")
	x_photo_24 = rs("photo 24")
	x_photo_25 = rs("photo 25")
	x_photo_26 = rs("photo 26")
	x_photo_27 = rs("photo 27")
	x_photo_28 = rs("photo 28")
	x_photo_29 = rs("photo 29")
	x_photo_30 = rs("photo 30")
%>
<tr bgcolor="<%= bgcolor %>">
<input type="hidden" name="key" value="<%= key %>">
<td width="30">
                            <p align="center"><font face="Arial" size="2"><%= x_ID %></font></td>
<td width="62">
                            <p align="center"><font face="Arial" size="2"><% response.write x_stock %></font></td>
<td width="62">
                            <p align="center"><font face="Arial" size="2"><%
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
%></font></td>
<td width="227">
<font face="Arial" size="2"><%
Select Case x_make
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
%>
&nbsp;
<% response.write x_model %>&nbsp;</font></td>
<td width="79">
                            <p align="right"><font face="Arial" size="2"><% if isnumeric(x_miles) then response.write formatnumber(x_miles,0,-2,-2,-2) else response.write x_miles end if %>&nbsp;
</font></td>
<td width="79">
                            <p align="right"><font face="Arial" size="2"><% if isnumeric(x_price) then response.write formatcurrency(x_price,0,-2,-2,-2) else response.write x_price end if %>&nbsp;
</font></td>
<td width="82">
                            <p align="center"><font face="Arial" size="2"><%
Select Case x_status
    Case "For Sale" response.write "For Sale"
    Case "Sold" response.write "Sold"
End Select
%></font></td>
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
<p>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="submit" name="Action" value="CONFIRM DELETE">
</form>
<font face="Arial" size="2" color="black"><b>&nbsp;&nbsp;&nbsp;&nbsp;</b></font><a href="adminlist.asp"><font face="Arial" size="2" color="black"><b><img src="images/back.gif" align="texttop" width="16" height="16" border="0"></b></font></a><font face="Arial" size="2" color="black"><b> 
            &nbsp;</b></font><a href="adminlist.asp"><font face="Arial" size="2" color="black"><b>Back to List</b></font></a><font face="Arial" size="2" color="black"><b><br>&nbsp;</b></font>
        </td>
    </tr>
</table>
<!--#include file="footer.asp"-->