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
key = Request.Querystring("key")
If key = "" OR IsNull(key) Then key = Request.Form("key")
If key = "" OR IsNull(key) Then Response.Redirect "carlist.asp"
'get action
a=Request.Form("a")
If a="" OR IsNull(a) Then
	a="I"	'display with input box
End If
' Open Connection to the database
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str
Select Case a
	Case "I": ' Get a record to display
		tkey = "" & key & ""
		strsql = "SELECT * FROM [car] WHERE [ID]=" & tkey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn
		If rs.EOF Then
			Response.Clear
			Response.Redirect "carlist.asp"
		Else
			rs.MoveFirst
		End If
		' Get the field contents
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
		rs.Close
		Set rs = Nothing
End Select
%>

<meta name="generator" content="Namo WebEditor v5.0(Trial)">
<div align="left">
    <table border="2" bordercolordark="black" bordercolorlight="black" cellspacing="0" width="668">
        <tr>
            <td width="660">

                <table cellpadding="0" cellspacing="0" bordercolordark="black" bordercolorlight="black" align="center" width="640">
                    <tr>
                        <td width="640" valign="top" colspan="4" align="center" height="24">
                <p align="center"><font face="Arial"><b><span style="font-size:28pt;"><u>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;BUYERS 
                GUIDE &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</u></span></b></font>                        </td>
                    </tr>
                    <tr>
                        <td width="640" colspan="4">
                            <p align="center"><font face="Arial"><span style="font-size:9pt;">IMPORTANT: 
                            Spoken promises are difficult to enforce. Ask the 
                            dealer to put all promises in writing. Keep this 
                            form.</span></font></p>
                        </td>
                    </tr>
                    <tr>
                        <td width="160">
                            <p>&nbsp;<u>&nbsp;&nbsp;&nbsp;</u><font face="Arial"><span style="font-size:11pt;"><u><%
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
%> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</u></span></font></p>
                        </td>
                        <td width="160">
                            <p>&nbsp;<u>&nbsp;&nbsp;&nbsp;</u><font face="Arial"><span style="font-size:11pt;"><u><% response.write x_model %> 
                            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</u></span></font></p>
                        </td>
                        <td width="115">
                            <p><u>&nbsp;&nbsp;&nbsp;</u><font face="Arial"><span style="font-size:11pt;"><u><%
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
%> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</u></span></font></p>
                        </td>
                        <td width="205">
                            <p>&nbsp;<u>&nbsp;&nbsp;&nbsp;&nbsp;</u><font face="Arial"><span style="font-size:11pt;"><u><% response.write x_vin %> 
                            &nbsp;&nbsp;&nbsp;&nbsp;</u></span></font></p>
                        </td>
                    </tr>
                    <tr>
                        <td width="160">
                            <p><font face="Arial"><span style="font-size:8pt;">&nbsp;Vehicle 
                            Name</span></font></p>
                        </td>
                        <td width="160">
                            <p><font face="Arial"><span style="font-size:8pt;">&nbsp;Model</span></font></p>
                        </td>
                        <td width="115">
                            <p><font face="Arial"><span style="font-size:8pt;">&nbsp;Year</span></font></p>
                        </td>
                        <td width="205">
                            <p><font face="Arial"><span style="font-size:8pt;">&nbsp;Vin 
                            Number</span></font></p>
                        </td>
                    </tr>
                    <tr>
                        <td width="320" colspan="2">
                            <p>&nbsp;<u>&nbsp;&nbsp;&nbsp;</u><font face="Arial"><span style="font-size:11pt;"><u><% response.write x_stock %> 
                            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</u></span></font></p>
                        </td>
                        <td width="115">
                            <p>&nbsp;</p>
                        </td>
                        <td width="205">
                            <p>&nbsp;</p>
                        </td>
                    </tr>
                    <tr>
                        <td width="640" colspan="4" height="15">
                            <p><font face="Arial"><span style="font-size:8pt;">Dealer 
                            Stock Number (Optional)</span></font></p>
                        </td>
                    </tr>
                    <tr>
                        <td width="320" colspan="2">
                            <p>&nbsp;</p>
                        </td>
                        <td width="115">
                            <p>&nbsp;</p>
                        </td>
                        <td width="205">
                            <p>&nbsp;</p>
                        </td>
                    </tr>
                    <tr>
                        <td width="640" colspan="4">
                            <p><font face="Arial"><span style="font-size:9pt;"><b><u>WARRANTIES 
                            FOR THIS VEHICLE &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</u></b></span></font></p>
                        </td>
                    </tr>
                    <tr>
                        <td width="320" colspan="2">
                            <p>&nbsp;<img src="images/x.gif" width="49" height="49" border="0"></p>
                        </td>
                        <td width="115">
                            <p>&nbsp;</p>
                        </td>
                        <td width="205">
                            <p>&nbsp;</p>
                        </td>
                    </tr>
                    <tr>
                        <td width="320" colspan="2">
                            <p>&nbsp;</p>
                        </td>
                        <td width="115">
                            <p>&nbsp;</p>
                        </td>
                        <td width="205">
                            <p>&nbsp;</p>
                        </td>
                    </tr>
                    <tr>
                        <td width="320" colspan="2">
                            <p>&nbsp;</p>
                        </td>
                        <td width="115">
                            <p>&nbsp;</p>
                        </td>
                        <td width="205">
                            <p>&nbsp;</p>
                        </td>
                    </tr>
                    <tr>
                        <td width="320" colspan="2">
                            <p>&nbsp;</p>
                        </td>
                        <td width="115">
                            <p>&nbsp;</p>
                        </td>
                        <td width="205">
                            <p>&nbsp;</p>
                        </td>
                    </tr>
                    <tr>
                        <td width="320" colspan="2">
                            <p>&nbsp;</p>
                        </td>
                        <td width="115">
                            <p>&nbsp;</p>
                        </td>
                        <td width="205">
                            <p>&nbsp;</p>
                        </td>
                    </tr>
                    <tr>
                        <td width="320" colspan="2">
                            <p>&nbsp;</p>
                        </td>
                        <td width="115">
                            <p>&nbsp;</p>
                        </td>
                        <td width="205">
                            <p>&nbsp;</p>
                        </td>
                    </tr>
                    <tr>
                        <td width="320" colspan="2">
                            <p>&nbsp;</p>
                        </td>
                        <td width="115">
                            <p>&nbsp;</p>
                        </td>
                        <td width="205">
                            <p>&nbsp;</p>
                        </td>
                    </tr>
                    <tr>
                        <td width="320" colspan="2">
                            <p>&nbsp;</p>
                        </td>
                        <td width="115">
                            <p>&nbsp;</p>
                        </td>
                        <td width="205">
                            <p>&nbsp;</p>
                        </td>
                    </tr>
                    <tr>
                        <td width="320" colspan="2">
                            <p>&nbsp;</p>
                        </td>
                        <td width="115">
                            <p>&nbsp;</p>
                        </td>
                        <td width="205">
                            <p>&nbsp;</p>
                        </td>
                    </tr>
                </table>
                <p align="center"><font face="Arial"><b><span style="font-size:26pt;">&nbsp;</span></b></font></p>
                <p align="center"><font face="Arial"><b><span style="font-size:26pt;">&nbsp;</span></b></font></p>
                <p align="center"><font face="Arial"><b><span style="font-size:10pt;">IMPORTANT</span></b></font></p>
            </td>
        </tr>
    </table>
</div>
<p align="left">&nbsp;</p>
<p align="left">&nbsp;</p>
<p align="left">&nbsp;</p>
<div align="left">
    <table cellpadding="0" cellspacing="0" width="595">
        <tr>
            <td width="595" colspan="6">
                <p>&nbsp;</p>
            </td>
        </tr>
        <tr>
            <td width="46" height="17">
                <p>&nbsp;</p>
            </td>
            <td width="133" height="17">            
<p><font face="Arial"><span style="font-size:11pt;"><%
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
%></span></font></td>
            <td width="130" height="17">            
<p><font face="Arial"><span style="font-size:11pt;"><% response.write x_model %></span></font></td>
            <td width="75" height="17"><font face="Arial"><span style="font-size:11pt;"><%
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
            <td width="176" height="17"><font face="Arial"><span style="font-size:11pt;"><% response.write x_vin %></span></font></td>
            <td width="35" height="17">
                <p>&nbsp;</p>
            </td>
        </tr>
        <tr>
            <td width="595" colspan="6" height="12">
                <p>&nbsp;</p>
            </td>
        </tr>
        <tr>
            <td width="46" height="10">
                <p>&nbsp;</p>
            </td>
            <td width="514" height="10" colspan="4"><font face="Arial"><span style="font-size:11pt;"><% response.write x_stock %></span></font></td>
            <td width="35" height="10">
                <p>&nbsp;</p>
            </td>
        </tr>
        <tr>
            <td width="595" colspan="6">
                <p>&nbsp;</p>
            </td>
        </tr>
    </table>
</div>
