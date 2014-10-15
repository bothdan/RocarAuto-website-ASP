
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
<body>



<script language="JavaScript">
<!--
function na_open_window(name, url, left, top, width, height, toolbar, menubar, statusbar, scrollbar, resizable)
{
  toolbar_str = toolbar ? 'yes' : 'no';
  menubar_str = menubar ? 'yes' : 'no';
  statusbar_str = statusbar ? 'yes' : 'no';
  scrollbar_str = scrollbar ? 'yes' : 'no';
  resizable_str = resizable ? 'yes' : 'no';
  window.open(url, name, 'left='+left+',top='+top+',width='+width+',height='+height+',toolbar='+toolbar_str+',menubar='+menubar_str+',status='+statusbar_str+',scrollbars='+scrollbar_str+',resizable='+resizable_str);
}

// -->
</script>
<a href="javascript:window.print()">print page</a>
<br>&nbsp;<div align="left">
<table width="664" bgcolor="white" cellspacing="0" bordercolordark="white" bordercolorlight="black" cellpadding="0">
    <tr>
        <td width="353">
            <font face="Arial"><span style="font-size:14pt;"><b><%
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
%> &nbsp;<%
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
%> &nbsp;<% response.write x_model %> &nbsp;<%
Select Case x_drivetrain
    Case "FWD" response.write "FWD"
    Case "RWD" response.write "RWD"
    Case "AWD" response.write "AWD"
End Select
%><br></b></span></font>        </td>
        <td width="10">
            <font face="Arial"><span style="font-size:12pt;"><b>&nbsp;</b></span></font>        </td>
        <td width="301">
            <font face="Arial"><span style="font-size:14pt;"><b>&nbsp;</b></span></font>        </td>
    </tr>
    <tr>
        <td width="353" valign="top">
                <div align="left">
            <table border="0" cellspacing="0" bordercolordark="white" bordercolorlight="#666666" bordercolor="#666666">
                <tr>
                    <td width="337">                <table border="0" cellspacing="0" bordercolordark="white" bordercolorlight="#0066FF" bordercolor="black" cellpadding="1" align="center">
                    <tr>
                        <td width="216">
                                    <table align="center" bordercolorlight="#333333" cellspacing="0" bordercolordark="#333333" cellpadding="0">
                                        <tr>
                                            <td width="240">
                                                <p><% If not isnull(x_photo_1) Then %>
<font face="Arial" size="2"><img src="car_ph.asp?key=<%= key %>&nr=1" border=0 width="250">
<% End If %></font></p>
                                            </td>
                                        </tr>
                                    </table>
</td>
                    </tr>
                    <tr>
                        <td width="270">
                        <p align="center"><font face="Arial" color="#990000"><span style="font-size:14pt;">Price: 
                        <% if isnumeric(x_price) then response.write formatcurrency(x_price,0,-2,-2,-2) else response.write x_price end if %></span></font></p>
</td>
                    </tr>
                </table>
                    </td>
                </tr>
            </table>
                </div>
        </td>
        <td width="10">
            <p>&nbsp;</p>
        </td>
        <td width="301" valign="top">
            <table align="center" cellpadding="0" cellspacing="0" bgcolor="white" bordercolordark="white" bordercolorlight="black" width="360">
                <tr>
                    <td width="360">
                        <p><font face="Arial" color="#990000"><span style="font-size:10pt;"><b>Call 
                        - (773) 334-0025<br></b></span></font></p>
                    </td>
                </tr>
                <tr>
                    <td width="360"><font face="Verdana"><b><span style="font-size:8pt;"><br>Vehicle Location: Rocar Auto Sales<BR>5136 N. Western Ave. Chicago, IL 
60625<br>&nbsp;</span></b></font></td>
                </tr>
                <tr>
                    <td width="227">
                        <p><font face="Verdana"><span style="font-size:8pt;"><b>Bodystyle: 
                        </b><%
Select Case x_doors
    Case "5" response.write "5"
    Case "4" response.write "4"
    Case "3" response.write "3"
    Case "2" response.write "2"
    Case "1" response.write "1"
End Select
%> door <%
Select Case x_type
    Case "Sedan" response.write "Sedan"
    Case "SUV" response.write "SUV"
    Case "Mini-Van" response.write "Mini-Van"
    Case "Wagon" response.write "Wagon"
    Case "Hatchback" response.write "Hatchback"
    Case "Coupe" response.write "Coupe"
    Case "Truck" response.write "Truck"
    Case "Convertible" response.write "Convertible"
    Case "Sport" response.write "Sport"
    Case "SUT" response.write "SUT"
End Select
%></span></font></p>
                    </td>
                </tr>
                <tr>
                    <td width="227">
                        <p><font face="Verdana"><span style="font-size:8pt;"><b>Engine: 
                        </b><% response.write x_engine %></span></font></p>
                    </td>
                </tr>
                <tr>
                    <td width="227"><font color="black" face="Verdana"><span style="font-size:8pt;"><b>Transmission: 
                        </b></span></font><font face="Verdana"><span style="font-size:8pt;"><%
Select Case x_transmission
    Case "Automatic" response.write "Automatic"
    Case "Manual" response.write "Manual"
End Select
%></span></font></td>
                </tr>
                <tr>
                    <td width="227">
                        <p><font face="Verdana"><span style="font-size:8pt;"><b>Ext. 
                        Color: </b></span></font><span style="font-size:8pt;"><font face="Verdana"><% response.write x_ext_color %></font></span></p>
                    </td>
                </tr>
                <tr>
                    <td width="227">
                        <p><b><span style="font-size:8pt;"><font face="Verdana">Int. 
                        Color: </font></span></b><span style="font-size:8pt;"><font face="Verdana"><% response.write x_int_color %></font></span></p>
                    </td>
                </tr>
                <tr>
                    <td width="227">
                        <p><b><span style="font-size:8pt;"><font face="Verdana">Mileage: 
                        </font></span></b><span style="font-size:8pt;"><font face="Verdana"><% if isnumeric(x_miles) then response.write formatnumber(x_miles,0,-2,-2,-2) else response.write x_miles end if %></font></span></p>
                    </td>
                </tr>
                <tr>
                    <td width="227">
                        <p><b><span style="font-size:8pt;"><font face="Verdana">Stock 
                        Number: </font></span></b><span style="font-size:8pt;"><font face="Verdana"><% response.write x_stock %></font></span></p>
                    </td>
                </tr>
                <tr>
                    <td width="360">
                        <p><b><span style="font-size:8pt;"><font face="Verdana">VIN 
                        Number: </font></span></b><span style="font-size:8pt;"><font face="Verdana"><% response.write x_vin %><br>&nbsp;</font></span></p>
                    </td>
                </tr>
                <tr>
                    <td width="227">
                        <table align="center" cellpadding="0" cellspacing="0" bordercolordark="white" bordercolorlight="#666666" width="176">
                            <tr>
                                <td width="65">
                                    <p><span style="font-size:8pt;"><font face="Verdana"><b>City 
                                    MPG</b></font></span></p>
                                </td>
                                <td width="46" rowspan="2">
                                    <p><img src="images/pumpicon.gif" width="38" height="36" border="0"></p>
                                </td>
                                <td width="65">
                                    <p><b><span style="font-size:8pt;"><font face="Verdana">Hwy 
                                    MPG</font></span></b></p>
                                </td>
                            </tr>
                            <tr>
                                <td width="65">
                                    <p align="center"><b><span style="font-size:16pt;"><font face="Verdana"><% response.write x_city_mpg %></font></span></b></td>
                                <td width="65">
                                    <p align="center"><font face="Verdana"><b><span style="font-size:16pt;"><% response.write x_hwy_mpg %></span></b></font></td>
                            </tr>
                            <tr>
                                <td width="176" colspan="3"><font face="Arial"><span style="font-size:7pt;">Actual rating will vary with options, driving conditions, habits and vehicle 
condition.<br>&nbsp;</span></font></td>
                            </tr>
                        </table>
                    </td>
                </tr>
                <tr>
                    <td width="360">
                            <p align="center">&nbsp;</p>
                    </td>
                </tr>
            </table>
        </td>
    </tr>
        <tr>
        <td width="668" valign="top" colspan="3">
                <p align="center"><font face="Arial"><span style="font-size:11pt;"><br>printed 
                on:</span></font><span style="font-size:11pt;"><font face="Arial"><% response.write date() %></font></span></p>
        </td>
        </tr>
</table>
</div>
<p>
&nbsp;</body>
</html>