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
If key = "" OR IsNull(key) Then Response.Redirect "findelist.asp"
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
		strsql = "SELECT * FROM [finde] WHERE [ID]=" & tkey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn
		If rs.EOF Then
			Response.Clear
			Response.Redirect "findelist.asp"
		Else
			rs.MoveFirst
		End If
		' Get the field contents
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
		rs.Close
		Set rs = Nothing
End Select
%>
<!--#include file="header.asp"-->
<meta name="generator" content="Namo WebEditor v5.0(Trial)">
<table align="center" cellpadding="0" cellspacing="0" width="801" bgcolor="white">
    <tr>
        <td>
            <p><font face="Arial" size="2"><br> </font></p>
<p>
<form>
<table border="0" cellspacing="0" cellpadding="0" bgcolor="#CCCCCC" width="553">
<tr>
<td bgcolor="white" width="54">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="124"><span style="font-size:14pt;"><b><font face="Arial">Find 
                            a Car:<br>&nbsp;</font></b></span></td>
<td bgcolor="white" width="375"><font face="Arial" color="white"><span style="font-size:10pt;"><%= x_ID %>&nbsp;</span></font></td>
</tr>
<tr>
<td bgcolor="white" width="54">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="124"><font face="Arial"><span style="font-size:10pt;"><b>First Name</b></span></font><span style="font-size:10pt;"><b><font face="Arial">:</font></b></span></td>
<td bgcolor="white" width="375"><font face="Arial"><span style="font-size:10pt;"><% response.write x_first_name %>&nbsp;</span></font></td>
</tr>
<tr>
<td bgcolor="white" width="54">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="124"><font face="Arial"><span style="font-size:10pt;"><b>Last Name</b></span></font><span style="font-size:10pt;"><b><font face="Arial">:</font></b></span></td>
<td bgcolor="white" width="375"><font face="Arial"><span style="font-size:10pt;"><% response.write x_last_name %>&nbsp;</span></font></td>
</tr>
<tr>
<td bgcolor="white" width="54">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="124"><font face="Arial"><span style="font-size:10pt;"><b>Home Phone</b></span></font><span style="font-size:10pt;"><b><font face="Arial">&nbsp;</font></b></span></td>
<td bgcolor="white" width="375"><font face="Arial"><span style="font-size:10pt;"><% response.write x_home_phone %>&nbsp;</span></font></td>
</tr>
<tr>
<td bgcolor="white" width="54">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="124"><font face="Arial"><span style="font-size:10pt;"><b>E-mail</b></span></font><span style="font-size:10pt;"><b><font face="Arial">&nbsp;Address:</font></b></span></td>
<td bgcolor="white" width="375"><font face="Arial"><span style="font-size:10pt;"><% response.write x_email %>&nbsp;</span></font></td>
</tr>
<tr>
<td bgcolor="white" width="54">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="124"><font face="Arial"><span style="font-size:10pt;"><b>Type</b></span></font><span style="font-size:10pt;"><b><font face="Arial">:</font></b></span></td>
<td bgcolor="white" width="375"><font face="Arial"><span style="font-size:10pt;"><% response.write x_type %>&nbsp;</span></font></td>
</tr>
<tr>
<td bgcolor="white" width="54">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="124"><span style="font-size:10pt;"><b><font face="Arial">&nbsp;</font></b></span></td>
<td bgcolor="white" width="375"><font face="Arial"><span style="font-size:10pt;">&nbsp;</span></font></td>
</tr>
<tr>
<td bgcolor="white" width="54">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="124"><font face="Arial"><span style="font-size:10pt;"><b>Years 
                            Range:</b></span></font></td>
<td bgcolor="white" width="375">                            <p align="left"><font face="Verdana"><span style="font-size:9pt;"><% response.write x_yearold %>&nbsp;- 
                            <% response.write x_yearnew %></span></font></td>
</tr>
<tr>
<td bgcolor="white" width="54">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="124"><font face="Arial"><span style="font-size:10pt;"><b>Make</b></span></font><span style="font-size:10pt;"><b><font face="Arial">:</font></b></span></td>
<td bgcolor="white" width="375"><font face="Arial"><span style="font-size:10pt;"><%
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
&nbsp;</span></font></td>
</tr>
<tr>
<td bgcolor="white" width="54">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="124"><font face="Arial"><span style="font-size:10pt;"><b>Model</b></span></font><span style="font-size:10pt;"><b><font face="Arial">:</font></b></span></td>
<td bgcolor="white" width="375"><font face="Arial"><span style="font-size:10pt;"><% response.write x_model %>&nbsp;</span></font></td>
</tr>
<tr>
<td bgcolor="white" width="54">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="124"><font face="Arial"><span style="font-size:10pt;"><b>Bodystyle</b></span></font><span style="font-size:10pt;"><b><font face="Arial">:</font></b></span></td>
<td bgcolor="white" width="375"><font face="Arial"><span style="font-size:10pt;"><%
Select Case x_bodystyle
    Case "Cupe" response.write "Cupe"
    Case "Convertible" response.write "Convertible"
    Case "Sedan" response.write "Sedan"
    Case "SUV" response.write "SUV"
    Case "Wagon" response.write "Wagon"
End Select
%>
&nbsp;</span></font></td>
</tr>
<tr>
<td bgcolor="white" width="54">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="124"><font face="Arial"><span style="font-size:10pt;"><b>Transmission</b></span></font><span style="font-size:10pt;"><b><font face="Arial">:</font></b></span></td>
<td bgcolor="white" width="375"><font face="Arial"><span style="font-size:10pt;"><%
Select Case x_transmission
    Case "Automatic" response.write "Automatic"
    Case "Manual" response.write "Manual"
End Select
%>
&nbsp;</span></font></td>
</tr>
<tr>
<td bgcolor="white" width="54">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="124"><font face="Arial"><span style="font-size:10pt;"><b>Mileage 
                            </b></span></font><span style="font-size:10pt;"><b><font face="Arial">R</font></b></span><font face="Arial"><span style="font-size:10pt;"><b>ange:</b></span></font></td>
<td bgcolor="white" width="375">                            <p align="left"> 
                            <font face="Arial" size="2"><% if isnumeric(x_mileagelow) then response.write formatnumber(x_mileagelow,0,-2,-2,-2) else response.write x_mileagelow end if %> 
                            </font><font face="Verdana"><span style="font-size:10pt;">-
                            </span></font><font face="Arial" size="2"><% if isnumeric(x_mileagehi) then response.write formatnumber(x_mileagehi,0,-2,-2,-2) else response.write x_mileagehi end if %></font></td>
</tr>
<tr>
<td bgcolor="white" width="54">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="124"><span style="font-size:10pt;"><b><font face="Arial">Price 
                            Range:</font></b></span></td>
<td bgcolor="white" width="375">                            <p align="left"><font face="Arial" size="2"><% if isnumeric(x_pricelow) then response.write formatcurrency(x_pricelow,0,-2,-2,-2) else response.write x_pricelow end if %></font><font face="Verdana"><span style="font-size:10pt;"> 
                            -
                            </span></font><font face="Arial" size="2"><% if isnumeric(x_pricehi) then response.write formatcurrency(x_pricehi,0,-2,-2,-2) else response.write x_pricehi end if %></font></td>
</tr>
<tr>
<td bgcolor="white" width="54">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="124"><span style="font-size:10pt;"><b><font face="Arial">&nbsp;</font></b></span></td>
<td bgcolor="white" width="375"><font face="Arial"><span style="font-size:10pt;">&nbsp;</span></font></td>
</tr>
<tr>
<td bgcolor="white" width="54">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="124"><font face="Arial"><span style="font-size:10pt;"><b>Comments</b></span></font><span style="font-size:10pt;"><b><font face="Arial">:</font></b></span></td>
<td bgcolor="white" width="375"><font face="Arial"><span style="font-size:10pt;"><%= replace(x_comments & "",chr(10),"<br>") %>&nbsp;</span></font></td>
</tr>
<tr>
<td bgcolor="white" width="54">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="124">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="375">
                            <p>&nbsp;</p>
</td>
</tr>
<tr>
<td bgcolor="white" width="54">
                            <p>&nbsp;</p>
</td>
<td bgcolor="white" width="124"><font face="Arial"><span style="font-size:10pt;"><b>Time 
                            </b></span></font><span style="font-size:10pt;"><b><font face="Arial">F</font></b></span><font face="Arial"><span style="font-size:10pt;"><b>rame:</b></span></font></td>
<td bgcolor="white" width="375"><font face="Arial"><span style="font-size:10pt;"><% response.write x_time %>&nbsp;</span></font></td>
</tr>
</table>
</form>
            <p><font face="Arial" size="2" color="black"><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</b></font><a href="findelist.asp"><font face="Arial" size="2" color="black"><b><img src="images/leftsm.gif" align="texttop" width="16" height="16" border="0"></b></font></a><font face="Arial" size="2" color="black"><b> 
            &nbsp;&nbsp;</b></font><a href="findelist.asp"><font face="Arial" size="2" color="black"><b>Back to Find a Car List</b></font></a><font face="Arial" size="2"><br>&nbsp;</font></p>
        </td>
    </tr>
</table>
<!--#include file="footer.asp"-->
