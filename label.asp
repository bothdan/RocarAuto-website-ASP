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
<p><font face="Arial" size="2">View TABLE: car<br><br><a href="carlist.asp">Back to List</a></font></p>
<p>
<form>
<table border="0" cellspacing="0" cellpadding="0" bgcolor="#CCCCCC">
<tr>
<td bgcolor="#708090"><font color="#FFFFFF"><font face="Arial" size="2">ID</font>&nbsp;</font></td>
<td bgcolor="#F5F5F5"><font face="Arial" size="2"><%= x_ID %></font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#708090"><font color="#FFFFFF"><font face="Arial" size="2">Year</font>&nbsp;</font></td>
<td bgcolor="#F5F5F5"><font face="Arial" size="2"><%
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
%>
</font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#708090"><font color="#FFFFFF"><font face="Arial" size="2">Make</font>&nbsp;</font></td>
<td bgcolor="#F5F5F5"><font face="Arial" size="2"><%
Select Case x_make
    Case "Audi" response.write "Audi"
    Case "Acura" response.write "Acura"
    Case "BMW" response.write "BMW"
    Case "" response.write ""
End Select
%>
</font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#708090"><font color="#FFFFFF"><font face="Arial" size="2">Model</font>&nbsp;</font></td>
<td bgcolor="#F5F5F5"><font face="Arial" size="2"><% response.write x_model %></font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#708090"><font color="#FFFFFF"><font face="Arial" size="2">Type</font>&nbsp;</font></td>
<td bgcolor="#F5F5F5"><font face="Arial" size="2"><%
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
%>
</font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#708090"><font color="#FFFFFF"><font face="Arial" size="2">Miles</font>&nbsp;</font></td>
<td bgcolor="#F5F5F5"><font face="Arial" size="2"><% if isnumeric(x_miles) then response.write formatnumber(x_miles,0,-2,-2,-2) else response.write x_miles end if %></font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#708090"><font color="#FFFFFF"><font face="Arial" size="2">Price</font>&nbsp;</font></td>
<td bgcolor="#F5F5F5"><font face="Arial" size="2"><% if isnumeric(x_price) then response.write formatcurrency(x_price,0,-2,-2,-2) else response.write x_price end if %></font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#708090"><font color="#FFFFFF"><font face="Arial" size="2">Doors</font>&nbsp;</font></td>
<td bgcolor="#F5F5F5"><font face="Arial" size="2"><%
Select Case x_doors
    Case "5" response.write "5"
    Case "4" response.write "4"
    Case "3" response.write "3"
    Case "2" response.write "2"
    Case "1" response.write "1"
End Select
%>
</font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#708090"><font color="#FFFFFF"><font face="Arial" size="2">Engine</font>&nbsp;</font></td>
<td bgcolor="#F5F5F5"><font face="Arial" size="2"><% response.write x_engine %></font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#708090"><font color="#FFFFFF"><font face="Arial" size="2">Transmission</font>&nbsp;</font></td>
<td bgcolor="#F5F5F5"><font face="Arial" size="2"><%
Select Case x_transmission
    Case "Automatic" response.write "Automatic"
    Case "Manual" response.write "Manual"
End Select
%>
</font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#708090"><font color="#FFFFFF"><font face="Arial" size="2">Drivetrain</font>&nbsp;</font></td>
<td bgcolor="#F5F5F5"><font face="Arial" size="2"><%
Select Case x_drivetrain
    Case "FWD" response.write "FWD"
    Case "RWD" response.write "RWD"
    Case "AWD" response.write "AWD"
End Select
%>
</font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#708090"><font color="#FFFFFF"><font face="Arial" size="2">Exterior color</font>&nbsp;</font></td>
<td bgcolor="#F5F5F5"><font face="Arial" size="2"><% response.write x_ext_color %></font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#708090"><font color="#FFFFFF"><font face="Arial" size="2">Interior color</font>&nbsp;</font></td>
<td bgcolor="#F5F5F5"><font face="Arial" size="2"><% response.write x_int_color %></font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#708090"><font color="#FFFFFF"><font face="Arial" size="2">Stock #</font>&nbsp;</font></td>
<td bgcolor="#F5F5F5"><font face="Arial" size="2"><% response.write x_stock %></font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#708090"><font color="#FFFFFF"><font face="Arial" size="2">VIN</font>&nbsp;</font></td>
<td bgcolor="#F5F5F5"><font face="Arial" size="2"><% response.write x_vin %></font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#708090"><font color="#FFFFFF"><font face="Arial" size="2">City MPG</font>&nbsp;</font></td>
<td bgcolor="#F5F5F5"><font face="Arial" size="2"><% response.write x_city_mpg %></font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#708090"><font color="#FFFFFF"><font face="Arial" size="2">Hwy MPG</font>&nbsp;</font></td>
<td bgcolor="#F5F5F5"><font face="Arial" size="2"><% response.write x_hwy_mpg %></font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#708090"><font color="#FFFFFF"><font face="Arial" size="2">Carfax</font>&nbsp;</font></td>
<td bgcolor="#F5F5F5"><font face="Arial" size="2"><%
Select Case x_carfax
    Case "Yes" response.write "Yes"
    Case "No" response.write "No"
End Select
%>
</font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#708090"><font color="#FFFFFF"><font face="Arial" size="2">Special</font>&nbsp;</font></td>
<td bgcolor="#F5F5F5"><font face="Arial" size="2"><%
Select Case x_special
    Case "Yes" response.write "Yes"
    Case "No" response.write "No"
End Select
%>
</font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#708090"><font color="#FFFFFF"><font face="Arial" size="2">Status</font>&nbsp;</font></td>
<td bgcolor="#F5F5F5"><font face="Arial" size="2"><%
Select Case x_status
    Case "For Sale" response.write "For Sale"
    Case "Sold" response.write "Sold"
End Select
%>
</font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#708090" colspan="2">&nbsp;</td>
</tr>
</table>
</form>
<p>
&nbsp;
<div align="left">
    <p class="MsoTableGrid" style="margin-left:-30.6pt; border-style:none; border-collapse:collapse;">&nbsp;</p>
</div>
<table class="MsoTableGrid" style="margin-left:-30.6pt; border-width:7px; border-style:none; border-collapse:collapse;" border="5" cellspacing="0" width="516" bordercolor="#003366" bordercolordark="#003366" bordercolorlight="#003366" align="center">
    <tr>
        <td width="502" height="136" colspan="2" bordercolor="#003366" bordercolordark="#003366" bordercolorlight="#003366">
            <p>&nbsp;</p>
        </td>
    </tr>
    <tr>
        <td width="302" height="242">
            <p>&nbsp;</p>
        </td>
        <td width="196" height="752" rowspan="3">
            <p>&nbsp;</p>
        </td>
    </tr>
    <tr>
        <td width="302" height="302">
            <p>&nbsp;</p>
        </td>
    </tr>
    <tr>
        <td width="302" height="206">
            <p>&nbsp;</p>
        </td>
    </tr>
</table>
<p class="MsoTableGrid" style="margin-left:-30.6pt; border-style:none; border-collapse:collapse;" align="left">&nbsp;</p>
