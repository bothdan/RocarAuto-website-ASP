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
<body>
<form>
<table border="0" cellspacing="0" cellpadding="0" bgcolor="#CCCCCC">
<tr>
<td bgcolor="#F5F5F5"><% If not isnull(x_photo_1) Then %>
<font face="Arial" size="2"><a href="car_photo_1_bv.asp?key=<%= key %>" target="_self"><img src="car_photo_1_bv.asp?key=<%= key %>" border=0></a>
<% End If %>
</font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#F5F5F5"><% If not isnull(x_photo_2) Then %>
<font face="Arial" size="2"><a href="car_photo_2_bv.asp?key=<%= key %>" target="_self"><img src="car_photo_2_bv.asp?key=<%= key %>" border=0></a>
<% End If %>
</font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#F5F5F5"><% If not isnull(x_photo_3) Then %>
<font face="Arial" size="2"><a href="car_photo_3_bv.asp?key=<%= key %>" target="_self"><img src="car_photo_3_bv.asp?key=<%= key %>" border=0></a>
<% End If %>
</font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#F5F5F5"><% If not isnull(x_photo_4) Then %>
<font face="Arial" size="2"><a href="car_photo_4_bv.asp?key=<%= key %>" target="_self"><img src="car_photo_4_bv.asp?key=<%= key %>" border=0></a>
<% End If %>
</font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#F5F5F5"><% If not isnull(x_photo_5) Then %>
<font face="Arial" size="2"><a href="car_photo_5_bv.asp?key=<%= key %>" target="blank"><img src="car_photo_5_bv.asp?key=<%= key %>" border=0></a>
<% End If %>
</font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#F5F5F5"><% If not isnull(x_photo_6) Then %>
<font face="Arial" size="2"><a href="car_photo_6_bv.asp?key=<%= key %>" target="blank"><img src="car_photo_6_bv.asp?key=<%= key %>" border=0></a>
<% End If %>
</font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#F5F5F5"><% If not isnull(x_photo_7) Then %>
<font face="Arial" size="2"><a href="car_photo_7_bv.asp?key=<%= key %>" target="blank"><img src="car_photo_7_bv.asp?key=<%= key %>" border=0></a>
<% End If %>
</font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#F5F5F5"><% If not isnull(x_photo_8) Then %>
<font face="Arial" size="2"><a href="car_photo_8_bv.asp?key=<%= key %>" target="blank"><img src="car_photo_8_bv.asp?key=<%= key %>" border=0></a>
<% End If %>
</font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#F5F5F5"><% If not isnull(x_photo_9) Then %>
<font face="Arial" size="2"><a href="car_photo_9_bv.asp?key=<%= key %>" target="blank"><img src="car_photo_9_bv.asp?key=<%= key %>" border=0></a>
<% End If %>
</font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#F5F5F5"><% If not isnull(x_photo_10) Then %>
<font face="Arial" size="2"><a href="car_photo_10_bv.asp?key=<%= key %>" target="blank"><img src="car_photo_10_bv.asp?key=<%= key %>" border=0></a>
<% End If %>
</font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#F5F5F5"><% If not isnull(x_photo_11) Then %>
<font face="Arial" size="2"><a href="car_photo_11_bv.asp?key=<%= key %>" target="blank"><img src="car_photo_11_bv.asp?key=<%= key %>" border=0></a>
<% End If %>
</font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#F5F5F5"><% If not isnull(x_photo_12) Then %>
<font face="Arial" size="2"><a href="car_photo_12_bv.asp?key=<%= key %>" target="blank"><img src="car_photo_12_bv.asp?key=<%= key %>" border=0></a>
<% End If %>
</font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#F5F5F5"><% If not isnull(x_photo_13) Then %>
<font face="Arial" size="2"><a href="car_photo_13_bv.asp?key=<%= key %>" target="blank"><img src="car_photo_13_bv.asp?key=<%= key %>" border=0></a>
<% End If %>
</font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#F5F5F5"><% If not isnull(x_photo_14) Then %>
<font face="Arial" size="2"><a href="car_photo_14_bv.asp?key=<%= key %>" target="blank"><img src="car_photo_14_bv.asp?key=<%= key %>" border=0></a>
<% End If %>
</font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#F5F5F5"><% If not isnull(x_photo_15) Then %>
<font face="Arial" size="2"><a href="car_photo_15_bv.asp?key=<%= key %>" target="blank"><img src="car_photo_15_bv.asp?key=<%= key %>" border=0></a>
<% End If %>
</font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#F5F5F5"><% If not isnull(x_photo_16) Then %>
<font face="Arial" size="2"><a href="car_photo_16_bv.asp?key=<%= key %>" target="blank"><img src="car_photo_16_bv.asp?key=<%= key %>" border=0></a>
<% End If %>
</font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#F5F5F5"><% If not isnull(x_photo_17) Then %>
<font face="Arial" size="2"><a href="car_photo_17_bv.asp?key=<%= key %>" target="blank"><img src="car_photo_17_bv.asp?key=<%= key %>" border=0></a>
<% End If %>
</font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#F5F5F5"><% If not isnull(x_photo_18) Then %>
<font face="Arial" size="2"><a href="car_photo_18_bv.asp?key=<%= key %>" target="blank"><img src="car_photo_18_bv.asp?key=<%= key %>" border=0></a>
<% End If %>
</font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#F5F5F5"><% If not isnull(x_photo_19) Then %>
<font face="Arial" size="2"><a href="car_photo_19_bv.asp?key=<%= key %>" target="blank"><img src="car_photo_19_bv.asp?key=<%= key %>" border=0></a>
<% End If %>
</font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#F5F5F5"><% If not isnull(x_photo_20) Then %>
<font face="Arial" size="2"><a href="car_photo_20_bv.asp?key=<%= key %>" target="blank"><img src="car_photo_20_bv.asp?key=<%= key %>" border=0></a>
<% End If %>
</font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#F5F5F5"><% If not isnull(x_photo_21) Then %>
<font face="Arial" size="2"><a href="car_photo_21_bv.asp?key=<%= key %>" target="blank"><img src="car_photo_21_bv.asp?key=<%= key %>" border=0></a>
<% End If %>
</font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#F5F5F5"><% If not isnull(x_photo_22) Then %>
<font face="Arial" size="2"><a href="car_photo_22_bv.asp?key=<%= key %>" target="blank"><img src="car_photo_22_bv.asp?key=<%= key %>" border=0></a>
<% End If %>
</font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#F5F5F5"><% If not isnull(x_photo_23) Then %>
<font face="Arial" size="2"><a href="car_photo_23_bv.asp?key=<%= key %>" target="blank"><img src="car_photo_23_bv.asp?key=<%= key %>" border=0></a>
<% End If %>
</font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#F5F5F5"><% If not isnull(x_photo_24) Then %>
<font face="Arial" size="2"><a href="car_photo_24_bv.asp?key=<%= key %>" target="blank"><img src="car_photo_24_bv.asp?key=<%= key %>" border=0></a>
<% End If %>
</font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#F5F5F5"><% If not isnull(x_photo_25) Then %>
<font face="Arial" size="2"><a href="car_photo_25_bv.asp?key=<%= key %>" target="blank"><img src="car_photo_25_bv.asp?key=<%= key %>" border=0></a>
<% End If %>
</font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#F5F5F5"><% If not isnull(x_photo_26) Then %>
<font face="Arial" size="2"><a href="car_photo_26_bv.asp?key=<%= key %>" target="blank"><img src="car_photo_26_bv.asp?key=<%= key %>" border=0></a>
<% End If %>
</font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#F5F5F5"><% If not isnull(x_photo_27) Then %>
<font face="Arial" size="2"><a href="car_photo_27_bv.asp?key=<%= key %>" target="blank"><img src="car_photo_27_bv.asp?key=<%= key %>" border=0></a>
<% End If %>
</font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#F5F5F5"><% If not isnull(x_photo_28) Then %>
<font face="Arial" size="2"><a href="car_photo_28_bv.asp?key=<%= key %>" target="blank"><img src="car_photo_28_bv.asp?key=<%= key %>" border=0></a>
<% End If %>
</font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#F5F5F5"><% If not isnull(x_photo_29) Then %>
<font face="Arial" size="2"><a href="car_photo_29_bv.asp?key=<%= key %>" target="blank"><img src="car_photo_29_bv.asp?key=<%= key %>" border=0></a>
<% End If %>
</font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#F5F5F5"><% If not isnull(x_photo_30) Then %>
<font face="Arial" size="2"><a href="car_photo_30_bv.asp?key=<%= key %>" target="blank"><img src="car_photo_30_bv.asp?key=<%= key %>" border=0></a>
<% End If %>
</font>&nbsp;</td>
</tr>
</table>
</form>
</body>
</html>