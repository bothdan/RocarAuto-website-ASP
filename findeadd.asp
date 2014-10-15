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
'get action
a = Request.Form("a")
If (a = "" OR IsNull(a)) Then
	key = Request.Querystring("key")
	If key <> "" Then
		a = "C" 'copy record
	Else
		a = "I" 'display blank record
	End If
End If
' Open Connection to the database
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str
Select Case a
	Case "C": ' Get a record to display
		tkey = "" & key & ""
		strsql = "SELECT * FROM [finde] WHERE [ID]=" & tkey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn
		If rs.EOF Then
			Response.Clear
			Response.Redirect "carlist.asp"
		Else
			rs.MoveFirst
		' Get the field contents
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
		End If
		rs.Close
		Set rs = Nothing
	Case "A": ' Add
		'get fields from form
x_ID = Request.Form("x_ID")
x_first_name = Request.Form("x_first_name")
x_last_name = Request.Form("x_last_name")
x_home_phone = Request.Form("x_home_phone")
x_email = Request.Form("x_email")
x_type = Request.Form("x_type")
x_yearold = Request.Form("x_yearold")
x_yearnew = Request.Form("x_yearnew")
x_make = Request.Form("x_make")
x_model = Request.Form("x_model")
x_bodystyle = Request.Form("x_bodystyle")
x_transmission = Request.Form("x_transmission")
x_mileagelow = Request.Form("x_mileagelow")
x_mileagehi = Request.Form("x_mileagehi")
x_pricelow = Request.Form("x_pricelow")
x_pricehi = Request.Form("x_pricehi")
x_comments = Request.Form("x_comments")
x_time = Request.Form("x_time")
		' Open record
		strsql = "SELECT * FROM [finde] WHERE 0 = 1"
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn, 1, 2
		rs.AddNew
		tmpFld = Trim(x_first_name)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("first name") = tmpFld
		tmpFld = Trim(x_last_name)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("last name") = tmpFld
		tmpFld = Trim(x_home_phone)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("home phone") = tmpFld
		tmpFld = Trim(x_email)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("email") = tmpFld
		tmpFld = Trim(x_type)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("type") = tmpFld
		tmpFld = Trim(x_yearold)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("yearold") = tmpFld
		tmpFld = Trim(x_yearnew)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("yearnew") = tmpFld
		tmpFld = Trim(x_make)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("make") = tmpFld
		tmpFld = Trim(x_model)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("model") = tmpFld
		tmpFld = Trim(x_bodystyle)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("bodystyle") = tmpFld
		tmpFld = Trim(x_transmission)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("transmission") = tmpFld
		tmpFld = Trim(x_mileagelow)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("mileagelow") = tmpFld
		tmpFld = Trim(x_mileagehi)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("mileagehi") = tmpFld
		tmpFld = Trim(x_pricelow)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("pricelow") = tmpFld
		tmpFld = Trim(x_pricehi)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("pricehi") = tmpFld
		tmpFld = Trim(x_comments)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("comments") = tmpFld
		tmpFld = Trim(x_time)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("time") = tmpFld
		rs.Update
		rs.Close
		Set rs = Nothing
		conn.Close
		Set conn = Nothing
		Response.Clear
		Response.Redirect "findethanks.asp"
End Select
%>
<!--#include file="header.asp"-->
<meta name="generator" content="Namo WebEditor v5.0(Trial)">
<body text="black" link="blue" vlink="purple" alink="red">
<table align="center" border="1" cellspacing="0" width="803" bordercolor="black" bordercolordark="white" bordercolorlight="#999999" bgcolor="whitesmoke">
    <tr>
        <td width="797"><script language="JavaScript" src="ew.js"></script>
<script language="JavaScript">
<!-- start Javascript
function  EW_checkMyForm(EW_this) {
return true;
}
// end JavaScript -->
</script>
            <table align="center" border="1" bgcolor="white" width="604" cellspacing="0" bordercolordark="white" bordercolorlight="black">
                <tr>
                    <td width="594" height="47"><font face="Verdana"><span style="font-size:9pt;">Try our CarFinder and get notified when cars that interest you arrive! We will 
send you photos and details of the vehicle(s) that interest you most, 
automatically ... no hassles, no obligation and your request is confidential. 
                        </span></font><span class="required" style="font-size:9pt;"><font face="Verdana">( </font><font face="Verdana" color="#660000">* indicates required fields.</font><font face="Verdana"> )</font></span><font face="Verdana"><span style="font-size:9pt;"> </span></font></td>
                </tr>
            </table>
<form onSubmit="return EW_checkMyForm(this);"  action="findeadd.asp" method="post">
<p>
<input type="hidden" name="a" value="A">
<table border="0" cellspacing="0" cellpadding="2" bgcolor="whitesmoke" align="center" width="669">
<tr>
<td bgcolor="whitesmoke" width="327" colspan="2" height="29"><font color="black"><img src="images/contactinfocustomer.gif" width="246" height="33" border="0"></font><font face="Arial" size="2"><%= x_ID %><input type="hidden" name="x_ID" value="<%= x_ID %>"></font></td>
<td bgcolor="#F5F5F5" width="334" height="29" colspan="2"><img src="images/carinfo.gif" width="346" height="33" border="0"></td>
</tr>
<tr>
<td bgcolor="whitesmoke" width="132" height="13">
                            <p>&nbsp;</p>
</td>
<td bgcolor="whitesmoke" width="191" height="13">
                            <p>&nbsp;</p>
</td>
<td bgcolor="#F5F5F5" width="117" height="13"><font face="Verdana"><span style="font-size:10pt;">&nbsp;</span></font></td>
<td bgcolor="#F5F5F5" width="213" height="13"><font face="Verdana"><span style="font-size:10pt;">&nbsp;</span></font></td>
</tr>
<tr>
<td bgcolor="whitesmoke" width="132">
                            <p align="right"><font color="#990000" face="Verdana"><span style="font-size:10pt;">*First Name:</span></font></td>
<td bgcolor="whitesmoke" width="191"><font face="Arial" size="2"><input type="text" name="x_first_name" size="15" maxlength=50 value="<%= Server.HtmlEncode(x_first_name&"") %>"></font></td>
<td bgcolor="#F5F5F5" width="117"><font color="black" face="Verdana"><span style="font-size:10pt;">Type:</span></font></td>
<td bgcolor="#F5F5F5" width="213"><font face="Verdana"><span style="font-size:10pt;"><%
x_typeList = "<SELECT name='x_type'><OPTION value=''>Please Select</OPTION>"
    x_typeList = x_typeList & "<OPTION value=""Pre-Owned"""
    If x_type = "Pre-Owned" Then
        x_typeList = x_typeList & " selected"
    End If
    x_typeList = x_typeList & ">" & "Pre-Owned" & "</option>" 
    x_typeList = x_typeList & "</select>"
response.write x_typeList
%></span></font></td>
</tr>
<tr>
<td bgcolor="whitesmoke" width="132">
                            <p align="right"><font color="#990000" face="Verdana"><span style="font-size:10pt;">*Last Name:</span></font></td>
<td bgcolor="whitesmoke" width="191"><font face="Arial" size="2"><input type="text" name="x_last_name" size="15" maxlength=50 value="<%= Server.HtmlEncode(x_last_name&"") %>"></font></td>
<td bgcolor="#F5F5F5" width="117"><font face="Verdana"><span style="font-size:10pt;">Year:</span></font></td>
<td bgcolor="#F5F5F5" width="213"><font face="Verdana"><span style="font-size:10pt;"><input type="text" name="x_yearold" size="4" maxlength="4" value="<%= Server.HtmlEncode(x_yearold&"") %>">&nbsp;to<input type="text" name="x_yearnew" size="4" maxlength="4" value="<%= Server.HtmlEncode(x_yearnew&"") %>"></span></font></td>
</tr>
<tr>
<td bgcolor="whitesmoke" width="132">
                            <p align="right"><font color="#990000" face="Verdana"><span style="font-size:10pt;">*Home Phone:</span></font></td>
<td bgcolor="whitesmoke" width="191"><font face="Arial" size="2"><input type="text" name="x_home_phone" size="15" maxlength=50 value="<%= Server.HtmlEncode(x_home_phone&"") %>"></font></td>
<td bgcolor="#F5F5F5" width="117"><font face="Verdana"><span style="font-size:10pt;">Make:</span></font></td>
<td bgcolor="#F5F5F5" width="213"><font face="Verdana"><span style="font-size:10pt;"><%

x_makeList = "<SELECT name='x_make'><OPTION value=''>Please Select</OPTION>"
    x_makeList = x_makeList & "<OPTION value=""Acura"""
    If x_make = "Acura" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Acura" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Alfa Romeo"""
    If x_make = "Alfa Romeo" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Alfa Romeo" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Am General"""
    If x_make = "Am General" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Am General" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Aston Martin"""
    If x_make = "Aston Martin" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Aston Martin" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Audi"""
    If x_make = "Audi" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Audi" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""BMW"""
    If x_make = "BMW" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "BMW" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Bentley"""
    If x_make = "Bentley" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Bentley" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Buick"""
    If x_make = "Buick" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Buick" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Cadillac"""
    If x_make = "Cadillac" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Cadillac" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Chevrolet"""
    If x_make = "Chevrolet" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Chevrolet" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Chrysler"""
    If x_make = "Chrysler" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Chrysler" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Dacia"""
    If x_make = "Dacia" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Dacia" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Daewoo"""
    If x_make = "Daewoo" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Daewoo" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Daihatsu"""
    If x_make = "Daihatsu" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Daihatsu" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Dodge"""
    If x_make = "Dodge" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Dodge" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Eagle"""
    If x_make = "Eagle" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Eagle" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Ferrari"""
    If x_make = "Ferrari" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Ferrari" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Ford"""
    If x_make = "Ford" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Ford" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""GMC"""
    If x_make = "GMC" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "GMC" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Geo"""
    If x_make = "Geo" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Geo" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Honda"""
    If x_make = "Honda" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Honda" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Hummer"""
    If x_make = "Hummer" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Hummer" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Hyundai"""
    If x_make = "Hyundai" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Hyundai" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Infiniti"""
    If x_make = "Infiniti" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Infiniti" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""International"""
    If x_make = "International" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "International" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Isuzu"""
    If x_make = "Isuzu" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Isuzu" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Jaguar"""
    If x_make = "Jaguar" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Jaguar" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Jeep"""
    If x_make = "Jeep" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Jeep" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Kia"""
    If x_make = "Kia" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Kia" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Lamborghini"""
    If x_make = "Lamborghini" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Lamborghini" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Land Rover"""
    If x_make = "Land Rover" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Land Rover" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Lexus"""
    If x_make = "Lexus" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Lexus" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Lincoln"""
    If x_make = "Lincoln" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Lincoln" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Lotus"""
    If x_make = "Lotus" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Lotus" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Maserati"""
    If x_make = "Maserati" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Maserati" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Maybach"""
    If x_make = "Maybach" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Maybach" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Mazda"""
    If x_make = "Mazda" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Mazda" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Mercedes-Benz"""
    If x_make = "Mercedes-Benz" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Mercedes-Benz" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Mercury"""
    If x_make = "Mercury" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Mercury" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Mini"""
    If x_make = "Mini" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Mini" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Mitsubishi"""
    If x_make = "Mitsubishi" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Mitsubishi" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Morgan"""
    If x_make = "Morgan" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Morgan" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Nissan"""
    If x_make = "Nissan" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Nissan" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Oldsmobile"""
    If x_make = "Oldsmobile" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Oldsmobile" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Panoz"""
    If x_make = "Panoz" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Panoz" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Peugeot"""
    If x_make = "Peugeot" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Peugeot" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Plymouth"""
    If x_make = "Plymouth" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Plymouth" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Pontiac"""
    If x_make = "Pontiac" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Pontiac" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Porsche"""
    If x_make = "Porsche" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Porche" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Rolls-Royce"""
    If x_make = "Rolls-Royce" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Rolls-Royce" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Saab"""
    If x_make = "Saab" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Saab" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Saleen"""
    If x_make = "Saleen" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Saleen" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Saturn"""
    If x_make = "Saturn" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Saturn" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Scion"""
    If x_make = "Scion" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Scion" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Smart"""
    If x_make = "Smart" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Smart" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Sterling"""
    If x_make = "Sterling" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Sterling" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Subaru"""
    If x_make = "Subaru" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Subaru" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Suzuki"""
    If x_make = "Suzuki" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Suzuki" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Tesla"""
    If x_make = "Tesla" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Tesla" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Toyota"""
    If x_make = "Toyota" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Toyota" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Volkswagen"""
    If x_make = "Volkswagen" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Volkswagen" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Volvo"""
    If x_make = "Volvo" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Volvo" & "</option>"
    x_makeList = x_makeList & "<OPTION value=""Yugo"""
    If x_make = "Yugo" Then
        x_makeList = x_makeList & " selected"
    End If
    x_makeList = x_makeList & ">" & "Yugo" & "</option>"
x_makeList = x_makeList & "</select>"
response.write x_makeList
%></span></font></td>
</tr>
<tr>
<td bgcolor="whitesmoke" width="132">
                            <p align="right"><font color="#990000" face="Verdana"><span style="font-size:10pt;">*E-Mail:</span></font></td>
<td bgcolor="whitesmoke" width="191"><font face="Arial" size="2"><input type="text" name="x_email" size="15" maxlength=50 value="<%= Server.HtmlEncode(x_email&"") %>"></font></td>
<td bgcolor="#F5F5F5" width="117"><font face="Verdana"><span style="font-size:10pt;">Model:</span></font></td>
<td bgcolor="#F5F5F5" width="213"><font face="Verdana"><span style="font-size:10pt;"><input type="text" name="x_model" size="15" maxlength=50 value="<%= Server.HtmlEncode(x_model&"") %>"></span></font></td>
</tr>
<tr>
<td bgcolor="whitesmoke" width="132">
                            <p align="right"><font color="black" face="Verdana">&nbsp;</font></td>
<td bgcolor="whitesmoke" width="191">&nbsp;</td>
<td bgcolor="#F5F5F5" width="117"><font color="black" face="Verdana"><span style="font-size:10pt;">Bodystyle:</span></font></td>
<td bgcolor="#F5F5F5" width="213"><font face="Verdana"><span style="font-size:10pt;"><%
x_bodystyleList = "<SELECT name='x_bodystyle'><OPTION value=''>Please Select</OPTION>"
    x_bodystyleList = x_bodystyleList & "<OPTION value=""Cupe"""
    If x_bodystyle = "Cupe" Then
        x_bodystyleList = x_bodystyleList & " selected"
    End If
    x_bodystyleList = x_bodystyleList & ">" & "Cupe" & "</option>"
    x_bodystyleList = x_bodystyleList & "<OPTION value=""Convertible"""
    If x_bodystyle = "Convertible" Then
        x_bodystyleList = x_bodystyleList & " selected"
    End If
    x_bodystyleList = x_bodystyleList & ">" & "Convertible" & "</option>"
    x_bodystyleList = x_bodystyleList & "<OPTION value=""Sedan"""
    If x_bodystyle = "Sedan" Then
        x_bodystyleList = x_bodystyleList & " selected"
    End If
    x_bodystyleList = x_bodystyleList & ">" & "Sedan" & "</option>"
    x_bodystyleList = x_bodystyleList & "<OPTION value=""SUV"""
    If x_bodystyle = "SUV" Then
        x_bodystyleList = x_bodystyleList & " selected"
    End If
    x_bodystyleList = x_bodystyleList & ">" & "SUV" & "</option>"
    x_bodystyleList = x_bodystyleList & "<OPTION value=""Wagon"""
    If x_bodystyle = "Wagon" Then
        x_bodystyleList = x_bodystyleList & " selected"
    End If
    x_bodystyleList = x_bodystyleList & ">" & "Wagon" & "</option>"
x_bodystyleList = x_bodystyleList & "</select>"
response.write x_bodystyleList
%></span></font></td>
</tr>
<tr>
<td bgcolor="whitesmoke" width="327" colspan="2">
                            <p>&nbsp;</p>
</td>
<td bgcolor="#F5F5F5" width="117">
                            <p><font color="black" face="Verdana"><span style="font-size:10pt;">Transmission:</span></font></p>
</td>
<td bgcolor="#F5F5F5" width="213">
                            <font face="Verdana"><span style="font-size:10pt;"><%
x_transmissionList = "<SELECT name='x_transmission'><OPTION value=''>Please Select</OPTION>"
    x_transmissionList = x_transmissionList & "<OPTION value=""Automatic"""
    If x_transmission = "Automatic" Then
        x_transmissionList = x_transmissionList & " selected"
    End If
    x_transmissionList = x_transmissionList & ">" & "Automatic" & "</option>"
    x_transmissionList = x_transmissionList & "<OPTION value=""Manual"""
    If x_transmission = "Manual" Then
        x_transmissionList = x_transmissionList & " selected"
    End If
    x_transmissionList = x_transmissionList & ">" & "Manual" & "</option>"
x_transmissionList = x_transmissionList & "</select>"
response.write x_transmissionList
%>
                            </span></font></td>
</tr>
<tr>
<td bgcolor="whitesmoke" width="327" colspan="2">
                            <p><img src="images/searchperiod.gif" width="246" height="33" border="0"></p>
</td>
<td bgcolor="#F5F5F5" width="117">
                            <p><font face="Verdana"><span style="font-size:10pt;">Mileage:</span></font></p>
</td>
<td bgcolor="#F5F5F5" width="213"><font face="Verdana"><span style="font-size:10pt;"><input type="text" name="x_mileagelow" size="6" maxlength="7" value="<%= Server.HtmlEncode(x_mileagelow&"") %>"> 
                            to <input type="text" name="x_mileagehi" size="6" maxlength="6" value="<%= Server.HtmlEncode(x_mileagehi&"") %>"></span></font></td>
</tr>
<tr>
<td bgcolor="whitesmoke" width="327" colspan="2"><font color="black">&nbsp;</font></td>
<td bgcolor="#F5F5F5" width="117"><font face="Verdana"><span style="font-size:10pt;">Price:</span></font></td>
<td bgcolor="#F5F5F5" width="213"><font face="Verdana"><span style="font-size:10pt;"><input type="text" name="x_pricelow" size="6" maxlength="6" value="<%= Server.HtmlEncode(x_pricelow&"") %>"> 
                            to <input type="text" name="x_pricehi" size="6" maxlength="6" value="<%= Server.HtmlEncode(x_pricehi&"") %>"></span></font></td>
</tr>
<tr>
<td bgcolor="whitesmoke" width="327" colspan="2" height="32">
<font color="black" face="Verdana"><span style="font-size:10pt;"><input type="radio" name="x_time" value="1 week">1 
                            week </span></font></td>
<td bgcolor="#F5F5F5" width="334" colspan="2" height="32"><font face="Verdana"><span style="font-size:10pt;">Comments:</span></font></td>
</tr>
<tr>
<td bgcolor="whitesmoke" width="327" colspan="2" height="30"><font color="black" face="Verdana"><span style="font-size:10pt;"><input type="radio" name="x_time" value="2 weeks">2 
                            weeks </span></font></td>
<td bgcolor="#F5F5F5" width="334" colspan="2" height="123" rowspan="4"><font face="Verdana"><span style="font-size:10pt;"><textarea cols="35" rows="6" name="x_comments"><%= x_comments %></textarea></span></font>
                            <p>&nbsp;</td>
</tr>
<tr>
<td bgcolor="whitesmoke" width="327" colspan="2" height="30"><font color="black" face="Verdana"><span style="font-size:10pt;"><input type="radio" name="x_time" value="4 weeks">4 
                            weeks </span></font></td>
</tr>
<tr>
<td bgcolor="whitesmoke" width="327" colspan="2" height="30"><font color="black" face="Verdana"><span style="font-size:10pt;"><input type="radio" name="x_time" value="8 weeks">8 
                            weeks </span></font></td>
</tr>
<tr>
<td bgcolor="whitesmoke" width="327" colspan="2">&nbsp;</td>
</tr>
</table>
<p align="right">
<input type="submit" name="Action" value="   Submit   "> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</form>
            <p>&nbsp;</p>
        </td>
    </tr>
</table>
<!--#include file="footer.asp"-->
</p>
