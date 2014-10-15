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
'get action
a=Request.Form("a")
Select Case a
	Case "S": ' Get Search Criteria
	'Construct search criteria for advance search, remove blank field
	search_criteria = ""
	x_ID = Request.Form("x_ID")
	z_ID = Request.Form("z_ID")
	If x_ID <> "" Then
		srchFld = x_ID
		this_search_criteria = "x_ID=" & Server.URLEncode(srchFld)
		this_search_criteria = this_search_criteria & "&z_ID=" & Server.URLEncode(z_ID)
	Else
		this_search_criteria = ""
	End If
	If this_search_criteria <> "" Then
		If search_criteria = "" Then
			search_criteria = this_search_criteria
		Else
			search_criteria = search_criteria & "&" & this_search_criteria
		End If
	End If
	x_year = Request.Form("x_year")
	z_year = Request.Form("z_year")
	If x_year <> "" Then
		srchFld = x_year
		this_search_criteria = "x_year=" & Server.URLEncode(srchFld)
		this_search_criteria = this_search_criteria & "&z_year=" & Server.URLEncode(z_year)
	Else
		this_search_criteria = ""
	End If
	If this_search_criteria <> "" Then
		If search_criteria = "" Then
			search_criteria = this_search_criteria
		Else
			search_criteria = search_criteria & "&" & this_search_criteria
		End If
	End If
	x_make = Request.Form("x_make")
	z_make = Request.Form("z_make")
	If x_make <> "" Then
		srchFld = x_make
		this_search_criteria = "x_make=" & Server.URLEncode(srchFld)
		this_search_criteria = this_search_criteria & "&z_make=" & Server.URLEncode(z_make)
	Else
		this_search_criteria = ""
	End If
	If this_search_criteria <> "" Then
		If search_criteria = "" Then
			search_criteria = this_search_criteria
		Else
			search_criteria = search_criteria & "&" & this_search_criteria
		End If
	End If
	x_model = Request.Form("x_model")
	z_model = Request.Form("z_model")
	If x_model <> "" Then
		srchFld = x_model
		this_search_criteria = "x_model=" & Server.URLEncode(srchFld)
		this_search_criteria = this_search_criteria & "&z_model=" & Server.URLEncode(z_model)
	Else
		this_search_criteria = ""
	End If
	If this_search_criteria <> "" Then
		If search_criteria = "" Then
			search_criteria = this_search_criteria
		Else
			search_criteria = search_criteria & "&" & this_search_criteria
		End If
	End If
	x_type = Request.Form("x_type")
	z_type = Request.Form("z_type")
	If x_type <> "" Then
		srchFld = x_type
		this_search_criteria = "x_type=" & Server.URLEncode(srchFld)
		this_search_criteria = this_search_criteria & "&z_type=" & Server.URLEncode(z_type)
	Else
		this_search_criteria = ""
	End If
	If this_search_criteria <> "" Then
		If search_criteria = "" Then
			search_criteria = this_search_criteria
		Else
			search_criteria = search_criteria & "&" & this_search_criteria
		End If
	End If
	x_miles = Request.Form("x_miles")
	z_miles = Request.Form("z_miles")
	If x_miles <> "" Then
		srchFld = x_miles
		this_search_criteria = "x_miles=" & Server.URLEncode(srchFld)
		this_search_criteria = this_search_criteria & "&z_miles=" & Server.URLEncode(z_miles)
	Else
		this_search_criteria = ""
	End If
	If this_search_criteria <> "" Then
		If search_criteria = "" Then
			search_criteria = this_search_criteria
		Else
			search_criteria = search_criteria & "&" & this_search_criteria
		End If
	End If
	x_price = Request.Form("x_price")
	z_price = Request.Form("z_price")
	If x_price <> "" Then
		srchFld = x_price
		this_search_criteria = "x_price=" & Server.URLEncode(srchFld)
		this_search_criteria = this_search_criteria & "&z_price=" & Server.URLEncode(z_price)
	Else
		this_search_criteria = ""
	End If
	If this_search_criteria <> "" Then
		If search_criteria = "" Then
			search_criteria = this_search_criteria
		Else
			search_criteria = search_criteria & "&" & this_search_criteria
		End If
	End If
	x_doors = Request.Form("x_doors")
	z_doors = Request.Form("z_doors")
	If x_doors <> "" Then
		srchFld = x_doors
		this_search_criteria = "x_doors=" & Server.URLEncode(srchFld)
		this_search_criteria = this_search_criteria & "&z_doors=" & Server.URLEncode(z_doors)
	Else
		this_search_criteria = ""
	End If
	If this_search_criteria <> "" Then
		If search_criteria = "" Then
			search_criteria = this_search_criteria
		Else
			search_criteria = search_criteria & "&" & this_search_criteria
		End If
	End If
	x_engine = Request.Form("x_engine")
	z_engine = Request.Form("z_engine")
	If x_engine <> "" Then
		srchFld = x_engine
		this_search_criteria = "x_engine=" & Server.URLEncode(srchFld)
		this_search_criteria = this_search_criteria & "&z_engine=" & Server.URLEncode(z_engine)
	Else
		this_search_criteria = ""
	End If
	If this_search_criteria <> "" Then
		If search_criteria = "" Then
			search_criteria = this_search_criteria
		Else
			search_criteria = search_criteria & "&" & this_search_criteria
		End If
	End If
	x_transmission = Request.Form("x_transmission")
	z_transmission = Request.Form("z_transmission")
	If x_transmission <> "" Then
		srchFld = x_transmission
		this_search_criteria = "x_transmission=" & Server.URLEncode(srchFld)
		this_search_criteria = this_search_criteria & "&z_transmission=" & Server.URLEncode(z_transmission)
	Else
		this_search_criteria = ""
	End If
	If this_search_criteria <> "" Then
		If search_criteria = "" Then
			search_criteria = this_search_criteria
		Else
			search_criteria = search_criteria & "&" & this_search_criteria
		End If
	End If
	x_drivetrain = Request.Form("x_drivetrain")
	z_drivetrain = Request.Form("z_drivetrain")
	If x_drivetrain <> "" Then
		srchFld = x_drivetrain
		this_search_criteria = "x_drivetrain=" & Server.URLEncode(srchFld)
		this_search_criteria = this_search_criteria & "&z_drivetrain=" & Server.URLEncode(z_drivetrain)
	Else
		this_search_criteria = ""
	End If
	If this_search_criteria <> "" Then
		If search_criteria = "" Then
			search_criteria = this_search_criteria
		Else
			search_criteria = search_criteria & "&" & this_search_criteria
		End If
	End If
	x_ext_color = Request.Form("x_ext_color")
	z_ext_color = Request.Form("z_ext_color")
	If x_ext_color <> "" Then
		srchFld = x_ext_color
		this_search_criteria = "x_ext_color=" & Server.URLEncode(srchFld)
		this_search_criteria = this_search_criteria & "&z_ext_color=" & Server.URLEncode(z_ext_color)
	Else
		this_search_criteria = ""
	End If
	If this_search_criteria <> "" Then
		If search_criteria = "" Then
			search_criteria = this_search_criteria
		Else
			search_criteria = search_criteria & "&" & this_search_criteria
		End If
	End If
	x_int_color = Request.Form("x_int_color")
	z_int_color = Request.Form("z_int_color")
	If x_int_color <> "" Then
		srchFld = x_int_color
		this_search_criteria = "x_int_color=" & Server.URLEncode(srchFld)
		this_search_criteria = this_search_criteria & "&z_int_color=" & Server.URLEncode(z_int_color)
	Else
		this_search_criteria = ""
	End If
	If this_search_criteria <> "" Then
		If search_criteria = "" Then
			search_criteria = this_search_criteria
		Else
			search_criteria = search_criteria & "&" & this_search_criteria
		End If
	End If
	x_stock = Request.Form("x_stock")
	z_stock = Request.Form("z_stock")
	If x_stock <> "" Then
		srchFld = x_stock
		this_search_criteria = "x_stock=" & Server.URLEncode(srchFld)
		this_search_criteria = this_search_criteria & "&z_stock=" & Server.URLEncode(z_stock)
	Else
		this_search_criteria = ""
	End If
	If this_search_criteria <> "" Then
		If search_criteria = "" Then
			search_criteria = this_search_criteria
		Else
			search_criteria = search_criteria & "&" & this_search_criteria
		End If
	End If
	x_vin = Request.Form("x_vin")
	z_vin = Request.Form("z_vin")
	If x_vin <> "" Then
		srchFld = x_vin
		this_search_criteria = "x_vin=" & Server.URLEncode(srchFld)
		this_search_criteria = this_search_criteria & "&z_vin=" & Server.URLEncode(z_vin)
	Else
		this_search_criteria = ""
	End If
	If this_search_criteria <> "" Then
		If search_criteria = "" Then
			search_criteria = this_search_criteria
		Else
			search_criteria = search_criteria & "&" & this_search_criteria
		End If
	End If
	x_city_mpg = Request.Form("x_city_mpg")
	z_city_mpg = Request.Form("z_city_mpg")
	If x_city_mpg <> "" Then
		srchFld = x_city_mpg
		this_search_criteria = "x_city_mpg=" & Server.URLEncode(srchFld)
		this_search_criteria = this_search_criteria & "&z_city_mpg=" & Server.URLEncode(z_city_mpg)
	Else
		this_search_criteria = ""
	End If
	If this_search_criteria <> "" Then
		If search_criteria = "" Then
			search_criteria = this_search_criteria
		Else
			search_criteria = search_criteria & "&" & this_search_criteria
		End If
	End If
	x_hwy_mpg = Request.Form("x_hwy_mpg")
	z_hwy_mpg = Request.Form("z_hwy_mpg")
	If x_hwy_mpg <> "" Then
		srchFld = x_hwy_mpg
		this_search_criteria = "x_hwy_mpg=" & Server.URLEncode(srchFld)
		this_search_criteria = this_search_criteria & "&z_hwy_mpg=" & Server.URLEncode(z_hwy_mpg)
	Else
		this_search_criteria = ""
	End If
	If this_search_criteria <> "" Then
		If search_criteria = "" Then
			search_criteria = this_search_criteria
		Else
			search_criteria = search_criteria & "&" & this_search_criteria
		End If
	End If
	x_carfax = Request.Form("x_carfax")
	z_carfax = Request.Form("z_carfax")
	If x_carfax <> "" Then
		srchFld = x_carfax
		this_search_criteria = "x_carfax=" & Server.URLEncode(srchFld)
		this_search_criteria = this_search_criteria & "&z_carfax=" & Server.URLEncode(z_carfax)
	Else
		this_search_criteria = ""
	End If
	If this_search_criteria <> "" Then
		If search_criteria = "" Then
			search_criteria = this_search_criteria
		Else
			search_criteria = search_criteria & "&" & this_search_criteria
		End If
	End If
	x_special = Request.Form("x_special")
	z_special = Request.Form("z_special")
	If x_special <> "" Then
		srchFld = x_special
		this_search_criteria = "x_special=" & Server.URLEncode(srchFld)
		this_search_criteria = this_search_criteria & "&z_special=" & Server.URLEncode(z_special)
	Else
		this_search_criteria = ""
	End If
	If this_search_criteria <> "" Then
		If search_criteria = "" Then
			search_criteria = this_search_criteria
		Else
			search_criteria = search_criteria & "&" & this_search_criteria
		End If
	End If
	x_status = Request.Form("x_status")
	z_status = Request.Form("z_status")
	If x_status <> "" Then
		srchFld = x_status
		this_search_criteria = "x_status=" & Server.URLEncode(srchFld)
		this_search_criteria = this_search_criteria & "&z_status=" & Server.URLEncode(z_status)
	Else
		this_search_criteria = ""
	End If
	If this_search_criteria <> "" Then
		If search_criteria = "" Then
			search_criteria = this_search_criteria
		Else
			search_criteria = search_criteria & "&" & this_search_criteria
		End If
	End If
		If search_criteria <> "" Then
			Response.Clear
			Response.Redirect "carlist.asp" & "?" & search_criteria
		End If
End Select
' Open Connection to the database
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str
%>
<!--#include file="header.asp"-->
<meta name="generator" content="Namo WebEditor v5.0(Trial)">
<p><font face="Arial" size="2">Search TABLE: car<br><br><a href="carlist.asp">Back to List</a></font></p>
<script language="JavaScript" src="ew.js"></script>
<script language="JavaScript">
<!-- start Javascript
function  EW_checkMyForm(EW_this) {
if (EW_this.x_miles && !EW_checkinteger(EW_this.x_miles.value)) {
        if (!EW_onError(EW_this, EW_this.x_miles, "TEXT", "Incorrect integer - Miles"))
            return false; 
        }
if (EW_this.x_price && !EW_checknumber(EW_this.x_price.value)) {
        if (!EW_onError(EW_this, EW_this.x_price, "TEXT", "Incorrect floating point number - Price"))
            return false; 
        }
return true;
}
// end JavaScript -->
</script>
<form onSubmit="return EW_checkMyForm(this);"  action="carsrch.asp" method="post">
<p>
<input type="hidden" name="a" value="S">
<table border="0" cellspacing="0" cellpadding="0" bgcolor="#CCCCCC">
<tr>
<td bgcolor="#708090"><font color="#FFFFFF"><font face="Arial" size="2">Year</font></font>&nbsp;</td>
<td bgcolor="#F5F5F5"><font face="Arial" size="2">LIKE
<input type="hidden" name="z_year" value="LIKE,'%,%'"></font>&nbsp;</td>
<td bgcolor="#F5F5F5"><font face="Arial" size="2"><%
x_yearList = "<SELECT name='x_year'><OPTION value=''>Please Select</OPTION>"
    x_yearList = x_yearList & "<OPTION value=""2010"""
    If x_year = "2010" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "2010" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""2009"""
    If x_year = "2009" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "2009" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""2008"""
    If x_year = "2008" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "2008" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""2007"""
    If x_year = "2007" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "2007" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""2006"""
    If x_year = "2006" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "2006" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""2005"""
    If x_year = "2005" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "2005" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""2004"""
    If x_year = "2004" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "2004" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""2003"""
    If x_year = "2003" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "2003" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""2002"""
    If x_year = "2002" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "2002" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""2001"""
    If x_year = "2001" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "2001" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""2000"""
    If x_year = "2000" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "2000" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""1999"""
    If x_year = "1999" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "1999" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""1998"""
    If x_year = "1998" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "1998" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""1997"""
    If x_year = "1997" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "1997" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""1996"""
    If x_year = "1996" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "1996" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""1995"""
    If x_year = "1995" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "1995" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""1994"""
    If x_year = "1994" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "1994" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""1993"""
    If x_year = "1993" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "1993" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""1992"""
    If x_year = "1992" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "1992" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""1991"""
    If x_year = "1991" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "1991" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""1990"""
    If x_year = "1990" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "1990" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""1989"""
    If x_year = "1989" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "1989" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""1988"""
    If x_year = "1988" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "1988" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""1887"""
    If x_year = "1887" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "1987" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""1886"""
    If x_year = "1886" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "1986" & "</option>"
    x_yearList = x_yearList & "<OPTION value=""1885"""
    If x_year = "1885" Then
        x_yearList = x_yearList & " selected"
    End If
    x_yearList = x_yearList & ">" & "1985" & "</option>"
x_yearList = x_yearList & "</select>"
response.write x_yearList
%>
</font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#708090"><font color="#FFFFFF"><font face="Arial" size="2">Make</font></font>&nbsp;</td>
<td bgcolor="#F5F5F5"><font face="Arial" size="2">LIKE
<input type="hidden" name="z_make" value="LIKE,'%,%'"></font>&nbsp;</td>
<td bgcolor="#F5F5F5"><font face="Arial" size="2"><%

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
%>
</font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#708090"><font color="#FFFFFF"><font face="Arial" size="2">Model</font></font>&nbsp;</td>
<td bgcolor="#F5F5F5"><font face="Arial" size="2">LIKE
<input type="hidden" name="z_model" value="LIKE,'%,%'"></font>&nbsp;</td>
<td bgcolor="#F5F5F5"><font face="Arial" size="2"><input type="text" name="x_model" size=30 maxlength=50 value="<%= Server.HtmlEncode(x_model&"") %>"></font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#708090"><font color="#FFFFFF"><font face="Arial" size="2">Type</font></font>&nbsp;</td>
<td bgcolor="#F5F5F5"><font face="Arial" size="2">LIKE
<input type="hidden" name="z_type" value="LIKE,'%,%'"></font>&nbsp;</td>
<td bgcolor="#F5F5F5"><font face="Arial" size="2"><%
x_typeList = "<SELECT name='x_type'><OPTION value=''>Please Select</OPTION>"
    x_typeList = x_typeList & "<OPTION value=""Sedan"""
    If x_type = "Sedan" Then
        x_typeList = x_typeList & " selected"
    End If
    x_typeList = x_typeList & ">" & "Sedan" & "</option>"
    x_typeList = x_typeList & "<OPTION value=""SUV"""
    If x_type = "SUV" Then
        x_typeList = x_typeList & " selected"
    End If
    x_typeList = x_typeList & ">" & "SUV" & "</option>"
    x_typeList = x_typeList & "<OPTION value=""Mini-Van"""
    If x_type = "Mini-Van" Then
        x_typeList = x_typeList & " selected"
    End If
    x_typeList = x_typeList & ">" & "Mini-Van" & "</option>"
    x_typeList = x_typeList & "<OPTION value=""Wagon"""
    If x_type = "Wagon" Then
        x_typeList = x_typeList & " selected"
    End If
    x_typeList = x_typeList & ">" & "Wagon" & "</option>"
    x_typeList = x_typeList & "<OPTION value=""Hatchback"""
    If x_type = "Hatchback" Then
        x_typeList = x_typeList & " selected"
    End If
    x_typeList = x_typeList & ">" & "Hatchback" & "</option>"
    x_typeList = x_typeList & "<OPTION value=""Coupe"""
    If x_type = "Coupe" Then
        x_typeList = x_typeList & " selected"
    End If
    x_typeList = x_typeList & ">" & "Coupe" & "</option>"
    x_typeList = x_typeList & "<OPTION value=""Truck"""
    If x_type = "Truck" Then
        x_typeList = x_typeList & " selected"
    End If
    x_typeList = x_typeList & ">" & "Truck" & "</option>"
    x_typeList = x_typeList & "<OPTION value=""Convertible"""
    If x_type = "Convertible" Then
        x_typeList = x_typeList & " selected"
    End If
    x_typeList = x_typeList & ">" & "Convertible" & "</option>"
    x_typeList = x_typeList & "<OPTION value=""Sport"""
    If x_type = "Sport" Then
        x_typeList = x_typeList & " selected"
    End If
    x_typeList = x_typeList & ">" & "Sport" & "</option>"
    x_typeList = x_typeList & "<OPTION value=""SUT"""
    If x_type = "SUT" Then
        x_typeList = x_typeList & " selected"
    End If
    x_typeList = x_typeList & ">" & "SUT" & "</option>"
x_typeList = x_typeList & "</select>"
response.write x_typeList
%>
</font>&nbsp;</td>
</tr>
<tr>
<td bgcolor="#708090"><font color="#FFFFFF"><font face="Arial" size="2">Stock #</font></font>&nbsp;</td>
<td bgcolor="#F5F5F5"><font face="Arial" size="2">LIKE
<input type="hidden" name="z_stock" value="LIKE,'%,%'"></font>&nbsp;</td>
<td bgcolor="#F5F5F5"><font face="Arial" size="2"><input type="text" name="x_stock" size=30 maxlength=50 value="<%= Server.HtmlEncode(x_stock&"") %>"></font>&nbsp;</td>
</tr>
</table>
<p>
<input type="submit" name="Action" value="Search">
</form>
<!--#include file="footer.asp"-->
<%
conn.close
Set conn = nothing
%>
