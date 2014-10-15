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

function na_preload_img()
{ 
  var img_list = na_preload_img.arguments;
  if (document.preloadlist == null) 
    document.preloadlist = new Array();
  var top = document.preloadlist.length;
  for (var i=0; i < img_list.length; i++) {
    document.preloadlist[top+i] = new Image;
    document.preloadlist[top+i].src = img_list[i+1];
  } 
}

function na_change_img_src(name, nsdoc, rpath, preload)
{ 
  var img = eval((navigator.appName.indexOf('Netscape', 0) != -1) ? nsdoc+'.'+name : 'document.all.'+name);
  if (name == '')
    return;
  if (img) {
    img.altsrc = img.src;
    img.src    = rpath;
  } 
}

function na_restore_img_src(name, nsdoc)
{
  var img = eval((navigator.appName.indexOf('Netscape', 0) != -1) ? nsdoc+'.'+name : 'document.all.'+name);
  if (name == '')
    return;
  if (img && img.altsrc) {
    img.src    = img.altsrc;
    img.altsrc = null;
  } 
}

// -->
</script>
<meta name="generator" content="Namo WebEditor v5.0(Trial)">
<body OnLoad="na_preload_img(false, 'images/ofover.gif', 'images/printvehicleover.gif', 'images/testdriveover.gif', 'images/financeover.gif', 'images/tradeinover.gif', 'images/carfinde.gif', 'images/carfaxover.gif', 'images/carfaxfreeover.gif');">

<!--#include file="header.asp"-->

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
<table align="center" width="802" bgcolor="white" cellspacing="0" bordercolordark="white" bordercolorlight="black" cellpadding="0">
    <tr>
        <td width="28">
            <p>&nbsp;</p>
        </td>
        <td width="377">
<p><font face="Arial"><a href="carlist.asp"><span style="font-size:12pt;"><b><img src="images/back.gif" width="16" height="16" border="0" align="texttop"></b></span></a><span style="font-size:12pt;"><b> 
            </b></span></font><a href="carlist.asp"><font face="Arial" color="black"><span style="font-size:10pt;"><b>Back To Inventory</b></span></font></a></p>
        </td>
        <td width="11">
            <font face="Arial"><span style="font-size:12pt;"><b>&nbsp;</b></span></font>        </td>
        <td width="386">
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
%> &nbsp;<%Select Case x_make
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
%><br>&nbsp;</b></span></font>        </td>
    </tr>
    <tr>
        <td width="28">
            <p>&nbsp;</p>
        </td>
        <td width="377" valign="top">
            <table align="center" border="1" cellspacing="0" bordercolordark="white" bordercolorlight="#666666" bordercolor="#666666" width="326">
                <tr>
                    <td width="320">                <font face="Arial"><span style="font-size:10pt;">Make us an offer! We will try our best to accommodate you. Your offer should 
generally be greater than the invoice price. </span></font><font face="Arial" size="2"><br></font>
                        <p><script language="JavaScript" src="ew.js"></script>
<script language="JavaScript">
<!-- start Javascript
function  EW_checkMyForm(EW_this) {
return true;
}
// end JavaScript -->
</script>
<form onSubmit="return EW_checkMyForm(this);"  action="makeofferadd.asp" method="post">
<p align="center">
<input type="hidden" name="a" value="A">
                            <b><font face="Arial"><span style="font-size:14pt;">Make An Offer 
                            &nbsp;&nbsp;&nbsp;</span></font></b><SPAN class=required style="font-size:8pt;"><font face="Arial" color="#990000">* indicates required fields.</font></SPAN><table border="0" cellspacing="0" cellpadding="0" bgcolor="whitesmoke" width="290" bordercolordark="#CCCCCC" bordercolorlight="#CCCCCC" align="center">
<tr>
<td width="129" bordercolordark="#CCCCCC" bordercolorlight="#CCCCCC" bordercolor="white">&nbsp;</td>
<td width="161" bordercolor="white">&nbsp;</td>
</tr>
<tr>
<td width="129" bordercolor="white"><font color="#990000" face="Arial"><span style="font-size:12pt;">* 
                                        First Name:</span></font></td>
<td width="161" bordercolor="white"><font face="Arial" size="2"><input type="text" name="x_first" size="20" maxlength=50 value="<%= Server.HtmlEncode(x_first&"") %>"></font>&nbsp;</td>
</tr>
<tr>
<td width="129" bordercolor="white"><font color="#990000" face="Arial"><span style="font-size:12pt;">* 
                                        Last Name:</span></font></td>
<td width="161" bordercolor="white"><font face="Arial" size="2"><input type="text" name="x_last" size="20" maxlength=50 value="<%= Server.HtmlEncode(x_last&"") %>"></font>&nbsp;</td>
</tr>
<tr>
<td width="129" bordercolor="white"><font color="black" face="Arial"><span style="font-size:12pt;">Phone 
                                        Number:</span></font></td>
<td width="161" bordercolor="white"><font face="Arial" size="2"><input type="text" name="x_phone" size="20" maxlength=50 value="<%= Server.HtmlEncode(x_phone&"") %>"></font>&nbsp;</td>
</tr>
<tr>
<td width="129" bordercolor="white"><font color="#990000" face="Arial"><span style="font-size:12pt;">* 
                                        E-mail Address:</span></font></td>
<td width="161" bordercolor="white"><font face="Arial" size="2"><input type="text" name="x_email" size="20" maxlength=50 value="<%= Server.HtmlEncode(x_email&"") %>"></font>&nbsp;</td>
</tr>
<tr>
<td width="129" bordercolor="white">
                                        <p>&nbsp;</p>
</td>
<td width="161" bordercolor="white">
                                        <p>&nbsp;</p>
</td>
</tr>
<tr>
<td width="129" bordercolor="white">
                                        <p><span style="font-size:12pt;"><font face="Arial" color="black">Asking 
                                        Price:</font></span></p>
</td>
<td width="161" bordercolor="white">
                        <p align="left"><font face="Arial"><span style="font-size:12pt;"><% if isnumeric(x_price) then response.write formatcurrency(x_price,0,-2,-2,-2) else response.write x_price end if %></span></font></p>
</td>
</tr>
<tr>
<td width="129" bordercolor="white"><font color="#990000" face="Arial"><span style="font-size:12pt;">* 
                                        Your Offer&nbsp;</span></font></td>
<td width="161" bordercolor="white"><font face="Arial" size="2"><input type="text" name="x_offer" size="10" maxlength=50 value="<%= Server.HtmlEncode(x_offer&"") %>"></font>&nbsp;</td>
</tr>
<tr>
<td width="129" bordercolor="white"><font color="#990000">&nbsp;</font></td>
<td width="161" bordercolor="white">
<p><font face="Arial" size="2"><input type="hidden" name="x_stock" value="<%= x_stock %>"><input type="hidden" name="x_make" value="<%= x_make %>"><input type="hidden" name="x_model" value="<%= x_model %>"><input type="hidden" name="x_year" value="<%= x_year %>"><input type="hidden" name="x_price" value="<%= x_price %>"><input type="hidden" name="x_ID" value="<%= x_ID %>"><input type="hidden" name="x_offerdate" value="<%= x_offerdate %>"></font></td>
</tr>
        <tr>
<td width="129" bordercolor="white">
                <p>&nbsp;</p>
</td>
<td width="161" bordercolor="white">
<p align="right">
<input type="submit" name="Action" value="   Submit   "> &nbsp;&nbsp;</td>
        </tr>
                                <tr>
<td width="129" bordercolor="white">
                                        <p>&nbsp;</p>
</td>
<td width="161" bordercolor="white">
                                        <p>&nbsp;</p>
</td>
                                </tr>
</table>
</form>
                    </td>
                </tr>
            </table>
        </td>
        <td width="11">
            <p>&nbsp;</p>
        </td>
        <td width="386" valign="top">
            <table align="center" cellpadding="0" cellspacing="0" bgcolor="white" bordercolordark="white" bordercolorlight="black" width="400">
                <tr>
                    <td width="400" colspan="2">
                        <p><font face="Arial" color="#990000"><span style="font-size:10pt;"><b>Call 
                        - (773) 334-0025<br></b></span></font></p>
                    </td>
                </tr>
                <tr>
                    <td width="400" colspan="2"><font face="Verdana"><b><span style="font-size:8pt;"><br>Vehicle Location: Rocar Auto Sales<BR>5136 N. Western Ave. Chicago, IL 
60625<br>&nbsp;</span></b></font></td>
                </tr>
                <tr>
                    <td width="198">
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
                    <td width="202">
                        <p align="center"><font face="Arial" color="#990000"><span style="font-size:14pt;">Price: 
                        <% if isnumeric(x_price) then response.write formatcurrency(x_price,0,-2,-2,-2) else response.write x_price end if %></span></font></p>
                    </td>
                </tr>
                <tr>
                    <td width="198">
                        <p><font face="Verdana"><span style="font-size:8pt;"><b>Engine: 
                        </b><% response.write x_engine %></span></font></p>
                    </td>
                    <td width="202" rowspan="2">
                        <p align="center"><a href="offer.asp?st=<%=x_stock%>&mk=<%=x_make%>&md=<%=x_model%>&yr=<%=x_year%>&pr=<%=x_price%>&key=<%=key%>" OnMouseOut="na_restore_img_src('offer2', 'document')" OnMouseOver="na_change_img_src('offer2', 'document', 'images/ofover.gif', true);"><img src="images/of.gif" width="161" height="37" border="0" align="absmiddle" name="offer2"></a></p>
                    </td>
                </tr>
                <tr>
                    <td width="198"><font color="black" face="Verdana"><span style="font-size:8pt;"><b>Transmission: 
                        </b></span></font><font face="Verdana"><span style="font-size:8pt;"><%
Select Case x_transmission
    Case "Automatic" response.write "Automatic"
    Case "Manual" response.write "Manual"
End Select
%></span></font></td>
                </tr>
                <tr>
                    <td width="198">
                        <p><font face="Verdana"><span style="font-size:8pt;"><b>Ext. 
                        Color: </b></span></font><span style="font-size:8pt;"><font face="Verdana"><% response.write x_ext_color %></font></span></p>
                    </td>
                    <td width="202">
                        <p>&nbsp;</p>
                    </td>
                </tr>
                <tr>
                    <td width="198">
                        <p><b><span style="font-size:8pt;"><font face="Verdana">Int. 
                        Color: </font></span></b><span style="font-size:8pt;"><font face="Verdana"><% response.write x_int_color %></font></span></p>
                    </td>
                    <td width="202">
                        <p>&nbsp;</p>
                    </td>
                </tr>
                <tr>
                    <td width="198">
                        <p><b><span style="font-size:8pt;"><font face="Verdana">Mileage: 
                        </font></span></b><span style="font-size:8pt;"><font face="Verdana"><% if isnumeric(x_miles) then response.write formatnumber(x_miles,0,-2,-2,-2) else response.write x_miles end if %><br>&nbsp;</font></span></p>
                    </td>
                    <td width="202">
                        <p>&nbsp;</p>
                    </td>
                </tr>
                <tr>
                    <td width="198">
                        <p><b><span style="font-size:8pt;"><font face="Verdana">Stock 
                        Number: </font></span></b><span style="font-size:8pt;"><font face="Verdana"><% response.write x_stock %></font></span></p>
                    </td>
                    <td width="202">
                        <p>&nbsp;</p>
                    </td>
                </tr>
                <tr>
                    <td width="400" colspan="2">
                        <p><b><span style="font-size:8pt;"><font face="Verdana">VIN 
                        Number: </font></span></b><span style="font-size:8pt;"><font face="Verdana"><% response.write x_vin %><br>&nbsp;</font></span></p>
                    </td>
                </tr>
                <tr>
                    <td width="198">
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
                                    <p align="center"><% If not isnull(x_city_mpg) Then %><b><span style="font-size:16pt;"><font face="Verdana"><% response.write x_city_mpg %></font></span></b><%ELSE%><b><span style="font-size:16pt;"><font face="Verdana"><% response.write "--"%><%end if%></font></span></b></td>
                                <td width="65">
                                    <p align="left"><% If not isnull(x_hwy_mpg) Then %><b><span style="font-size:16pt;"><font face="Verdana"><% response.write x_hwy_mpg %></font></span></b><%ELSE%><b><span style="font-size:16pt;"><font face="Verdana"><% response.write "--"%><%end if%></font></span></b></td>
                            </tr>
                            <tr>
                                <td width="176" colspan="3"><font face="Arial"><span style="font-size:7pt;">Actual rating will vary with options, driving conditions, habits and vehicle 
condition.<br>&nbsp;</span></font></td>
                            </tr>
                        </table>
                    </td>
                    <td width="202">
                        <p>&nbsp;</p>
                    </td>
                </tr>
                <tr>
                    <td width="400" colspan="2">
                        <table align="center" border="0" cellspacing="0" bordercolordark="#071734" bordercolorlight="#071734" width="347" bordercolor="#071734">
                            <tr>
                                <td width="307" bordercolor="#CCCCCC" bgcolor="#CCCCCC">
                                    <p><a href="javascript:na_open_window('win', 'carviewprint.asp?export=html&key=<%=key%>', 0, 0, 600, 500, 0, 0, 0, 1, 1)" OnMouseOut="na_restore_img_src('printvehicle1', 'document')" OnMouseOver="na_change_img_src('printvehicle1', 'document', 'images/printvehicleover.gif', true)"><img src="images/printvehicle.gif" width="312" height="30" border="0" name="printvehicle1"></a></p>
                                </td>
                                <td width="30" rowspan="6" bgcolor="#071734" height="214">
                                    <p align="center"><img src="images/tools.gif" width="13" height="95" border="0"></p>
                                </td>
                            </tr>
                            <tr>
                                <td width="307" bordercolor="#CCCCCC" bgcolor="#CCCCCC">
                                    <p><a OnMouseOut="na_restore_img_src('testdrive1', 'document')" OnMouseOver="na_change_img_src('testdrive1', 'document', 'images/testdriveover.gif', true);" href="moreinfo.asp?st=<%=x_stock%>&mk=<%=x_make%>&md=<%=x_model%>&yr=<%=x_year%>&pr=<%=x_price%>&key=<%=key%>"><img src="images/testdrive.gif" width="312" height="30" border="0" name="testdrive1"></a></p>
                                </td>
                            </tr>
                            <tr>
                                <td width="307" bordercolor="#CCCCCC" bgcolor="#CCCCCC">
                                    <p><a OnMouseOut="na_restore_img_src('finance1', 'document')" OnMouseOver="na_change_img_src('finance1', 'document', 'images/financeover.gif', true);" href="creditadd.asp?st=<%=x_stock%>&k=<%=key%>&v=<%=x_vin%>&mi=<%=x_miles%>&p=<%=x_price%>&mk=<%=x_make%>&m=<%=x_model%>&y=<%=x_year%>"><img src="images/finance.gif" width="312" height="30" border="0" name="finance1"></a></p>
                                </td>
                            </tr>
                            <tr>
                                <td width="307" bordercolor="#CCCCCC" bgcolor="#CCCCCC">
                                    <p><a OnMouseOut="na_restore_img_src('tradein1', 'document')" OnMouseOver="na_change_img_src('tradein1', 'document', 'images/tradeinover.gif', true);" href="tradeinadd.asp?st=<%=x_stock%>"><img src="images/tradein.gif" width="312" height="30" border="0" name="tradein1"></a></p>
                                </td>
                            </tr>
                            <tr>
                                <td width="307" bordercolor="#CCCCCC" bgcolor="#CCCCCC">
                                    <p><a href="findeadd.asp" OnMouseOut="na_restore_img_src('carfinde1', 'document')" OnMouseOver="na_change_img_src('carfinde1', 'document', 'images/carfinde.gif', true);"><img src="images/carfind.gif" width="312" height="30" border="0" name="carfinde1"></a></p>
                                </td>
                            </tr>
                            <tr>
                                <td width="307" bgcolor="#CCCCCC" height="67" bordercolor="white">
                                    <p align="center"><% If x_carfax = "No" Then %><a OnMouseOut="na_restore_img_src('carfax', 'document')" OnMouseOver="na_change_img_src('carfax', 'document', 'images/carfaxover.gif', true)" href="javascript:na_open_window('win', 'http://www.carfax.com/cfm/check_order.cfm?partner=CDM_H&VIN=<%=x_vin%>', 0, 0, 800, 600, 1, 1, 1, 1, 1)" target="_self"><img src="images/carfax.gif" width="125" height="51" border="0" name="carfax"></a>

<font face="Arial" size="2"><%Else%><a OnMouseOut="na_restore_img_src('freecarfax', 'document')" OnMouseOver="na_change_img_src('freecarfax', 'document', 'images/carfaxfreeover.gif', true)" href="javascript:na_open_window('win', 'http://www.cars.com/go/search/logCarfaxClick.jsp?linklocation=detail&VIN=<%=x_vin%>&pa_id=136993453&dlr_id=162277&CPO=N&aff=national', 0, 0, 800, 600, 1, 1, 1, 1, 1)" target="_self"><img src="images/carfaxfree.gif" width="173" height="61" border="0" name="freecarfax"></a><% End If %></font></td>
                            </tr>
                        </table>
                    </td>
                </tr>
            </table>
        </td>
    </tr>
    <tr>
        <td width="28">
            <p>&nbsp;</p>
        </td>
        <td width="377">
            <p>&nbsp;</p>
        </td>
        <td width="11">
            <p>&nbsp;</p>
        </td>
        <td width="386">
            <p>&nbsp;</p>
        </td>
    </tr>
    <tr>
        <td width="28">
            <p>&nbsp;</p>
        </td>
        <td width="377">
            <p>&nbsp;</p>
        </td>
        <td width="11">
            <p>&nbsp;</p>
        </td>
        <td width="386">
            <p>&nbsp;</p>
        </td>
    </tr>
</table>
<!--#include file="footer.asp"-->


</body>
</html>