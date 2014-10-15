
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
displayRecs = 20
recRange = 10
%>
<%
dbwhere = ""
masterdetailwhere = ""
searchwhere = ""
a_search = ""
b_search = ""
whereClause = ""
%>
<%
' Get search criteria for advance search
x_ID = Request.QueryString("x_ID")
z_ID = Request.QueryString("z_ID")
arrfieldopr = Split(z_ID,",")
If x_ID <> "" Then
	x_ID = Replace(x_ID,"'","''")
	x_ID = Replace(x_ID,"[","[[]")
	a_search = a_search & "[ID] " 'add field
	a_search = a_search	& arrfieldopr(0) & " " ' add operator
	If Ubound(arrfieldopr) >= 1 Then
		a_search = a_search & arrfieldopr(1) 'add search prefix
	End If
	a_search = a_search & x_ID 'add input parameter
	If Ubound(arrfieldopr) >=2 Then
		a_search = a_search & arrfieldopr(2) 'add search suffix
	End If
	a_search = a_search	 & " AND "
End If
x_year = Request.QueryString("x_year")
z_year = Request.QueryString("z_year")
arrfieldopr = Split(z_year,",")
If x_year <> "" Then
	x_year = Replace(x_year,"'","''")
	x_year = Replace(x_year,"[","[[]")
	a_search = a_search & "[year] " 'add field
	a_search = a_search	& arrfieldopr(0) & " " ' add operator
	If Ubound(arrfieldopr) >= 1 Then
		a_search = a_search & arrfieldopr(1) 'add search prefix
	End If
	a_search = a_search & x_year 'add input parameter
	If Ubound(arrfieldopr) >=2 Then
		a_search = a_search & arrfieldopr(2) 'add search suffix
	End If
	a_search = a_search	 & " AND "
End If
x_make = Request.QueryString("x_make")
z_make = Request.QueryString("z_make")
arrfieldopr = Split(z_make,",")
If x_make <> "" Then
	x_make = Replace(x_make,"'","''")
	x_make = Replace(x_make,"[","[[]")
	a_search = a_search & "[make] " 'add field
	a_search = a_search	& arrfieldopr(0) & " " ' add operator
	If Ubound(arrfieldopr) >= 1 Then
		a_search = a_search & arrfieldopr(1) 'add search prefix
	End If
	a_search = a_search & x_make 'add input parameter
	If Ubound(arrfieldopr) >=2 Then
		a_search = a_search & arrfieldopr(2) 'add search suffix
	End If
	a_search = a_search	 & " AND "
End If
x_model = Request.QueryString("x_model")
z_model = Request.QueryString("z_model")
arrfieldopr = Split(z_model,",")
If x_model <> "" Then
	x_model = Replace(x_model,"'","''")
	x_model = Replace(x_model,"[","[[]")
	a_search = a_search & "[model] " 'add field
	a_search = a_search	& arrfieldopr(0) & " " ' add operator
	If Ubound(arrfieldopr) >= 1 Then
		a_search = a_search & arrfieldopr(1) 'add search prefix
	End If
	a_search = a_search & x_model 'add input parameter
	If Ubound(arrfieldopr) >=2 Then
		a_search = a_search & arrfieldopr(2) 'add search suffix
	End If
	a_search = a_search	 & " AND "
End If
x_type = Request.QueryString("x_type")
z_type = Request.QueryString("z_type")
arrfieldopr = Split(z_type,",")
If x_type <> "" Then
	x_type = Replace(x_type,"'","''")
	x_type = Replace(x_type,"[","[[]")
	a_search = a_search & "[type] " 'add field
	a_search = a_search	& arrfieldopr(0) & " " ' add operator
	If Ubound(arrfieldopr) >= 1 Then
		a_search = a_search & arrfieldopr(1) 'add search prefix
	End If
	a_search = a_search & x_type 'add input parameter
	If Ubound(arrfieldopr) >=2 Then
		a_search = a_search & arrfieldopr(2) 'add search suffix
	End If
	a_search = a_search	 & " AND "
End If
x_miles = Request.QueryString("x_miles")
z_miles = Request.QueryString("z_miles")
arrfieldopr = Split(z_miles,",")
If x_miles <> "" Then
	x_miles = Replace(x_miles,"'","''")
	x_miles = Replace(x_miles,"[","[[]")
	a_search = a_search & "[miles] " 'add field
	a_search = a_search	& arrfieldopr(0) & " " ' add operator
	If Ubound(arrfieldopr) >= 1 Then
		a_search = a_search & arrfieldopr(1) 'add search prefix
	End If
	a_search = a_search & x_miles 'add input parameter
	If Ubound(arrfieldopr) >=2 Then
		a_search = a_search & arrfieldopr(2) 'add search suffix
	End If
	a_search = a_search	 & " AND "
End If
x_price = Request.QueryString("x_price")
z_price = Request.QueryString("z_price")
arrfieldopr = Split(z_price,",")
If x_price <> "" Then
	x_price = Replace(x_price,"'","''")
	x_price = Replace(x_price,"[","[[]")
	a_search = a_search & "[price] " 'add field
	a_search = a_search	& arrfieldopr(0) & " " ' add operator
	If Ubound(arrfieldopr) >= 1 Then
		a_search = a_search & arrfieldopr(1) 'add search prefix
	End If
	a_search = a_search & x_price 'add input parameter
	If Ubound(arrfieldopr) >=2 Then
		a_search = a_search & arrfieldopr(2) 'add search suffix
	End If
	a_search = a_search	 & " AND "
End If
x_doors = Request.QueryString("x_doors")
z_doors = Request.QueryString("z_doors")
arrfieldopr = Split(z_doors,",")
If x_doors <> "" Then
	x_doors = Replace(x_doors,"'","''")
	x_doors = Replace(x_doors,"[","[[]")
	a_search = a_search & "[doors] " 'add field
	a_search = a_search	& arrfieldopr(0) & " " ' add operator
	If Ubound(arrfieldopr) >= 1 Then
		a_search = a_search & arrfieldopr(1) 'add search prefix
	End If
	a_search = a_search & x_doors 'add input parameter
	If Ubound(arrfieldopr) >=2 Then
		a_search = a_search & arrfieldopr(2) 'add search suffix
	End If
	a_search = a_search	 & " AND "
End If
x_engine = Request.QueryString("x_engine")
z_engine = Request.QueryString("z_engine")
arrfieldopr = Split(z_engine,",")
If x_engine <> "" Then
	x_engine = Replace(x_engine,"'","''")
	x_engine = Replace(x_engine,"[","[[]")
	a_search = a_search & "[engine] " 'add field
	a_search = a_search	& arrfieldopr(0) & " " ' add operator
	If Ubound(arrfieldopr) >= 1 Then
		a_search = a_search & arrfieldopr(1) 'add search prefix
	End If
	a_search = a_search & x_engine 'add input parameter
	If Ubound(arrfieldopr) >=2 Then
		a_search = a_search & arrfieldopr(2) 'add search suffix
	End If
	a_search = a_search	 & " AND "
End If
x_transmission = Request.QueryString("x_transmission")
z_transmission = Request.QueryString("z_transmission")
arrfieldopr = Split(z_transmission,",")
If x_transmission <> "" Then
	x_transmission = Replace(x_transmission,"'","''")
	x_transmission = Replace(x_transmission,"[","[[]")
	a_search = a_search & "[transmission] " 'add field
	a_search = a_search	& arrfieldopr(0) & " " ' add operator
	If Ubound(arrfieldopr) >= 1 Then
		a_search = a_search & arrfieldopr(1) 'add search prefix
	End If
	a_search = a_search & x_transmission 'add input parameter
	If Ubound(arrfieldopr) >=2 Then
		a_search = a_search & arrfieldopr(2) 'add search suffix
	End If
	a_search = a_search	 & " AND "
End If
x_drivetrain = Request.QueryString("x_drivetrain")
z_drivetrain = Request.QueryString("z_drivetrain")
arrfieldopr = Split(z_drivetrain,",")
If x_drivetrain <> "" Then
	x_drivetrain = Replace(x_drivetrain,"'","''")
	x_drivetrain = Replace(x_drivetrain,"[","[[]")
	a_search = a_search & "[drivetrain] " 'add field
	a_search = a_search	& arrfieldopr(0) & " " ' add operator
	If Ubound(arrfieldopr) >= 1 Then
		a_search = a_search & arrfieldopr(1) 'add search prefix
	End If
	a_search = a_search & x_drivetrain 'add input parameter
	If Ubound(arrfieldopr) >=2 Then
		a_search = a_search & arrfieldopr(2) 'add search suffix
	End If
	a_search = a_search	 & " AND "
End If
x_ext_color = Request.QueryString("x_ext_color")
z_ext_color = Request.QueryString("z_ext_color")
arrfieldopr = Split(z_ext_color,",")
If x_ext_color <> "" Then
	x_ext_color = Replace(x_ext_color,"'","''")
	x_ext_color = Replace(x_ext_color,"[","[[]")
	a_search = a_search & "[ext_color] " 'add field
	a_search = a_search	& arrfieldopr(0) & " " ' add operator
	If Ubound(arrfieldopr) >= 1 Then
		a_search = a_search & arrfieldopr(1) 'add search prefix
	End If
	a_search = a_search & x_ext_color 'add input parameter
	If Ubound(arrfieldopr) >=2 Then
		a_search = a_search & arrfieldopr(2) 'add search suffix
	End If
	a_search = a_search	 & " AND "
End If
x_int_color = Request.QueryString("x_int_color")
z_int_color = Request.QueryString("z_int_color")
arrfieldopr = Split(z_int_color,",")
If x_int_color <> "" Then
	x_int_color = Replace(x_int_color,"'","''")
	x_int_color = Replace(x_int_color,"[","[[]")
	a_search = a_search & "[int_color] " 'add field
	a_search = a_search	& arrfieldopr(0) & " " ' add operator
	If Ubound(arrfieldopr) >= 1 Then
		a_search = a_search & arrfieldopr(1) 'add search prefix
	End If
	a_search = a_search & x_int_color 'add input parameter
	If Ubound(arrfieldopr) >=2 Then
		a_search = a_search & arrfieldopr(2) 'add search suffix
	End If
	a_search = a_search	 & " AND "
End If
x_stock = Request.QueryString("x_stock")
z_stock = Request.QueryString("z_stock")
arrfieldopr = Split(z_stock,",")
If x_stock <> "" Then
	x_stock = Replace(x_stock,"'","''")
	x_stock = Replace(x_stock,"[","[[]")
	a_search = a_search & "[stock] " 'add field
	a_search = a_search	& arrfieldopr(0) & " " ' add operator
	If Ubound(arrfieldopr) >= 1 Then
		a_search = a_search & arrfieldopr(1) 'add search prefix
	End If
	a_search = a_search & x_stock 'add input parameter
	If Ubound(arrfieldopr) >=2 Then
		a_search = a_search & arrfieldopr(2) 'add search suffix
	End If
	a_search = a_search	 & " AND "
End If
x_vin = Request.QueryString("x_vin")
z_vin = Request.QueryString("z_vin")
arrfieldopr = Split(z_vin,",")
If x_vin <> "" Then
	x_vin = Replace(x_vin,"'","''")
	x_vin = Replace(x_vin,"[","[[]")
	a_search = a_search & "[vin] " 'add field
	a_search = a_search	& arrfieldopr(0) & " " ' add operator
	If Ubound(arrfieldopr) >= 1 Then
		a_search = a_search & arrfieldopr(1) 'add search prefix
	End If
	a_search = a_search & x_vin 'add input parameter
	If Ubound(arrfieldopr) >=2 Then
		a_search = a_search & arrfieldopr(2) 'add search suffix
	End If
	a_search = a_search	 & " AND "
End If
x_city_mpg = Request.QueryString("x_city_mpg")
z_city_mpg = Request.QueryString("z_city_mpg")
arrfieldopr = Split(z_city_mpg,",")
If x_city_mpg <> "" Then
	x_city_mpg = Replace(x_city_mpg,"'","''")
	x_city_mpg = Replace(x_city_mpg,"[","[[]")
	a_search = a_search & "[city_mpg] " 'add field
	a_search = a_search	& arrfieldopr(0) & " " ' add operator
	If Ubound(arrfieldopr) >= 1 Then
		a_search = a_search & arrfieldopr(1) 'add search prefix
	End If
	a_search = a_search & x_city_mpg 'add input parameter
	If Ubound(arrfieldopr) >=2 Then
		a_search = a_search & arrfieldopr(2) 'add search suffix
	End If
	a_search = a_search	 & " AND "
End If
x_hwy_mpg = Request.QueryString("x_hwy_mpg")
z_hwy_mpg = Request.QueryString("z_hwy_mpg")
arrfieldopr = Split(z_hwy_mpg,",")
If x_hwy_mpg <> "" Then
	x_hwy_mpg = Replace(x_hwy_mpg,"'","''")
	x_hwy_mpg = Replace(x_hwy_mpg,"[","[[]")
	a_search = a_search & "[hwy_mpg] " 'add field
	a_search = a_search	& arrfieldopr(0) & " " ' add operator
	If Ubound(arrfieldopr) >= 1 Then
		a_search = a_search & arrfieldopr(1) 'add search prefix
	End If
	a_search = a_search & x_hwy_mpg 'add input parameter
	If Ubound(arrfieldopr) >=2 Then
		a_search = a_search & arrfieldopr(2) 'add search suffix
	End If
	a_search = a_search	 & " AND "
End If
x_carfax = Request.QueryString("x_carfax")
z_carfax = Request.QueryString("z_carfax")
arrfieldopr = Split(z_carfax,",")
If x_carfax <> "" Then
	x_carfax = Replace(x_carfax,"'","''")
	x_carfax = Replace(x_carfax,"[","[[]")
	a_search = a_search & "[carfax] " 'add field
	a_search = a_search	& arrfieldopr(0) & " " ' add operator
	If Ubound(arrfieldopr) >= 1 Then
		a_search = a_search & arrfieldopr(1) 'add search prefix
	End If
	a_search = a_search & x_carfax 'add input parameter
	If Ubound(arrfieldopr) >=2 Then
		a_search = a_search & arrfieldopr(2) 'add search suffix
	End If
	a_search = a_search	 & " AND "
End If
x_special = Request.QueryString("x_special")
z_special = Request.QueryString("z_special")
arrfieldopr = Split(z_special,",")
If x_special <> "" Then
	x_special = Replace(x_special,"'","''")
	x_special = Replace(x_special,"[","[[]")
	a_search = a_search & "[special] " 'add field
	a_search = a_search	& arrfieldopr(0) & " " ' add operator
	If Ubound(arrfieldopr) >= 1 Then
		a_search = a_search & arrfieldopr(1) 'add search prefix
	End If
	a_search = a_search & x_special 'add input parameter
	If Ubound(arrfieldopr) >=2 Then
		a_search = a_search & arrfieldopr(2) 'add search suffix
	End If
	a_search = a_search	 & " AND "
End If
x_status = Request.QueryString("x_status")
z_status = Request.QueryString("z_status")
arrfieldopr = Split(z_status,",")
If x_status <> "" Then
	x_status = Replace(x_status,"'","''")
	x_status = Replace(x_status,"[","[[]")
	a_search = a_search & "[status] " 'add field
	a_search = a_search	& arrfieldopr(0) & " " ' add operator
	If Ubound(arrfieldopr) >= 1 Then
		a_search = a_search & arrfieldopr(1) 'add search prefix
	End If
	a_search = a_search & x_status 'add input parameter
	If Ubound(arrfieldopr) >=2 Then
		a_search = a_search & arrfieldopr(2) 'add search suffix
	End If
	a_search = a_search	 & " AND "
End If
If Len(a_search) > 4 Then
	a_search = Mid(a_search,1,Len(a_search)-4)
End If
%>
<%
'Build search criteria
If a_search <> "" Then
	searchwhere = a_search 'advance search
ElseIf b_search <> "" Then
	searchwhere = b_search 'basic search
End If
'Save search criteria
If searchwhere <> "" Then
	Session("car_searchwhere") = searchwhere
	'reset start record counter (new search)
	startRec = 1
	Session("car_REC") = startRec
Else
	searchwhere = Session("car_searchwhere")
End If
%>
<%
'Get clear search cmd
If Request.QueryString("cmd").Count > 0 Then
	cmd = Request.QueryString("cmd")
	If UCase(cmd) = "RESET" Then
		'reset search criteria
		searchwhere = ""
		Session("car_searchwhere") =searchwhere
    ElseIf UCase(cmd) = "RESETALL" Then
		'reset search criteria
		searchwhere = ""
		Session("car_searchwhere") =searchwhere
	End If
	'reset start record counter (reset command)
	startRec = 1
	Session("car_REC") = startRec
End If
'construct dbwhere
If masterdetailwhere <> "" Then
	dbwhere = dbwhere & "(" & masterdetailwhere & ") AND "
End If
If searchwhere <> "" Then
	dbwhere = dbwhere & "(" & searchwhere & ") AND "
End If
If Len(dbwhere) > 5 Then
	dbwhere = Mid(dbwhere, 1, Len(dbwhere)-5) 'trim right most AND
End If
%>
<%
' Load Default Order
DefaultOrder = "make"
DefaultOrderType = "ASC"
'No Default Filter
DefaultFilter = ""
' Check for an Order parameter
OrderBy = ""
If Request.QueryString("order").Count > 0 Then
	OrderBy = Request.QueryString("order")
	' Check If an ASC/DESC toggle is required
	If Session("car_OB") = OrderBy Then
		If Session("car_OT") = "ASC" Then
			Session("car_OT") = "DESC"
		Else
			Session("car_OT") = "ASC"
		End if
	Else
		Session("car_OT") = "ASC"
	End If
	Session("car_OB") = OrderBy
	Session("car_REC") = 1
Else
	OrderBy = Session("car_OB")
	If OrderBy = "" Then
		OrderBy = DefaultOrder
		Session("car_OB") = OrderBy
		Session("car_OT") = DefaultOrderType
	End If
End If
' Open Connection to the database
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str
' Build SQL
strsql = "SELECT * FROM [car]"
If DefaultFilter <> "" Then
	whereClause = whereClause & "(" & DefaultFilter & ") AND "
End If
If dbwhere <> "" Then
	whereClause = whereClause & "(" & dbwhere & ") AND "
End If
If Right(whereClause, 5)=" AND " Then whereClause = Left(whereClause, Len(whereClause)-5)
If whereClause <> "" Then
	strsql = strsql & " WHERE " & whereClause
End If
If OrderBy <> "" Then 
	strsql = strsql & " ORDER BY [" & OrderBy & "] " & Session("car_OT")
End If	
'Response.Write strsql
Set rs = Server.CreateObject("ADODB.Recordset")
rs.cursorlocation = 3
rs.Open strsql, conn, 1, 2
totalRecs = rs.RecordCount
' Check for a START parameter
If Request.QueryString("start").Count > 0 Then
	startRec = Request.QueryString("start")
	Session("car_REC") = startRec
ElseIf Request.QueryString("pageno").Count > 0 Then
	pageno = Request.QueryString("pageno")
	If IsNumeric(pageno) Then
		startRec = (pageno-1)*displayRecs+1
		If startRec <= 0 Then
			startRec = 1
		ElseIf startRec >= ((totalRecs-1)\displayRecs)*displayRecs+1 Then
			startRec = ((totalRecs-1)\displayRecs)*displayRecs+1
		End If
		Session("car_REC") = startRec
	Else
		startRec = Session("car_REC")
		If Not IsNumeric(startRec) Or startRec = "" Then
			'reset start record counter
			startRec = 1
			Session("car_REC") = startRec
		End If
	End If
Else
	startRec = Session("car_REC")
	If Not IsNumeric(startRec) Or startRec = "" Then
		'reSet start record counter
		startRec = 1
		Session("car_REC") = startRec
	End If
End If
%>
<!--#include file="header.asp"-->
<script language="JavaScript">
<!--
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

<body bgcolor="white" text="black" link="blue" vlink="purple" alink="red" OnLoad="na_preload_img(false, 'images/mapprintover.gif');">
<table width="801" align="center" bgcolor="white" cellspacing="0" bordercolordark="white" bordercolorlight="black" cellpadding="0">
    <tr>
        <td width="800" align="center">
            <p><script language="JavaScript" src="ew.js"></script>
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
    
            
<div align="left">
<table border="0" cellpadding="0" cellspacing="0" width="781" bgcolor="#CCCCCC">
        <tr>
            <td width="781" bgcolor="white" height="34"><form onSubmit="return EW_checkMyForm(this);"  action="carsrch.asp" method="post">

<table border="0" cellspacing="0" cellpadding="0" bgcolor="#CCCCCC" width="746" align="center">
<tr>
<td bgcolor="white" width="24" nowrap height="20">
<p>
<input type="hidden" name="a" value="S"></td>
<td bgcolor="white" width="55" nowrap height="20">
<font face="Arial" size="2"><input type="hidden" name="z_make" value="LIKE,'%,%'"><input type="hidden" name="z_model" value="LIKE,'%,%'"><input type="hidden" name="z_type" value="LIKE,'%,%'"></font></td>
<td bgcolor="white" width="113" nowrap height="20"><SPAN class="required" style="font-size:10pt;"><B><font face="Verdana" color="#660000">Narrow Search:</font></B></SPAN></td>
<td bgcolor="white" width="108" nowrap height="20"><font face="Verdana"><span style="font-size:8pt;"><%
x_makeList = "<SELECT name='x_make'><OPTION value=''>All Makes</OPTION>"
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
<td bgcolor="white" width="153" nowrap height="20">
<p align="center"><font face="Verdana"><span style="font-size:8pt;">&nbsp;&nbsp;<input type="text" name="x_model" size="15" maxlength=50 value="<%= Server.HtmlEncode(x_model&"") %>"> 
                                            &nbsp;</span></font></td>
<td bgcolor="white" width="118" nowrap height="20">
<p><font face="Verdana"><span style="font-size:8pt;"><%
x_typeList = "<SELECT name='x_type'><OPTION value=''>All Body Styles</OPTION>"
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
</span></font></td>
<td bgcolor="white" width="73" nowrap height="20">
<input type="submit" name="Action" value=" Search "></td>
<td bgcolor="white" width="102" nowrap valign="middle" height="49" align="right" rowspan="2">
<font size="1" face="Arial"><a href="javascript:window.print()" OnMouseOut="na_restore_img_src('carprin1', 'document')" OnMouseOver="na_change_img_src('carprin1', 'document', 'images/mapprintover.gif', true);"><br><img src="images/mapprint.gif" width="79" height="54" border="0" name="carprin1" align="bottom"></a></font></td>
</tr>
                            <tr>
<td bgcolor="white" width="192" nowrap height="11" colspan="3">
                                               
                                   
</td>
<td bgcolor="white" width="108" nowrap height="11" valign="top">
                                    <font color="#5C7394" face="Verdana"><span style="font-size:10pt;"><b>«</b></span></font><a href="carlist.asp?cmd=reset"><span style="font-size:10pt;"><font face="Verdana" color="#5C7394"><b>Start 
                                    Over</b></font></span></a>
</td>
<td bgcolor="white" width="344" nowrap height="11" colspan="3">	<div class="smartbrowse">
		&nbsp;</div>
</td>
                            </tr>
</table>
</form>
            </td>
        </tr>
    </table>
</div>
            
<table border="1" width="779" cellspacing="0" bordercolordark="white" bordercolorlight="#999999" bordercolor="black" align="center">
    <tr>
        <td width="120" height="15" bgcolor="#CCCCCC"><table border="0" cellspacing="0" cellpadding="0" width="95"><tr><td width="95" height="13">
                        <font size="1" face="Arial">&nbsp;</font><font face="Verdana"><span style="font-size:8pt;">&nbsp;&nbsp;<% If CLng(startRec) > CLng(totalRecs) Then startRec = totalRecs
	stopRec = startRec + displayRecs - 1
	recCount = totalRecs - 1
	If rsEOF Then recCount = totalRecs
	If stopRec > recCount Then stopRec = recCount %><%= startRec %>-<%= stopRec %> of <%= totalRecs %>
                        </span></font>
</td></tr></table>
        </td>
        <td width="273" height="15" bgcolor="#CCCCCC" nowrap>
                        <p><FONT class=smalltext id=inventoryStatusBar face="Verdana"><span style="font-size:8pt;">Use the buttons below to sort our 
inventory.</span></FONT></p>
        </td>
        <td width="372" height="15" bgcolor="#CCCCCC" nowrap><p>
<font face="Arial"><span style="font-size:9pt;">&nbsp;&nbsp;&nbsp;<%
' Display page numbers
If totalRecs > 0 Then
	rsEof = (totalRecs < (startRec + displayRecs))
	' Find out if there should be Backward or Forward Buttons on the table.
	If 	startRec = 1 Then
		isPrev = False
	Else
		isPrev = True
		PrevStart = startRec - displayRecs
		If PrevStart < 1 Then PrevStart = 1 %>	
	</span></font><a href="carlist.asp?start=<%=PrevStart%>"><font face="Arial"><b><span style="font-size:9pt;"><img src="images/carleft.gif" width="60" height="26" border="0" align="texttop"></span></b></font></a><font face="Arial"><span style="font-size:9pt;">
	</span><span style="font-size:16pt;"><b><%
	End If
	If (isPrev OR (NOT rsEof)) Then
		x = 1
		y = 1
		dx1 = ((startRec-1)\(displayRecs*recRange))*displayRecs*recRange+1
		dy1 = ((startRec-1)\(displayRecs*recRange))*recRange+1
		If (dx1+displayRecs*recRange-1) > totalRecs Then
			dx2 = (totalRecs\displayRecs)*displayRecs+1
			dy2 = (totalRecs\displayRecs)+1
		Else
			dx2 = dx1+displayRecs*recRange-1
			dy2 = dy1+recRange-1
		End If
		While x <= totalRecs
			If x >= dx1 AND x <= dx2 Then
				If CLng(startRec) = CLng(x) Then %>
	</b></span><b><span style="font-size:16pt;"><%=y%></span></b><span style="font-size:16pt;"><b>
				<%	Else %>
	</b></span></font><a href="carlist.asp?start=<%=x%>"><font face="Arial"><b><span style="font-size:16pt;"><%=y%></span></b></font></A><font face="Arial"><span style="font-size:16pt;"><b>
				<%	End If
				x = x + displayRecs
				y = y + 1
			ElseIf x >= (dx1-displayRecs*recRange) AND x <= (dx2+displayRecs*recRange) Then
				If x+recRange*displayRecs < totalRecs Then %>
	</b></span></font><a href="carlist.asp?start=<%=x%>"><font face="Arial"><b><span style="font-size:16pt;"><%=y%>-<%=y+recRange-1%></span></b></font></A><font face="Arial"><span style="font-size:16pt;"><b>
				<% Else
					ny=(totalRecs-1)\displayRecs+1
						If ny = y Then %>
	</b></span></font><a href="carlist.asp?start=<%=x%>"><font face="Arial"><b><span style="font-size:16pt;"><%=y%></span></b></font></A><font face="Arial"><span style="font-size:16pt;"><b>
						<% Else %>
	</b></span></font><a href="carlist.asp?start=<%=x%>"><font face="Arial"><b><span style="font-size:16pt;"><%=y%>-<%=ny%></span></b></font></A><font face="Arial"><span style="font-size:16pt;"><b>
						<%	End If
				End If
				x=x+recRange*displayRecs
				y=y+recRange
			Else
				x=x+recRange*displayRecs
				y=y+recRange
			End If
		Wend
	End If
	' Next link
	If NOT rsEof Then
		NextStart = startRec + displayRecs
		isMore = True %></b></span><span style="font-size:9pt;">
	</span></font><a href="carlist.asp?start=<%=NextStart%>"><font face="Arial"><b><span style="font-size:9pt;"><img src="images/carright.gif" width="60" height="26" border="0" align="texttop"></span></b></font></a><font face="Arial"><span style="font-size:9pt;">
	<% Else
		isMore = False
	End If %>
		
	
<% End If %></span></font>                    </td>
    </tr>
    <tr>
        <td width="773" colspan="3"><form method="post">
<table border="0" cellspacing="0" cellpadding="0" bgcolor="#CCCCCC" width="785">
<tr bgcolor="#708090">
<td height="23" width="128" bgcolor="white" background="images/sort_mid.gif" nowrap valign="top">
                            <p><font face="Arial Narrow"><span style="font-size:4pt;">&nbsp;</span></font></p>

</td>
<td height="23" background="images/sort_mid.gif" width="52" nowrap valign="top">
                            <p><a href="carlist.asp?order=<%= Server.URLEncode("year") %>"><font color="#FFFFFF" face="Arial Narrow"><span style="font-size:4pt;"><img src="images/sort_year.gif" width="33" height="18" border="0"></span></font><font face="Arial Narrow" color="#999999"><span style="font-size:4pt;"> 
                            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></font><font color="#FFFFFF" face="Arial Narrow"><span style="font-size:4pt;"><% If OrderBy <> "year" Then %><img src="images/sort_right.gif" width="12" height="18" border="0"><% End If %><% If OrderBy = "year" Then %><% If Session("car_OT") = "ASC" Then %><img src="images/sort_right.gif" width="12" height="18" border="0"><% ElseIf Session("car_OT") = "DESC" Then %><img src="images/sort_select.gif" width="12" height="18" border="0"><% End If %><% End If %></span></font></a></p>

</td>
<td height="23" background="images/sort_mid.gif" width="114" nowrap valign="top">
                            <p><a href="carlist.asp?order=<%= Server.URLEncode("make") %>"><font color="#FFFFFF" face="Arial Narrow"><span style="font-size:4pt;"><img src="images/sort_make.gif" width="38" height="18" border="0"></span></font><font color="#999999" face="Arial Narrow"><span style="font-size:4pt;"> 
                            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></font></a><a href="carlist.asp?order=<%= Server.URLEncode("year") %>"><font color="#FFFFFF" face="Arial Narrow"><span style="font-size:4pt;"><% If OrderBy <> "make" Then %><img src="images/sort_right.gif" width="12" height="18" border="0"><% End If %></span></font></a><a href="carlist.asp?order=<%= Server.URLEncode("make") %>"><font color="#FFFFFF" face="Arial Narrow"><span style="font-size:4pt;"><% If OrderBy = "make" Then %><% If Session("car_OT") = "ASC" Then %><img src="images/sort_blank.gif" width="12" height="18" border="0"><% ElseIf Session("car_OT") = "DESC" Then %><img src="images/sort_select.gif" width="12" height="18" border="0"><% End If %><% End If %></span></font></a></p>

</td>
<td height="23" width="231" background="images/sort_mid.gif" nowrap valign="top">
                            <p><a href="carlist.asp?order=<%= Server.URLEncode("model") %>"><font color="#FFFFFF" face="Arial Narrow"><span style="font-size:4pt;"><img src="images/sort_model.gif" width="42" height="18" border="0"></span></font></a><font color="#FFFFFF" face="Arial Narrow"><span style="font-size:4pt;"> 
                            &nbsp;</span></font><font face="Arial Narrow"><span style="font-size:4pt;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></font><font color="#FFFFFF" face="Arial Narrow"><span style="font-size:4pt;">&nbsp;&nbsp;</span></font><a href="carlist.asp?order=<%= Server.URLEncode("year") %>"><font color="#FFFFFF" face="Arial Narrow"><span style="font-size:4pt;"><% If OrderBy <> "model" Then %><img src="images/sort_right.gif" width="12" height="18" border="0"><% End If %></span></font></a><a href="carlist.asp?order=<%= Server.URLEncode("model") %>"><font color="#FFFFFF" face="Arial Narrow"><span style="font-size:4pt;"><% If OrderBy = "model" Then %><% If Session("car_OT") = "ASC" Then %><img src="images/sort_blank.gif" width="12" height="18" border="0"><% ElseIf Session("car_OT") = "DESC" Then %><img src="images/sort_select.gif" width="12" height="18" border="0"><% End If %><% End If %></span></font></a></p>

</td>
<td height="23" width="81" background="images/sort_mid.gif" nowrap valign="top">
                            <p><a href="carlist.asp?order=<%= Server.URLEncode("type") %>"><font color="#FFFFFF" face="Arial Narrow"><span style="font-size:4pt;"><img src="images/sort_type.gif" width="33" height="18" border="0"></span></font></a><font color="#FFFFFF" face="Arial Narrow"><span style="font-size:4pt;"> 
                            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></font><a href="carlist.asp?order=<%= Server.URLEncode("year") %>"><font color="#FFFFFF" face="Arial Narrow"><span style="font-size:4pt;"><% If OrderBy <> "type" Then %><img src="images/sort_right.gif" width="12" height="18" border="0"><% End If %></span></font></a><a href="carlist.asp?order=<%= Server.URLEncode("type") %>"><font color="#FFFFFF" face="Arial Narrow"><span style="font-size:4pt;"><% If OrderBy = "type" Then %><% If Session("car_OT") = "ASC" Then %><img src="images/sort_blank.gif" width="12" height="18" border="0"><% ElseIf Session("car_OT") = "DESC" Then %><img src="images/sort_select.gif" width="12" height="18" border="0"><% End If %><% End If %></span></font></a></p>

</td>
<td height="23" width="87" background="images/sort_mid.gif" nowrap valign="top">
                            <p><a href="carlist.asp?order=<%= Server.URLEncode("miles") %>"><font color="#FFFFFF" face="Arial Narrow"><span style="font-size:4pt;"><img src="images/sort_mileage.gif" width="39" height="18" border="0"></span></font></a><font color="#FFFFFF" face="Arial Narrow"><span style="font-size:4pt;"> 
                            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></font><a href="carlist.asp?order=<%= Server.URLEncode("year") %>"><font color="#FFFFFF" face="Arial Narrow"><span style="font-size:4pt;"><% If OrderBy <> "miles" Then %><img src="images/sort_right.gif" width="12" height="18" border="0"><% End If %></span></font></a><a href="carlist.asp?order=<%= Server.URLEncode("miles") %>"><font color="#FFFFFF" face="Arial Narrow"><span style="font-size:4pt;"><% If OrderBy = "miles" Then %><% If Session("car_OT") = "ASC" Then %><img src="images/sort_blank.gif" width="12" height="18" border="0"><% ElseIf Session("car_OT") = "DESC" Then %><img src="images/sort_select.gif" width="12" height="18" border="0"><% End If %><% End If %></span></font></a></p>

</td>
<td height="23" width="70" background="images/sort_mid.gif" nowrap valign="top">
                            <p><a href="carlist.asp?order=<%= Server.URLEncode("price") %>"><font color="#FFFFFF" face="Arial Narrow"><span style="font-size:4pt;"><img src="images/sort_price.gif" width="37" height="18" border="0"></span></font></a><font color="#FFFFFF" face="Arial Narrow"><span style="font-size:4pt;"> 
                            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></font><a href="carlist.asp?order=<%= Server.URLEncode("year") %>"><font color="#FFFFFF" face="Arial Narrow"><span style="font-size:4pt;"><% If OrderBy <> "price" Then %><img src="images/sort_right.gif" width="12" height="18" border="0"><% End If %></span></font></a><a href="carlist.asp?order=<%= Server.URLEncode("price") %>"><font color="#FFFFFF" face="Arial Narrow"><span style="font-size:4pt;"><% If OrderBy = "price" Then %><% If Session("car_OT") = "ASC" Then %><img src="images/sort_blank.gif" width="12" height="18" border="0"><% ElseIf Session("car_OT") = "DESC" Then %><img src="images/sort_select.gif" width="12" height="18" border="0"><% End If %><% End If %></span></font></a></p>

</td>
<td height="23" width="22" bgcolor="white" nowrap valign="top" background="images/sort_mid.gif">
                            <p><font face="Arial Narrow"><span style="font-size:4pt;">&nbsp;</span></font></p>
</td>
</tr>
<%
'Avoid starting record > total records
If CLng(startRec) > CLng(totalRecs) Then
	startRec = totalRecs
End If
'Set the last record to display
stopRec = startRec + displayRecs - 1
'Move to first record directly for performance reason
recCount = startRec - 1
If NOT rs.EOF Then
	rs.MoveFirst
	rs.Move startRec - 1
End If
recActual = 0
Do While (NOT rs.EOF) AND (recCount < stopRec)
	recCount = recCount + 1
	If CLng(recCount) >= CLng(startRec) Then 
		recActual = recActual + 1 %>
<%
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
	'Load Key for record
	key = rs("ID")
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
<td width="128" height="66" align="center" valign="top">
                            <table border="1">
                                <tr>
                                    <td width="100"><font face="Arial" size="2"><% If not isnull(x_photo_1) Then %>
<a href="carview.asp?key=<%= key %>&nr=1"><img src="car_ph.asp?key=<%= key %>&nr=1" border=0 width="110" height="82"></a></font><strong><font size="-1"><%else%>
        <a href="carview.asp?key=<%= key %>"><img src="images/sedan.gif" border=0 align="absmiddle" width="110" height="82"></a></font></strong><font face="Arial"><span style="font-size:8pt;"><% End If %></span></font></td>
                                </tr>
                            </table>
</td>
<td nowrap width="52" height="66" valign="top"><font face="Arial" size="2">&nbsp;<%
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
<td nowrap width="114" height="66" valign="top"><font face="Arial" size="2">&nbsp;<%Select Case x_make
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
%></font></td>
<td nowrap width="231" height="66" valign="top"><font face="Arial" size="2">&nbsp;<% response.write x_model %>&nbsp;<% response.write x_engine %></font></td>
<td nowrap width="81" height="66" valign="top"><font face="Arial" size="2">&nbsp;<%
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
%>&nbsp;</font></td>
<td nowrap width="87" height="66" valign="top">
                            <p align="right"><font face="Arial" size="2">&nbsp;<% if isnumeric(x_miles) then response.write formatnumber(x_miles,0,-2,-2,-2) else response.write x_miles end if %> 
                            &nbsp;&nbsp;&nbsp;</font></td>
<td nowrap width="70" height="66" valign="top">
                            <p align="right"><font face="Arial" size="2">&nbsp;<% if isnumeric(x_price) then response.write formatcurrency(x_price,0,-2,-2,-2) else response.write x_price end if %> 
                            &nbsp;&nbsp;&nbsp;</font></td>
<td nowrap width="22" height="66" valign="top">&nbsp;
                                        <p>&nbsp;
                                        <p> 
                            &nbsp;</p>
</td>
</tr>
<%
	End If
	rs.MoveNext
Loop
%>
</table>
<% If recActual > 0 Then %><% End If %>
</form>
                        <p align="center"><%
' Close recordSet and connection
rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing %><font face="Arial"><span style="font-size:9pt;">&nbsp;</span></font>
                        <p align="center"><font face="Arial"><span style="font-size:9pt;">&nbsp;<%
' Display page numbers
If totalRecs > 0 Then
	rsEof = (totalRecs < (startRec + displayRecs))
	' Find out if there should be Backward or Forward Buttons on the table.
	If 	startRec = 1 Then
		isPrev = False
	Else
		isPrev = True
		PrevStart = startRec - displayRecs
		If PrevStart < 1 Then PrevStart = 1 %>	
	</span></font><a href="carlist.asp?start=<%=PrevStart%>"><font face="Arial"><b><span style="font-size:12pt;"><img src="images/carleft.gif" width="70" height="30" border="0" align="texttop"></span></b></font></a><font face="Arial"><span style="font-size:12pt;"><b>
	</b></span><span style="font-size:20pt;"><b><%
	End If
	If (isPrev OR (NOT rsEof)) Then
		x = 1
		y = 1
		dx1 = ((startRec-1)\(displayRecs*recRange))*displayRecs*recRange+1
		dy1 = ((startRec-1)\(displayRecs*recRange))*recRange+1
		If (dx1+displayRecs*recRange-1) > totalRecs Then
			dx2 = (totalRecs\displayRecs)*displayRecs+1
			dy2 = (totalRecs\displayRecs)+1
		Else
			dx2 = dx1+displayRecs*recRange-1
			dy2 = dy1+recRange-1
		End If
		While x <= totalRecs
			If x >= dx1 AND x <= dx2 Then
				If CLng(startRec) = CLng(x) Then %>
	</b></span><b><span style="font-size:20pt;"><%=y%></span></b><span style="font-size:20pt;"><b>
				<%	Else %>
	</b></span></font><a href="carlist.asp?start=<%=x%>"><font face="Arial"><b><span style="font-size:20pt;"><%=y%></span></b></font></A><font face="Arial"><span style="font-size:20pt;"><b>
				<%	End If
				x = x + displayRecs
				y = y + 1
			ElseIf x >= (dx1-displayRecs*recRange) AND x <= (dx2+displayRecs*recRange) Then
				If x+recRange*displayRecs < totalRecs Then %>
	</b></span></font><a href="carlist.asp?start=<%=x%>"><font face="Arial"><b><span style="font-size:20pt;"><%=y%>-<%=y+recRange-1%></span></b></font></A><font face="Arial"><span style="font-size:20pt;"><b>
				<% Else
					ny=(totalRecs-1)\displayRecs+1
						If ny = y Then %>
	</b></span></font><a href="carlist.asp?start=<%=x%>"><font face="Arial"><b><span style="font-size:20pt;"><%=y%></span></b></font></A><font face="Arial"><span style="font-size:20pt;"><b>
						<% Else %>
	</b></span></font><a href="carlist.asp?start=<%=x%>"><font face="Arial"><b><span style="font-size:20pt;"><%=y%>-<%=ny%></span></b></font></A><font face="Arial"><span style="font-size:20pt;"><b>
						<%	End If
				End If
				x=x+recRange*displayRecs
				y=y+recRange
			Else
				x=x+recRange*displayRecs
				y=y+recRange
			End If
		Wend
	End If
	' Next link
	If NOT rsEof Then
		NextStart = startRec + displayRecs
		isMore = True %>
	</b></span></font><a href="carlist.asp?start=<%=NextStart%>"><font face="Arial"><b><span style="font-size:12pt;"><img src="images/carright.gif" width="70" height="30" border="0" align="texttop"></span></b></font></a><font face="Arial"><span style="font-size:12pt;"><b>
	<% Else
		isMore = False
	End If %>
		
	
<% End If %></b></span></font></p>


        </td>
    </tr>
</table>
<form action="carlist.asp">
                
</form>
        </td>
    </tr>
</table>
<!--#include file="footer.asp"-->

















