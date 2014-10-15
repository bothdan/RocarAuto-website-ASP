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
<table align="center" border="1" cellspacing="0" width="803" bordercolordark="white" bordercolorlight="black" bgcolor="whitesmoke">
    <tr>
        <td width="797"><script language="JavaScript" src="ew.js"></script>
<script language="JavaScript">
<!-- start Javascript
function  EW_checkMyForm(EW_this) {
return true;
}
// end JavaScript -->
</script>
            <p>&nbsp;</p>

            <table align="center" border="1" bgcolor="white" width="604" cellspacing="0" bordercolordark="white" bordercolorlight="black">
                <tr>
                    <td width="594" height="47">
                        <p align="center"><font face="Verdana" color="#990000"><span style="font-size:16pt;">Thank 
                        You!</span></font></p>
                        <p align="center"><font face="Verdana"><span style="font-size:14pt;">Soon one 
                        of our representativ will contact you.</span></font></p>
                        <p align="center">&nbsp;</p>
</td>
                </tr>
            </table>
            <p align="center"><a href="carlist.asp?cmd=resetall"><span style="font-size:14pt;"><font face="Verdana" color="black">Back 
            to ROCAR Inventory</font></span></a></p>
            <p>&nbsp;</p>
        </td>
    </tr>
</table>
<!--#include file="footer.asp"-->

