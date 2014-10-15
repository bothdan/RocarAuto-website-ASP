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
		strsql = "SELECT * FROM [contact] WHERE [ID]=" & tkey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn
		If rs.EOF Then
			Response.Clear
			Response.Redirect "contactlist.asp"
		Else
			rs.MoveFirst
		' Get the field contents
		x_first_name = rs("first name")
		x_last_name = rs("last name")
		x_email = rs("email")
		x_phone = rs("phone")
		x_comments = rs("comments")
		End If
		rs.Close
		Set rs = Nothing
	Case "A": ' Add
		'get fields from form
x_ID = Request.Form("x_ID")
x_first_name = Request.Form("x_first_name")
x_last_name = Request.Form("x_last_name")
x_email = Request.Form("x_email")
x_phone = Request.Form("x_phone")
x_comments = Request.Form("x_comments")
		' Open record
		strsql = "SELECT * FROM [contact] WHERE 0 = 1"
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn, 1, 2
		rs.AddNew
		tmpFld = Trim(x_first_name)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("first name") = tmpFld
		tmpFld = Trim(x_last_name)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("last name") = tmpFld
		tmpFld = Trim(x_email)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("email") = tmpFld
		tmpFld = Trim(x_phone)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("phone") = tmpFld
		tmpFld = Trim(x_comments)
		If trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("comments") = tmpFld
		rs.Update
		rs.Close
		Set rs = Nothing
		conn.Close
		Set conn = Nothing
		Response.Clear
		Response.Redirect "contactlist.asp"
End Select
%>
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
<body OnLoad="na_preload_img(false, 'images/mapprintover.gif');">
<table cellpadding="0" cellspacing="0" width="800" align="center">
    <tr>
        <td width="800" colspan="2" background="images/contactbg.gif">
            <p><img src="images/maptop.gif" width="439" height="32" border="0"></p>
        </td>
    </tr>
    <tr>
        <td width="569" bgcolor="white" height="406" align="center" valign="top">
            <p align="center"><font face="Arial"><b><span style="font-size:16pt;"><br><a href="javascript:na_open_window('Directions', 'images/map.gif', 0, 0, 554, 561, 0, 0, 0, 0, 0)" target="_self"><img src="images/mapsm.gif" width="450" height="350" border="0"></a><br></span></b><span style="font-size:9pt;">click 
            to enlarge</span></font></td>
        <td width="231" bgcolor="white" valign="top" height="465" rowspan="2">
            <table align="center" cellpadding="0" cellspacing="0" width="214">
                <tr>
                    <td width="12" height="9">
                        <p><font face="Arial"><span style="font-size:5pt;"><img src="images/upleft.gif" width="12" height="10" border="0"></span></font></p>
                    </td>
                    <td width="194" background="images/up.gif" height="9">
                        <p><font face="Arial"><span style="font-size:5pt;">&nbsp;</span></font></p>
                    </td>
                    <td width="8" height="9">
                        <p><font face="Arial"><span style="font-size:5pt;"><img src="images/upright.gif" width="13" height="10" border="0"></span></font></p>
                    </td>
                </tr>
                <tr>
                    <td width="12" rowspan="18" background="images/leftbg.gif" height="314">
                        <p><font face="Arial"><span style="font-size:8pt;">&nbsp;</span></font></p>
                    </td>
                    <td width="194" nowrap><SPAN style="font-size:10pt;"><font face="Arial"><b>Contact Information</b></font></SPAN></td>
                    <td width="8" rowspan="18" background="images/rightbg.gif" height="314">
                        <p><font face="Arial"><span style="font-size:8pt;">&nbsp;</span></font></p>
                    </td>
                </tr>
                <tr>
                    <td width="194" nowrap>
                        <p><font face="Arial"><span style="font-size:8pt;">&nbsp;</span></font></p>
                    </td>
                </tr>
                <tr>
                    <td width="194" nowrap height="17">
                        <p><font face="Arial"><span style="font-size:9pt;">5136 
                        N. Western Ave.</span></font></p>
                    </td>
                </tr>
                <tr>
                    <td width="194" nowrap height="17">
                        <p><font face="Arial"><span style="font-size:9pt;">Chicago, 
                        IL 60625-2533</span></font></p>
                    </td>
                </tr>
                <tr>
                    <td width="194" nowrap height="17"><span style="font-size:9pt;"><font face="Arial">Phone: (773) 334-0025 </font></span></td>
                </tr>
                <tr>
                    <td width="194" nowrap height="17"><span style="font-size:9pt;"><font face="Arial">Fax: 773-334-5036 </font></span></td>
                </tr>
                <tr>
                    <td width="194" nowrap height="17">
                        <p><font face="Arial"><span style="font-size:9pt;">Email: 
                        </span></font><a href="contactadd.asp"><span style="font-size:9pt;"><b><font face="Arial" color="teal">Email 
                        Us</font></b></span></a></p>
                    </td>
                </tr>
                <tr>
                    <td width="194" nowrap>
                        <p><font face="Arial"><span style="font-size:8pt;">&nbsp;</span></font></p>
                    </td>
                </tr>
                <tr>
                    <td width="194" nowrap><SPAN style="font-size:10pt;"><b><font face="Arial">Dealership Hours</font></b></SPAN></td>
                </tr>
                <tr>
                    <td width="194" nowrap>
                        <p><font face="Arial"><span style="font-size:8pt;">&nbsp;</span></font></p>
                    </td>
                </tr>
                <tr>
                    <td width="194" nowrap height="17"><span style="font-size:9pt;"><font face="Arial">Monday:&nbsp;10:00 AM - 7:00 PM </font></span></td>
                </tr>
                <tr>
                    <td width="194" nowrap height="17"><span style="font-size:9pt;"><font face="Arial">Tuesday:&nbsp;10:00 AM - 7:00 PM </font></span></td>
                </tr>
                <tr>
                    <td width="194" nowrap height="17"><span style="font-size:9pt;"><font face="Arial">Wednesday:&nbsp;10:00 AM - 7:00 PM </font></span></td>
                </tr>
                <tr>
                    <td width="194" nowrap height="17"><span style="font-size:9pt;"><font face="Arial">Thursday:&nbsp;10:00 AM - 7:00 PM </font></span></td>
                </tr>
                <tr>
                    <td width="194" nowrap height="17"><span style="font-size:9pt;"><font face="Arial">Friday:&nbsp;10:00 AM - 7:00 PM </font></span></td>
                </tr>
                <tr>
                    <td width="194" nowrap><span style="font-size:9pt;"><font face="Arial">Saturday:&nbsp;10:00 AM - 6:00 PM</font></span></td>
                </tr>
                <tr>
                    <td width="194" nowrap height="17"><span style="font-size:9pt;"><font face="Arial">Sunday:&nbsp;Closed&nbsp;</font></span></td>
                </tr>
                <tr>
                    <td width="194" nowrap height="20">
                        <p align="center"><a href="javascript:window.print()" OnMouseOut="na_restore_img_src('mapprint1', 'document')" OnMouseOver="na_change_img_src('mapprint1', 'document', 'images/mapprintover.gif', true);"><br><img src="images/mapprint.gif" width="79" height="54" border="0" name="mapprint1"></a></p>
                    </td>
                </tr>
                <tr>
                    <td width="12" height="12">
                        <p><font face="Arial"><span style="font-size:8pt;"><img src="images/dwleft.gif" width="12" height="11" border="0"></span></font></p>
                    </td>
                    <td width="194" background="images/dwbg.gif" height="12">
                        <p><font face="Arial"><span style="font-size:5pt;">&nbsp;</span></font></p>
                    </td>
                    <td width="8" height="12">
                        <p><font face="Arial"><span style="font-size:5pt;"><img src="images/dwright.gif" width="13" height="11" border="0"></span></font></p>
                    </td>
                </tr>
            </table>
        </td>
    </tr>
    <tr>
        <td width="569" bgcolor="white" height="26">
            <p>&nbsp;&nbsp;<a href="http://www.mapquest.com/" target="_blank"><img src="images/mapquest_logo.gif" width="225" height="53" border="0"></a><br> 
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font face="Arial"><b><span style="font-size:10pt;"><a href="http://www.mapquest.com/" target="_blank">www.MapQuest.com</a><br>&nbsp;</span></b></font></p>
        </td>
    </tr>
</table>
<!--#include file="footer.asp"-->