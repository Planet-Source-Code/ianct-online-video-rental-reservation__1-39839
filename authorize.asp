<HTML>
<TITLE>Video West Reservations Online</TITLE>

<style type="text/css">
/* body properties */
body         { font-family : Verdana, sans-serif; color : #000000; font-size : 12px }
a            {     color : #333333; } a:link    { color : #333333; text-decoration : none; }
a:visited    {   color : #333333; text-decoration : none; }
a:hover      {   color : #333333; text-decoration : underline; }
a:active     {     color : #333333; text-decoration : underline; } /* normal table properties */
table        { font-size : 12px; } table.border { background-color : #443088; } /* normal font properties */
/* font         { font-family : Verdana, sans-serif; color : #333333; font-size : 10px; } font.sub-text { font-size   : 10px;} */  /* menu properties */
tr.menu      { background-color: #CCCCCC; font-family : Verdana, sans-serif; font-size : 13px; color : #333333; font-weight : bold; } /* item properties */
tr.item      { background-color: #EEEEEE; font-family : Verdana, sans-serif; font-size : 11px; color : #333333; }

BODY {
 		scrollbar-face-color: #000000;
		scrollbar-shadow-color: #222222;
 		scrollbar-highlight-color: #bbbbbb; 
 		scrollbar-3dlight-color: #000000;
 		scrollbar-darkshadow-color: #000000;
 		scrollbar-track-color: #333333;
 		scrollbar-arrow-color: #ffffff;
}  


</style>
<body bgcolor="#FFFFFF">
Video West Online Reservations<br>
[Verify Information]<br>
<br>
<%
strCustID = request.form("CustID")
strCustPhone = request.form("CustPhone")
strReserveTitle = request.form("ReserveTitle")
strReserveFormat = request.form("ReserveFormat")
strDate = request.form("ReserveDate")

if len(strDate) = 0 or len(strCustID) = 0 or len(strCustPhone) = 0 then
	%>
	Sorry, ReserveDate, Customer ID <b>and</b> Customer Phone are required items to place movies on reserve.<br>
	<br>
	<a href="javascript:history.go(-1)">Go back</a><br>

	<font size=1>iReserve by <a href="http://www.IanThurston.com">IanThurston.com</a></font><br>
	<%
	response.end
	end if


	set conn = server.createobject("adodb.connection")
	DSNtemp="DRIVER={Microsoft Access Driver (*.mdb)}; "
        DSNtemp=dsntemp & "DBQ=" & server.mappath("reservations.mdb")

	Conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
           "Data Source=" & Server.MapPath("reservations.mdb")
	
	SQLstmt = "SELECT * FROM Customers WHERE CustID='" & strCustID & "' AND CustPhone='" & strCustPhone & "'"

'	SQLstmt = "INSERT INTO Guestbook (Name,City,State,Country,Email,URL,Comments)"
'	SQLstmt = SQLstmt & " VALUES (" 
'	SQLstmt = SQLstmt & "'" & Name & "',"
'	SQLstmt = SQLstmt & "'" & City & "',"
'	SQLstmt = SQLstmt & "'" & State & "',"
'	SQLstmt = SQLstmt & "'" & Country & "',"
'	SQLstmt = SQLstmt & "'" & Email & "',"
'	SQLstmt = SQLstmt & "'" & URL & "',"
'	SQLstmt = SQLstmt & "'" & Comments & "'"
'	SQLstmt = SQLstmt & ")"


	Set rs = Server.CreateObject("ADODB.Recordset")
	'Set RS = conn.execute(SQLstmt)
	rs.Open sqlstmt, conn, 3, 3

	if rs.recordcount>0 then
		strCustFirstName = rs("CustFirstName")
		strCustLastName = rs("CustLastName")
		else
		%>
		<u><b>Parameters</b></u><br>
		CustID: <%=strCustID%><br>
		CustPhone: <%=strCustPhone%><br>
		ReserveTitle: <%=strReserveTitle%><br>
		<br><br>
		No records found
		<%
		response.end
		end if
   	Conn.Close
	Set conn = nothing
	%>



<TABLE cellpadding=1 cellspacing=1 bgcolor="#000000">
<TR class="MENU">
 <TD>
 Date
 </TD>
 <TD>
 CustID
 </TD>
 <TD>
 Phone Number
 </TD>
 <TD>
 Last Name
 </TD>
 <TD>
 First Name
 </TD>
 <TD>
 Reservation Title
 </TD>
 <TD>
 Reservation Format
 </TD>
</TR>
<TR class="ITEM">
<TD>
 <%=strDate%>
 </TD>

<%


if not isEmpty(strCustID) and len(strCustID) > 0 then
	%>
 	<TD>
	<%=strCustID%>
  	</TD>
	<%
	else
	%>
	<TD></TR>
	<%
	end if

if not isEmpty(strCustPhone) and len(strCustPhone) > 0 then
	%>
	<TD>
	<%=strCustPhone%>
	</TD>
	<%
	else
	%>
	<TD></TD>
	<%
	end if

if not isEmpty(strCustLastName) and len(strCustLastName) > 0 then
	%>
	<TD>
	<%=strCustLastName%>
	</TD>
	<%
	else
	%>
	<TD></TD>
	<%
	end if

if not isEmpty(strCustFirstName) and len(strCustFirstName) > 0 then
	%>
	<TD>
	<%=strCustFirstName%>
	</TD>
	<%
	else
	%>
	<TD></TD>
	<%
	end if

if not isEmpty(strReserveTitle) and len(strReserveTitle) > 0 then
	%>
	<TD>
	<%=strReserveTitle%>
	</TD>
	<%
	else
	%>
	<TD></TD>
	<%
	end if

if not isEmpty(strReserveFormat) and len(strReserveFormat) > 0 then
	%>
	<TD>
	<%=strReserveFormat%>
	</TD>
	<%
	else
	%>
	<TD></TD>
	<%
	end if
	%>

</TR>
</TABLE>
<FORM action="reserve.asp" method="post">
<input type="hidden" name="CustID" value="<%=strCustID%>">
<input type="hidden" name="CustPhone" value="<%=strCustPhone%>">
<input type="hidden" name="CustFirstName" value="<%=strCustFirstName%>">
<input type="hidden" name="CustLastName" value="<%=strCustLastName%>">
<input type="hidden" name="ReserveTitle" value="<%=strReserveTitle%>">
<input type="hidden" name="ReserveFormat" value="<%=strReserveFormat%>">
<input type="hidden" name="ReserveDate" value="<%=strDate%>">

Please verify the above information and <input type="Submit" value="Click Here" STYLE="font-size:8pt; font-family:Verdana"> to make your reservation.
<br>
<br>
<br>
<font size=1>iReserve by <a href="http://www.IanThurston.com">IanThurston.com</a></font><br>
</BODY>
</HTML>