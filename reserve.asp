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
[Reserved]<br>
<br>
<%
strCustID = request.form("CustID")
strCustPhone = request.form("CustPhone")
strReserveTitle = request.form("ReserveTitle")
strReserveFormat = request.form("ReserveFormat")
strFirstName = request.form("CustFirstName")
strLastName = request.form("CustLastName")
strDate = request.form("ReserveDate")


	set conn = server.createobject("adodb.connection")
	DSNtemp="DRIVER={Microsoft Access Driver (*.mdb)}; "
        DSNtemp=dsntemp & "DBQ=" & server.mappath("reservations.mdb")

	Conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
           "Data Source=" & Server.MapPath("reservations.mdb")
	
	SQLstmt = "INSERT INTO Reservations (ReserveDate, CustID, CustPhone, CustFirstName, CustLastName, ReserveTitle, ReserveFormat)"
	SQLstmt = SQLstmt & " VALUES (" 
	SQLstmt = SQLstmt & "'" & strDate & "',"
	SQLstmt = SQLstmt & "'" & strCustID & "',"
	SQLstmt = SQLstmt & "'" & strCustPhone & "',"
	SQLstmt = SQLstmt & "'" & strFirstName & "',"
	SQLstmt = SQLstmt & "'" & strLastName & "',"
	SQLstmt = SQLstmt & "'" & strReserveTitle & "',"
	SQLstmt = SQLstmt & "'" & strReserveFormat & "'"
	SQLstmt = SQLstmt & ")"

'	response.write(sqlstmt)


	Set RS = conn.execute(SQLstmt)

	If err.number>0 then
	response.write "VBScript Errors Occured:" & "<P>"
    	response.write "Error Number=" & err.number & "<P>"
    	response.write "Error Descr.=" & err.description & "<P>"
    	response.write "Help Context=" & err.helpcontext & "<P>" 
    	response.write "Help Path=" & err.helppath & "<P>"
    	response.write "Native Error=" & err.nativeerror & "<P>"
    	response.write "Source=" & err.source & "<P>"
    	response.write "SQLState=" & err.sqlstate & "<P>"
   	end if
   	IF conn.errors.count> 0 then
    	response.write "Database Errors Occured" & "<P>"
    	response.write SQLstmt & "<P>"
   		for counter= 0 to conn.errors.count
    		response.write "Error #" & conn.errors(counter).number & "<P>"
    		response.write "Error desc. -> " & conn.errors(counter).description & "<P>"
   		next
   	else
    	response.write "<br><font face='verdana' size=3><b>"
	response.write "Thank you! Your reservation has been added to the system.</font></b><p>"
    	response.write "<font face='verdana' size=2>"
	response.write "Important: Your reservation is NOT guaranteed.<br>"

   	end if
   	Conn.Close
	Set conn = nothing
	%>



<br>

<font size=1>iReserve by <a href="http://www.IanThurston.com">IanThurston.com</a></font><br>
</BODY>
</HTML>