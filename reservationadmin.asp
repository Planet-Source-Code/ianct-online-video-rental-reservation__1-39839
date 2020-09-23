<html>
<head>
<title>Reservation Administration</title>

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
</head>

<body bgcolor="#FFFFFF">
Video West Online Reservations<br>
[Reservation Administration]<br>
<a href="customeradmin.asp">[Customer Administration]<br>
<a href="default.asp">[Make a Reservation]</a><br>
<br>
<%
Flag = request.form("Flag")
Select Case Flag
	Case 0
		Flag = Flag + 1
%>
<form method="post" action="reservationadmin.asp">
<table border=0 cellpadding=1 cellspacing=1 bgcolor="#000000">
  <tr class="ITEM">
    <td>
	<b>UserID</b>
    </td>

    <td>
	<input type="text" name="userid" STYLE="font-size:8pt; font-family:Verdana">
    </td>
  </tr>
  
  <tr class="ITEM">
    <td>
	<b>Password:</b>
    </td>
    <td>
	<input type="password" name="password" STYLE="font-size:8pt; font-family:Verdana">
    </td>
  </tr>

  <tr class="ITEM">
    <td>
	<b>RetrieveDate:</b><br>Leave blank for All
    </td>
    <td>
	<input type="text" name="retrievedate" STYLE="font-size:8pt; font-family:Verdana" value="<%=Date%>" size=10>
    </td>
  </tr>
  
  <tr class="ITEM">
	<td colspan=2 align="center">
	<input type='hidden' name="flag" value="<%= Flag %>">
	<input type="submit" value="Submit" STYLE="font-size:8pt; font-family:Verdana">&nbsp;<input type="reset" STYLE="font-size:8pt; font-family:Verdana">
	</td>
  </tr>
</table>
  
<%
    Case 1
		UserID = request.form("userid")
		Password = request.form("password")
		RetrieveDate = request.form("retrievedate")
		' **** be sure to change "admin" and "password" 
		' **** to your choice of userid and password
		If UserID <> "admin" OR Password <> "password" then
			response.write "<font face='verdana' color='red'>"
			response.write "Invalid login. Please try again.</font>"
		Else
'		set conn = server.createobject("adodb.connection")
'		DSNtemp="DRIVER={Microsoft Access Driver (*.mdb)}; "
'		DSNtemp=dsntemp & "DBQ=" & server.mappath("reservations.mdb")
'		conn.Open DSNtemp

	set conn = server.createobject("adodb.connection")
	DSNtemp="DRIVER={Microsoft Access Driver (*.mdb)}; "
        DSNtemp=dsntemp & "DBQ=" & server.mappath("reservations.mdb")

	Conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
           "Data Source=" & Server.MapPath("reservations.mdb")

		if len(retrievedate) > 0 then 
	    		sqlstmt = "SELECT * from Reservations WHERE ReserveDate='" & retrievedate & "' Order by ReserveDate Desc"
			else
	    		sqlstmt = "SELECT * from Reservations Order by ReserveDate Desc"
			end if

 			Set RS = conn.execute(sqlstmt)
  			If rs.eof then
      			response.write "<center>There are no records in the database"
	  			response.write "<br>Please check back later</center>"
	  			response.end
				End If
			%>
			<table bgcolor="#000000" border=0 cellpadding=1 cellspacing=1>
			<tr class="MENU">
			 <td colspan=7 align='center'>
			 <font face='verdana' size=3>
			 <b>Delete Records</b>
			 </td>
			</tr>

			<form action="reservationadmin.asp" method="post">
			<TR class="ITEM">
			 <TD colspan=7 align="center">
			 <input type="Submit" value="Show by Date" STYLE="font-size:8pt; font-family:Verdana"><br>
			 <input type="text" size=10 name="retrievedate" value="">
 			 <br>Leave Blank for All
			 <input type="hidden" name="UserID" value="<%= UserID %>">
			 <input type="hidden" name="Password" value="<%= Password %>">
			 <input type="hidden" name='Flag' value='1'>
			 </TD>
			</TR>
			</form>

			<TR class="ITEM">
			 <TD>Del</TD>
			 <TD>Reserve Date</TD>
			 <TD>Customer ID</TD>
			 <TD>Customer Phone</TD>
			 <TD>Customer Name</TD>
			 <TD>Reserve Title</TD>
			 <TD>Reserve Format</TD>
			</TR>

			<TR class="MENU">
			 <TD align="center" colspan="7">
			  <% RotateDate = rs("ReserveDate") %>
			 <%=RotateDate%>
			 </TD>
			</TR>
			<%
			response.write "<form action='reservationadmin.asp' method='post'>"

  	   		Do while not rs.eof
		  		' The database has an autonumber field set as
		  		' the primary key, so we will use that field
		  		' to specify which record we want to modify
	ID = rs("ID")
        strCustID = rs("CustID")
	strCustPhone = rs("CustPhone")
	strReserveDate = rs("ReserveDate")
	strName = rs("CustLastName") & ", " & rs("CustFirstName")
	strReserveTitle = rs("ReserveTitle")
	strReserveFormat = rs("ReserveFormat")


	if strReserveDate <> RotateDate then
		%>
		<TR class="MENU">
		 <TD align="center" colspan="7">
		  <% RotateDate = rs("ReserveDate") %>
		 <%=RotateDate%>
		 </TD>
		</TR>
		<%
		RotateDate = strReserveDate
		end if

	if UseColor = "#c0c0FF" then UseColor = "#8080FF" else UseColor="#c0c0FF"

  %>


  <tr class="ITEM">
    <td width=15 bgcolor="<%=UseColor%>">
	<input type="checkbox" name="ID" value="<%= ID %>">
	</td>
	
	<td bgcolor="<%=UseColor%>">
        	<font size=-2 face="verdana"><%= strReserveDate %></font>
	</td>

	<td bgcolor="<%=UseColor%>">
		<font size=-2 face="verdana"><%= strCustID %></font>
	</td>

	<td bgcolor="<%=UseColor%>">
		<font size=-2 face="verdana"><%= strCustPhone %></font>
	</td>

	<td bgcolor="<%=UseColor%>">
		<font size=-2 face="verdana"><%= strName %></font>
	</td>

	<td bgcolor="<%=UseColor%>">
		<font size=-2 face="verdana"><%= strReserveTitle %></font>
	</td>

	<td bgcolor="<%=UseColor%>">
        	<font size=-2 face="verdana"><%= strReserveFormat %></font>
	</td>
  </tr>
  <%
   		rs.MoveNext
 	   		loop 
  %>
    <tr class="ITEM">
    <td colspan=7 align="center">
	<input type='hidden' name="retrievedate" value="<%=retrievedate%>">
	<input type='hidden' name="UserID" value="<%= UserID %>">
	<input type='hidden' name="Password" value="<%= Password %>">
	<input type='hidden' name='Flag' value='2'>
	<input type="submit" value="Delete Record(s)" STYLE="font-size:8pt; font-family:Verdana">
	</td>
	</form>

  </tr>
<%
  		End If
	Case 2
	    If IsEmpty(request.form("ID")) then
			response.write "<font face='verdana' size=3 color='red'>"
			response.write "Oops! You have to check a "
			response.write "box for this to work!"
			response.write "<br>Please hit your Back"
			response.write " button and try again."
			response.end
		End If
		set rs = nothing
	    ID = request.form("ID")
	    UserID = request.form("UserID")
	    Password = request.form("Password")
	    retrievedate = request.form("retrievedate")

	set conn = server.createobject("adodb.connection")
	DSNtemp="DRIVER={Microsoft Access Driver (*.mdb)}; "
        DSNtemp=dsntemp & "DBQ=" & server.mappath("reservations.mdb")

	Conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
           "Data Source=" & Server.MapPath("reservations.mdb")

		For each record in request("ID")
    		sqlstmt = "DELETE * from Reservations WHERE ID=" & record
			Set RS = conn.execute(sqlstmt)
		Next

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
			response.write "<font face='verdana' size=4><b>"
	    		response.write "The record has been deleted.</b></font>"
			response.write "<form action='reservationadmin.asp' method='post'>"
			response.write "<input type='hidden' name='flag' value='1'>"
			response.write "<input type='hidden' name='UserID' value='" & UserID& "'>"
			response.write "<input type='hidden' name='Password' value='" & Password & "'>"
			response.write "<input type='hidden' name='retrievedate' value='" & retrievedate & "'>"
			response.write "<input type='submit' value='Back to Administration List'>"
			response.write "</form>"
   		end if
End Select
set rs = nothing
set conn = nothing
%>
</table>
</center>
<br>
<font size="1">iReserve by <a href="http://www.IanThurston.com">IanThurston.com</a>

</body>
</html>
