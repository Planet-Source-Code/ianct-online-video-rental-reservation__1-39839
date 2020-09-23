<%
Response.Expires = 0

if len(trim(Request.QueryString("cDate"))) > 0 and isDate(Request.QueryString("cDate")) = True then
curDate = cdate(Request.QueryString("cDate"))
curMonth = month(curDate)
curYear = year(curDate)

else

if len(trim(curMonth)) < 1 then
curMonth = Request("cMonth")
end if
if len(trim(curMonth)) < 1 then
curMonth = month(date)
end if

curYear = Request("cYear")
if len(trim(curYear)) < 1 then
curYear = year(date)
end if

'curDate = cdate(day(date) & "." & cdbl(curMonth) & "." & cdbl(curYear))
curDate = cdate(day(date) & "/" & cdbl(curMonth) & "/" & cdbl(curYear))

end if

if weekday(dateadd("d",((day(curDate)-1)* -1),curdate),2) = 1 then
Firstday = dateadd("d",((day(curDate)-1 + 7)* -1),curdate)
else
Firstday = dateadd("d",((day(curDate)-1 + weekday(dateadd("d",((day(curDate))* -1),curdate),2)) * -1),curdate)
end if

curWeek = DatePart("ww", Firstday, 2, vbFirstFourDays)

%>
<HTML>
<Head>
<Title>Reservation Date</Title>
<%
RecObj = Request.QueryString("backf")
Response.Write "<Script language='javascript'> "
Response.Write "function SetDate(aDate) "
Response.Write "{ "
Response.Write "window.opener." & RecObj & ".value = aDate; "
Response.Write "window.close(); "
Response.Write "} "
Response.Write "</Script> "
%>
<Script language="javascript">

window.moveTo((window.screen.width - 285)/2,(window.screen.height - 200)/2);

function cib(o,Nr)
{
if (Nr == "on")
{
o.style.background = "#F7BF47";
}
else
{
o.style.background = "#D5D1C8";
}
}

</Script>

<Style type="text/css">
 TD {Font-family:tahoma, Arial, Verdana; font-weight:400; font-size: 8pt; color:#000000;}
 .INPUTComb {Font-family:tahoma, Arial, Verdana; font-weight:400; font-size: 7pt; color:#000000;}
 .CalDay {Font-family:tahoma, Arial, Verdana; font-weight:600; font-size: 8pt; color:#0000a0; text-align: left; width:35px; height:25px; background-color: #D5D1C8; border-left: 1px #ffffff solid; border-top: 1px #ffffff solid; border-right: 1px #000000 solid; border-bottom: 1px #000000 solid; cursor:hand}
 .INPUTBUTTON {Font-family:Tahoma, Verdana, Arial; font-weight:400; font-size: 8pt; color:#0000a0; background-color: #D5D1C8; border-left: 1px #ffffff solid; border-top: 1px #ffffff solid; border-right: 1px #000000 solid; border-bottom: 1px #000000 solid; cursor:hand}
 A:link {Font-family: Tahoma, Arial, Verdana; Font-size: 12pt; Font-weight: 600; color: #000000; text-decoration: none;}
 A:active {Font-family: Tahoma, Arial, Verdana; Font-size: 12pt; Font-weight: 600; color: #000000; text-decoration: none;}
 A:visited {Font-family: Tahoma, Arial, Verdana; Font-size: 12pt; Font-weight: 600; color: #000000; text-decoration: none;}
 A:hover {Font-family: Tahoma, Arial, Verdana; Font-size: 12pt; Font-weight: 600; color: #00cc00; text-decoration: none;}
</Style>

</Head>

<Body onLoad="javascript:window.focus()" bgcolor="#e9e9e9" leftmargin=5 topmargin=2 rightmargin=5 bottommargin=5>

<Form method="Post" action="datepicker.asp?backf=<%Response.Write RecObj%>" name="CalForm">

<Table border="0" width="100%" height="30" cellpadding="0" cellspacing="0">
<TR>
<TD width="20"><Input type="Button" class="INPUTBUTTON" onClick="javascript:window.document.location.href='datepicker.asp?backf=<%Response.Write RecObj%>&cDate=<%Response.write dateadd("m", -12, curDate)%>'" Value="<<" id=Button1 name=Button1></TD>
<TD width="20"><Input type="Button" class="INPUTBUTTON" onClick="javascript:window.document.location.href='datepicker.asp?backf=<%Response.Write RecObj%>&cDate=<%Response.write dateadd("m", -1, curDate)%>'" Value="<" id=Button1b name=Button1b></TD>
<TD width="80" align=center><Font size=2><B><%Response.write monthname(curMonth,2)%>&nbsp;<%Response.write curYear%></B></Font></TD>
<TD align=right><SELECT class="INPUTComb" onChange="window.document.CalForm.submit()" id="cMonth" name="cMonth">
<%For i = 1 to 12%>
<OPTION Value="<%Response.write i%>" <%if i = cdbl(curmonth) then response.write "selected" end if%>><%Response.write monthname(i)%></OPTION>
<%Next%>
</SELECT>&nbsp;
<SELECT class="INPUTComb" onChange="window.document.CalForm.submit()" id="cYear" name="cYear">
<%For i = 2050 to 1900 step -1%>
<OPTION Value="<%Response.write i%>" <%if i = cdbl(curYear) then response.write "selected" end if%>><%Response.write i%></OPTION>
<%Next%>
</SELECT></TD>
<TD width="20" align=right>&nbsp;<Input type="Button" class="INPUTBUTTON" onClick="javascript:window.document.location.href='datepicker.asp?backf=<%Response.Write RecObj%>&cDate=<%Response.write dateadd("m", 1, curDate)%>'" Value=">" id=Button2 name=Button2></TD>
<TD width="20"><Input type="Button" class="INPUTBUTTON" onClick="javascript:window.document.location.href='datepicker.asp?backf=<%Response.Write RecObj%>&cDate=<%Response.write dateadd("m", 12, curDate)%>'" Value=">>" id=Button2b name=Button2b></TD>
</TR>
</Table>

<Table border="0" width="280" cellpadding="0" cellspacing="0">
<TR>
<TD>&nbsp;<SELECT class="INPUTComb" onChange="javascript:window.document.location.href='datepicker.asp?backf=<%Response.Write RecObj%>&cDate='+window.CalForm.cWeek.value" id="cWeek" name="cWeek">
<%For i = 1 to 52%>
<OPTION Value="<%if curWeek = 52 then Response.write dateadd("ww", i - 1, Firstday) else Response.write dateadd("ww", i - curWeek, Firstday) end if%>" <%If i = curWeek then Response.write "selected" end if%>><%Response.write i%></OPTION>
<%Next%>
</SELECT></TD>
<%For i = 1 to 7%>
<TD width=35 align=center><%if i = 7 then Response.write "<Font color='#cc0000'><B>" end if%><%if i = 6 then Response.write "<Font color='#ff0000'><B>" end if%><%Response.Write weekdayname(i,2,2)%><%if i = 6 then Response.write "</B></Font>" end if%><%if i = 7 then Response.write "</B></Font>" end if%></TD>
<%next%>
</TR>
</Table>

<Table border="0" width="280" cellpadding="0" cellspacing="0">
<TR>
<TD width=50 align=center><%Response.Write DatePart("ww", FirstDay, vbMonday, vbFirstFourDays)%></TD>
<%For i = 0 to 41%>
<%If (i = 7 or i = 14 or i = 21 or i = 28 or i = 35) then%></TR><TR><TD width=50 align=center><%Response.Write DatePart("ww", dateadd("d", i, FirstDay), vbMonday, vbFirstFourDays)%></TD><%end if%>
<TD width=35 align=center><Input type="Button" <%If month(dateadd("d", i, FirstDay)) <> month(curDate) then Response.Write "style='color: #808080; font-weight:400;'" end if%> <%If dateadd("d", i, FirstDay) = Date then Response.write "style='font-size=10pt; border-left: #000000 1px solid; border-top: #000000 1px solid; border-right: #c0c0c0 1px solid; border-bottom: #c0c0c0 1px solid; '" end if%> class="CalDay" onmouseover="cib(this,'on')" onmouseout="cib(this,'off')" onClick="javascript:SetDate('<%Response.Write dateadd("d", i, FirstDay)%>')" Value="&nbsp;<%Response.write day(dateadd("d", i, FirstDay))%>"></TD>
<%Next%>
</TR>
</Table>

</Form>

</Body>
</HTML>

