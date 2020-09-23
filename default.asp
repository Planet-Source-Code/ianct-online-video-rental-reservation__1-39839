<%
Response.Expires = 0
%>
<HTML>
<TITLE>Video West Reservations Online</TITLE>

<Script language="javascript">

function ShowDate(oDoc,cDate)
{
window.open("datepicker.asp?backf="+oDoc+"&cDate="+cDate,"window", "height=210, width=300, resizeable=no,","");
}

</Script>

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

<script language="JAVASCRIPT">

function submitIt()
{
  var strCustID = document.MainForm.CustID.value;
  var strPhone = document.MainForm.CustPhone.value;
  var strReserveTitle = document.MainForm.ReserveTitle.value;

 if (strCustID == "") 
  {
    alert( "Customer ID cannot be blank." );
  }
  else
 if (strPhone == "")
  {
    alert( "Customer Phone cannot be blank." );
    end
  }
  else
 if (strReserveTitle == "")
  {
    alert( "Reserve Title cannot be blank." );
    end
  }
  else
  {
    document.MainForm.submit();
  }
}
</script>


<body bgcolor="#FFFFFF">
Video West Online Reservations<br>
<form action="authorize.asp" method="post" name="MainForm">
<TABLE cellpadding=1 cellspacing=1 bgcolor="#000000">
 <TR class="MENU">
 <TD colspan="2">
 Date to Reserve
 </TD>
 </TR>
 <TR class="ITEM">
 <TD>
 <input type="text" onClick="javascript:ShowDate('MainForm.ReserveDate',MainForm.ReserveDate.value)" class="INPUTTEXT" name="ReserveDate" size="10" STYLE="font-size:8pt; font-family:Verdana" id="VonDatum" size=10 Value=<%Response.write DateAdd("d","+1",Date)%> maxlength=12>
<!-- <A href="javascript:ShowDate('MainForm.ReserveDate',MainForm.ReserveDate.value)">pick</A>-->
 </TD>
 <TD>
 Reserve date<br><b>cannot</b> be today
 </TD>
 </TR>

 <TR class="MENU">
 <TD> <% if strFlag = 1 then response.write("*") %>
 Customer ID
 </TD>
 <TD>
 Customer Phone
 </TD>
</TR>
<TR class="ITEM">
 <TD>
 <input name="CustID" size="6" maxlength="10" STYLE="font-size:8pt; font-family:Verdana">
 </TD>
 <TD>
 <input name="CustPhone" size="12" maxlength="12" STYLE="font-size:8pt; font-family:Verdana">
 </TD>
</TR>
<TR class="MENU">
 <TD>
 Title Name
 </TD>
 <TD>
 Format
 </TD>
</TR>
<TR class="ITEM">
 <TD>
 <input name="ReserveTitle" STYLE="font-size:8pt; font-family:Verdana">
 </TD>
 <TD>
 <SELECT name="ReserveFormat" STYLE="font-size:8pt; font-family:Verdana">
 <OPTION>VHS
 <OPTION>DVD
 </SELECT>
 </TD>
</TR>
<TR class="ITEM">
 <TD colspan=2 align="center">
 <input type="hidden" name="flag" value=1>
 <input type="Button" onClick="javascript:submitIt();" value="Reserve!" STYLE="font-size:8pt; font-family:Verdana">
 <input type="Reset" value="Clear" STYLE="font-size:8pt; font-family:Verdana">
 </TD>
 </TR>
</TABLE>
</FORM>
<br>
<font size=1>iReserve by <a href="http://www.IanThurston.com">IanThurston.com</a></font><br>
</BODY>
</HTML>
