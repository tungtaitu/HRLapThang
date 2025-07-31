<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>

<SCRIPT LANGUAGE=javascript>
var sMon = new Array(12);
	sMon[0] = "Jan"
	sMon[1] = "Feb"
	sMon[2] = "Mar"
	sMon[3] = "Apr"
	sMon[4] = "May"
	sMon[5] = "Jun"
	sMon[6] = "Jul"
	sMon[7] = "Aug"
	sMon[8] = "Sep"
	sMon[9] = "Oct"
	sMon[10] = "Nov"
	sMon[11] = "Dec"

function calendar(t) {
	var sPath = "calendar1.htm";
	strFeatures = "dialogWidth=206px;dialogHeight=228px;center=yes;help=no";
	st = t.value;
	sDate = showModalDialog(sPath,st,strFeatures);
	t.value = formatDate(sDate, 0);
	
}

function checkDate(t) {
	dDate = new Date(t.value);
	if (dDate == "NaN") {t.value = ""; return;}

	iYear = dDate.getFullYear()

	if ((iYear > 1899)&&(iYear < 1950)) {

		sYear = "" + iYear + ""
		if (t.value.indexOf(sYear,1) == -1) {
			iYear += 100
			sDate = (dDate.getMonth() + 1) + "/" + dDate.getDate() + "/" + iYear
			dDate = new Date(sDate)
		}
	}
	t.value = formatDate(dDate);
}

function formatDate(sDate) {
	var sScrap = "";
	var dScrap = new Date(sDate);
	if (dScrap == "NaN") return sScrap;
	
	iDay = dScrap.getDate();
	iMon = dScrap.getMonth();
	iYea = dScrap.getFullYear();
 if ((iMon+1)<=9 )
  {sMon="0"+(iMon+1) }
  else
 {sMon=iMon+1 }
  if (iDay<=9 )
  {sDay="0"+iDay }
  else
  {sDay=iDay }
	sScrap = iYea + "/" + sMon + "/" + sDay ;
	return sScrap;
}
</SCRIPT> 
<!-- #include file = "../../GetSQLServerConnection.fun" --> 
<!-- #include file="../../ADOINC.inc" -->
<%
'on error resume next   
session.codepage="65001"
SELF = "catime"
 
Set conn = GetSQLServerConnection()	  
Set rs = Server.CreateObject("ADODB.Recordset")   

DAT1 = REQUEST("DAT1")
DAT2 = REQUEST("DAT2")

IF DAT1="" THEN DAT1=DATE()-1
IF DAT2="" THEN DAT2=DATE()-1


FUNCTION FDT(D)
	Response.Write YEAR(D)&"/"&RIGHT("00"&MONTH(D),2)&"/"&RIGHT("00"&DAY(D),2) 
	
END FUNCTION 

nowmonth = year(date())&right("00"&month(date()),2)  
if month(date())="01" then  
	calcmonth = year(date()-1)&"12" 
else
	calcmonth =  year(date())&right("00"&month(date())-1,2)   
end if 	 

%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
<link rel="stylesheet" href="../../Include/style.css" type="text/css">
<link rel="stylesheet" href="../../Include/style2.css" type="text/css"> 
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
'-----------------enter to next field
function enterto()
	if window.event.keyCode = 13 then window.event.keyCode =9 		
end function  

function f()
	<%=self%>.DAT1.focus()	
	<%=self%>.DAT1.SELECT()
end function   


-->
</SCRIPT>  


</head>   
<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload="f()">
<form name="<%=self%>" method="post" action="acceptedcatime.asp">
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
	
<table width="460" border="0" cellspacing="0" cellpadding="0">
	<tr><TD >
	<img border="0" src="../../image/icon.gif" align="absmiddle">
	轉入差勤資料 </TD></tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=500>		 
<br><br>	
<TABLE WIDTH=460 BORDER=0>   
	<TR height=25 >
		<TD nowrap align=right WIDTH=150>接收日期：</TD>
		<TD   >
			<INPUT NAME=DAT1 SIZE=12 CLASS=INPUTBOX  onDblClick="calendar(this)" VALUE="<%=FDT(DAT1)%>"> ~ 	
			<INPUT NAME=DAT2 SIZE=12 CLASS=INPUTBOX  onDblClick="calendar(this)" VALUE="<%=FDT(DAT2)%>">
		</TD> 
	</TR>	 	 
	<TR>
		<td COLSPAN=2 ALIGN=CENTER HEIGHT=50>
			<input type="button" name="send" value="確　定" class=button  onclick="go()" >
			<input type="RESET"  name="send" value="取　消" class=button >	
		</td>	
	</TR>
</table>	 
<BR>  
</form>

</body>
</html>

<script language=vbscript>
function BACKMAIN()	
	open "../main.asp" , "_self"
end function     

function go()
	<%=self%>.action="forwait_tmp2.asp"
	<%=self%>.target="Fore"
	<%=self%>.submit()
end function 
</script>

