<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<!-- #include file="../Include/func.inc" -->
<!--#include file="../include/sideinfo.inc"-->
<%
Set conn = GetSQLServerConnection()
self="YEDQ01"


nowmonth = year(date())&right("00"&month(date()),2)
if month(date())="01" then
	'calcmonth = year(date()-1)&"12"
	calcmonth = nowmonth
else
	'calcmonth =  year(date())&right("00"&month(date())-1,2)
	calcmonth = nowmonth 
end if

if day(date())<=11 then
	if month(date())="01" then
		calcmonth = year(date()-1)&"12"
		calcmonth = nowmonth 
	else
		calcmonth =  year(date())&right("00"&month(date())-1,2)
		calcmonth = nowmonth 
	end if
else
	calcmonth = nowmonth
end if
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
function f()
	<%=self%>.YYMM.focus()
	<%=self%>.YYMM.SELECT()
end function
-->
</SCRIPT>
</head>
<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload='f()' >
<form name="<%=self%>" method="post" action="EMPWORK.FORE.ASP">
<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>

<table width="100%"  cellpadding=0 cellspacing=0><tr><td align="center">
	<table id="myTableForm" width="40%">
		<tr><td colspan=2 height="35px">&nbsp;</td></tr>
		<tr>
			<TD nowrap align=right style="width:20%">查詢年月：</TD>
			<TD style="width:80%"><INPUT  type="text" style="width:100px" NAME=YYMM   VALUE="<%=calcmonth%>"></TD>
		</TR>		 
		<TR>
			<td nowrap align=right ><a href="vbscript:gotemp()"><font color=blue><u>員工編號：</u></font></a></td>
			<td>
				<input  type="text" style="width:48%" name=empid  maxlength=5 onchange=strchg(1)>
				<input  type="text" style="width:48%" name=empname maxlength=5 > 				
			</td>
		</TR> 	
		<tr >
			<td align=center colspan=2 height="50px">
				<input type=button  name=btm class="btn btn-sm btn-danger" value="確   認" onclick="go()" onkeydown="go()">
				<input type=reset  name=btm class="btn btn-sm btn-outline-secondary" value="取   消">
			</td>
		</tr>
	</table>
</td></tr></table>

</body>
</html>


<script language=vbs> 
function gotemp()
	open "../getempdata.asp?formName="&"<%=self%>", "Back"
	parent.best.cols="50%,50%"
end function 
function dataclick(a)
	if a = 1 then
		open "empbasic/empbasic.asp" , "_self"
	elseif a = 2 then
		open "empfile/empfile.asp" , "_self"
	elseif a = 3 then
		open "empworkHour/empwork.asp" , "_self"
	elseif a = 4 then
		open "holiday/empholiday.asp" , "_self"
	elseif a = 5 then
		open "AcceptCaTime/main.asp" , "_self"
	elseif a = 6 then
		open "../report/main.asp" , "_self"
	end if
end function

function strchg(a)
	if a=1 then
		<%=self%>.empid.value = Ucase(<%=self%>.empid.value)
	elseif a=2 then
		<%=self%>.empid.value = Ucase(<%=self%>.empid.value)
	end if
end function

function go()
 	<%=self%>.action="<%=self%>.ForeGnd.asp"
 	<%=self%>.submit()
end function


'*******檢查日期*********************************************
FUNCTION date_change(a)

if a=1 then
	INcardat = Trim(<%=self%>.indat1.value)
elseif a=2 then
	INcardat = Trim(<%=self%>.indat2.value)
end if

IF INcardat<>"" THEN
	ANS=validDate(INcardat)
	IF ANS <> "" THEN
		if a=1 then
			Document.<%=self%>.indat1.value=ANS
		elseif a=2 then
			Document.<%=self%>.indat2.value=ANS
		end if
	ELSE
		ALERT "EZ0067:輸入日期不合法 !!"
		if a=1 then
			Document.<%=self%>.indat1.value=""
			Document.<%=self%>.indat1.focus()
		elseif a=2 then
			Document.<%=self%>.indat2.value=""
			Document.<%=self%>.indat2.focus()
		end if
		EXIT FUNCTION
	END IF

ELSE
	'alert "EZ0015:日期該欄位必須輸入資料 !!"
	EXIT FUNCTION
END IF

END FUNCTION
</script> 