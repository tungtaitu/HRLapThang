<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<!-- #include file="../Include/func.inc" -->
<!--#include file="../include/sideinfo.inc"-->
<%
Set conn = GetSQLServerConnection()
self="YEDE02"


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
	<%=self%>.D1.focus()
	<%=self%>.D1.SELECT()
end function
-->
</SCRIPT>
</head>
<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload='f()' >
<form name="<%=self%>" method="post" action="EMPWORK.FORE.ASP">
<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>
<table width="100%" BORDER=0 align=center cellpadding=0 cellspacing=0 >
	<tr>
		<td align="center">
			<table id="myTableForm" width="50%"> 
				<tr><td height="35px" colspan=4>&nbsp;</td></tr>
				<tr height=30 >
					<TD nowrap align=right>日期：</TD>
					<TD >
						<input type="text" style="width:100px" name=D1  maxlength=10 onblur="date_change(1)" value="<%=DD2%>">~
						<input type="text" style="width:100px" name=D2  maxlength=10 onblur="date_change(2)" value="<%=DD2%>">
					</TD>
					<TD nowrap align=right height=30 >國籍：</TD>
					<TD >
						<select name=F_country   style='width:120px'  >
							
							<%
							if Session("RIGHTS") <="1" then  					
							%><option value="">全部 </option>
							<%
								SQL="SELECT * FROM BASICCODE WHERE FUNC='country'  ORDER BY SYS_type desc  "
							else
								SQL="SELECT * FROM BASICCODE WHERE FUNC='country' and sys_type='VN'  ORDER BY SYS_type desc  "
							end if 	
							SET RST = CONN.EXECUTE(SQL)
							WHILE NOT RST.EOF
							%>
							<option value="<%=RST("SYS_TYPE")%>" ><%=RST("SYS_VALUE")%></option>
							<%
							RST.MOVENEXT
							WEND
							%>
						</SELECT>
						<%SET RST=NOTHING %>
					</TD>
				</tr>
				<tr>
					<TD nowrap align=right height=30 >廠別：</TD>
					<TD >
						<select name=F_WHSNO   >					
							<% 
							if Session("RIGHTS")="0" then%>
							<option value="">全部 </option>
							<% 
								SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' ORDER BY SYS_TYPE "
							else
								SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' and sys_type='"& Session("NETWHSNO") &"' ORDER BY SYS_TYPE "
							end if 	
							SET RST = CONN.EXECUTE(SQL)
							WHILE NOT RST.EOF
							%>
							<option value="<%=RST("SYS_TYPE")%>"><%=RST("SYS_VALUE")%></option>
							<%
							RST.MOVENEXT
							WEND
							%>
						</SELECT>
						<%SET RST=NOTHING %>
					</TD>
					<TD nowrap align=right >組/部門：</TD>
					<TD >
						<select name=F_GROUPID    >
						<option value="" selected >全部 </option>
						<%
						SQL="SELECT * FROM BASICCODE WHERE FUNC='GROUPID' and sys_type <>'AAA' ORDER BY SYS_TYPE "
						'SQL="SELECT * FROM BASICCODE WHERE FUNC='GROUPID'  ORDER BY SYS_TYPE "
						SET RST = CONN.EXECUTE(SQL)
						'RESPONSE.WRITE SQL
						WHILE NOT RST.EOF
						%>
						<option value="<%=RST("SYS_TYPE")%>" ><%=RST("SYS_TYPE")%><%=RST("SYS_VALUE")%></option>
						<%
						RST.MOVENEXT
						WEND
						%>
						</SELECT>
						<%SET RST=NOTHING %>
					</td>
				</tr>
				<TR  height=30 >
					<td nowrap align=right >員工編號：</td>
					<td >
						<input type="text" style="width:100px" name=F_empid  maxlength=6 onchange=strchg(1)>

					</td>
					<td nowrap align=right >班別：</td>
					<td  >
						<select name=F_shift >
							<option value="">全部</option>
							<option value="ALL">常日班</option>
							<option value="A">A班</option>
							<option value="B">B班</option>		
							<option value="N">ca diem夜班</option>					
							<option value="M">ca diem < 7H 夜班 </option>							
						</select>
					</td>
				</TR>
				<tr >
					<td align=center colspan=4 height="50px">
						<input type=button  name=btm class="btn btn-sm btn-danger" value="確   認" onclick="go()" onkeydown="go()">
						<input type=reset  name=btm class="btn btn-sm btn-outline-secondary" value="取   消">
					</td>
				</tr>
			</table>
		</td></tr></table>

</body>
</html>


<script language=vbs>
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
		<%=self%>.F_empid.value = Ucase(<%=self%>.F_empid.value)
	elseif a=2 then
		<%=self%>.empid2.value = Ucase(<%=self%>.empid2.value)
	end if
end function

function go()
 	<%=self%>.action="<%=self%>.ForeGnd.asp"
 	<%=self%>.submit()
end function


'*******檢查日期*********************************************
FUNCTION date_change(a)

if a=1 then
	INcardat = Trim(<%=self%>.D1.value)
elseif a=2 then
	INcardat = Trim(<%=self%>.D2.value)
end if

IF INcardat<>"" THEN
	ANS=validDate(INcardat)
	IF ANS <> "" THEN
		if a=1 then
			Document.<%=self%>.D1.value=ANS
		elseif a=2 then
			Document.<%=self%>.D2.value=ANS
		end if
	ELSE
		ALERT "EZ0067:輸入日期不合法 !!"
		if a=1 then
			Document.<%=self%>.D1.value=""
			Document.<%=self%>.D1.focus()
		elseif a=2 then
			Document.<%=self%>.D2.value=""
			Document.<%=self%>.D2.focus()
		end if
		EXIT FUNCTION
	END IF

ELSE
	'alert "EZ0015:日期該欄位必須輸入資料 !!"
	EXIT FUNCTION
END IF

END FUNCTION
</script> 