<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../../GetSQLServerConnection.fun" -->
<!-- #include file="../../ADOINC.inc" -->
<!-- #include file="../../Include/func.inc" -->
<!--#include file="../../include/sideinfolev2.inc"-->

<%
Set conn = GetSQLServerConnection()
self="EMPWORKFORE1"


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
<link rel="stylesheet" href="../../Include/style.css" type="text/css">
<link rel="stylesheet" href="../../Include/style2.css" type="text/css">
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
<table width="100%" BORDER=0 align=center cellpadding=0 cellspacing=0 >
	<tr>
		<td align="center">
			<table id="myTableForm" width="50%"> 
				<tr><td height="35px" colspan=4>&nbsp;</td></tr>
				<tr height=30 >
					<TD nowrap align=right>查詢年月<br>Tìm kiếm ngày：</TD>
					<TD ><INPUT type="text" style="width:100px" NAME=YYMM   VALUE="<%=calcmonth%>" ></TD>
				
					<TD nowrap align=right height=30 >國籍<br>Quốc Gia：</TD>
					<TD >
						<select name=country   style='width:120px'  >
							<%if session("rights")<>"9" then %><option value="">全部<br>Toàn Bộ </option><%end if%>
							<%
							if session("rights")="9" then 
								SQL="SELECT * FROM BASICCODE WHERE FUNC='country'  ORDER BY SYS_type desc  "
							else
								SQL="SELECT * FROM BASICCODE WHERE FUNC='country'  ORDER BY SYS_type desc  "
							end if 	
							SET RST = CONN.EXECUTE(SQL)
							WHILE NOT RST.EOF
							%>
							<option value="<%=RST("SYS_TYPE")%>" <%if RST("SYS_TYPE")="VN" then%>selected<%end if%> ><%=RST("SYS_VALUE")%></option>
							<%
							RST.MOVENEXT
							WEND
							%>
						</SELECT>
						<%SET RST=NOTHING %>
					</TD>
				</tr>
				<tr>
					<TD nowrap align=right height=30 >廠別<br>Xưởng：</TD>
					<TD >
						<select name=WHSNO style="width:100px" >					
							<% 
							if Session("RIGHTS")="0" then%>
							<option value="">全部<br>Toàn Bộ </option>
							<% 
								SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' ORDER BY SYS_TYPE "
							else
								SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' and sys_type='"& Session("NETWHSNO") &"' ORDER BY SYS_TYPE "
							end if 	
							SET RST = CONN.EXECUTE(SQL)
							WHILE NOT RST.EOF
							%>
							<option value="<%=RST("SYS_TYPE")%>" <%if RST("SYS_TYPE")=session("mywhsno") then%>selected<%end if%>><%=RST("SYS_VALUE")%></option>
							<%
							RST.MOVENEXT
							WEND
							%>
						</SELECT>
						<%SET RST=NOTHING %>
					</TD>
					<TD nowrap align=right >組/部門<br>Tô/Bộ phận：</TD>
					<TD >
						<select name=GROUPID style="width:100px"  >
						<option value="" selected >全部<br>Toàn Bộ </option>
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
					<td nowrap align=right >員工編號<br>Số nhân viên：</td>
					<td >
						<input type="text" style="width:100px" name=empid1 maxlength=6 onchange=strchg(1)>
					</td>
					<td nowrap align=right >班別/<br>Ca：</td>
					<td  >
						<select name=shift  style="width:120px">
							<option value="">全部 Toàn Bộ</option>
							<option value="ALL">常日班 Ca bình thường</option>
							<option value="A">A班 Ca A</option>
							<option value="B">B班 Ca B</option>
							<option value="D">其他 Khác</option>
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
		<%=self%>.empid1.value = Ucase(<%=self%>.empid1.value)
	elseif a=2 then
		<%=self%>.empid2.value = Ucase(<%=self%>.empid2.value)
	end if
end function

function go()
 	<%=self%>.action="EMPWORK.FORE.asp"
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