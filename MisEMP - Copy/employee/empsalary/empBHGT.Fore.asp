<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../../GetSQLServerConnection.fun" -->
<!-- #include file="../../ADOINC.inc" -->
<!-- #include file="../../Include/func.inc" -->
<%
Set conn = GetSQLServerConnection()
self="EMPBHGT"

nowmonth = year(date())&right("00"&month(date()),2)  
if month(date())="1" then  
	calcmonth = year(date())-1&"12"  	
else
	calcmonth =  year(date())&right("00"&month(date())-1,2)   
end if 	   

if day(date())<=11 then 
	if month(date())="1" then  
		calcmonth = year(date())-1&"12" 
	else
		calcmonth =  year(date())&right("00"&month(date())-1,2)   
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
<form name="<%=self%>" method="post" action="EMPFILE.SALARY.ASP">
<table width="460" border="0" cellspacing="0" cellpadding="0">
	<tr><TD>
	<img border="0" src="../../image/icon.gif" align="absmiddle">
	員工保險與工團</TD></tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=500>
<BR><BR>
<table width=500  ><tr><td >
	<table width=400 align=center border=0 cellspacing="2" cellpadding="2"  >
		<tr   >
			<TD nowrap align=right>計薪年月<br><font class=txt8>Tháng Năm</font></TD>
			<TD ><INPUT NAME=YYMM  CLASS=INPUTBOX VALUE="<%=calcmonth%>" SIZE=10></TD>
		</TR>
		<TR>
		 	<TD nowrap align=right >國籍<br><font class=txt8>Quốc tịch</font></TD>
			<TD >
				<select name=country  class=font9   >
					<%SQL="SELECT * FROM BASICCODE WHERE FUNC='country' and sys_type='VN' ORDER BY SYS_type desc  "
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
			<TD nowrap align=right   >廠別<br><font class=txt8>Xưởng</font></TD>
			<TD >
				<select name=WHSNO  class=font9 >
					<option value="">全部(Toan bo) </option>
					<%SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' ORDER BY SYS_TYPE "
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
		</TR>
		<tr>
		<TR   >
			<TD nowrap align=right >組/部門<br><font class=txt8>Bộ phận</font></TD>
			<TD >
				<select name=GROUPID  class=font9  >
				<option value="" selected >全部(Toan bo) </option>
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
		</tr>
		<!--tr height=30 >
			<TD nowrap align=right >職等：</TD>
			<TD >
				<select name=JOB  class=font9  >
				<option value="">全部 </option>
				<%SQL="SELECT * FROM BASICCODE WHERE FUNC='LEV'  ORDER BY SYS_TYPE "
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
		</TR-->
		<TR   >
			<td nowrap align=right >員工編號<br><font class=txt8>So The</font></td>
			<td colspan=3>
				<input name=empid1 class=inputbox size=15 maxlength=5 onchange=strchg(1)>

			</td>
		</TR>
		<TR   >
			<td nowrap align=right >員工統計<br><font class=txt8>Thong ke Nhan vien</font></td>
			<td colspan=3>
			 	<select name=outemp class=font9>
			 		<option value="">全部(Toan bo)</option>
			 		<option value="N">本月在職</option>
			 		<option value="D">本月離職</option>
			 	</select>
			</td>
		</TR>
		<TR   >
			<td nowrap align=right >班別<br><font class=txt8>Ca</font></td>
			<td colspan=3>
			 	<select name=shift class=font9>
			 	<option value="">全部(Toan bo)</option>
				<%
				SQL="SELECT * FROM BASICCODE WHERE FUNC='shift'   ORDER BY SYS_TYPE "				
				SET RST = CONN.EXECUTE(SQL) 
				WHILE NOT RST.EOF
				%>
				<option value="<%=RST("SYS_TYPE")%>" ><%=RST("SYS_TYPE")%>-<%=RST("SYS_VALUE")%></option>
				<%
				RST.MOVENEXT
				WEND
				%>
				</SELECT>
				<%SET RST=NOTHING %>
			 	</select>
			</td>
		</TR>
		<!--TR   >
			<td colspan=2 align=center><input type=checkbox name="chk1"  onclick="recalchg()">本月重新計算
			<input type="hidden"  name="recalc" value="N" size=1>
			</td>
		</tr-->
	</table><BR>
	<table width=450 align=center>
		<tr >
			<td align=center>
				<input type=button  name=btm class=button value="確   定"  type="hiddent" onclick="go()" onkeydown="go()">
				<input type=reset  name=btm class=button value="取   消">
			</td>
		</tr>
	</table>

</td></tr></table>

</body>
</html>
<script language=vbs>  
function recalchg()
	if <%=self%>.chk1.checked= true then 
		<%=self%>.recalc.value="Y"
	elseif <%=self%>.chk1.checked= false then 
		<%=self%>.recalc.value="N"
	end if 
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
		<%=self%>.empid1.value = Ucase(<%=self%>.empid1.value)
	elseif a=2 then
		<%=self%>.empid2.value = Ucase(<%=self%>.empid2.value)
	end if
end function

function go()
 	'if <%=self%>.country.value="TA" then
 	'	<%=self%>.action="EMPFILE.VNSALARY.asp"
 	'	<%=self%>.submit()
 	'else
 		<%=self%>.action="EMPBHGT.ForeGnd.asp"
 		<%=self%>.submit()
 	'end if
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