<%@LANGUAGE=VBSCRIPT CODEPAGE=950%>
<!-- #include file = "../../GetSQLServerConnection.fun" -->
<!-- #include file="../../ADOINC.inc" -->
<!-- #include file="../../Include/func.inc" -->
<!--#include file="../../include/checkpower.asp"-->  
<%
Set conn = GetSQLServerConnection()
self="SALARYCP08"


nowmonth = year(date())&right("00"&month(date()),2)
if right("00"&month(date()),2)="01" then
	calcmonth = year(date())-1&"12"
else
	calcmonth =  year(date())&right("00"&month(date())-1,2)
end if

if day(date())<=11 then
	if right("00"&month(date()),2)="01" then
		calcmonth = year(date())-1&"12"
	else
		calcmonth =  year(date())&right("00"&month(date())-1,2)
	end if
else
	calcmonth = nowmonth
end if

if right(calcmonth,2)="01" then
	sgym = left(calcmonth,4)-1 & "12"
else
	sgym = left(calcmonth,4)&right("00"&right(calcmonth,2)-1,2)
end if
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
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
<form name="<%=self%>" method="post" action="<%=SELF%>.GETRPT.ASP">
<table width="460" border="0" cellspacing="0" cellpadding="0">
	<tr><TD>
	<img border="0" src="../../image/icon.gif" align="absmiddle">
	7.績效獎金明細表</TD></tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=500><BR>
<table width=500  ><tr><td >
	<table width=300 align=center border=0 cellspacing="2" cellpadding="2" class=txt  >
		<tr height=30 >
			<TD nowrap align=right>計薪年月</TD>
			<TD ><INPUT NAME=YYMM  CLASS=INPUTBOX VALUE="<%=calcmonth%>" SIZE=10></TD>
		</TR>
		<tr height=30 >
			<TD nowrap align=right>績效年月</TD>
			<TD ><INPUT NAME=JXYM  CLASS=INPUTBOX VALUE="<%=sgym%>" SIZE=10></TD>
		</TR>
		<TR>
		 	<TD nowrap align=right height=30 >國籍<BR>Quoc Tinh</TD>
			<TD >
				<select name=country  class=font9 style='width:75'  >
					<%if Session("NETWHSNO")="ALL" or Session("RIGHTS")<="1" or Session("RIGHTS")="8" then%>
					<option value="">全部 </option>
					<% 
						SQL="SELECT * FROM BASICCODE WHERE FUNC='country' ORDER BY SYS_TYPE "
					else
						SQL="SELECT * FROM BASICCODE WHERE FUNC='country' and sys_type='VN' ORDER BY SYS_TYPE "
					end if 
					'SQL="SELECT * FROM BASICCODE WHERE FUNC='country'  ORDER BY SYS_type desc  "
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
			<TD nowrap align=right height=30 >廠別<BR>Xuong</TD>
			<TD > 
				<select name=WHSNO  class=font9 >
					<% 
					if Session("RIGHTS")<="1" or Session("RIGHTS")="8" then%>
					<option value="">全部ALL</option>
					<% 
						SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' ORDER BY SYS_TYPE "
					else
						SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' and sys_type='"& Session("NETWHSNO") &"' ORDER BY SYS_TYPE "
					end if 
					'SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' ORDER BY SYS_TYPE "
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
		<TR height=30 >
			<TD nowrap align=right >部門<BR>Bo phan</TD>
			<TD >
				<select name=GROUPID  class=font9  >				
				<%if Session("RIGHTS")<="2" or Session("RIGHTS")>="8" then%>
					<option value="">全部ALL </option>
					<% 
						SQL="SELECT * FROM BASICCODE WHERE FUNC='GROUPID' and sys_type <>'AAA' ORDER BY SYS_TYPE " 
					else
						SQL="SELECT * FROM BASICCODE WHERE FUNC='GROUPID' and sys_type <>'AAA' and  sys_type= '"& session("NETG1") &"' ORDER BY SYS_TYPE "
					end if   
					
					SET RST = CONN.EXECUTE(SQL)
					RESPONSE.WRITE SQL 
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
		<tr height=30 >
			<TD nowrap align=right >單位<BR>Don vi</TD>
			<TD >
				<select name=zuno  class=font9  >
				<option value="">全部 </option>
				<%SQL="SELECT * FROM BASICCODE WHERE FUNC='zuno' and sys_type<>'XXX'  ORDER BY SYS_TYPE "
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
		</TR>
		<TR  height=30 >
			<td nowrap align=right >員工編號<BR>So the</td>
			<td colspan=3>
				<input name=empid1 class=inputbox size=15 maxlength=5 onchange=strchg(1)>

			</td>
		</TR>
		<TR  height=30 >
			<td nowrap align=right >班別<BR>Ca</td>
			<td colspan=3>
			 	<select name="shift" class=font9>
			 		<option value="">全部</option>
			 		<option value="ALL">常日班</option>
			 		<option value="A">A班</option>
			 		<option value="B">B班</option>
			 	</select>
			</td>
		</TR>
		<TR  height=30 >
			<td nowrap align=right >類別<BR>Loai</td>
			<td colspan=3>
			 	<select name="Loai" class=font9>
			 		<option value="">明細表(Detail)</option>
			 		<option value="T">統計表(Total)</option>			 		
			 	</select>
			</td>
		</TR>	
		<TR  height=30 >
			<td nowrap align=right >排序</td>
			<td colspan=3>
			 	<select name=sortby  class=font9>
			 		<option value="A">全部(ALL)</option>
			 		<option value="B">依工號(so the)</option>
			 	</select>
			</td>
		</TR>				
		<TR  height=30 >			
			<td colspan=4 align="center">
			 	<input name="func" type="checkbox" onclick="funcchg()">印已結帳績效獎金(In)
				<input name="op" value="" type="hidden" >
			</td>
		</TR>						
	</table><BR>
	<table width=450 align=center>
		<tr >
			<td align=center>
				<input type=button  name=btm class=button value="(Y)確認Confirm" onclick="go()" onkeydown="go()">
				<input type=reset  name=btm class=button value="(N)取消Cancel">
			</td>
		</tr>
	</table>

</td></tr></table>

</body>
</html>


<script language=vbs> 

function funcchg()
	if <%=self%>.func.checked=true   then 
		<%=self%>.op.value="Y" 
	else	
		<%=self%>.op.value="" 
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

 	<%=self%>.action="<%=self%>.getrpt.asp"
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