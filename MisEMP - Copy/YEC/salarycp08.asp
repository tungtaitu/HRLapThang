<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%> 
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->
<!-- #include file="../Include/func.inc" -->
<!--#include file="../include/checkpower.asp"-->  
<!-- #include file="../Include/SIDEINFO.inc" -->
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
<body  topmargin="60" leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload='f()' >
<form name="<%=self%>" method="post" action="<%=SELF%>.GETRPT.ASP"> 
<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>

<table width="100%" BORDER=0 align=center cellpadding=0 cellspacing=0 >
	<tr>
		<td align="center">
			<table id="myTableForm" width="50%"> 
				<tr><td height="35px" colspan=4>&nbsp;</td></tr>
				<tr >
					<TD nowrap align=right>計薪年月<br>Tien luong</TD>
					<TD ><INPUT NAME=YYMM  class="form-control form-control-sm mb-2 mt-2" VALUE="<%=calcmonth%>" SIZE=10></TD>		
					<TD nowrap align=right>績效年月<br>Tien Thuong</TD>
					<TD ><INPUT NAME=JXYM  class="form-control form-control-sm mb-2 mt-2" VALUE="<%=sgym%>" SIZE=10></TD>
				</TR>
				<TR>
					<TD nowrap align=right >國籍<BR>Quoc Tinh</TD>
					<TD >
						<select name=country  class="form-control form-control-sm mb-2 mt-2"  >
							<%if Session("NETWHSNO")="ALL" or Session("RIGHTS")<="1" or Session("RIGHTS")="8" then%>
							<option value="">--ALL--</option>
							<% 
								SQL="SELECT * FROM BASICCODE WHERE FUNC='country' ORDER BY SYS_TYPE "
							else
								SQL="SELECT * FROM BASICCODE WHERE FUNC='country' and sys_type='VN' ORDER BY SYS_TYPE "
							end if 
							'SQL="SELECT * FROM BASICCODE WHERE FUNC='country'  ORDER BY SYS_type desc  "
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
					</TD>		 
					<TD nowrap align=right >廠別<BR>Xuong</TD>
					<TD > 
						<select name=WHSNO  class="form-control form-control-sm mb-2 mt-2" >
							<% 
							if Session("RIGHTS")<="1" or Session("RIGHTS")="8" then%>
							<option value="">--ALL--</option>
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
				<TR >
					<TD nowrap align=right >部門<BR>Bo phan</TD>
					<TD >
						<select name=GROUPID  class="form-control form-control-sm mb-2 mt-2"  >				
						<%if Session("RIGHTS")<="2" or Session("RIGHTS")>="8" then%>
							<option value="">--ALL--</option>
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
						<%SET RST=NOTHING 
						
						%> 
						<input type="hidden" name="zuno" >
					</td>
					<td nowrap align=right >員工編號<BR>So the</td>
					<td>
						<input name=empid1 class="form-control form-control-sm mb-2 mt-2" size=15 maxlength=5 onchange=strchg(1)>
					</td>
				</TR>
				<TR>
					<td nowrap align=right >班別<BR>Ca</td>
					<td>
						<select name="shift" class="form-control form-control-sm mb-2 mt-2">
						<option value="">--ALL--</option>
						<%SQL="SELECT * FROM BASICCODE WHERE FUNC='shift' and sys_type<>'XXX'  ORDER BY len(sys_type) desc, SYS_TYPE "
						SET RST = CONN.EXECUTE(SQL)
						WHILE NOT RST.EOF
						%>
						<option value="<%=RST("SYS_TYPE")%>" ><%=RST("SYS_VALUE")%></option>
						<%
						RST.MOVENEXT
						WEND
						%>
						</SELECT>
						<%SET RST=NOTHING 
						conn.close 
						set conn=nothing
						%>
					</td>
					<td nowrap align=right >類別<BR>Loai</td>
					<td >
						<select name="Loai" class="form-control form-control-sm mb-2 mt-2">
							<option value="">明細表(Detail)</option>
							<option value="T">統計表(Total)</option>			 		
						</select>
					</td>		
				</TR>
				<TR  height=30 >
					<td nowrap align=right >排序</td>
					<td >
						<select name=sortby  class="form-control form-control-sm mb-2 mt-2">
							<option value="A">全部(ALL)</option>
							<option value="B">依工號(so the)</option>
						</select>
					</td>			
					<td align="right">
						<input name="func" type="checkbox" onclick="funcchg()">
						<input name="op" value="" type="hidden" >
					</td>
					<td nowrap>
						印已結帳績效獎金(In)
					</td>
				</TR>
				<tr >
					<td align=center colspan=4 height="50px">
						<input type=button  name=btm class="btn btn-sm btn-danger" value="(Y)確認Confirm" onclick="go()" onkeydown="go()">
						<input type=reset  name=btm class="btn btn-sm btn-outline-secondary" value="(N)取消Cancel">
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>

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

 	<%=self%>.action="<%=session("rpt")%>"&"rpt/"&"<%=self%>.getrpt.asp"
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