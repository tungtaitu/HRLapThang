<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<!-- #include file="../Include/func.inc" -->
<!--#include file="../include/sideinfo.inc"-->
<%
Set conn = GetSQLServerConnection()
self="yeDp04"


 
w1=session("mywhsno")

indat1 = request("indat1")
indat2 = request("indat2")
country = request("country")
WHSNO = request("WHSNO")
GROUPID = request("GROUPID")
shift = request("shift")
T1 = request("T1")
eid = request("eid") 

if request("indat1")="" then indat1=DD2
if request("indat2")="" then indat2=DD2
if request("country")="" then country="VN"
if request("WHSNO")="" then WHSNO=session("mywhsno")
w1 = session("mywhsno")
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
	<%=self%>.indat1.focus()
end function
-->
</SCRIPT>
</head>
<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload='f()' >
<form name="<%=self%>" method="post" >
<input type="hidden" name="netuser"  value="<%=session("netuser")%>">

<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>
<BR><BR>
<table width="100%" ><tr><td align="center">
	<table id="myTableForm" width="50%"> 
		<tr><td height="35px" colspan=4>&nbsp;</td></tr>
		<TR height=30>
			<td nowrap align=right >日期範圍<br>Ngay</td>
			<td nowrap>
				<input type="text" style="width:45%" name=indat1 maxlength=10 onblur="date_change(1)" value="<%=indat1%>">~
				<input type="text" style="width:45%" name=indat2 maxlength=10 onblur="date_change(2)" value="<%=indat2%>">
			</td>
		 	<TD nowrap align=right>國籍<br>Quoc Tich</TD> 
			<TD >
				<select name=country  style="width:120px">
					<option value="">--ALL--- </option>
					<%SQL="SELECT * FROM BASICCODE WHERE FUNC='country'  ORDER BY SYS_type desc  "
					SET RST = CONN.EXECUTE(SQL)
					WHILE NOT RST.EOF
					%>
					<option value="<%=RST("SYS_TYPE")%>" <%if rst("sys_type")=country then%>selected<%end if%>><%=RST("SYS_VALUE")%></option>
					<%
					RST.MOVENEXT
					WEND
					%>
				</SELECT>
				<%SET RST=NOTHING %>
			</TD>
		</tr>
		<tr height=30>
			<TD nowrap align=right>廠別<br>Xuong</TD>
			<TD >
				<select name=WHSNO   >
					<option value="">--ALL--- </option>
					<%SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' ORDER BY SYS_TYPE "
					SET RST = CONN.EXECUTE(SQL)
					WHILE NOT RST.EOF
					%>
					<option value="<%=RST("SYS_TYPE")%>" <%if rst("sys_type")=whsno  then%>selected<%end if%>><%=RST("SYS_VALUE")%></option>
					<%
					RST.MOVENEXT
					WEND
					%>
				</SELECT>
				<%SET RST=NOTHING %>
			</TD>
			<TD nowrap align=right >部門<br>Bo phan</TD>
			<TD >
				<select name=GROUPID   onchange="grpchg()"  >
				<option value="" selected >---ALL---</option>
				<%
				SQL="SELECT * FROM BASICCODE WHERE FUNC='GROUPID' and sys_type <>'AAA' ORDER BY SYS_TYPE "
				'SQL="SELECT * FROM BASICCODE WHERE FUNC='GROUPID'  ORDER BY SYS_TYPE "
				SET RST = CONN.EXECUTE(SQL)
				'RESPONSE.WRITE SQL
				WHILE NOT RST.EOF
				%>
				<option value="<%=RST("SYS_TYPE")%>" <%if rst("sys_type")=groupid then%>selected<%end if%>><%=RST("SYS_TYPE")%> <%=RST("SYS_VALUE")%></option>
				<%
				RST.MOVENEXT
				WEND
				%>
				</SELECT>
				<%SET RST=NOTHING %>
			</td>
		</tr>
		<TR height=30 >
			<TD nowrap align=right valign="top">單位(組)<br>Don vi</TD>
			<TD>
				<select name=zuno    MULTIPLE size=7  style="width:180" >
				<option value="" selected >---ALL--- </option>
				<%
				if trim(groupid)<>"" then 
					SQL="SELECT * FROM BASICCODE WHERE FUNC='zuno' and SYS_TYPE<>'XX' and left(sys_type,4) like '"& groupid &"%'  ORDER BY SYS_TYPE "				
				else
					SQL="SELECT * FROM BASICCODE WHERE FUNC='zuno' and SYS_TYPE<>'XX' ORDER BY SYS_TYPE "				
				end if 	
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
			<td nowrap align=right >班別<br>Ca</td>
			<td >
			 	<select name=shift >
			 		<option value="">--ALL--- </option>
					<%SQL="SELECT * FROM BASICCODE WHERE FUNC='shift' ORDER BY len(SYS_TYPE) desc, SYS_TYPE "				
						SET RST = CONN.EXECUTE(SQL)
						'RESPONSE.WRITE SQL
						WHILE NOT RST.EOF					%>
						<option value="<%=RST("SYS_TYPE")%>" <%if rst("sys_type")=shift then %>selected<%end if%>><%=RST("SYS_TYPE")%>-<%=RST("SYS_VALUE")%></option>
					<%RST.MOVENEXT
						WEND					%>
				</SELECT>
				<%SET RST=NOTHING 
				conn.close
				set conn=nothing
				%>			 					 	
			</td>
		</TR>
		<TR  height=30 >
			<td nowrap align=right >加班時間(起)<br>Time(Up)</td>
			<td>
				<input  type="text" style="width:100px" name=T1   maxlength=5 onblur=t1chg()   value="<%=T1%>" > 				
			</td>
			<td nowrap align=right >員工編號<br>So The</td>
			<td>
				<input  type="text" style="width:100px" name=eid   maxlength=5  value="<%=eid%>" > 				
			</td>
		</TR> 
		<tr height="50px">
			<td align=center colspan=4>
				<input type=button  name=btm class="btn btn-sm btn-danger" value="(Y)Confirm" onclick="go()" onkeydown="go()">
				<input type=reset  name=btm class="btn btn-sm btn-outline-secondary" value="(N)Cancel">
			</td>
		</tr>

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

function grpchg()
	<%=self%>.action ="<%=self%>.fore.asp"
	<%=self%>.submit()
end function 

function t1chg()
	if trim(<%=self%>.t1.value)<>"" then 
		<%=self%>.t1.value = left(trim(<%=self%>.t1.value),2) &":"&right(trim(<%=self%>.t1.value),2)
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