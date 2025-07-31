<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<!-- #include file="../Include/func.inc" -->
<!--#include file="../include/sideinfo.inc"-->
<%
Set conn = GetSQLServerConnection()
self="yefp04"


indat1 = request("indat1")
indat2 = request("indat2")
country = request("country")
WHSNO = request("WHSNO")
GROUPID = request("GROUPID")
shift = request("shift")
T1 = request("T1")
eid = request("empid1") 
showby= request("showby") 

if request("indat1")="" then indat1=DD2
if request("indat2")="" then indat2=DD2
if request("country")="" then country="VN"
if request("WHSNO")="" then WHSNO=session("mywhsno") 

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
<input  type="hidden" name="netuser" value="<%=session("netuser")%>">
<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>

<table width="100%"><tr><td align="center">
	<table id="myTableForm" width="50%"> 
		<tr><td height="35px" colspan=4>&nbsp;</td></tr>
		<TR height=30>
			<td nowrap align=right >日期範圍<br>Ngay</td>
			<td>
				<input  type="text" style="width:45%" name=indat1 maxlength=10 onblur="date_change(1)" value="<%=indat1%>">~
				<input  type="text" style="width:45%" name=indat2 maxlength=10 onblur="date_change(2)" value="<%=indat2%>">
			</td>
		 	<TD nowrap align=right>國籍<br>Quoc Tich</TD>
			<TD >			
				<select name=country    style="width:150px" >					
					<%if session("rights")<>"9" then %><option value="">---ALL---</option><%end if%>
					<%
					if session("rights")="9" then 
						SQL="SELECT * FROM BASICCODE WHERE FUNC='country' and SYS_TYPE='VN' ORDER BY SYS_type desc  "
					else
						SQL="SELECT * FROM BASICCODE WHERE FUNC='country'  ORDER BY SYS_type desc  "
					end if 	
					SET RST = CONN.EXECUTE(SQL)
					WHILE NOT RST.EOF
					%>
					<option value="<%=RST("SYS_TYPE")%>"  <%if rst("sys_type")=country then%>selected<%end if%> ><%=RST("SYS_TYPE")%>-<%=RST("SYS_VALUE")%></option>
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
				<select name=WHSNO >
					<%if session("rights")<>"9" then %><option value="">--ALL--</option><%end if%>
					<%
					if session("rights")="9"  then 
						SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' and sys_type='LT' ORDER BY SYS_TYPE "
					else
						SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' ORDER BY SYS_TYPE "
					end if 	
					SET RST = CONN.EXECUTE(SQL)
					WHILE NOT RST.EOF
					%>
					<option value="<%=RST("SYS_TYPE")%>"  <%if rst("sys_type")=whsno  then%>selected<%end if%>><%=RST("SYS_VALUE")%></option>
					<%
					RST.MOVENEXT
					WEND
					%>
				</SELECT>
				<%SET RST=NOTHING %>
			</TD>
			<TD nowrap align=right >部門<BR>Bo phan</TD>
			<TD >
				<select name=GROUPID      onchange="grpchg()" >
				<option value="" selected >---ALL---</option>
				<%
				SQL="SELECT * FROM BASICCODE WHERE FUNC='GROUPID' and sys_type <>'AAA' ORDER BY SYS_TYPE "
				'SQL="SELECT * FROM BASICCODE WHERE FUNC='GROUPID'  ORDER BY SYS_TYPE "
				SET RST = CONN.EXECUTE(SQL)
				'RESPONSE.WRITE SQL
				WHILE NOT RST.EOF
				%>
				<option value="<%=RST("SYS_TYPE")%>" <%if rst("sys_type")=groupid  then%>selected<%end if%> ><%=RST("SYS_TYPE")%>-<%=RST("SYS_VALUE")%></option>
				<%
				RST.MOVENEXT
				WEND
				%>
				</SELECT>
				<%SET RST=NOTHING %>
			</td>
		</tr>
		<TR height=30 >
			<TD nowrap align=right valign="top">單位(組)<br>Don Vi</TD>
			<TD >
				<select name=zuno  class="txt8"  MULTIPLE size=7   style="width:180">
				<option value="" selected >---ALL--- </option>
				<%if trim(groupid)<>"" then 
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
			<td nowrap align=right >員工編號<BR>So the</td>
			<td>
				<input name=empid1  size=15 maxlength=6 value="<%=eid%>" >  
			</td>
		</TR>
		<TR  height=30 >
			<td nowrap align=right >員工統計<br>Loai</td>
			<td>
			 	<select name=outemp >
			 		<option value="">---ALL--</option>			 		
			 		<option value="D">(Thoi viec)已離職</option>
			 	</select>
			</td>
			<td nowrap align=right >班別<BR>Ca</td>
			<td>
			 	<select name=shift >
			 		<option value="">---ALL---</option>		
					<%SQL="SELECT * FROM BASICCODE WHERE FUNC='shift'  ORDER BY len(sys_type) desc, SYS_TYPE "
					SET RST = CONN.EXECUTE(SQL)
					WHILE NOT RST.EOF
					%>
					<option value="<%=RST("SYS_TYPE")%>" <%if rst("sys_type")=shift then %>selected<%end if%>><%=RST("SYS_TYPE")%>-<%=RST("SYS_VALUE")%></option>
					<%
					RST.MOVENEXT
					WEND
					%>					
					<%SET RST=NOTHING %>									
			 	</select>
			</td>
		</TR>
		<TR height=30 >
			<td nowrap align=right >顯示方式<br>Sap xep</td>
			<td>
				<select  name=showby  >					
					<option value="">ALL(依工號)</option>
					<option value="B" <%if showby="B" then%>selected<%end if%>>B.(theo nagy)依班別/工號</option>					
					<option value="C" <%if showby="C" then%>selected<%end if%>>C.(theo bo phan, ca)依組別/班別/工號</option>
			</td >
			<td colspan=2></td>
		</TR> 
		<TR height="50px" >			
			<td align=center colspan="4">
				<input type=button  name=btm class="btn btn-sm btn-danger" value="(Y)Confrim" onclick="go()" onkeydown="go()">
				<input type=reset  name=btm class="btn btn-sm btn-outline-secondary" value="(N)Cancel">
			</td>					
		</TR> 
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
 	<%=self%>.action="<%=session("rpt")%>"&"yef/"&"<%=self%>.getrpt.asp"
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

function grpchg()
	<%=self%>.action ="<%=self%>.fore.asp"
	<%=self%>.submit()
end function 

</script> 