<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<!-- #include file="../Include/func.inc" -->
<!--#include file="../include/sideinfo.inc"-->
<%
Set conn = GetSQLServerConnection()
self="YECE0202"  
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

if request("YYMM")<>"" then  calcmonth = request("YYMM")
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
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
<form name="<%=self%>" method="post"  >
<input name="flag" type="hidden" value="F">	
	<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>	
	<table width="100%" BORDER=0 align=center cellpadding=0 cellspacing=0 >
		<tr>
			<td align="center">
				<table class="txt" cellpadding=3 cellspacing=3>
					<tr height=30 >
						<TD nowrap align=right>計薪年月<br>Tien Luong</TD>
						<TD><INPUT type="text" style="width:100px" NAME=YYMM  VALUE="<%=calcmonth%>" SIZE=10></TD>			
						<td>( yyyymm)</td>
						<TD nowrap align=right>部門<br>bo phan</TD>
						<TD nowrap >
							<select name=GROUPID style="width:120px">
							<option value="" selected >--All--</option>
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
						<TD nowrap align=right>工號<br>So the</TD>
						<TD ><INPUT type="text" style="width:100px" NAME=eid  VALUE="<%=eid%>" SIZE=8></TD>			
					</tr>
					<%if  request("flag")="" then %>
					<tr height=60>
						<td  colspan=7 align="center">
							<input type=button  name=btm class="btn btn-sm btn-danger" value="(Y)確   認" onclick="go()" onkeydown="go()">
							<input type=reset  name=btm class="btn btn-sm btn-outline-secondary" value="(N)取   消">
						</td>
					</tr>
					<%end if%>
					<tr>
						<td  colspan=7 align="center">
							<%if request("flag")="F" then %>
							<font color="red"> 資料處理中...請稍候 ....... !! </font>
							<%end if%>
						</td>
					</tr>
				</table>
			</td>
		</tr>					
	</table>
			
</form>
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
	
	<%=self%>.action="<%=SELF%>.Fore.asp"
 	<%=self%>.submit()  
	ym = <%=self%>.yymm.value
	g1 = <%=self%>.GROUPID.value
	eid = <%=self%>.eid.value
	open "<%=self%>.updnew.asp?yymm="&ym &"&g1="& g1 &"&eid="&eid , "Back"
	
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