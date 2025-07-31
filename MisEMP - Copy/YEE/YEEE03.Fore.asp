<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<!-- #include file="../Include/func.inc" -->
<!--#include file="../include/sideinfo.inc"-->
<%
Set conn = GetSQLServerConnection()
self="yeee03" 
if  instr(conn,"168")>0 then 
	w1="LA"
elseif  instr(conn,"169")>0 then 
	w1="DN"	
elseif  instr(conn,"47")>0 then 
	w1="BC"	
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
	<%=self%>.yymm.focus()
end function
-->
</SCRIPT>
</head>
<body  topmargin="5" leftmargin="0"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload='f()' >
<form name="<%=self%>" method="post" >

	<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>
	<table width="100%" border=0 >
		<tr>
			<td align="center">
				
				<table id="myTableForm" width="70%"> 
					<tr><td height="35px" colspan=4>&nbsp;</td></tr>
					<TR height=30>
						<td nowrap align=right >統計年度：</td>
						<td>
							<input type="text" style="width:100px" name=yymm maxlength=4 > 
						</td>
						<TD nowrap align=right>國籍：</TD>
						<TD >
							<select name=country   style='width:120px'  >
								<%if session("rights")<>"9" then %><option value="">全部 </option><%end if%>
								<%					
								SQL="SELECT * FROM BASICCODE WHERE FUNC='country'  ORDER BY SYS_type desc  "					
								SET RST = CONN.EXECUTE(SQL)
								WHILE NOT RST.EOF
								%>
								<option value="<%=RST("SYS_TYPE")%>" <%if rst("sys_type")="VN" then %>selected<%end if%> ><%=RST("SYS_VALUE")%></option>
								<%
								RST.MOVENEXT
								WEND
								%>
							</SELECT>
							<%SET RST=NOTHING %>
						</TD>
					</tr>
					<tr height=30>
						<TD nowrap align=right>廠別：</TD>
						<TD >
							<select name=WHSNO   >
								<%if session("rights")<>"9" then %><option value="">全部 </option><%end if%>
								<%
								if session("rights")="9"  then 
									SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' and sys_type='LA' ORDER BY SYS_TYPE "
								else
									SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' ORDER BY SYS_TYPE "
								end if 
								SET RST = CONN.EXECUTE(SQL)
								WHILE NOT RST.EOF
								%>
								<option value="<%=RST("SYS_TYPE")%>" <%if w1=rst("sys_type") then%>selected<%end if%> ><%=RST("SYS_VALUE")%></option>
								<%
								RST.MOVENEXT
								WEND
								%>
							</SELECT>
							<%SET RST=NOTHING %>
						</TD>
						<TD nowrap align=right >部門：</TD>
						<TD >
							<select name=GROUPID    >
							<option value="" selected >全部 </option>
							<%
							SQL="SELECT * FROM BASICCODE WHERE FUNC='GROUPID' and sys_type <>'AAA' ORDER BY SYS_TYPE "
							'SQL="SELECT * FROM BASICCODE WHERE FUNC='GROUPID'  ORDER BY SYS_TYPE "
							SET RST = CONN.EXECUTE(SQL)
							'RESPONSE.WRITE SQL
							WHILE NOT RST.EOF
							%>
							<option value="<%=RST("SYS_TYPE")%>" ><%=RST("SYS_TYPE")%> <%=RST("SYS_VALUE")%></option>
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
						<td>
							<input name=empid1  size=15 maxlength=5 >
						</td>
						<td nowrap align=right >顯示方式：</td>
						<td >
							<select  name=showby  >					
								<option value="A">A.依部門/工號</option>
								<option value="B">B.依工號</option>			
								<option value="">ALL</option>
							</select>
						</td>
					</TR> 
					<tr height="50px">
						<td align=center colspan=4>
							<input type=button  name=btm class="btn btn-sm btn-danger" value="確   認" onclick="go()" onkeydown="go()">
							<input type=reset  name=btm class="btn btn-sm btn-outline-secondary" value="取   消">
							<input type=button  name=btn class="btn btn-sm btn-outline-secondary" value="save To Excel" onclick=goexcel() style='background-color:#ffccff' >
						</td>
					</tr>
				</table>				
			</td>
		</tr>
	</table>
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
	if <%=self%>.yymm.value="" then 
		alert "請輸入[統計年度]!!"
		<%=self%>.yymm.focus()
		exit function 
	end if 	
	parent.best.cols="100%,0%"
 	<%=self%>.action="<%=self%>.Foregnd.asp"
	<%=self%>.target="Fore"
 	<%=self%>.submit()
end function

function goexcel()
	if <%=self%>.yymm.value="" then 
		alert "請輸入[統計年度]!!"
		<%=self%>.yymm.focus()
		exit function 
	end if 	
	'open "<%=self%>.toexcel.asp" , "Back" 
	<%=self%>.action="<%=self%>.toexcel.asp"
	<%=self%>.target="Back"
	<%=self%>.submit()
	'parent.best.cols="50%,50%"
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