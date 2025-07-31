<%@LANGUAGE=VBSCRIPT CODEPAGE=950%> 
<!-- #include file = "../../GetSQLServerConnection.fun" --> 
<!-- #include file="../../ADOINC.inc" -->
<!-- #include file="../../Include/func.inc" -->
<!--#include file="../../include/checkpower.asp"-->  
<%
Set conn = GetSQLServerConnection()	  
self="SALARYCP04"  


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
<form name="<%=self%>" method="post" action="EMPFILE.SALARY.ASP">
<table width="460" border="0" cellspacing="0" cellpadding="0">
	<tr><TD>
	<img border="0" src="../../image/icon.gif" align="absmiddle">
	4.薪資總表</TD></tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=500>		
<BR><BR>
<table width=500  ><tr><td >
	<table width=400 align=center border=0 cellspacing="0" cellpadding="0"  > 
		<tr height=30 >
			<TD nowrap align=right>計薪年月：</TD>
			<TD ><INPUT NAME=YYMM  CLASS=INPUTBOX VALUE="<%=calcmonth%>" SIZE=10></TD>	
		</TR>
		<TR>
		 	<TD nowrap align=right height=30 >國籍：</TD>
			<TD >
				<select name=country  class=font9 style='width:75'  >
					<%if Session("NETWHSNO")="ALL" or Session("RIGHTS")<="1" or Session("RIGHTS")="8" or Session("RIGHTS")="5"  then%>
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
			<TD nowrap align=right height=30 >廠別：</TD>
			<TD > 
				<select name=WHSNO  class=font9 >
					<%if Session("RIGHTS")<="1" or Session("RIGHTS")="8" then%>
					<option value="">全部 </option>
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
			</TD> 
		</TR>		
		<tr>	 
		<TR height=30 >
			<TD nowrap align=right >組/部門：</TD>
			<TD >
				<select name=GROUPID  class=font9  >				
				<%if Session("RIGHTS")<="2" or Session("RIGHTS")>="8" or Session("RIGHTS")="5" then%>
					<option value="">全部 </option>
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
		<TR  height=30 >
			<td nowrap align=right >員工編號：</td>
			<td colspan=3>
				<input name=empid1 class=inputbox size=15 maxlength=5 onchange=strchg(1)> 
				
			</td>
		</TR>
		<TR  height=35 > 
			<td colspan=4 align=center>
			 	
			 	<INPUT type="radio" id=radio1 name=radio1 onclick=typechg(0) checked > 印結帳薪資 &nbsp;
				<INPUT type="radio" id=radio1 name=radio1 onclick=typechg(1)  > 印目前薪資
			 	<input size=1 name=job type=hidden value="A">
			 	
			</td>
		</TR>				
		<!--TR  height=30 >
			<td nowrap align=right >員工統計：</td>
			<td colspan=3>
			 	<select name=outemp class=font9> 
			 		<option value="">全部</option>
			 		<option value="N">在職</option>
			 		<option value="D">已離職</option>
			 	</select>	
			</td>
		</TR-->
	</table><BR>	
	<table width=450 align=center>
		<tr >
			<td align=center>
				<input type=button  name=btm class=button value="確   認" onclick="go()" onkeydown="go()">
				<input type=reset  name=btm class=button value="取   消">
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