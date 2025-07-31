<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%> 
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->
<!-- #include file="../Include/func.inc" -->
<!-- #include file="../Include/SIDEINFO.inc" -->
<%
Set conn = GetSQLServerConnection()	  
self="YECE03"   
  

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
<form name="<%=self%>" method="post" action="<%=self%>.SALARY.ASP">
 
<BR><BR>
<table width=500  ><tr><td >
	<table width=300 align=center border=0 cellspacing="2" cellpadding="2" class="txt8" > 
		<tr height=30 >
			<TD nowrap align=right>計薪年月<br>Tien Luong</TD>
			<TD ><INPUT NAME=YYMM  CLASS=INPUTBOX VALUE="<%=calcmonth%>" SIZE=10 maxlength="6"> (yyyymm)</TD>	
		</TR>
		<TR>
		 	<TD nowrap align=right height=30 >國籍<br>Quoc Tich</TD>
			<TD >
				<select name=country  class=font9 style='width:75'  >					
					<%SQL="SELECT * FROM BASICCODE WHERE FUNC='country' and sys_type='TA' ORDER BY SYS_type desc  "
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
			<TD nowrap align=right height=30 >廠別<br>Xuong</TD>
			<TD > 
				<select name=WHSNO  class=font9 >
					<option value="">----</option>
					<%SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' ORDER BY SYS_TYPE "
					SET RST = CONN.EXECUTE(SQL)
					WHILE NOT RST.EOF  
					%>
					<option value="<%=RST("SYS_TYPE")%>" <%if wx=rst("sys_type") then %>selected<%end if%> ><%=RST("SYS_VALUE")%></option>				 
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
			<TD nowrap align=right >部門<br>Bo phan</TD>
			<TD >
				<select name=GROUPID  class=font9  >
				<option value="" selected >----</option>
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
		<TR  height=30 >
			<td nowrap align=right >員工編號<br>So the</td>
			<td colspan=3>
				<input name=empid1 class=inputbox size=15 maxlength=5 onchange=strchg(1)> 
				
			</td>
		</TR>
		<TR  height=30 >
			<td nowrap align=right >員工統計<br>Thong ke</td>
			<td colspan=3>
			 	<select name=outemp class=font9> 
			 		<option value="">---</option>
			 		<option value="N">(Toan Bo N.V)在職</option>
			 		<option value="D">(Thoi viec thang nay)本月離職</option>
			 	</select>	
			</td>
		</TR>
		<!--TR  height=30 >
			<td align=right><input type=checkbox  name=chk1  onclick="chksts()" ></td>
			<td colspan=3 >
				重新計算本月				
			</td>		
		</TR-->  
		<input type=hidden name=recalc value="N" size=1> 
	</table><BR>	
	<table width=450 align=center>
		<tr >
			<td align=center>
				<input type=button  name=btm class=button value="(Y)COnfirm" onclick="go()" onkeydown="go()">
				<input type=reset  name=btm class=button value="(N)Cancel">
			</td>
		</tr>	
	</table>	

</td></tr></table> 

</body>
</html>


<script language=vbs> 
function chksts()
	if <%=self%>.chk1.checked=true then 
		<%=self%>.recalc.value="Y"
	else
		<%=self%>.recalc.value="N"
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
 	<%=self%>.action="<%=self%>.SALARY.asp"
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