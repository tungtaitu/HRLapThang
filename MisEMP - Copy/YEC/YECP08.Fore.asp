<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%> 
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->
<!-- #include file="../Include/func.inc" -->
<!--#include file="../include/checkpower.asp"-->  
<!-- #include file="../Include/SIDEINFO.inc" -->
<%
Set conn = GetSQLServerConnection()	  
self="YECP08"  


nowmonth = year(date())&right("00"&month(date()),2)  
if month(date())="01" then  
	calcmonth = year(date()-1)&"12"  	
else
	calcmonth =  year(date())&right("00"&month(date())-1,2)   
end if 	   

if day(date())<=11 then 
	if month(date())="01" then  
		calcmonth = year(date()-1)&"12" 
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
<body  topmargin="50" leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload='f()' >
<form name="<%=self%>" method="post" action="EMPFILE.SALARY.ASP">
<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>

<table width="100%" BORDER=0 align=center cellpadding=0 cellspacing=0 >
	<tr>
		<td align="center">
			<table id="myTableForm" width="50%"> 
				<tr><td height="35px" colspan=4>&nbsp;</td></tr>
				<tr height=30 >
					<TD nowrap align=right>計薪年月<br>Tien Luong</TD>
					<TD ><INPUT type="text" style="width:100px" NAME=YYMM   VALUE="<%=calcmonth%>" ></TD>	
				
					<TD nowrap align=right height=30 valign="top" >國籍<br>Quoc Tich<br>(可複選)</TD>
					<TD >
						<select name=country   style='width:120px' >
							<% 
							if Session("NETWHSNO")="ALL" or Session("RIGHTS")<="1"  then%>					
							<% 
								SQL="SELECT * FROM BASICCODE WHERE FUNC='country' ORDER BY SYS_TYPE desc"
							else
								SQL="SELECT * FROM BASICCODE WHERE FUNC='country' and sys_type='VN' ORDER BY SYS_TYPE desc"
							end if 	
							'SQL="SELECT * FROM BASICCODE WHERE FUNC='country'  ORDER BY SYS_type desc  "
							SET RST = CONN.EXECUTE(SQL)
							WHILE NOT RST.EOF  
							%>
							<option value="<%=RST("SYS_TYPE")%>"  <% if rst("sys_type")="VN" then %>selected<%end if%>><%=RST("SYS_VALUE")%></option>				 
							<%
							RST.MOVENEXT
							WEND 
							%>
						</SELECT>
						<%SET RST=NOTHING %>
					</TD>	
				</tr>	
				<TR>
					<TD nowrap align=right height=30 >廠別<br>Xuong</TD>
					<TD >
						<select name=whsno   style='width:120'  >
							<%if Session("NETWHSNO")="ALL" or Session("RIGHTS")<="1" or Session("RIGHTS")="8" or Session("RIGHTS")="5"  then%>
							<option value="">----</option>
							<% 
								SQL="SELECT * FROM BASICCODE WHERE FUNC='whsno' ORDER BY SYS_TYPE "
							else
								SQL="SELECT * FROM BASICCODE WHERE FUNC='whsno'  ORDER BY SYS_TYPE "
							end if 
							'SQL="SELECT * FROM BASICCODE WHERE FUNC='country'  ORDER BY SYS_type desc  "
							SET RST = CONN.EXECUTE(SQL)
							WHILE NOT RST.EOF  
							%>
							<option value="<%=RST("SYS_TYPE")%>" <%if wx=rst("sys_type") then%>selected<%end if%>><%=RST("SYS_VALUE")%></option>				 
							<%
							RST.MOVENEXT
							WEND 
							%>
						</SELECT>
						<%SET RST=NOTHING %>
					</TD>
					<TD nowrap align=right >部門<br>Bo Phan</TD>
					<TD >
						<select name=GROUPID  style='width:120px'  >				
						<%if Session("RIGHTS")<="2" or Session("RIGHTS")>="8" or Session("RIGHTS")="5"  then%>
							<option value="">-----</option>
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
					</td>
				</tr>
				<tr height=30 >	
					<TD nowrap align=right >職務<vr>Chuc vu</TD>			
					<TD >
						<select name=JOB  style='width:120px'  >	
						<option value="">-----</option>		 
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
						<%SET RST=NOTHING 
						conn.close
						set conn=nothing 				
						%>
					</TD>
					<td nowrap align=right >員工編號<br>SO the</td>
					<td>
						<input type="text" style="width:100px" name=empid1  maxlength=5 onchange=strchg(1)> 						
					</td>
				</TR>
				<tr >
					<td align=center  colspan=4 height="50px">						
						<input type=button  name=btm class="btn btn-sm btn-danger" value="(Y)Confirm" onclick="go()" onkeydown="go()">&nbsp;&nbsp;
						<input type=reset  name=btm class="btn btn-sm btn-outline-secondary" value="(N)Cancel">&nbsp;&nbsp;
						<input type=button  name=btm class="btn btn-sm btn-outline-secondary" value="SaveTo EXCEL" onclick=goexcel() style='background-color:#e4e4e4'>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table> 

</body>
</html>


<script language=vbs> 
function goexcel()
	'open "<%=self%>.toexcel.asp" , "Back" 
	<%=self%>.action="<%=self%>.toexcel.asp"
	<%=self%>.target="Back"
	<%=self%>.submit()
	parent.best.cols="100%,00%"
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
 	<%=self%>.target="Fore"
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