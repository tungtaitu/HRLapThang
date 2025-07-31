<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%> 
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->
<!-- #include file="../Include/func.inc" -->
<!--#include file="../include/sideinfo.inc"-->
<%
Set conn = GetSQLServerConnection()	  
self="YEFP10" 


nowmonth = year(date())&right("00"&month(date()),2)
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
	<%=self%>.whsno.focus()	
	'<%=self%>.yymm.select()
end function    
-->
</SCRIPT>   
</head> 
<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload='f()' >
<form name="<%=self%>" method="post" >
<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>
<table width="100%"  ><tr><td align="center">
	<table id="myTableForm" width="50%"> 
		<tr><td height="35px" colspan=4>&nbsp;</td></tr>  
		<TR  height=35 >
			<td nowrap align=right >國籍<br>quoc tich</td>
			<td >
				<select name=country style="width:120px">
					<option value="">----</option>
					<%sql="select * from basicCOde where func='country' and sys_type='VN' " 
					  set rds=conn.execute(Sql)
					  while not rds.eof
					%>
					<option value="<%=rds("sys_type")%>"><%=rds("sys_type")%>-<%=rds("sys_value")%></option>
					<%rds.movenext
					wend
					set rds=nothing
					%>
				</select>				
			</td>
			<td nowrap align=right >廠別<br>Xuong</td>
			<td >
				<select name=whsno style="width:120px">
					<option value="">----</option>
					<%sql="select * from basicCOde where func='whsno' " 
					  set rds=conn.execute(Sql)
					  while not rds.eof
					%>
					<option value="<%=rds("sys_type")%>"><%=rds("sys_type")%>-<%=rds("sys_value")%></option>
					<%rds.movenext
					wend
					set rds=nothing
					%>
				</select>
			</td>
		</TR>		
		<TR  height=35 >
			<td nowrap align=right >員工編號<br>so the</td>
			<td>
				<input type="text" style="width:100px" name=empid1 maxlength=5 onchange=strchg(1)> 
			</td>
			<td nowrap align=right >起迄日期<br>Ngay</td>
			<td nowrap>
				<input  type="text" style="width:100px" name="indat1" onblur="date_change(1)">~
				<input  type="text" style="width:100px" name="indat2" onblur="date_change(2)">
			</td>
		</TR>  
		<tr height="50px">
			<td align=center colspan=4>
				<input type=button  name=btm class="btn btn-sm btn-danger" value="確   認" onclick="go()" onkeydown="go()">
				<input type=reset  name=btm class="btn btn-sm btn-outline-secondary" value="取   消">
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
</script> 