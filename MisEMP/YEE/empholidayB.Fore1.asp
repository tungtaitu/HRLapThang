<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%> 
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->
<!-- #include file="../Include/func.inc" -->
<!--#include file="../include/sideinfo.inc"-->
<%
Set conn = GetSQLServerConnection()	  
self="EMPHOLIDAYB"  


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
	<%=self%>.ym1.focus()	
	<%=self%>.ym1.SELECT()
end function    
function ym1chg()
	if <%=self%>.ym1.value<>"" then 
		<%=self%>.ym2.value = <%=self%>.ym1.value 
		<%=self%>.ym2.select()
	end if 	
end function 
-->
</SCRIPT>   
</head> 
<body   onkeydown="enterto()" onload='f()' >
<form name="<%=self%>" method="post"  >

	<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>
	
				<table width="100%" BORDER=0 align=center cellpadding=0 cellspacing=0 >
					<tr>
						<td align="center">							
							<table id="myTableForm" width="50%"> 
								<tr><td height="35px" colspan=4>&nbsp;</td></tr> 
								<tr>
									<td nowrap align=right>年月<br>Thang Nam</td>
									<td nowrap>
										<input type="text" style="width:45%" name=ym1 maxlength=6  onblur="ym1chg()">~
										<input type="text" style="width:45%" name=ym2 maxlength=6  >													
									</td>								
									<td nowrap align=right >日期<br>Ngay</td>
									<td>
										<input type="text" style="width:45%" name=dat1 maxlength=10 onblur="date_change(1)">~
										<input type="text" style="width:45%" name=dat2 maxlength=10 onblur="date_change(2)">
									</td> 
								</tr>
								<tr>
									<TD nowrap align=right>國籍<br>Quoc Tich</TD>
									<TD >										
										<select name=country   type="text" style="width:160px" >
											<option value="">----</option>
											<%SQL="SELECT * FROM BASICCODE WHERE FUNC='country'  ORDER BY SYS_type desc  "
											SET RST = CONN.EXECUTE(SQL)
											WHILE NOT RST.EOF  
											%>
											<option value="<%=RST("SYS_TYPE")%>" ><%=RST("SYS_VALUE")%></option>				 
											<%
											RST.MOVENEXT
											WEND 
											SET RST=NOTHING
											%>
										</SELECT>
									</TD>										 
									<TD nowrap align=right height=30 >廠別<br>Xuong</TD>
									<TD> 
										<select name=WHSNO  style="width:160px" >
											<option value="">---</option>
											<%SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' ORDER BY SYS_TYPE "
											SET RST = CONN.EXECUTE(SQL)
											WHILE NOT RST.EOF  
											%>
											<option value="<%=RST("SYS_TYPE")%>"><%=RST("SYS_VALUE")%></option>				 
											<%
											RST.MOVENEXT
											WEND 
											SET RST=NOTHING
											%>
										</SELECT>
									</TD>
								</tr>
								<tr>
									<TD nowrap align=right >部門<br>Bo Phan</TD>
									<TD>
										<select name=GROUPID    >
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
									</td>	
									<TD nowrap align=right >假別<br>Loai phep</TD>			
									<TD>
										<select name=JB    >	
										<option value="">---</option>		 
										<%SQL="SELECT * FROM BASICCODE WHERE FUNC='JB'  ORDER BY SYS_TYPE "
										SET RST = CONN.EXECUTE(SQL)
										WHILE NOT RST.EOF  
										%>
										<option value="<%=RST("SYS_TYPE")%>" ><%=RST("SYS_TYPE")%>.<%=RST("SYS_VALUE")%></option>				 
										<%
										RST.MOVENEXT
										WEND 
										%>		 				 
										</SELECT>
										<%SET RST=NOTHING %>
									</TD>
								</tr>
								<tr>
									<td nowrap align=right >員工編號<br>So The</td>
									<td colspan=3>
										<input  type="text" style="width:100px" name=empid1  maxlength=6 onchange=strchg(1)> 
									</td>
								</TR>
								<tr height="50px">
									<td align=center colspan=4>
										<input type=button  name=btm class="btn btn-sm btn-danger" value="(Y)Confirm" onclick="go()" onkeydown="go()">
										<input type=reset  name=btm class="btn btn-sm btn-outline-secondary" value="(N)Cancel">
										<input type=button  name=btmexcel class="btn btn-sm btn-outline-secondary" value="SaveTo EXCEL" onclick=goexcel()>
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
	<%=self%>.action="<%=self%>.Toexcel.asp?"
	<%=self%>.target="Back"
	<%=self%>.submit()
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
 	'IF <%=SELF%>.DAT1.VALUE="" AND <%=SELF%>.DAT2.VALUE="" THEN  
 	'	ALERT "必須輸入日期"
 	'	<%=SELF%>.DAT1.FOCUS() 
 	'	EXIT function  
 	'END IF 	
 	<%=self%>.action="<%=SELF%>.FORE.asp"
 	<%=self%>.submit() 
end function   
	

'*******檢查日期*********************************************
FUNCTION date_change(a)	

if a=1 then
	INcardat = Trim(<%=self%>.dat1.value)  		
elseif a=2 then
	INcardat = Trim(<%=self%>.dat2.value)
end if		
		    
IF INcardat<>"" THEN
	ANS=validDate(INcardat)
	IF ANS <> "" THEN
		if a=1 then
			Document.<%=self%>.dat1.value=ANS			
		elseif a=2 then
			Document.<%=self%>.dat2.value=ANS		 			
		end if		
	ELSE
		ALERT "EZ0067:輸入日期不合法 !!" 
		if a=1 then
			Document.<%=self%>.dat1.value=""
			Document.<%=self%>.dat1.focus()
		elseif a=2 then
			Document.<%=self%>.dat2.value=""
			Document.<%=self%>.dat2.focus()
		end if		
		EXIT FUNCTION
	END IF
		 
ELSE
	'alert "EZ0015:日期該欄位必須輸入資料 !!" 		
	EXIT FUNCTION
END IF 
   
END FUNCTION
</script> 