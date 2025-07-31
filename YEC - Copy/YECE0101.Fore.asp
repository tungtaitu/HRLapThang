<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%> 
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->
<!-- #include file="../Include/func.inc" -->
<!--#include file="../include/checkpower.asp"-->  
<!-- #include file="../Include/SIDEINFO.inc" -->
<%
Set conn = GetSQLServerConnection()	  
self="YECE0101"  


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
<body  leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload='f()' >
<form name="<%=self%>" method="post" action="EMPFILE.SALARY.ASP">
 
	<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>
	<table width="100%" border=0 >
		<tr>
			<td>
				<table width="100%" BORDER=0 align=center cellpadding=0 cellspacing=0 >
					<tr>
						<td align="center">
							<table id="myTableForm" width="50%"> 
								<tr><td colspan=4 height="40px">&nbsp;</td></tr>
								<tr>
									<TD nowrap align=right>計薪年月<br>Tien Luong</TD>
									<TD ><INPUT type="text" style="width:100px" NAME=YYMM VALUE="<%=calcmonth%>"></TD>	
									<TD nowrap align=right >國籍<br>Quoc Tich</TD>
									<TD >
										<select name=F_country style="width:120px">
											<% 
											if Session("NETWHSNO")="ALL" or Session("RIGHTS")<="1"  then%>					
											<% 
												SQL="SELECT * FROM BASICCODE WHERE FUNC='country' ORDER BY SYS_TYPE desc"
											else
												SQL="SELECT * FROM BASICCODE WHERE FUNC='country' and sys_type='VN' ORDER BY SYS_TYPE desc"
											end if 	
											SQL="SELECT * FROM BASICCODE WHERE FUNC='country'  ORDER BY SYS_type desc  "
											SET RST = CONN.EXECUTE(SQL)
											WHILE NOT RST.EOF  
											%>
											<option value="<% if (RST("SYS_TYPE") ="TW" OR RST("SYS_TYPE") ="MA"  OR RST("SYS_TYPE") ="CN" OR RST("SYS_TYPE") ="TA" ) and Session("netuser")="DARK" THEN %>  <%ELSE %> <%=RST("SYS_TYPE")%> <% END IF %>"  <% if rst("sys_type")="VN" then %>selected<%end if%>>   <% if (RST("SYS_TYPE") ="TW" OR RST("SYS_TYPE") ="MA" OR RST("SYS_TYPE") ="CN" OR RST("SYS_TYPE") ="TA" ) and Session("netuser")="DARK" then %>  <% else %>  <%=RST("SYS_VALUE") %> <%end if %> </option>				 
											<%
											RST.MOVENEXT
											WEND 
											%>
										</SELECT>
										<%SET RST=NOTHING %>
									</TD>	
								</tr>
								<tr>		 
									<TD nowrap align=right>廠別<br>Xuong</TD>
									<TD > 
										<select name=F_WHSNO style="width:120px"  >
											<% 
											if Session("RIGHTS")="0"  or Session("netuser")="L0197" then%>											
											<% 
												SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' ORDER BY SYS_TYPE "
											else
												SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO'  ORDER BY SYS_TYPE "
											end if 	
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
									<TD nowrap align=right >組/部門<br>Bo Phan</TD>
									<TD >
										<select name=GROUPID style="width:120px" >
										<option value="">----</option>
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
								</tr>
								<tr>	
									<TD nowrap align=right >職等<br>Chuc vu</TD>			
									<TD >
										<select name=JOB style="width:120px">	
										<option value="">----</option>	 
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
									<td nowrap align=right >員工編號<br>So The</td>
									<td>
										<input type="text" style="width:120px" name=empid1  maxlength=6 onchange=strchg(1)> 				
									</td>
								</TR>
								<TR>
									<td nowrap align=right >員工統計<br>Thong ke</td>
									<td>
										<select name=outemp style="width:120px"> 
											<option value="">----</option>
											<option value="N">Toan bo N.V(在職)</option>
											<option value="D">Thoi viec thang nay(本月離職)</option>
											<option value="<%=nowmonth%>">N.V moi  thang nay(本月新進)</option>
											<option value="<%=calcmonth%>">N.V moi thang truoc(上月新進)</option>
										</select>	
									</td>
									<td nowrap align=right >員工年資<br>So thang lam viec</td>
									<td>
										<select name=nzs style="width:120px">					
											<option value="">----</option>
											<option value="<12">( <12thang )1年以下</option>
											<option value="=12">( >=12thang )滿1年</option>
											<option value="=24">( >=24thang )滿2年</option>
											<option value="=36">( =36 thang )滿3年</option>
											<option value=">36">( >=36thang )3年以上</option>
										</select>
									</td>
								</TR>
								<tr>
									<td align=center colspan=4 height="50px">
										<input type=button  name=btm class="btn btn-sm btn-danger" value="(Y)Confirm" onclick="go()" onkeydown="go()">
										<input type=reset  name=btm class="btn btn-sm btn-outline-secondary" value="(N)Cancel">
									</td>
								</tr>
							</table>
						</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>

</body>
</html>


<script language=vbs>  

function strchg(a)
	if a=1 then 
		<%=self%>.empid1.value = Ucase(<%=self%>.empid1.value)
	elseif a=2 then 	
		<%=self%>.empid2.value = Ucase(<%=self%>.empid2.value)
	end if 	
end function 
	
function go()  
	if <%=self%>.F_country.value="" then 
		alert "請選擇國籍!! xin danh lai Quoc tich"
		<%=self%>.F_country.focus()
		exit function 
	end if 
 	<%=self%>.action="YECE0101.ForeGnd.asp"
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