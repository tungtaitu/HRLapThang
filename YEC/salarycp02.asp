<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%> 
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->
<!-- #include file="../Include/func.inc" -->
<!--#include file="../include/checkpower.asp"-->  
<!-- #include file="../Include/SIDEINFO.inc" -->
<%
Set conn = GetSQLServerConnection()	  
self="SALARYCP02"  


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
<body  topmargin="50" leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload='f()' >
<form name="<%=self%>" method="post" action="EMPFILE.SALARY.ASP">
 <input type="hidden" name="uid" value="<%=session("netuser")%>">

	<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>
	
				<table width="100%" BORDER=0 cellpadding=0 cellspacing=0 >
					<tr>
						<td align="center">
							<table id="myTableForm" width="50%"> 
								<tr><td height="35px">&nbsp;</td></tr>
								<tr>
									<TD nowrap align=right>計薪年月<br>Tien luong</TD>
									<TD ><INPUT type="text" style="width:100px" NAME=YYMM   VALUE="<%=calcmonth%>"></TD>								
									<TD nowrap align=right height=30 width=150 >國籍<br>Quoc tich</TD>
									<TD >
										<select name=country style="width:120px"  >
											<%if Session("NETWHSNO")="ALL" or Session("RIGHTS")<="1" or Session("RIGHTS")="8" or Session("RIGHTS")="5"  then%>
											<option value="">---ALL---</option>
											<% 
												SQL="SELECT * FROM BASICCODE WHERE FUNC='country' ORDER BY SYS_TYPE "
											else
												SQL="SELECT * FROM BASICCODE WHERE FUNC='country'   ORDER BY SYS_TYPE "
											end if 
											'SQL="SELECT * FROM BASICCODE WHERE FUNC='country'  ORDER BY SYS_type desc  "
											SET RST = CONN.EXECUTE(SQL)
											WHILE NOT RST.EOF  
											%>
											<option value="<%=RST("SYS_TYPE")%>" ><%=RST("SYS_TYPE")%>-<%=RST("SYS_VALUE")%></option>				 
											<%
											RST.MOVENEXT
											WEND
											rst.close	
											SET RST=NOTHING
											%>											
										</SELECT>
										
									</TD>	
								</tr>
								<tr class="txt">		 
									<TD nowrap align=right>廠別<BR>Xuong</TD>
									<TD > 
										<select name=WHSNO  style="width:120px"  >
											<% 
											if Session("RIGHTS")<="1" or Session("RIGHTS")="8" then%>
											<option value="">---ALL---</option>
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
											rst.close
											%>
										</SELECT>
										<%SET RST=NOTHING %>
									</TD>
									<TD nowrap align=right >3.組/部門<br>Bo Phan</TD>
									<TD >
										<select name=GROUPID    style="width:120px" >				
										<%if Session("RIGHTS")<="2" or Session("RIGHTS")>="8" or Session("RIGHTS")="5"  then%>
											<option value="">---ALL---</option>
											<% 
												SQL="SELECT * FROM BASICCODE WHERE FUNC='GROUPID' and sys_type <>'AAA' ORDER BY SYS_TYPE " 
											else
												SQL="SELECT * FROM BASICCODE WHERE FUNC='GROUPID' and sys_type <>'AAA' and  sys_type= '"& session("NETG1") &"' ORDER BY SYS_TYPE "
											end if   
											
											SET RST = CONN.EXECUTE(SQL)
											RESPONSE.WRITE SQL 
											WHILE NOT RST.EOF  
										%>
											<option value="<%=RST("SYS_TYPE")%>" <% if RST("SYS_TYPE")=session("mywhsno")  then  %>selected<%end if%> ><%=RST("SYS_TYPE")%><%=RST("SYS_VALUE")%></option>				 
										<%
											RST.MOVENEXT
											WEND 
											rst.close
										%>
										</SELECT>
										<%SET RST=NOTHING 
										conn.close
										set conn=nothing 
										%>
									</TD>
								</tr> 
								<TR class="txt" >
									<td nowrap align=right >員工編號<BR>So The</td>
									<td>
										<input type="text" style="width:100px" name=empid1  onchange=strchg(1)> 										
									</td>
									<td nowrap align=right >員工統計<br>Tong ke</td>
									<td colspan=3>
										<select name=outemp  style="width:120px">
											<option value="">全部(ALL)</option>			 		
											<option value="D">本月離職(T.C)</option>
										</select>
									</td>
								</TR> 
								<TR>
									<td nowrap align=right ></td>
									<td nowrap colspan=5>
										<input type=checkbox name="func" onclick="funcchg()"> 不印績效獎金(Khong in tien thuong)
										<input size=1 name=nojx type=hidden value=""><br>																		
										<INPUT type="radio" id=radio1 name=radio1 onclick=typechg(0) checked > 印結帳薪資 &nbsp;
										<INPUT type="radio" id=radio1 name=radio1 onclick=typechg(1)  > 印目前薪資
										<input size=1 name=job type=hidden value="A">										
									</td>
								</TR>	
								<tr >
									<td align=center colspan=6 height="50px">
										<input type=button  name=btm class="btn btn-sm btn-danger" value="(Y)確認Print" onclick="go()" onkeydown="go()">
										<input type=reset  name=btm class="btn btn-sm btn-outline-secondary" value="(N)取消Cancel">
									</td>
								</tr>
							</table>
						</td>
					</tr>					
				</table>
			 

</body>
</html>


<script language=vbs>
function funcchg()
	if <%=self%>.func.checked=true then 
		<%=self%>.nojx.value="Y"
	else
		<%=self%>.nojx.value=""
	end if 
end function   


function typechg(a)
	if a=0 then 
		<%=self%>.job.value="A"
	elseif a=1 then 
		<%=self%>.job.value=""
	else
		<%=self%>.job.value="A"	
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
 
 	<%=self%>.action="<%=session("rpt")%>"&"yec/"&"<%=self%>.getrpt.asp"
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