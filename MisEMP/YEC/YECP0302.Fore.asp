<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%> 
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->
<!-- #include file="../Include/func.inc" -->
<!--#include file="../include/checkpower.asp"--> 
<!--#include file="../include/sideinfo.inc"--> 
<%
Set conn = GetSQLServerConnection()	  
self="YECP0302"  


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
 
function f()
	<%=self%>.YYMM.focus()	
	<%=self%>.YYMM.SELECT()
end function    
 
</SCRIPT>   
</head> 
<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload='f()' >
<form name="<%=self%>" method="post" action="EMPFILE.SALARY.ASP">
<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>

<table width="100%" BORDER=0 align=center cellpadding=0 cellspacing=0 >
	<tr>
		<td align="center">
			<table id="myTableForm" width="50%"> 
				<tr><td height="35px" colspan=4>&nbsp;</td></tr>
				<tr>
					<TD nowrap align=right>計薪年月<BR><font class=txt8>Tien Luong</font></TD>
					<TD ><INPUT type="text" style="width:100px" NAME=YYMM   VALUE="<%=calcmonth%>"></TD>
					<TD nowrap align=right>績效年月<BR><font class=txt8>Tien TThuong</font></TD>
					<TD ><INPUT type="text" style="width:100px" NAME=jxym   VALUE=""></TD>	
				</TR>		
				<TR>
					<TD nowrap align=right height=30 >國籍<BR><font class=txt8>Quoc Tich</font></TD>
					<TD >
						<select name=country style="width:120px"  >
							<%if Session("NETWHSNO")="ALL" or Session("RIGHTS")<="1" or Session("RIGHTS")="8" then%>
							<option value="">全部(Toan bo) </option>
							<% 
								SQL="SELECT * FROM BASICCODE WHERE FUNC='country' ORDER BY SYS_TYPE "
							else
								SQL="SELECT * FROM BASICCODE WHERE FUNC='country' and sys_type='VN' ORDER BY SYS_TYPE "
							end if 
							'SQL="SELECT * FROM BASICCODE WHERE FUNC='country'  ORDER BY SYS_type desc  "
							SET RST = CONN.EXECUTE(SQL)
							WHILE NOT RST.EOF  
							%>
							<option value="<%=RST("SYS_TYPE")%>" <%if rst("sys_type")="VN" then%>selected<%end if%>><%=RST("SYS_VALUE")%></option>				 
							<%
							RST.MOVENEXT
							WEND 
							rst.close
							%>
						</SELECT>
						<%SET RST=NOTHING %>
					</TD>
					<TD nowrap align=right height=30 >廠別<BR><font class=txt8>Xuong</font></TD>
					<TD > 
						<select name=WHSNO  style="width:120px" >
							<%if Session("RIGHTS")<="1" or Session("RIGHTS")="8" then%>
							<option value="">全部(Toan bo) </option>
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
				</TR>
				<TR >
					<TD nowrap align=right >組/部門<BR><font class=txt8>Bo Phan</font></TD>
					<TD >
						<select name=GROUPID   style="width:120px" >				
						<%if Session("RIGHTS")<="2" or Session("RIGHTS")="8" or Session("RIGHTS")="9" then%>
							<option value="">全部(Toan bo) </option>
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
							rst.close
						%>
						</SELECT>
						<%SET RST=NOTHING %>
					</td>
					<TD nowrap align=right >職等<BR><font class=txt8>Chuc vu</font></TD>			
					<TD >
						<select name=JOB    >	
						<option value="">全部(Toan bo) </option>		 
						<%SQL="SELECT * FROM BASICCODE WHERE FUNC='LEV'  ORDER BY SYS_TYPE "
						SET RST = CONN.EXECUTE(SQL)
						WHILE NOT RST.EOF  
						%>
						<option value="<%=RST("SYS_TYPE")%>" ><%=RST("SYS_VALUE")%></option>				 
						<%
						RST.MOVENEXT
						WEND 
						rst.close
						%>		 				 
						</SELECT>
						<%SET RST=NOTHING 
						conn.close
						set con=nothing
						%>
					</TD>
				</TR>
				<TR>
					<td nowrap align=right >員工編號<BR><font class=txt8>So The</font></td>
					<td>
						<input type="text" style="width:100px" name=empid1  maxlength=5 onchange=strchg(1)> 				
					</td>
					<td nowrap align=right >員工統計<BR><font class=txt8>Thong ke Nhan vien</font></td>
					<td>
						<select name=outemp style="width:120px">
							<option value="">--All--</option>			 		
							<option value="D">本月離職(Thoi viec thang nay)</option>
						</select>
					</td>
				</TR>	
				<TR >
					<td nowrap align=right >類別<BR><font class=txt8>Loai</font></td>
					<td colspan=3>
						<input type="radio" name=radio1 checked onclick="radiochg(0)">薪資Tien Luong 
						<input type="radio" name=radio1 onclick="radiochg(1)">績效Tien Thuong 
						<input type="hidden" value="A" name="rptType">
					</td>
				</TR>
				<tr >
					<td align=center height="50px" colspan=4>						
						<input type=button  name=btm class="btn btn-sm btn-outline-secondary" value="(Y)Print" onclick="go()" onkeydown="go()">&nbsp;&nbsp;
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
	if <%=self%>.rpttype.value="B" then 
		if <%=self%>.jxym.value="" then 
			alert "請輸入績效年月Xin danh lai tien thuong (YYYYMM)"
			<%=self%>.jxym.focus()
			exit function 
		end if 
	end if 		 
	if <%=self%>.rpttype.value="B" then 
		<%=self%>.action="<%=self%>.jx.toexcel.asp"
		<%=self%>.target="Back"
		<%=self%>.submit()
	else
		<%=self%>.action="<%=self%>.toexcel.asp"
		<%=self%>.target="Back"
		<%=self%>.submit()
	end if 	
	parent.best.cols="100%,0%"
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
	if <%=self%>.rpttype.value="B" then 
		if <%=self%>.jxym.value="" then 
			alert "請輸入績效年月Xin danh lai tien thuong (YYYYMM)"
			<%=self%>.jxym.focus()
			exit function 
		end if 
	end if 	
 	<%=self%>.action="<%=session("rpt")%>"&"yec/"&"<%=self%>.getrpt.asp"
 	<%=self%>.submit() 
end function   

function radiochg(index)	 
	if index= 0 then 	
		<%=self%>.rptType.value="A"
	end if
	if index = 1 then 
		<%=self%>.rptType.value="B"					
	end if 
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