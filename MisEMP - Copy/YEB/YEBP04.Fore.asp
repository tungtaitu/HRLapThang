<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%> 
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->
<!-- #include file="../Include/func.inc" -->
<!--#include file="../include/checkpower.asp"-->
<!-- #include file="../Include/SIDEINFO.inc" -->
<%
Set conn = GetSQLServerConnection()	  
self="YEBP04"  
'YEBP04.getrpt.asp

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
	'<%=self%>.country.focus()	
	'<%=self%>.country.SELECT()
	<%=self%>.inym.focus()
end function    
-->
</SCRIPT>   
</head> 
<body  topmargin="50" leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload='f()' >
<form name="<%=self%>" method="post"  >
 
	<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>
	
				<table width="100%" BORDER=0 align=center cellpadding=0 cellspacing=0 >
					<tr>
						<td align="center">
							<table id="myTableForm" width="50%"> 
								<tr><td colspan=4 height="40px">&nbsp;</td></tr>
								<TR>
									<TD nowrap align=right width="10%">到職年月<br><font class=txt8>Thong ke Thang Nam</font></TD>
									<TD  width="40%">
										 <input name=inym  size=15 maxlength=6  value=""> 
									</TD>								
									<TD nowrap align=right  width="10%">國籍<br><font class=txt8>Quoc Tich</font></TD>
									<TD  width="40%">
										<select name=country  >
											<option value="">ALL </option>
											<%SQL="SELECT * FROM BASICCODE WHERE FUNC='country'  ORDER BY SYS_type desc  "
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
										<%SET RST=NOTHING %>
									</TD>	
								</tr>
								<tr>		 
									<TD nowrap align=right height=30 >廠別<br><font class=txt8>Loai Xuong</font></TD>
									<TD > 
										<select name=WHSNO   >
											<option value="">全部 </option>
											<%SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' ORDER BY SYS_TYPE "
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
									<TD nowrap align=right >組/部門<br><font class=txt8>Don Vi</font></TD>
									<TD >
										<select name=GROUPID    >
										<option value="" selected >全部ALL</option>
										<option value="" selected > </option>
										<option value="WB"   >WB-外包THẦU NGOÀI</option>
										<option value="" selected ><hr size=0></option>
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
										rst.close					
										%>
										</SELECT>
										<%SET RST=NOTHING %>
									</td>	
								</tr>								
								<TR>
									<td nowrap align=right >員工編號<br><font class=txt8>So The</font></td>
									<td colspan=3>
										<input name=empid1  size=6 maxlength=5 onchange=strchg(1)>~
										<input name=empid2  size=6 maxlength=5 onchange=strchg(2)>
									</td>
								</TR>
								<TR>
									<td nowrap align=right >外包類別<br><font class=txt8>Loai THẦU NGOÀI</font></td>
									<td colspan=3>
										<select name=loai    >	
										<option value="">全部 </option>		 
										<%SQL="SELECT * FROM BASICCODE WHERE FUNC='WB'  ORDER BY SYS_TYPE "
										SET RST = CONN.EXECUTE(SQL)
										WHILE NOT RST.EOF  
										%>
										<option value="<%=RST("SYS_TYPE")%>" ><%=RST("SYS_TYPE")%>-<%=RST("SYS_VALUE")%></option>				 
										<%
										RST.MOVENEXT
										WEND 
										rst.close
										%>		 				 
										</SELECT>
										<%SET RST=NOTHING %>
										<%
										conn.close 
										set conn=nothing 
										%>
									</td>
								</TR>
								<tr >
									<td align=center colspan=4 height="60px">
										<input type=button  name=btm class="btn btn-sm btn-danger" value="(Y)Confirm" onclick="go()" onkeydown="go()">
										<input type=reset  name=btm class="btn btn-sm btn-outline-secondary" value="(N)Cancel">
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
 	<%=self%>.action="<%=session("rpt")%>"&"yeb/"&"<%=self%>.getrpt.asp"
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