<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%> 
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->
<!-- #include file="../Include/func.inc" -->
<!--#include file="../include/sideinfo.inc"-->
<%
Set conn = GetSQLServerConnection()	  
self="YEBBP03" 
%>

<html> 
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css">
<SCRIPT  LANGUAGE=javascript>

function f(){
	<%=self%>.country.focus();	
}   

</SCRIPT>   
</head> 
<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload='f()' >
<form name="<%=self%>" method="post" action="emp_basicbiao.getrpt.asp">

	<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>
	
				<table width="100%" BORDER=0 align=center cellpadding=0 cellspacing=0 >
					<tr>
						<td align="center">
							<table id="myTableForm" width="60%"> 
								<tr><td colspan=4 height="40px">&nbsp;</td></tr>
								<tr>
									<TD nowrap align=right>國籍：</TD>
									<TD>
										<select name=country style="width:120px">
											<%if Session("NETWHSNO")="ALL" or Session("RIGHTS")<="2" or Session("RIGHTS")="8" then%>
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
									<td nowrap align=right >員工編號：</td>
									<td nowrap>
										<input type="text"  style="width:100px" name=empid1 maxlength=5 onchange=strchg(1)>~
										<input type="text"  style="width:100px" name=empid2 maxlength=5 onchange=strchg(2)>
									</td>
								</TR>
								<TR>
									<TD nowrap align=right>廠別：</TD>
									<TD> 
										<select name=WHSNO  style="width:120px">
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
										</SELECT>
										<%SET RST=NOTHING %>
									</TD>
									<td nowrap align=right >到職日期：</td>
									<td nowrap>
										<input type="text"  style="width:100px" name=indat1 maxlength=10 onblur="date_change(1)">~
										<input type="text"  style="width:100px" name=indat2 maxlength=10 onblur="date_change(2)">
									</td>
								</TR>
								<TR>
									<TD nowrap align=right >組/部門：</TD>
									<TD >
										<select name=GROUPID  style="width:130px">
										<option value="" selected >全部 </option>
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
									<td nowrap align=right >簽約日期：</td>
									<td nowrap>
										<input type="text"  style="width:100px" name=bhdat1 maxlength=10 onblur="date_change(3)">~
										<input type="text"  style="width:100px" name=bhdat2 maxlength=10 onblur="date_change(4)">	
									</td>
								</TR>								
								<TR>
									<TD nowrap align=right >職等：</TD>			
									<TD >
										<select name=JOB   style="width:180px">	
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
										<%SET RST=NOTHING 
										conn.close 
										set conn=nothing 
										
										%>
									</TD>
									<td nowrap align=right >員工簽約統計：</td>
									<td>
										<select  name=empTJ  style="width:120px">					
											<option value="">全部</option>
											<option value="A">待(未)簽約</option>
											<option value="B">已簽約</option>
											<option value="C">需續約</option>
										</select>
									</td>									
								</TR>
								<tr>
									<td align=center colspan=4 height="50px">
										<input type=button  name=btm class="btn btn-sm btn-danger" value="確   認" onclick="go()" onkeydown="go()">
										<input type=reset  name=btm class="btn btn-sm btn-outline-secondary" value="取   消">
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
elseif a=3 then
	INcardat = Trim(<%=self%>.bhdat1.value)
elseif a=4 then
	INcardat = Trim(<%=self%>.bhdat2.value)		
end if		
		    
IF INcardat<>"" THEN
	ANS=validDate(INcardat)
	IF ANS <> "" THEN
		if a=1 then
			Document.<%=self%>.indat1.value=ANS			
		elseif a=2 then
			Document.<%=self%>.indat2.value=ANS		 			
		elseif a=3 then
			Document.<%=self%>.bhdat1.value=ANS	
		elseif a=4 then
			Document.<%=self%>.bhdat2.value=ANS			
		end if		
	ELSE
		ALERT "EZ0067:輸入日期不合法 !!" 
		if a=1 then
			Document.<%=self%>.indat1.value=""
			Document.<%=self%>.indat1.focus()
		elseif a=2 then
			Document.<%=self%>.indat2.value=""
			Document.<%=self%>.indat2.focus()
		elseif a=3 then
			Document.<%=self%>.bhdat1.value=""
			Document.<%=self%>.bhdat1.focus()
		elseif a=4 then
			Document.<%=self%>.bhdat2.value=""
			Document.<%=self%>.bhdat2.focus()		
		end if		
		EXIT FUNCTION
	END IF
		 
ELSE
	'alert "EZ0015:日期該欄位必須輸入資料 !!" 		
	EXIT FUNCTION
END IF 
   
END FUNCTION
</script> 