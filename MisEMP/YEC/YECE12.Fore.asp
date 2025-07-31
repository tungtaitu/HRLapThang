<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%> 
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->
<!-- #include file="../Include/func.inc" -->
<!-- #include file="../Include/SIDEINFO.inc" -->
<%
Set conn = GetSQLServerConnection()	  
self="yece12"  


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

'response.write session("mywhsno")

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
<form name="<%=self%>" method="post" action="EMPFILEHW.SALARY.ASP">

	<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>
	<table width="100%" BORDER=0 align=center cellpadding=0 cellspacing=0 >
		<tr>
			<td align="center">
				<table id="myTableForm" width="50%">
					<tr><td height="40px">&nbsp;</td></tr>
					<tr >
						<TD nowrap align=right>計薪年月<br>Tien luong</TD>
						<TD ><INPUT type="text" style="width:100px" NAME=yymm   VALUE="<%=calcmonth%>"></TD>								
						<TD nowrap align=right >國籍<br>Quoc tich</TD>
						<TD >
							<select name=country  style="width:120px">					
								<%
								if request("CT")="VN" then 
									SQL="SELECT * FROM BASICCODE WHERE FUNC='country' and sys_type='VN'  ORDER BY SYS_type desc  "
								elseif request("CT")="CN" then 	
									SQL="SELECT * FROM BASICCODE WHERE FUNC='country' and sys_type='CN'  ORDER BY SYS_type desc  "
								elseif request("CT")="CT" then 	
									SQL="SELECT * FROM BASICCODE WHERE FUNC='country' and sys_type='CT'  ORDER BY SYS_type desc  "	
								elseif request("CT")="TM" then 	
									SQL="SELECT * FROM BASICCODE WHERE FUNC='country' and sys_type='X' ORDER BY SYS_type desc  "		
								else 	
									if session("rights")<="0" then 
										SQL="SELECT * FROM BASICCODE WHERE FUNC='country'  ORDER BY SYS_type desc  "		 
									else
										SQL="SELECT * FROM BASICCODE WHERE FUNC='country' and sys_type='VN' ORDER BY SYS_type desc  "		 
									end if 	
								end if 	
								SET RST = CONN.EXECUTE(SQL)
								WHILE NOT RST.EOF  
								%>
								<option value="<%=RST("SYS_TYPE")%>" ><%=RST("SYS_TYPE")%>-<%=RST("SYS_VALUE")%></option>				 
								<%
								RST.MOVENEXT
								WEND 
								%>
								<%if request("CT")="TM"  then %>
								<option value="TW">TW</option>
								<%end if%>								
							</SELECT>
							<%SET RST=NOTHING %>
						</TD>	
					</tr>
					<tr>		 
						<TD nowrap align=right >廠別<br>Xuong</TD>
						<TD > 
							<select name=whsno style="width:120px" >
								<option value="">--ALL---</option>
								<%SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' ORDER BY SYS_TYPE "
								SET RST = CONN.EXECUTE(SQL)
								WHILE NOT RST.EOF  
								%>
								<option value="<%=RST("SYS_TYPE")%>" <%if session("mywhsno")=RST("SYS_TYPE") then%>selected<%end if%>><%=RST("SYS_VALUE")%></option>				 
								<%
								RST.MOVENEXT
								WEND 
								%>					
							</SELECT>
							<%SET RST=NOTHING %>
						</TD>								
						<TD nowrap align=right >組/部門<br>Bo phan</TD>
						<TD >
							<select name=groupid  style="width:120px"  >
							<option value="" selected >--ALL---</option>
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
					<TR >
						<td nowrap align=right >員工編號<br>So the</td>
						<td>
							<input type="text" style="width:100px" name=empid1  maxlength=6 onchange=strchg(1)> 										
						</td>								
						<td nowrap align=right >員工統計<br>Loai</td>
						<td>
							<select name=outemp  style="width:120px"> 			 		
								<option value="">在職(ALL CN)</option>
								<option value="D">本月離職(thoi viec thang nay)</option>
							</select>	
						</td>
					</TR>	
					<%if request("CT")="VN" or request("CT")="CT" then %>		
					<TR>
						<td nowrap align=right >年假代金<br>Tien thuong phep<br>(only VN)</td>
						<td colspan=5>
							計算(tinh)<input  type="text" style="width:120px" name="NJYY" class="inputbox" maxlength=4  >(Nam)年假未修代金
						</td>
					</TR>		
					<%else%>
						<input size=5 name="NJYY" class="inputbox" maxlength=4  type="hidden" >
					<%end if%>
					<tr >
						<td align=center colspan=6 height="50px">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<input type=button  name=btm class="btn btn-sm btn-danger" value="(Y)Confirm" onclick="go()" onkeydown="go()">
							<input type=reset  name=btm class="btn btn-sm btn-outline-secondary" value="(N)Cancel">
							&nbsp;&nbsp;&nbsp;&nbsp;
							<input type=button  name=btn class="btn btn-sm btn-outline-secondary" value="(T)特別獎金" onclick="gob()" >
						</td>
					</tr>
				</table>
			</td>
		</tr>		
	</table>
			
</form>
</body>
</html>

<script  type="text/javascript">
function gob(){
	var m = document.forms[0];
	var c1 = m.yymm.value ; 
	var c2 = m.country.value ; 
	var c3 = m.whsno.value ; 
	var c4 = m.groupid.value ; 
	var c5 = m.empid1.value ; 
	//alert (c1);
	window.open ("yece12.foreB.asp?yymm="+c1+"&country="+c2+"&whsno="+c3+"&g1="+c4+"&eid="+c5,"_blank","top=100,left=150,width=800,height=500,scrollbars=yes,resizable=yes" ) ;
}
</script>

<script language=vbs> 

function chksts()
	if <%=self%>.chk1.checked=true then 
		<%=self%>.recalc.value="Y"
	else
		<%=self%>.recalc.value="N"
	end if 
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

	if <%=self%>.whsno.value="" then 
		alert "Hay chon xuong"
		<%=self%>.whsno.focus()
	else	
		<%=self%>.action="<%=SELF%>.SALARY.asp"
		<%=self%>.submit() 
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