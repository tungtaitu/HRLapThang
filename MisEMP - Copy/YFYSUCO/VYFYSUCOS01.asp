<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%> 
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->
<!-- #include file="../Include/func.inc" -->
<!--#include file="../include/sideinfo.inc"-->
<%
Set conn = GetSQLServerConnection()	  
self="vyfysucos01"  


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

if right(calcmonth,2)="01" then 
	sgym = left(calcmonth,4)-1 & "12" 
else
	sgym = left(calcmonth,4)&right("00"&right(calcmonth,2)-1,2)
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
	<%=self%>.YYMM1.focus()		
	
end function    
-->
</SCRIPT>   
</head> 
<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload='f()' >
<form name="<%=self%>" method="post"  >

	<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>	 
	<table width="100%" border=0 >
		<tr>
			<td>
				<table width="80%" BORDER=0 align=center cellpadding=0 cellspacing=0 >
					<tr>
						<td>
							<table class="txt" cellpadding=3 cellspacing=3> 
								<tr height=30 >
									<TD nowrap align=right>統計年月<br>Năm tháng thống kê</TD>
									<TD nowrap colspan=3>
										<INPUT  type="text" style="width:100px" NAME=YYMM1   VALUE="" SIZE=10 maxlength=6>~
										<INPUT  type="text" style="width:100px" NAME=YYMM2   VALUE="" SIZE=10 maxlength=6>
									</TD>
								</TR>
								<tr>
									<TD nowrap align=right>績效年月<br>Năm tháng kết thúc</TD>
									<TD><INPUT  type="text" style="width:100px" NAME=JXYM></TD>
									<TD nowrap align=right height=30 >國籍<br>Quốc tịch</TD>
									<TD nowrap>
										<select name=country   style='width:75'  >
											<option value="">全部-Tất cả </option>
											<%SQL="SELECT * FROM BASICCODE WHERE FUNC='country'  ORDER BY SYS_type desc  "
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
									</td>
								</tr>
								<TR>
									
									<TD nowrap align=right height=30 >廠別<br>Xưởng</TD>
									<TD > 
										<select name=WHSNO   >
											<option value="">全部-Tất cả</option>
											<%SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' ORDER BY SYS_TYPE "
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
									<TD nowrap align=right >組/部門<br>Nhóm/Bộ phận</TD>
									<TD>
										<select name=GROUPID  style="width:120px"  >
										<option value="" selected >全部-Tất cả </option>
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
									<td nowrap align=right >員工編號<br>Số thẻ</td>
									<td >
										<input  type="text" style="width:100px" name=empid1 maxlength=5 onchange=strchg(1)> 										
									</td>
									<TD nowrap align=right >扣款對象<br>Đối tượng khấu trừ</TD>			
									<TD >
										<select name=cfGroup style="width:100px">
											<option value="">全部-Tất cả</option>
											<option value="A">公司員工-Nhân viên</option>
											<option value="B">貨車司機-Tài xế nhà xe</option>
											<option value="C">廠商-Nhà cung ứng</option>
											<option value="D">原紙廠商-Nhà cung ứng giấy cuộn</option>					
										</select>
									</TD>									
								</TR> 
								<tr >
									<td align=center colspan=4>
										<input type=button  name=btm class="btn btn-sm btn-danger" value="確認 Xác nhận" onclick="go()" onkeydown="go()">
										<input type=reset  name=btm class="btn btn-sm btn-outline-secondary" value="取消 Hủy bỏ">				
									</td>
								</tr>
							</table>
						</td>
					</tr>					
				</table>
			</td>
		</tr>
	</table>
</form>
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
 	<%=self%>.action="<%=session("rpt")%>rpt/<%=SELF%>.GETRPT.asp"
 	<%=self%>.submit()
end function   

FUNCTION GETDATA()
	<%=self%>.action="<%=SELF%>.asp"
 	<%=self%>.submit()
END FUNCTION  
	

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


'_________________DATE CHECK___________________________________________________________________

function validDate(d)
	if len(d) = 8 and isnumeric(d) then
		d = left(d,4) & "/" & mid(d, 5, 2) & "/" & right(d,2)
		if isdate(d) then
			validDate = formatDate(d)
		else
			validDate = ""
		end if
	elseif len(d) >= 8 and isdate(d) then
			validDate = formatDate(d)
		else
			validDate = ""
		end if
end function

function formatDate(d)
		formatDate = Year(d) & "/" & _
		Right("00" & Month(d), 2) & "/" & _
		Right("00" & Day(d), 2)
end function
'________________________________________________________________________________________  

 
</script> 