<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%> 
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->
<!-- #include file="../Include/func.inc" -->
<!--#include file="../include/checkpower.asp"--> 
<!--#include file="../include/sideinfo.inc"--> 
<%
Set conn = GetSQLServerConnection()	  
self="YEcb01"  


nowmonth = year(date())&right("00"&month(date()),2)  
if month(date())="01" then  
	calcmonth = year(date())-1&"12"  	
else
	calcmonth =  year(date())&right("00"&month(date())-1,2)   
end if 	   

if day(date())<=11 then 
	if month(date())="01" then  
		calcmonth = year(date())-1&"12" 
	else
		calcmonth =  year(date())&right("00"&month(date())-1,2)   
	end if 	 	
else
	calcmonth = nowmonth 
end if 

if session("netuser")="" then 
	response.write "UserID is Empty Please Login again !!!<BR>"
	response.write "Vao mang trong rong , hoac doi lau , hay nhan nut nhap mang tu dau !!! "
	response.end 
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
	<%=self%>.closeym.focus()	
	<%=self%>.closeym.SELECT()
end function    
-->
</SCRIPT>   
</head> 
<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload='f()' >
<form name="<%=self%>" method="post" action="EMPFILE.SALARY.ASP">

	<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>
	
	<table width="70%" BORDER=0 align=center cellpadding=0 cellspacing=0 >					
		<tr>
			<td>
				<table class="txt"> 
					<tr>
						<td>&nbsp;</td>
						<TD ><font color=blue>
							1. 請確認所有資料正確後執行關帳<BR>
							2. 關帳後薪資等相關資料無法異動(所有國籍),不能修改<BR>
							<font color=red>3. 必須關帳才能列印當月薪資單</font><BR>
							4. 關帳後自動備份資料<BR>
							</font>
						</TD>
					</TR>		 
				</table>
			</td>
		</tr>
		<tr>
			<td>
				<table class="txt" cellpadding=3 cellspacing=3>
					<tr>
						<TD align=right nowrap>關帳年月：</TD>
						<TD><INPUT type="text" style="width:100px" NAME=CloseYM VALUE="<%=calcmonth%>" maxlength=6></TD>
						<TD align=right nowrap>廠別：</TD>
						<TD>
							<select name=WHSNO style="width:150px">					
								<%
								'if session("rights")=0 then 
								'	SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' and sys_type='LA' ORDER BY SYS_TYPE "
								'else
								'	SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' and sys_type='"& session("NETWHSNO") &"' ORDER BY SYS_TYPE "
								'end if	
								SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' ORDER BY SYS_TYPE "
								SET RST = CONN.EXECUTE(SQL)
								WHILE NOT RST.EOF
								%>
								<option value="<%=RST("SYS_TYPE")%>"><%=RST("SYS_VALUE")%></option>
								<%
								RST.MOVENEXT
								WEND
								SET RST=NOTHING
								%>
							</select>
						</TD>
						<td>
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

function strchg(a)
	if a=1 then 
		<%=self%>.empid1.value = Ucase(<%=self%>.empid1.value)
	elseif a=2 then 	
		<%=self%>.empid2.value = Ucase(<%=self%>.empid2.value)
	end if 	
end function 
	
function go()  
	if <%=self%>.closeym.value="" then 
		alert "請輸入關帳年月"
		<%=self%>.closeym.focus()
		exit function 
	elseif len(<%=self%>.closeym.value)<>6 then 
		alert "關帳年月輸入錯誤!!"
		<%=self%>.closeym.value=""
		<%=self%>.closeym.focus()
		exit function 
	end if	
	if <%=self%>.WHSNO.value ="" then 
		alert "請選擇廠別"
		<%=self%>.WHSNO.focus()
		exit function 
	end if 	
 	<%=self%>.action="<%=self%>.upd.asp"
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