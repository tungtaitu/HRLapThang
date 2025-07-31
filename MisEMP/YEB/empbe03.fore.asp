<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%> 
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->
<!-- #include file="../Include/func.inc" -->
<!-- #include file="../Include/SIDEINFO.inc" -->
<%
'Set conn = GetSQLServerConnection()	  
self="empbe03"  


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
	<%=self%>.empid(0).focus()	
	'<%=self%>.country.SELECT()
end function    
-->
</SCRIPT>   
</head> 
<body  onkeydown="enterto()" onload='f()' >
<form name="<%=self%>" method="post"  >

	<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>
	<table width="100%" border=0 >
		<tr>
			<td>
				<table width="90%" BORDER=0 align=center cellpadding=0 cellspacing=0 >					
					<tr>
						<td class="p-3 bg-lightgray ">
							<h3 >簽證資料新增 <font class="text-secondary fa" style="font-size:1.3rem;">THÊM MỚI DỮ LIỆU VISA</font></h3>
							<table id="myTableGrid" width="98%">								
								<tr bgcolor="#DCDCDC" height=25>
									<td align=center>工號<br>Số thẻ</td>
									<td align=center>姓名<br>Họ tên</td>
									<td align=center>簽証號碼<br>Số VISA</td>
									<td align=center>有效期(起)<br>Từ ngày</td>
									<td align=center>有效期(迄)<br>Đến ngày</td>
									<td align=center>費用(VND)<br>Chi phí(VND)</td>
									<td align=center>備註<br>Ghi chú</td>
								</tr>
								<% for i = 1 to 10 %>
								<tr>
									<td>
										<input name=empid size=6 class="form-control form-control-sm mb-2 mt-2" ondblclick="getempdata(<%=i-1%>)"  onchange="chkempid(<%=i-1%>)">
									</td>	
									<td>
										<input name=empname size=15 class="form-control form-control-sm mb-2 mt-2" readonly >
										<input type="hidden" name="f_country" value="" >
									</td>	
									<td>
										<input name=visano size=10 class="form-control form-control-sm mb-2 mt-2"  onblur="visanochg(<%=i-1%>)" maxlength=9>
									</td>	
									<td>
										<input name=dat1 size=11 class="form-control form-control-sm mb-2 mt-2" onblur="dat1chg(<%=i-1%>)">
									</td>	
									<td>
										<input name=dat2 size=11 class="form-control form-control-sm mb-2 mt-2" onblur="dat2chg(<%=i-1%>)" >
									</td>	
									<td>
										<input name=visaAmt size=10 class="form-control form-control-sm mb-2 mt-2" onblur="amtchg(<%=i-1%>)" value="0" style='text-align:right'>
									</td>	
									<td>
										<input name=memo size=15 class="form-control form-control-sm mb-2 mt-2" onblur="visanochg(<%=i-1%>)" >
									</td>	
								</tr>
								<%next%>
							</table>
						</td>
					</tr>
					<tr>
						<td align=center>
							<table class="table-borderless table-sm bg-white text-secondary">
								<tr >
									<td>
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
function getempdata(index) 
	ncols="visano"
	open "Getempdata.asp?pself="& "<%=self%>" &"&index=" & index &"&ncols="& ncols , "Back" 
	parent.best.cols="50%,50%"
end function  


function chkempid(index)	
	if <%=self%>.empid(index).value<>"" then 
		code1=Ucase(trim(<%=self%>.empid(index).value))
		open "<%=self%>.back.asp?func=chkempid&index=" & index &"&code1=" & code1 , "Back" 
		'parent.best.cols="70%,30%"
	end if 
end  function 

function visanochg(index)	
	<%=self%>.visano(index).value=Ucase(<%=self%>.visano(index).value)
	<%=self%>.memo(index).value=Ucase(<%=self%>.memo(index).value)
end  function  

function amtchg(index)	
	if <%=self%>.visaAmt(index).value<>"" then 	
		if isnumeric(<%=self%>.visaAmt(index).value)=false then 
			alert "請輸入數字!!"
			<%=self%>.visaAmt(index).value="0"
			<%=self%>.visaAmt(index).focus()
			<%=self%>.visaAmt(index).select()
		end  if 
	end if 		
end  function   


function go() 
	empstr=""
	for x = 1 to 10 
		if <%=self%>.empid(x-1).value<>"" then 
			if <%=self%>.visaNo(x-1).value="" then 
				alert "請輸入 "&<%=self%>.empid(x-1).value&" 簽證號碼!!"
				<%=self%>.visaNo(x-1).focus()
				exit function
			elseif 	<%=self%>.dat1(x-1).value="" then 
				alert "請輸入 "&<%=self%>.empid(x-1).value&" 有效期(起)!!"
				<%=self%>.dat1(x-1).focus()
				exit function
			elseif 	<%=self%>.dat2(x-1).value="" then 
				alert "請輸入 "&<%=self%>.empid(x-1).value&" 有效期(迄)!!"
				<%=self%>.dat2(x-1).focus()
				exit function
			'elseif 	<%=self%>.visaamt(x-1).value=""  or  <%=self%>.visaamt(x-1).value="0" then 
			'	alert "請輸入 "&<%=self%>.empid(x-1).value&" 簽證費用!!"
			'	<%=self%>.visaamt(x-1).focus()	
			'	<%=self%>.visaamt(x-1).select()
			'	exit function
			end if 
		end  if
		empstr = empstr & Ucase(<%=self%>.empid(x-1).value)
	next 
	if len(empstr)=0 then 
		alert "請輸入資料!!"
		<%=self%>.empid(0).focus()
		exit function 
	else	
	 	<%=self%>.action="<%=self%>.Upd.asp"
	 	<%=self%>.submit() 
	 end  if 	
end function   
	

'*******檢查日期*********************************************
FUNCTION dat1chg(index)	

	INcardat = Trim(<%=self%>.dat1(index).value)  		
		    
IF INcardat<>"" THEN
	ANS=validDate(INcardat)
	IF ANS <> "" THEN		
		Document.<%=self%>.dat1(index).value=ANS					
	ELSE
		ALERT "EZ0067:輸入日期不合法 !!" 		
		Document.<%=self%>.dat1(index).value=""
		Document.<%=self%>.dat1(index).focus() 		
		EXIT FUNCTION
	END IF
		 
ELSE
	'alert "EZ0015:日期該欄位必須輸入資料 !!" 		
	EXIT FUNCTION
END IF 
   
END FUNCTION 

FUNCTION dat2chg(index)	
	INcardat = Trim(<%=self%>.dat2(index).value)  		
		    
IF INcardat<>"" THEN
	ANS=validDate(INcardat)
	IF ANS <> "" THEN		
		Document.<%=self%>.dat2(index).value=ANS					
	ELSE
		ALERT "EZ0067:輸入日期不合法 !!" 		
		Document.<%=self%>.dat2(index).value=""
		Document.<%=self%>.dat2(index).focus() 		
		EXIT FUNCTION
	END IF
		 
ELSE
	'alert "EZ0015:日期該欄位必須輸入資料 !!" 		
	EXIT FUNCTION
END IF    
END FUNCTION 
</script> 