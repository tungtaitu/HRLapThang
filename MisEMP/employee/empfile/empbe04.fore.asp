<%@LANGUAGE=VBSCRIPT CODEPAGE=950%> 
<!-- #include file = "../../GetSQLServerConnection.fun" --> 
<!-- #include file="../../ADOINC.inc" -->
<!-- #include file="../../Include/func.inc" -->
<%
Set conn = GetSQLServerConnection()	  
self="empbe04"  


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
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
<link rel="stylesheet" href="../../Include/style.css" type="text/css">
<link rel="stylesheet" href="../../Include/style2.css" type="text/css">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
function f()
	<%=self%>.empid(0).focus()	
	'<%=self%>.country.SELECT()
end function    
-->
</SCRIPT>   
</head> 
<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload='f()' >
<form name="<%=self%>" method="post"  >
<table width="460" border="0" cellspacing="0" cellpadding="0">
	<tr><TD colspan=3>
	<img border="0" src="../../image/icon.gif" align="absmiddle">
	幹部合同資料維護</TD>
	</tr>
	<tr><td colspan=3><hr size=0	style='border: 1px dotted #999999;' align=left width=500></td></tr>
	<tr height=40 >
		<td width=150 align=center valign=middle>
			<font color="Brown"><b>幹部合同資料新增</b></font>
		</td>
		<td width=180 align=center>
			<img border="0" src="../../picture/icon02.gif" align="absmiddle"> 
			<a href="empbe0401.asp" target="_parent">幹部合同資料異動查詢</a>
		</td>
		<td></td>
	</tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=500>		
<table width=500  ><tr><td >
	<table width=450 class=txt9>
		<tr bgcolor="#DCDCDC" height=25>
			<td align=center>工號</td>
			<td align=center>姓名</td>
			<td align=center>到職日期</td>
			<td align=center>合同期(起)</td>
			<td align=center>合同期(迄)</td>			
			<td align=center>備註</td>
		</tr>
		<% for i = 1 to 10 %>
		<tr>
			<td>
				<input name=empid size=6 class=inputbox ondblclick="getempdata(<%=i-1%>)"  onchange="chkempid(<%=i-1%>)">
			</td>	
			<td>
				<input name=empname size=15 class=readonly8 readonly >
				<input type=hidden name=country  >
			</td>	
			<td>
				<input name=indate size=10 class=readonly8 readonly >
			</td>	
			<td>
				<input name=dat1 size=11 class=inputbox onblur="dat1chg(<%=i-1%>)">
			</td>	
			<td>
				<input name=dat2 size=11 class=inputbox onblur="dat2chg(<%=i-1%>)" >
			</td>				
			<td>
				<input name=memo size=25 class=inputbox onblur="visanochg(<%=i-1%>)" >
			</td>	
		</tr>
		<%next%>
	</table>	
	<table width=450 align=center>
		<tr >
			<td align=center>
				<input type=button  name=btm class=button value="確   認" onclick="go()" onkeydown="go()">
				<input type=reset  name=btm class=button value="取   消">
			</td>
		</tr>	
	</table>	

</td></tr></table> 

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
			if 	<%=self%>.dat1(x-1).value="" then 
				alert "請輸入 "&<%=self%>.empid(x-1).value&" 有效期(起)!!"
				<%=self%>.dat1(x-1).focus()
				exit function
			elseif 	<%=self%>.dat2(x-1).value="" then 
				alert "請輸入 "&<%=self%>.empid(x-1).value&" 有效期(迄)!!"
				<%=self%>.dat2(x-1).focus()
				exit function			
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