<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%> 
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->
<!-- #include file="../Include/func.inc" -->
<!--#include file="../include/sideinfo.inc"-->
<%
'Set conn = GetSQLServerConnection()	  
self="YEBBB0202"  


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
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css">
<SCRIPT LANGUAGE=javascript>

function f(){
	<%=self%>.empid1.focus();
}   

</SCRIPT>   
</head> 
<body onkeydown="enterto()" onload='f()' >
<form name="<%=self%>" method="post"  >
	<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>
	
	<table BORDER=0 align=center cellpadding=3 cellspacing=3 >
		<tr>
			<td align="center">
				<table class="txt" >
					<tr>
						<td nowrap align=right >員工編號-So The</td>
						<td>
							<input type="text" style="width:150px" name="empid1" size=15 maxlength=5 onchange=strchg(1)> 				
						</td>	
						<td align=center>
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


<script language=javascript>

	function strchg(a){
		if (a==1) 
			<%=self%>.empid1.value = <%=self%>.empid1.value.toUpperCase();
		else if (a==2) 	
			<%=self%>.empid2.value = <%=self%>.empid2.value.toUpperCase();
			
	} 
	
	function go() {
		if(<%=self%>.empid1.value=="")
		{
			alert("請輸入員工編號!!");
			<%=self%>.empid1.focus();
		}else{	
			<%=self%>.action="<%=self%>.ForeGnd.asp";
			<%=self%>.submit();
		}		
	} 
	

</script> 