<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%> 
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->
<!-- #include file="../Include/func.inc" -->
<!--#include file="../include/sideinfo.inc"-->
<%
Set conn = GetSQLServerConnection()	  
self="YEBBB0101"  


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
  
</head> 
<body   onkeydown="enterto()">
<form name="<%=self%>" method="post"  >
	<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>
	<table width="100%" border=0 >
		<tr>
			<td align="center">
				<table id="myTableForm" width="50%">
					<tr>
						<td align="center" colspan=2>&nbsp;</td>
					</tr>
					<TR>
						<TD align=right width="20%">國籍<br>Quoc Tich</TD>
						<TD >
							<select name=country style="width:98%">
								<option value="">----</option>
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
						</TD>	
					</tr>
					<tr>		 
						<TD align=right  >廠別<br>Xuong</TD>
						<TD > 
							<select name=WHSNO   >
								<option value="">----</option>
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
					</TR>	 
					<TR>
						<TD nowrap align=right >組/部門<br>Bo Phan</TD>
						<TD >
							<select name=GROUPID    >
							<option value="" selected >----</option>
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
					<TR  >
						<td nowrap align=right >員工編號<br>So the</td>
						<td>
							<input type="text"  name=empid1  size=15 maxlength=5 onchange="strchg(1)"> 							
						</td>
					</TR>
					<TR>
						<td nowrap align=right >簽約統計<br>ky hop dong</td>
						<td >
							<select name=outemp > 
								<option value="">----</option>
								<option value="Y">Da ky hop Dong(已簽約)</option>
								<option value="N">Chua ky hop dong(未簽約)</option>
							</select>	
						</td>
					</TR>
					<TR>
						<td nowrap align=right >員工統計<br>Thong ke </td>
						<td >
							<select name=IOemp > 
								<option value="Y">Tai Chuc(在職)</option>
								<option value="">ALL全部</option>
								<option value="N">Thoi Viec(已離職)</option>
							</select>	
						</td>
					</TR>
					<tr >
						<td colspan=2 align=center>
							<input type=button  name=btm class="btn btn-sm btn-danger" value="(Y)Confirm" onclick="go()" onkeydown="go()">
							<input type=reset  name=btm class="btn btn-sm btn-outline-secondary" value="(N)Cancel">
						</td>
					</tr>
					<tr>
						<td align="center" colspan=2>&nbsp;</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
	 
	<%
	conn.close 
	set conn=nothing
	%>


</body>
</html>


<script language=javascript>

function strchg(a){
	if (a==1) 
		<%=self%>.empid1.value = <%=self%>.empid1.value.toUpperCase();
	else if(a==2) 	
		<%=self%>.empid2.value = <%=self%>.empid2.value.toUpperCase();
} 
	
function go(){ 
 	<%=self%>.action="<%=self%>.ForeGnd.asp";
 	<%=self%>.submit();
} 

</script> 