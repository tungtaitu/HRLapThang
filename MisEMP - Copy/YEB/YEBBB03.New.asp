<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%> 
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->
<!-- #include file="../Include/func.inc" -->
<!--#include file="../include/sideinfo.inc"-->
<%
Set conn = GetSQLServerConnection()	  
self="yebbb03"  


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

empid=request("empid")
if empid<>"" then 
	sql="select *from view_empfile where empid='"& empid &"' "
	'response.write sql
	set rs=conn.execute(sql)
	if rs.eof then %>
		<script language=javascript>
			alert("員工編號輸入錯誤!!")
			open("<%=self%>.new.asp", "_self");
		</script>
<%	
	else
		empid=rs("empid")
		empname=rs("empnam_cn")&" "&rs("empnam_vn")
		indat=rs("nindat")
		gstr = rs("gstr")&"-"&rs("zstr")
		jstr = rs("jstr")&" " 
	end if 
end if 	
set rs=nothing  

rp_dat = request("rp_dat") 
  
rp_type = request("rp_type")
rp_func = request("rp_func")
rp_method = request("rp_method")
rp_memo = request("rp_memo")
fileno = request("fileno")
WHSNO = request("WHSNO")
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
	if(<%=self%>.empid.value=="") 
		<%=self%>.whsno.focus()	;
	else
		<%=self%>.rp_dat.focus();	
	
}  

function gotemp(){
	open("../getempdata.asp?formName="+"<%=self%>", "Back");
	parent.best.cols="65%,35%";
}  

function empidchg(){
	if (<%=self%>.empid.value !="" || <%=self%>.rp_type.value !=""){ 
	 	<%=self%>.action = "<%=self%>.new.asp";
	 	<%=self%>.submit();
	} 	
}
 
</SCRIPT>   
</head> 
<body onload='f()' >
<form name="<%=self%>" method="post"  >
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=hidden NAME=nowmonth VALUE="<%=nowmonth%>">

	<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>	
	<table width="94%" align="center">
		<tr>
			<td>
				<table class="txt" cellpadding=3 cellspacing=3> 		
					<tr>
						<Td align=right>工號So the</td>
						<td>
							<input type="text" style="width:100px" name="empid" size=8  ondblclick='gotemp()' onblur='empidchg()' value="<%=empid%>">
						</td>								
						<td align=right>姓名Ho ten</td>
						<td><input type="text" style="width:150px" name="empname" size=25  readonly class="readonly" value="<%=empname%>"></td>	
						<Td align=right>到職日NVX</td>
						<td><input type="text" style="width:120px" name="indat" size=15  readonly class="readonly" value="<%=indat%>"></td>
					</tr>
					<tr>
						<Td align=right>所屬廠別Xuong</td> 
						<TD> 
							<select name=WHSNO  >
								<option value="">全部 </option>
								<%SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' and sys_type like '"& session("rpwhsno") &"%' ORDER BY SYS_TYPE "
								SET RST = CONN.EXECUTE(SQL)
								WHILE NOT RST.EOF  
								%>
								<option value="<%=RST("SYS_TYPE")%>" <%if whsno=RST("SYS_TYPE") then%>selected<%end if%>><%=RST("SYS_VALUE")%></option>				 
								<%
								RST.MOVENEXT
								WEND 
								%>
							</SELECT>
							<%SET RST=NOTHING %>
						</TD>
						<Td align=right>單位Don Vi</td>
						<td>
							<input type="text" name="gstr" size=15  readonly class="readonly" value="<%=gstr%>">						
						</td>
						<td align=right>職務Chuc vu</td>
						<td><input type="text" name="jstr" size=18  readonly class="readonly" value="<%=jstr%>"></td>
					</tr>			
					<tr>
						<Td align=right>獎懲類別thuong phat</td>
						<td colspan=2>						
							<select name=rp_type   onchange='empidchg()'>
								<option value="" <%if rp_type="" then%>selected<%end if%>>---</option>
								<option value="R" <%if rp_type="R" then%>selected<%end if%>>R-獎thuong</option>
								<option value="P" <%if rp_type="P" then%>selected<%end if%>>P-懲phat</option>
							</select>
						</td>									
						<td>
							<select name=rp_func  onchange='empidchg()'>
								<%if rp_type="" then%>
									<option value="">---</option>
								<%else%>
									<%
									if rp_type="R" then 
										sql="select* from basicCode  where func='goods' order by sys_type"
									elseif 	rp_type="P" then 
										sql="select* from basicCode  where func='bads' order by sys_type"
									end if 	
									SET RST=COnn.execute(sql)
									while not rst.eof 
									%>
									<option value="<%=rst("sys_type")%>" <%if trim(rp_func)=trim(rst("sys_type")) then %>selected<%end if%>><%=rst("sys_type")%>-<%=rst("sys_value")%></option>
									<%rst.movenext%>
									<%wend
									rst.close
									set rst=nothing
								end if 
								conn.close 
								set conn=nothing
									%> 
							</select>
						</td>
						<Td align=right>事件日期Ngay</td>
						<td>
							<input type="text" name="rp_dat" id="rp_dat" size=15  value="<%=rp_dat%>" onblur="chkdat('rp_dat')">
						</td>
					</tr>				 		
				</table>
			</td>
		</tr>
		<tr>
			<td align="center">
				<table  id="myTableForm" width="98%">
					<tr>
						<td height="40px">&nbsp;&nbsp;處理方式說明phuong thuc xu ly </td>
					</tr>
					<tr>
						<td align=center><input type="text" style="width:98%" name=rp_method  size=75 maxlength=255 value=<%=rp_method%>> </td>
					</tr>
					<tr>
						<td  height="30px">&nbsp;&nbsp;事件內容說明thuyet minh</td>
					</tr>
					<tr>
						<td align=center>
							<TEXTAREA rows=10 cols=75 name=rp_memo  STYLE='HEIGHT:AUTO' wrap="PHYSICAL"><%=rp_memo%></TEXTAREA>				
						</td>
					</tr>	
					<tr>
						<td height="30px">&nbsp;&nbsp;文件編號</td>
					</tr>
					<tr>
						<td align=center>
							<input type="text" style="width:98%"  name=fileno  size=75 maxlength=255 value=<%=fileno%> >
						</td>
					</tr>			
					<Tr>
						<Td align=center height="50px">				
							<INPUT TYPE=button name=btn value="(Y)Confirm" onclick=go() class="btn btn-sm btn-danger">
							<INPUT TYPE=button name=btn value="(N)Cancel" onclick=goclr() class="btn btn-sm btn-outline-secondary">
							<INPUT TYPE=button name=btn value="(M)回主畫面" onclick="gom()" class="btn btn-sm btn-outline-secondary">
						</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
			
</body>
</html>


<script language=javascript>
function gom(){
	open("<%=self%>.Fore.asp?whsno="+"<%=session("rpwhsno")%>" ,"_self");
}

function goclr() {
	<%=self%>.empid.value="";
	<%=self%>.action = "<%=self%>.new.asp";
	<%=self%>.submit();
}  
 
	
function go() {
	if(<%=self%>.WHSNO.value==""){ 
		alert("請選擇廠別!!");
		<%=self%>.whsno.focus();
	}else if(<%=self%>.empid.value==""){ 
		alert("請輸入員工編號!!");
		<%=self%>.empid.focus();
	}else if(<%=self%>.rp_dat.value==""){ 
		alert("請輸入事件日期!!");
		<%=self%>.rp_dat.focus();
	}else if(<%=self%>.rp_type.value==""){ 
		alert("請輸入獎懲類別!!");
		<%=self%>.rp_type.focus();
	}else if(<%=self%>.rp_func.value==""){ 
		alert("請輸入獎懲方式!!");
		<%=self%>.rp_func.focus();
	}else if(<%=self%>.rp_memo.value==""){ 
		alert("請輸入事件內容說明!!");
		<%=self%>.rp_memo.focus();
	}else{				
		<%=self%>.action="<%=self%>.upd.asp";
		<%=self%>.submit() ;
	}
} 

</script> 